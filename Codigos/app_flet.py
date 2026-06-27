"""Interface Flet Android-first do RoboKOF.

Mantém os fluxos de leitura PDF/Excel, validação, Rede BH, fila RoboKOF e TXT EDI.
Integrações Outlook e SAP não fazem parte desta edição móvel.
"""
from __future__ import annotations

import asyncio
import mimetypes
import shutil
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

import flet as ft
import pandas as pd

import cadastro_service
import config
import fila_service
import importador_clientes
import operacao_service
import rastreabilidade_layouts
from layouts.rede_bh.processor import (
    build_kof_queue_from_lote,
    normalize_lote_date,
    process_batch as process_bh_batch,
)
from mobile_health import run_health_check
from mobile_queue import montar_fila_do_excel_validacao


COLOR_BG = "#FFF7F7"
COLOR_SURFACE = "#FFFFFF"
COLOR_SURFACE_ALT = "#FFF1F1"
COLOR_RED = "#E41E2B"
COLOR_RED_DARK = "#B91620"
COLOR_RED_SOFT = "#FDE7E9"
COLOR_DARK = "#2A1013"
COLOR_TEXT = "#1F2937"
COLOR_MUTED = "#6F5D60"
COLOR_BORDER = "#F0C9CE"
COLOR_SUCCESS = "#1F8B4C"
COLOR_WARNING = "#C98300"
COLOR_INFO = "#2563EB"
SUPPORTED_EXTENSIONS = ["pdf", "xlsx", "xls", "xlsm"]


def detectar_tipo_arquivo(path: str | Path) -> str:
    ext = Path(path).suffix.lower()
    if ext == ".pdf":
        return "PDF"
    if ext in {".xlsx", ".xls", ".xlsm"}:
        return "EXCEL"
    return "OUTRO"


def inicializar_item(path: str | Path) -> dict[str, Any]:
    path = Path(path)
    return {
        "id": uuid.uuid4().hex,
        "caminho_arquivo": str(path),
        "nome_arquivo": path.name,
        "tipo_arquivo": detectar_tipo_arquivo(path),
        "layout_id": "",
        "layout_nome": "",
        "status": "Aguardando rastreabilidade",
        "mensagem": "",
        "qtd_linhas_lidas": 0,
        "qtd_linhas_validas": 0,
        "qtd_linhas_descartadas": 0,
        "qtd_linhas_inseridas": 0,
        "modo_rastreabilidade": "NAO",
        "rastreabilidade": {},
        "alertas": [],
        "data_remessa_manual": "",
    }


def format_size(size: int) -> str:
    value = float(size)
    for unit in ("B", "KB", "MB", "GB"):
        if value < 1024 or unit == "GB":
            return f"{value:.1f} {unit}"
        value /= 1024
    return f"{size} B"


def color_for_status(status: str) -> str:
    text = str(status).upper()
    if any(word in text for word in ("ERRO", "FALHA", "INVÁL")):
        return COLOR_RED
    if any(word in text for word in ("ALERTA", "AVISO", "VALIDAR", "SUGESTÃO", "SUGESTAO")):
        return COLOR_WARNING
    if any(word in text for word in ("GERADA", "SUCESSO", "PRONTO", "OK")):
        return COLOR_SUCCESS
    return COLOR_MUTED


class RoboKOFApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.file_picker = ft.FilePicker()
        self.share = ft.Share()
        self.page.services.extend([self.file_picker, self.share])

        self.items: list[dict[str, Any]] = []
        self.bh_files: list[str] = []
        self.bh_last_result: Any = None
        self.outputs: list[Path] = []
        self.current_index = 0
        self.busy = False
        self.log_lines: list[str] = []

        self.content = ft.Container(expand=True)
        self.progress = ft.ProgressBar(value=0, color=COLOR_RED, bgcolor=COLOR_RED_SOFT, visible=False)
        self.busy_text = ft.Text("", size=12, color=COLOR_MUTED)
        self.status_bar = ft.Container(
            content=ft.Column([self.progress, self.busy_text], spacing=4),
            padding=ft.Padding.only(left=14, right=14, bottom=6),
            bgcolor=COLOR_SURFACE,
        )

        self.nav = ft.NavigationBar(
            selected_index=0,
            bgcolor=COLOR_SURFACE,
            indicator_color=COLOR_RED_SOFT,
            on_change=self.on_nav_change,
            destinations=[
                ft.NavigationBarDestination(icon=ft.Icons.HOME_OUTLINED, selected_icon=ft.Icons.HOME, label="Início"),
                ft.NavigationBarDestination(icon=ft.Icons.UPLOAD_FILE_OUTLINED, selected_icon=ft.Icons.UPLOAD_FILE, label="Converter"),
                ft.NavigationBarDestination(icon=ft.Icons.STORE_OUTLINED, selected_icon=ft.Icons.STORE, label="Rede BH"),
                ft.NavigationBarDestination(icon=ft.Icons.SETTINGS_OUTLINED, selected_icon=ft.Icons.SETTINGS, label="Operação"),
                ft.NavigationBarDestination(icon=ft.Icons.FOLDER_OUTLINED, selected_icon=ft.Icons.FOLDER, label="Saídas"),
            ],
        )

    def configure_page(self) -> None:
        self.page.title = "RoboKOF"
        self.page.bgcolor = COLOR_BG
        self.page.padding = 0
        self.page.theme_mode = ft.ThemeMode.LIGHT
        self.page.theme = ft.Theme(
            color_scheme_seed=COLOR_RED,
            font_family="Roboto",
            visual_density=ft.VisualDensity.COMFORTABLE,
        )
        self.page.navigation_bar = self.nav

    def build(self) -> None:
        self.configure_page()
        self.refresh_outputs()
        self.render()
        self.page.add(
            ft.Column(
                [self.header(), self.status_bar, self.content],
                spacing=0,
                expand=True,
            )
        )

    def header(self) -> ft.Control:
        return ft.Container(
            bgcolor=COLOR_RED,
            padding=ft.Padding.symmetric(horizontal=16, vertical=14),
            content=ft.Row(
                [
                    ft.Container(
                        width=44,
                        height=44,
                        border_radius=12,
                        bgcolor=ft.Colors.WHITE,
                        alignment=ft.Alignment.CENTER,
                        content=ft.Text("RK", weight=ft.FontWeight.BOLD, color=COLOR_RED, size=18),
                    ),
                    ft.Column(
                        [
                            ft.Text("RoboKOF", color=ft.Colors.WHITE, size=21, weight=ft.FontWeight.BOLD),
                            ft.Text("Conversor de pedidos · Android", color="#FFE8EA", size=12),
                        ],
                        spacing=1,
                        expand=True,
                    ),
                    ft.IconButton(
                        icon=ft.Icons.REFRESH,
                        icon_color=ft.Colors.WHITE,
                        tooltip="Atualizar",
                        on_click=lambda _: self.render(),
                    ),
                ]
            ),
        )

    async def on_nav_change(self, event: ft.Event[ft.NavigationBar]) -> None:
        self.current_index = int(event.control.selected_index)
        self.render()

    def render(self) -> None:
        views = [self.home_view, self.import_view, self.bh_view, self.operation_view, self.outputs_view]
        self.content.content = views[self.current_index]()
        self.page.update()

    def section(self, title: str, subtitle: str, body: list[ft.Control], icon: ft.IconData | None = None) -> ft.Control:
        heading = [
            ft.Row(
                [
                    ft.Icon(icon or ft.Icons.CHEVRON_RIGHT, color=COLOR_RED),
                    ft.Column(
                        [
                            ft.Text(title, size=20, weight=ft.FontWeight.BOLD, color=COLOR_DARK),
                            ft.Text(subtitle, size=12, color=COLOR_MUTED),
                        ],
                        spacing=1,
                        expand=True,
                    ),
                ]
            )
        ]
        return ft.ListView(
            controls=heading + body,
            spacing=12,
            padding=16,
            expand=True,
        )

    def info_card(self, label: str, value: str, icon: ft.IconData, color: str = COLOR_RED) -> ft.Control:
        return ft.Container(
            col={"xs": 6, "sm": 3},
            bgcolor=COLOR_SURFACE,
            border=ft.Border.all(1, COLOR_BORDER),
            border_radius=14,
            padding=14,
            content=ft.Row(
                [
                    ft.Container(
                        width=38,
                        height=38,
                        border_radius=10,
                        bgcolor=COLOR_RED_SOFT,
                        alignment=ft.Alignment.CENTER,
                        content=ft.Icon(icon, color=color, size=21),
                    ),
                    ft.Column(
                        [
                            ft.Text(value, size=20, weight=ft.FontWeight.BOLD, color=COLOR_DARK),
                            ft.Text(label, size=11, color=COLOR_MUTED),
                        ],
                        spacing=0,
                    ),
                ]
            ),
        )

    def home_view(self) -> ft.Control:
        layouts = cadastro_service.listar_layouts_ativos()
        fila_count = 0
        try:
            fila_count = len(fila_service.carregar_fila())
        except Exception:
            pass
        validation_count = len(list(config.PEDIDOS_A_VALIDAR_DIR.rglob("*.xlsx")))
        txt_count = len(list(config.OUT_TXT_DIR.rglob("*.txt")))
        stats = ft.ResponsiveRow(
            [
                self.info_card("Layouts ativos", str(len(layouts)), ft.Icons.VIEW_MODULE_OUTLINED),
                self.info_card("Na fila", str(fila_count), ft.Icons.QUEUE_OUTLINED, COLOR_INFO),
                self.info_card("Validações", str(validation_count), ft.Icons.FACT_CHECK_OUTLINED, COLOR_WARNING),
                self.info_card("TXTs gerados", str(txt_count), ft.Icons.DESCRIPTION_OUTLINED, COLOR_SUCCESS),
            ],
            spacing=10,
            run_spacing=10,
        )
        hero = ft.Container(
            bgcolor=COLOR_SURFACE,
            border=ft.Border.all(1, COLOR_BORDER),
            border_radius=18,
            padding=18,
            content=ft.Column(
                [
                    ft.Text("Pedidos de clientes para o padrão RoboKOF", size=19, weight=ft.FontWeight.BOLD, color=COLOR_DARK),
                    ft.Text(
                        "Importe PDFs ou planilhas, identifique o layout, confira o Excel de validação e gere os arquivos operacionais.",
                        color=COLOR_MUTED,
                        size=13,
                    ),
                    ft.Row(
                        [
                            ft.FilledButton("Converter arquivos", icon=ft.Icons.UPLOAD_FILE, on_click=lambda _: self.go_to(1)),
                            ft.OutlinedButton("Ver saídas", icon=ft.Icons.FOLDER, on_click=lambda _: self.go_to(4)),
                        ],
                        wrap=True,
                    ),
                ],
                spacing=12,
            ),
        )
        flow = ft.Container(
            bgcolor=COLOR_SURFACE_ALT,
            border_radius=16,
            padding=16,
            content=ft.Column(
                [
                    ft.Text("Fluxo seguro", weight=ft.FontWeight.BOLD, color=COLOR_DARK),
                    ft.Text("1. Arquivo → 2. Layout → 3. Extração → 4. Excel de validação → 5. Fila → 6. TXT", color=COLOR_MUTED),
                    ft.Text("Nenhuma leitura envia pedido direto para TXT. A conferência permanece obrigatória.", color=COLOR_WARNING, size=12),
                ],
                spacing=5,
            ),
        )
        return self.section("Visão geral", "Operação local, sem SAP e sem Outlook", [stats, hero, flow], ft.Icons.DASHBOARD_OUTLINED)

    def go_to(self, index: int) -> None:
        self.current_index = index
        self.nav.selected_index = index
        self.render()

    def import_view(self) -> ft.Control:
        action_bar = ft.Container(
            bgcolor=COLOR_SURFACE,
            border=ft.Border.all(1, COLOR_BORDER),
            border_radius=14,
            padding=12,
            content=ft.Row(
                [
                    ft.FilledButton("Adicionar", icon=ft.Icons.ADD, on_click=self.pick_import_files, disabled=self.busy),
                    ft.OutlinedButton("Processar", icon=ft.Icons.PLAY_ARROW, on_click=self.process_imports, disabled=self.busy or not self.items),
                    ft.TextButton("Limpar", icon=ft.Icons.DELETE_SWEEP_OUTLINED, on_click=self.clear_imports, disabled=self.busy or not self.items),
                ],
                wrap=True,
            ),
        )
        controls: list[ft.Control] = [action_bar]
        if not self.items:
            controls.append(self.empty_state("Nenhum arquivo adicionado", "Selecione PDFs ou planilhas de pedidos.", ft.Icons.UPLOAD_FILE))
        else:
            controls.extend(self.import_item_card(item) for item in self.items)
        return self.section(
            "Conversão de layouts",
            "PDF/Excel → extração → validação. O processamento pode levar alguns minutos em lotes grandes.",
            controls,
            ft.Icons.UPLOAD_FILE,
        )

    def import_item_card(self, item: dict[str, Any]) -> ft.Control:
        layouts = cadastro_service.listar_layouts_ativos(item["tipo_arquivo"])
        options = [
            ft.DropdownOption(key=str(row["layout_id"]), text=str(row["nome_layout"]))
            for _, row in layouts.sort_values("nome_layout").iterrows()
        ]

        async def layout_changed(event: ft.Event[ft.Dropdown]) -> None:
            layout_id = str(event.control.value or "")
            layout = cadastro_service.buscar_layout(layout_id=layout_id)
            item["layout_id"] = layout_id
            item["layout_nome"] = str((layout or {}).get("nome_layout", ""))
            item["status"] = "Pronto para validação" if layout_id else "Aguardando layout"
            item["mensagem"] = "Layout selecionado manualmente."
            self.render()

        def remove(_: Any) -> None:
            self.items = [candidate for candidate in self.items if candidate["id"] != item["id"]]
            self.render()

        status_color = color_for_status(item.get("status", ""))
        subtitle = item.get("mensagem") or f"Tipo: {item['tipo_arquivo']}"
        metrics = ""
        if item.get("qtd_linhas_lidas"):
            metrics = (
                f"Lidas {item.get('qtd_linhas_lidas', 0)} · "
                f"Válidas {item.get('qtd_linhas_validas', 0)} · "
                f"Descartadas {item.get('qtd_linhas_descartadas', 0)}"
            )
        return ft.Card(
            bgcolor=COLOR_SURFACE,
            content=ft.Container(
                padding=14,
                content=ft.Column(
                    [
                        ft.Row(
                            [
                                ft.Icon(ft.Icons.PICTURE_AS_PDF if item["tipo_arquivo"] == "PDF" else ft.Icons.TABLE_VIEW, color=COLOR_RED),
                                ft.Column(
                                    [
                                        ft.Text(item["nome_arquivo"], weight=ft.FontWeight.BOLD, color=COLOR_DARK, max_lines=2),
                                        ft.Text(item.get("status", ""), size=12, color=status_color, weight=ft.FontWeight.BOLD),
                                    ],
                                    spacing=2,
                                    expand=True,
                                ),
                                ft.IconButton(ft.Icons.CLOSE, tooltip="Remover", on_click=remove, disabled=self.busy),
                            ]
                        ),
                        ft.Dropdown(
                            label="Layout",
                            value=str(item.get("layout_id") or "") or None,
                            options=options,
                            enable_search=True,
                            enable_filter=True,
                            on_select=layout_changed,
                            disabled=self.busy,
                        ),
                        ft.Text(subtitle, size=11, color=COLOR_MUTED, max_lines=4),
                        ft.Text(metrics, size=11, color=COLOR_INFO, visible=bool(metrics)),
                    ],
                    spacing=9,
                ),
            ),
        )

    async def pick_import_files(self, _: Any = None) -> None:
        files = await self.file_picker.pick_files(
            dialog_title="Selecione pedidos",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=SUPPORTED_EXTENSIONS,
            allow_multiple=True,
            with_data=True,
        )
        if not files:
            return
        await self.set_busy(True, "Copiando arquivos para a área de trabalho…", indeterminate=True)
        try:
            paths = [self.materialize_file(file, "importacoes") for file in files]
            for path in paths:
                if detectar_tipo_arquivo(path) == "OUTRO":
                    continue
                item = inicializar_item(path)
                self.items.append(item)
                await asyncio.to_thread(self.trace_item, item)
            self.toast(f"{len(paths)} arquivo(s) adicionado(s).")
        except Exception as exc:
            self.show_error("Falha ao importar arquivos", exc)
        finally:
            await self.set_busy(False)
            self.render()

    def materialize_file(self, selected: Any, subfolder: str) -> Path:
        destination_dir = config.TEMP_DIR / subfolder
        destination_dir.mkdir(parents=True, exist_ok=True)
        safe_name = Path(str(selected.name)).name
        destination = destination_dir / f"{uuid.uuid4().hex[:8]}_{safe_name}"
        data = getattr(selected, "bytes", None)
        source_path = getattr(selected, "path", None)
        if data:
            destination.write_bytes(bytes(data))
        elif source_path and Path(source_path).exists():
            shutil.copy2(source_path, destination)
        else:
            raise ValueError(f"O Android não disponibilizou os dados de {safe_name}.")
        return destination

    def trace_item(self, item: dict[str, Any]) -> None:
        try:
            result = rastreabilidade_layouts.rastrear_layout_arquivo(
                item["caminho_arquivo"], item["tipo_arquivo"], permitir_auto=True
            )
            if result.aplicar_automaticamente:
                rastreabilidade_layouts.aplicar_rastreabilidade_no_item(item, result)
            else:
                rastreabilidade_layouts.registrar_sugestao_no_item(item, result)
                if result.sucesso and result.layout_id_referencia:
                    item["layout_id"] = result.layout_id_referencia
                    item["layout_nome"] = result.nome_layout_referencia
                    item["status"] = "Layout sugerido — confirme"
        except Exception as exc:
            item["status"] = "Aguardando layout"
            item["mensagem"] = f"Rastreabilidade indisponível: {exc}"

    async def process_imports(self, _: Any = None) -> None:
        pending = [item for item in self.items if item.get("layout_id")]
        without_layout = [item for item in self.items if not item.get("layout_id")]
        if without_layout:
            self.toast(f"Selecione o layout de {len(without_layout)} arquivo(s).", error=True)
            return
        if not pending:
            self.toast("Nenhum arquivo pronto para processar.", error=True)
            return

        await self.set_busy(True, "Extraindo e validando os pedidos…", indeterminate=True)
        try:
            result = await asyncio.to_thread(importador_clientes.processar_importacao, pending)
            by_name = {row.get("nome_arquivo"): row for row in result.get("resultados_por_arquivo", [])}
            for item in self.items:
                processed = by_name.get(item["nome_arquivo"])
                if processed:
                    item.update(processed)
            self.refresh_outputs()
            self.toast(result.get("mensagem_geral", "Processamento concluído."))
        except Exception as exc:
            self.show_error("Erro no processamento", exc)
        finally:
            await self.set_busy(False)
            self.render()

    def clear_imports(self, _: Any = None) -> None:
        self.items.clear()
        self.render()

    def bh_view(self) -> ft.Control:
        date_field = ft.TextField(
            label="Data do lote",
            value=datetime.now().strftime("%d.%m.%Y"),
            hint_text="DD.MM.AAAA",
            data="bh_date",
        )
        controls: list[ft.Control] = [
            ft.Container(
                bgcolor=COLOR_SURFACE,
                border=ft.Border.all(1, COLOR_BORDER),
                border_radius=14,
                padding=14,
                content=ft.Column(
                    [
                        date_field,
                        ft.Row(
                            [
                                ft.FilledButton("Adicionar PDFs", icon=ft.Icons.ADD, on_click=self.pick_bh_files, disabled=self.busy),
                                ft.OutlinedButton(
                                    "Processar lote",
                                    icon=ft.Icons.PLAY_ARROW,
                                    data=date_field,
                                    on_click=self.process_bh,
                                    disabled=self.busy or not self.bh_files,
                                ),
                                ft.TextButton("Limpar", icon=ft.Icons.DELETE_SWEEP, on_click=self.clear_bh, disabled=self.busy),
                            ],
                            wrap=True,
                        ),
                    ],
                    spacing=12,
                ),
            )
        ]
        if self.bh_files:
            controls.append(
                ft.Container(
                    bgcolor=COLOR_SURFACE,
                    border_radius=14,
                    padding=12,
                    content=ft.Column(
                        [ft.Text(f"{len(self.bh_files)} PDF(s) no lote", weight=ft.FontWeight.BOLD)]
                        + [ft.Text(f"• {Path(path).name}", size=11, color=COLOR_MUTED) for path in self.bh_files[:30]],
                        spacing=4,
                    ),
                )
            )
        else:
            controls.append(self.empty_state("Nenhum PDF da Rede BH", "Adicione os pedidos do lote.", ft.Icons.STORE))

        if self.bh_last_result is not None:
            result = self.bh_last_result
            controls.append(
                ft.Container(
                    bgcolor=COLOR_SURFACE_ALT,
                    border_radius=14,
                    padding=14,
                    content=ft.Column(
                        [
                            ft.Text(f"{result.push_label} concluído", weight=ft.FontWeight.BOLD, color=COLOR_SUCCESS),
                            ft.Text(
                                f"PDFs: {result.total_pdfs} · Itens: {result.total_items} · Alertas/erros: {result.total_errors}",
                                color=COLOR_MUTED,
                            ),
                            ft.Row(
                                [
                                    ft.FilledButton("Gerar fila BH", icon=ft.Icons.QUEUE, on_click=self.generate_bh_queue),
                                    ft.OutlinedButton(
                                        "Compartilhar pacote",
                                        icon=ft.Icons.SHARE,
                                        data=str(result.zip_file),
                                        on_click=self.share_from_event,
                                    ),
                                ],
                                wrap=True,
                            ),
                        ],
                        spacing=9,
                    ),
                )
            )
        return self.section(
            "Rede BH",
            "Processamento dedicado, controle de duplicidades, consolidado e fila segura.",
            controls,
            ft.Icons.STORE,
        )

    async def pick_bh_files(self, _: Any = None) -> None:
        files = await self.file_picker.pick_files(
            dialog_title="PDFs Rede BH",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["pdf"],
            allow_multiple=True,
            with_data=True,
        )
        if not files:
            return
        await self.set_busy(True, "Copiando PDFs da Rede BH…", indeterminate=True)
        try:
            self.bh_files.extend(str(self.materialize_file(file, "rede_bh")) for file in files)
            self.toast(f"{len(files)} PDF(s) adicionado(s) ao lote BH.")
        except Exception as exc:
            self.show_error("Falha ao importar PDFs BH", exc)
        finally:
            await self.set_busy(False)
            self.render()

    async def process_bh(self, event: ft.Event[ft.Button]) -> None:
        date_field = event.control.data
        date_value = str(date_field.value or "").strip()
        try:
            normalized = normalize_lote_date(date_value)
        except Exception as exc:
            self.show_error("Data do lote inválida", exc)
            return
        await self.set_busy(True, "Processando lote Rede BH…", indeterminate=True)
        try:
            self.bh_last_result = await asyncio.to_thread(
                process_bh_batch,
                self.bh_files,
                config.BH_OUTPUT_ROOT,
                config.BH_BASE_PATH,
                normalized,
                False,
            )
            self.refresh_outputs()
            self.toast(f"{self.bh_last_result.push_label} processado com sucesso.")
        except Exception as exc:
            self.show_error("Erro no lote Rede BH", exc)
        finally:
            await self.set_busy(False)
            self.render()

    def clear_bh(self, _: Any = None) -> None:
        self.bh_files.clear()
        self.bh_last_result = None
        self.render()

    async def generate_bh_queue(self, _: Any = None) -> None:
        if self.bh_last_result is None:
            return
        await self.set_busy(True, "Montando fila segura da Rede BH…", indeterminate=True)
        try:
            result = await asyncio.to_thread(
                build_kof_queue_from_lote,
                self.bh_last_result.date_dir,
                datetime.now().strftime("%d/%m/%Y"),
                False,
                False,
            )
            df = result.get("df")
            if df is None or df.empty:
                raise ValueError("Nenhuma linha segura foi liberada. Verifique duplicidades e alertas.")
            inserted = await asyncio.to_thread(fila_service.inserir_na_fila, df)
            self.refresh_outputs()
            self.toast(f"{inserted.get('linhas_inseridas', 0)} linha(s) BH inserida(s) na fila.")
        except Exception as exc:
            self.show_error("Fila BH não gerada", exc)
        finally:
            await self.set_busy(False)
            self.render()

    def operation_view(self) -> ft.Control:
        try:
            queue_df = fila_service.carregar_fila()
        except Exception:
            queue_df = pd.DataFrame(columns=config.FILA_COLUMNS)
        orders = 0
        if not queue_df.empty and {"Matricula", "Nº Pedido"}.issubset(queue_df.columns):
            orders = len(queue_df[["Matricula", "Nº Pedido"]].drop_duplicates())
        queue_path = fila_service.obter_caminho_fila()
        body = [
            ft.ResponsiveRow(
                [
                    self.info_card("Linhas na fila", str(len(queue_df)), ft.Icons.FORMAT_LIST_NUMBERED, COLOR_INFO),
                    self.info_card("Pedidos únicos", str(orders), ft.Icons.RECEIPT_LONG, COLOR_SUCCESS),
                ],
                spacing=10,
            ),
            ft.Container(
                bgcolor=COLOR_SURFACE,
                border=ft.Border.all(1, COLOR_BORDER),
                border_radius=14,
                padding=15,
                content=ft.Column(
                    [
                        ft.Text("Geração operacional", weight=ft.FontWeight.BOLD, size=17, color=COLOR_DARK),
                        ft.Text("Gera os arquivos RoboKOF e os TXTs EDI a partir da fila validada.", color=COLOR_MUTED),
                        ft.Row(
                            [
                                ft.FilledButton("Gerar RoboKOF + TXT", icon=ft.Icons.PLAY_ARROW, on_click=self.generate_operational, disabled=self.busy or queue_df.empty),
                                ft.OutlinedButton("Compartilhar fila", icon=ft.Icons.SHARE, data=str(queue_path), on_click=self.share_from_event),
                                ft.TextButton("Esvaziar fila", icon=ft.Icons.DELETE_FOREVER_OUTLINED, on_click=self.confirm_clear_queue, disabled=self.busy or queue_df.empty),
                            ],
                            wrap=True,
                        ),
                    ],
                    spacing=11,
                ),
            ),
            ft.Container(
                bgcolor=COLOR_RED_SOFT,
                border_radius=14,
                padding=14,
                content=ft.Text(
                    "Nesta edição móvel não existem botões de Outlook ou SAP. O resultado final é compartilhado pelo Android para o destino escolhido pelo usuário.",
                    color=COLOR_DARK,
                    size=12,
                ),
            ),
        ]
        return self.section("Operação RoboKOF", "Fila validada, arquivo padrão e TXT EDI", body, ft.Icons.SETTINGS)

    async def generate_operational(self, _: Any = None) -> None:
        await self.set_busy(True, "Gerando arquivos RoboKOF e TXT EDI…", indeterminate=True)
        try:
            result = await asyncio.to_thread(operacao_service.main)
            self.refresh_outputs()
            message = (
                f"RoboKOF: {result.get('robokof_gerados', 0)} · "
                f"TXTs: {result.get('txts_gerados', 0)} · "
                f"Erros: {result.get('erros_gerados', 0)}"
            )
            self.toast(message, error=bool(result.get("erros_gerados", 0)))
        except Exception as exc:
            self.show_error("Falha na geração operacional", exc)
        finally:
            await self.set_busy(False)
            self.render()

    def confirm_clear_queue(self, _: Any = None) -> None:
        async def confirm(_: Any) -> None:
            self.page.pop_dialog()
            try:
                empty = pd.DataFrame(columns=config.FILA_COLUMNS)
                await asyncio.to_thread(fila_service.salvar_fila, empty)
                self.toast("Fila esvaziada.")
                self.render()
            except Exception as exc:
                self.show_error("Não foi possível esvaziar a fila", exc)

        dialog = ft.AlertDialog(
            modal=True,
            title="Esvaziar fila?",
            content=ft.Text("Todos os itens atuais serão removidos da fila operacional."),
            actions=[
                ft.TextButton("Cancelar", on_click=lambda _: self.page.pop_dialog()),
                ft.FilledButton("Esvaziar", icon=ft.Icons.DELETE_FOREVER, on_click=confirm),
            ],
        )
        self.page.show_dialog(dialog)

    def outputs_view(self) -> ft.Control:
        self.refresh_outputs()
        body: list[ft.Control] = [
            ft.Row(
                [
                    ft.FilledButton("Importar Excel validado", icon=ft.Icons.FACT_CHECK, on_click=self.pick_validated_excel, disabled=self.busy),
                    ft.OutlinedButton("Diagnóstico", icon=ft.Icons.HEALTH_AND_SAFETY_OUTLINED, on_click=self.health_check, disabled=self.busy),
                    ft.IconButton(ft.Icons.REFRESH, tooltip="Atualizar lista", on_click=lambda _: self.render()),
                ],
                wrap=True,
            )
        ]
        if not self.outputs:
            body.append(self.empty_state("Nenhuma saída gerada", "Os arquivos aparecerão aqui após o processamento.", ft.Icons.FOLDER_OPEN))
        else:
            body.extend(self.output_card(path) for path in self.outputs[:100])
        return self.section(
            "Saídas e governança",
            "Compartilhe validações, filas, pacotes, planilhas e TXTs pelo Android.",
            body,
            ft.Icons.FOLDER,
        )

    def output_card(self, path: Path) -> ft.Control:
        is_validation = path.suffix.lower() in {".xlsx", ".xls"} and (
            "VALIDACAO" in path.name.upper() or "FILA_KOF" in path.name.upper()
        )
        actions: list[ft.Control] = [
            ft.IconButton(ft.Icons.SHARE, tooltip="Compartilhar", data=str(path), on_click=self.share_from_event)
        ]
        if is_validation:
            actions.insert(0, ft.IconButton(ft.Icons.QUEUE, tooltip="Enviar para a fila", data=str(path), on_click=self.confirm_queue_from_output))
        try:
            modified = datetime.fromtimestamp(path.stat().st_mtime).strftime("%d/%m/%Y %H:%M")
            size = format_size(path.stat().st_size)
        except Exception:
            modified, size = "", ""
        return ft.Card(
            bgcolor=COLOR_SURFACE,
            content=ft.ListTile(
                leading=ft.Icon(self.icon_for_file(path), color=COLOR_RED),
                title=ft.Text(path.name, weight=ft.FontWeight.BOLD, size=13, max_lines=2),
                subtitle=ft.Text(f"{modified} · {size}\n{self.relative_output(path)}", size=10, color=COLOR_MUTED, max_lines=3),
                trailing=ft.Row(actions, spacing=0, tight=True),
            ),
        )

    async def pick_validated_excel(self, _: Any = None) -> None:
        files = await self.file_picker.pick_files(
            dialog_title="Excel validado",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["xlsx", "xls"],
            allow_multiple=False,
            with_data=True,
        )
        if not files:
            return
        try:
            path = self.materialize_file(files[0], "validacoes_reimportadas")
            self.confirm_queue_path(path)
        except Exception as exc:
            self.show_error("Excel não importado", exc)

    def confirm_queue_from_output(self, event: ft.Event[ft.IconButton]) -> None:
        self.confirm_queue_path(Path(str(event.control.data)))

    def confirm_queue_path(self, path: Path) -> None:
        try:
            preview = montar_fila_do_excel_validacao(path)
        except Exception as exc:
            self.show_error("Excel incompatível com a fila", exc)
            return

        async def confirm(_: Any) -> None:
            self.page.pop_dialog()
            await self.insert_validation_queue(path)

        dialog = ft.AlertDialog(
            modal=True,
            title="Inserir na fila RoboKOF?",
            content=ft.Column(
                [
                    ft.Text(path.name, weight=ft.FontWeight.BOLD),
                    ft.Text(f"Aba: {preview['sheet']}"),
                    ft.Text(f"Linhas seguras: {preview['safe_rows']}"),
                    ft.Text(f"Linhas bloqueadas: {preview['blocked_rows']}", color=COLOR_WARNING),
                    ft.Text("Confirme somente após conferir o Excel de validação.", size=12, color=COLOR_RED),
                ],
                tight=True,
                spacing=5,
            ),
            actions=[
                ft.TextButton("Cancelar", on_click=lambda _: self.page.pop_dialog()),
                ft.FilledButton("Confirmar", icon=ft.Icons.QUEUE, on_click=confirm),
            ],
        )
        self.page.show_dialog(dialog)

    async def insert_validation_queue(self, path: Path) -> None:
        await self.set_busy(True, "Validando e inserindo linhas seguras…", indeterminate=True)
        try:
            result = await asyncio.to_thread(montar_fila_do_excel_validacao, path)
            df = result["df"]
            if df is None or df.empty:
                details = "\n".join(result.get("alerts", [])[:8])
                raise ValueError("Nenhuma linha segura foi encontrada." + (f"\n{details}" if details else ""))
            inserted = await asyncio.to_thread(fila_service.inserir_na_fila, df)
            self.toast(
                f"{inserted.get('linhas_inseridas', 0)} linha(s) inserida(s); "
                f"{result.get('blocked_rows', 0)} bloqueada(s)."
            )
        except Exception as exc:
            self.show_error("Fila não alimentada", exc)
        finally:
            await self.set_busy(False)
            self.render()

    async def health_check(self, _: Any = None) -> None:
        await self.set_busy(True, "Validando bases, cadastros e módulos…", indeterminate=True)
        try:
            result = await asyncio.to_thread(run_health_check)
            rows = [
                ft.Text(
                    f"{'✓' if row['status'] == 'OK' else '✕'} {row['item']}: {row['detalhe']}",
                    color=COLOR_SUCCESS if row["status"] == "OK" else COLOR_RED,
                    size=11,
                )
                for row in result["checks"]
            ]
            self.page.show_dialog(
                ft.AlertDialog(
                    title="Diagnóstico RoboKOF",
                    content=ft.ListView(rows, spacing=5, height=420, width=520),
                    actions=[ft.TextButton("Fechar", on_click=lambda _: self.page.pop_dialog())],
                )
            )
        except Exception as exc:
            self.show_error("Diagnóstico não concluído", exc)
        finally:
            await self.set_busy(False)

    def refresh_outputs(self) -> None:
        extensions = {".xlsx", ".xls", ".txt", ".zip", ".csv", ".log"}
        if not config.RESULTADOS_DIR.exists():
            self.outputs = []
            return
        paths = [
            path for path in config.RESULTADOS_DIR.rglob("*")
            if path.is_file() and path.suffix.lower() in extensions and not path.name.startswith("~$")
        ]
        self.outputs = sorted(paths, key=lambda path: path.stat().st_mtime if path.exists() else 0, reverse=True)

    def relative_output(self, path: Path) -> str:
        try:
            return str(path.relative_to(config.RESULTADOS_DIR))
        except Exception:
            return str(path)

    def icon_for_file(self, path: Path) -> ft.IconData:
        return {
            ".xlsx": ft.Icons.TABLE_VIEW,
            ".xls": ft.Icons.TABLE_VIEW,
            ".txt": ft.Icons.DESCRIPTION,
            ".zip": ft.Icons.ARCHIVE,
            ".csv": ft.Icons.DATA_OBJECT,
            ".log": ft.Icons.RECEIPT_LONG,
        }.get(path.suffix.lower(), ft.Icons.INSERT_DRIVE_FILE)

    async def share_from_event(self, event: ft.Event[Any]) -> None:
        await self.share_path(Path(str(event.control.data)))

    async def share_path(self, path: Path) -> None:
        if not path.exists():
            self.toast("Arquivo não encontrado.", error=True)
            return
        try:
            mime, _ = mimetypes.guess_type(path.name)
            result = await self.share.share_files(
                [ft.ShareFile(path=str(path), name=path.name, mime_type=mime or "application/octet-stream")],
                title="RoboKOF",
                text=f"Arquivo gerado pelo RoboKOF: {path.name}",
            )
            if getattr(result, "status", None) == "unavailable":
                await self.export_file(path)
        except Exception:
            await self.export_file(path)

    async def export_file(self, path: Path) -> None:
        try:
            await self.file_picker.save_file(
                dialog_title="Salvar arquivo RoboKOF",
                file_name=path.name,
                allowed_extensions=[path.suffix.lstrip(".")],
                src_bytes=path.read_bytes(),
            )
        except Exception as exc:
            self.show_error("Não foi possível compartilhar/exportar", exc)

    async def set_busy(self, busy: bool, message: str = "", *, indeterminate: bool = False) -> None:
        self.busy = busy
        self.progress.visible = busy
        self.progress.value = None if busy and indeterminate else (0 if busy else 0)
        self.busy_text.value = message if busy else ""
        self.page.update()
        await asyncio.sleep(0)

    def toast(self, message: str, *, error: bool = False) -> None:
        snack = ft.SnackBar(
            content=ft.Text(message, color=ft.Colors.WHITE),
            bgcolor=COLOR_RED if error else COLOR_DARK,
        )
        self.page.show_dialog(snack)

    def show_error(self, title: str, exc: Exception | str) -> None:
        message = str(exc)
        self.log_lines.append(f"{datetime.now():%d/%m/%Y %H:%M:%S} | {title} | {message}")
        self.page.show_dialog(
            ft.AlertDialog(
                title=title,
                content=ft.Text(message, selectable=True),
                actions=[ft.TextButton("Fechar", on_click=lambda _: self.page.pop_dialog())],
            )
        )

    def empty_state(self, title: str, subtitle: str, icon: ft.IconData) -> ft.Control:
        return ft.Container(
            bgcolor=COLOR_SURFACE,
            border=ft.Border.all(1, COLOR_BORDER),
            border_radius=16,
            padding=28,
            alignment=ft.Alignment.CENTER,
            content=ft.Column(
                [
                    ft.Icon(icon, size=48, color=COLOR_BORDER),
                    ft.Text(title, weight=ft.FontWeight.BOLD, color=COLOR_DARK, text_align=ft.TextAlign.CENTER),
                    ft.Text(subtitle, color=COLOR_MUTED, size=12, text_align=ft.TextAlign.CENTER),
                ],
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                spacing=7,
            ),
        )


def main(page: ft.Page) -> None:
    RoboKOFApp(page).build()
