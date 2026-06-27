from __future__ import annotations

import flet as ft

from bootstrap import install_code_path, prepare_runtime

prepare_runtime()
install_code_path()

from app_flet import main as app_main  # noqa: E402


def main(page: ft.Page):
    return app_main(page)


if __name__ == "__main__":
    ft.run(main, assets_dir="assets")
