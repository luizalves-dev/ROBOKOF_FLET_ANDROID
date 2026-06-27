# Autor: Kauê Melo
"""Leitor fallback simples para arquivos .xls BIFF8.
Usado quando xlrd não está instalado ou falha no ambiente.
Foco: extrair valores de células para layouts tabulares simples, como Iquegami.
"""
from __future__ import annotations
import struct
from typing import Any, Dict, Tuple

ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF

class CFBReader:
    def __init__(self, path: str):
        self.data = open(path, "rb").read()
        if self.data[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
            raise ValueError("Arquivo não parece ser .xls OLE/BIFF válido.")
        self.sector_size = 1 << struct.unpack_from("<H", self.data, 30)[0]
        self.mini_sector_size = 1 << struct.unpack_from("<H", self.data, 32)[0]
        self.num_fat = struct.unpack_from("<I", self.data, 44)[0]
        self.first_dir = struct.unpack_from("<I", self.data, 48)[0]
        self.mini_cutoff = struct.unpack_from("<I", self.data, 56)[0]
        self.first_mini_fat = struct.unpack_from("<I", self.data, 60)[0]
        self.num_mini_fat = struct.unpack_from("<I", self.data, 64)[0]
        self.first_difat = struct.unpack_from("<I", self.data, 68)[0]
        self.num_difat = struct.unpack_from("<I", self.data, 72)[0]
        difat = list(struct.unpack_from("<109I", self.data, 76))
        sid = self.first_difat
        for _ in range(self.num_difat):
            if sid in (ENDOFCHAIN, FREESECT):
                break
            sec = self.sector(sid)
            difat.extend(struct.unpack_from("<127I", sec, 0))
            sid = struct.unpack_from("<I", sec, 127 * 4)[0]
        fat = []
        entries_per_sector = self.sector_size // 4
        for sid in difat:
            if sid in (FREESECT, ENDOFCHAIN):
                continue
            if len(fat) // entries_per_sector >= self.num_fat:
                break
            fat.extend(struct.unpack("<%dI" % entries_per_sector, self.sector(sid)))
        self.fat = fat
        self.dir_stream = self.stream_chain(self.first_dir)
        self.entries = []
        for i in range(0, len(self.dir_stream), 128):
            e = self.dir_stream[i:i + 128]
            if len(e) < 128:
                continue
            nlen = struct.unpack_from("<H", e, 64)[0]
            name = e[:max(0, nlen - 2)].decode("utf-16le", "ignore") if nlen >= 2 else ""
            typ = e[66]
            start = struct.unpack_from("<I", e, 116)[0]
            size = struct.unpack_from("<Q", e, 120)[0]
            self.entries.append({"name": name, "type": typ, "start": start, "size": size})
        root = self.entries[0]
        self.root_stream = self.stream_chain(root["start"]) if root["start"] not in (FREESECT, ENDOFCHAIN) else b""
        mf = b""
        if self.first_mini_fat not in (FREESECT, ENDOFCHAIN) and self.num_mini_fat:
            mf = self.stream_chain(self.first_mini_fat, self.num_mini_fat * self.sector_size)
        self.mini_fat = list(struct.unpack("<%dI" % (len(mf) // 4), mf[:len(mf) // 4 * 4])) if mf else []

    def sector(self, sid: int) -> bytes:
        off = (sid + 1) * self.sector_size
        return self.data[off:off + self.sector_size]

    def stream_chain(self, start: int, maxbytes: int | None = None) -> bytes:
        out = []
        sid = start
        seen = set()
        total = 0
        while sid not in (ENDOFCHAIN, FREESECT) and sid < len(self.fat) and sid not in seen:
            seen.add(sid)
            sec = self.sector(sid)
            out.append(sec)
            total += len(sec)
            if maxbytes and total >= maxbytes:
                break
            sid = self.fat[sid]
        data = b"".join(out)
        return data[:maxbytes] if maxbytes else data

    def mini_stream_chain(self, start: int, size: int) -> bytes:
        out = []
        sid = start
        seen = set()
        total = 0
        while sid not in (ENDOFCHAIN, FREESECT) and sid < len(self.mini_fat) and sid not in seen and total < size:
            seen.add(sid)
            off = sid * self.mini_sector_size
            out.append(self.root_stream[off:off + self.mini_sector_size])
            total += self.mini_sector_size
            sid = self.mini_fat[sid]
        return b"".join(out)[:size]

    def get_stream(self, name: str) -> bytes:
        for e in self.entries:
            if e["name"].lower() == name.lower():
                if e["size"] < self.mini_cutoff and e["type"] == 2:
                    return self.mini_stream_chain(e["start"], e["size"])
                return self.stream_chain(e["start"], e["size"])
        raise KeyError(name)


def _decode_rk(raw: bytes) -> float:
    val = struct.unpack("<I", raw)[0]
    mult = 0.01 if (val & 0x01) else 1.0
    if val & 0x02:
        i = val >> 2
        if i & 0x20000000:
            i -= 0x40000000
        return i * mult
    b1 = struct.pack("<I", val & 0xFFFFFFFC) + b"\x00\x00\x00\x00"
    d = struct.unpack("<d", b1)[0]
    if abs(d) < 1e-300 or abs(d) > 1e300:
        b2 = b"\x00\x00\x00\x00" + struct.pack("<I", val & 0xFFFFFFFC)
        d = struct.unpack("<d", b2)[0]
    return d * mult


def _records(stream: bytes):
    pos = 0
    while pos + 4 <= len(stream):
        sid, ln = struct.unpack_from("<HH", stream, pos)
        pos += 4
        payload = stream[pos:pos + ln]
        pos += ln
        yield sid, payload


def _parse_sst(records: list[tuple[int, bytes]]) -> list[str]:
    sst = []
    for idx, (sid, payload) in enumerate(records):
        if sid != 0x00FC:
            continue
        blob = bytearray(payload)
        j = idx + 1
        while j < len(records) and records[j][0] == 0x003C:
            blob.extend(records[j][1])
            j += 1
        data = bytes(blob)
        if len(data) < 8:
            return sst
        _, unique = struct.unpack_from("<II", data, 0)
        p = 8
        for _ in range(unique):
            if p + 3 > len(data):
                break
            cch = struct.unpack_from("<H", data, p)[0]
            p += 2
            opts = data[p]
            p += 1
            rich = opts & 0x08
            ext = opts & 0x04
            is16 = opts & 0x01
            rt = 0
            sz = 0
            if rich and p + 2 <= len(data):
                rt = struct.unpack_from("<H", data, p)[0]
                p += 2
            if ext and p + 4 <= len(data):
                sz = struct.unpack_from("<I", data, p)[0]
                p += 4
            if is16:
                raw = data[p:p + cch * 2]
                p += cch * 2
                s = raw.decode("utf-16le", "ignore")
            else:
                raw = data[p:p + cch]
                p += cch
                s = raw.decode("latin1", "ignore")
            if rich:
                p += 4 * rt
            if ext:
                p += sz
            sst.append(s)
        break
    return sst


def read_xls_cells(path: str) -> Dict[Tuple[int, int], Any]:
    cfb = CFBReader(path)
    names = [e["name"] for e in cfb.entries]
    stream = cfb.get_stream("Workbook") if "Workbook" in names else cfb.get_stream("Book")
    records = list(_records(stream))
    sst = _parse_sst(records)
    cells: Dict[Tuple[int, int], Any] = {}
    for sid, payload in records:
        if sid == 0x0203 and len(payload) >= 14:  # NUMBER
            r, c, _ = struct.unpack_from("<HHH", payload, 0)
            cells[(r, c)] = struct.unpack_from("<d", payload, 6)[0]
        elif sid == 0x027E and len(payload) >= 10:  # RK
            r, c, _ = struct.unpack_from("<HHH", payload, 0)
            cells[(r, c)] = _decode_rk(payload[6:10])
        elif sid == 0x00BD and len(payload) >= 6:  # MULRK
            r, cfirst = struct.unpack_from("<HH", payload, 0)
            clast = struct.unpack_from("<H", payload, len(payload) - 2)[0]
            p = 4
            for c in range(cfirst, clast + 1):
                if p + 6 <= len(payload) - 2:
                    cells[(r, c)] = _decode_rk(payload[p + 2:p + 6])
                p += 6
        elif sid == 0x00FD and len(payload) >= 10:  # LABELSST
            r, c, _ = struct.unpack_from("<HHH", payload, 0)
            ix = struct.unpack_from("<I", payload, 6)[0]
            cells[(r, c)] = sst[ix] if ix < len(sst) else ""
        elif sid == 0x0204 and len(payload) >= 8:  # LABEL
            r, c, _ = struct.unpack_from("<HHH", payload, 0)
            ln = struct.unpack_from("<H", payload, 6)[0]
            cells[(r, c)] = payload[8:8 + ln].decode("latin1", "ignore")
    return cells


def cells_to_matrix(cells: Dict[Tuple[int, int], Any]) -> list[list[Any]]:
    if not cells:
        return []
    max_r = max(r for r, _ in cells)
    max_c = max(c for _, c in cells)
    matrix = [[None for _ in range(max_c + 1)] for _ in range(max_r + 1)]
    for (r, c), v in cells.items():
        matrix[r][c] = v
    return matrix
