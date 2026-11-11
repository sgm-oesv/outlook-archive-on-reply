#!/usr/bin/env python3
"""Erzeugt die Platzhalter-Icons (assets/icon-{16,32,64,80}.png)."""
import zlib
import struct
from pathlib import Path

ASSETS = Path(__file__).resolve().parent.parent / "assets"


def png(path: Path, size: int) -> None:
    def chunk(tag: bytes, data: bytes) -> bytes:
        payload = tag + data
        return struct.pack(">I", len(data)) + payload + struct.pack(">I", zlib.crc32(payload) & 0xFFFFFFFF)

    ihdr = struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0)
    row = b"\x00" + b"\x28\x6e\xa8" * size
    path.write_bytes(
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", ihdr)
        + chunk(b"IDAT", zlib.compress(row * size))
        + chunk(b"IEND", b"")
    )


def main() -> None:
    ASSETS.mkdir(parents=True, exist_ok=True)
    for s in (16, 32, 64, 80):
        png(ASSETS / f"icon-{s}.png", s)
    print("Icons erstellt in", ASSETS)


if __name__ == "__main__":
    main()
