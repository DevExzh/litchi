#!/usr/bin/env python3
"""Audit the live DOC CHPX pointers in two OLE output artifacts.

The audit intentionally uses only the Python standard library.  It compares
the FIB-selected PlcBteChpx and the referenced ChpxFkp payloads while treating
unreferenced stream tails as allocation, not formatting state.
"""

from __future__ import annotations

import argparse
import hashlib
import struct
from dataclasses import dataclass
from pathlib import Path


END_OF_CHAIN = 0xFFFFFFFE
FREE_SECTOR = 0xFFFFFFFF
OLE_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
CHPX_FIB_INDEX = 12
FIB_FC_LCB = 154
FIB_COUNT_OFFSET = 152
WORD_FLAGS_OFFSET = 10
CHPX_PAGE_BYTES = 512


class CompoundFileError(ValueError):
    """Raised when an OLE compound file cannot be safely inspected."""


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def u16(data: bytes, offset: int) -> int:
    try:
        return struct.unpack_from("<H", data, offset)[0]
    except struct.error as error:
        raise CompoundFileError(f"truncated u16 at {offset}") from error


def u32(data: bytes, offset: int) -> int:
    try:
        return struct.unpack_from("<I", data, offset)[0]
    except struct.error as error:
        raise CompoundFileError(f"truncated u32 at {offset}") from error


def u64(data: bytes, offset: int) -> int:
    try:
        return struct.unpack_from("<Q", data, offset)[0]
    except struct.error as error:
        raise CompoundFileError(f"truncated u64 at {offset}") from error


@dataclass(frozen=True)
class DirectoryEntry:
    name: str
    kind: int
    start_sector: int
    size: int


class CompoundFile:
    """Small read-only OLE reader sufficient for DOC stream auditing."""

    def __init__(self, path: Path) -> None:
        self.path = path
        self.data = path.read_bytes()
        if self.data[:8] != OLE_SIGNATURE:
            raise CompoundFileError(f"{path}: not an OLE compound file")
        self.sector_size = 1 << u16(self.data, 30)
        self.mini_sector_size = 1 << u16(self.data, 32)
        self.num_fat_sectors = u32(self.data, 44)
        self.first_directory_sector = u32(self.data, 48)
        self.mini_stream_cutoff = u32(self.data, 56)
        self.first_mini_fat_sector = u32(self.data, 60)
        self.num_mini_fat_sectors = u32(self.data, 64)
        self.first_difat_sector = u32(self.data, 68)
        self.num_difat_sectors = u32(self.data, 72)
        if self.sector_size not in (512, 4096):
            raise CompoundFileError(f"{path}: unsupported sector size")
        if self.mini_sector_size != 64:
            raise CompoundFileError(f"{path}: unsupported mini-sector size")
        self.fat = self._load_fat()
        directory = self._read_chain(self.first_directory_sector, self.fat)
        self.entries = self._read_directory(directory)
        self.root = next((entry for entry in self.entries if entry.kind == 5), None)
        if self.root is None:
            raise CompoundFileError(f"{path}: root storage is missing")
        self.mini_fat = self._load_mini_fat()
        self.mini_stream = self._read_regular(self.root.start_sector, self.root.size)

    def _sector(self, sector: int) -> bytes:
        if sector in (END_OF_CHAIN, FREE_SECTOR):
            raise CompoundFileError(f"{self.path}: invalid sector marker")
        start = 512 + sector * self.sector_size
        end = start + self.sector_size
        if start < 512 or end > len(self.data):
            raise CompoundFileError(f"{self.path}: sector {sector} is outside file")
        return self.data[start:end]

    def _load_difat(self) -> list[int]:
        entries_per_sector = self.sector_size // 4
        difat = [u32(self.data, 76 + index * 4) for index in range(109)]
        sector = self.first_difat_sector
        for _ in range(self.num_difat_sectors):
            if sector in (END_OF_CHAIN, FREE_SECTOR):
                raise CompoundFileError(f"{self.path}: truncated DIFAT chain")
            raw = self._sector(sector)
            difat.extend(u32(raw, index * 4) for index in range(entries_per_sector - 1))
            sector = u32(raw, (entries_per_sector - 1) * 4)
        selected = [sector for sector in difat if sector != FREE_SECTOR]
        if len(selected) < self.num_fat_sectors:
            raise CompoundFileError(f"{self.path}: DIFAT has too few FAT sectors")
        return selected[: self.num_fat_sectors]

    def _load_fat(self) -> list[int]:
        entries_per_sector = self.sector_size // 4
        fat: list[int] = []
        for sector in self._load_difat():
            raw = self._sector(sector)
            fat.extend(u32(raw, index * 4) for index in range(entries_per_sector))
        return fat

    def _read_chain(self, start: int, links: list[int]) -> bytes:
        if start in (END_OF_CHAIN, FREE_SECTOR):
            return b""
        chunks: list[bytes] = []
        seen: set[int] = set()
        sector = start
        while sector not in (END_OF_CHAIN, FREE_SECTOR):
            if sector in seen or sector >= len(links):
                raise CompoundFileError(f"{self.path}: invalid sector chain")
            seen.add(sector)
            chunks.append(self._sector(sector))
            sector = links[sector]
        return b"".join(chunks)

    def _read_regular(self, start: int, size: int) -> bytes:
        if size == 0:
            return b""
        raw = self._read_chain(start, self.fat)
        if len(raw) < size:
            raise CompoundFileError(f"{self.path}: regular stream is truncated")
        return raw[:size]

    def _load_mini_fat(self) -> list[int]:
        if self.num_mini_fat_sectors == 0:
            return []
        raw = self._read_chain(self.first_mini_fat_sector, self.fat)
        count = len(raw) // 4
        return [u32(raw, index * 4) for index in range(count)]

    def _read_mini(self, start: int, size: int) -> bytes:
        if size == 0:
            return b""
        if not self.mini_fat:
            raise CompoundFileError(f"{self.path}: mini FAT is missing")
        chunks: list[bytes] = []
        seen: set[int] = set()
        sector = start
        while sector not in (END_OF_CHAIN, FREE_SECTOR):
            if sector in seen or sector >= len(self.mini_fat):
                raise CompoundFileError(f"{self.path}: invalid mini-sector chain")
            seen.add(sector)
            begin = sector * self.mini_sector_size
            end = begin + self.mini_sector_size
            if end > len(self.mini_stream):
                raise CompoundFileError(f"{self.path}: mini-sector is outside root stream")
            chunks.append(self.mini_stream[begin:end])
            sector = self.mini_fat[sector]
        raw = b"".join(chunks)
        if len(raw) < size:
            raise CompoundFileError(f"{self.path}: mini stream is truncated")
        return raw[:size]

    @staticmethod
    def _read_directory(raw: bytes) -> list[DirectoryEntry]:
        entries: list[DirectoryEntry] = []
        for offset in range(0, len(raw) - 127, 128):
            name_length = u16(raw, offset + 64)
            if name_length < 2:
                continue
            if name_length > 64 or name_length % 2:
                raise CompoundFileError("invalid directory entry name length")
            name = raw[offset : offset + name_length - 2].decode("utf-16le")
            entries.append(
                DirectoryEntry(
                    name=name,
                    kind=raw[offset + 66],
                    start_sector=u32(raw, offset + 116),
                    size=u64(raw, offset + 120),
                )
            )
        return entries

    def stream(self, name: str) -> bytes:
        entry = next((entry for entry in self.entries if entry.name == name and entry.kind == 2), None)
        if entry is None:
            raise CompoundFileError(f"{self.path}: stream {name!r} is missing")
        if entry.size < self.mini_stream_cutoff:
            return self._read_mini(entry.start_sector, entry.size)
        return self._read_regular(entry.start_sector, entry.size)

    def stream_names(self) -> list[str]:
        return sorted(entry.name for entry in self.entries if entry.kind == 2)


@dataclass(frozen=True)
class ChpxView:
    word: bytes
    table: bytes
    pairs: tuple[tuple[int, int], ...]
    chpx_pair: tuple[int, int]
    fcs: tuple[int, ...]
    page_numbers: tuple[int, ...]
    pages: tuple[bytes, ...]


def inspect(path: Path) -> tuple[CompoundFile, ChpxView]:
    compound = CompoundFile(path)
    word = compound.stream("WordDocument")
    table_name = "1Table" if u16(word, WORD_FLAGS_OFFSET) & 0x0200 else "0Table"
    table = compound.stream(table_name)
    count = u16(word, FIB_COUNT_OFFSET)
    pairs = tuple(
        (u32(word, FIB_FC_LCB + index * 8), u32(word, FIB_FC_LCB + index * 8 + 4))
        for index in range(count)
    )
    offset, length = pairs[CHPX_FIB_INDEX]
    bte = table[offset : offset + length]
    if len(bte) != length or length < 12 or (length - 4) % 8:
        raise CompoundFileError(f"{path}: malformed CHPX BTE")
    page_count = (length - 4) // 8
    fcs = tuple(u32(bte, index * 4) for index in range(page_count + 1))
    page_numbers = tuple(u32(bte, (page_count + 1) * 4 + index * 4) & 0x003F_FFFF for index in range(page_count))
    pages = tuple(
        word[page * CHPX_PAGE_BYTES : (page + 1) * CHPX_PAGE_BYTES]
        for page in page_numbers
    )
    if any(len(page) != CHPX_PAGE_BYTES for page in pages):
        raise CompoundFileError(f"{path}: CHPX FKP page exceeds WordDocument")
    return compound, ChpxView(word, table, pairs, (offset, length), fcs, page_numbers, pages)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--candidate", type=Path, required=True)
    parser.add_argument("--baseline", type=Path, required=True)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    candidate_file, candidate = inspect(args.candidate)
    baseline_file, baseline = inspect(args.baseline)
    candidate_sha = sha256(args.candidate.read_bytes())
    baseline_sha = sha256(args.baseline.read_bytes())
    print(f"candidate_artifact={args.candidate} bytes={args.candidate.stat().st_size} sha256={candidate_sha}")
    print(f"baseline_artifact={args.baseline} bytes={args.baseline.stat().st_size} sha256={baseline_sha}")
    if candidate_file.stream_names() != baseline_file.stream_names():
        raise CompoundFileError("candidate and baseline stream inventories differ")
    unchanged = []
    for name in candidate_file.stream_names():
        candidate_stream = candidate_file.stream(name)
        baseline_stream = baseline_file.stream(name)
        if name not in {"WordDocument", "1Table", "0Table"}:
            if candidate_stream != baseline_stream:
                raise CompoundFileError(f"non-DOC stream changed: {name}")
            unchanged.append((name, len(candidate_stream), sha256(candidate_stream)))
    print("unchanged_streams=" + ", ".join(f"{name}:{size}:{digest}" for name, size, digest in unchanged))
    if len(candidate.pairs) != len(baseline.pairs):
        raise CompoundFileError("candidate and baseline FIB pair counts differ")
    pair_diffs = [
        (index, left, right)
        for index, (left, right) in enumerate(zip(candidate.pairs, baseline.pairs))
        if left != right
    ]
    if [index for index, _, _ in pair_diffs] != [CHPX_FIB_INDEX]:
        raise CompoundFileError(f"unexpected FIB pair differences: {pair_diffs}")
    print(f"fib12_candidate={candidate.chpx_pair} fib12_baseline={baseline.chpx_pair}")
    if candidate.fcs != baseline.fcs or candidate.pages != baseline.pages:
        raise CompoundFileError("live CHPX ranges or page payloads differ")
    print(f"chpx_page_count={len(candidate.pages)}")
    for index, (start, end), candidate_page, baseline_page in zip(
        range(len(candidate.pages)), zip(candidate.fcs, candidate.fcs[1:]), candidate.page_numbers, baseline.page_numbers
    ):
        payload_sha = sha256(candidate.pages[index])
        print(f"chpx[{index}] fc={start}..{end} candidate_pn={candidate_page} baseline_pn={baseline_page} fkp_sha256={payload_sha}")
    print("result=PASS: skipped blocks are not reachable through the FIB-selected PlcBteChpx")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except CompoundFileError as error:
        print(f"audit_error: {error}")
        raise SystemExit(2)
