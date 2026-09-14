#!/usr/bin/env python3
"""Shared-string composition of one XLS fixture, read straight from its SST.

Independent of the repository's own code: `olefile` plus the standard library.
It is used by change 0576 to justify two statements the record makes — that one
of `54016.xls`'s 7,893 strings is empty (which is why exactly two allocations
short of `2 x 7,893` are removed), and that rich text is rare and `ExtRst` absent
on these fixtures (which is why the measure path leaves both materializing).

    python3 -B sst_composition.py [fixture ...]
"""

import struct
import sys

import olefile

DEFAULT = [
    "test-data/poi/test-data/spreadsheet/54016.xls",
    "test-data/ole/xls/WithCustomViews.xls",
    "test-data/ole/xls/ConditionalFormattingSamples.xls",
]


def globals_stream(path):
    return olefile.OleFileIO(path).openstream("Workbook").read()


def sst_segments(data):
    """The SST record payload and its Continue run, in globals order."""
    offset, segments, started = 0, [], False
    while offset + 4 <= len(data):
        kind, length = struct.unpack_from("<HH", data, offset)
        body = data[offset + 4 : offset + 4 + length]
        if kind == 0x00FC and not started:
            started = True
            segments.append(body)
        elif started and kind == 0x003C:
            segments.append(body)
        elif started or kind == 0x000A:
            break
        offset += 4 + length
    return segments


def boundsheets(data):
    """`BoundSheet8` count and the bytes of name payload they carry.

    A source-backed open decodes those names through `utils::parse_string_record`,
    which is where the `String::from_utf16` instructions that survive this change
    come from.
    """
    offset, count, name_bytes = 0, 0, 0
    while offset + 4 <= len(data):
        kind, length = struct.unpack_from("<HH", data, offset)
        if kind == 0x000A:
            break
        if kind == 0x0085:
            count += 1
            name_bytes += length - 6
        offset += 4 + length
    return count, name_bytes


class Cursor:
    def __init__(self, segments):
        self.segments, self.index, self.offset = segments, 0, 0

    def remaining(self):
        if self.index >= len(self.segments):
            return 0
        return len(self.segments[self.index]) - self.offset

    def advance(self):
        self.index += 1
        self.offset = 0

    def byte(self):
        while self.remaining() == 0:
            self.advance()
        value = self.segments[self.index][self.offset]
        self.offset += 1
        return value

    def take(self, count):
        out = b""
        while len(out) < count:
            if self.remaining() == 0:
                self.advance()
            chunk = min(self.remaining(), count - len(out))
            out += self.segments[self.index][self.offset : self.offset + chunk]
            self.offset += chunk
        return out


def analyse(path):
    data = globals_stream(path)
    sheets, name_bytes = boundsheets(data)
    cursor = Cursor(sst_segments(data))
    total, unique = struct.unpack("<II", cursor.take(8))
    empty = rich = extrst = continued = 0
    for _ in range(unique):
        count = struct.unpack("<H", cursor.take(2))[0]
        flags = cursor.byte()
        runs = struct.unpack("<H", cursor.take(2))[0] if flags & 0x08 else 0
        extension = struct.unpack("<I", cursor.take(4))[0] if flags & 0x04 else 0
        empty += count == 0
        rich += bool(flags & 0x08)
        extrst += bool(flags & 0x04)
        high, done, crossed = flags & 0x01, 0, False
        while done < count:
            per = 2 if high else 1
            take = min(cursor.remaining() // per, count - done)
            cursor.offset += take * per
            done += take
            if done == count:
                break
            cursor.advance()
            high = cursor.segments[cursor.index][0] == 1
            cursor.offset = 1
            crossed = True
        continued += crossed
        cursor.take(runs * 4)
        cursor.take(extension)
    return {
        "path": path,
        "unique": unique,
        "total": total,
        "empty_strings": empty,
        "rich_text": rich,
        "extrst": extrst,
        "crossing_a_continue": continued,
        "bound_sheets": sheets,
        "sheet_name_bytes": name_bytes,
    }


if __name__ == "__main__":
    for fixture in sys.argv[1:] or DEFAULT:
        row = analyse(fixture)
        print(
            "{path}: unique={unique} total={total} empty={empty_strings} "
            "rich_text={rich_text} extrst={extrst} crossing_a_continue={crossing_a_continue} "
            "bound_sheets={bound_sheets} sheet_name_bytes={sheet_name_bytes}".format(**row)
        )
