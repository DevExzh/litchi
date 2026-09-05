#!/usr/bin/env python3
"""Generate an independent descriptor-bearing ZIP64 OPC corpus or verify a ZIP.

Python's standard-library zipfile/zlib are independent of soapberry-zip.
Payloads are generated and drained through fixed-size buffers. The archive is
caller-selected scratch storage, never an implicit product spill capability.
"""

import argparse
import hashlib
import json
from pathlib import Path
import platform
import zipfile
import zlib


class SequentialSink:
    def __init__(self, output):
        self.output = output

    def write(self, data):
        return self.output.write(data)

    def flush(self):
        self.output.flush()


def info(name):
    entry = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
    entry.compress_type = zipfile.ZIP_DEFLATED
    entry.create_system = 3
    entry.external_attr = 0o100644 << 16
    return entry


def generate(path, size):
    content_types = (
        b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        b'<Default Extension="bin" ContentType="application/octet-stream"/>'
        b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        b'<Default Extension="xml" ContentType="application/xml"/></Types>'
    )
    relationships = (
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        b'<Relationship Id="rId1" Type="urn:litchi:zip64-probe" Target="document.xml"/>'
        b'</Relationships>'
    )
    with path.open("wb") as output:
        with zipfile.ZipFile(SequentialSink(output), "w", allowZip64=True) as archive:
            archive.writestr(info("[Content_Types].xml"), content_types)
            archive.writestr(info("_rels/.rels"), relationships)
            archive.writestr(info("document.xml"), b'<document>independent ZIP64 corpus</document>')
            with archive.open(info("large.bin"), "w", force_zip64=True) as entry:
                block = bytes(65536)
                remaining = size
                while remaining:
                    chunk = block[:min(remaining, len(block))]
                    entry.write(chunk)
                    remaining -= len(chunk)


def verify(path):
    with path.open("rb") as source:
        archive_sha256 = hashlib.file_digest(source, "sha256").hexdigest()
    result = {
        "python": platform.python_version(),
        "zlib": zlib.ZLIB_RUNTIME_VERSION,
        "archive_bytes": path.stat().st_size,
        "archive_sha256": archive_sha256,
        "buffer_bytes": 65536,
        "entries": [],
    }
    with zipfile.ZipFile(path) as archive:
        for entry in archive.infolist():
            count = 0
            crc = 0
            digest = hashlib.sha256()
            with archive.open(entry) as source:
                while chunk := source.read(65536):
                    count += len(chunk)
                    crc = zlib.crc32(chunk, crc)
                    digest.update(chunk)
            assert count == entry.file_size
            assert crc == entry.CRC
            result["entries"].append({
                "name": entry.filename,
                "uncompressed_bytes": count,
                "compressed_bytes": entry.compress_size,
                "crc32": crc,
                "sha256": digest.hexdigest(),
                "flags": entry.flag_bits,
                "version_needed": entry.extract_version,
            })
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("mode", choices=("generate", "verify"))
    parser.add_argument("path", type=Path)
    parser.add_argument("--size", type=int, default=2**32)
    args = parser.parse_args()
    if args.size < 0:
        parser.error("size must be nonnegative")
    if args.mode == "generate":
        generate(args.path, args.size)
    print(json.dumps(verify(args.path), indent=2, sort_keys=True))


if __name__ == "__main__":
    main()
