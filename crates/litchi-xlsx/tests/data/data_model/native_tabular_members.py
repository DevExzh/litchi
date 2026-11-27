"""Reproduce metadata observations for the one pinned native fixture.

This is an evidence extractor, not a general XLDM reader or an Xpress decoder.
Run with Python 3 from any directory; JSON is written to standard output.
"""

import hashlib
import json
from pathlib import Path
import struct
import xml.etree.ElementTree as ET
import zipfile


def checksum(payload):
    """Native complemented, non-reflected CRC observed in this fixture."""
    crc = 0xFFFFFFFF
    for byte in payload:
        crc ^= byte << 24
        for _ in range(8):
            crc = ((crc << 1) ^ (0x04C11DB7 if crc & 0x80000000 else 0)) & 0xFFFFFFFF
    return crc ^ 0xFFFFFFFF


def main():
    fixture = Path(__file__).with_name("tdf167689_x15_namespace.xlsx")
    source = fixture.read_bytes()
    digest = hashlib.sha256(source).hexdigest()
    assert digest == "b9064db5694c5e60928a7b6732aabf440c7d2f997d907834e38f927342ec8688"
    with zipfile.ZipFile(fixture) as archive:
        data = archive.read("xl/model/item.data")
    text = data[:4096].decode("utf-16le").rstrip("\0")
    header = ET.fromstring(text[text.index("<BackupLog>"):])
    directory_offset = int(header.findtext("m_cbOffsetHeader"))
    directory_size = int(header.findtext("DataSize"))
    directory = ET.fromstring(data[directory_offset:directory_offset + directory_size].decode("utf-16le"))
    entries = directory.findall("BackupFile")
    assert len(entries) == int(header.findtext("Files")) == 48
    ranges = []
    end = int(header.findtext("m_cbOffsetData"))
    for entry in entries:
        offset = int(entry.findtext("m_cbOffsetHeader"))
        size = int(entry.findtext("Size"))
        assert offset == end and size >= 4
        end = offset + size
        payload = data[offset:end - 4]
        assert checksum(payload) == struct.unpack_from("<I", data, end - 4)[0]
        ranges.append((entry.findtext("Path"), offset, payload))
    assert end <= directory_offset
    _, _, log_bytes = ranges[-1]
    log = ET.fromstring(log_bytes.decode("utf-16"))
    files = log.findall(".//FileList/BackupFile")
    decoded_sizes = {file.findtext("StoragePath"): int(file.findtext("Size")) for file in files}
    assert len(decoded_sizes) == len(files) == 46
    members = []
    total_raw = total_compressed = 0
    for storage_path, offset, payload in ranges[1:-1]:
        position = decoded = 0
        frames = []
        while position < len(payload):
            original, compressed = struct.unpack_from("<HH", payload, position)
            position += 4
            assert position + compressed <= len(payload)
            raw = original == compressed
            frames.append({"original": original, "stored": compressed, "raw": raw})
            total_raw += raw
            total_compressed += not raw
            decoded += original
            position += compressed
        assert position == len(payload) and decoded == decoded_sizes[storage_path]
        members.append({
            "storage_path": storage_path, "offset": offset,
            "stored_without_crc": len(payload), "decoded_from_backup_log": decoded,
            "frames": frames,
        })
    print(json.dumps({
        "fixture": fixture.name, "sha256": digest,
        "model_sha256": hashlib.sha256(data).hexdigest(),
        "header_version": header.findtext("BackupRestoreSyncVersion"),
        "backup_log_version": log.findtext("BackupRestoreSyncVersion"),
        "directory_offset": directory_offset, "directory_size": directory_size,
        "complemented_crc_matches": len(ranges),
        "raw_frames": total_raw, "compressed_frames": total_compressed,
        "members": members,
    }, indent=2))


if __name__ == "__main__":
    main()
