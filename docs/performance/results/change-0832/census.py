"""Bounded physical-column-record census of the checked XLSX test fixtures.

This is a structural observation, not a benchmark or production-parser oracle.
It counts direct SpreadsheetML cols/col records without interpreting MCE.
"""
import collections
import hashlib
import io
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
NAMESPACES = (
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "http://purl.oclc.org/ooxml/spreadsheetml/main",
)
MAX_ARCHIVE = 16 * 1024 * 1024
MAX_XML = 16 * 1024 * 1024
MAX_XML_PER_ARCHIVE = 64 * 1024 * 1024


def sha(data):
    return hashlib.sha256(data).hexdigest()


def derive():
    rows = []
    totals = collections.Counter()
    distribution = collections.Counter()
    for path in sorted((ROOT / "test-data").rglob("*.xlsx")):
        assert not path.is_symlink()
        row = {"path": str(path.relative_to(ROOT)), "bytes": path.stat().st_size}
        if row["bytes"] > MAX_ARCHIVE:
            row["status"] = "archive-size-limit"
            rows.append(row)
            continue
        data = path.read_bytes()
        row["sha256"] = sha(data)
        row["worksheets"] = []
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as archive:
                names = [entry.filename for entry in archive.infolist()]
                if len(names) != len(set(names)):
                    raise ValueError("duplicate ZIP member names")
                budget = MAX_XML_PER_ARCHIVE
                for entry in archive.infolist():
                    if not (entry.filename.startswith("xl/worksheets/")
                            and entry.filename.endswith(".xml")
                            and "/" not in entry.filename[len("xl/worksheets/"):]):
                        continue
                    sheet = {"member": entry.filename, "bytes": entry.file_size}
                    row["worksheets"].append(sheet)
                    if entry.file_size > MAX_XML or entry.file_size > budget:
                        sheet["status"] = "xml-size-limit"
                        continue
                    budget -= entry.file_size
                    try:
                        with archive.open(entry) as stream:
                            xml = stream.read(MAX_XML + 1)
                        if len(xml) != entry.file_size or len(xml) > MAX_XML:
                            raise ValueError("XML length disagrees with bounded ZIP metadata")
                        sheet["sha256"] = sha(xml)
                        if b"<!DOCTYPE" in xml or b"<!ENTITY" in xml:
                            raise ValueError("DTD/entity declaration excluded from census")
                        root = ET.fromstring(xml)
                        namespaces = [ns for ns in NAMESPACES if root.tag == "{" + ns + "}worksheet"]
                        if len(namespaces) != 1:
                            raise ValueError("not a recognized SpreadsheetML worksheet root")
                        ns = namespaces[0]
                        containers = root.findall("{" + ns + "}cols")
                        records = sum(len(item.findall("{" + ns + "}col")) for item in containers)
                        sheet.update(status="counted", cols_elements=len(containers), records=records)
                    except (ValueError, ET.ParseError, RuntimeError, zipfile.BadZipFile,
                            NotImplementedError, OSError) as error:
                        sheet.update(status="excluded", reason=str(error))
            row["status"] = "inspected"
        except (ValueError, RuntimeError, zipfile.BadZipFile, NotImplementedError, OSError) as error:
            row.update(status="excluded", reason=str(error))
        rows.append(row)
    for row in rows:
        totals["archives_" + row["status"]] += 1
        for sheet in row.get("worksheets", []):
            totals["worksheets_" + sheet["status"]] += 1
            if sheet["status"] == "counted":
                distribution[sheet["records"]] += 1
                totals["worksheets_zero_records" if sheet["records"] == 0 else
                       "worksheets_one_record" if sheet["records"] == 1 else
                       "worksheets_multiple_records"] += 1
                if sheet["cols_elements"]:
                    totals["worksheets_with_cols"] += 1
    return {"schema": "litchi.performance.0832.column-census.v1",
            "base": "7eeaab48c0281b06f53527d4c4f4ea79050d27e7",
            "script_sha256": sha(Path(__file__).read_bytes()),
            "limits": {"archive_bytes": MAX_ARCHIVE, "xml_bytes": MAX_XML,
                       "xml_bytes_per_archive": MAX_XML_PER_ARCHIVE},
            "scope": "All test-data/**/*.xlsx paths; physical direct cols/col only; no MCE processing, parser admission, producer frequency, or performance inference.",
            "totals": dict(sorted(totals.items())),
            "record_count_distribution": {str(k): v for k, v in sorted(distribution.items())},
            "archives": rows}


def main():
    value = derive()
    path = P / "census.json"
    if sys.argv[1:] == ["--check"]:
        assert json.loads(path.read_text()) == value
    else:
        assert not sys.argv[1:]
        with path.open("x") as stream:
            json.dump(value, stream, indent=2, sort_keys=True)
            stream.write("\n")
    print(json.dumps(value["totals"], sort_keys=True))


if __name__ == "__main__":
    main()
