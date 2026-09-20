#!/usr/bin/env python3
"""Bind real-deck capture digests to exact ZIP XML members."""
import hashlib
import json
import zipfile
from pathlib import Path
P = Path(__file__).resolve().parent
ROOT = P.parents[3]

def prepare():
    source = Path("test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx")
    sha = lambda value: hashlib.sha256(value).hexdigest()
    members = []
    with zipfile.ZipFile(ROOT / source) as archive:
        for info in archive.infolist():
            if info.filename.endswith(".xml"):
                data = archive.read(info.filename)
                members.append(dict(member=info.filename, bytes=len(data), sha256=sha(data)))
    summary = json.loads((P / "focused-summary.json").read_text())
    pairs = next(row for row in summary["rows"] if row["name"]=="real-noop-r0")["capture_pairs"]
    bindings = []
    for pair in pairs:
        matches = [row for row in members if row["sha256"]==pair["raw_sha256"] and row["bytes"]==int(pair["raw_len"])]
        assert len(matches)==1, matches
        bindings.append(dict(member=matches[0]["member"],raw_sha256=pair["raw_sha256"],
            raw_len=int(pair["raw_len"]),output_sha256=pair["output_sha256"],
            output_len=int(pair["output_len"]),output_capacity=int(pair["output_capacity"])))
    assert len(bindings)==14
    return dict(real=dict(path=str(source), archive_sha256=sha((ROOT/source).read_bytes()),
        archive_bytes=(ROOT/source).stat().st_size,capture_member_bindings=bindings),
        generated=dict(recipe="generated:12x8",recipe_source="probe/src/main.rs",
        recipe_source_sha256=sha((P/"probe/src/main.rs").read_bytes()),archive_digest_recorded=False,
        scope="Recipe and observed raw/output SHA-256 pairs are retained; original generated ZIP bytes are not retained."))

if __name__ == "__main__":
    (P / "corpus.json").write_text(json.dumps(prepare(),indent=2)+"\n")
    print("PASS: 14 capture digest pairs uniquely map to real archive XML members")
