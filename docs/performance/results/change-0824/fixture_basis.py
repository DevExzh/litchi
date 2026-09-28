"""Independently identify compaction bytes in the pinned real-file reference."""
import difflib
import hashlib
import json
from pathlib import Path
import sys
import zipfile

P = Path(__file__).resolve().parent
ROOT = P.parents[3]


def derive():
    real = json.loads((P / "corpus-inputs.json").read_text())["real"]
    descriptors = [real[key] for key in ("input", "reference")]
    payloads = []
    for descriptor in descriptors:
        path = ROOT / descriptor["path"]
        data = path.read_bytes()
        assert len(data) == descriptor["bytes"]
        assert hashlib.sha256(data).hexdigest() == descriptor["sha256"]
        with zipfile.ZipFile(path) as archive:
            payloads.append(archive.read("ppt/slides/slide1.xml"))
    source, published = payloads
    selected = b"<a:t>Learning PPTX</a:t>"
    replacement = b"<a:t>litchi-perf-0638-ordinary-save</a:t>"
    assert source.count(selected) == 1
    staged = source.replace(selected, replacement, 1)
    changes = []
    for tag, start, end, output_start, output_end in difflib.SequenceMatcher(
            None, staged, published, autojunk=False).get_opcodes():
        if tag != "equal":
            changes.append({"kind": tag, "source_span": [start, end],
                            "output_span": [output_start, output_end],
                            "source_hex": staged[start:end].hex(),
                            "output_hex": published[output_start:output_end].hex()})
    assert changes == [{"kind": "delete", "source_span": [55, 57],
                        "output_span": [55, 55], "source_hex": "0d0a", "output_hex": ""}]
    assert staged[:55].endswith(b"?>") and staged[57:].startswith(b"<p:sld ")
    assert staged[:55] + staged[57:] == published
    return {"schema": "litchi.performance.0824.fixture-basis.v1",
            "input": descriptors[0], "reference": descriptors[1],
            "part": "ppt/slides/slide1.xml", "input_part_bytes": len(source),
            "staged_part_bytes": len(staged), "published_part_bytes": len(published),
            "staged_sha256": hashlib.sha256(staged).hexdigest(),
            "published_sha256": hashlib.sha256(published).hexdigest(),
            "compaction_changes_after_exact_target_text_replacement": changes,
            "scope": "Byte-difference evidence for one pinned fixture; no timing or general semantic proof."}


if __name__ == "__main__":
    assert sys.argv[1:] in (["--write"], ["--check"])
    encoded = json.dumps(derive(), indent=2, sort_keys=True) + "\n"
    output = P / "fixture-basis.json"
    if sys.argv[1] == "--write":
        assert not output.exists()
        output.write_text(encoded)
    else:
        assert output.read_text() == encoded
    print("0824 fixture basis PASS: only outer CRLF removed after exact text replacement")
