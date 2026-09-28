"""Replay the measured Scene call paths motivating the 0823 trial.

Counts are inclusive stack-presence diagnostics, not additive phase fractions.
The complete 0822 record remains the authority for sampling qualifications.
"""
import gzip
import hashlib
import json
from pathlib import Path
import re
import sys

P = Path(__file__).resolve().parent
PREVIOUS = P.parent / "change-0822"
OWNER = "pptx_edit_profile_0822::edit_region_0822"
SCENE = "litchi_pptx::shape::reader::Scene::read_with"
FUNCTIONS = [
    SCENE,
    "quick_xml::reader::ns_reader::NsReader<R>::process_event",
    "quick_xml::name::NamespaceResolver::resolve_event",
    "quick_xml::events::attributes::IterState::next",
    "litchi_pptx::opened::transaction::compaction_scene",
]
HEADER = re.compile(r"^\S+\s+\d+\s+\d+\.\d+:\s+(\d+)\s+cycles:u:\s*$")
FRAME = re.compile(r"^\s*([0-9a-fA-F]+)\s+(.+)\s+\((.+)\)\s*$")


def derive():
    binary = json.loads((PREVIOUS / "build.json").read_text())["binaries"]["fp"]["artifact"]
    rows = []
    for member in json.loads((PREVIOUS / "perf/compression.json").read_text()):
        if member["kind"] != "frames":
            continue
        stored = (PREVIOUS / "perf" / Path(member["compressed"]["path"]).name).read_bytes()
        assert len(stored) == member["compressed"]["bytes"]
        assert hashlib.sha256(stored).hexdigest() == member["compressed"]["sha256"]
        raw = gzip.decompress(stored)
        assert len(raw) == member["original"]["bytes"]
        assert hashlib.sha256(raw).hexdigest() == member["original"]["sha256"]
        stacks, current = [], None
        for line in raw.decode().splitlines():
            if HEADER.fullmatch(line):
                current = []
                stacks.append(current)
            elif line.strip():
                frame = FRAME.fullmatch(line)
                assert frame and current is not None, repr(line)
                current.append((re.sub(r"\+0x[0-9a-fA-F]+$", "", frame[2]), frame[3]))
            else:
                current = None
        owned = [s[:s.index((OWNER, binary["path"])) + 1]
                 for s in stacks if (OWNER, binary["path"]) in s]
        scenes = [s for s in owned if (SCENE, binary["path"]) in s]
        rows.append({
            "repeat": member["repeat"],
            "compressed_frames": member["compressed"],
            "whole_samples": len(stacks),
            "owner_samples": len(owned),
            "scene_samples": len(scenes),
            "inclusive_presence_in_scene_stacks": {
                f: sum((f, binary["path"]) in s for s in scenes) for f in FUNCTIONS
            },
        })
    assert len(rows) == 2
    return {"schema": "litchi.performance.0823.profile-basis.v1",
            "scope": "0822 exact-owner call-path diagnostics; overlapping counts, no Amdahl estimate",
            "repeats": rows}


if __name__ == "__main__":
    assert sys.argv[1:] in (["--write"], ["--check"])
    encoded = json.dumps(derive(), indent=2, sort_keys=True) + "\n"
    output = P / "profile-basis.json"
    if sys.argv[1] == "--write":
        assert not output.exists()
        output.write_text(encoded)
    else:
        assert output.read_text() == encoded
    print("0823 profile basis PASS")
