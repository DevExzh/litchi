"""Recheck the seven naming-only implementation bodies from the repo root."""
from pathlib import Path
import hashlib
import json
import re
import subprocess


proof = json.loads(Path(__file__).with_name("body-parity.json").read_text())
names = {row["before"]: row["after"] for row in proof["functions"]}


def normalized_body(source, name):
    start = re.search(r"pub fn " + name + r"\b", source).start()
    start = source.index("{", start)
    end = source.index("\n}", start) + 2
    body = source[start:end]
    for old, new in names.items():
        body = re.sub(r"\b" + old + r"\b", new, body)
    return re.sub(r",(?=\))", "", re.sub(r"\s+", "", body))


for row in proof["functions"]:
    previous = subprocess.check_output(
        ["git", "show", proof["base_commit"] + ":" + row["file"]], text=True
    )
    current = Path(row["file"]).read_text()
    before = normalized_body(previous, row["before"])
    after = normalized_body(current, row["after"])
    assert before == after, row["after"]
    digest = hashlib.sha256(after.encode()).hexdigest()
    assert digest == row["normalized_body_sha256"], row["after"]
print("Seven canonical implementation bodies match their explicit-limit originals.")
