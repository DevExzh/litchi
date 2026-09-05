#!/usr/bin/env python3
"""Replay the retained 0416 evidence without task-local worktrees or binaries."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile


def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def main() -> None:
    root = Path(__file__).resolve().parent
    repo = root.parents[3]
    identities = json.loads((root / "identities.json").read_text())
    for name, expected in identities["candidate_sources"].items():
        result = subprocess.run(
            ["git", "show", f"{identities['candidate_revision']}:{name}"],
            cwd=repo, capture_output=True,
        )
        actual = result.stdout if result.returncode == 0 else (repo / name).read_bytes()
        assert sha(actual) == expected, ("candidate source", name)
    for name, expected in identities["probe_sources"].items():
        assert sha((root / name).read_bytes()) == expected, ("probe source", name)
    for name, expected in identities["diagnostic_artifacts"].items():
        assert sha((root / name).read_bytes()) == expected, ("diagnostic artifact", name)

    subprocess.run([sys.executable, str(root / "interop.py"), "verify"], check=True)
    manifest = json.loads((root / "manifest.json").read_text())
    corpus = {item["archive_sha256"]: item["archive_bytes"] for item in manifest["fixtures"]}
    for item in json.loads((root / "native-fixtures.json").read_text())["files"]:
        data = (repo / item["path"]).read_bytes()
        assert len(data) == item["bytes"] and sha(data) == item["sha256"], item["path"]
        corpus[item["sha256"]] = item["bytes"]

    captures = [root]
    if (root / "followup/capture.json").is_file():
        captures.append(root / "followup")
    for capture_root in captures:
        capture = json.loads((capture_root / "capture.json").read_text())
        for role in ("control", "candidate"):
            recorded = capture["identities"][role]
            assert recorded["revision"] == identities[f"{role}_revision"], role
            assert recorded["sha256"] == identities["binaries"][role]["sha256"], role
        for group in ("fixtures", "indexed_fixtures", "capability_fixtures"):
            for item in capture.get(group, {}).values():
                assert corpus.get(item["sha256"]) == item["bytes"], (group, item)
        expected = [
            (leg, label, mode)
            for leg in ("A1", "B1", "B2", "A2")
            for group, modes in (("fixtures", ("borrowed", "indexed")),
                                 ("indexed_fixtures", ("indexed",)))
            for label in sorted(capture.get(group, {}))
            for mode in modes
        ]
        assert [(r["leg"], r["fixture"], r["mode"]) for r in capture["runs"]] == expected
        previous_finish = ""
        for run in capture["runs"] + capture["capability_runs"]:
            assert run["exit_code"] == 0 and run["started"] >= previous_finish
            previous_finish = run["finished"]
            role = run["role"]
            fixture = run["fixture"]
            if "leg" in run:
                assert role == ("control" if run["leg"].startswith("A") else "candidate")
                group = "indexed_fixtures" if run.get("indexed_only") else "fixtures"
                tail = ["index", run["mode"], capture[group][fixture]["path"],
                        str(capture["samples"]), str(capture["warmups"])]
            else:
                tail = ["capability", run["mode"], capture["capability_fixtures"][fixture]["path"]]
            assert run["argv"] == [
                "taskset", "-c", str(capture["cpu"]), "/usr/bin/time", "-v", "-o",
                run["time_v"], capture["identities"][role]["binary"], *tail,
            ]
            assert run["finished"] >= run["started"]
        assert len(capture["capability_runs"]) == 4 * len(capture["capability_fixtures"])
        with tempfile.TemporaryDirectory(prefix="litchi-goal-0416-replay-") as scratch:
            output = Path(scratch) / "summary.json"
            subprocess.run(
                [sys.executable, str(root / "summarize.py"), "--root", str(capture_root),
                 "--samples", str(capture["samples"]), "--warmups", str(capture["warmups"]),
                 "--output", str(output)], check=True,
            )
            assert output.read_bytes() == (capture_root / "guard-summary.json").read_bytes()

    original = json.loads((root / "capture.json").read_text())
    counts = json.loads((root / "observations/capture.json").read_text())
    for run in counts["runs"]:
        role, label = run["role"], run["fixture"]
        assert run["argv"] == ["taskset", "-c", "2",
                               original["identities"][role]["binary"], "count", run["source"]["path"]]
        assert corpus.get(run["source"]["sha256"]) == run["source"]["bytes"]
        counted = json.loads((root / "observations" / run["report"]).read_text())
        capability = label in original["capability_fixtures"]
        reference = (root / "capability" / f"{role}-{label}-indexed.json" if capability else
                     root / "guards" / f"{'A1' if role == 'control' else 'B1'}-{label}-indexed.json")
        assert counted["observation"] == json.loads(reference.read_text())["observation"]
        assert counted["reader_at"]["calls"] > 0 and counted["reader_at"]["bytes"] > 0

    for role in ("control", "candidate"):
        expected = json.loads((root / "guards" / f"{'A1' if role == 'control' else 'B1'}-zip32-many256-indexed.json").read_text())
        for kind in ("stat", "profile"):
            diagnostic = json.loads((root / "profile" / f"{role}-{kind}-probe.json").read_text())
            assert diagnostic["sample_count"] == len(diagnostic["samples_ns"]) == 20000
            assert diagnostic["warmups"] == 100
            assert diagnostic["observation"] == expected["observation"]

    for name in identities["passing_checks"]:
        record = json.loads((root / "checks" / f"{name}.json").read_text())
        assert record["exit_code"] == 0, ("check", name)
        assert (root / "checks" / record["log"]).is_file(), ("log", name)
    print("OK: source/corpus/capture bindings, raw-vector summaries and required check outcomes")


if __name__ == "__main__":
    main()
