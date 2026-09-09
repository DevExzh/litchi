#!/usr/bin/env python3
"""Export the six bound perf captures without overwriting prior outputs."""
import subprocess
from common import ROOT, meta, read


def main():
    manifest = read(ROOT / "route-profiles/profiles1-summary-perf-script-commands.json")
    assert len(manifest["records"]) == 6
    for record in manifest["records"]:
        source = record["input_perf_data"]
        assert meta(ROOT / source["path"]) == {k: source[k] for k in ("bytes", "sha256")}
        output = ROOT / record["output"]["path"]
        output.parent.mkdir(parents=True, exist_ok=True)
        with output.open("xb") as stdout, output.with_suffix(".stderr").open("xb") as stderr:
            subprocess.run(record["command"], cwd=record["working_directory"], stdout=stdout,
                           stderr=stderr, check=True, timeout=300)
        assert output.stat().st_size > 0
        print(record["profile_label"], meta(output), flush=True)


if __name__ == "__main__":
    main()
