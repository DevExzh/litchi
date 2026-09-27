"""Capture candidate qualification and expose every old-oracle difference.

No changed field is automatically accepted. The receipt distinguishes native
execution/projection validity from exact pre-change oracle equality.
"""
import importlib.util
import json
import pathlib
import subprocess
import time

import compare
from build import P, ROOT, census, sha


if __name__ == "__main__":
    spec = importlib.util.spec_from_file_location("candidate_build", P / "build-candidate.py")
    builder = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(builder)
    builder.guard_inputs()
    sources = {arm: json.loads((P / name).read_text()) for arm, name in
               (("baseline", "source.json"), ("candidate", "candidate-source.json"))}
    assert census() == sources["candidate"]["files"]
    builds = {arm: {row["lane"]: row for row in json.loads((P / name).read_text())}
              for arm, name in (("baseline", "build.json"), ("candidate", "candidate-build.json"))}
    oracle = json.loads((P / "oracle.json").read_text())
    attempt = 0
    while (P / f"candidate-qualification-{attempt}").exists():
        attempt += 1
    folder = P / f"candidate-qualification-{attempt}"
    folder.mkdir()
    rows = []
    for lane in ("native", "allocation"):
        for case in oracle:
            index = len(rows)
            report = folder / f"{index:02d}.json"
            row = {"arm": "candidate", "lane": lane, "case": case, "samples": 1, "warmups": 0}
            binary = pathlib.Path(builds["candidate"][lane]["binary"])
            assert sha(binary) == builds["candidate"][lane]["binary_sha256"]
            command = compare.expected_command(row, builds) + [str(report)]
            row |= {"command": command, "report": str(report.relative_to(P)),
                    "monotonic_start": time.monotonic_ns(), "started": time.time()}
            with report.with_suffix(".stdout").open("w") as out, report.with_suffix(".stderr").open("w") as err:
                result = subprocess.run(command, cwd=ROOT, stdout=out, stderr=err)
            row |= {"exit": result.returncode, "monotonic_end": time.monotonic_ns(), "ended": time.time()}
            row["files"] = {str(path.relative_to(P)): sha(path)
                            for path in (report, report.with_suffix(".stdout"), report.with_suffix(".stderr"))
                            if path.is_file()}
            rows.append(row)
            (folder / "manifest.json").write_text(json.dumps(rows, indent=2) + "\n")
            assert result.returncode == 0, row
            projection = compare.projection(json.loads(report.read_text()), row, "candidate", builds, sources)
            row["projection"] = projection
            row["differences"] = compare.deep_diff(oracle[case], projection)
            row["exact_prior_oracle"] = not row["differences"]
            (folder / "manifest.json").write_text(json.dumps(rows, indent=2) + "\n")
            assert census() == sources["candidate"]["files"]
            builder.guard_inputs()
            print(f"candidate qualification {index + 1}/4: {len(row['differences'])} prior-oracle differences", flush=True)
    raise SystemExit(0 if all(row["exact_prior_oracle"] for row in rows) else 1)
