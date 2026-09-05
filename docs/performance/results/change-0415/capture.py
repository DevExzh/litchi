#!/usr/bin/env python3
"""Run the predeclared eight-row sequential ABBA guard with distinct binaries."""
import argparse
import datetime
import hashlib
import json
from pathlib import Path
import subprocess


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    for role in ("control", "candidate"):
        parser.add_argument(f"--{role}", type=Path, required=True)
        parser.add_argument(f"--{role}-tree", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--samples", type=int, default=300)
    parser.add_argument("--warmups", type=int, default=30)
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    (args.output / "guards").mkdir(exist_ok=True)
    identities = {}
    for name in ("control", "candidate"):
        binary = getattr(args, name).resolve()
        tree = getattr(args, f"{name}_tree").resolve()
        status = subprocess.check_output(["git", "status", "--porcelain"], cwd=tree, text=True)
        assert not status, (tree, status)
        with binary.open("rb") as source:
            digest = hashlib.file_digest(source, "sha256").hexdigest()
        identities[name] = {
            "binary": str(binary), "sha256": digest, "tree": str(tree),
            "revision": subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=tree, text=True).strip(),
        }
    runs = []
    for leg in ("A1", "B1", "B2", "A2"):
        role = "control" if leg.startswith("A") else "candidate"
        for mode in ("borrowed", "owned"):
            for payload in ("zeros", "mixed"):
                for size in (16384, 1048576):
                    stem = f"{leg}-{mode}-{payload}-{size}"
                    report = (args.output / "guards" / f"{stem}.json").resolve()
                    timing = (args.output / "guards" / f"{stem}.time.txt").resolve()
                    argv = ["taskset", "-c", "2", "/usr/bin/time", "-v", "-o", str(timing),
                            identities[role]["binary"], "guard", mode, payload, str(size),
                            str(args.samples), str(args.warmups)]
                    started = datetime.datetime.now(datetime.timezone.utc).isoformat()
                    with report.open("w") as output:
                        result = subprocess.run(argv, cwd=identities[role]["tree"], stdout=output,
                                                stderr=subprocess.PIPE, text=True, check=False)
                    runs.append({"leg": leg, "stem": stem, "argv": argv, "started": started,
                                 "finished": datetime.datetime.now(datetime.timezone.utc).isoformat(),
                                 "exit_code": result.returncode, "stderr": result.stderr})
                    (args.output / "capture.json").write_text(json.dumps({"identities": identities, "runs": runs}, indent=2) + "\n")
                    if result.returncode:
                        raise SystemExit(f"{stem}: {result.stderr}")
                    print(stem, flush=True)


if __name__ == "__main__":
    main()
