"""Freeze the already specified 0905 plan after both builds finish."""
from pathlib import Path
import hashlib
import json

P = Path(__file__).resolve().parent
ROOT = P.parents[3]


def descriptor(path):
    path = Path(path).resolve()
    return {"path": str(path), "bytes": path.stat().st_size,
            "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}


def main():
    plan = json.loads((P / "plan-template.json").read_text())
    plan["status"] = "frozen"
    plan["workload_cwd"] = str(ROOT)
    plan["source"] = {
        "baseline_manifest": descriptor(P / "baseline-source.json"),
        "candidate_manifest": descriptor(P / "candidate-source.json"),
        "file_count": 10773,
        "candidate_files": ["crates/xml-minifier/src/audit.rs"],
        "root_worktree": str(ROOT), "source_root": "/tmp/litchi-0905-src",
    }
    for role in ("baseline", "candidate"):
        for kind in ("native", "allocator"):
            receipt = P / f"{role}-{kind}-build-run.json"
            raw = json.loads(receipt.read_text())
            if raw["exit_code"] != 0 or raw["inputs_unchanged"] is not True:
                raise RuntimeError(f"unsuccessful build: {receipt}")
            binary = descriptor(P / "binaries" / f"{role}-{kind}")
            saved = json.loads((P / f"{role}-{kind}-binary.json").read_text())
            if binary != saved:
                raise RuntimeError(f"binary custody differs: {role}/{kind}")
            binary["build_receipt"] = descriptor(receipt)
            plan["roles"][role][kind] = binary
    plan["semantic_validator"] = descriptor(P.parent / "change-0705/analyze.py")
    plan["helpers"] = {name: descriptor(P / name)
                       for name in ("run_v2.py", "analyze_v3.py", "freeze.py")}
    plan["environment_receipt"] = descriptor(P / "environment.json")
    plan["method"] = descriptor(P / "METHOD.md")
    plan["semantic_parity"] = descriptor(P / "parity-comparison.json")
    with (P / "plan.json").open("x") as output:
        json.dump(plan, output, indent=2, sort_keys=True)
        output.write("\n")


if __name__ == "__main__":
    main()
