#!/usr/bin/env python3
"""Capture serialized whole-command source-backed allocation diagnostics.

The historical no-argument invocation remains the control capture.  The
optional role/output arguments let ``profile-candidate.py`` reuse the exact
artifact and verifier protocol for a candidate binary without touching the
retained control ``runs/`` tree.
"""
import argparse
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
spec = importlib.util.spec_from_file_location("identity", ROOT / "pinned/source-identity.py")
identity = importlib.util.module_from_spec(spec)
spec.loader.exec_module(identity)


def write(path, value):
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n")


def artifact(path):
    return {"name": path.name, "bytes": path.stat().st_size, "sha256": identity.sha256_file(path)}


def require_ancestor(worktree, ancestor, revision):
    result = subprocess.run(
        ["git", "merge-base", "--is-ancestor", ancestor, revision],
        cwd=worktree, capture_output=True, text=True, check=False,
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"{ancestor} is not an ancestor of {revision}: "
            f"{(result.stderr or result.stdout).strip()}"
        )


def load_build(build_path, role, protocol):
    build = json.loads(build_path.read_text())
    if role == "control":
        source = build["source"]
        binary = build["binary"]
        worktree = Path(source["worktree"])
        assert build["status"] == "pass"
        assert identity.source_identity(worktree) == source
        assert identity.binary_identity(Path(binary["path"]), binary["label"]) == binary
    else:
        assert build["status"] == "pass" and build.get("exit_code") == 0
        before = build["source_before"]
        after = build["source_after"]
        assert before == after and before["clean"] is True
        source = after
        worktree = Path(source["worktree"])
        assert identity.source_identity(worktree) == source
        binary = build["binaries"]["normal"]
        observed = identity.binary_identity(Path(binary["path"]), binary.get("label", "candidate/normal"))
        for field in ("sha256", "binary_sha256", "bytes", "binary_bytes", "mode_bits", "executable", "profile"):
            if field in binary:
                assert observed[field] == binary[field], f"candidate binary {field} changed"
    return build, source, binary, worktree


def validate_protocol(protocol, role, source):
    assert protocol["samples"] == 1 and protocol["warmups"] == 0
    assert protocol["corpora"] == ["plain", "media_rich"]
    assert protocol["cpu"] == 2 and protocol["workers"] == 1
    baseline = protocol["control_revision"]
    if role == "control":
        assert baseline == source["revision"]
    else:
        assert protocol.get("profile_role") == "candidate"
        require_ancestor(Path(source["worktree"]), baseline, source["revision"])


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--role", choices=("control", "candidate"), default="control")
    parser.add_argument("--output-root", type=Path, default=ROOT)
    parser.add_argument("--build", type=Path)
    parser.add_argument("--protocol", type=Path)
    parser.add_argument("--verifier", type=Path)
    parser.add_argument("--repo-root", type=Path)
    args = parser.parse_args()
    output_root = args.output_root.expanduser().resolve()
    output_root.mkdir(parents=True, exist_ok=True)
    build_path = (args.build or (ROOT / ("build-control.json" if args.role == "control" else "measurement-build-candidate.json"))).expanduser().resolve()
    protocol_path = (args.protocol or (ROOT / "protocol.json")).expanduser().resolve()
    build = json.loads(build_path.read_text())
    protocol = json.loads(protocol_path.read_text())
    source = build.get("source") if args.role == "control" else build.get("source_after")
    assert isinstance(source, dict)
    build, source, binary, worktree = load_build(build_path, args.role, protocol)
    validate_protocol(protocol, args.role, source)
    verifier = ROOT / "pinned/verify-report.py"
    verifier = (args.verifier or verifier).expanduser().resolve()
    repo_root = (args.repo_root or REPO).expanduser().resolve()
    verifier_hash = identity.sha256_file(verifier)
    versions = {}
    for tool in ("heaptrack", "heaptrack_print"):
        versions[tool] = subprocess.check_output([tool, "--version"], text=True).strip()
    (output_root / "runs").mkdir(exist_ok=False)
    for corpus in protocol["corpora"]:
        folder = output_root / "runs" / corpus
        folder.mkdir()
        selector = f"pptx_source_backed_cross_copy_{corpus}_lifecycle"
        argv = ["taskset", "-c", "2", "/usr/bin/time", "-v", "-o", str(folder / "time-v.txt"),
                "heaptrack", "--record-only", "-o", str(folder / "trace"), binary["path"],
                "--case", selector, *protocol["common_flags"], "--samples", "1", "--warmup", "0",
                "--json", str(folder / "report.json"), "--corpus-manifest", str(folder / "catalog.json")]
        record = {"change": 424, "corpus": corpus, "selector": selector, "status": "running",
                  "started_utc": identity.utc_now(), "scope": protocol["scope"], "argv": argv,
                  "protocol_sha256": identity.sha256_file(protocol_path),
                  "build_sha256": identity.sha256_file(build_path), "source": source,
                  "binary": binary, "tools": versions, "verifier_sha256": verifier_hash,
                  "claim_authorized": False, "performance_claim": None,
                  "environment": {"RUSTUP_TOOLCHAIN": "1.98.1", "DEBUGINFOD_URLS": ""}}
        if args.role == "candidate":
            record["role"] = "candidate"
            record["build_path"] = str(build_path)
            record["protocol_path"] = str(protocol_path)
        write(folder / "capture.json", record)
        print(f"{corpus}: capturing", flush=True)
        with (folder / "stdout.txt").open("wb") as out, (folder / "stderr.txt").open("wb") as err:
            result = subprocess.run(argv, cwd=worktree,
                                    env=os.environ | record["environment"], stdout=out, stderr=err)
        record["exit_code"] = result.returncode
        record["source_unchanged"] = identity.source_identity(worktree) == source
        record["binary_unchanged"] = identity.binary_identity(Path(binary["path"]), binary["label"]) == binary
        record["verifier_unchanged"] = identity.sha256_file(verifier) == verifier_hash
        record["finished_utc"] = identity.utc_now()
        record["status"] = "failed"
        try:
            assert result.returncode == 0 and all(record[key] for key in
                ("source_unchanged", "binary_unchanged", "verifier_unchanged"))
            report = json.loads((folder / "report.json").read_text())
            assert report["binary_identity"]["binary_sha256"] == binary["sha256"]
            assert report["binary_identity"]["binary_bytes"] == binary["bytes"]
            assert report["environment"]["git_revision"] == source["revision"]
            assert report["environment"]["git_worktree_dirty"] is False
            verify = [sys.executable, "-B", str(verifier), "--repo-root", str(repo_root),
                      "--report", str(folder / "report.json"), "--catalog", str(folder / "catalog.json"),
                      "--selector", selector, "--lane", "normal", "--contract", "functional",
                      "--samples", "1", "--warmups", "0"]
            record["verify_argv"] = verify
            checked = subprocess.run(verify, capture_output=True, text=True)
            (folder / "verification.txt").write_text(checked.stdout)
            (folder / "verification-stderr.txt").write_text(checked.stderr)
            record["verify_exit_code"] = checked.returncode
            proof = json.loads(checked.stdout)
            assert checked.returncode == 0 and proof["status"] == "pass"
            assert proof["claim_authorized"] is False and proof["performance_claim"] is None
            assert proof["report_count"] == 1 and proof["selector"] == selector
            assert proof["lane"] == "normal" and proof["samples"] == 1 and proof["warmups"] == 0
            assert proof["reports"][0]["report_sha256"] == identity.sha256_file(folder / "report.json")
            record["status"] = "pass"
        finally:
            record["files"] = [artifact(p) for p in sorted(folder.iterdir()) if p.name != "capture.json"]
            write(folder / "capture.json", record)
        print(f"{corpus}: {record['status']}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
