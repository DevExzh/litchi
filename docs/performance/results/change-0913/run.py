"""Create-only source, build, and matched four-variant capture custody for 0913.

The coordinator prepares ``/tmp/litchi-0913-src`` from the frozen 0910 source,
leaving the 0910 source files in place and a ``.owned-0913`` marker at its
root.  ``freeze`` copies only the explicitly named candidate overlays from the
workspace, records an inventory, and writes the capture plan before any build
or benchmark process is started.  ``build`` compiles the candidate with the
0910 release settings and retains a byte-bound copy of the executable.
``capture`` then runs the retained 0910 executable and the candidate executable
in the frozen baseline/candidate AB/BA schedule.

This helper deliberately does not run while being authored; all side effects
described here happen only when the coordinator invokes one of its modes.
Every packet file is create-only so an accidental rerun cannot silently
replace evidence.
"""

from __future__ import annotations

import argparse
import datetime
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
from typing import Callable


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
SOURCE = Path("/tmp/litchi-0913-src")
TARGET = ROOT / "target/litchi-0913-build"
BASELINE_BINARY = ROOT / "docs/performance/results/change-0910/binaries/native"
BASELINE_SOURCE_MANIFEST = ROOT / "docs/performance/results/change-0910/source-v1.json"

PARENT = "66b944b07fcf4c6fa37f10e7ed83a02417455f57"
SOURCE_MARKER = ".owned-0913"
SOURCE_MARKER_CONTENT = "owned-direct-read-0913\n"

DEFAULT_OVERLAYS = (
    "crates/soapberry-zip/src/office.rs",
    "crates/litchi-opc/src/payload.rs",
    "crates/litchi-opc/tests/owned_direct_concurrency.rs",
)

WARMUP = 30
SAMPLES = 20
PAIRS_PER_ROUTE = 12
BLOCKS = PAIRS_PER_ROUTE * 2
CPU = 12
ROUTES = ("materialized", "owned")
VARIANTS = ("baseline", "candidate")
BOOTSTRAP_DRAWS = 10_000
BOOTSTRAP_SEED = 913


class EvidenceError(ValueError):
    """A requested operation would create contradictory custody."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def now() -> str:
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def path_string(path: Path, base: Path = ROOT) -> str:
    path = path.resolve()
    try:
        return str(path.relative_to(base.resolve()))
    except ValueError:
        return str(path)


def descriptor(path: Path, base: Path = ROOT) -> dict[str, object]:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"not a regular file: {path}")
    return {"path": path_string(path, base), "bytes": path.stat().st_size, "sha256": sha256(path)}


def write_new(path: Path, value: object) -> None:
    require(not path.exists() and not path.is_symlink(), f"refusing to overwrite {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, indent=2, sort_keys=True)
        stream.write("\n")


def write_new_bytes(path: Path, data: bytes) -> None:
    require(not path.exists() and not path.is_symlink(), f"refusing to overwrite {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("xb") as stream:
        stream.write(data)
        stream.flush()
        os.fsync(stream.fileno())


def marker(path: Path, expected: str) -> None:
    require(path.is_file() and not path.is_symlink(), f"missing marker: {path}")
    require(path.read_text(encoding="utf-8") == expected, f"marker differs: {path}")


def assert_roots() -> None:
    marker(SOURCE / SOURCE_MARKER, SOURCE_MARKER_CONTENT)
    marker(TARGET / SOURCE_MARKER, SOURCE_MARKER_CONTENT)


def inventory(root: Path) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for path in sorted(root.rglob("*")):
        if path.is_file() and not path.is_symlink():
            rows.append({"path": str(path.relative_to(root)), "bytes": path.stat().st_size, "sha256": sha256(path)})
    return rows


def baseline_inventory() -> dict[str, dict[str, object]]:
    require(BASELINE_SOURCE_MANIFEST.is_file(), f"missing frozen baseline source manifest: {BASELINE_SOURCE_MANIFEST}")
    raw = json.loads(BASELINE_SOURCE_MANIFEST.read_text(encoding="utf-8"))
    require(isinstance(raw, dict) and raw.get("parent") == "468fe2931d6532c32f4c6608590816b8eaa88e31", "0910 source manifest parent differs")
    files = raw.get("files")
    require(isinstance(files, list), "0910 source manifest has no files")
    result: dict[str, dict[str, object]] = {}
    for row in files:
        require(isinstance(row, dict) and isinstance(row.get("path"), str), "malformed 0910 source inventory")
        path = str(row["path"])
        require(path not in result, f"duplicate 0910 source path: {path}")
        result[path] = {"path": path, "bytes": int(row["bytes"]), "sha256": str(row["sha256"])}
    return result


def parse_overlays(values: list[str] | None) -> list[str]:
    overlays = list(DEFAULT_OVERLAYS if values is None else values)
    require(overlays, "at least one candidate overlay is required")
    require(len(set(overlays)) == len(overlays), "candidate overlay paths must be unique")
    for name in overlays:
        path = Path(name)
        require(not path.is_absolute() and ".." not in path.parts, f"unsafe overlay path: {name}")
        require((ROOT / path).is_file() and not (ROOT / path).is_symlink(), f"missing workspace overlay: {name}")
    return overlays


def freeze(attempt: str, overlay_values: list[str] | None) -> None:
    assert_roots()
    overlays = parse_overlays(overlay_values)
    baseline = baseline_inventory()
    overlay_set = set(overlays)

    for name in overlays:
        source_path = SOURCE / name
        source_path.parent.mkdir(parents=True, exist_ok=True)
        write_new_bytes(source_path, (ROOT / name).read_bytes()) if not source_path.exists() else source_path.write_bytes((ROOT / name).read_bytes())

    current = {row["path"]: row for row in inventory(SOURCE)}
    require(SOURCE_MARKER in current, "candidate source marker is missing from inventory")
    allowed_new = {SOURCE_MARKER} | overlay_set
    for name, row in baseline.items():
        require(name in current, f"candidate source removed baseline file: {name}")
        if name not in overlay_set:
            require(current[name]["bytes"] == row["bytes"] and current[name]["sha256"] == row["sha256"], f"unrequested source changed: {name}")
    for name in current:
        require(name in baseline or name in allowed_new, f"unrequested candidate source file: {name}")
    require(set(current) == set(baseline) | allowed_new, "candidate source inventory does not equal baseline plus declared overlays/marker")

    with tempfile.TemporaryDirectory(prefix="litchi-0913-overlay-index-") as directory:
        index = Path(directory) / "index"
        env = dict(os.environ, GIT_INDEX_FILE=str(index))
        subprocess.run(["git", "read-tree", PARENT], cwd=ROOT, env=env, check=True)
        subprocess.run(["git", "add", "--", *overlays], cwd=ROOT, env=env, check=True)
        patch = subprocess.check_output(["git", "diff", "--cached", "--binary", PARENT, "--", *overlays], cwd=ROOT, env=env)

    patch_path = HERE / f"source-overlay-{attempt}.patch"
    write_new_bytes(patch_path, patch)
    source_manifest = {
        "status": "frozen_candidate_source",
        "parent": PARENT,
        "source_root": str(SOURCE),
        "baseline_source_manifest": descriptor(BASELINE_SOURCE_MANIFEST),
        "overlay_paths": overlays,
        "overlays": [
            {
                "path": name,
                "workspace": descriptor(ROOT / name),
                "source": {"path": name, "bytes": current[name]["bytes"], "sha256": current[name]["sha256"]},
            }
            for name in overlays
        ],
        "overlay_patch": descriptor(patch_path),
        "files": list(current.values()),
    }
    manifest_path = HERE / f"source-{attempt}.json"
    write_new(manifest_path, source_manifest)
    plan = make_plan(attempt, descriptor(manifest_path))
    write_new(HERE / "capture-plan.json", plan)
    print(f"frozen candidate source: {len(current)} files; {len(overlays)} overlays")


def make_plan(attempt: str, source_manifest: dict[str, object]) -> dict[str, object]:
    legs: list[dict[str, object]] = []
    route_pair_counts = {route: 0 for route in ROUTES}
    block = 0
    # Interleave route-specific pairs.  Each route receives six AB and six BA
    # pairs, while every process remains serial and globally ordered.
    for round_index in range(PAIRS_PER_ROUTE):
        for route in ROUTES:
            pair = route_pair_counts[route]
            route_pair_counts[route] += 1
            first = VARIANTS if pair % 2 == 0 else tuple(reversed(VARIANTS))
            for position, variant in enumerate(first):
                legs.append({
                    "label": f"block-{block:02d}-pair-{pair:02d}-{route}-{variant}",
                    "block": block,
                    "route": route,
                    "pair": pair,
                    "variant": variant,
                    "position": position,
                })
            block += 1
    return {
        "schema": "litchi.performance.0913.pptx-binary-variant-plan.v1",
        "status": "frozen_before_capture",
        "created_utc": now(),
        "attempt": attempt,
        "parent": PARENT,
        "source_manifest": source_manifest,
        "warmup": WARMUP,
        "samples": SAMPLES,
        "pairs_per_route": PAIRS_PER_ROUTE,
        "blocks": BLOCKS,
        "cpu": CPU,
        "routes": list(ROUTES),
        "variants": list(VARIANTS),
        "primary": "Matched candidate-versus-retained-0910 baseline binary timings for each route; source, output, sink, and semantic guards are required to be byte-identical.",
        "comparison": "Route-specific paired binary control. Ratios are baseline over candidate within one route; no materialized-versus-owned performance claim is made.",
        "statistics": f"Twelve paired processes per route, six AB and six BA; pooled p50/p95/p99 per route and variant; {BOOTSTRAP_DRAWS} paired-block bootstrap draws with seed {BOOTSTRAP_SEED}.",
        "scope": "Warm generated 200-slide media-rich PPTX fixture; CPU 12; 30 warmups and 20 measured samples per child; /usr/bin/time whole-process maximum RSS.",
        "adoption_gate": "Owned total paired median ratio baseline/candidate at least 1.05 with 95% bootstrap lower bound above 1; flag any route total, p95, p99 or whole-process RSS regression above 5% for review.",
        "legs": legs,
    }


def env_for_build() -> dict[str, str]:
    env = dict(os.environ)
    env.update({
        "CARGO_TARGET_DIR": str(TARGET),
        "CARGO_BUILD_JOBS": "2",
        "CARGO_INCREMENTAL": "0",
        "CARGO_PROFILE_RELEASE_LTO": "false",
        "CARGO_PROFILE_RELEASE_CODEGEN_UNITS": "16",
        "CARGO_PROFILE_RELEASE_DEBUG": "0",
        "LC_ALL": "C",
        "LANG": "C",
        "TZ": "UTC",
    })
    for key in ("RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS"):
        env.pop(key, None)
    return env


def common_receipt(env: dict[str, str], manifest: Path) -> dict[str, object]:
    return {
        "cwd": str(SOURCE),
        "environment_overrides": {key: env[key] for key in (
            "CARGO_TARGET_DIR", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL",
            "CARGO_PROFILE_RELEASE_LTO", "CARGO_PROFILE_RELEASE_CODEGEN_UNITS",
            "CARGO_PROFILE_RELEASE_DEBUG", "LC_ALL", "LANG", "TZ",
        )},
        "removed_environment_keys": ["RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS"],
        "source": descriptor(manifest),
        "helper": descriptor(HERE / "run.py"),
        "toolchain": {
            name: {
                "launcher": shutil.which(name),
                "version": subprocess.check_output([name, "--version"], cwd=SOURCE, env=env, text=True).strip(),
            }
            for name in ("cargo", "rustc")
        },
    }


def execute(
    folder: Path,
    argv: list[str],
    env: dict[str, str],
    common: dict[str, object],
    extra: dict[str, object] | None = None,
    after_success: Callable[[], dict[str, object]] | None = None,
) -> dict[str, object]:
    require(not folder.exists() and not folder.is_symlink(), f"refusing to overwrite command directory: {folder}")
    folder.mkdir(parents=True, exist_ok=False)
    start = dict(common, argv=argv, start_utc=now(), **(extra or {}))
    write_new(folder / "start.json", start)
    with (folder / "stdout.txt").open("x", encoding="utf-8") as stdout, (folder / "stderr.txt").open("x", encoding="utf-8") as stderr:
        result = subprocess.run(argv, cwd=SOURCE, env=env, stdout=stdout, stderr=stderr)
    receipt = dict(start, end_utc=now(), exit_code=result.returncode)
    receipt["logs"] = [descriptor(folder / "stdout.txt"), descriptor(folder / "stderr.txt")]
    if (folder / "artifacts").is_dir():
        receipt["artifacts"] = [descriptor(path) for path in sorted((folder / "artifacts").rglob("*")) if path.is_file()]
    if (folder / "process.txt").is_file():
        receipt["process_metrics"] = descriptor(folder / "process.txt")
    if after_success is not None and result.returncode == 0:
        receipt.update(after_success())
    write_new(folder / "command.json", receipt)
    print(f"{folder.name} exit {result.returncode}", flush=True)
    if result.returncode:
        raise SystemExit(result.returncode)
    return receipt


def check_manifest(manifest_path: Path) -> None:
    raw = json.loads(manifest_path.read_text(encoding="utf-8"))
    require(isinstance(raw, dict) and raw.get("parent") == PARENT, "candidate source manifest parent differs")
    for row in raw.get("files", []):
        path = SOURCE / str(row["path"])
        require(path.is_file() and path.stat().st_size == row["bytes"] and sha256(path) == row["sha256"], f"source changed after freeze: {path}")
    require(len(raw.get("files", [])) == len(inventory(SOURCE)), "source inventory cardinality changed after freeze")


def build(attempt: str) -> None:
    assert_roots()
    manifest = HERE / f"source-{attempt}.json"
    plan = HERE / "capture-plan.json"
    require(manifest.is_file() and plan.is_file(), "freeze must complete before build")
    check_manifest(manifest)
    env = env_for_build()
    common = common_receipt(env, manifest)
    binary = TARGET / "release/examples/pptx_owned_benchmark"
    retained = HERE / "binaries/candidate"

    def retain_candidate() -> dict[str, object]:
        require(binary.is_file(), "candidate build did not produce its executable")
        write_new_bytes(retained, binary.read_bytes())
        retained.chmod(0o755)
        write_new(HERE / f"build-{attempt}/binary.json", {"original": descriptor(binary), "retained": descriptor(retained)})
        return {
            "binary": descriptor(binary),
            "retained_binary": descriptor(retained),
            "baseline_binary": descriptor(BASELINE_BINARY),
        }

    receipt = execute(
        HERE / f"build-{attempt}",
        ["cargo", "build", "--manifest-path", str(SOURCE / "tools/perf-baseline/Cargo.toml"), "--locked", "--offline", "--release", "--example", "pptx_owned_benchmark", "-j", "2"],
        env,
        common,
        {"plan": descriptor(plan)},
        retain_candidate,
    )
    print(f"candidate retained: {retained}", flush=True)


def capture(attempt: str) -> None:
    assert_roots()
    plan_path = HERE / "capture-plan.json"
    manifest_path = HERE / f"source-{attempt}.json"
    build_path = HERE / f"build-{attempt}/command.json"
    binary_path = HERE / "binaries/candidate"
    require(plan_path.is_file() and manifest_path.is_file() and build_path.is_file(), "freeze and build must complete before capture")
    plan = json.loads(plan_path.read_text(encoding="utf-8"))
    require(plan.get("status") == "frozen_before_capture", "capture plan is not frozen")
    build = json.loads(build_path.read_text(encoding="utf-8"))
    build_binary = json.loads((HERE / f"build-{attempt}/binary.json").read_text(encoding="utf-8"))
    require(build.get("exit_code") == 0, "candidate build receipt did not pass")
    require(build_binary["retained"] == descriptor(binary_path), "retained candidate binary changed")
    check_manifest(manifest_path)
    env = dict(os.environ)
    env.update({"LC_ALL": "C", "LANG": "C", "TZ": "UTC"})
    common = {
        "cwd": str(SOURCE),
        "environment_overrides": {"LC_ALL": "C", "LANG": "C", "TZ": "UTC"},
        "removed_environment_keys": [],
        "source": descriptor(manifest_path),
        "helper": descriptor(HERE / "run.py"),
        "plan": descriptor(plan_path),
        "candidate_binary": build_binary["retained"],
        "baseline_binary": descriptor(BASELINE_BINARY),
    }
    for leg in plan["legs"]:
        folder = HERE / "capture" / str(leg["label"])
        folder.mkdir(parents=True, exist_ok=False)
        variant_binary = BASELINE_BINARY if leg["variant"] == "baseline" else binary_path
        argv = [
            "/usr/bin/time", "-f",
            "max_rss_kib=%M\nuser_seconds=%U\nsystem_seconds=%S\nelapsed_seconds=%e\nmajor_faults=%F\nminor_faults=%R\nvoluntary_context_switches=%w\ninvoluntary_context_switches=%c",
            "-o", str(folder / "process.txt"),
            "taskset", "-c", str(CPU), str(variant_binary), str(leg["route"]), str(WARMUP), str(SAMPLES), str(folder / "artifacts"),
        ]
        execute(folder, argv, env, common, {"leg": leg, "variant": leg["variant"], "binary": descriptor(variant_binary), "plan": descriptor(plan_path)})


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("mode", choices=("freeze", "build", "capture"))
    parser.add_argument("--attempt", default="v1")
    parser.add_argument("--overlay", action="append", dest="overlays", help="workspace-relative candidate overlay; repeat for each file")
    args = parser.parse_args()
    require(args.attempt.replace("-", "").isalnum(), "invalid attempt")
    try:
        if args.mode == "freeze":
            freeze(args.attempt, args.overlays)
        elif args.mode == "build":
            build(args.attempt)
        else:
            capture(args.attempt)
    except (EvidenceError, OSError, subprocess.CalledProcessError, json.JSONDecodeError) as error:
        print(f"run.py: error: {error}", file=os.sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
