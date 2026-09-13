#!/usr/bin/env python3
"""Source-bound driver for the supplemental XLSX cap-boundary guard.

The main 0544 runner owns the source-manifest and receipt protocol.  This
driver imports it with a rebound evidence directory, while retaining the same
repository and target tree.  The cap-boundary example is built separately and
its normal binary is retained under the cap lane's child of the main target.
No allocator or profiler lane is defined here: this is a bounded valid-path
cliff guard whose only admission rule is the frozen p50/mean envelope.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import shutil
import sys


HERE = Path(__file__).resolve().parent
MAIN_BUNDLE = HERE.parent
REPO = MAIN_BUNDLE.parents[3]
TARGET = Path("/home/zhuhe/litchi-goal-0544-target")
SCRATCH = TARGET / "cap-boundary-binaries"
EXAMPLE = "perf_cap_boundary"
BINARY = SCRATCH / "{stage}-cap-boundary"
WARMUP = 10
SAMPLES = 100
CPU = 2
SIZES = (160, 164, 256)
REPEATS = (1, 2)
NATIVE_ORDER = (("baseline", 1), ("candidate", 1),
                ("candidate", 2), ("baseline", 2))
TIME_FORMAT = (
    '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,'
    '"system_seconds":%S}'
)
SCRIPT_PATH = Path(__file__).resolve()
MAIN_RUN_PATH = MAIN_BUNDLE / "run.py"

# Use the frozen main runner as the custody authority.  Rebinding HERE makes
# all manifest, receipt, plan and artifact paths local to this subbundle;
# REPO and TARGET remain the main campaign's exact values.
sys.dont_write_bytecode = True
sys.path.insert(0, str(MAIN_BUNDLE))
import run as main_run  # noqa: E402  (the main runner is the custody authority)

main_run.HERE = HERE
main_run.REPO = REPO
main_run.TARGET = TARGET
main_run.SCRATCH = SCRATCH
run = main_run


def sha(path: Path) -> str:
    return run.sha(path)


def read_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, value) -> None:
    run.write(path, value)


def plan_data():
    return run.plan_data()


def run_bound(stage: str, name: str, command: list[str], binary: Path | None = None,
              *, retained_baseline: bool = False) -> None:
    """Run through the main custody protocol and bind this wrapper too."""

    run.run(stage, name, command, binary, retained_baseline=retained_baseline)
    receipt_path = HERE / stage / f"{name}.receipt.json"
    receipt = read_json(receipt_path)
    receipt["cap_driver_sha256"] = sha(SCRIPT_PATH)
    write_json(receipt_path, receipt)


def stage_binary(stage: str) -> Path:
    if stage not in ("baseline", "candidate"):
        raise ValueError(f"unsupported stage {stage}")
    return Path(str(BINARY).format(stage=stage))


def build_command() -> list[str]:
    return [
        "env",
        "CARGO_BUILD_JOBS=2",
        "CARGO_INCREMENTAL=0",
        "cargo",
        "build",
        "--release",
        "--locked",
        "-p",
        "litchi-xlsx",
        "--all-features",
        "--example",
        EXAMPLE,
        "--target-dir",
        str(TARGET),
    ]


def job_name(stage: str, repeat: int, size: int) -> str:
    if stage not in ("baseline", "candidate"):
        raise ValueError(f"unsupported stage {stage}")
    if repeat not in REPEATS:
        raise ValueError(f"unsupported repeat {repeat}")
    if size not in SIZES:
        raise ValueError(f"unsupported cap-boundary size {size}")
    return f"cap-native-r{repeat}-{size}"


def ordered_sizes(repeat: int) -> tuple[int, ...]:
    if repeat == 1:
        return SIZES
    if repeat == 2:
        return tuple(reversed(SIZES))
    raise ValueError(f"unsupported repeat {repeat}")


def expected_jobs() -> list[tuple[str, int, int]]:
    return [
        (job_name(stage, repeat, size), repeat, size)
        for stage, repeat in NATIVE_ORDER
        for size in ordered_sizes(repeat)
    ]


def _has_receipt(stage: str, repeat: int, size: int) -> bool:
    return (HERE / stage / f"{job_name(stage, repeat, size)}.receipt.json").is_file()


def assert_capture_slot(stage: str, repeat: int) -> None:
    target = (stage, repeat)
    try:
        target_index = NATIVE_ORDER.index(target)
    except ValueError as error:
        raise ValueError(f"unsupported capture slot {target}") from error
    for index, (previous_stage, previous_repeat) in enumerate(NATIVE_ORDER):
        complete = all(
            _has_receipt(previous_stage, previous_repeat, size)
            for size in SIZES
        )
        if index < target_index:
            if not complete:
                raise RuntimeError(
                    "capture order is incomplete before "
                    f"{stage} repeat {repeat}: "
                    f"{previous_stage} repeat {previous_repeat}"
                )
        elif index == target_index:
            if complete:
                raise RuntimeError(
                    f"capture slot {stage} repeat {repeat} is already complete"
                )
        elif complete:
            raise RuntimeError(
                f"capture order advanced past {stage} repeat {repeat}"
            )


def freeze(stage: str) -> None:
    if not (REPO / "crates/litchi-xlsx/examples" / f"{EXAMPLE}.rs").is_file():
        raise FileNotFoundError(
            f"{EXAMPLE}.rs must exist before the cap-boundary source freeze"
        )
    run.freeze(stage)


def build(stage: str) -> None:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    name = "build-cap-boundary"
    run_bound(stage, name, build_command())
    source = TARGET / "release" / "examples" / EXAMPLE
    if not source.is_file():
        raise FileNotFoundError(f"cap-boundary example binary is missing: {source}")
    binary = stage_binary(stage)
    shutil.copy2(source, binary)
    if sha(source) != sha(binary):
        raise RuntimeError("cap-boundary binary copy changed bytes")
    write_json(HERE / stage / "binary-cap-boundary.json", {
        "path": str(binary),
        "sha256": sha(binary),
        "bytes": binary.stat().st_size,
        "build_receipt_sha256": sha(HERE / stage / f"{name}.receipt.json"),
        "source_manifest_sha256": sha(HERE / stage / "source-manifest.json"),
    })


def capture(stage: str, repeat: int) -> None:
    plan = plan_data()
    if plan.get("status") == "draft":
        raise RuntimeError("cap-boundary plan must be frozen before capture")
    assert_capture_slot(stage, repeat)
    identity = read_json(HERE / stage / "binary-cap-boundary.json")
    binary = stage_binary(stage)
    if sha(binary) != identity.get("sha256"):
        raise RuntimeError(f"{stage} cap-boundary binary identity differs")
    for size in ordered_sizes(repeat):
        name = job_name(stage, repeat, size)
        report = HERE / stage / f"{name}.json"
        fixture = HERE / stage / f"{name}.fixture.bin"
        rss = HERE / stage / f"{name}.rss.json"
        command = [
            "taskset", "-c", str(plan.get("cpu", CPU)),
            "/usr/bin/time", "-f", TIME_FORMAT, "-o", str(rss),
            str(binary),
            "--size", str(size),
            "--warmup", str(WARMUP),
            "--samples", str(SAMPLES),
            "--json", str(report),
            "--fixture-out", str(fixture),
        ]
        run_bound(stage, name, command, binary,
                  retained_baseline=(stage == "baseline"))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage", choices=("baseline", "candidate"))
    parser.add_argument("action", choices=("freeze", "build", "capture"))
    parser.add_argument("repeat", nargs="?", type=int,
                        help="capture repeat (1 or 2); required for capture")
    args = parser.parse_args()
    if args.action == "freeze":
        if args.repeat is not None:
            parser.error("freeze does not take a repeat")
        freeze(args.stage)
    elif args.action == "build":
        if args.repeat is not None:
            parser.error("build does not take a repeat")
        build(args.stage)
    else:
        if args.repeat is None:
            parser.error("capture requires repeat 1 or 2")
        capture(args.stage, args.repeat)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
