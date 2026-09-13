"""Serial, source-bound XLSX merge build and measurement driver.

This driver deliberately keeps the benchmark implementation in the Rust
binary.  It only freezes stage identities, builds one retained binary at a
time, launches one case/shape child at a time, and records custody for every
child.  The output stage and the live execution stage are separate so the
ABBA baseline-R2 leg can run the retained baseline binary while checking the
candidate source manifest.
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
import sys
import time


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
TARGET = Path("/home/zhuhe/litchi-goal-0557-target")
SCRATCH_ROOT = TARGET / "retained"
LOCK_BINDING = HERE / "workspace-lock.json"
LOCK_COPY = HERE / "workspace-Cargo.lock"
PLAN = HERE / "plan.json"
STAGES = ("baseline", "candidate", "final")

STAGE = "baseline"
EXECUTION_STAGE = "baseline"
FOLDER = HERE / STAGE
SCRATCH = SCRATCH_ROOT / STAGE

TIME_FORMAT = (
    '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,'
    '"system_seconds":%S}'
)
ENVIRONMENT_KEYS = (
    "TMPDIR",
    "CARGO_TARGET_DIR",
    "CARGO_BUILD_JOBS",
    "CARGO_INCREMENTAL",
    "RUSTFLAGS",
    "CARGO_ENCODED_RUSTFLAGS",
    "RUSTDOCFLAGS",
    "LD_PRELOAD",
    "MALLOC_CONF",
    "GLIBC_TUNABLES",
)


def configure(stage: str, execution_stage: str | None = None) -> None:
    """Select the evidence folder and the source identity used by a child."""
    global STAGE, EXECUTION_STAGE, FOLDER, SCRATCH
    if stage not in STAGES:
        raise ValueError(f"unknown output stage: {stage}")
    EXECUTION_STAGE = execution_stage or stage
    if EXECUTION_STAGE not in STAGES:
        raise ValueError(f"unknown execution stage: {EXECUTION_STAGE}")
    STAGE = stage
    FOLDER = HERE / STAGE
    SCRATCH = SCRATCH_ROOT / STAGE


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def write_json_exclusive(path: Path, value: object) -> None:
    """Write an evidence object without ever replacing an existing file."""
    with path.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, indent=2, sort_keys=True)
        stream.write("\n")


def write_bytes_exclusive(path: Path, value: bytes) -> None:
    with path.open("xb") as stream:
        stream.write(value)


def now() -> str:
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def plan() -> dict:
    with PLAN.open(encoding="utf-8") as stream:
        value = json.load(stream)
    if not isinstance(value, dict):
        raise ValueError("plan.json must contain an object")
    return value


def lock_identity() -> dict:
    """Validate both the ignored workspace lock and its retained copy."""
    binding = json.loads(LOCK_BINDING.read_text(encoding="utf-8"))
    if set(binding) != {"path", "sha256", "scope"}:
        raise ValueError("workspace-lock.json has an unexpected schema")
    if binding["path"] != "Cargo.lock":
        raise ValueError("workspace-lock.json names the wrong lock")
    expected = binding["sha256"]
    if not isinstance(expected, str) or len(expected) != 64:
        raise ValueError("workspace-lock.json has an invalid digest")
    repo_digest = sha(REPO / "Cargo.lock")
    copy_digest = sha(LOCK_COPY)
    if repo_digest != expected or copy_digest != expected:
        raise ValueError("workspace Cargo.lock identity changed")
    return {
        "workspace_lock_sha256": repo_digest,
        "workspace_lock_copy_sha256": copy_digest,
        "workspace_lock_binding_sha256": sha(LOCK_BINDING),
        "workspace_lock_binding": binding,
    }


def ensure_lock_files() -> dict:
    """Create the retained ignored-lock evidence once, then validate it."""
    repo_lock = REPO / "Cargo.lock"
    if not repo_lock.is_file():
        raise FileNotFoundError(repo_lock)
    digest = sha(repo_lock)
    if not LOCK_BINDING.exists():
        write_json_exclusive(
            LOCK_BINDING,
            {
                "path": "Cargo.lock",
                "sha256": digest,
                "scope": "Ignored workspace lock additionally bound before and after every 0557 child; the harness lock is in the source manifest",
            },
        )
    if not LOCK_COPY.exists():
        write_bytes_exclusive(LOCK_COPY, repo_lock.read_bytes())
    return lock_identity()


def source_paths() -> list[str]:
    tracked = subprocess.check_output(
        [
            "git",
            "ls-files",
            "-z",
            "crates",
            "tools/perf-baseline",
            "Cargo.toml",
            "Cargo.lock",
            ".cargo",
            "rust-toolchain.toml",
        ],
        cwd=REPO,
    ).split(b"\0")
    untracked = subprocess.check_output(
        [
            "git",
            "ls-files",
            "--others",
            "--exclude-standard",
            "-z",
            "--",
            "crates",
            "tools/perf-baseline",
        ],
        cwd=REPO,
    ).split(b"\0")
    names = {item.decode() for item in tracked if item}
    names.update(item.decode() for item in untracked if item)
    return sorted(name for name in names if (REPO / name).is_file())


def freeze() -> None:
    """Freeze one exact source stage before its first build or child."""
    selected_plan = plan()
    expected_revision = selected_plan.get("revision")
    actual_revision = subprocess.check_output(
        ["git", "rev-parse", "HEAD"], cwd=REPO, text=True
    ).strip()
    if actual_revision != expected_revision:
        raise RuntimeError(
            f"HEAD {actual_revision} differs from prospective base {expected_revision}"
        )
    ensure_lock_files()
    if FOLDER.exists():
        raise FileExistsError(FOLDER)
    TARGET.mkdir(parents=True, exist_ok=True)
    (TARGET / "tmp").mkdir(parents=True, exist_ok=True)
    FOLDER.mkdir()
    SCRATCH.mkdir(parents=True, exist_ok=False)
    manifest = {name: sha(REPO / name) for name in source_paths()}
    write_json_exclusive(FOLDER / "source-manifest.json", manifest)
    patch = subprocess.check_output(
        ["git", "diff", "--", "crates", "tools/perf-baseline"], cwd=REPO
    )
    write_bytes_exclusive(FOLDER / "source.patch", patch)
    print(
        "Frozen",
        STAGE,
        "manifest",
        sha(FOLDER / "source-manifest.json"),
        "lock",
        lock_identity()["workspace_lock_sha256"],
        flush=True,
    )


def manifest_identity(stage: str) -> dict:
    manifest_path = HERE / stage / "source-manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if not isinstance(manifest, dict) or not manifest:
        raise ValueError(f"{stage} source manifest is empty or malformed")
    for name, expected in manifest.items():
        if not isinstance(name, str) or not isinstance(expected, str):
            raise ValueError(f"{stage} source manifest has a malformed entry")
        if len(expected) != 64:
            raise ValueError(f"{stage} source manifest has an invalid digest")
    return {
        "stage": stage,
        "manifest_path": str(manifest_path),
        "manifest_sha256": sha(manifest_path),
        "source_file_count": len(manifest),
        "locks": lock_identity(),
    }


def validate_live_manifest(
    stage: str, manifest: dict[str, str], observed: dict[str, str]
) -> None:
    """Purely validate a live source map against a frozen stage manifest."""
    if set(observed) != set(manifest):
        raise ValueError(f"{stage} source path set changed")
    for name, expected in manifest.items():
        if observed.get(name) != expected:
            raise ValueError(f"{stage} source changed: {name}")


def output_stage_identity(stage: str) -> dict:
    """Read the retained output manifest without comparing it to live source."""
    return {**manifest_identity(stage), "guard_scope": "manifest-only"}


def execution_stage_identity(stage: str) -> dict:
    """Validate the execution manifest and every currently visible source file."""
    identity = manifest_identity(stage)
    manifest = json.loads(
        (HERE / stage / "source-manifest.json").read_text(encoding="utf-8")
    )
    observed = {name: sha(REPO / name) for name in source_paths()}
    validate_live_manifest(stage, manifest, observed)
    return {**identity, "guard_scope": "live-source"}


def guarded_identity(stage: str, *, verify_live: bool) -> dict:
    """Return a serializable guard result instead of losing a later receipt."""
    try:
        identity = (
            execution_stage_identity(stage)
            if verify_live
            else output_stage_identity(stage)
        )
        return {"ok": True, "identity": identity}
    except Exception as error:  # the error is part of the evidence
        return {
            "ok": False,
            "error": f"{type(error).__name__}: {error}",
        }


def source_guard_pure_test(output_stage: str, execution_stage: str) -> None:
    """Prove the asymmetric ABBA guard contract without mutating source."""
    output = output_stage_identity(output_stage)
    execution = execution_stage_identity(execution_stage)
    if output["guard_scope"] != "manifest-only":
        raise AssertionError("output guard unexpectedly checks live source")
    if execution["guard_scope"] != "live-source":
        raise AssertionError("execution guard does not check live source")
    manifest = json.loads(
        (HERE / execution_stage / "source-manifest.json").read_text(
            encoding="utf-8"
        )
    )
    observed = {name: sha(REPO / name) for name in source_paths()}
    if not observed:
        raise AssertionError("execution source map is empty")
    mutated = dict(observed)
    first_name = next(iter(mutated))
    mutated[first_name] = "0" * 64
    try:
        validate_live_manifest(execution_stage, manifest, mutated)
    except ValueError:
        print(
            "source guard pure test passed:",
            f"{output_stage} output + {execution_stage} execution;",
            "synthetic execution mutation refused",
            flush=True,
        )
    else:
        raise AssertionError("synthetic execution mutation was accepted")


def compiler_observations() -> list[dict]:
    observations = []
    for proc in Path("/proc").iterdir():
        if not proc.name.isdigit():
            continue
        try:
            comm = (proc / "comm").read_text(encoding="utf-8").strip()
            if comm in ("cargo", "rustc", "rustdoc"):
                observations.append(
                    {
                        "pid": int(proc.name),
                        "comm": comm,
                        "cwd": os.readlink(proc / "cwd"),
                    }
                )
        except OSError:
            pass
    return sorted(observations, key=lambda item: item["pid"])


def command_environment() -> dict[str, str]:
    environment = os.environ.copy()
    environment.update(
        {
            "TMPDIR": str(TARGET / "tmp"),
            "CARGO_TARGET_DIR": str(TARGET),
            "CARGO_BUILD_JOBS": "2",
            "CARGO_INCREMENTAL": "0",
        }
    )
    return environment


def reserve_paths(paths: list[Path]) -> None:
    for path in paths:
        if path.exists() or path.is_symlink():
            raise FileExistsError(f"refusing to overwrite evidence path {path}")


def run_child(
    name: str,
    command: list[str],
    *,
    binary: Path | None = None,
    binary_descriptor: Path | None = None,
    artifacts: list[Path] | None = None,
    environment_overrides: dict[str, str] | None = None,
    allow_failure: bool = False,
) -> dict:
    """Run one child and write its receipt even if the post-run guard fails.

    The receipt is written before a non-zero result or identity failure is
    raised.  This ordering is intentional: source mutation during a child is
    itself valuable evidence and must never erase the failed-run record.
    """
    if not name or any(character in name for character in "/\\"):
        raise ValueError(f"invalid child name: {name!r}")
    if not FOLDER.is_dir():
        raise FileNotFoundError(f"stage is not frozen: {FOLDER}")
    artifact_paths = list(artifacts or [])
    receipt_path = FOLDER / f"{name}.receipt.json"
    host_path = FOLDER / f"{name}.host.json"
    stdout_path = FOLDER / f"{name}.stdout"
    stderr_path = FOLDER / f"{name}.stderr"
    reserve_paths([receipt_path, host_path, stdout_path, stderr_path, *artifact_paths])

    before_output = guarded_identity(STAGE, verify_live=False)
    before_execution = guarded_identity(EXECUTION_STAGE, verify_live=True)
    binary_before = None
    binary_error = None
    if binary is not None:
        try:
            binary_before = sha(binary)
        except Exception as error:  # retain a receipt for a missing binary
            binary_error = f"{type(error).__name__}: {error}"
    descriptor_before = None
    descriptor_error = None
    if binary_descriptor is not None:
        try:
            descriptor_before = sha(binary_descriptor)
        except Exception as error:  # retain a receipt for a missing descriptor
            descriptor_error = f"{type(error).__name__}: {error}"

    write_json_exclusive(
        host_path,
        {
            "observed_utc": now(),
            "compiler_processes": compiler_observations(),
            "scope": "Accessible compiler processes only; no host quiescence guarantee",
        },
    )
    environment = command_environment()
    if environment_overrides:
        environment.update(environment_overrides)
    command = [str(item) for item in command]
    start_utc = now()
    tick = time.monotonic()
    result_code = None
    child_started = False
    spawn_error = None
    preflight_ok = bool(before_output["ok"] and before_execution["ok"])
    if binary is not None and binary_error is not None:
        preflight_ok = False
    if binary_descriptor is not None and descriptor_error is not None:
        preflight_ok = False
    try:
        with stdout_path.open("x", encoding="utf-8") as stdout, stderr_path.open(
            "x", encoding="utf-8"
        ) as stderr:
            if preflight_ok:
                child_started = True
                result = subprocess.run(
                    command,
                    cwd=REPO,
                    stdout=stdout,
                    stderr=stderr,
                    env=environment,
                    check=False,
                )
                result_code = result.returncode
            else:
                stderr.write("child refused before launch because a guard failed\n")
    except Exception as error:
        spawn_error = f"{type(error).__name__}: {error}"

    after_output = guarded_identity(STAGE, verify_live=False)
    after_execution = guarded_identity(EXECUTION_STAGE, verify_live=True)
    binary_after = None
    if binary is not None and binary_error is None:
        try:
            binary_after = sha(binary)
        except Exception:
            binary_after = None
    descriptor_after = None
    if binary_descriptor is not None and descriptor_error is None:
        try:
            descriptor_after = sha(binary_descriptor)
        except Exception:
            descriptor_after = None
    binary_unchanged = None
    if binary is not None:
        binary_unchanged = (
            binary_before is not None
            and binary_after is not None
            and binary_before == binary_after
        )
    descriptor_unchanged = None
    if binary_descriptor is not None:
        descriptor_unchanged = (
            descriptor_before is not None
            and descriptor_after is not None
            and descriptor_before == descriptor_after
        )
    output_manifest_before = before_output.get("identity", {}).get(
        "manifest_sha256"
    )
    output_manifest_after = after_output.get("identity", {}).get(
        "manifest_sha256"
    )
    output_manifest_unchanged = (
        before_output.get("ok")
        and after_output.get("ok")
        and output_manifest_before == output_manifest_after
    )
    execution_manifest_before = before_execution.get("identity", {}).get(
        "manifest_sha256"
    )
    execution_manifest_after = after_execution.get("identity", {}).get(
        "manifest_sha256"
    )
    execution_manifest_unchanged = (
        before_execution.get("ok")
        and after_execution.get("ok")
        and execution_manifest_before == execution_manifest_after
    )
    success = bool(
        preflight_ok
        and child_started
        and result_code == 0
        and after_output["ok"]
        and after_execution["ok"]
        and output_manifest_unchanged
        and execution_manifest_unchanged
        and (binary is None or binary_unchanged)
        and (binary_descriptor is None or descriptor_unchanged)
        and spawn_error is None
    )
    receipt = {
        "schema": "xlsx_linear_merge_0557_run_receipt_v1",
        "name": name,
        "command": command,
        "start_utc": start_utc,
        "end_utc": now(),
        "seconds": time.monotonic() - tick,
        "child_started": child_started,
        "exit_code": result_code,
        "success": success,
        "stage": STAGE,
        "execution_stage": EXECUTION_STAGE,
        "output_stage_before": before_output,
        "execution_stage_before": before_execution,
        "output_stage_after": after_output,
        "execution_stage_after": after_execution,
        "output_manifest_unchanged": output_manifest_unchanged,
        "execution_manifest_unchanged": execution_manifest_unchanged,
        "output_stage_manifest_sha256": (
            before_output.get("identity", {}).get("manifest_sha256")
            if before_output.get("ok")
            else None
        ),
        "execution_stage_manifest_sha256": (
            before_execution.get("identity", {}).get("manifest_sha256")
            if before_execution.get("ok")
            else None
        ),
        "source_manifest_sha256": (
            before_output.get("identity", {}).get("manifest_sha256")
            if before_output.get("ok")
            else None
        ),
        "execution_manifest_sha256": (
            before_execution.get("identity", {}).get("manifest_sha256")
            if before_execution.get("ok")
            else None
        ),
        "workspace_lock_sha256": (
            before_output.get("identity", {})
            .get("locks", {})
            .get("workspace_lock_sha256")
            if before_output.get("ok")
            else None
        ),
        "workspace_lock_copy_sha256": (
            before_output.get("identity", {})
            .get("locks", {})
            .get("workspace_lock_copy_sha256")
            if before_output.get("ok")
            else None
        ),
        "workspace_lock_binding_sha256": (
            before_output.get("identity", {})
            .get("locks", {})
            .get("workspace_lock_binding_sha256")
            if before_output.get("ok")
            else None
        ),
        "binary_path": str(binary) if binary is not None else None,
        "binary_descriptor_path": (
            str(binary_descriptor) if binary_descriptor is not None else None
        ),
        "binary_before_sha256": binary_before,
        "binary_after_sha256": binary_after,
        "binary_sha256": binary_before,
        "binary_unchanged": binary_unchanged,
        "binary_error": binary_error,
        "binary_descriptor_before_sha256": descriptor_before,
        "binary_descriptor_after_sha256": descriptor_after,
        "binary_descriptor_unchanged": descriptor_unchanged,
        "binary_descriptor_error": descriptor_error,
        "spawn_error": spawn_error,
        "environment": {
            key: environment.get(key) for key in ENVIRONMENT_KEYS
        },
        "plan_sha256": sha(PLAN),
        "script_sha256": sha(Path(__file__)),
        "artifacts": {
            path.name: sha(path)
            for path in sorted(FOLDER.glob(f"{name}.*"))
            if path.is_file() and path != receipt_path
        },
    }
    write_json_exclusive(receipt_path, receipt)
    print(name, "exit", result_code, "success", success, flush=True)
    if not success and not allow_failure:
        raise RuntimeError(f"child {name} failed; receipt retained at {receipt_path}")
    return receipt


def copy_exclusive(source: Path, destination: Path) -> None:
    if destination.exists() or destination.is_symlink():
        raise FileExistsError(f"refusing to overwrite retained binary {destination}")
    with source.open("rb") as source_stream, destination.open("xb") as destination_stream:
        shutil.copyfileobj(source_stream, destination_stream)
    shutil.copystat(source, destination)


def build(kind: str) -> None:
    if kind not in ("normal", "alloc"):
        raise ValueError(kind)
    executable = "litchi-perf-baseline" + ("-alloc" if kind == "alloc" else "")
    command = [
        "env",
        f"TMPDIR={TARGET / 'tmp'}",
        "CARGO_BUILD_JOBS=2",
        "CARGO_INCREMENTAL=0",
        "cargo",
        "build",
        "--release",
        "--locked",
        "--manifest-path",
        "tools/perf-baseline/Cargo.toml",
        "--bin",
        executable,
        "--target-dir",
        str(TARGET),
    ]
    if kind == "alloc":
        command.extend(["--features", "allocator-metrics"])
    target_binary = TARGET / "release" / executable
    run_child("build-" + kind, command)
    if not target_binary.is_file():
        raise FileNotFoundError(target_binary)
    retained = SCRATCH / kind
    copy_exclusive(target_binary, retained)
    write_json_exclusive(
        FOLDER / f"binary-{kind}.json",
        {
            "schema": "xlsx_linear_merge_0557_binary_identity_v1",
            "kind": kind,
            "path": str(retained),
            "sha256": sha(retained),
            "bytes": retained.stat().st_size,
            "build_receipt_sha256": sha(FOLDER / f"build-{kind}.receipt.json"),
            "source_manifest_sha256": sha(FOLDER / "source-manifest.json"),
            "workspace_lock": lock_identity(),
            "stage": STAGE,
        },
    )
    print("Retained", kind, sha(retained), flush=True)


def workload_entries(selected_plan: dict, primary_only: bool) -> list[dict]:
    workloads = selected_plan["workloads"]
    entries = list(workloads["primary"])
    if not primary_only:
        entries += list(workloads["controls"])
    if not entries:
        raise ValueError("the plan has no workloads")
    return entries


def rows(selected_plan: dict, primary_only: bool, reverse: bool) -> list[tuple[str, str]]:
    entries = workload_entries(selected_plan, primary_only)
    output = [
        (entry["case"], shape)
        for shape in selected_plan["corpus"]["shapes"]
        for entry in entries
    ]
    if reverse:
        output.reverse()
    return output


def binary_identity(kind: str) -> tuple[Path, dict]:
    binary = SCRATCH / kind
    identity_path = FOLDER / f"binary-{kind}.json"
    identity = json.loads(identity_path.read_text(encoding="utf-8"))
    if identity.get("sha256") != sha(binary):
        raise ValueError(f"retained {kind} binary changed")
    if identity.get("source_manifest_sha256") != sha(
        FOLDER / "source-manifest.json"
    ):
        raise ValueError(f"retained {kind} binary is bound to another output stage")
    lock = identity.get("workspace_lock", {})
    current_lock = lock_identity()
    for field in (
        "workspace_lock_sha256",
        "workspace_lock_copy_sha256",
        "workspace_lock_binding_sha256",
    ):
        if lock.get(field) != current_lock[field]:
            raise ValueError(f"retained {kind} binary lock identity changed")
    return binary, identity


def lane_configuration(selected_plan: dict, lane: str) -> dict:
    timing = selected_plan["timing"]
    if lane == "noise":
        return timing["noise"]
    if lane == "native":
        return timing["normal"]
    if lane == "alloc":
        return timing["allocation"]
    raise ValueError(lane)


def capture(lane: str, repeat: int | None = None) -> None:
    if lane not in ("noise", "native", "alloc"):
        raise ValueError(lane)
    selected_plan = plan()
    config = lane_configuration(selected_plan, lane)
    if lane == "noise":
        if STAGE != "baseline" or EXECUTION_STAGE != "baseline":
            raise ValueError("baseline noise is restricted to baseline/baseline")
        primary_only = True
        kind = "normal"
    else:
        primary_only = False
        kind = "alloc" if lane == "alloc" else "normal"
    repeats = int(config["repeats"])
    if repeat is not None and not 1 <= repeat <= repeats:
        raise ValueError(f"repeat must be within 1..{repeats}")
    selected_repeats = [repeat] if repeat is not None else list(range(1, repeats + 1))
    binary = None
    if lane != "noise" or STAGE == "baseline":
        binary, _ = binary_identity(kind)
    for current_repeat in selected_repeats:
        reverse = lane != "noise" and current_repeat == 2
        for case, shape in rows(selected_plan, primary_only, reverse):
            name = f"{lane}-r{current_repeat}-{shape}-{case}"
            output = FOLDER / f"{name}.json"
            catalog = FOLDER / f"{name}.catalog.json"
            rss = FOLDER / f"{name}.rss.json"
            command = [
                "taskset",
                "-c",
                str(selected_plan["timing"]["cpu"]),
                "/usr/bin/time",
                "-f",
                TIME_FORMAT,
                "-o",
                str(rss),
                str(binary),
                "--case",
                case,
                "--xlsx-cell-crud-shape",
                shape,
                "--warmup",
                str(config["warmup"]),
                "--samples",
                str(config["samples"]),
                "--json",
                str(output),
                "--corpus-manifest",
                str(catalog),
            ]
            run_child(
                name,
                command,
                binary=binary,
                binary_descriptor=FOLDER / f"binary-{kind}.json",
                artifacts=[output, catalog, rss],
            )


def run_check(
    name: str, command: list[str], environment_overrides: dict[str, str]
) -> None:
    if not command:
        raise ValueError("check requires an argv after --")
    if command[0] == "--":
        command = command[1:]
    if not command:
        raise ValueError("check requires an argv after --")
    run_child(name, command, environment_overrides=environment_overrides)


def parse_args() -> argparse.Namespace:
    raw_arguments = sys.argv[1:]
    command = []
    if "check" in raw_arguments and "--" in raw_arguments:
        separator = raw_arguments.index("--")
        command = raw_arguments[separator + 1 :]
        raw_arguments = raw_arguments[:separator]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "action",
        choices=[
            "freeze",
            "guard-test",
            "build-normal",
            "build-alloc",
            "noise",
            "native",
            "alloc",
            "check",
        ],
    )
    parser.add_argument("--stage", choices=STAGES, default="baseline")
    parser.add_argument("--execution-stage", choices=STAGES)
    parser.add_argument("--repeat", type=int, choices=[1, 2])
    parser.add_argument("--name", help="receipt name for the check action")
    parser.add_argument(
        "--env",
        action="append",
        default=[],
        metavar="KEY=VALUE",
        help="environment override for check (repeatable)",
    )
    parsed = parser.parse_args(raw_arguments)
    parsed.command = command
    return parsed


def main() -> None:
    args = parse_args()
    execution_stage = args.execution_stage
    if args.action == "guard-test" and execution_stage is None:
        execution_stage = "candidate"
    configure(args.stage, execution_stage)
    if args.action == "freeze":
        if args.repeat is not None:
            raise ValueError("freeze does not accept --repeat")
        freeze()
    elif args.action == "guard-test":
        if args.repeat is not None:
            raise ValueError("guard-test does not accept --repeat")
        source_guard_pure_test(args.stage, execution_stage)
    elif args.action.startswith("build-"):
        if args.repeat is not None:
            raise ValueError("build does not accept --repeat")
        build(args.action[6:])
    elif args.action == "check":
        if not args.name:
            raise ValueError("check requires --name")
        overrides = {}
        for assignment in args.env:
            key, separator, value = assignment.partition("=")
            if not separator or not key or "=" in key:
                raise ValueError(f"--env requires KEY=VALUE: {assignment!r}")
            overrides[key] = value
        run_check(args.name, args.command, overrides)
    else:
        capture(args.action, args.repeat)


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(f"0557 driver failed: {error}", file=sys.stderr)
        raise
