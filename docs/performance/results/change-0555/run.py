"""Serial source-bound matched CFB/OLE2 build and measurement driver.

The driver owns only child execution and receipt creation.  It keeps the
output stage (where a binary or capture is retained) separate from the live
execution stage so the final ABBA baseline-r2 leg can use the retained
baseline binary while binding the candidate source manifest explicitly.
"""

import argparse
import datetime
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import time


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
TARGET = Path("/home/zhuhe/litchi-goal-0555-target")
SCRATCH_ROOT = TARGET / "retained"
STAGE = "baseline"
EXECUTION_STAGE = "baseline"
FOLDER = HERE / STAGE
SCRATCH = SCRATCH_ROOT / STAGE
LOCK_BINDING = HERE / "workspace-lock.json"
LOCK_COPY = HERE / "workspace-Cargo.lock"


def configure(stage, execution_stage=None):
    global STAGE, EXECUTION_STAGE, FOLDER, SCRATCH
    assert stage in ("baseline", "candidate", "final")
    EXECUTION_STAGE = execution_stage or stage
    assert EXECUTION_STAGE in ("baseline", "candidate", "final")
    STAGE = stage
    FOLDER = HERE / stage
    SCRATCH = SCRATCH_ROOT / stage


def sha(path):
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def write(path, value):
    with path.open("x") as stream:
        json.dump(value, stream, indent=2)
        stream.write("\n")


def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def lock_digest():
    """Validate the separately ignored workspace lock before every child."""
    binding = json.loads(LOCK_BINDING.read_text())
    assert set(binding) == {"path", "sha256", "scope"}
    assert binding["path"] == "Cargo.lock"
    digest = binding["sha256"]
    assert isinstance(digest, str) and len(digest) == 64
    assert sha(REPO / "Cargo.lock") == digest
    assert sha(LOCK_COPY) == digest
    return digest


def source_paths():
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
    names = {name.decode() for name in tracked if name}
    names |= {name.decode() for name in untracked if name.endswith(b".rs")}
    return sorted(name for name in names if (REPO / name).is_file())


def freeze():
    """Freeze the exact source tree used by this stage before its first child."""
    plan = json.loads((HERE / "plan.json").read_text())
    assert subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=REPO,
                                  text=True).strip() == plan["revision"]
    lock_digest()
    FOLDER.mkdir(exist_ok=False)
    SCRATCH.mkdir(parents=True, exist_ok=True)
    manifest = {name: sha(REPO / name) for name in source_paths()}
    write(FOLDER / "source-manifest.json", manifest)
    patch = subprocess.check_output(
        ["git", "diff", "--", "crates", "tools/perf-baseline"], cwd=REPO
    )
    (FOLDER / "source.patch").write_bytes(patch)
    print("Frozen", STAGE, "manifest", sha(FOLDER / "source-manifest.json"),
          "lock", lock_digest(), flush=True)


def check_source():
    """Fail closed if source or either lock identity changed during a child."""
    lock_digest()
    manifest = json.loads(
        (HERE / EXECUTION_STAGE / "source-manifest.json").read_text()
    )
    for name, digest in manifest.items():
        assert sha(REPO / name) == digest, name


def compiler_observations():
    observations = []
    for proc in Path("/proc").iterdir():
        if not proc.name.isdigit():
            continue
        try:
            comm = (proc / "comm").read_text().strip()
            if comm in ("cargo", "rustc"):
                observations.append(
                    dict(pid=int(proc.name), comm=comm,
                         cwd=os.readlink(proc / "cwd"))
                )
        except OSError:
            pass
    return observations


def run(name, command, binary=None, allow_failure=False):
    """Run one serial child and bind all source, lock, plan, and driver IDs."""
    receipt_path = FOLDER / (name + ".receipt.json")
    assert not receipt_path.exists(), receipt_path
    check_source()
    before = sha(binary) if binary else None
    write(
        FOLDER / (name + ".host.json"),
        dict(
            observed_utc=now(),
            compiler_processes=compiler_observations(),
            scope="Accessible compiler processes; no host quiescence guarantee",
        ),
    )
    start, tick = now(), time.monotonic()
    environment = {
        **os.environ,
        "TMPDIR": str(TARGET / "tmp"),
        "CARGO_BUILD_JOBS": "2",
        "CARGO_INCREMENTAL": "0",
    }
    with (FOLDER / (name + ".stdout")).open("x") as out, \
            (FOLDER / (name + ".stderr")).open("x") as err:
        result = subprocess.run(command, cwd=REPO, stdout=out, stderr=err,
                                env=environment)
    check_source()
    assert binary is None or sha(binary) == before
    stage_manifest = HERE / STAGE / "source-manifest.json"
    execution_manifest = HERE / EXECUTION_STAGE / "source-manifest.json"
    receipt = dict(
        schema="ole2_0555_run_receipt_v1",
        command=command,
        start_utc=start,
        end_utc=now(),
        seconds=time.monotonic() - tick,
        exit_code=result.returncode,
        stage=STAGE,
        execution_stage=EXECUTION_STAGE,
        execution_manifest_sha256=sha(execution_manifest),
        binary_sha256=before,
        source_manifest_sha256=sha(stage_manifest),
        workspace_lock_sha256=lock_digest(),
        workspace_lock_binding_sha256=sha(LOCK_BINDING),
        script_sha256=sha(Path(__file__)),
        plan_sha256=sha(HERE / "plan.json"),
        environment={
            key: environment.get(key)
            for key in (
                "TMPDIR",
                "CARGO_TARGET_DIR",
                "CARGO_BUILD_JOBS",
                "CARGO_INCREMENTAL",
                "RUSTFLAGS",
                "CARGO_ENCODED_RUSTFLAGS",
                "LD_PRELOAD",
                "MALLOC_CONF",
                "GLIBC_TUNABLES",
            )
        },
    )
    receipt_path_artifacts = {
        p.name: sha(p)
        for p in sorted(FOLDER.glob(name + ".*"))
        if p.is_file() and p != receipt_path
    }
    receipt["artifacts"] = receipt_path_artifacts
    write(receipt_path, receipt)
    assert allow_failure or result.returncode == 0, (name, result.returncode)
    print(name, "exit", result.returncode, flush=True)


def build(kind):
    assert kind in ("normal", "alloc")
    executable = "litchi-perf-baseline" + ("-alloc" if kind == "alloc" else "")
    command = [
        "env",
        "TMPDIR=" + str(TARGET / "tmp"),
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
        command += ["--features", "allocator-metrics"]
    run("build-" + kind, command)
    source = TARGET / "release" / executable
    binary = SCRATCH / kind
    shutil.copy2(source, binary)
    assert sha(source) == sha(binary)
    write(
        FOLDER / ("binary-" + kind + ".json"),
        dict(
            path=str(binary),
            sha256=sha(binary),
            bytes=binary.stat().st_size,
            build_receipt_sha256=sha(FOLDER / ("build-" + kind + ".receipt.json")),
            source_manifest_sha256=sha(FOLDER / "source-manifest.json"),
            workspace_lock_sha256=lock_digest(),
        ),
    )


def jobs(lane):
    plan = json.loads((HERE / "plan.json").read_text())
    config = plan["allocation" if lane == "alloc" else lane]
    output = []
    for repeat in range(1, config["repeats"] + 1):
        if lane == "profile":
            selections = [("xls-owned", {
                "cases": ["xls_owned_source_open_one_cell"]
            })]
            selections += [
                ("cfb-" + shape, {
                    "cases": ["cfb_open"],
                    "shapes": [shape],
                    "payload": plan["groups"]["cfb"]["payload"],
                })
                for shape in plan["groups"]["cfb"]["shapes"]
            ]
        else:
            names = ["xls", "cfb"] if repeat == 1 else ["cfb", "xls"]
            selections = [(name, plan["groups"][name]) for name in names]
        for group, selection in selections:
            output.append(
                dict(
                    name=f"{lane}-r{repeat}-{group}",
                    repeat=repeat,
                    group=group,
                    selection=selection,
                    samples=config["samples"],
                    warmup=config["warmup"],
                )
            )
    return output


def capture(lane, repeat=None):
    assert lane in ("native", "alloc", "profile")
    plan = json.loads((HERE / "plan.json").read_text())
    kind = "alloc" if lane == "alloc" else "normal"
    binary = SCRATCH / kind
    binary_record = json.loads((FOLDER / ("binary-" + kind + ".json")).read_text())
    assert sha(binary) == binary_record["sha256"]
    assert binary_record["workspace_lock_sha256"] == lock_digest()
    for job in jobs(lane):
        if repeat is not None and job["repeat"] != repeat:
            continue
        name, selection = job["name"], job["selection"]
        command = ["taskset", "-c", str(plan["cpu"])]
        if lane == "native":
            command += [
                "/usr/bin/time",
                "-f",
                '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}',
                "-o",
                str(FOLDER / (name + ".rss.json")),
            ]
        elif lane == "profile":
            owner = plan["profile"][
                "cfb_owner" if job["group"].startswith("cfb-") else "xls_owner"
            ]
            command += [
                "valgrind",
                "--vgdb=no",
                "--tool=callgrind",
                "--collect-atstart=no",
                "--toggle-collect=" + owner,
                "--zero-before=" + owner,
                "--dump-after=" + owner,
                "--callgrind-out-file=" + str(FOLDER / (name + ".callgrind")),
            ] + plan["profile"]["instruction_flags"]
        command += [
            str(binary),
            "--case",
            ",".join(selection["cases"]),
            "--warmup",
            str(job["warmup"]),
            "--samples",
            str(job["samples"]),
            "--json",
            str(FOLDER / (name + ".json")),
            "--corpus-manifest",
            str(FOLDER / (name + ".catalog.json")),
        ]
        if "shapes" in selection:
            command += [
                "--shape",
                ",".join(selection["shapes"]),
                "--payload",
                selection["payload"],
            ]
        run(name, command, binary)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "action",
        choices=["freeze", "build-normal", "build-alloc", "native", "profile", "alloc"],
    )
    parser.add_argument("--stage", choices=["baseline", "candidate", "final"],
                        default="baseline")
    parser.add_argument("--execution-stage", choices=["baseline", "candidate", "final"])
    parser.add_argument("--repeat", type=int, choices=[1, 2])
    args = parser.parse_args()
    configure(args.stage, args.execution_stage)
    if args.action == "freeze":
        freeze()
    elif args.action.startswith("build-"):
        build(args.action[6:])
    else:
        capture(args.action, args.repeat)
