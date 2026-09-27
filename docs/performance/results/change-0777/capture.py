"""Reproduce the 0777 XML attribute integration capture.

The runner deliberately keeps the two library builds in separate target
directories and executes every process in one serial order.  It writes a new
``capture-0`` directory and refuses to reuse an existing one.  The candidate
revision is read from the coordinator's ``final-source.json`` so a capture
cannot silently measure a different checkout.

This script is a capture description and runner.  It does not contain any
production implementation and must be run from the 0777 worktree by the root
coordinator after the candidate source has been frozen.
"""

from __future__ import annotations

import hashlib
import json
import os
import platform
import shutil
import subprocess
import sys
import time
import tomllib
from pathlib import Path
from typing import Any, Iterable


PACKET = Path(__file__).resolve().parent
AFTER = PACKET.parents[3]
BEFORE = Path("/home/zhuhe/code/litchi")
CAPTURE = PACKET / "capture-0"
BASE_REF = "87e926fcc6e71360073a8bf97ad572ad9d187a67"
TARGETS = {
    "before": Path("/home/zhuhe/code/litchi-target-0777-release-before"),
    "after": Path("/home/zhuhe/code/litchi-target-0777-release-after"),
}
CPU = 12
SAMPLES = 9
WARMUP = 2
MUTATE_LIMIT = 1_048_576
LEGS = ("before", "after", "after", "before", "before", "after")

ROOT_RELATIVE = (Path("test-data"), Path("docs/performance/results"))
PROBE_INPUTS = {
    "attribute_checks.rs": Path("probe-src/attribute_checks.rs"),
    "attribute_checks_equivalence.rs": Path("probe-src/attribute_checks_equivalence.rs"),
    "main.rs": Path("probe-src/main.rs"),
    "Cargo.toml.template": Path("probe-src/Cargo.toml.template"),
}
PROBE_ORIGINS = {
    "attribute_checks.rs": Path("tools/perf-baseline/src/bin/attribute_checks.rs"),
    "attribute_checks_equivalence.rs": Path(
        "tools/perf-baseline/src/bin/attribute_checks_equivalence.rs"
    ),
    "main.rs": Path("docs/performance/results/change-0776/probe-src/main.rs"),
    "Cargo.toml.template": Path(
        "docs/performance/results/change-0777/probe-src/Cargo.toml.template"
    ),
}
PROBE_FILES = tuple(PROBE_INPUTS)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def json_write(path: Path, value: Any) -> None:
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n")


def json_read(path: Path) -> Any:
    return json.loads(path.read_text())


def display_path(path: Path) -> str:
    try:
        return str(path.relative_to(PACKET))
    except ValueError:
        return str(path)


def git(root: Path, *arguments: str) -> str:
    return subprocess.check_output(
        ["git", *arguments], cwd=root, text=True
    ).strip()


def regular_files(root: Path) -> Iterable[Path]:
    """Yield regular files while matching Rust's symlink_metadata traversal.

    In particular, symlinked directories and files are excluded.  The packet
    owns no symlinked input roots, but this makes the receipt explicit and
    keeps 0777's inventory consistent with the two Rust collectors.
    """

    if root.is_symlink():
        return
    if root.is_dir():
        for child in sorted(root.iterdir()):
            yield from regular_files(child)
    elif root.is_file():
        yield root


def is_zip_candidate(path: Path) -> bool:
    try:
        with path.open("rb") as source:
            return source.read(4) == b"PK\x03\x04"
    except OSError:
        return False


def relative_after(path: Path) -> str:
    return str(path.relative_to(AFTER))


def source_manifest(root: Path, reference: str) -> dict[str, Any]:
    paths = (
        "crates",
        "tools/perf-baseline",
        "Cargo.toml",
        "clippy.toml",
        ".cargo/config.toml",
        "rust-toolchain.toml",
    )
    subprocess.run(
        ["git", "diff", "--quiet", reference, "--", *paths],
        cwd=root,
        check=True,
    )
    names = subprocess.check_output(
        ["git", "ls-files", "-z", "--", *paths], cwd=root
    ).decode().split("\0")
    files = {name: sha256(root / name) for name in names if name}
    lock = root / "Cargo.lock"
    if not lock.is_file():
        raise RuntimeError(f"missing root Cargo.lock: {lock}")
    files["Cargo.lock"] = sha256(lock)
    return {
        "root": str(root),
        "reference": reference,
        "head": git(root, "rev-parse", "HEAD"),
        "files": files,
    }


def verify_final_source(
    candidate_manifest: dict[str, Any], candidate_ref: str, source: dict[str, Any]
) -> None:
    declared_files = candidate_manifest.get("files", {})
    if not isinstance(declared_files, dict):
        raise RuntimeError("final-source.json files is not an object")
    for name, expected in declared_files.items():
        actual = source["files"].get(name)
        if actual != expected:
            raise RuntimeError(
                f"final-source.json disagrees for {name}: {actual!r} != {expected!r}"
            )
    if candidate_manifest.get("base") != BASE_REF[:10]:
        raise RuntimeError(
            f"final-source.json base is {candidate_manifest.get('base')!r}, "
            f"expected {BASE_REF[:10]!r}"
        )
    if candidate_manifest.get("candidate") != candidate_ref:
        raise RuntimeError("final-source.json candidate does not resolve to HEAD")


def fixture_manifest() -> dict[str, Any]:
    files: dict[str, Any] = {}
    packages: dict[str, Any] = {}
    for path in regular_files(AFTER / "test-data"):
        name = relative_after(path)
        entry = {"bytes": path.stat().st_size, "sha256": sha256(path)}
        files[name] = entry
        if is_zip_candidate(path):
            packages[name] = entry
    return {
        "root": "test-data",
        "symlinks_skipped": True,
        "files": files,
        "package_candidates": packages,
    }


def package_inventory() -> dict[str, Any]:
    inventory: dict[str, Any] = {}
    for root in ROOT_RELATIVE:
        absolute = AFTER / root
        for path in regular_files(absolute):
            if not is_zip_candidate(path):
                continue
            relative = relative_after(path)
            inventory[relative] = {
                "bytes": path.stat().st_size,
                "root": str(root),
                "sha256": sha256(path),
            }
    return {
        "roots": [str(root) for root in ROOT_RELATIVE],
        "symlinks_skipped": True,
        "regular_pk_candidates": inventory,
    }


def probe_input_manifest() -> dict[str, Any]:
    manifest: dict[str, Any] = {}
    for name, relative in PROBE_INPUTS.items():
        source = PACKET / relative
        if not source.is_file():
            raise RuntimeError(f"missing probe input {source}")
        origin = AFTER / PROBE_ORIGINS[name]
        if not origin.is_file():
            raise RuntimeError(f"missing probe origin {origin}")
        source_hash = sha256(source)
        origin_hash = sha256(origin)
        if source_hash != origin_hash:
            raise RuntimeError(
                f"archived probe source differs from its origin: {relative}"
            )
        manifest[name] = {
            "path": str(relative),
            "origin": str(PROBE_ORIGINS[name]),
            "bytes": source.stat().st_size,
            "sha256": source_hash,
            "origin_sha256": origin_hash,
        }
    return manifest


def rendered_probe(
    leg: str, root: Path, input_manifest: dict[str, Any]
) -> tuple[Path, dict[str, Any]]:
    probe = CAPTURE / f"probe-{leg}"
    (probe / "src/bin").mkdir(parents=True)
    copies = {
        "src/bin/attribute_checks.rs": "attribute_checks.rs",
        "src/bin/attribute_checks_equivalence.rs": "attribute_checks_equivalence.rs",
        "src/bin/mce-stream-probe.rs": "main.rs",
    }
    copied: dict[str, Any] = {}
    for destination, source_name in copies.items():
        source = PACKET / PROBE_INPUTS[source_name]
        target = probe / destination
        shutil.copy2(source, target)
        copied[destination] = {
            "source": source_name,
            "bytes": target.stat().st_size,
            "sha256": sha256(target),
        }
        if copied[destination]["sha256"] != input_manifest[source_name]["sha256"]:
            raise RuntimeError(f"copied probe source mismatch: {destination}")
    template = (PACKET / PROBE_INPUTS["Cargo.toml.template"]).read_text()
    manifest = probe / "Cargo.toml"
    manifest.write_text(template.replace("@SRC@", str(root)))
    copied["Cargo.toml"] = {
        "source": "Cargo.toml.template",
        "bytes": manifest.stat().st_size,
        "sha256": sha256(manifest),
    }
    receipt = {
        "root": str(root),
        "leg": leg,
        "files": copied,
        "template_sha256": input_manifest["Cargo.toml.template"]["sha256"],
    }
    json_write(CAPTURE / f"source-map-{leg}.json", receipt)
    return manifest, receipt


def environment(leg: str, target: Path, references: dict[str, str]) -> dict[str, Any]:
    allowed = sorted(os.sched_getaffinity(0)) if hasattr(os, "sched_getaffinity") else []
    if CPU not in allowed:
        raise RuntimeError(f"CPU {CPU} is not in the current affinity: {allowed}")
    env = os.environ.copy()
    env.update(
        {
            "CARGO_BUILD_JOBS": "2",
            "CARGO_INCREMENTAL": "0",
            "CARGO_NET_OFFLINE": "true",
            "RUSTUP_TOOLCHAIN": "1.95.0",
            "CARGO_TARGET_DIR": str(target),
            "PYTHONDONTWRITEBYTECODE": "1",
        }
    )
    selected = {
        key: env.get(key)
        for key in (
            "CARGO_BUILD_JOBS",
            "CARGO_INCREMENTAL",
            "CARGO_NET_OFFLINE",
            "RUSTUP_TOOLCHAIN",
            "CARGO_TARGET_DIR",
            "RUSTFLAGS",
            "CARGO_ENCODED_RUSTFLAGS",
            "RUSTC_WRAPPER",
            "PATH",
            "LANG",
            "LC_ALL",
        )
        if env.get(key) is not None
    }
    return {
        "leg": leg,
        "platform": platform.platform(),
        "machine": platform.machine(),
        "processor": platform.processor(),
        "cpu": CPU,
        "allowed_cpus": allowed,
        "rustc": subprocess.check_output(
            ["rustc", "-Vv"], env=env, text=True
        ),
        "cargo": subprocess.check_output(["cargo", "-V"], env=env, text=True).strip(),
        "profile": {
            "release": True,
            "opt_level": 3,
            "lto": True,
            "panic": "abort",
            "debug": False,
        },
        "environment": selected,
        "references": references,
    }, env


def artifact_receipt(path: Path) -> dict[str, Any]:
    if not path.is_file():
        raise RuntimeError(f"missing artifact: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def run_command(
    runs: list[dict[str, Any]],
    *,
    kind: str,
    name: str,
    command: list[str],
    cwd: Path,
    env: dict[str, str],
    log: Path,
    metadata: dict[str, Any] | None = None,
) -> subprocess.CompletedProcess[str]:
    started = time.time_ns()
    with log.open("w") as stream:
        result = subprocess.run(
            command,
            cwd=cwd,
            env=env,
            stdout=stream,
            stderr=subprocess.STDOUT,
            text=True,
        )
    ended = time.time_ns()
    row: dict[str, Any] = {
        "kind": kind,
        "name": name,
        "command": command,
        "cwd": str(cwd),
        "exit": result.returncode,
        "started_ns": started,
        "ended_ns": ended,
        "log": artifact_receipt(log),
    }
    if metadata:
        row.update(metadata)
    runs.append(row)
    json_write(CAPTURE / "runs.json", runs)
    return result


def require_success(result: subprocess.CompletedProcess[str], name: str, log: Path) -> None:
    if result.returncode:
        raise RuntimeError(f"{name} failed with {result.returncode}; see {log}")


def binary_receipts(target: Path, names: Iterable[str]) -> dict[str, Any]:
    return {
        name: artifact_receipt(target / "release" / name) for name in names
    }


def check_report(report: Path, case: str, n: int | None) -> dict[str, Any]:
    value = json_read(report)
    if value.get("case") != case:
        raise RuntimeError(f"wrong case in {report}: {value.get('case')!r}")
    if value.get("samples") != SAMPLES or value.get("warmup") != WARMUP:
        raise RuntimeError(f"wrong sample settings in {report}")
    samples = value.get("durations_ns", value.get("elapsed_ns"))
    if not isinstance(samples, list) or len(samples) != SAMPLES:
        raise RuntimeError(f"wrong timing sample count in {report}")
    if not value.get("outcomes"):
        raise RuntimeError(f"no outcomes in {report}")
    if n is not None and value.get("n") != n:
        raise RuntimeError(f"wrong n in {report}: {value.get('n')!r} != {n}")
    return {
        "case": case,
        "n": n,
        "samples": len(samples),
        "outcomes": value["outcomes"],
        "report": artifact_receipt(report),
    }


def main() -> int:
    if CAPTURE.exists():
        raise RuntimeError(f"refusing to overwrite existing output: {CAPTURE}")
    if not BEFORE.is_dir() or not AFTER.is_dir():
        raise RuntimeError(f"missing before/after worktree: {BEFORE} / {AFTER}")
    if any(target.exists() for target in TARGETS.values()):
        raise RuntimeError("refusing to reuse an existing 0777 release target")
    if not Path("/usr/bin/time").is_file():
        raise RuntimeError("/usr/bin/time is required for RSS receipts")
    final_source_path = PACKET / "final-source.json"
    if not final_source_path.is_file():
        raise RuntimeError("root must write final-source.json before capture")

    CAPTURE.mkdir()
    runner = artifact_receipt(Path(__file__).resolve())
    json_write(CAPTURE / "runner.json", runner)

    final_source = json_read(final_source_path)
    candidate_ref = git(AFTER, "rev-parse", "HEAD")
    if final_source.get("candidate") != candidate_ref:
        raise RuntimeError(
            f"final-source candidate {final_source.get('candidate')!r} != HEAD {candidate_ref}"
        )
    if final_source.get("base") != BASE_REF[:10]:
        raise RuntimeError("final-source base does not match the 0777 baseline")
    references = {"before": BASE_REF, "after": candidate_ref}

    base_source = source_manifest(BEFORE, BASE_REF)
    after_source = source_manifest(AFTER, candidate_ref)
    verify_final_source(final_source, candidate_ref, after_source)
    if base_source["head"] != BASE_REF:
        raise RuntimeError(f"base HEAD changed: {base_source['head']}")
    if base_source["files"].get("Cargo.lock") != after_source["files"].get("Cargo.lock"):
        raise RuntimeError("root Cargo.lock differs between capture legs")
    json_write(CAPTURE / "source-before.json", base_source)
    json_write(CAPTURE / "source-after.json", after_source)
    json_write(CAPTURE / "final-source.json", final_source)

    input_manifest = probe_input_manifest()
    json_write(CAPTURE / "probe-inputs.json", input_manifest)
    fixtures_before = fixture_manifest()
    inventory_before = package_inventory()
    json_write(CAPTURE / "fixtures-before.json", fixtures_before)
    json_write(CAPTURE / "package-inventory-before.json", inventory_before)
    json_write(
        CAPTURE / "scan-roots.json",
        {
            "roots": [str(root) for root in ROOT_RELATIVE],
            "absolute_roots": [str(AFTER / root) for root in ROOT_RELATIVE],
            "collector": "regular files with PK\\x03\\x04 prefix; symlinks skipped",
            "mutate_limit": MUTATE_LIMIT,
        },
    )

    references_json = {name: value for name, value in references.items()}
    environment_receipts: dict[str, Any] = {}
    environments: dict[str, dict[str, str]] = {}
    for leg in ("before", "after"):
        receipt, env = environment(leg, TARGETS[leg], references_json)
        environment_receipts[leg] = receipt
        environments[leg] = env
    json_write(CAPTURE / "environment.json", environment_receipts)

    manifests: dict[str, Path] = {}
    source_maps: dict[str, Any] = {}
    for leg, root in (("before", BEFORE), ("after", AFTER)):
        manifest, source_map = rendered_probe(leg, root, input_manifest)
        manifests[leg] = manifest
        source_maps[leg] = source_map

    runs: list[dict[str, Any]] = []
    lock_log = CAPTURE / "lock-before.log"
    lock_command = [
        "cargo",
        "generate-lockfile",
        "--offline",
        "--manifest-path",
        str(manifests["before"]),
    ]
    lock_result = run_command(
        runs,
        kind="standalone-lock",
        name="generate-before",
        command=lock_command,
        cwd=AFTER,
        env=environments["before"],
        log=lock_log,
    )
    require_success(lock_result, "standalone lock generation", lock_log)
    before_lock = manifests["before"].parent / "Cargo.lock"
    after_lock = manifests["after"].parent / "Cargo.lock"
    shutil.copy2(before_lock, after_lock)
    resolved = tomllib.loads(before_lock.read_text())["package"]
    assert [x["version"] for x in resolved if x["name"] == "quick-xml"] == ["0.41.0"], "reviewed quick-xml version changed"

    lock_receipt = {
        "before": artifact_receipt(before_lock),
        "after": artifact_receipt(after_lock),
        "copied_before_to_after": True,
    }
    if lock_receipt["before"]["sha256"] != lock_receipt["after"]["sha256"]:
        raise RuntimeError("standalone lock copy changed bytes")
    json_write(CAPTURE / "standalone-lock.json", lock_receipt)

    build_commands = {
        "before": [
            "cargo",
            "build",
            "--release",
            "--offline",
            "--locked",
            "--no-default-features",
            "--manifest-path",
            str(manifests["before"]),
            "--bin",
            "attribute_checks",
            "--bin",
            "mce-stream-probe",
        ],
        "after": [
            "cargo",
            "build",
            "--release",
            "--offline",
            "--locked",
            "--no-default-features",
            "--features",
            "candidate-equivalence",
            "--manifest-path",
            str(manifests["after"]),
            "--bin",
            "attribute_checks",
            "--bin",
            "attribute_checks_equivalence",
            "--bin",
            "mce-stream-probe",
        ],
    }
    json_write(CAPTURE / "build-commands.json", build_commands)
    for leg in ("before", "after"):
        log = CAPTURE / f"build-{leg}.log"
        result = run_command(
            runs,
            kind="build",
            name=leg,
            command=build_commands[leg],
            cwd=AFTER,
            env=environments[leg],
            log=log,
            metadata={"target": str(TARGETS[leg])},
        )
        require_success(result, f"{leg} release build", log)

    binaries = {
        "before": binary_receipts(
            TARGETS["before"], ("attribute_checks", "mce-stream-probe")
        ),
        "after": binary_receipts(
            TARGETS["after"],
            ("attribute_checks", "attribute_checks_equivalence", "mce-stream-probe"),
        ),
    }
    json_write(CAPTURE / "binaries.json", binaries)

    root_args = [str(root) for root in ROOT_RELATIVE]
    equivalence_report = CAPTURE / "equivalence.json"
    equivalence_log = CAPTURE / "equivalence.log"
    equivalence_command = [
        "taskset",
        "-c",
        str(CPU),
        str(TARGETS["after"] / "release/attribute_checks_equivalence"),
        "--json",
        str(equivalence_report),
        *root_args,
    ]
    equivalence_result = run_command(
        runs,
        kind="equivalence",
        name="candidate",
        command=equivalence_command,
        cwd=AFTER,
        env=environments["after"],
        log=equivalence_log,
        metadata={"report": artifact_receipt(equivalence_report)
                  if equivalence_report.is_file() else None},
    )
    require_success(equivalence_result, "candidate equivalence", equivalence_log)
    runs[-1]["report"] = artifact_receipt(equivalence_report)
    json_write(CAPTURE / "runs.json", runs)
    json_write(CAPTURE / "equivalence-receipt.json", artifact_receipt(equivalence_report))

    for leg in ("before", "after"):
        report = CAPTURE / f"differential-{leg}.json"
        log = CAPTURE / f"differential-{leg}.log"
        command = [
            "taskset",
            "-c",
            str(CPU),
            str(TARGETS[leg] / "release/attribute_checks"),
            "differential",
            "--mutate-limit",
            str(MUTATE_LIMIT),
            "--json",
            str(report),
            *root_args,
        ]
        result = run_command(
            runs,
            kind="differential",
            name=leg,
            command=command,
            cwd=AFTER,
            env=environments[leg],
            log=log,
            metadata={"report": artifact_receipt(report)
                      if report.is_file() else None},
        )
        require_success(result, f"{leg} differential", log)
        runs[-1]["report"] = artifact_receipt(report)
        json_write(CAPTURE / "runs.json", runs)
        json_write(CAPTURE / f"differential-{leg}-receipt.json", artifact_receipt(report))

    cases: list[tuple[str, str, int | None]] = [
        ("mce_benign_worksheet", "mce-stream-probe", None),
        ("mce_benign_document", "mce-stream-probe", None),
        ("mce_stream_count_worksheet", "mce-stream-probe", None),
        ("mce_stream_count_document", "mce-stream-probe", None),
        ("mce_prefixed_1", "mce-stream-probe", None),
        ("mce_prefixed_2", "mce-stream-probe", None),
        ("mce_prefixed_8", "mce-stream-probe", None),
        ("mce_prefixed_9", "mce-stream-probe", None),
        ("mce_prefixed_32", "mce-stream-probe", None),
    ]
    cases.extend(
        ("opc_relationship_declarations", "attribute_checks", n)
        for n in (0, 8, 29, 30, 32, 33, 256, 1024, 4096, 16384)
    )
    native_receipts: list[dict[str, Any]] = []
    for case, binary_name, n in cases:
        for index, leg in enumerate(LEGS):
            suffix = "" if n is None else f"-n{n}"
            stem = f"{case}{suffix}-{index:02d}-{leg}"
            report = CAPTURE / f"{stem}.json"
            log = CAPTURE / f"{stem}.log"
            rss = CAPTURE / f"{stem}.rss"
            for output in (report, log, rss):
                if output.exists():
                    raise RuntimeError(f"refusing to overwrite native output: {output}")
            binary = TARGETS[leg] / f"release/{binary_name}"
            if n is None:
                command = [
                    "/usr/bin/time",
                    "-f",
                    "%M",
                    "-o",
                    str(rss),
                    "taskset",
                    "-c",
                    str(CPU),
                    str(binary),
                    "adversarial",
                    "--case",
                    case,
                    "--samples",
                    str(SAMPLES),
                    "--warmup",
                    str(WARMUP),
                    "--json",
                    str(report),
                ]
            else:
                command = [
                    "/usr/bin/time",
                    "-f",
                    "%M",
                    "-o",
                    str(rss),
                    "taskset",
                    "-c",
                    str(CPU),
                    str(binary),
                    "probe",
                    "--case",
                    case,
                    "--n",
                    str(n),
                    "--samples",
                    str(SAMPLES),
                    "--warmup",
                    str(WARMUP),
                    "--json",
                    str(report),
                ]
            result = run_command(
                runs,
                kind="native",
                name=stem,
                command=command,
                cwd=AFTER,
                env=environments[leg],
                log=log,
                metadata={
                    "leg": leg,
                    "case": case,
                    "n": n,
                    "report": artifact_receipt(report) if report.is_file() else None,
                    "rss": artifact_receipt(rss) if rss.is_file() else None,
                    "binary": binaries[leg][binary_name],
                    "order_index": index,
                },
            )
            require_success(result, stem, log)
            runs[-1]["report"] = artifact_receipt(report)
            runs[-1]["rss"] = artifact_receipt(rss)
            json_write(CAPTURE / "runs.json", runs)
            checked = check_report(report, case, n)
            rss_text = rss.read_text().strip()
            if not rss_text.isdigit():
                raise RuntimeError(f"invalid RSS receipt in {rss}: {rss_text!r}")
            checked.update(
                {
                    "leg": leg,
                    "order_index": index,
                    "rss_kib": int(rss_text),
                    "rss": artifact_receipt(rss),
                }
            )
            native_receipts.append(checked)
            json_write(CAPTURE / "native-receipts.json", native_receipts)
            print(f"{stem} done", flush=True)

    fixtures_after = fixture_manifest()
    inventory_after = package_inventory()
    if fixtures_after != fixtures_before:
        raise RuntimeError("test-data fixture inventory changed during capture")
    if inventory_after != inventory_before:
        raise RuntimeError("scan-root package inventory changed during capture")
    json_write(CAPTURE / "fixtures-after.json", fixtures_after)
    json_write(CAPTURE / "package-inventory-after.json", inventory_after)

    final_base = source_manifest(BEFORE, BASE_REF)
    final_after = source_manifest(AFTER, candidate_ref)
    if final_base != base_source or final_after != after_source:
        raise RuntimeError("source changed during capture")
    json_write(CAPTURE / "source-final.json", {"before": final_base, "after": final_after})
    for name, relative in PROBE_INPUTS.items():
        if sha256(PACKET / relative) != input_manifest[name]["sha256"]:
            raise RuntimeError(f"probe input changed during capture: {name}")
    for leg, expected in binaries.items():
        for name, receipt in expected.items():
            current = artifact_receipt(Path(receipt["path"]))
            if current != receipt:
                raise RuntimeError(f"binary changed during capture: {leg}/{name}")

    assert artifact_receipt(Path(__file__).resolve()) == runner, "capture runner changed"

    complete = {
        "schema": "litchi-0777-xml-attribute-capture-v1",
        "complete": True,
        "serial": True,
        "base": BASE_REF,
        "candidate": candidate_ref,
        "order": list(LEGS),
        "samples": SAMPLES,
        "warmup": WARMUP,
        "cpu": CPU,
        "mutate_limit": MUTATE_LIMIT,
        "equivalence": "candidate",
        "differential_legs": ["before", "after"],
        "native_cases": len(cases),
        "native_processes": len(native_receipts),
        "known_styles_refusal_included": False,
        "source_unchanged": True,
        "probe_inputs_unchanged": True,
        "fixtures_unchanged": True,
        "package_inventory_unchanged": True,
        "standalone_lock_shared": True,
        "binaries_unchanged": True,
    }
    json_write(CAPTURE / "complete.json", complete)
    print(json.dumps(complete, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:  # preserve a useful failure in the terminal
        print(f"capture failed: {error}", file=sys.stderr)
        raise
