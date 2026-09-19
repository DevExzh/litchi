#!/usr/bin/env python3
"""Run a temporary 0693 MCE/source-slice trace; never timing evidence."""
from __future__ import annotations

import argparse
import difflib
import hashlib
import json
import os
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
CODEC = ROOT / "crates/litchi-ooxml-common/src/mce/codec.rs"
MODEL = ROOT / "crates/litchi-pptx/src/opened/model.rs"
PARTS = ROOT / "crates/litchi-pptx/src/parts/mod.rs"
REAL = ROOT / "test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx"
CONTROL = P / "marker-control.pptx"

CODEC_OLD = b"""pub fn process_ooxml(x: &[u8]) -> R<Cow<'_, [u8]>> {
    process_markup_compatibility(x, &Capabilities::default(), &Limits::default()).map(|x| x.xml)
}
"""
CODEC_NEW = b"""static LITCHI_0693_TRACE_CALL: std::sync::atomic::AtomicU64 =
    std::sync::atomic::AtomicU64::new(0);
pub fn process_ooxml(x: &[u8]) -> R<Cow<'_, [u8]>> {
    let call = LITCHI_0693_TRACE_CALL.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
    let profile = std::env::var("LITCHI_0693_TRACE_PROFILE").unwrap_or_else(|_| "unset".to_owned());
    let stage = std::env::var("LITCHI_0693_TRACE_STAGE").unwrap_or_else(|_| "unset".to_owned());
    let input_ptr = x.as_ptr() as usize;
    let input_len = x.len();
    let result = process_markup_compatibility(x, &Capabilities::default(), &Limits::default()).map(|x| x.xml);
    match &result {
        Ok(Cow::Borrowed(output)) => eprintln!("LITCHI0693_CODEC call={} profile={} stage={} input_ptr=0x{:x} input_len={} output_ptr=0x{:x} output_len={} mode=borrow status=ok", call, profile, stage, input_ptr, input_len, output.as_ptr() as usize, output.len()),
        Ok(Cow::Owned(output)) => eprintln!("LITCHI0693_CODEC call={} profile={} stage={} input_ptr=0x{:x} input_len={} output_ptr=0x{:x} output_len={} mode=owned status=ok", call, profile, stage, input_ptr, input_len, output.as_ptr() as usize, output.len()),
        Err(_) => eprintln!("LITCHI0693_CODEC call={} profile={} stage={} input_ptr=0x{:x} input_len={} output_ptr=0x0 output_len=0 mode=error status=error", call, profile, stage, input_ptr, input_len),
    }
    result
}
"""

PARTS_OLD = b"""pub(crate) fn processed_xml(part: &dyn Part) -> Result<Cow<'_, [u8]>> {
    if part.blob().len() > MAX_PART_XML_BYTES {
        return Err(Error::Limit {
            resource: "PresentationML part XML",
            limit: MAX_PART_XML_BYTES,
        });
    }
    Ok(process_part(part)?)
}
"""
PARTS_NEW = b"""pub(crate) fn processed_xml(part: &dyn Part) -> Result<Cow<'_, [u8]>> {
    if part.blob().len() > MAX_PART_XML_BYTES {
        return Err(Error::Limit {
            resource: "PresentationML part XML",
            limit: MAX_PART_XML_BYTES,
        });
    }
    let source = part.blob();
    eprintln!(
        "LITCHI0693_PART profile={} stage=processed_xml uri={} raw_ptr=0x{:x} raw_len={}",
        std::env::var("LITCHI_0693_TRACE_PROFILE").unwrap_or_else(|_| "unset".to_owned()),
        part.partname().as_str(),
        source.as_ptr() as usize,
        source.len()
    );
    Ok(litchi_ooxml_common::mce::process_ooxml(source)?)
}
"""
CAPTURE_BEGIN = b"""    eprintln!(
        "LITCHI0693_MODEL profile={} stage=capture_begin parts={}",
        std::env::var("LITCHI_0693_TRACE_PROFILE").unwrap_or_else(|_| "unset".to_owned()),
        package.part_count()
    );
"""
CAPTURE_SUCCESS = b"""    eprintln!(
        "LITCHI0693_MODEL profile={} stage=capture_success parts={} slides={}",
        std::env::var("LITCHI_0693_TRACE_PROFILE").unwrap_or_else(|_| "unset".to_owned()),
        package.part_count(),
        slides.len()
    );
"""
CAPTURE_LAYOUT = b"""    let slide_entry_bytes = captured
        .slides
        .first()
        .map_or(0usize, size_of_val);
    eprintln!(
        "LITCHI0693_MODEL profile={} stage=capture_layout proof_entry_bytes={} slide_entry_bytes={}",
        std::env::var("LITCHI_0693_TRACE_PROFILE").unwrap_or_else(|_| "unset".to_owned()),
        size_of::<crate::notes::SlideRootProof<'_>>(),
        slide_entry_bytes
    );
"""


def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def fsha(path: Path) -> str:
    return sha(path.read_bytes())


def write_json(path: Path, value: object) -> None:
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n")


def rel(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def one(data: bytes, old: bytes, new: bytes, label: str) -> bytes:
    if data.count(old) != 1:
        raise RuntimeError(f"{label}: expected one source seam")
    return data.replace(old, new, 1)


def patch_source_helper(data: bytes) -> bytes:
    marker = b"pub(crate) fn processed_xml_with_source"
    start = data.find(marker)
    end = data.find(b"\n}\n\npub(crate) fn processed_xml(", start)
    if start < 0 or end < 0:
        raise RuntimeError(f"{PARTS}: source-bound preprocessing seam missing")
    block = data[start : end + 3]
    trace = b"""    eprintln!(
        \"LITCHI0693_PART profile={} stage=source_bound uri={} raw_ptr=0x{:x} raw_len={}\",
        std::env::var(\"LITCHI_0693_TRACE_PROFILE\").unwrap_or_else(|_| \"unset\".to_owned()),
        part.partname().as_str(),
        source.as_ptr() as usize,
        source.len()
    );
"""
    patched = one(block, b"    let source = part.blob();\n", b"    let source = part.blob();\n" + trace, str(PARTS))
    return data[:start] + patched + data[end + 3 :]


def patch_parts(data: bytes) -> bytes:
    if b"use litchi_ooxml_common::mce::process_part;" in data:
        data = data.replace(b"use litchi_ooxml_common::mce::process_part;\n", b"", 1)
    elif b"use litchi_ooxml_common::mce::process_ooxml;" not in data:
        raise RuntimeError(f"{PARTS}: unexpected MCE import")
    data = one(data, PARTS_OLD, PARTS_NEW, str(PARTS))
    return patch_source_helper(data)


def patch_model(data: bytes) -> bytes:
    start = data.find(b"fn capture_internal(")
    if start < 0:
        raise RuntimeError(f"{MODEL}: capture_internal seam missing")
    begin = data.find(b"    let presentation = PresentationPart::from_package(package)?;", start)
    end = data.find(b"\n}\n\n/// The `litchi-pptx-opened-v2` complete-package revision", begin)
    if begin < 0 or end < 0:
        raise RuntimeError(f"{MODEL}: capture body seam missing")
    data = data[:begin] + CAPTURE_BEGIN + data[begin:]
    end += len(CAPTURE_BEGIN)
    block = data[start : end + 2]
    block = one(
        block,
        b"    let captured = view.capture_slides()?;\n",
        b"    let captured = view.capture_slides()?;\n" + CAPTURE_LAYOUT,
        str(MODEL),
    )
    block = one(block, b"    Ok(Snapshot {\n", CAPTURE_SUCCESS + b"    Ok(Snapshot {\n", str(MODEL))
    return data[:start] + block + data[end + 2 :]


def changes(model: bool) -> list[tuple[str, Path, bytes, bytes]]:
    before = CODEC.read_bytes()
    out = [(rel(CODEC), CODEC, before, one(before, CODEC_OLD, CODEC_NEW, str(CODEC)))]
    before = PARTS.read_bytes()
    out.append((rel(PARTS), PARTS, before, patch_parts(before)))
    if model:
        before = MODEL.read_bytes()
        out.append((rel(MODEL), MODEL, before, patch_model(before)))
    return out


def source_binding() -> dict[str, str]:
    for candidate in (P / "builds-candidate.json", P / "candidate-source.json"):
        if not candidate.is_file():
            continue
        value = json.loads(candidate.read_text())
        rows = value if isinstance(value, list) else [value]
        for row in rows:
            if row.get("phase") in (None, "candidate") and row.get("source_sha256"):
                return row["source_sha256"]
    return {}


def verify(cs: list[tuple[str, Path, bytes, bytes]]) -> None:
    expected = source_binding()
    for name, path, before, _ in cs:
        actual = fsha(path)
        if actual != sha(before):
            raise RuntimeError(f"{name}: source changed while preparing trace")
        if expected.get(name) and actual != expected[name]:
            raise RuntimeError(f"{name}: bytes differ from candidate source binding")


def archive(run: Path, cs: list[tuple[str, Path, bytes, bytes]]) -> None:
    source_dir = run / "sources"
    source_dir.mkdir()
    patch: list[str] = []
    files = []
    for name, _, before, after in cs:
        stem = name.replace("/", "__")
        (source_dir / (stem + ".before")).write_bytes(before)
        (source_dir / (stem + ".after")).write_bytes(after)
        patch += difflib.unified_diff(
            before.decode().splitlines(True), after.decode().splitlines(True),
            fromfile="a/" + name, tofile="b/" + name,
        )
        files.append({"path": name, "before_sha256": sha(before), "after_sha256": sha(after)})
    (run / "trace.patch").write_text("".join(patch))
    write_json(run / "source-hashes.json", {"files": files, "patch_sha256": fsha(run / "trace.patch")})


def restore(run: Path, cs: list[tuple[str, Path, bytes, bytes]]) -> list[dict[str, object]]:
    records = []
    errors = []
    for name, path, before, after in reversed(cs):
        try:
            current = path.read_bytes()
            path.write_bytes(before)
            restored = path.read_bytes()
            records.append({"path": name, "current_sha256": sha(current), "was_patched": current == after,
                            "restored_sha256": sha(restored), "restored_exact": restored == before})
            if restored != before:
                errors.append(name)
        except OSError as exc:
            records.append({"path": name, "error": str(exc), "restored_exact": False})
            errors.append(name)
    write_json(run / "restoration.json", records)
    if errors:
        raise RuntimeError("failed exact restoration: " + ", ".join(errors))
    return records


def parse() -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--profile", default="baseline")
    p.add_argument("--model-trace", action="store_true")
    p.add_argument("--cargo", default="cargo")
    p.add_argument("--manifest-path", default=str(P / "probe/Cargo.toml"))
    p.add_argument("--target-dir", default=str(ROOT.parent / "litchi-target-0693"))
    p.add_argument("--binary-name", default="probe0693")
    p.add_argument("--jobs", type=int, default=2)
    p.add_argument("--feature", action="append", default=[])
    p.add_argument("--rustflags", default="-D warnings")
    p.add_argument("--skip-build", action="store_true")
    p.add_argument("--binary")
    p.add_argument("--real-source", default=str(REAL))
    p.add_argument("--control-source", default=str(CONTROL))
    p.add_argument("--generated-source", default="generated:12x8")
    p.add_argument("--with-generated", action="store_true")
    p.add_argument("--with-control", action="store_true")
    p.add_argument("--operation", action="append", choices=("capture", "apply"))
    p.add_argument("--capture-count", type=int, default=1)
    p.add_argument("--apply-count", type=int, default=1)
    p.add_argument("--probe-arg", action="append", default=[])
    p.add_argument("--cpu", type=int)
    p.add_argument("--timeout", type=float)
    p.add_argument("--output-root", default=str(P / "trace-runs"))
    return p.parse_args()


def path(value: str) -> Path:
    candidate = Path(value).expanduser()
    return candidate if candidate.is_absolute() else ROOT / candidate


def run(argv: list[str], stem: Path, env: dict[str, str], timeout: float | None) -> int:
    with stem.with_suffix(".stdout").open("wb") as out, stem.with_suffix(".stderr").open("wb") as err:
        return subprocess.run(argv, cwd=ROOT, env=env, stdout=out, stderr=err,
                              timeout=timeout, check=False).returncode


def main() -> int:
    args = parse()
    if args.jobs < 1 or args.capture_count < 1 or args.apply_count < 1:
        raise ValueError("jobs/counts must be positive")
    if any(c not in "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_.:-" for c in args.profile):
        raise ValueError("unsafe profile label")
    root = path(args.output_root)
    root.mkdir(parents=True, exist_ok=True)
    run_dir = root / f"{datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S%fZ')}-{os.getpid()}"
    run_dir.mkdir()
    cs = changes(args.model_trace)
    verify(cs)
    archive(run_dir, cs)
    real = path(args.real_source)
    if not real.is_file():
        raise RuntimeError(f"missing real source: {real}")
    cases = [("real", str(real))]
    if args.with_generated:
        cases.append(("generated", args.generated_source))
    if args.with_control:
        control = path(args.control_source)
        if not control.is_file():
            raise RuntimeError(f"missing control source: {control}")
        cases.append(("control", str(control)))
    operations = args.operation or ["capture", "apply"]
    manifest: dict[str, object] = {
        "schema": "litchi-0693-trace-v1", "diagnostic_only": True,
        "timing_evidence": False, "profile": args.profile,
        "model_trace": args.model_trace, "cases": cases,
        "operations": operations, "status": "started",
    }
    write_json(run_dir / "plan.json", manifest)
    failure: BaseException | None = None
    records = []
    try:
        for _, source, _, after in cs:
            source.write_bytes(after)
        if args.skip_build:
            if not args.binary:
                raise RuntimeError("--skip-build requires --binary")
            binary = path(args.binary)
            build = {"skipped": True}
        else:
            target = path(args.target_dir)
            target.mkdir(parents=True, exist_ok=True)
            command = [args.cargo, "build", "--release", "--locked", "--manifest-path",
                       str(path(args.manifest_path)), "--target-dir", str(target),
                       "--bin", args.binary_name, "-j", str(args.jobs)]
            if args.feature:
                command += ["--features", ",".join(args.feature)]
            env = os.environ.copy()
            env["RUSTFLAGS"] = args.rustflags
            code = run(command, run_dir / "build", env, None)
            build = {"command": command, "exit_code": code}
            write_json(run_dir / "build.json", build)
            if code:
                raise RuntimeError(f"trace build failed with {code}")
            binary = target / "release" / args.binary_name
        if not binary.is_file():
            raise RuntimeError(f"missing trace binary: {binary}")
        env = os.environ.copy()
        env["LITCHI_0693_TRACE_PROFILE"] = args.profile
        for case, source in cases:
            for operation in operations:
                count = args.capture_count if operation == "capture" else args.apply_count
                label = f"{case}-{operation}-{count}"
                command = [str(binary), "prefix", source, operation, str(count), *args.probe_arg]
                if args.cpu is not None:
                    command = ["taskset", "-c", str(args.cpu), *command]
                env["LITCHI_0693_TRACE_STAGE"] = f"prefix:{case}:{operation}"
                code = run(command, run_dir / label, env, args.timeout)
                out, err = run_dir / f"{label}.stdout", run_dir / f"{label}.stderr"
                records.append({"case": case, "source": source, "operation": operation,
                                "count": count, "command": command, "exit_code": code,
                                "stdout_sha256": fsha(out), "stderr_sha256": fsha(err)})
                write_json(run_dir / "probe-runs.json", records)
                if code:
                    raise RuntimeError(f"probe failed for {label} with {code}")
        manifest.update({"status": "completed", "binary": str(binary),
                         "binary_sha256": fsha(binary), "build": build,
                         "probe_runs": records})
    except BaseException as exc:
        failure = exc
        manifest.update({"status": "failed", "error": f"{type(exc).__name__}: {exc}",
                         "probe_runs": records})
    finally:
        try:
            restored = restore(run_dir, cs)
            manifest["restoration"] = restored
            manifest["source_restored_exact"] = all(x["restored_exact"] for x in restored)
        except BaseException as exc:
            failure = failure or exc
            manifest.update({"status": "restore-failed", "error": str(exc)})
        write_json(run_dir / "manifest.json", manifest)
    print(run_dir)
    if failure:
        print(f"trace failed: {failure}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, RuntimeError, ValueError) as exc:
        print(f"trace setup failed: {exc}", file=sys.stderr)
        raise SystemExit(2)
