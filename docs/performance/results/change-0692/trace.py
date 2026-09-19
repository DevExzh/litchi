#!/usr/bin/env python3
"""Temporary 0692 MCE trace; stderr is diagnostic only, never timing evidence."""
from __future__ import annotations
import argparse, difflib, hashlib, json, os, subprocess, sys
from datetime import datetime, timezone
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
BASE = P / "baseline.json"
CODEC = ROOT / "crates/litchi-ooxml-common/src/mce/codec.rs"
MODEL = ROOT / "crates/litchi-pptx/src/opened/model.rs"
PARTS = ROOT / "crates/litchi-pptx/src/parts/mod.rs"
REAL = ROOT / "test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx"
CONTROL = P / "marker-control.pptx"

CODEC_OLD = b"""pub fn process_ooxml(x: &[u8]) -> R<Cow<'_, [u8]>> {
    process_markup_compatibility(x, &Capabilities::default(), &Limits::default()).map(|x| x.xml)
}
"""
CODEC_NEW = b"""static LITCHI_0692_TRACE_CALL: std::sync::atomic::AtomicU64 =
    std::sync::atomic::AtomicU64::new(0);
pub fn process_ooxml(x: &[u8]) -> R<Cow<'_, [u8]>> {
    let call = LITCHI_0692_TRACE_CALL.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
    let profile = std::env::var("LITCHI_0692_TRACE_PROFILE").unwrap_or_else(|_| "unset".to_owned());
    let stage = std::env::var("LITCHI_0692_TRACE_STAGE").unwrap_or_else(|_| "unset".to_owned());
    let input_ptr = x.as_ptr() as usize;
    let input_len = x.len();
    let result = process_markup_compatibility(x, &Capabilities::default(), &Limits::default()).map(|x| x.xml);
    match &result {
        Ok(Cow::Borrowed(o)) => eprintln!("LITCHI0692_CODEC call={} profile={} stage={} input_ptr=0x{:x} input_len={} output_ptr=0x{:x} output_len={} mode=borrow status=ok", call, profile, stage, input_ptr, input_len, o.as_ptr() as usize, o.len()),
        Ok(Cow::Owned(o)) => eprintln!("LITCHI0692_CODEC call={} profile={} stage={} input_ptr=0x{:x} input_len={} output_ptr=0x{:x} output_len={} mode=owned status=ok", call, profile, stage, input_ptr, input_len, o.as_ptr() as usize, o.len()),
        Err(_) => eprintln!("LITCHI0692_CODEC call={} profile={} stage={} input_ptr=0x{:x} input_len={} output_ptr=0x0 output_len=0 mode=error status=error", call, profile, stage, input_ptr, input_len),
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
    let visible = part.blob();
    let (arc_struct, arc_bytes, arc_len, alias) = {
        let shared = part.blob_arc();
        (std::sync::Arc::as_ptr(&shared) as usize, shared.as_ptr() as usize, shared.len(),
         shared.len() == visible.len() && std::ptr::eq(shared.as_slice(), visible))
    };
    let profile = std::env::var("LITCHI_0692_TRACE_PROFILE").unwrap_or_else(|_| "unset".to_owned());
    eprintln!("LITCHI0692_PART profile={} stage=process_part uri={} visible_ptr=0x{:x} visible_len={} arc_struct_ptr=0x{:x} arc_bytes_ptr=0x{:x} arc_len={} alias={}", profile, part.partname().as_str(), visible.as_ptr() as usize, visible.len(), arc_struct, arc_bytes, arc_len, alias);
    if visible.len() > MAX_PART_XML_BYTES {
        return Err(Error::Limit { resource: "PresentationML part XML", limit: MAX_PART_XML_BYTES });
    }
    Ok(process_part(part)?)
}
"""
CAPTURE_OLD = b""") -> Result<Snapshot> {
    let presentation = PresentationPart::from_package(package)?;"""
CAPTURE_NEW = b""") -> Result<Snapshot> {
    eprintln!("LITCHI0692_MODEL profile={} stage=capture_begin parts={}", std::env::var("LITCHI_0692_TRACE_PROFILE").unwrap_or_else(|_| "unset".to_owned()), package.part_count());
    let presentation = PresentationPart::from_package(package)?;"""
CAPTURE_END_OLD = b"""    Ok(Snapshot {
"""
CAPTURE_END_NEW = b"""    eprintln!("LITCHI0692_MODEL profile={} stage=capture_success parts={} slides={}", std::env::var("LITCHI_0692_TRACE_PROFILE").unwrap_or_else(|_| "unset".to_owned()), package.part_count(), slides.len());
    Ok(Snapshot {
"""

def sha(data: bytes) -> str: return hashlib.sha256(data).hexdigest()
def fsha(path: Path) -> str: return sha(path.read_bytes())
def write_json(path: Path, value: object) -> None: path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n")
def rel(path: Path) -> str: return path.relative_to(ROOT).as_posix()
def one(data: bytes, old: bytes, new: bytes, label: str) -> bytes:
    if data.count(old) != 1: raise RuntimeError(f"{label}: expected one source seam")
    return data.replace(old, new, 1)
def changes(model: bool) -> list[tuple[str, Path, bytes, bytes]]:
    before = CODEC.read_bytes()
    out = [(rel(CODEC), CODEC, before, one(before, CODEC_OLD, CODEC_NEW, str(CODEC)))]
    if model:
        before = PARTS.read_bytes()
        out.append((rel(PARTS), PARTS, before, one(before, PARTS_OLD, PARTS_NEW, str(PARTS))))
        before = MODEL.read_bytes()
        after = one(before, CAPTURE_OLD, CAPTURE_NEW, str(MODEL))
        out.append((rel(MODEL), MODEL, before, one(after, CAPTURE_END_OLD, CAPTURE_END_NEW, str(MODEL))))
    return out
def verify(cs: list[tuple[str, Path, bytes, bytes]]) -> None:
    expected = json.loads((P / "builds-candidate.json").read_text())[0]["source_sha256"]
    for name, path, before, _ in cs:
        if fsha(path) != expected.get(name) or fsha(path) != sha(before):
            raise RuntimeError(f"{name}: bytes differ from frozen baseline")
def archive(run: Path, cs: list[tuple[str, Path, bytes, bytes]]) -> None:
    source_dir = run / "sources"; source_dir.mkdir()
    patch: list[str] = []; files = []
    for name, _, before, after in cs:
        stem = name.replace("/", "__")
        (source_dir / (stem + ".before")).write_bytes(before)
        (source_dir / (stem + ".after")).write_bytes(after)
        patch += difflib.unified_diff(before.decode().splitlines(True), after.decode().splitlines(True), fromfile="a/"+name, tofile="b/"+name)
        files.append({"path": name, "before_sha256": sha(before), "after_sha256": sha(after)})
    (run / "trace.patch").write_text("".join(patch))
    write_json(run / "source-hashes.json", {"files": files, "patch_sha256": fsha(run / "trace.patch")})
def restore(run: Path, cs: list[tuple[str, Path, bytes, bytes]]) -> list[dict[str, object]]:
    result = []; errors = []
    for name, path, before, after in reversed(cs):
        try:
            current = path.read_bytes(); path.write_bytes(before); restored = path.read_bytes()
            result.append({"path": name, "current_sha256": sha(current), "was_patched": current == after, "restored_sha256": sha(restored), "restored_exact": restored == before})
            if restored != before: errors.append(name)
        except OSError as exc:
            result.append({"path": name, "error": str(exc), "restored_exact": False}); errors.append(name)
    write_json(run / "restoration.json", result)
    if errors: raise RuntimeError("failed exact restoration: " + ", ".join(errors))
    return result

def parse() -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--profile", default="baseline"); p.add_argument("--model-trace", action="store_true")
    p.add_argument("--cargo", default="cargo"); p.add_argument("--manifest-path", default=str(P/"probe/Cargo.toml"))
    p.add_argument("--target-dir", default=str(ROOT.parent/"litchi-target-0692-trace")); p.add_argument("--binary-name", default="probe0692")
    p.add_argument("--jobs", type=int, default=2); p.add_argument("--feature", action="append", default=[])
    p.add_argument("--rustflags", default="-D warnings"); p.add_argument("--skip-build", action="store_true"); p.add_argument("--binary")
    p.add_argument("--real-source", default=str(REAL)); p.add_argument("--control-source", default=str(CONTROL)); p.add_argument("--generated-source", default="generated:12x8")
    p.add_argument("--with-generated", action="store_true"); p.add_argument("--with-control", action="store_true")
    p.add_argument("--operation", action="append", choices=("capture", "apply")); p.add_argument("--capture-count", type=int, default=1); p.add_argument("--apply-count", type=int, default=1)
    p.add_argument("--probe-arg", action="append", default=[]); p.add_argument("--cpu", type=int); p.add_argument("--timeout", type=float); p.add_argument("--output-root", default=str(P/"trace-runs"))
    return p.parse_args()
def path(value: str) -> Path:
    p = Path(value).expanduser(); return p if p.is_absolute() else ROOT/p
def run(argv: list[str], stem: Path, env: dict[str, str], timeout: float | None) -> int:
    with stem.with_suffix(".stdout").open("wb") as out, stem.with_suffix(".stderr").open("wb") as err:
        return subprocess.run(argv, cwd=ROOT, env=env, stdout=out, stderr=err, timeout=timeout, check=False).returncode

def main() -> int:
    a = parse()
    if a.jobs < 1 or a.capture_count < 1 or a.apply_count < 1: raise ValueError("jobs/counts must be positive")
    if any(c not in "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_.:-" for c in a.profile): raise ValueError("unsafe profile label")
    root = path(a.output_root); root.mkdir(parents=True, exist_ok=True); run_dir = root / f"{datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S%fZ')}-{os.getpid()}"; run_dir.mkdir()
    cs = changes(a.model_trace); verify(cs); archive(run_dir, cs)
    real = path(a.real_source)
    if not real.is_file(): raise RuntimeError(f"missing real source: {real}")
    cases = [("real", str(real))]
    if a.with_generated: cases.append(("generated", a.generated_source))
    if a.with_control:
        control = path(a.control_source)
        if not control.is_file(): raise RuntimeError(f"missing control source: {control}")
        cases.append(("control", str(control)))
    target = path(a.target_dir); ops = a.operation or ["capture", "apply"]
    manifest: dict[str, object] = {"schema": "litchi-0692-trace-v1", "diagnostic_only": True, "timing_evidence": False, "profile": a.profile, "model_trace": a.model_trace, "cases": cases, "operations": ops, "status": "started"}
    write_json(run_dir/"plan.json", manifest); failure: BaseException | None = None; records = []
    try:
        for _, source, _, after in cs: source.write_bytes(after)
        if a.skip_build:
            if not a.binary: raise RuntimeError("--skip-build requires --binary")
            binary = path(a.binary); build = {"skipped": True}
        else:
            target.mkdir(parents=True, exist_ok=True)
            command = [a.cargo, "build", "--release", "--locked", "--manifest-path", str(path(a.manifest_path)), "--target-dir", str(target), "--bin", a.binary_name, "-j", str(a.jobs)]
            if a.feature: command += ["--features", ",".join(a.feature)]
            env = os.environ.copy(); env["RUSTFLAGS"] = a.rustflags; code = run(command, run_dir/"build", env, None)
            build = {"command": command, "exit_code": code}; write_json(run_dir/"build.json", build)
            if code: raise RuntimeError(f"trace build failed with {code}")
            binary = target/"release"/a.binary_name
        if not binary.is_file(): raise RuntimeError(f"missing trace binary: {binary}")
        env = os.environ.copy(); env["LITCHI_0692_TRACE_PROFILE"] = a.profile
        for case, source in cases:
            for op in ops:
                count = a.capture_count if op == "capture" else a.apply_count; label = f"{case}-{op}-{count}"
                command = [str(binary), "prefix", source, op, str(count), *a.probe_arg]
                if a.cpu is not None: command = ["taskset", "-c", str(a.cpu), *command]
                env["LITCHI_0692_TRACE_STAGE"] = f"prefix:{case}:{op}"; code = run(command, run_dir/label, env, a.timeout)
                out, err = run_dir/f"{label}.stdout", run_dir/f"{label}.stderr"
                records.append({"case": case, "source": source, "operation": op, "count": count, "command": command, "exit_code": code, "stdout_sha256": fsha(out), "stderr_sha256": fsha(err)})
                write_json(run_dir/"probe-runs.json", records)
                if code: raise RuntimeError(f"probe failed for {label} with {code}")
        manifest.update({"status": "completed", "binary": str(binary), "binary_sha256": fsha(binary), "build": build, "probe_runs": records})
    except BaseException as exc:
        failure = exc; manifest.update({"status": "failed", "error": f"{type(exc).__name__}: {exc}", "probe_runs": records})
    finally:
        try:
            restored = restore(run_dir, cs); manifest["restoration"] = restored; manifest["source_restored_exact"] = all(x["restored_exact"] for x in restored)
        except BaseException as exc:
            failure = failure or exc; manifest.update({"status": "restore-failed", "error": str(exc)})
        write_json(run_dir/"manifest.json", manifest)
    print(run_dir)
    if failure: print(f"trace failed: {failure}", file=sys.stderr); return 1
    return 0

if __name__ == "__main__":
    try: raise SystemExit(main())
    except (OSError, RuntimeError, ValueError) as exc:
        print(f"trace setup failed: {exc}", file=sys.stderr); raise SystemExit(2)
