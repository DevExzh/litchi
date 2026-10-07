#!/usr/bin/env python3
"""Fail-closed custody checker for the external 0944 XLSX profile packet."""
from __future__ import annotations
import hashlib, json, os, pathlib, stat, sys
from typing import Any

PACKET = pathlib.Path(__file__).resolve().parent.parent
SCHEMA = "litchi.xlsx-publication-profile-0944.v2"
RUNTIME_ENV = {"PATH":"/usr/bin:/bin", "LANG":"C", "LC_ALL":"C", "TZ":"UTC", "RUST_BACKTRACE":"0"}
HEX = set("0123456789abcdef")

class EvidenceError(ValueError): pass

def reject_pairs(pairs):
    out={}
    for key,value in pairs:
        if key in out: raise EvidenceError(f"duplicate JSON key: {key}")
        out[key]=value
    return out

def load_json(path: pathlib.Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=reject_pairs)
    except (OSError, UnicodeError, json.JSONDecodeError, EvidenceError) as e:
        raise EvidenceError(f"cannot load {path}: {e}") from e

def safe_rel(value: Any, label: str) -> pathlib.Path:
    if not isinstance(value,str) or not value or "\x00" in value:
        raise EvidenceError(f"{label} is not a relative path")
    p=pathlib.PurePosixPath(value)
    if p.is_absolute() or ".." in p.parts or "." in p.parts:
        raise EvidenceError(f"{label} escapes packet: {value}")
    if not value.isascii():
        raise EvidenceError(f"{label} is not ASCII custody path: {value}")
    return pathlib.Path(*p.parts)

def resolve(value: Any, label: str) -> pathlib.Path:
    rel=safe_rel(value,label)
    path=PACKET/rel
    current=PACKET
    for part in rel.parts:
        current=current/part
        if current.is_symlink():
            raise EvidenceError(f"{label} traverses symlink: {value}")
    resolved=path.resolve()
    try: resolved.relative_to(PACKET.resolve())
    except ValueError as e: raise EvidenceError(f"{label} resolves outside packet") from e
    return path

def digest(path: pathlib.Path) -> tuple[int,str]:
    if not path.is_file() or path.is_symlink(): raise EvidenceError(f"not a regular file: {path}")
    h=hashlib.sha256(); n=0
    with path.open("rb") as f:
        for b in iter(lambda:f.read(1<<20),b""): h.update(b); n+=len(b)
    return n,h.hexdigest()

def descriptor(value: Any, label: str, *, relative=True) -> dict[str,Any]:
    if not isinstance(value,dict): raise EvidenceError(f"{label} is not an object")
    raw=value.get("path")
    path=resolve(raw,label+".path") if relative else pathlib.Path(raw)
    n,h=digest(path)
    if value.get("bytes") != n or value.get("sha256") != h:
        raise EvidenceError(f"{label} digest differs for {raw}")
    if not isinstance(value.get("sha256"),str) or len(value["sha256"])!=64 or set(value["sha256"]) - HEX:
        raise EvidenceError(f"{label} digest is malformed")
    return value

def source_manifest(plan:dict[str,Any]) -> dict[str,Any]:
    md=descriptor(plan["source_manifest"],"source_manifest")
    mp=resolve(md["path"],"source_manifest.path")
    m=load_json(mp)
    if m.get("schema") != "litchi-xlsx-profile-source-closure-v1": raise EvidenceError("source closure schema differs")
    if m.get("source_root") != plan.get("source_root"): raise EvidenceError("source root differs")
    if m.get("origin_commit") != plan.get("head"): raise EvidenceError("source commit differs")
    if m.get("file_count") != len(m.get("files",[])): raise EvidenceError("source file count differs")
    root=resolve(plan["source_root"],"source_root")
    files=m.get("files")
    if not isinstance(files,list) or not files: raise EvidenceError("source closure is empty")
    manifest_paths=[]
    for item in files:
        if not isinstance(item,dict): raise EvidenceError("source file entry is not an object")
        manifest_paths.append(str(safe_rel(item.get("path"),"source manifest path")))
    if len(set(manifest_paths)) != len(manifest_paths): raise EvidenceError("source manifest has duplicate paths")
    actual_paths=[]
    for candidate in root.rglob("*"):
        if candidate.is_symlink(): raise EvidenceError(f"source closure contains symlink: {candidate}")
        if candidate.is_file(): actual_paths.append(str(candidate.relative_to(root)))
    if set(actual_paths) != set(manifest_paths):
        missing=sorted(set(manifest_paths)-set(actual_paths)); extra=sorted(set(actual_paths)-set(manifest_paths))
        raise EvidenceError(f"source manifest path set differs (missing={missing[:3]}, extra={extra[:3]})")
    for index,item in enumerate(files):
        if not isinstance(item,dict): raise EvidenceError(f"source file {index} is not an object")
        rel=safe_rel(item.get("path"),f"source file {index}")
        if not str(rel).startswith(('.cargo'+os.sep,'crates'+os.sep,'test-data'+os.sep,'tools'+os.sep)) and str(rel) not in {"Cargo.toml","README.md","clippy.toml","deny.toml","rust-toolchain.toml","rustfmt.toml"}:
            raise EvidenceError(f"source file outside frozen closure: {rel}")
        path=(root/rel).resolve()
        try: path.relative_to(root.resolve())
        except ValueError as e: raise EvidenceError(f"source file escapes root: {rel}") from e
        n,h=digest(path)
        if item.get("kind") != "file" or item.get("bytes") != n or item.get("sha256") != h:
            raise EvidenceError(f"source file digest differs: {rel}")
        if item.get("mode") != stat.S_IMODE(path.stat().st_mode): raise EvidenceError(f"source mode differs: {rel}")
    return m

def check(plan_path: pathlib.Path) -> dict[str,Any]:
    plan=load_json(plan_path)
    if plan.get("schema") != SCHEMA: raise EvidenceError("plan schema differs")
    if plan.get("head") != "933ee45cbb5d2ef9dc29583f7ec890ad234e04fb": raise EvidenceError("authoritative HEAD differs")
    if plan.get("source_root") != "source-build": raise EvidenceError("source root differs")
    if plan.get("environment") != RUNTIME_ENV: raise EvidenceError("runtime environment is not frozen clean env")
    if plan.get("cpu") != 12 or plan.get("case") != "xlsx_source_backed_cell_values_one_percent_edit_save": raise EvidenceError("profile identity differs")
    if plan.get("shapes") != ["medium","dense-sparse"] or plan.get("repeats") != 2 or plan.get("samples") != 1 or plan.get("warmup") != 0: raise EvidenceError("sample policy differs")
    options=plan.get("options")
    expected=["--tool=callgrind","--collect-atstart=no","--toggle-collect="+plan.get("owner", ""),"--zero-before="+plan.get("owner", ""),"--dump-after="+plan.get("owner", "")]
    if options != expected: raise EvidenceError("Callgrind options differ")
    sched=plan.get("schedule")
    expected_sched=[{"index":0,"repeat":0,"shape":"medium"},{"index":1,"repeat":0,"shape":"dense-sparse"},{"index":2,"repeat":1,"shape":"dense-sparse"},{"index":3,"repeat":1,"shape":"medium"}]
    if sched != expected_sched: raise EvidenceError("schedule differs")
    for key in ("source_manifest","full_source_manifest","build","build_start","build_process_group_audit","binary","source_checker","semantic_validator","reader","raw_analyzer","capture_runner","raw_adapter","build_runner","audit_controls","audit_control_results","retained_helpers","native_context","prior_semantic_validator"):
        descriptor(plan.get(key),key)
    for key,value in plan.get("tools",{}).items(): descriptor(value,"tool."+key,relative=False)
    source_manifest(plan)
    accepted=plan.get("accepted_0905")
    if not isinstance(accepted,dict): raise EvidenceError("accepted 0905 binding is missing")
    accepted_path=resolve(accepted.get("path"),"accepted_0905.path")
    accepted_n,accepted_h=digest(accepted_path)
    if accepted_h != accepted.get("sha256") or accepted_h != "eddc437156c976c55ce36b0bf8f0ef47cc4943d296735a1ac8d6ef120b8754e7": raise EvidenceError("accepted 0905 audit source binding differs")
    custody=load_json(resolve(plan["retained_helpers"]["path"],"retained_helpers.path"))
    helper_entries=custody.get("files")
    if not isinstance(helper_entries,list) or len(helper_entries)!=6: raise EvidenceError("retained helper custody cardinality differs")
    seen=set()
    for item in helper_entries:
        if not isinstance(item,dict) or item.get("path") in seen: raise EvidenceError("retained helper custody has duplicate/malformed entry")
        seen.add(item.get("path")); descriptor(item,"retained_helper")
    build_start=load_json(resolve(plan["build_start"]["path"],"build_start.path"))
    if build_start.get("schema") != "litchi-xlsx-profile-build-start-v1" or build_start.get("argv")[:2] != ["/home/zhuhe/.cargo/bin/cargo","build"] or build_start.get("cwd_relative") != "source-build" or build_start.get("deadline_seconds") != 900 or build_start.get("process_group") is not True: raise EvidenceError("build start receipt differs")
    build_runner=descriptor(plan["build_runner"],"build_runner")
    if build_start.get("runner") != build_runner.get("path") or build_start.get("runner_sha256") != build_runner.get("sha256"): raise EvidenceError("build runner custody differs")
    if build_start.get("env",{}).get("CARGO_NET_OFFLINE") != "true" or build_start.get("env",{}).get("CARGO_BUILD_JOBS") != "2" or build_start.get("unset_env") != ["RUSTFLAGS","CARGO_ENCODED_RUSTFLAGS","RUSTC_WRAPPER","CARGO_BUILD_RUSTC_WRAPPER"]: raise EvidenceError("build environment binding differs")
    if build_start.get("inputs",{}).get("source_manifest_sha256") != plan["source_manifest"]["sha256"]: raise EvidenceError("build source manifest binding differs")
    build=load_json(resolve(plan["build"]["path"],"build.path"))
    if build.get("schema") != "litchi-xlsx-profile-build-terminal-v1" or build.get("returncode") != 0 or build.get("timed_out") is not False or build.get("terminal_before_parse") is not True or build.get("process_group_id") is None: raise EvidenceError("build terminal receipt is not successful")
    if build.get("binary",{}).get("sha256") != plan["binary"].get("sha256") or build.get("binary",{}).get("bytes") != plan["binary"].get("bytes"): raise EvidenceError("build binary binding differs")
    process_audit=load_json(resolve(plan["build_process_group_audit"]["path"],"build_process_group_audit.path"))
    if process_audit.get("schema") != "litchi-xlsx-profile-build-process-group-audit-v1" or process_audit.get("process_group_clear") is not True or process_audit.get("pid_proc_entry_exists") is not False or process_audit.get("terminal",{}).get("sha256") != plan["build"]["sha256"]: raise EvidenceError("build process-group audit is not clear/bound")
    if plan["binary"].get("path") == plan["source_root"] or pathlib.PurePosixPath(plan["binary"]["path"]).is_absolute(): raise EvidenceError("binary custody path must be relative")
    audit=load_json(resolve(plan["audit_control_results"]["path"],"audit_control_results.path"))
    if audit.get("schema") != "litchi-xlsx-profile-audit-controls-v2" or audit.get("status") != "pass" or not isinstance(audit.get("control_count"),int) or audit["control_count"] < 14: raise EvidenceError("audit control result is incomplete")
    if plan.get("expected_measured_calls") != 1 or plan.get("historical_expected_lifecycle_calls") != 5: raise EvidenceError("owner call policy differs")
    return {"status":"pass","schema":SCHEMA,"source_file_count":len(source_manifest(plan)["files"]),"source_manifest_sha256":plan["source_manifest"]["sha256"],"binary_sha256":plan["binary"]["sha256"]}

def main():
    path=pathlib.Path(sys.argv[1]) if len(sys.argv)>1 else PACKET/'docs/performance/results/change-0944/plan.json'
    try: print(json.dumps(check(path),sort_keys=True))
    except (EvidenceError,OSError) as e: print(f"source_check_v2: {e}",file=sys.stderr); return 1
    return 0
if __name__=='__main__': raise SystemExit(main())
