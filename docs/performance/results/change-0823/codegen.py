"""Retain exact-binary Scene disassembly after all native work is terminal."""
import gzip
import hashlib
import json
from pathlib import Path
import subprocess
import time

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
SYMBOL = "litchi_pptx::shape::reader::Scene::read_with"


def digest(data):
    return hashlib.sha256(data).hexdigest()


def artifact(path):
    raw = path.read_bytes()
    return {"path": str(path.relative_to(ROOT)), "bytes": len(raw), "sha256": digest(raw)}


def main():
    for lane in ("native", "allocation"):
        assert json.loads((P / lane / "complete.json").read_text())["status"] == "pass"
    out = P / "codegen"
    assert not out.exists()
    out.mkdir()
    result = {"schema": "litchi.performance.0823.codegen.v1", "symbol": SYMBOL, "rows": []}
    for leg in ("before", "after"):
        binary = json.loads((P / f"build-{leg}/build.json").read_text())["binaries"]["real-native"]["artifact"]
        raw = Path(binary["path"]).read_bytes()
        assert len(raw) == binary["bytes"] and digest(raw) == binary["sha256"]
        raw_symbols, demangled_symbols = [], []
        commands = []
        for label, options in (("raw", []), ("demangled", ["-C"])):
            command = ["nm", "--defined-only", "--print-size", *options, binary["path"]]
            started = time.time()
            completed = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, cwd=ROOT)
            stderr = out / f"{leg}-{label}.stderr"
            stderr.write_bytes(completed.stderr)
            output = out / f"{leg}-{label}.nm.gz"
            output.write_bytes(gzip.compress(completed.stdout, mtime=0))
            commands.append({"argv": command, "started": started, "ended": time.time(),
                             "exit_code": completed.returncode, "stderr": artifact(stderr),
                             "stdout_gzip": artifact(output)})
            (out / f"{leg}-commands.json").write_text(json.dumps(commands, indent=2) + "\n")
            assert completed.returncode == 0
            rows = [s.split(maxsplit=3) for s in completed.stdout.decode().splitlines()]
            if label == "raw":
                raw_symbols = rows
            else:
                demangled_symbols = rows
        matches = [r for r in demangled_symbols if len(r) == 4 and r[3] == SYMBOL]
        assert len(matches) == 1, matches
        matched = matches[0]
        mangled = [r for r in raw_symbols if len(r) == 4 and r[:3] == matched[:3]]
        assert len(mangled) == 1, mangled
        command = ["objdump", "-d", "-C", "--no-show-raw-insn", f"--disassemble={mangled[0][3]}", binary["path"]]
        # GNU objdump applies --disassemble after demangling when -C is set.
        command[4] = f"--disassemble={SYMBOL}"
        started = time.time()
        completed = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, cwd=ROOT)
        assembly = out / f"{leg}.asm"
        assembly.write_bytes(completed.stdout)
        stderr = out / f"{leg}.stderr"
        stderr.write_bytes(completed.stderr)
        commands.append({"argv": command, "started": started, "ended": time.time(),
                         "exit_code": completed.returncode, "stdout": artifact(assembly), "stderr": artifact(stderr)})
        (out / f"{leg}-commands.json").write_text(json.dumps(commands, indent=2) + "\n")
        assert completed.returncode == 0
        text = completed.stdout.decode()
        assert f"<{SYMBOL}>:" in text and len(text.splitlines()) > 20
        calls = [line.strip() for line in text.splitlines() if "call" in line]
        wrappers = {name: [line for line in calls if name in line] for name in (
            "NsReader<R>::process_event", "NamespaceResolver::resolve_event", "NamespaceResolver::push")}
        result["rows"].append({"leg": leg, "binary": binary,
                               "raw_symbol": mangled[0], "demangled_symbol": matched,
                               "assembly": artifact(assembly), "calls": wrappers})
    (out / "result.json").write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    print("0823 exact-binary Scene disassembly retained")


if __name__ == "__main__":
    main()
