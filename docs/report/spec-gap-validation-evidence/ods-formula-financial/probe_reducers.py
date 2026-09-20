#!/usr/bin/env python3
"""Replay snapshotted reducer unit tests without claiming evaluator coverage."""
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile


def main():
    if len(sys.argv) != 2:
        raise SystemExit("usage: probe_reducers.py CHECKOUT")
    base = Path(sys.argv[1]) / "crates/litchi-ods/src/codec/formula/evaluation"
    with tempfile.TemporaryDirectory(prefix="litchi-financial-reducer-probe-") as directory:
        temporary = Path(directory)
        for name, source in (("numerics.rs", base / "numerics.rs"),
                             ("reducers.rs", base / "financial/reducers.rs")):
            snapshot = source.read_bytes()
            (temporary / name).write_bytes(snapshot)
            print(name, hashlib.sha256(snapshot).hexdigest(), flush=True)
        harness = """#![allow(dead_code)]
mod codec {pub mod formula {pub mod evaluation {
#[derive(Clone,Copy,Debug,PartialEq,Eq)]enum ScalarError{Number,Value,DivisionByZero}
#[path=NUMERICS_PATH] mod numerics;
mod financial {#[path=REDUCERS_PATH] mod reducers;}
}}}
"""
        harness = harness.replace("NUMERICS_PATH", json.dumps(str(temporary / "numerics.rs")))
        harness = harness.replace("REDUCERS_PATH", json.dumps(str(temporary / "reducers.rs")))
        source = temporary / "probe.rs"
        source.write_text(harness)
        binary = temporary / "probe"
        subprocess.run(["rustc", "--edition=2024", "--test", str(source), "-o", str(binary)], check=True)
        subprocess.run([str(binary), "--quiet"], check=True)


if __name__ == "__main__":
    main()
