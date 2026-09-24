#!/usr/bin/env python3
"""Run standalone numerical solver probes, without claiming evaluator coverage.

Pass the checkout containing the solver as the sole argument. Temporary files
are removed on success and failure. The source is snapshotted before rustc runs.
"""

import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile


HARNESS = r'''#![allow(dead_code)]
#[derive(Clone,Copy,Debug,PartialEq,Eq)]enum ScalarError{Number,Value,DivisionByZero}
#[derive(Debug)]enum EvaluationFailure{Cancelled}
type EvaluationResult<T>=Result<T,EvaluationFailure>;
mod codec {pub mod formula {pub mod evaluation {
use crate::{ScalarError,EvaluationFailure,EvaluationResult};
mod financial {#[path=SOLVER_PATH]mod solver;
#[test]fn two_cashflow_scale_invariance(){
 for scale in [1e-300_f64,1.0,1e300]{for rate in [-0.9_f64,-0.5,0.0,0.1,1.0,100.0,1e8]{
 let values=[-scale,scale*(1.0+rate)];
 let actual=solver::irr(&values,0.1,&mut |_|Ok(())).unwrap().unwrap();
 assert!((actual-rate).abs()<=1e-10*(1.0+rate.abs()),"scale={scale},rate={rate},actual={actual}");
 }}}
#[test]fn dated_two_cashflows_scale_invariance(){
 for scale in [1e-300_f64,1.0,1e300]{for rate in [-0.9_f64,-0.5,0.0,0.1,1.0,100.0]{
 let values=[-scale,scale*(1.0+rate)];
 let actual=solver::xirr(&values,&[0.0,365.0],0.1,&mut |_|Ok(())).unwrap().unwrap();
 assert!((actual-rate).abs()<=1e-10*(1.0+rate.abs()),"scale={scale},rate={rate},actual={actual}");
 }}}
#[test]fn rate_minus_one_boundary(){assert_eq!(solver::rate(2.0,-50.0,100.0,50.0,0.0,-1.0,&mut |_|Ok(())).unwrap(),Ok(-1.0));}
}}}}
'''


def main():
    if len(sys.argv) != 2:
        raise SystemExit("usage: probe_solver.py CHECKOUT")
    source = Path(sys.argv[1]) / "crates/litchi-ods/src/codec/formula/evaluation/financial/solver.rs"
    snapshot = source.read_bytes()
    print("solver SHA256:", hashlib.sha256(snapshot).hexdigest(), flush=True)
    with tempfile.TemporaryDirectory(prefix="litchi-financial-solver-probe-") as directory:
        temporary = Path(directory)
        solver = temporary / "solver.rs"
        solver.write_bytes(snapshot)
        harness = temporary / "probe.rs"
        harness.write_text(HARNESS.replace("SOLVER_PATH", json.dumps(str(solver))))
        binary = temporary / "probe"
        subprocess.run(
            ["rustc", "--edition=2024", "--test", str(harness), "-o", str(binary)],
            check=True,
        )
        subprocess.run([str(binary), "--quiet"], check=True)


if __name__ == "__main__":
    main()
