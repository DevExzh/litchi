#!/usr/bin/env python3
"""Compare the production fraction helper with Python's exact Fraction oracle."""
import hashlib
import json
import math
from pathlib import Path
import random
import struct
import subprocess
import sys
import tempfile
from fractions import Fraction

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
SOURCE = ROOT / "crates/litchi-ods/src/codec/formula/evaluation/text/fraction.rs"
SEED = 20260920


def digest(data):
    return hashlib.sha256(data).hexdigest()


def cases():
    rng = random.Random(SEED)
    values = [0.0, -0.0, 1.0, 0.5, math.nextafter(1.0, 0.0),
              math.ldexp(1.0, -1074), math.ldexp(1.0, -1022),
              0.11805555555555555, 0.13392857142857142,
              0.2361111111111111, 0.005050505050505051]
    values += [rng.random() for _ in range(256)]
    values += [math.ldexp(1.0 + rng.random(), -rng.randint(1, 74)) for _ in range(256)]
    values += [math.ldexp(rng.random(), -rng.randint(0, 1074)) for _ in range(256)]
    for bound in (1, 9, 99, 999, 9999, 99999, 999999):
        midpoint = 0.5 / bound
        for value in (*values, math.nextafter(midpoint, 0.0), midpoint,
                      math.nextafter(midpoint, 1.0)):
            yield value, bound


def main():
    before = digest(SOURCE.read_bytes())
    rows = list(cases())
    stdin = "".join(f"{struct.unpack('>Q', struct.pack('>d', value))[0]} {bound}\n"
                    for value, bound in rows)
    program = """use std::io::{self, BufRead};
#[derive(Debug)]
enum ScalarError { Number }
mod text {
    #[path = PATH]
    mod fraction;
    pub fn nearest(value: f64, bound: u64) -> Result<(u64, u64), super::ScalarError> {
        fraction::nearest(value, bound, |_| Ok::<(), ()>(())).expect("callback")
    }
}
fn main() {
    for line in io::stdin().lock().lines() {
        let line = line.expect("input");
        let mut fields = line.split_whitespace();
        let bits: u64 = fields.next().expect("bits").parse().expect("u64");
        let bound: u64 = fields.next().expect("bound").parse().expect("u64");
        let (p, q) = text::nearest(f64::from_bits(bits), bound).expect("valid fraction");
        println!("{p} {q}");
    }
}
""".replace("PATH", json.dumps(str(SOURCE)))
    compiler = subprocess.check_output(["rustc", "-Vv"], text=True).strip()
    with tempfile.TemporaryDirectory(prefix="litchi-text-fraction-root-") as directory:
        directory = Path(directory)
        source = directory / "main.rs"
        binary = directory / "probe"
        source.write_text(program)
        subprocess.run(["rustc", "--edition=2024", "-D", "warnings", "-O",
                        str(source), "-o", str(binary)], check=True)
        run = subprocess.run([str(binary)], input=stdin, capture_output=True,
                             text=True, check=True)
    output = run.stdout.splitlines()
    assert len(output) == len(rows)
    for (value, bound), line in zip(rows, output):
        expected = Fraction.from_float(value).limit_denominator(bound)
        actual = tuple(map(int, line.split()))
        assert actual == (expected.numerator, expected.denominator), (value.hex(), bound, actual, expected)
    assert before == digest(SOURCE.read_bytes()), "source changed during reproduction"
    receipt = {
        "verified": True, "cases": len(rows), "seed": SEED,
        "source_sha256": before, "verifier_sha256": digest(Path(__file__).read_bytes()),
        "input_sha256": digest(stdin.encode()), "output_sha256": digest(run.stdout.encode()),
        "python": sys.version, "rustc": compiler, "temporary_tree_cleaned": True,
        "scope": "Pure production helper compiled with an equivalent Number-error enum; exact Python Fraction comparison for this retained deterministic corpus, not an integration or universal rounding claim",
    }
    destination = HERE / "fraction-reproduction.json"
    if sys.argv[1:] == ["--check"]:
        retained = json.loads(destination.read_text())
        for key, value in receipt.items():
            if key not in ("python", "rustc"):
                assert retained[key] == value, key
    else:
        assert not sys.argv[1:], "only --check is supported"
        destination.write_text(json.dumps(receipt, indent=2) + "\n")
    print(json.dumps(receipt, sort_keys=True))


if __name__ == "__main__":
    main()
