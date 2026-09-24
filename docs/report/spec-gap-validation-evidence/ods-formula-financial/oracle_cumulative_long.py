#!/usr/bin/env python3
"""Independent long-span cumulative oracle; no production evaluator imports."""
import argparse
import decimal
import hashlib
import importlib.util
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent
CORPUS = HERE / "oracle-cumulative-long-vectors.json"


def document():
    script = HERE / "oracle_reducers.py"
    spec = importlib.util.spec_from_file_location("financial_reducer_oracle", script)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    rows = []
    with decimal.localcontext() as context:
        context.prec = 256
        for periods in (64, 512, 2048):
            for payment_type in (0, 1):
                arguments = {"rate": "0.1", "periods": str(periods), "value": "100",
                             "start": "1", "end": str(periods - 1), "type": str(payment_type)}
                for function in ("CUMIPMT", "CUMPRINC"):
                    expected = module.FUNCTIONS[function](arguments)
                    rows.append({
                        "function": function,
                        "formula": f"={function}(0.1;{periods};100;1;{periods - 1};{payment_type})",
                        "inputs": arguments.copy(), "expected": str(expected),
                        "relative_tolerance": "1e-12", "absolute_tolerance": "0",
                    })
    return {
        "schema": "financial-cumulative-long-span-oracle-v1",
        "precision_digits": 256,
        "basis": "Independent Decimal evaluation of the documented per-period IPMT/PPMT equations using binary64-converted inputs; no production kernel calls.",
        "oracle_script_sha256": hashlib.sha256(script.read_bytes()).hexdigest(),
        "contract_sha256": hashlib.sha256((HERE / "contract.md").read_bytes()).hexdigest(),
        "vectors": rows,
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--write", action="store_true")
    group.add_argument("--check", action="store_true")
    args = parser.parse_args()
    expected = document()
    if args.write:
        CORPUS.write_text(json.dumps(expected, indent=2) + "\n")
    else:
        actual = json.loads(CORPUS.read_text())
        if actual != expected:
            raise SystemExit("long-span cumulative corpus differs from regenerated oracle")
        print(f"verified {len(expected['vectors'])} long-span cumulative vectors")


if __name__ == "__main__":
    main()
