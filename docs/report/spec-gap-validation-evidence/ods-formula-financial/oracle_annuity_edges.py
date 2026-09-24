#!/usr/bin/env python3
"""Independent Decimal oracle for finite annuity overflow regressions."""
import argparse
from decimal import Decimal, localcontext
import hashlib
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent
CORPUS = HERE / "oracle-annuity-edge-vectors.json"
FORMULAS = (
    "=PPMT(2;1;2;1e308;0;1)",
    "=IPMT(0.1;512;512;100;1;0)",
    "=IPMT(1e-10;1;2;1e308;0;0)",
    "=PPMT(1e-10;1;2;1e308;0;0)",
    "=PPMT(1e-320;1;2;100;0;0)",
    "=IPMT(1e-320;2;2;100;0;0)",
    "=PPMT(1e-10;1;2;1e308;1e308;0)",
    "=IPMT(1e-10;2;2;1e308;0;0)",
    "=IPMT(1e-10;2;10;1e308;1e308;0)",
)


def document():
    rows = []
    with localcontext() as context:
        # More than 320 digits are needed to retain 1 + a subnormal rate.
        context.prec = 800
        for formula in FORMULAS:
            function, arguments = formula[1:-1].split("(")
            rate, period, periods, pv, fv, due = (
                Decimal.from_float(float(value)) for value in arguments.split(";")
            )
            base = 1 + rate
            growth = base ** int(periods)
            payment = -(pv * growth + fv) * rate / ((growth - 1) * (1 + rate * due))
            prior_growth = base ** (int(period) - 1)
            balance = pv * prior_growth + payment * (1 + rate * due) * (prior_growth - 1) / rate
            interest = -rate * balance
            if due == 1:
                interest = Decimal(0) if period == 1 else interest / base
            result = interest if function == "IPMT" else payment - interest
            rows.append({"formula": formula, "expected_binary64": repr(float(result))})
    return {
        "schema": "financial-annuity-edge-oracle-v1",
        "precision_digits": 800,
        "basis": "Forward payment and balance equations with exact binary64 inputs; no production imports.",
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
    elif json.loads(CORPUS.read_text()) != expected:
        raise SystemExit("annuity edge corpus differs from regenerated oracle")
    else:
        print(f"verified {len(expected['vectors'])} annuity edge vectors")


if __name__ == "__main__":
    main()
