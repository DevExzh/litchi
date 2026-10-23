#!/usr/bin/env python3
"""Create the external Pages footnote workload from tracked provenance.

The source package is intentionally kept in ``test-data/iwork``.  The
qualification fixture directory is disposable and external to the checkout;
this small deterministic generator gives the synthetic workload a reviewed
input identity without adding an unreviewable generated package to git.  The
owner probe itself performs the body-footnote insertion and same-family
no-op, so this copy is only fixture preparation.
"""

from __future__ import annotations

import argparse
from pathlib import Path
import shutil


ROOT = Path(__file__).resolve().parents[4]
SOURCE = ROOT / "test-data/iwork/pages/basic.pages"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, default=SOURCE)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    source = args.input.resolve(strict=True)
    output = args.output.resolve()
    if source != SOURCE.resolve():
        raise SystemExit(f"input must be the tracked Pages fixture: {SOURCE}")
    if output.is_relative_to(ROOT):
        raise SystemExit(f"output must be outside the checkout: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source, output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
