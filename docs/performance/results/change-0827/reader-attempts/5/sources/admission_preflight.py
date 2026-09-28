"""Execution-free qualification admission preflight for the 0827 packet."""

from __future__ import annotations

from pathlib import Path
import sys


P = Path(__file__).resolve().parent
if str(P) not in sys.path:
    sys.path.insert(0, str(P))

import admission  # noqa: E402


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    assert args in ([], ["--write"])
    value = admission.qualification_preflight()
    assert value["status"] == "pass"
    assert value["reports"] == 13
    assert value["sample_envelopes"] == 13
    assert value["new_measurements"] == 0
    print("0827 admission preflight PASS: 12 prior qualification reports plus retained 0825 failure")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
