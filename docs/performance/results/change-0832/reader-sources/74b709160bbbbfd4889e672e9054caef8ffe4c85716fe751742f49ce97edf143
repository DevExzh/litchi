"""Exercise current source/plan/install custody before comparative capture."""
import analyze as a
import audit as independent
import driver as d


def main():
    d.check("after")
    a.require_plan()
    a.require_inputs()
    a.require_install()
    independent.validate_plan()
    custody = independent.validate_custody()
    assert custody["normative"]["count"] == 35
    print("0832 custody preflight PASS: four source transitions and 35 normative files")


if __name__ == "__main__":
    main()
