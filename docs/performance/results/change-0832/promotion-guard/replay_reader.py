"""Supply the omitted SOURCES constant without modifying the frozen reader.

This additive correction changes no report parser, statistic, fixture, workload,
or threshold. Every corrected admission/analysis is separately hash-bound.
"""
import hashlib
import importlib.util
import json
from pathlib import Path
import sys

P = Path(__file__).resolve().parent
SOURCES = (
    "crates/litchi-xlsx/src/column.rs",
    "crates/litchi-xlsx/src/raw/worksheet/codec.rs",
    "crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/write/columns.rs",
    "crates/litchi-xlsx/src/raw/worksheet/edit/validation.rs",
)


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read(path):
    return json.loads(path.read_text())


def write(path, value):
    with path.open("x") as stream:
        json.dump(value, stream, indent=2, sort_keys=True)
        stream.write("\n")


def binding():
    frozen = read(P / "freeze.json")
    assert sha(P / "reader.py") == frozen["scripts"]["reader.py"]["sha256"]
    assert all(set(frozen["source_archives"][leg]) == set(SOURCES) for leg in ("before", "after"))
    failure = P / "attempts/06-reader"
    assert read(failure / "receipt.json")["exit_code"] == 1
    assert "NameError: name 'SOURCES' is not defined" in (failure / "output.log").read_text()
    return {"schema": "litchi.performance.0832.promotion-reader-correction.v1",
            "status": "pass", "sources": list(SOURCES),
            "reason": "Supply only the missing SOURCES global; frozen parsing/statistics/workload code is unchanged.",
            "freeze_sha256": sha(P / "freeze.json"),
            "reader_sha256": sha(P / "reader.py"),
            "loader_sha256": sha(Path(__file__)),
            "qualification_sha256": sha(P / "qualification.json"),
            "failure_receipt_sha256": sha(failure / "receipt.json"),
            "failure_log_sha256": sha(failure / "output.log")}


def main():
    args = sys.argv[1:]
    assert args in ([], ["prepare"], ["--admit"], ["--check"])
    correction = P / "reader-correction.json"
    expected = binding()
    spec = importlib.util.spec_from_file_location("frozen_promotion_reader", P / "reader.py")
    reader = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(reader)
    assert not hasattr(reader, "SOURCES")
    reader.SOURCES = SOURCES
    if args == ["prepare"]:
        assert not (P / "capture.json").exists() and not (P / "native").exists()
        reader.verify_freeze()
        write(correction, expected)
        print("0832 additive reader correction prepared; frozen reader unchanged")
        return
    assert read(correction) == expected
    code = reader.main(args)
    assert code == 0
    if args == ["--admit"]:
        write(P / "corrected-admission.json", {"status": "pass",
              "correction_sha256": sha(correction),
              "admission_sha256": sha(P / "qualification-admission.json")})
    else:
        result = {"status": "pass", "correction_sha256": sha(correction),
                  "analysis_sha256": sha(P / "analysis.json"),
                  "corrected_admission_sha256": sha(P / "corrected-admission.json")}
        if args == ["--check"]:
            assert read(P / "corrected-analysis.json") == result
        else:
            write(P / "corrected-analysis.json", result)


if __name__ == "__main__":
    main()
