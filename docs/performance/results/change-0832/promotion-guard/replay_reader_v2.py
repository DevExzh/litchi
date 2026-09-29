"""Add the omitted source constant and write-once JSON helper to frozen replay.

Version 1 and both failed admissions remain retained. Neither correction
changes report parsing, statistics, workloads, or regression thresholds.
"""
import importlib.util
from pathlib import Path
import symtable
import builtins
import sys

import replay_reader as previous

P = Path(__file__).resolve().parent
sha, read, write = previous.sha, previous.read, previous.write


def binding():
    original = previous.binding()
    assert read(P / "reader-correction.json") == original
    failure = P / "attempts/08-replay_reader"
    assert read(failure / "receipt.json")["exit_code"] == 1
    assert "NameError: name 'write_once' is not defined" in (failure / "output.log").read_text()
    table = symtable.symtable((P / "reader.py").read_text(), "reader.py", "exec")
    known = {s.get_name() for s in table.get_symbols() if s.is_assigned() or s.is_imported()}
    known |= set(dir(builtins)) | {"__file__", "__name__"}
    missing = set()
    def scan(scope):
        missing.update(s.get_name() for s in scope.get_symbols()
                       if s.is_global() and s.is_referenced() and s.get_name() not in known)
        for child in scope.get_children():
            scan(child)
    scan(table)
    assert missing == {"SOURCES", "write_once"}
    return {**original, "schema": "litchi.performance.0832.promotion-reader-correction.v2",
            "reason": "Supply only missing SOURCES and write_once globals; frozen validation/statistics remain unchanged.",
            "previous_loader_sha256": original["loader_sha256"],
            "previous_correction_sha256": sha(P / "reader-correction.json"),
            "loader_sha256": sha(Path(__file__)), "missing_globals": sorted(missing),
            "second_failure_receipt_sha256": sha(failure / "receipt.json"),
            "second_failure_log_sha256": sha(failure / "output.log")}


def main():
    args = sys.argv[1:]
    assert args in ([], ["prepare"], ["--admit"], ["--check"])
    correction = P / "reader-correction-v2.json"
    expected = binding()
    spec = importlib.util.spec_from_file_location("frozen_promotion_reader", P / "reader.py")
    reader = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(reader)
    assert not hasattr(reader, "SOURCES") and not hasattr(reader, "write_once")
    reader.SOURCES = previous.SOURCES
    reader.write_once = write
    if args == ["prepare"]:
        assert not (P / "capture.json").exists() and not (P / "native").exists()
        frozen, _, _ = reader.verify_freeze()
        reader.load_qualification_reports(frozen)
        write(correction, expected)
        print("0832 correction v2 prepared; all qualification records validate; frozen reader unchanged")
        return
    assert read(correction) == expected
    assert reader.main(args) == 0
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
