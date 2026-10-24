# 0362ul10e — OLE owner child PATH binding repair

- Date: 2026-10-24.
- Parent packet: 0362ul10d.
- Status: source-only false-acceptance repair; native execution and acceptance
  remain pending coordinator review.

The paired runner now assigns `PATH` from the selected toolchain receipt when
constructing every build and ABBA child environment.  Previously it recorded
the selected PATH but copied a later ambient `os.environ`, allowing a PATH
mutation after toolchain selection to change child lookup while post-run
validation only noticed the mismatch.  The focused runner test now changes
ambient PATH after selection and proves the child retains the locked value.

Source-only checks performed:

```text
python3 -m unittest tools.test_ole_owner_paired_runner tools.test_validate_ole_owner_run -q
python3 -m py_compile tools/native_qualification/ole_owner_paired_runner.py tools/native_qualification/validate_ole_owner_run.py tools/test_ole_owner_paired_runner.py tools/test_validate_ole_owner_run.py
git diff --check
```

No Cargo build, native binary, timing result, or performance acceptance claim
is made by this record.
