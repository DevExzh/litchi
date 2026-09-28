# Offline reader recovery

The first failure-aware reader pass was retained under `reader-0/`. The
quality-1 input snapshot stores the allowlisted test as the flattened file
`quality-1/input-snapshot/xlsx_planning_allocations.rs`; the reader initially
looked for the relative source path below that directory. The reader now uses
that retained basename and still binds its digest to the quality-1 source
witness.

The next pass was retained under `reader-1/`. `load_quality_recovery` used a
missing local `custody_value` name while checking the old frozen inputs. Its
caller now passes the current custody value explicitly, so the recovery check
continues to compare the archived root inputs, locks, corpus, architecture,
host, and unrelated-file witness.

After these bounded corrections, `analyze.py --write`, `analyze.py --check`,
and `validate.py` all completed successfully. No workload, Cargo command, or
frozen audit/driver was changed by these reader recoveries.
