# Accepted candidate capture notes

This capture uses the corrected frozen source. The reserve-ceiling diagnostic is retained separately in `candidate-pre-reserve-fix`; it is excluded from the matched comparison. The corrected source uses fallible one-edge-at-a-time growth in `capture_edges`, avoiding a capacity reservation proportional to the configured relationship limit.

The report's `bindings` column is the total number of connection bindings in the fixture. The raw row field `validation.reference_count` counts bindings pointing to the first storage (`uid-0000`) before a rename, or to `uid-renamed-0000` after the rename. The large fixture has 512 total bindings and 449 first-storage references; medium has 128 total and 113 first-storage references; small has 16 total and 15 first-storage references.

Each raw JSONL row was checked against its per-sample JSON before redundant JSON files were removed. Empty stderr files were removed; the retained raw JSONL, `/usr/bin/time -v` sidecars, and result files are listed in `retained-files.sha256`.
