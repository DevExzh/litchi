# Baseline capture notes

The report's `bindings` column is the total number of connection bindings in the
fixture. The raw row field `validation.reference_count` is narrower: it counts
bindings pointing to the first storage (`uid-0000`) before a rename, or to the
renamed storage (`uid-renamed-0000`) after the rename. The fixture sends one
binding to each non-first storage until those storages are covered, then sends
remaining bindings to the first storage. Therefore the large fixture has 512
total bindings and 449 first-storage references; medium has 128 total and 113
first-storage references; small has 16 total and 15 first-storage references.

The no-op lane produced an empty transaction (`changed == false` and an empty
patch), and all 45 serialized output hashes matched their source fixture hash.
The remove/inverse lane also restored the source fixture hash for all 45 rows.
The per-sample JSON rows were checked against their JSONL rows before the
redundant JSON files were removed. Empty stderr files were removed; the retained
raw JSONL, `/usr/bin/time -v` sidecars, and all other result files are listed in
`retained-files.sha256`.
