# Change 0955 evidence

See [the report](../../0955-xlsx-managed-compression-qualification.md).

From the repository root, build and run into a new directory:

```sh
CARGO_TARGET_DIR=/path/to/external-target CARGO_INCREMENTAL=0 \
  cargo +1.95.0 build --locked --offline \
  --manifest-path tools/perf-baseline/Cargo.toml \
  --example xlsx_managed_compression -j2
/path/to/external-target/debug/examples/xlsx_managed_compression /path/to/new-result
```

`manifest.json` is the final deterministic result. `plan.json` records the nine
prospective case expectations; plan workload IDs combine with their worker field,
while result IDs explicitly append the worker count. `validation.json` records
matrix validation, replay and overwrite refusal. `source-final.json` binds the
Git base plus the four harness overlay files, toolchain, build command and binary.
The complete source census exceeds 1 MiB and remains external at its named path,
with its size and SHA-256 committed in that binding. Regenerate it by walking the
recorded Git base plus overlay and hashing each file; `freeze.py` contains the
exact census procedure. Source revisions, intermediate receipts and binaries
remain under `/home/zhuhe/litchi-performance-0955`; no binaries are committed.

The final source came from `git archive` of the recorded base with only the four
listed harness files overlaid. All other snapshot files were independently
matched to their Git blob IDs. The original `source-v1` path is an integration
directory, not an assertion that every development build used its final bytes;
`revisions/` and the versioned input receipts retain the earlier states.

`run.py` is the command/log/deadline recorder. Build receipts use the original
external target path; closure retains the final executable at the binary path
recorded in `cleanup.json`. Follow its SHA-256 identity when locating it.
`review.md` approves the final module hash. `boundary.json` is the full main-tree
crate-topology check; this standalone example adds no Cargo dependency edges.
The stopped wrong-directory build and subsequent counter-placement revision are
recorded explicitly, and no timing result from either is used.
