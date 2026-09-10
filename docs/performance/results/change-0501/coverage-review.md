# CRUD baseline coverage review: change 0501

This is a read-only audit of the current CRUD index, its checked identity and
the `tools/perf-baseline` full-run contract. It also documents the owned
capture runner in `baseline-capture.py`. No baseline lane, build, or source
edit was run while preparing this review.

## Current gap

The checked default identity is 37 cases, 201 case/corpus rows and 31
deterministic corpora. The representative non-iWork index has 15 categories
and 33 mapped selectors: 11 measured mappings and 22 correctness-only
mappings. The measured mappings currently account for 48 measured rows. The
checked catalog is `litchi-perf-corpus-v2`; its current checked catalog SHA-256
is
`d2c35126ee4e862ada539944ddb6cc2c654b82fe1e1465f505034fd1a9f7a84f`.

Those measured rows are the eight-corpus `cfb_list_streams`, `cfb_read_one`
and `opc_open` families plus three-shape rows for
`xlsx_full_cell_scan`, `xlsx_list_sheets`, `doc_fresh_write_to`,
`xls_fresh_write_to`, `ppt_fresh_write_to`, `xlsx_one_cell_commit`,
`xlsx_one_percent_commit` and `odp_existing_append_lifecycle`. The remaining
mapped scenarios retain correctness-only evidence or an explicit unsupported
scope; they are not silently promoted by a full default run.

The contract-only check succeeds:

```text
python3 tools/validate_crud_coverage_index.py
validated non-iWork CRUD coverage index: 15 categories, 33 mapped selectors (contract-only; no run timing report supplied)
```

The timing gate remains open because the expected ignored report
`target/perf/container-baseline.json` is absent. The explicit report check
therefore fails with a missing-report error. A static index pass does not
create timing evidence or promote any mapping. The current run must be
validated against the report and the generated schema-2 catalog sidecar, then
the checked catalog remains the identity reference for review.

The immediate blocker is source/binary finalization, not the coverage harness:
the normal executable currently being built is a candidate while the PPTX
controls and optimization review are still in progress. A run against it now
would produce useful descriptive timing, but it would not be authoritative for
the final candidate and would require a second two-lane run after any source
change. Once the final executable and matching source manifest are frozen, the
two serial lanes materially close the missing current full-run baseline gate;
they do not promote the representative index into exhaustive CRUD coverage.

## Required full run

The final candidate executable must be frozen before capture. Its source hash
manifest is a required second argument to the runner. The manifest must contain
a lowercase forty-character candidate base identity under `revision`,
`candidate_base`, or `candidate_base_revision`, plus a `files` map of
repository-relative paths to SHA-256 values. The runner rehashes every listed
file before and after each lane and after the static tests. It records the
candidate base and the complete hash map. The observed Git HEAD is recorded as
host state only; it is not used to invent a candidate revision.

From the repository root, the final command is:

```sh
python3 -B docs/performance/results/change-0501/baseline-capture.py \
  /absolute/path/to/frozen/litchi-perf-baseline \
  /absolute/path/to/final-candidate-source-manifest.json
```

The driver runs these two lanes serially:

```text
/usr/bin/time -v -o docs/performance/results/change-0501/captures/R1-normal/resource.log \
  taskset -c 2 /absolute/path/to/frozen/litchi-perf-baseline \
  --workers 1 --samples 15 --warmup 3 \
  --json docs/performance/results/change-0501/captures/R1-normal/report.json \
  --corpus-manifest docs/performance/results/change-0501/captures/R1-normal/corpus-catalog.json
```

`R2-normal` has the same arguments and fresh output paths. There is no
`--case`, shape, payload, writer-shape or other selector restriction, so each
report must contain all 201 rows, all 37 default cases and 15 retained samples
per row after three warmups. The sidecar must contain 201 bindings and schema
version 2. Each lane retains `started.json`, the exact command, report,
catalog, GNU `time -v` resource log, stdout/stderr, validator output and an
exclusive receipt. The runner then executes the static test module with the
correct package form:

```sh
python3 -B -m unittest tools.test_crud_coverage_index
```

The runner passes each generated lane catalog explicitly to the repository
gate:

```sh
python3 -B tools/validate_crud_coverage_index.py \
  --catalog docs/performance/results/change-0501/captures/R1-normal/corpus-catalog.json \
  --report docs/performance/results/change-0501/captures/R1-normal/report.json
```

Using the generated sidecar is required: a current run can have a different
provenance `catalog_sha256` while preserving the checked `content_set_sha256`.
The report's catalog reference must match its own sidecar exactly, and the
report must still satisfy the checked 37/201 identity and binary hash gates.

Every child receives `RUSTUP_TOOLCHAIN=1.98.1`, an empty `DEBUGINFOD_URLS`,
`PYTHONDONTWRITEBYTECODE=1`, `LC_ALL=C`, and the explicit owned temporary
directory:

```text
TMPDIR=/tmp/litchi-goal-0501/default-corpora
TMP=/tmp/litchi-goal-0501/default-corpora
TEMP=/tmp/litchi-goal-0501/default-corpora
```

The runner creates that directory with an ownership marker immediately before
each lane, refuses to start if it already exists, removes only the marked
directory after the child exits, and records that the path is absent. A stale
directory or missing marker is a hard failure. No build is part of this
command.

## Resource and storage envelope

Historical 0465 normal full runs used the same one-worker, CPU-2,
three-warmup/fifteen-sample, 201-row protocol. They took about 53.6 seconds
per lane, reported 161,524 KiB and 152,224 KiB maximum RSS for the two normal
runs, and retained reports of about 1.16 MiB and catalogs of about 222 KiB.
The two current lanes should therefore be budgeted at roughly two minutes of
serial child time plus Python validation and cleanup, with at least a few MiB
of retained evidence. These are historical planning values, not measurements
of the final candidate or a current performance claim.

The repository had previously approached critically low free space during
large Rust builds. The parent run uses the dedicated target
`/tmp/litchi-goal-0501-target`; this runner only writes the two lane trees,
the summary/static-test logs and the explicit temporary corpus directory. The
temporary corpus directory is removed after each lane. Keep the lane reports,
catalogs and GNU-time logs because they are the portable evidence; do not
substitute a regenerated ignored report after capture.

## Historical distinction and decision

0465 is retained descriptive evidence from an older normal binary bound to
revision `161cf53b20d8bb65fe79d4567b9ea7768430de7b`. Its two full reports are
useful for protocol shape, row/sample counts, resource scale and historical
context, but they cannot satisfy the current-source report gate. The 0501
capture must use the final post-PPTX candidate executable and its matching
source hash manifest, even if that candidate has the same harness behavior.

If both lanes pass their report shape, binary identity, sidecar, source
custody, temporary-directory cleanup and repository-validator gates, this
materially closes the missing authoritative full-run baseline requirement for
the current candidate. It still does not complete `docs/GOAL.md`: the index is
representative rather than exhaustive, the run is serial and warm local
execution, and it makes no native-producer, cold-cache, remote/range, physical
I/O, allocation-local, worker-scaling, or whole-goal claim. The PPTX source
optimization decision remains separately dependent on its matched controls,
profile and semantic review.
