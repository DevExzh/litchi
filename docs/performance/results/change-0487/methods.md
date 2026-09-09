# 0487 comparison method

`compare.py` retains a matched before/after measurement of the private OPC
`SpliceAuditReader` retention across short replay reads. It imports the immutable 0484
`measure_routes.py` and report validators from
`../change-0484/`; the 0484 files and their hashes are inputs to the new
protocol. The new lane has its own `support.py`, gate, protocol, capture
receipts, and analysis output. Importing the helper is read-only. It never
builds or captures unless an explicit capture subcommand is selected.

The matrix contains the following exact arms for each of
`s64-a64-short-c64`, `s64-a16384-short-c64`, and
`s131072-a64-short-c64`:

* deterministic input: `owned`, `file`, `short-read`, and `latency`;
* `memory_store` with owned input; and
* `file_store` with owned input.

That is 18 arms. Each arm runs with the normal and allocator binary, in
repeat 1's declared order and repeat 2's reverse order. Both the retained 0485
binary and the new 0487 binary are run, giving 18 × 2 × 2 × 2 = 144 formal
children. Every child uses 30 measured samples after 3 warmups. The optional
pilot lane uses the same phase, arm, and role inventory with 3 samples and 1
warmup; pilots are excluded from the formal summary.

## Custody and commands

The protocol is `comparison-protocol.json`. It records the selected arms and
run order, all canonical helper hashes, the old route protocol and source
manifest, the three old 0484 file-input fixture hashes, machine metadata, and
the expected build commands. The sealed `before` bindings point to
`change-0485/build-{normal,allocator}.json`. The
`after` bindings point to `change-0487/build-{normal,allocator}.json`. A
protocol may be frozen before the after build exists; in that case the after
record hash is deferred and is bound by each capture receipt when the build
becomes available. A non-deferred hash is immutable.

Use a unique attempt token for every new capture or analysis output:

```text
python3 -B compare.py plan
python3 -B compare.py freeze
python3 -B compare.py pilot-all --phase before --attempt pilot-before1
python3 -B compare.py pilot-all --phase after  --attempt pilot-after1
python3 -B compare.py capture-all --phase before --attempt formal-before1
python3 -B compare.py capture-all --phase after  --attempt formal-after1
# Single-child retry/diagnostic forms use the same frozen arm and build bindings:
python3 -B compare.py capture --phase after --attempt formal-after-one \
  --arm deterministic-owned-s64-a64-short-c64 --role normal --repeat 1
python3 -B compare.py pilot --phase after --attempt pilot-after-one \
  --arm deterministic-owned-s64-a64-short-c64 --role normal
python3 -B compare.py analyze \
  --before-attempt formal-before1 \
  --after-attempt formal-after1
```

The formal capture commands must be run through the new gate by the
coordinator, serially, so the gate's source snapshot and shared CPU lock cover
the entire command. The analyzer requires these exact gate receipt labels and
binds their command, helper hashes, environment, and source snapshot:

```text
python3 -B docs/performance/results/change-0487/gate.py --attempt formal1 \
  capture-before python3 -B docs/performance/results/change-0487/compare.py \
  capture-all --phase before --attempt formal1
python3 -B docs/performance/results/change-0487/gate.py --attempt formal1 \
  capture-after python3 -B docs/performance/results/change-0487/compare.py \
  capture-all --phase after --attempt formal1
```

The gate does not replace an existing receipt. `--timeout-seconds`
defaults to 1,800. A timeout terminates the whole child process group, keeps
the started/failed receipt and partial artifacts, and stops the matrix so the
failure can be reviewed rather than silently skipped.

Before a child launches, the driver writes `started.json`, binds the exact
binary and build-record hashes, records the exact argv and environment, and
creates a private replay directory for the file-store arm. The command is
wrapped by GNU `/usr/bin/time -v` and pinned with `/usr/bin/taskset -c 2`, as
in the frozen 0484 route protocol. After a successful report validation, a
file-store replay directory must be empty. The driver removes that empty
directory and writes `replay-cleanup.json` with its device, inode, and empty
precondition. Nonempty or failed scratch state is retained for inspection.

The source records in the two build files are content-addressed manifests.
The before manifest must be the retained 0485 after-build manifest. The after manifest must
be identical between normal and allocator builds and must differ from before
only under the protocol's OPC implementation or focused splice/DOCX test
prefixes. The comparison summary lists every changed source path and fails
closed on an unexpected path. It also binds each build gate receipt, copied
binary metadata, command, source manifest, and environment.

## Report validation and metrics

Every retained report is revalidated during capture and again by `analyze`.
The route arms call the frozen 0484 route validator; deterministic file,
short-read, and latency arms call its frozen `_check_axis_report` validator.
The old validator checks the complete report shell, source and authored
proofs, candidate oracle, input profile, replay counters, and allocator
counter conservation. The report argv and binary metadata are recomputed from
the frozen arm, so a plausible report cannot be substituted for the child
that produced it.

For every process, the summary retains the raw 30 `elapsed_ns` values and
their p50, p95, and p99 statistics. Allocator rows also retain the raw
operation heap values
`region_peak_live_bytes - live_bytes_before` and their p50, p95, and p99
statistics. This is the operation-scoped allocator metric; it is separate
from whole-process RSS. The allocator call, byte, live-byte, and peak-counter
statistics remain in each allocator process row as well as in the validated
report. GNU time's `Maximum resident set size` is parsed as exactly one
observation per child and is explicitly reported as `n=1`.

The summary keeps one row for every phase, arm, role, and repeat. It does not
average different workloads or hide a row behind an aggregate. Matched
before/after comparison rows are emitted per repeat for elapsed time, the
operation heap on allocator rows, and whole-child GNU time RSS. RSS remains a
single observation (`n=1`) per process, so its repeated p50/p95/p99 comparison
columns are labels for that one observation rather than independent samples.
Each change carries its signed percentage, lower-value direction, and a review
flag when the absolute change exceeds 5 percent. A flag is a review trigger,
not a causal speedup claim; two process repeats do not support a strong
confidence interval.

The source archive, authored proof, decoded candidate main XML, candidate
semantic oracle, and preservation flags must match for every arm across
phases, roles, and repeats. Candidate ZIP archive byte length and SHA-256 are
retained independently. If ZIP framing or compression changes while decoded
content remains identical, the summary records that archive identity change
instead of treating it as a content failure.

The JSON summary is `comparison-summary.json`; the complete row tables are
also rendered to `comparison-summary.md`. The results remain scoped to these
three workloads and selected arms. They do not establish a global memory
bound, a broad provider/input interaction result, or a causal performance
claim without review of the retained source diff, receipts, and flagged rows.

`observe_host.py` records point-in-time load, per-CPU counters and visible
compiler/profiler process affinities around each formal phase. These records
are context, not proof of a reserved host or continuous noise monitoring.
The normal and allocator children remain pinned to CPU 2 and run serially.

## This capture

The accepted before attempt is `formal2`; after uses `formal1`. Pilots are
omitted because the same canonical matrix and retained baseline were already
validated in 0485; formal samples and all controls are retained. The first
before command (`formal1`) failed before launching a benchmark because its
loader required the not-yet-built after executable. Its protocol and helpers
are retained under `development/protocol1`. The current helper loads only the
selected capture phase; analysis still requires both builds.

The new helper also corrects an inherited protocol-hash field that previously
hashed the last fixture after validating fixture identities. The 0485 sealed
inventory still binds its actual protocol bytes. No old artifact is changed.
The 0487 verifier independently checks the actual protocol file hash.
