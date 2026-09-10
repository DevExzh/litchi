# 0497 publication syscall profile

This is a bounded diagnostic outside the formal timing matrix. It runs the
retained after `docx_replayable_tail_append` executable once for each of the
two after-only publication routes:

* `counting` is the sequential, non-retaining sink arm (`counting_sink`).
* `atomic` is the new filesystem route (`atomic_path`). Its timed interval
  includes the sibling temporary file, temporary-file `fsync`, replacement,
  and parent-directory synchronization. Readback, hash/semantic checks,
  inverse checks, and removal of the destination are outside that interval.

Both arms use the same bounded file authored provider and `--replay-sync data`.
That makes the authored replay `fdatasync`/`fsync` visible in both traces and
keeps it separate from output synchronization. The fixed diagnostic case is
source count `8192`, authored count `256`, replay-window chunks, near-limit
text, one warmup, and one measured sample. The Rust report remains the oracle
for source identity, preflight, candidate XML/archive, semantic identity,
untouched members, physical order, and exact inverse restoration.

The outer coordinator must hold `/home/zhuhe/.cache/litchi-goal-0484/cpu.lock`
for the complete command. Root's 0497 gate does not own that lock itself, so
the launch is wrapped by `flock`; `profile.py` pins each child to CPU 2 but
does not take the lock. It gives every route a unique private directory under
`/home/zhuhe/.cache/litchi-goal-0497/publication-profiles/<attempt>/` and a
private `TMPDIR` under `/home/zhuhe/.cache/litchi-goal-0497`. A timeout kills
the complete process group. Empty private replay/TMPDIR trees are removed;
nonempty trees are retained and fail custody.

The exact launch contract, after the retained after build receipt exists, is:

```sh
python3 -B docs/performance/results/change-0497/profile.py capture strace1 \
  --build docs/performance/results/change-0497/build-after-<attempt>.json \
  --binary /home/zhuhe/.cache/litchi-goal-0497/retained/after/normal/docx_replayable_tail_append \
  --modes counting,atomic --samples 1 --warmups 1 \
  --source-count 8192 --authored-count 256 --chunk window --text near \
  --provider file-store --replay-sync data --replay-max-bytes 67108864
```

Run that command through the 0497 caller gate with the lock held around the
whole profiler invocation. The concrete lock wrapper is:

```sh
flock -x /home/zhuhe/.cache/litchi-goal-0484/cpu.lock -- \
  python3 -B docs/performance/results/change-0497/profile.py capture strace1 \
  --build docs/performance/results/change-0497/build-after-<attempt>.json \
  --binary /home/zhuhe/.cache/litchi-goal-0497/retained/after/normal/docx_replayable_tail_append \
  --modes counting,atomic --samples 1 --warmups 1 \
  --source-count 8192 --authored-count 256 --chunk window --text near \
  --provider file-store --replay-sync data --replay-max-bytes 67108864
```

The gate should record that exact wrapped command in its terminal receipt.
The helper's route flags are the benchmark CLI contract shared with the DOCX
provider harness: `--source-counts`, `--authored-counts`, `--chunks`,
`--text`, `--samples`, `--warmups`, `--authored-provider`, `--replay-dir`,
`--replay-max-bytes`, `--replay-sync`, `--publication`, and `--json`.
The profile labels `counting` and `atomic` are emitted to the child as the
canonical `counting_sink` and `atomic_path` values used by the provider
harness.
`--publication-path-flag` is available if the final harness exposes an
explicit destination flag; the default atomic implementation derives the
destination below `TMPDIR`.

After capture, revalidate without launching a child:

```sh
python3 -B docs/performance/results/change-0497/profile.py verify strace1
```

Each route retains `started.json`, `terminal.json`, `stdout.txt`,
`stderr.txt`, the actual `report.json`, the unmodified `strace.raw`, and
`replay-cleanup.json`. The raw trace uses `strace -f -yy -ttt -T -s 512` and
the selected filesystem/read/write calls `open*`, `read*`, `write*`,
`close`, `fsync`, `fdatasync`, `rename*`, `unlink*`, `stat*`, and `mkdir*`.
The parser preserves syscall call/error counts and every observed descriptor
path. It reports these scopes separately:

* `authored_file_store`: the replay file below the private replay directory;
  its sync rows are the authored provider's `--replay-sync data` calls.
* `output_atomic`: the generated `published.docx`, its `.litchi-*.tmp`
  sibling, and their parent directory below `TMPDIR`; this is where atomic
  write, file-sync, rename, and directory-sync activity is recorded.
* `report`: the benchmark's JSON report file.

The atomic arm fails validation if the raw trace does not contain exact
destination and sibling paths, output writes, a sync, and a rename. The
report's `destination_path` and `private_parent_path` are bound one-for-one to
the raw trace's exact destination and generated parent paths; the parent must
be the destination's parent. For every destination, the raw trace must show a
successful sibling write and file sync before replacement rename, followed by
a successful sync of that private parent directory. The file-store replay
sync is kept in its own scope and does not satisfy the output file or parent
sync checks. The child executes the atomic route during warmups but emits
publication records only for measured samples. Therefore the parser requires
exactly `warmups + samples` private destinations, validates the lifecycle for
all of them, labels the first warmup destinations as trace-only evidence, and
binds only the final measured destination sequence to the report records. It
does not synthesize report records for warmups. The counting arm fails if an
atomic destination or sibling appears. Missing unsupported events remain
absent with zero observed counts;
the helper never invents counts. The profile is descriptive syscall evidence
and does not authorize a speedup or assign a syscall to an individual timed
sample. The final report's inverse scope is checked literally as
`untimed_fixture_publication_inverse_exact; timed_atomic_publication_inverse_not_reexecuted`:
the destination is reopened/read back and the fixture publication inverse
oracle is checked after timed publication, while inverse replay is not claimed
as part of the timed atomic interval.
