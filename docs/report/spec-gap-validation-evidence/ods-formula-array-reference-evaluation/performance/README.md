# Candidate array/reference capture

`run_candidate.py` is the serial release-capture wrapper for the candidate
scalar, value, and worksheet harnesses. It is a diagnostic custody and
comparison tool; it does not accept a performance result.

The wrapper requires an isolated copied workspace. Before building, it imports
the `hashes(root)` function from `gates/run.py` and requires that the workspace
and canonical checkout each have the same production source closure, byte for
byte. The closure is defined by the gate hash function and its current file
count is recorded in the receipt; the wrapper does not reject a legitimate
closure extension. It also requires the eight scalar/value harness inputs in
the workspace to match the canonical harness inputs. The receipt calls this
source identity a `dirty-copied-snapshot`; workspace and canonical `HEAD`
values are recorded as supplemental provenance only, because `HEAD` alone does
not identify dirty copied bytes.

Use a new output path for every run. The path must not be the retained
`baseline/` directory or either build directory:

```sh
python3 -B docs/report/spec-gap-validation-evidence/ods-formula-array-reference-evaluation/performance/run_candidate.py \
  --workspace /home/zhuhe/code/litchi-array-worktree \
  --output /home/zhuhe/code/litchi-array-candidate-release-01
```

The default target and temporary paths are
`/home/zhuhe/code/litchi-array-target` and
`/home/zhuhe/code/litchi-array-tmp`. Both must already exist. The wrapper runs
two `cargo build --locked --offline --release` commands serially, using the
existing target and never invoking `cargo clean`. The retained scalar baseline
is the ELF at
`/home/zhuhe/code/litchi-array-target/retained/scalar-baseline` by default;
`--baseline-binary` can name another retained ELF. The baseline raw CSV and
that ELF are required, hashed before and after the run, and are never used as
output paths. Toolchain versions, build flags, command logs, source and
harness hashes, and candidate executable hashes are retained in the output
root.

After a successful build the wrapper runs these fresh child output
directories, serially on CPU 6 through the existing runners:

| Output | Runner | Expected rows |
| --- | --- | ---: |
| `scalar/` | scalar, all phases and 94 cases | 282 |
| `value/` | value, all phases and 99 cases | 396 |
| `worksheet/` | uninstrumented worksheet adapter | 88 |
| `worksheet-instrumented/` | instrumented worksheet adapter | 88 |

The value runner may retain failed or refused diagnostic cases in its raw
output; the wrapper continues to the other lanes, records their status, and
returns failure when a lane is incomplete or unsuccessful. It checks executable
hashes after each lane and at the end. It also writes
`source-closure-before.json`, `source-closure-after.json`,
`input-hashes-before.json`, `input-hashes-after.json`, `source-identity.json`,
and `run.json`. `compare.json` is produced by the existing `compare.py` against
the committed `baseline/raw.csv`; `review-triggers.json` repeats every row
whose measured latency or RSS metric exceeds the 5% review threshold. Review
triggers are retained even when the compare process succeeds, and
`automatic_acceptance` remains false.

The two worksheet lanes measure different operations: instrumentation adds
provider counters and has its own overhead. A one-shot candidate capture is
not a final performance claim. The retained baseline was captured on a shared
host, so any regression or improvement requires repeated paired baseline and
candidate windows with the same corpus and build inputs. Host metadata is kept
lightweight and does not collect process tables or process command lines.

The gate suite itself is not run by this wrapper; only its source-closure hash
function is used for the preflight and final custody checks. Run the full gates
separately before invoking this capture.
