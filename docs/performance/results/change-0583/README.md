# change-0583 evidence packet: a neighbour span bound that reads the local size

Change record:
[`docs/performance/0583-zip-local-size-span-bound.md`](../../0583-zip-local-size-span-bound.md).
Defect record: [`0582`](../../0582-zip-strict-scope-differential-fuzz.md), finding 2.
Changed behaviour this narrows: [`0580`](../../0580-zip-target-scoped-strict-layout.md).

Disposition: retained. `performance_claim: none` — **no timing, allocation,
read-count or byte-count improvement is measured or claimed.** This is a
correctness and safety fix. The one quantitative result about cost is a
*non-change*: change 0580's corpus census is byte-identical.

## Contents

| Path | What it is |
| --- | --- |
| `run.sh` | The whole gate, end to end: three `git archive` extractions of `93a610ded`, the corpus, the differential, the classification, change 0580's census probe, the three residual witnesses, the pre-fix test run, the mutation matrix, and the lint and test gates. |
| `prefix_archive.py` | Reconstructs change 0580's `archive.rs` by undoing change 0583's two hunks, and asserts the result's sha256 against the file change 0582 measured (`df9ed280…`). What `run.sh` builds the pre-fix tree from. |
| `classification-summary.json` | Every divergence count in the record: before (`93a610ded`) against the fixed build, over change 0582's 22,875-input corpus. Produced by change 0582's own `classify.py`, unchanged. |
| `delta.py`, `delta-summary.json` | Which verdicts change 0583 moved *relative to change 0580 alone*, and what the pre-change build said about each. This is the file that shows the direction property: 32,472 `err → ok → err` and **zero** verdicts made readable that change 0580 refused. |
| `crafted-verdicts.txt` | The 1,816 crafted-family verdict rows that are not identical across all three builds, each with its before / 0580 / 0580+0583 verdict. The finding-2 witnesses are `crafted/neighbour-local-csize-smuggles*.zip`. |
| `mutate.py`, `mutation-matrix.txt` | Five one-hunk mutants of the shipped fix, and the six new tests run against each. Every test is killed by exactly one mutant and every mutant is killed, so no test passes vacuously. |
| `residual_witness.py` | Builds the three archives that price what this fix does **not** close. Imports change 0582's `build_zip`, so the witnesses come from the same generator as the corpus. |
| `residual-verdicts.txt` | Those three archives' verdicts on the pre-change build and on the fixed build. All three are accepted by the fixed build and refused by the pre-change one: they are residuals of change 0580 that change 0583 does not reach. |
| `error-vocabulary.txt` | The distinct `InvalidInput` identities each of the three builds emits over the whole corpus, with the name/method/reason payloads collapsed, plus both diffs. `0580 → 0580+0583` is empty: this change adds and removes no error identity. |
| `prechange-tests.txt` | The six new tests run against the pre-fix tree. Exactly one fails; the other five are regression guards, priced by the mutation matrix instead. |
| `test-summary.txt` | Passed/failed/ignored per crate for the six crates on the reachable path, plus the fmt and clippy verdicts. |

## Result

```
divergence classification, before (93a610ded) -> after, over 22,875 inputs

                                              0580 alone   0580 + 0583
  panics, either side                                  0             0
  intra-build oracle failures                          0             0
  C  accept -> refuse                                  0             0
  D  accept -> accept, different bytes                 0             0
  A  refuse -> accept, pre-change was overlap     37,176        37,160
  B  refuse -> accept, pre-change was other    1,094,221     1,061,765
  E  refuse -> refuse, different identity         351,525       383,981

  32,472 member verdicts that change 0580 made readable are refused again.
  0 member verdicts that change 0580 refused become readable.
  0 verdicts change on any real archive or seed, on either side.
```

## Replay

The counts this record cites, from the retained files:

```sh
python3 -c "
import json
d = json.load(open('docs/performance/results/change-0583/classification-summary.json'))
print('divergences      ', d['divergences'])
print('panics           ', d['panics'])
print('oracle failures  ', d['oracle_failure_count'])
print('per family       ', d['per_family'])"
```

What change 0583 moved, and in which direction:

```sh
python3 -c "
import json
d = json.load(open('docs/performance/results/change-0583/delta-summary.json'))
print(d['transitions_base_0580_fixed'])
print('made readable that 0580 refused:', d['fix_made_something_readable_count'])"
```

The finding-2 witnesses, all three builds:

```sh
grep -E 'neighbour-local-(c|u)size' \
  docs/performance/results/change-0583/crafted-verdicts.txt | grep ' target'
```

That no new test passes vacuously:

```sh
grep -E '^=== |^  test result' docs/performance/results/change-0583/mutation-matrix.txt
```

That the pre-fix reconstruction is change 0580's code and not an approximation:

```sh
python3 docs/performance/results/change-0583/prefix_archive.py \
    crates/soapberry-zip/src/archive.rs /tmp/prefix-archive.rs
```

Corpus convergence, without rebuilding: change 0580's `census-after.txt` is the
fixed build's census byte for byte. `run.sh` regenerates and `diff`s it.

## What is not here

No timing, allocation, cold-cache or cross-platform capture, and no
re-derivation of change 0580's read counts beyond the census diff that shows
they did not move.

The full differential reports are 238,685,636 and 235,006,954 bytes and are
**not** retained; `run.sh` regenerates them.

The `parse_zip` fuzz target was **not** run. `cargo-fuzz` is not installed on
this host and no nightly toolchain is present, exactly as change 0582 recorded.
That gate remains outstanding.
