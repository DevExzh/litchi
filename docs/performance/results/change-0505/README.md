# Change 0505 evidence

The shipped candidate is the **filter-only** six-line `candidate.patch`.
Primary timing evidence is **final/capture.json**, **final/summary.json** and
**final/before/** / **final/after/**. Both formal repeats improve all four
corpora without a >5% adverse measured flag. See the
[change record](../../changes/0505-odg-attribute-name-prefilter.md) for results,
validation, attribution and limitations.

Root `capture.json`, `summary.json`, `before/` and `after/` describe the earlier
**combined** filter plus borrowed-QName experiment in `combined.patch`.
Its two plain-large RSS regressions remain recorded; it is not shipped.
`rss-review/` contains six short supplementary children, not acceptance data.
Root before pilot files contain a separate 400-sample pilot. Do not pool them.

`profile-before` / `callgrind-before` / `heap-before` are fresh baseline
whole-child profiles; `filter-only` names identify the final candidate.
`after` names identify the combined experiment. `heap-plain-*` also pertains
to the combined experiment. Raw tool output is retained byte-for-byte.
Callgrind records instruction references only. Heaptrack includes child setup
and profiler overhead; unchanged rounded peaks do not imply exact equality.

Reproduction: build the unchanged probe using the command in environment.json
at the recorded clean base, freeze that binary, apply candidate.patch, rebuild
and freeze the candidate. Run `python3 capture.py --before /absolute/before
--after /absolute/after --output /absolute/new-results` (one shell line), then
`python3 analyze.py --root /absolute/new-results`. The analyzer imports the
retained 0502 bootstrap helper through this script's repository location.
CPU 2 and /usr/bin/time are required by the capture. Do not overwrite retained
results. Source, probe lock, binary and report hashes identify the captures;
the final rebuild is byte-identical to the timed filter-only executable.

`gates.json` lists final checks and explains retained initial fixture/format
failures. The final focused four-test rerun is included in the 96 unique tests,
not added to that total. `cleanup.json` records removal of batch-owned scratch;
`custody.json` hashes every other file in this evidence directory.
