# Change 0506 evidence

The shipped shape-only candidate is `candidate.patch`. See the
[change record](../../changes/0506-odg-shape-attribute-span-batch.md) for results,
mechanism, correctness scope and the accepted **plain-large RSS regression**.

`capture.json`, `summary.json`, `before/` and `after/` contain the primary
A1/B1/B2/A2 comparison: 16 children, 25 warmups and 200 samples per child.
The `before/*-pilot.json` files are a separate 400-sample pilot and are not
pooled with formal reports. Two >5% adverse RSS flags are retained; they are
not removed by profiling or averaged away.

To reproduce, build the unchanged 0502 probe at the clean base revision using
the command in `environment.json`; freeze the executable, apply the recorded
candidate patch and rebuild/freeze the candidate. Run `capture.py --before
/absolute/before --after /absolute/after --output /absolute/fresh-results`
with Python (one shell line), then `analyze.py --root /absolute/fresh-results`.
The analyzer imports the retained 0502 bootstrap implementation from this
repository location. CPU 2 and /usr/bin/time are required. Do not overwrite
retained reports. The source manifest supplies exact identities, including the
unchanged probe lockfile and final verification executable.

Raw Callgrind and heaptrack reports and traces use before/after names;
`heap-plain-*` profiles plain-large, while the other profiles use metadata-large.
Commands, scope and tool versions are recorded in profiling/environment JSON.
Profiles contain setup, hashing and a preflight open in addition to the one
sample, with zero warmups; counts are whole-child. Hardware counters were
unavailable. Profiler RSS is separate from the /usr/bin/time formal results.

`gates.py` reproduces the scoped gate commands with explicit output and target
paths. `gates.json` lists observed outcomes. Focused reruns are not additional
unique tests. `adr-review.md` and `adr-manifest.json` record the design review
and exact current accepted-ADR inputs. `cleanup.json` records removal of only
batch-owned scratch. `custody.json` hashes all other evidence files.
