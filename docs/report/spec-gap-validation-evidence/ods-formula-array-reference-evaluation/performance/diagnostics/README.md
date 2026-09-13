# Semantic preflight captures

These archives retain debug-build diagnostic captures from 2026-09-13. They
validate the corpus, runner protocol, and selected provider-read bounds. They
are not release performance measurements or final acceptance gates.

| Frozen diagnostic snapshot | Workload | Phases | Successful lanes |
| --- | --- | --- | --- |
| revision 7 | Value fixture | setup, parse, evaluate, parse/evaluate | 386 / 396 |
| revision 7 | Worksheet resolver, without instrumentation | setup, construct, evaluate, parse/evaluate | 88 / 88 |
| revision 7 | Worksheet resolver, instrumented | setup, construct, evaluate, parse/evaluate | 88 / 88 |
| revision 9 | Value fixture | evaluate | 99 / 99 |

Revision 7 failed the five aggregate sizes 4, 16, 256, 1,024 and 4,096 in both
evaluation phases. Smaller failures reported repeated complete scans; larger
ones exceeded the unchanged one-million-unit Work limit. Revision 9 passed
the same one-scan preflight oracles after aggregate reuse was extended to
references without absolute-coordinate markers. These are read-count and
budget observations, not latency improvement claims.

Every child used zero warmups and one measured iteration. The host was not
reserved for profiling, so the captured timing percentiles and RSS must not
support performance conclusions. The instrumented worksheet text case
reported 128 pointer checks and 128 matches; its uninstrumented counterpart
reports no provider counters. Both retained a 176-byte index reservation for
that fixture.

Each archive includes raw CSV, child stdout/stderr/status/time files,
commands, runner metadata, executable and harness hashes, and an ODS
source/test manifest. The manifest covers the ODS source snapshot, not every
dependency, compiler setting, or external input required to reproduce a
build. Revision labels are local diagnostic identifiers, not Git commits.
The executable and four harness inputs were unchanged during every capture.
Archives were verified against the adjacent per-file hash manifests before
the completed loose captures were removed from the working temporary folder.

The runner returned nonzero for revision 7's failed value capture and zero
for the successful captures. A separate overwrite check against an existing
22-case worksheet capture refused execution and left all 97 files unchanged.
The final stable-source gates and release scalar baseline comparison remain
required; neither these captures nor the 99-case preflight covers every
nested lazy-shape regression in the integration tests.
