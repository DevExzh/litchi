# 0361fda native IWA qualification harness

This directory contains the replacement for the interrupted 0361fd native
qualification attempts. The constructor probe retains all 15 routes from the
0361fb retained-header protocol. The owner probe retains the four measured
routes: Numbers header settings, Keynote media replacement, Pages footnote
insertion, and Pages section-name replacement.

The probes do all semantic and byte checks before starting their timed loops:

- constructor routes prove canonical and retained-header round trips and
  report the exact typed tight-header rejection;
- owner routes prove changed output, replay, inverse, and exact no-op output;
- every timed result carries source/output SHA-256 identities and a semantic
  checksum.

`iwa_native_qualification.py` snapshots both source commits, freezes the
shared source and lockfile identities, records toolchain and binary hashes,
requires native and allocation validation for both endpoints, and then runs
calibrated five-block ABBA. A build/profiler/benchmark process seen before,
during, or after a leg aborts the attempt; no contaminated block is retained
as qualified timing.

All generated projects, binaries, receipts, raw output, and timing evidence
must live outside this checkout. The owner fixture directory must contain
these files, copied from the tracked `test-data/iwork` fixtures (and the
small synthetic Pages footnote fixture used by the existing owner workload):

```text
basic.numbers
media-replacement-native.key
tone.wav
footnote-fixture.pages
basic.pages
```

It must also contain `manifest.json` with schema
`litchi.iwa.owner-fixtures.v1` and an absolute, SHA-256/byte-count artifact
row for each of the five files. The runner checks those identities before
validation and again before timing.

Run the Python unit tests without Cargo:

```sh
python3 -m unittest discover -s docs/unified-ops/results/0361fda -p 'test_*.py'
```

The host coordinator must supply a quiet window before the `run` phase. The
build phase is deliberately separate so source and binary provenance can be
reviewed before any native timing is attempted.
