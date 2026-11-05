# DOCX durable Ink qualification

This directory contains the committed probe and runners for the native
durable-Ink comparison. It is a DOCX workload; the legacy DOC body toggle
(`owneddoc_toggle`) is deliberately not an accepted binary or process name.

The build phase takes the final refs at execution time and records their
resolved commits, source archives, dependency lock, compiler, probe digest,
release settings, validation outputs, and protected binary digests. Keep its
output outside the checkout:

~~~sh
python3 docs/unified-ops/results/0362l/ink-qualify.py build "$OUT" \
  --base "$BASE_REF" --candidate "$CANDIDATE_REF" --lock "$PWD/Cargo.lock"
~~~

`build` runs the untimed correctness gate before any native timing is
possible. Both endpoints must produce the same source, forward, inverse and
no-op package bytes and the same forward, inverse and no-op durable patch JSON.
The probe also reopens
independent content-matching base and target packages, applies forward and
inverse patches to those packages, reads them through a fixed 4-KiB reader,
and publishes forward/inverse/no-op packages to a fixed-state digest sink.
The stream receipt records byte counts and digests without retaining output in
a `Vec`.

Only after a coordinator grants a quiet timing window, run the matched native
ABBA. The build manifest is mandatory: the runner rejects binaries whose path
or digest is not the protected pair produced by that manifest, and checks
those digests before every leg.

~~~sh
python3 docs/unified-ops/results/0362l/ink-abba.py \
  "$OUT/build-base/ink-native" "$OUT/build-candidate/ink-native" \
  "$NATIVE_OUT" --build-manifest "$OUT/build.json" --cpu "$CPU"
~~~

Each route is calibrated independently to at least 500 ms, warmed twice, and
measured in five ABBA blocks with host-overlap snapshots. Raw receipts,
source trees, archives, binaries and timing files remain outside Git.
`ink-collect.py` verifies the complete provenance chain and writes only a
reduced evidence summary.

The allocation phase is optional and is not a native timing claim:

~~~sh
python3 docs/unified-ops/results/0362l/ink-qualify.py allocation "$OUT"
python3 docs/unified-ops/results/0362l/ink-collect.py "$OUT" --out "$SUMMARY"
~~~

Final baseline/candidate refs, a fresh external output directory, a
coordinator-approved quiet CPU, and completed release builds are prerequisites
for native qualification. No binary or timing result is committed by this
source-only preparation change.

The source-only failure-case guards can be run without Cargo:

~~~sh
python3 -B -m unittest \
  docs/unified-ops/results/0362l/test_ink_qualification.py
~~~
