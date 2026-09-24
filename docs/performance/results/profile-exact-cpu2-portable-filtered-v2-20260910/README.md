# Self-contained corrected CPU2 exact-region profile evidence

This evidence-only V2 candidate is based on the corrected d1bd archive
`d1bd39b9cf522f5aa97856712cc0eed67730764f78a4dacc5a7b83a5d0438623` and
preserves the immutable original capture archive
`dc26ba06db7264abb72ee8eb87de1969096afad4523d7607b3544fbfe5a841da`.
The `publication-bundle.tar.zst` SHA-256 is `58384afa7434d55cc95c4f99900fd00f8c4f384e969a2b27f6c2555d1fb980e2` and its extracted
inner manifest records 52 payload files totaling 6,833,355 bytes.

The extracted bundle includes four gzip-compressed raw `perf.json` inputs and
an input manifest that binds their compressed and uncompressed hashes to the
capture manifest. Its default verifier is therefore self-contained and runs
raw derivation without `/var/tmp` inputs. An external `--capture-root` remains
available as an optional cross-check against the frozen raw capture.

Extract and run the standalone verifier:

```sh
tar --zstd -xf publication-bundle.tar.zst
cd litchi-profile-exact-cpu2-portable-filtered-v2-20260910
PYTHONDONTWRITEBYTECODE=1 python3 provenance/verify_filtered_profile.py
```

Optional external cross-check:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 provenance/verify_filtered_profile.py \
  --capture-root /var/tmp/litchi-profile-exact-cpu2-20260910
```

`preparer-validation.json` records the independent raw recomputation,
source/depfile proof, deterministic repack, and diagnostic PMU boundary. `root-verification.json` records the successful standalone root check from a
fresh extraction; the temporary extraction was removed afterward.

The measured source is `cf98bee37455e8e6d0e73002ff2e94299ef546cb` with
only the profiling hook overlay. This diagnostic profile covers XLSX commit/save
and three OPC open/read paths. Operation stack denominators exclude handshake
samples; raw PMU totals retain the small enable/disable boundary. Instrumented
elapsed time supports no production latency or speedup claim. The separate
plain-cell optimization is absent from these binaries.
