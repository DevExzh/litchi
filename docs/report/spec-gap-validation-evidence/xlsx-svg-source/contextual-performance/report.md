# XLSX SVG source contextual evidence

This report records exploratory source-bound evidence from two detached frozen
snapshots. It does not claim XLSX lifecycle performance, package memory,
native Office acceptance, or an optimization factor.

The before snapshot is `8b5838c591775990747b2cbce82fb2eea372b58b`; the after
snapshot is `fc44c4e6c945ab07ded7447f40670d898839eeb3`. The shared DrawingML
codec hash is unchanged. The after snapshot is the corrected host freeze with
raw-fragment-cap preflight and its boundary regression test.

Both bundles passed the sealed runner and verifier:

* 39 bounded lanes;
* three fresh processes per lane;
* two warmups and twenty measured samples per process;
* 2,340 samples and 117 retained empty stderr files per snapshot;
* allocator equations, semantic readbacks, output-cap refusals, SHA-256
  corpus identity, binary/source provenance, and source-manifest stability.

The normalized cross-snapshot manifest admits only the two intended host
source/test file deltas. Its digest is recorded in
`results/normalized-manifest.txt`. The raw JSON receipts, `/usr/bin/time -v`
RSS files, build provenance, binary hashes, corpus hashes, and verifier output
are under `results/before/` and `results/after/`; `results/comparison.md` is a
mechanical side-by-side index.

The source projection behavior is explicit in the receipts. For the synthetic
32-picture, 128-namespace lane, the before snapshot reports zero contextual
handles, zero `source()==None` values, and 4,262,358 retained raw-source bytes;
the after snapshot reports one shared context-storage node, 32 lazy values,
and 3,510 retained raw-source bytes. The exact 149,433-byte native-inspired
corpus reports 4,205,334 before and 1,078 after retained raw-source bytes.
These are scoped observations of the frozen scanner projection; they are not
whole-document memory claims. The after verifier requires those contextual
counts for every lane, while the before verifier requires the complete-source
projection.

Standalone export and scalar edit lanes check every expected embedded ID
(`rIdSvgN` or `rIdEdited`), reject linked references in readback, and preserve
the synthetic opaque QName fields. The 128-byte refusal lanes require every
selected picture to refuse without a partial-success receipt.
