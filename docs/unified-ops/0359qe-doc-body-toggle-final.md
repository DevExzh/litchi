# 0359qe — DOC body-toggle final integration qualification

- **Status:** deterministic profile complete; native ABBA pending; batch raw-byte
  fidelity is explicitly recorded for coordinator review.
- **Date:** 2026-10-23.
- **Integration:** `3cad59be707c7431bc1a8c115e1ff534a048dc57`.
- **Scope:** the existing private `litchi-doc` body transaction and the existing
  W5 probe. No public API or alternate editing family was added.

## Source assemblies

The candidate is the `3cad` tree as checked out. The baseline is an immutable
source overlay on the same tree: only the body-publication diffs from
`ae094d7d055c`, `6456aa3223`, and `160a00559b` were reversed in
`body_text.rs` and `tracked_revision/package.rs`. The current `PatchConflict`
versus composition `Conflict` taxonomy, current probe, lockfile, and every
other `3cad` file remain unchanged. The baseline is not described as pristine
`81323113be`.

The candidate implementation hashes are body
`614ce42c2dfe8ad985dc37d2c949602916bc155bd51149f2e843f24fdf378224` and
revision package
`0fd67fae678e870f6e7216f0ec5d64ea62a0da7db6df7f87297a965fd9a2da86`.
The baseline overlay hashes are body
`3e09f67abe5c40248959fe35ca90ea9c8ceb52609394267970b1307be09465d5` and
revision package
`613beaf162033c6153906ef2b1299d21e0ece58b4283a43400924c2503dbe40a`.
The exact two-file overlay diff has SHA-256
`10765febb64c4e099f5c3ef560366124dc1d28d540d138768ae0f7cf782da2ab`.

Both legs use probe SHA
`9e1cd649268ad5eefe4314b76493f61cd047bf6ec4114c7e7872b2182b1af689` and
the same release command:

```text
cargo build --offline --locked --release -p litchi-ops \
  --features doc,xls,rtf --example legacy_w5_source_probe
```

The build used toolchain 1.95.0, one Cargo job, incremental compilation off,
release debuginfo off, and the 20 GB memory-capped shim. Binaries and all
profiles remain outside the repository.

## Correctness and fidelity

The ordinary route (`doc/leaf_write`) produced identical raw output SHA
`6bb97058e47f747c29e568395eea01db1226aa5ab8962087e1d78a1b9b82ff11` on both
legs. The mixed-formatting test, refusal atomicity, inverse, and exact no-op
checks passed in the candidate's full DOC library run (1,210 passed, 2
ignored); the baseline overlay also compiled and its retained-render test
passed. The all-feature `litchi-ops` `doc_unified_ops` integration suite also
passed all 13 tests.

The 1k route preserves rendered output but not raw container bytes. The
candidate emits a compact 182,784-byte OLE file (`bd016a…`); the baseline emits
7,715,328 bytes (`4a5bcc…`) because it retains superseded CHPX/1Table history
from publishing every toggle. LibreOffice produced 18 pages for each file;
the 72-dpi rendered page hash list and extracted text are identical. This is
an explicit fidelity review item, not an unreported equal-byte claim.

## Structural CHPX pointer audit

The raw-byte difference is now covered by a reproducible structural audit,
not only by LibreOffice text and render hashes. The two OLE artifacts remain
outside the repository:

| role | artifact | bytes | SHA-256 |
|---|---|---:|---|
| candidate | `final-results/candidate/batch-output.bin` | 182,784 | `bd016a87f5d0015d6f8b910862eeeef9c0ce490033b4951c32ee5f403b7eebdf` |
| baseline | `final-results/baseline-batch-output.bin` | 7,715,328 | `4a5bcc2888ab819488fbb5a1ced1b088010440bd8293f10556e0903993e448d3` |

Regenerate the audit from the repository checkout and the external artifacts
with no Cargo build or timing run:

```text
python3 tools/audit_doc_chpx_history.py \
  --candidate (deleted-coord-artifact) \
  --baseline (deleted-coord-artifact)
```

The script uses only the Python standard library. It checks the OLE stream
inventory, compares every FIB pair, follows the FIB-selected `PlcBteChpx`,
and compares each referenced 512-byte `ChpxFkp` payload. The only FIB pair
difference is index 12 (`fcPlcfBteChpx`): candidate `(16384, 228)` versus
baseline `(135292, 228)`. Both live BTEs contain the same 28 `aFC` ranges and
the same 28 FKP payloads; only their physical page numbers differ. The exact
audit rows are:

| # | `aFC` range | candidate page | baseline page | referenced FKP SHA-256 |
|---:|---:|---:|---:|---|
| 0 | 1536..5640 | 280 | 14644 | `6ef9b380ff0af24eb1855657ef8507ca2d0f8c55c5226af137eeca9f7042591a` |
| 1 | 5640..9744 | 281 | 14645 | `b4f8f327130ca211c6d18bf9192bc8c7bbfdee20411d60a5c2b8101b597139c1` |
| 2 | 9744..13848 | 282 | 14646 | `3612c9a0059ae65ce1b0f3f2b605ca375f61cf33c455c6a7992aa6c8e2447fcc` |
| 3 | 13848..17952 | 283 | 14647 | `ff26025ec9bfe6b06602d953309b05a7270c8bb6ad19615c594a1e0bb9974d11` |
| 4 | 17952..22056 | 284 | 14648 | `6bb563dea3ee53b575786981b771bb826c95150ac83dd244e0ecc61e323e9288` |
| 5 | 22056..26160 | 285 | 14649 | `47050d7d0f43708415aef08b687c51550fdd664d16e626813e844fc0490f123d` |
| 6 | 26160..30264 | 286 | 14650 | `2789df9829efe2d2262dddc11cde9504a15fdbcfc5a2a09ec3dc12a8ea7cf452` |
| 7 | 30264..34368 | 287 | 14651 | `35c834b04437682d3083360dd73a119b9c284995dd76a2af69e2896313c5e958` |
| 8 | 34368..38472 | 288 | 14652 | `db003621369b20e56b281684c11f5d9538caa7fc4d1ecb394e368cb31e6804f3` |
| 9 | 38472..42576 | 289 | 14653 | `338a09625684a0f6d684a0856e84c858e7137b7e5a0c367f1eb1ff76b3eef2b5` |
| 10 | 42576..46680 | 290 | 14654 | `90961f4f58be52a6686f06175284e7d6959c44332a2b2c76adcad106a3541b82` |
| 11 | 46680..50784 | 291 | 14655 | `98427d9eef1cc164b535ae93c0ee7e5a087c21489ce500abaa69c5663522820b` |
| 12 | 50784..54888 | 292 | 14656 | `109302854889e47a2431945e31872b365cf176b44a53b5c4e56d128cc5cb4e53` |
| 13 | 54888..58992 | 293 | 14657 | `da1764c57c31cb28b7cf034c034aa713726dd13b4e6818fe5c2e2cd3199cbe32` |
| 14 | 58992..63096 | 294 | 14658 | `c8fe0139318279cb68379288098fb61211ddd9902e03e0e395b01712916be3e2` |
| 15 | 63096..67200 | 295 | 14659 | `72c6a8f1e4184eaf73b8f647f28c11d8057a0a726828243cc6c2f657d5078e59` |
| 16 | 67200..71304 | 296 | 14660 | `6d79467e697dbdf5ce3ac739c14fac5c3736e5f42b915463a36d9a16e8543621` |
| 17 | 71304..75408 | 297 | 14661 | `e5c926bb70e5bcd2d1c86d09c1962332df363f61d117e5997e30d3e680516264` |
| 18 | 75408..79512 | 298 | 14662 | `98d257f87ab41ed5160d04958751e66428da992dcb28f997a238b11309396319` |
| 19 | 79512..83616 | 299 | 14663 | `e2d5a2c41211f09fde80dadfaed8c7718a28bdb30df8230d726b1b2775d52e2f` |
| 20 | 83616..87720 | 300 | 14664 | `0c9733eb4b527d2885d47b6afa66bb6e2e6e202a2ac76c1d64207a893aa8a5fe` |
| 21 | 87720..91824 | 301 | 14665 | `70d6d42c1744286fddd909355873505f7d5c81071e37a080ba3cc6e23f636af6` |
| 22 | 91824..95928 | 302 | 14666 | `5f2cd7d98617b09ff0f934c0c8e583cde257d538c852d0daa6df1c009e87519c` |
| 23 | 95928..100032 | 303 | 14667 | `25ec5901597f3494811019959d605bd3782cd7fbc1bffa7e1b84bb9811a4211f` |
| 24 | 100032..104136 | 304 | 14668 | `48bc55bcd2dcd6d61726ae6477cd700e0dca576e35ea1ca9bc84f2b2eb29a39e` |
| 25 | 104136..108240 | 305 | 14669 | `d527b1855670edf8a22012b09d6bdd0ee5afaf962268f0de56733eb0d7eb9e5f` |
| 26 | 108240..112344 | 306 | 14670 | `069448dbd495043090b0890f35878e4c7c7fa64f879ff7e705f6d400559f6e12` |
| 27 | 112344..115536 | 307 | 14671 | `61f95fd666dc4129a22421cf5a746245f6e54ff82a0df80b1d5d9c9747372db6` |

The non-DOC streams are unchanged: `\u0001CompObj` is 106 bytes with SHA-256
`6f1ec717cd671c4e65321a1129f6fe3036e0effc4f87b3c72160eeed71483760`,
`\u0001Ole` is 20 bytes with SHA-256
`c36c8a4b7dee703b9ce6e288032033b718feef01ca283cfaa4332a8334b2adf3`, and
`Data` is 4,096 bytes with SHA-256
`ad7facb2586fc6e966c004d7d1d16b024f5805ff7cb47c7a85dabd8b48892ca7`.

In the integration source, `RevisionEditor::rewrite_chpx` in
`crates/litchi-doc/src/tracked_revision/package.rs` (lines 3366–3412) appends
the current FKP pages and replaces the FIB's CHPX BTE pair; `parse_bte` and
`parse_chpx` in `crates/litchi-doc/src/tracked_revision/codec.rs` (lines
648–666 and 466–496) read only that selected pair and its page numbers. This
is also the MS-DOC rule: `PlcBteChpx` maps each live WordDocument FC range to
its referenced `ChpxFkp` page. Therefore the
baseline-only blocks between its old and final BTEs are superseded allocation,
not a formatting or tracked-revision history reachable by the document. The
candidate's compact output retains the same live ranges and FKP payloads.

This structural result does not waive the qualification boundary: quiet-host
native ABBA remains pending, and all performance measurements above remain
diagnostic with no performance acceptance claim.

## Performance

The selected ordinary route has four calls (expected output, two warmups, one
measured call). Callgrind inclusive Ir is 257,848,943 candidate versus
256,622,212 baseline (1.004785×). `Edit::commit` is 79,843,981 versus
73,325,557 (1.088818×), while the setter falls from 32,195,265 to 22,463,186
(0.697731×). The ordinary Massif peak is 7,910,440 versus 8,228,648 bytes
(0.961329×).

The full 1k commit-plus-write route is 241,104 µs candidate versus 1,319,299
µs baseline in three-round wall diagnostics. Callgrind inclusive Ir is
24,673,833,943 versus 60,148,185,975 (0.410228×); the setter is 0.408337×,
and `write_to` is 0.036455×. Massif peak is 8,769,392 versus 107,252,816
bytes (0.081765×). These measurements account for the historical 1.1–1.7 ms
per-1k residual and include the final `write_to`.

The host load was approximately 37 during native attempts, so the wall
numbers are diagnostic only. A quiet-host native ABBA block remains required;
no native regression or acceptance claim is made here.

Reduced measurements and exact source/binary/profile hashes are in
[`results/0359qe/qualification.json`](results/0359qe/qualification.json).
No large evidence or executable is committed.
