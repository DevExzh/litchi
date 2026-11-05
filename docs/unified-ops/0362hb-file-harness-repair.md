# 0362hb — unified File qualification workload repair

- Date: 2026-10-22.
- Scope: the qualification probe only; no production format or facade code.
- Base: W8 Part A source at `3ce77d193a`, whose private `Commit` and
  `BatchCommit` results are consumed with `into_file()`.
- Status: source, correctness, lint and untimed route checks pass; native
  qualification remains pending the final W4/W4b-integrated source and a
  coordinator timing window.

The earlier probe did not perform the same read on both endpoints. Its leaf
side scanned whole-document text for DOC, RTF, XLS and PPT, counted XLSB sheet
names, and used broad text projections for other families, while the facade
read the selected marker. `LeafPackage::read_marker` now reads the same
logical target as the facade for all eleven families:

| Families | Native target paired with the facade marker |
| --- | --- |
| DOCX, ODT, DOC, RTF | first body paragraph |
| XLSX, XLSB, XLS, ODS | `Sheet1!L1` / first worksheet `L1` |
| PPTX, PPT, ODP | first shape on the first slide |

The generated DOCX fixture has 1,000 ordinary, mergeable paragraphs and a
real cached `NUMCHARS` field after them. The refresh workload computes the
non-whitespace character count and updates that field on both the native and
facade paths. The field's published value is checked after reopening the
output, so refresh is an actual cache edit rather than an empty update list.
The ODT native route uses its paragraph index rather than the facade's path
parser. XLSB import stays within the fixture's admitted cell-record range, and
RTF's batch count stays within its native 256-operation limit. The ODP fixture
now creates 100 slides with one shape per slide, so the batch route edits 100
distinct scene fragments while still calling `File::batch` and native
`splice_scene`. The superseded one-slide fixture put all 100 shapes in one
scene; its 100-item scene splice was rejected for overlapping that fragment.
That failure is retained in the external receipt named in the validation
receipt, rather than being hidden by reducing the workload to one item.

Before staging every route, the probe publishes both untouched source owners
and records their bytes and SHA-256 values. After staging, it publishes those
same owners again outside the timed region and fails if either byte sequence
changes. Validation then checks the edited semantic value, leaf/facade
semantic equality, changed output identity, exact no-op bytes, repeated output
bytes, and the source before/after digests. The JSON includes
`source_state_exact` and all four source-owner digest fields; semantic flags
are asserted by the validation runner rather than inferred from process exit
status.

The committed probe exposes 109 routes for the final untimed matrix:

- 33 open/get, save and repeated-save routes;
- 44 set and batch commit/publication routes;
- 22 merge commit/publication routes;
- 8 workbook import commit/publication routes; and
- 2 DOCX refresh commit/publication routes.

The targeted ODP routes passed these checks before the complete matrix. The
complete matrix asserted `semantic_equal`, `semantic_state_valid`,
`source_state_preserved`, `source_state_exact` and `exact_noop` for all 109
routes. Every `open_get` route also compared the native `read_marker` length
with the facade marker length. The command used for the probe build was:

```text
source (deleted-coord-artifact)
cargo build --offline --locked -p litchi-ops --all-features \
  --example file_w8_qualified_probe
cargo clippy --offline --locked -p litchi-ops --all-features \
  --example file_w8_qualified_probe -- -D warnings
python3 -B -m py_compile docs/unified-ops/results/0362h/*.py
```

The route checks used the reproducible
`docs/unified-ops/results/0362h/validate.py` runner, which invokes the built
example with `--mode validate --route FAMILY/ROUTE`, checks all five semantic
flags, checks equal well-formed before/after source-owner hashes, and writes
one JSONL artifact containing every route result. The reduced receipt is
committed at
`docs/unified-ops/results/0362h/0362hb-untimed-validation-receipt.json`;
it is 38,435 bytes with SHA-256
`e8d4f5cc21b538fd77ffa93b35aba8f5c549a7cd658e3a62bf578712d7107797`.
The external raw route artifact is
`(deleted-coord-artifact)/routes.jsonl`,
238,193 bytes with SHA-256
`877802466f1a0b4338b57e58a2cc9ad614427869ad4eef69e1bf7f6b78f71761`.
The reduced receipt also retains the superseded one-scene ODP failure as
`(deleted-coord-artifact)/superseded-odp-one-scene-batch100.receipt.json`,
971 bytes with SHA-256
`155b26508fa00ead8ef1a7d35a465a11dc8533107969905cffd72c72c161c658`.
Its exact development probe variant is 72,713 bytes with SHA-256
`07c6fcf2097df901f0a8adf79d2c53e50aadcd10ba8300cfd0587b5b10cb1693`;
the reduced failure receipt states that raw stderr was unavailable.
The exact source archive used by this rerun is
`(deleted-coord-artifact)/source.tar`,
245,760 bytes with SHA-256
`07cd6ff373721adcbda8c80190f312546ebe04c7a47f8e63f97a9d1a74f85ec5`. The receipt records the exact source archive, probe, binary and lock paths,
sizes and hashes; the protected binary is 244,723,872 bytes with SHA-256
`154268b7c7510ff383aa7d52fb01df7c19d409d0298de2f2b4a2dfcda730088e`.
No timing ratio or performance acceptance is claimed by this record. The
existing guarded ABBA runner remains responsible for paired release
measurements after the final integrated source and endpoint manifests are
available.
