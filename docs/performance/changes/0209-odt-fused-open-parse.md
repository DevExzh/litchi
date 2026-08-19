# Change 0209: ODT source-backed open fuses validation and content-style scans

Date: 2026-08-19

## Decision

**Banked.** The fused parse eliminates one of the two complete
`content.xml` tokenizations on the ODT source-backed open path. Frozen
cross-binary CPU-2 A/B/B/A measurement accepts the executed phase —
`odt_file_source_open` p50/mean/p95 14.63%-18.99% lower in both paired
directions with clean drifts — matching the profile attribution (the
removed scan was 6.58% inclusive and the surviving validation pass no
longer pays a second tokenizer setup; 26.23% + 6.58% collapsed to one
pass). The only adverse both-directions pattern (eager-open p50/mean,
sub-1.2%, on a byte-identical phase) did not reproduce in the single
permitted rerun, clearing the pre-floor block.

## Mechanism and invariants

Profiling attributed 26.23% inclusive of the `odt_file_source_open`
workload to `litchi_odf_common::core::family::validate_content_document_part`
and 6.58% to `litchi_odt::elements::style::StyleRegistry::from_xml` — two
complete `content.xml` tokenizations in
`SourceBackedDocument::from_package` (`document/source.rs`). The change
fuses them into one `NsReader` tokenization with two handlers
(`document/open_parse.rs`), the litchi-odt analog of banked change 0201:

- `ValidateHandler` replicates every check of
  `validate_content_document_part` for the text family — root, single
  body, family body, duplicates, depth, doctype/declaration/trailing
  content — with identical messages and positions, including the
  end-of-stream check. Validation errors abort the fused scan and return
  immediately, before styles.xml is fetched (the historical
  early-return); tokenization failures map to the validation-side
  message because validation historically tokenized the same bytes
  first.
- `StyleHandler` replicates `StyleRegistry::from_xml` byte-exactly:
  literal raw-qname `== b"style:style"` match (prefix-sensitive, no
  namespace resolution), raw undecoded attribute values, identical
  error strings, identical `Element::try_new` / `try_set_attribute` /
  `Style::from_element` / `try_add_style` call order; record-first-error.
- Error precedence preserved exactly: validation (incl. end-of-stream)
  → styles.xml parse → recorded content-style error → `try_extend`.

The standalone shells (`validate_content_document_part`,
`StyleRegistry::from_xml`, `StyleElements::parse_styles`) and the owned
`Document` open path stay byte-identical; the eager facade is untouched.

Invariants: identical observable outcomes — same registry contents or
same error message with the same precedence — on every input; no public
API change; changes confined to litchi-odt.

Verification: an equivalence oracle cross-checks fused vs sequential on
69 fixtures (23 `.odt`, 46 `.fodt`, recursive under `test-data/`) plus
15 synthetic malformed cases (wrong root, incomplete/duplicate body,
wrong family body, doctype, trailing text, content after root, late
declaration, invalid char reference, mismatched end tag, duplicate
style attribute, style-error-before-malformed-XML precedence, undecoded
entity in style name, custom-prefix-not-collected and
canonical-prefix-foreign-URI-collected literal-match pins), comparing
sorted style-name/family/parent/property projections. The full
litchi-odt suite passes (888 tests, +3 oracle tests); fmt, clippy
(`-D warnings`), rustdoc (`-D warnings`), and
`tools/check_crate_boundaries.py` pass. Facade suite: 269 passed, 8
failed — all 8 are pre-existing HEAD RTF fixture/parser mismatches
(trailing NUL in `test-data/rtf/testUnicode.rtf`), untouched by this
series.

## Matched release timing

Two frozen release binaries differ only in the fused ODT open parse;
both carry changes 0192-0196, 0198-0202, 0204, 0206, and 0207 (0208 was
withheld and reverted). Control SHA-256
`57270d24894a7047682146f4a6a68d428ecc51b3d1270e5b93d90d0fddcb284b` (the
banked 0207 binary; the tree was verified to rebuild bit-exact to it
after the 0208 revert), candidate SHA-256
`41b5f923638fa6fb0065318ce09edaf27588cf7aecd83c6c923c6d050f011efd`.
Fresh CPU-2-pinned processes ran `A1 control, B1 candidate, B2 candidate,
A2 control`, 30 warmups and 500 retained samples per leg, drift ceilings
5%/5%/10%/15% (p50/mean/p95/p99). The 0205 floor is litchi-ods-calibrated
and does NOT apply to litchi-odt (rule 4): banking uses the pre-floor
rule — accepts require lower in both paired directions with clean drifts;
any adverse both-directions pattern blocks unless cleared by the single
permitted rerun of that workload.

### odt_file_source_open (the executed phase)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 16.39% | 14.63% | -1.55% | 0.53% | ACCEPTED |
| mean | 16.81% | 14.67% | -1.10% | 1.45% | ACCEPTED |
| p95 | 18.99% | 14.85% | -2.66% | 2.32% | ACCEPTED |
| p99 | 21.58% | 11.89% | 16.80% | 31.24% | withheld (drift over ceiling) |

### odt_file_eager_open (no changed code — eager path byte-identical)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | -0.50% | -0.73% | 0.28% | 0.51% | withheld; adverse both dirs → single permitted rerun |
| mean | -0.08% | -1.17% | 0.37% | 1.47% | withheld; adverse both dirs → single permitted rerun |
| p95 | 1.75% | -2.95% | 1.16% | 6.00% | withheld (disagreeing directions) |
| p99 | 8.55% | -25.85% | -4.71% | 31.14% | withheld (disagreeing directions, drift over ceiling) |

### odt_file_source_open_full_text_lifecycle (open portion executes changed code)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 6.30% | -0.13% | -1.47% | 5.29% | withheld (disagreeing directions) |
| mean | 4.53% | 0.71% | -0.81% | 3.16% | ACCEPTED |
| p95 | -6.43% | 3.81% | 2.20% | -7.63% | withheld (disagreeing directions) |
| p99 | -11.71% | 16.62% | 8.76% | -18.82% | withheld (disagreeing directions, drift over ceiling) |

### odt_file_eager_open_full_text_lifecycle (no changed code)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 2.75% | 1.28% | 1.64% | 3.17% | ACCEPTED |
| mean | 2.02% | 1.62% | 1.65% | 2.07% | ACCEPTED |
| p95 | -1.86% | 1.34% | 1.49% | -1.70% | withheld (disagreeing directions) |
| p99 | -7.20% | 5.61% | 3.00% | -9.31% | withheld (disagreeing directions) |

### odt_file_eager_open — rule-2 rerun (clears the primary adverse reading)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 1.62% | -0.37% | -0.79% | 1.22% | withheld (disagreeing directions) — primary adverse pattern NOT reproduced |
| mean | 0.83% | 0.49% | 0.45% | 0.80% | ACCEPTED — primary adverse pattern NOT reproduced |
| p95 | -1.41% | 4.46% | 4.76% | -1.30% | withheld (disagreeing directions) |
| p99 | -6.22% | 17.69% | 23.26% | -4.49% | withheld (disagreeing directions, drift over ceiling) |

The primary-run p50/mean adverse both-directions reading (on a
byte-identical phase) did not reproduce in the single permitted rerun;
the block is cleared (same outcome as the 0201 eager-open rerun).

## Verdict

**Banked.** Claim scope, frozen cross-binary CPU-2 A/B/B/A (30 warmups,
500 samples), pre-floor acceptance (both directions lower, drifts within
5%/5%/10%/15%):

- `odt_file_source_open` p50/mean/p95: **14.63%-18.99% lower** (min-paired
  14.63%/14.67%/14.85%; p99 withheld on tail drift). The executed phase:
  one complete `content.xml` tokenization eliminated.
- `odt_file_source_open_full_text_lifecycle` mean: 0.71%-4.53% lower
  (p50/p95/p99 withheld, disagreeing directions).
- `odt_file_eager_open_full_text_lifecycle` p50/mean: 1.28%-2.75% lower
  (byte-identical phase; accepted but mechanism-absent, recorded for
  completeness).
- `odt_file_eager_open` mean: 0.49%-0.83% lower in the single permitted
  rerun (byte-identical phase; mechanism-absent).

No allocation/RSS, physical-I/O, cold-cache, producer, or broad-ODF claim
is made. Harness rebuild verified bit-exact to the measured candidate
(`41b5f923…`); this binary is the control for the next change. Raw
artifacts: `docs/performance/results/*-0209-*` and `*-0209r-*` (rerun).
