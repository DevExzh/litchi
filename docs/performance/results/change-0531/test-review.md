# 0531 candidate test review

Historical static review of the initial test draft. The first execution later found four incorrect expectations; `test-fix-review.md` and the preserved failed receipts document the corrections and baseline reference run. This review is not a claim that the initial draft passed.

This review compares the frozen candidate in [`source-diff.json`](candidate/source-diff.json)
and [`source-manifest.json`](candidate/source-manifest.json) with the requirements in
[`source-review.md`](source-review.md). The production change is one hunk in
`litchi-ooxml-common/src/mce/codec.rs`; the only added test source is the nine-test
public integration file bound by [`test-source-binding.json`](test-source-binding.json).
No Rust build or test was run for this review, and no Rust source was edited.

## Coverage conclusion

The installed tests are adequate for this one-hunk detector substitution when read
with the source-level byte predicate proof. The old branch is

```text
xml.windows(NAMESPACE.len()).any(|w| w == NAMESPACE.as_bytes())
```

and the candidate asks whether

```text
memchr::memmem::find(xml, NAMESPACE.as_bytes()).is_some()
```

`NAMESPACE` is a fixed, non-empty ASCII string, so its string and byte lengths are
equal. For every byte slice, `memmem::find` reports a match exactly when one of the
same contiguous `windows` equals that non-empty needle, including short input,
boundaries, arbitrary bytes, and repeated or overlapping occurrences. The tests
therefore exercise the public consequences of the proven predicate rather than
coupling the regression suite to a search implementation.

## Requirement matrix

| source-review requirement | installed coverage | assessment |
| --- | --- | --- |
| Empty/short input, invalid UTF-8, NUL, and no-match borrowing | `absent_namespace_bytes_are_borrowed_without_xml_or_utf8_validation` | Covered, including exact input/output limits and source pointer identity. |
| Exact URI at beginning/end and near boundaries | `exact_namespace_occurrences_cross_search_boundary_offsets`, plus `exact_namespace_at_raw_boundaries_retains_parser_errors` | Covered at offsets 31, 32, 63, 64, 127, and 128, and at raw start/end with parser errors retained. |
| Repeated/overlapping matches and one-byte near matches | `repeated_namespace_occurrences_remain_owned_and_unmodified` and `every_namespace_byte_mutation_and_near_match_remain_borrowed` | Repeated matches and every URI byte mutation are covered. There is no separate self-overlap fixture, but only an existential boolean is consumed and the proof covers it; this is not a material gap for the fixed needle. |
| Exact URI in element/text, ordinary attribute, comment, and CDATA | `exact_namespace_occurrences_in_xml_contexts_enter_the_parser` | Covered for all listed lexical contexts. The namespace-declaration form is exercised by the real AlternateContent test. |
| Malformed input containing the URI and limit ordering/boundaries | `exact_namespace_at_raw_boundaries_retains_parser_errors` and `input_limit_precedes_output_limit_and_both_paths_report_exact_resources` | Covered for both parser error paths, input-before-output precedence, exact limits, and both match/no-match branches. |
| Semantic MCE behavior after detection | `real_alternate_content_processing_still_selects_the_fallback`, together with the existing common MCE suite | Fallback selection/reporting is covered directly. Existing tests retain choice/ignore/process/preserve behavior, `MustUnderstand`, malformed inputs, deep limits, DTD/entity handling, and POI/LibreOffice fixtures. |
| Ownership wrappers | `string_and_part_arc_wrappers_preserve_ownership_contract` | Covers borrowed `process_str`, owned transformed `process_str`, shared no-op `process_part_arc`, and distinct owned output after a lexical MCE hit. Existing DOCX source-backed refusal and PPTX source-offset tests remain applicable. |
| Affected OOXML callers | Frozen `quality-plan.json` test target includes common, XLSX, DOCX, PPTX, DrawingML, and XLSB packages | The new test need not duplicate the same common detector fixture in every caller. The candidate quality receipt must record this package command and the focused integration target. |

## Differential-oracle disposition

The source review requested comparison with a `windows(...).any(...)` reference, but
the installed tests intentionally use observable branch, output, ownership, report,
limit, and parser-error behavior. That omission is acceptable here: the source diff
is an exact replacement of the predicate, `memmem` has the same non-empty byte-search
semantics, and the fixtures probe the adversarial offsets and bytes that could expose
a mismatch. A literal reference implementation in the test would be useful for an
exhaustive randomized harness, but it is not required to establish correctness for
this frozen substitution and should not block the candidate.

The tests do not independently repeat every DOCX/PPTX/XLSX caller or the unchanged
`active_offsets`/streaming paths. Those paths delegate to unchanged code or are outside
the production hunk; the frozen package test command is the appropriate cross-caller
check. The present comment-only `process_part_arc` case verifies the lexical ownership
transition, while the real AlternateContent cases verify semantic transformation.

## Required receipts and commands

Root should retain the candidate only with a successful receipt for the focused test
and the frozen affected-package test command:

```sh
env CARGO_TARGET_DIR=/home/zhuhe/litchi-goal-0531-target CARGO_BUILD_JOBS=2 \
  CARGO_INCREMENTAL=0 RUSTDOCFLAGS=-D warnings \
  cargo test --locked -p litchi-ooxml-common --test mce_namespace_search \
  --all-features

env CARGO_TARGET_DIR=/home/zhuhe/litchi-goal-0531-target CARGO_BUILD_JOBS=2 \
  CARGO_INCREMENTAL=0 RUSTDOCFLAGS=-D warnings \
  cargo test --locked -p litchi-ooxml-common -p litchi-xlsx -p litchi-docx \
  -p litchi-pptx -p litchi-drawingml -p litchi-xlsb --all-features \
  -- --test-threads=2
```

Static coverage disposition: **pass, with no material test gap**. The focused and
affected-package receipts, existing semantic MCE tests, and the frozen performance
and quality gates still decide whether the optimization is retained.
