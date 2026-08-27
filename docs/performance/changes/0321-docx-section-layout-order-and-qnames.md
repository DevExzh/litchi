# Change 0321: DOCX source-backed section layout order and namespace fidelity

## Scope

This change narrows the lossless source-backed section-layout contract for
missing `w:pgSz`, `w:pgMar`, and `w:cols` children. A missing modeled child is
inserted before later standard Word section barriers such as `w:pgBorders`,
`w:docGrid`, and `w:sectPrChange`. Foreign direct children, comments, and
unselected ZIP members remain byte-preserved. If an unknown direct Word child
makes the safe insertion position indeterminate, the edit returns the typed
`Error::UnsafeEdit` result instead of guessing.

The coverage also exercises locally shadowed prefixes on page size, margins,
columns, and nested columns. Foreign declarations and attributes remain in the
source fragment, while newly generated modeled attributes use the exact Word
namespace of the section and reopen with the requested semantic values. The
existing strict and default-namespace cases continue to cover generated
qualified names for both WordprocessingML namespace families.

Exact semantic no-ops continue to publish the complete source artifact byte for
byte, including opaque section content and unrelated package members.

## Correctness boundary

This is a narrow ordering, namespace, and lossless-preservation claim for
source-backed DOCX section-layout edits. It does not claim support for unknown
layout-child editing, arbitrary direct-child reordering, or unsupported section
property mutation. Such cases remain subject to typed capability errors.

## Verification status

The focused integration target passed with 19 tests:

`cargo test -p litchi-docx --test source_backed_sections`

The full `--lib --tests` run passed 918 library tests and all preceding
targets, but the persistent unrelated
`source_backed_file::replacing_the_path_reports_source_changed_without_retargeting`
test failed in both the full run and an isolated rerun. A second full run with
only that test skipped passed all 918 library tests and every other integration
target. The skipped failure is not attributed to this change.

The broad Clippy run still reports only unrelated baseline failures in
`source_backed.rs` and unfulfilled lint expectations in package, settings, and
writer files. The changed target passed Clippy with `-D warnings` using only
`-A clippy::redundant_closure_call` and `-A unfulfilled-lint-expectations`.
Rustdoc with `-D warnings`, rustfmt, and `git diff --check` passed.

No streaming, allocation, RSS, OOM, throughput, latency, or broad performance
claim is made here. Those claims require separately measured corpus and
process-level evidence.
