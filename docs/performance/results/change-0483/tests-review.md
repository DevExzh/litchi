# DOCX bounded tail-append test scope

This file records the focused integration-test scope for change 0483. The
tests live in `crates/litchi-docx/tests/source_backed_tail_append.rs` and
exercise `source_backed::Package::tail_append_plain_paragraph`, the explicit
`tail_append_noop` route, bounded `Limits`, and cancellation `Options`.

The fixture is a source-backed OPC package whose main document contains direct
plain paragraphs and a final opaque `w:sectPr`. The section-properties span is
checked byte-for-byte, including unusual namespace prefixes, attributes,
children, whitespace, Unicode QNames, and strict/transitional WordprocessingML
roots. A second physical member is retained so the test can compare its raw ZIP local
record, compressed payload, data descriptor, and central-directory record after
publication. The positive path checks paragraph placement, XML escaping,
reopening, and untouched-member preservation. The fixture profile uses a finite
4 KiB token window, depth 32, and 2 MiB semantic workspace ceiling; that is an
explicit envelope for the checked OPC audit formula rather than an unbounded
parser quota.
Main-story admission also covers 4,200 direct plain paragraphs carrying
`xml:space="preserve"`, with a narrow depth and 4 KiB token window, so the
aggregate attribute budget remains independent of per-event token scratch.

The refusal matrix covers non-final, duplicate, nested, and wrapped
`w:sectPr`; unsupported direct body children and paragraph grammar; invalid
XML QNames, duplicate or misplaced declarations, unsupported declaration
attributes, XML 1.1, and non-UTF-8 declarations;
protected, signed, external, or unsupported package topology; malformed XML;
overlong tokens; and both strict and transitional namespace forms. Refusals are
asserted before publication and leave the output sink empty.

Lifecycle coverage includes an exact no-op, an empty authored paragraph, a
changed source-backed append, an immediate exact inverse, source-version
staleness, append/source/token limit failures, cancellation before preparation
and publication, cancellation from inside managed and options-only short-write
publication callbacks, and a short sequential sink after output has started.
Managed `ExecutionContext` coverage observes memory while source callbacks are active,
checks release after plan drop, and exercises typed Memory and Work refusals
with no output. One-byte and one-time Interrupted source reads cover the
positional reader boundary. Store and Deflate main-document fixtures are used;
raw untouched members and their ZIP metadata are compared in both cases.
Settings protection and relationship closure are checked in strict and
transitional fixtures, including duplicate scalar attributes and elements,
malformed on/off tokens, and enabled/disabled order permutations. Full
mail-merge schema validation covers namespace spoofing,
duplicate structural owners and children, ordering, and malformed on/off
values, alongside fail-closed unsupported MCE input. Cache diagnostics also
verify that bounded preparation and publication do not retain the main XML
payload. Adversarial settings profiles keep the source/token guards below the
payload ceiling while checking inherited namespace-copy accounting and nested
MCE directive accounting. Small controls are admitted; larger fixtures refuse
with the typed `settings XML workspace` limit before publication. The MCE
fixture includes character-reference encoded colons and spaces so decoded
directive tokens and their resolved long namespace URI are charged.
An empty self-closing settings root is also admitted for both Word namespace
dialects, while a declaration placed inside that root is rejected before
publication. A valid MCE `AlternateContent` fallback is admitted while the
source settings member remains byte-identical; an unknown `MustUnderstand`
namespace is rejected through the typed document MCE error before output. An
exact source-size settings ceiling separately rejects the larger reinjected
MCE output before publication.

This is correctness and publication-contract coverage. It does not claim a
complete package-memory bound, a large authored-stream implementation, or
end-to-end DOCX performance evidence. Cargo, build, test, and formatting
commands were intentionally not run because the root agent serializes those
checks.
