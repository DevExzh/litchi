# XML bounded-reader audit review

Status: independent review complete and accepted. Root's
`xml-dev-tests-04` run passed the full `xml-minifier` suite with 16
independent stream tests, 14 legacy audit tests, 8 internal tests, and 5 asset
tests (1 ignored). The subsequent
`xml-fuzz-dev-build-01` and `xml-fuzz-dev-smoke-01` runs passed the ASan/
libFuzzer build and 10,000-iteration smoke over all 11 deterministic seeds.
The public memory helper is documented as a checked implementation envelope,
not an allocator or RSS claim, with its fixed BOM bookkeeping now accounted for.
The public `StreamError` documentation now links its `Input` and `Audit`
variants explicitly; the initial rustdoc link failure was documentation-only
and introduced no semantic XML change. Fresh final-gate aliases and the final
fuzz receipt are root-owned evidence records still being queued.

This review covers the existing `xml_minifier::audit` contract, quick-xml
0.41.0's buffered reader behavior, and the bounded reader needed by the DOCX
decoded splice. The relevant accepted contracts are ADR 0005 (finite resource
budgets, explicit scratch, source authority, and no plaintext spill) and ADR
0006 (validation must be deterministic, typed, and non-mutating).

## Existing slice contract that the reader must preserve

`verify` and `verify_authored` first account for the complete input byte length,
then require UTF-8, then parse with `quick_xml::Reader::from_str` and
`trim_text(false)`. The parser configuration is observable: end-name checking
is enabled, dangling ampersands are rejected, trailing closing-tag whitespace
is trimmed for parser matching, and comment-content checking is disabled. The
auditor separately checks raw event spans for compactness. Its counters include
the EOF event, raw event lengths for `TokenBytes`, raw text/reference lengths
for `TextBytes` (CDATA uses payload length), and all input bytes including an
XML declaration.

The current implementation validates raw event spans against the original
slice, so a leading UTF-8 BOM is not uniformly accepted or rejected. The
observed slice matrix is: `BOM+<a/>` is `Malformed@0`; a declaration is
`Malformed@0`; a leading comment followed by the root is `Malformed@8`; a
leading PI followed by the root is `Malformed@5`; text, whitespace, CDATA,
DOCTYPE, and a second BOM fail at offsets `0`, `0`, `0`, `0`, and `0`
respectively (the BOM-only input reaches the missing-root check at `3`). A
reader that strips the BOM and returns one generic error changes this raw-span
contract. The integration test keeps these cases explicit across one-, two-,
and three-byte prefix chunking.

The slice function also calls `from_utf8` before constructing the XML reader.
Consequently an invalid byte later in the input wins over an earlier malformed
XML construct. A one-pass non-replayable `BufRead` parser cannot reproduce that
global precedence without retaining/spooling the bytes or requiring a separate
validation pass over a replayable source. The API must either make a replayable
source/pass explicit, or document and test the streaming precedence. It must
not claim complete slice error parity while returning a malformed error before
an invalid byte that the slice path would report as `Encoding`.

`verify_authored` carries ambiguous-space state across adjacent text events but
flushes it at every markup event. Therefore a space-only text event before a
comment or processing instruction is rejected even if later text contains
non-whitespace. An explicit `xml:space="preserve"` scope suppresses both
formatting and authored-space rejection; `xml:space="default"` resets the
scope. Attribute values are decoded and normalized before deciding the scope,
while raw attribute layout is checked against exactly one ASCII separator.

## Critical buffering finding

`quick_xml::Reader<BufRead>::read_event_into` cannot by itself satisfy the
pre-growth token contract. Its `read_with` and `read_bang_element` paths append
every `fill_buf()` slice to the supplied `Vec<u8>` until a delimiter is found.
Returning tiny slices from a `BufRead` only changes the number of appends; it
does not cap the event buffer. A wrapper that reports at most `max_token_bytes`
per `fill_buf` therefore still permits a token much larger than the configured
limit to grow in the parser buffer.

quick-xml pushes the current start name into its private `opened_buffer` and
`opened_starts` before the audit loop can reject a depth overrun. The audit's
own `spaces` stack is entered only after its depth check, but the parser stack
has already observed the `max_depth + 1` start token. A zero depth limit still
permits that one parser event before refusal. The final bound accounts for a
depth-plus-one envelope with checked geometric capacity for each vector, and
the public reader documentation now states the same `(depth + 1)` rule.

The event buffer is not the only token-sized temporary. Attribute iteration is
borrowed, but `BytesStart::decoded_and_normalized_value` can allocate an owned
decoded value for an entity or normalized whitespace in `xml:space`. A bound
that includes the event buffer but omits one decoded-value window is incomplete
for inputs such as `xml:space="pre&#115;erve"`. The current helper reserves two
token-sized attribute temporaries, which covers the decoded and normalized
`Cow` values while normalization is active. The same review applies to any
temporary normalized namespace/name projection introduced by the new seam.
Documented exclusions for the caller's `BufRead`, fixed parser values, and
error strings are reasonable, but accepted and over-limit error paths must be
covered consistently.

For small profiles, ordinary `Vec` minimum capacities also matter: an empty
quick-xml open-name/index vector and the audit space stack can allocate a
minimum capacity greater than `2 * depth` when the configured depth is one or
zero. The current helper includes depth-plus-one and minimum floors. Its
attribute scratch term covers the duplicate-key `Vec<Range<usize>>` and a
fourfold `(u64, control-byte)` hash-table envelope after quick-xml's
32-attribute threshold. That factor is conservative for the pinned quick-xml
implementation, but it is implementation-specific: a future hash-table
layout or allocator would require rechecking bucket/load-factor and resize
peaks.

The current raw-span path has two token-sized guarded-source vectors
(`captured` and `exposed`) and a third token-sized `bom_raw` vector for BOM
inputs. The helper's six token capacities include those three vectors, the
parser event buffer, and the two attribute temporaries. The historical review
found that the old `+3` fixed-byte term named only one live BOM array. That is
resolved: the implementation now accounts for the 3-byte probe, 3-byte
history, and 6-byte history-combine scratch as a checked `+12` term, while
other fixed parser values, caller `BufRead` storage, allocator metadata, and
error strings remain explicitly outside the envelope.

The buffered quick-xml reader's BOM helper examines only the first
`fill_buf()` result. A one-, two-, or three-byte chunking source therefore has
different BOM behavior unless the wrapper coalesces a fixed three-byte prefix
or handles BOM policy before constructing quick-xml. Test all prefix lengths,
including a total-byte limit below three, and assert absolute offsets. The
workspace uses quick-xml's default feature set (no `encoding` feature), so a
feature change must be deliberate and parity-tested; enabling alternate
encoding detection would change declaration/BOM behavior.

The bounded implementation must either lexically frame each event into a
bounded scratch area before handing it to quick-xml, or use an equivalent
state machine that counts and rejects a token while discarding bytes after the
limit. It must not call `read_event_into` on an unbounded `Vec` and check the
length afterward. For an unterminated tag, comment, CDATA section, PI,
reference, or declaration, the scanner must also stop before scratch growth
exceeds the token limit and report a typed limit or malformed error according
to the documented precedence.

Counting bytes while discarding an over-limit token is safe only if the count
itself is checked. It must not attempt to retain the complete hostile token in
order to report its final length. If the streaming contract reports the first
observed excess as `limit + 1`, document that deliberate difference from the
slice implementation, whose `actual` is the complete raw event length. If
slice/reader error parity is required, the reader must continue boundedly to
the lexical boundary before reporting the final count, and must handle an
unterminated source without an unbounded loop or allocation.

## Parser and lexical edge cases

The focused integration tests should compare the reader against the slice
function over all of these inputs, using one-byte and several irregular chunk
sizes:

* declaration, PI, comment, CDATA, general references, numeric-looking
  references, and unknown named references (`&unknown;`); quick-xml exposes
  these as separate events and the auditor intentionally does not resolve the
  entity;
* text containing `<` only through markup boundaries, `&` split across input
  chunks, a lone `&`, `&` followed by `<`, and a reference that ends exactly at
  EOF;
* a long start/end tag, long quoted attribute value, long comment, long PI,
  long CDATA payload, and long text run, each one byte below, exactly at, and
  one byte above `TokenBytes`;
* `--` inside a comment. The current slice reader leaves comment validation
  disabled, so the streaming reader must not silently enable a stricter
  policy;
* malformed and truncated `<`, `</`, `<?`, `<!`, `<!--`, `<![CDATA[`, a
  declaration, a quote, a reference, and an internal-subset DOCTYPE;
* mismatched and unmatched end tags, multiple roots, non-whitespace outside the
  root, CDATA outside the root, and a BOM;
* UTF-8 multibyte code points split at every possible byte boundary, invalid
  leading bytes, invalid continuation bytes, truncated sequences at EOF, and
  invalid bytes inside names, attributes, comments, CDATA, PI, and text;
* `xml:space` values represented with entities or numeric references, inherited
  preserve/default scopes, duplicate `xml:space` attributes, and attributes
  whose quote or `>` occurs inside the value;
* exact raw-layout defects: repeated ASCII spaces, tabs/newlines/CR between
  attributes, spaces around `=`, whitespace before `>`, `/>`, or `</name >`,
  and a legal single ASCII separator.

The reader must retain absolute source offsets while consuming a `BufRead`.
Offsets in `Error::Malformed`, `Error::Encoding`, `Error::NotCompact`, and
`Error::Limit` are byte offsets in the original stream, not offsets in the
current read buffer or token scratch. The focused tests require exact offsets
for encoding, compactness, token, depth, and raw-layout errors. Generic parser
syntax errors are compared by typed category because quick-xml's buffered path
reports its event start while the existing slice path sometimes reports the
parser's post-read position; that diagnostic difference is documented rather
than silently presented as complete error-field parity.

## Accounting and precedence cases

Use profiles that isolate one resource at a time. Check exact acceptance and
one-under rejection for bytes, events (including EOF), attributes, depth,
token bytes, and aggregate text bytes. For each rejection assert the resource,
inclusive limit, actual value, and offset. Exercise checked accumulation near
`usize` limits through the public builder where possible; no counter may wrap
or saturate to an apparently valid report.

The byte-limit precedence is intentionally source-ordered for the reader: the
slice API can reject the complete input at offset zero before parsing, while a
reader reports the first observed overrun at the current absolute source
position. The integration test records that distinction while requiring the
reader position to remain stable across chunk sizes. Token limits are checked
before parser-buffer growth and report the first bounded lookahead byte.

The reader must not classify an event as compact before its full raw lexical
span has been checked. In particular, a token that is both over the token
budget and malformed needs a documented precedence; the implementation should
return the same precedence for every chunking pattern. Likewise, invalid UTF-8
must never be hidden by a lexical or compactness result merely because the
invalid byte appears after a delimiter in the same source chunk.

Check that an input read error is preserved as an I/O error in the public API
or mapped to the documented typed error, with its source error available when
the API promises it. A reader that returns `Interrupted` once must retry if the
contract follows quick-xml's behavior; a permanent error must terminate rather
than spin. Test an error in the middle of text, in a tag, and immediately after
the last complete event. No partial `Report` may be returned as success.

## Suggested public test shape

The API should expose a reader form accepting `&mut impl BufRead` (or an
equivalent generic source) and the same `Limits` plus authored-policy choice.
Tests can use a custom `BufRead` that returns one requested chunk at a time,
injects `Interrupted` once, or returns a permanent `io::Error`. Keep the test
source bytes borrowed and compare:

```text
slice_result(input, limits, policy)
        == reader_result(chunked(input, chunk_size), limits, policy)
```

For successful reports compare every field. For failures compare the stable
variant and all public fields; compare `Malformed.detail` only if the new API
promises quick-xml diagnostic parity, otherwise assert the stable offset and
category. Run the same corpus through `verify` and `verify_authored`.

A successful reader test should also prove that arbitrary chunking does not
change the authored-space state machine: split exactly at every byte in
`<p><b>a</b> <i>b</i></p>`, at each entity boundary, and around comments and
CDATA. A separate allocation/instrumentation test should demonstrate that a
token rejected above `TokenBytes` does not allocate proportional to its full
payload; a functional test alone cannot prove that property.

## Current integration review

The working tree now exposes `verify_reader` and
`verify_authored_reader`, returning the typed `StreamError::Input` or
`StreamError::Audit` variants. The independent integration suite is
`crates/xml-minifier/tests/stream_audit.rs`. It exercises successful report
parity for declarations, PI, comments, CDATA, entities, text accumulation,
and authored `xml:space` state; exact resource boundaries; pre-growth token
rejection for complete and truncated token kinds; raw end-tag and declaration
layout; malformed category parity; invalid UTF-8 offsets for each event kind;
one-time `Interrupted` retry and permanent source errors; and the finite-memory
helper. It also checks BOM token offsets across prefix splits and runs 1,000
deterministic ASCII mutations over valid seeds at chunk capacities 1 and 3.
Focused differential cases run through chunk patterns `[1]`, `[2]`, `[3]`,
`[7]`, and an irregular sequence.

The validation command used by root is:

```text
RUSTUP_TOOLCHAIN=1.98.1 cargo test -p xml-minifier --test stream_audit -- --nocapture
```

The root run passed all 16 independent stream tests. The BOM matrix now keeps
the existing slice outcomes for a bare BOM, declaration, comment, PI, text,
whitespace, CDATA, DOCTYPE, and a second BOM across one-, two-, and three-byte
prefix chunking. Raw end-tag whitespace (`</a >` and `</a\t>`) is checked from
the captured lexical span, while byte-limit diagnostics retain the documented
source-order difference: the slice reports offset zero after its upfront
whole-input check, whereas the reader reports the absolute first observed
overrun at byte 14 for every chunking. The interrupted-source cases also prove
that a one-time `Interrupted` result is retried and resumes to the same report;
permanent failures remain typed `StreamError::Input` errors.

## Fuzz review

The replacement target is
`crates/xml-minifier/fuzz/fuzz_targets/minify_xml.rs`. It drives both slice
policies and both bounded reader policies through an allocation-free
`BufRead`; chunk sizes from one through 97 bytes are selected by the input's
first 256 bytes. Inputs above 64 KiB are skipped, and the target uses a 4 KiB
token ceiling, depth 32, 1,000 attributes, 128 KiB events, and 64 KiB text.
Successful reports are compared exactly. Valid-UTF-8 failures retain category
and resource/limit parity, with documented exceptions for source-order UTF-8
precedence and truncated-token malformed-versus-lookahead-limit precedence.

The deterministic corpus is recorded under
`docs/performance/results/change-0482/fuzz/seeds/` with nine XML seeds and two
binary invalid-UTF-8 seeds. `xml-fuzz-dev-build-01` built the target and
`xml-fuzz-dev-smoke-01` completed 10,000 seeded iterations with exit code zero;
the retained smoke receipt is `fuzz/smoke.json`. The producer minifier remains
a compile-time macro API, so the target includes a fixed macro smoke invocation
while runtime fuzz input exercises the public audit APIs.
