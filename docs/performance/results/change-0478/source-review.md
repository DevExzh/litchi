# 0478 generated-name source review

Review scope: the symbolic generated-name capability described by
[`api-plan.md`](api-plan.md), with particular attention to the proof boundary
between ZIP, OPC, and PPTX. At the time of this review the new lower-level
plan module was still being assembled, so this records the invariants that the
implementation and its tests must demonstrate. This is an independent review
record, not a completion or performance claim.

## Review gate

The generated route can remove retained name indexes only if the plan is a
checked capability. A finite descriptor representation and a scalar cursor
are sufficient, but accepting a descriptor whose expansion has not been
proved equivalent to the production normalizers would make the memory result
unsound. The implementation should therefore fail closed with a typed
unprovable-plan error whenever a symbolic comparison cannot be decided.

The proof must run before the output sink is touched. In particular, plan
construction must not defer PackURI validation, path normalization, or range
overflow checks until the first entry. A provider may be initialized after
the plan has been validated; a preflight failure must leave the output sink
unchanged.

## Required symbolic cases

The family language is safe only with explicit boundary rules. A numeric slot
in the final component must have a non-digit boundary on both sides: a prefix
ending in a digit or a suffix beginning in a digit makes the decimal split
ambiguous and must be rejected. A slash in the suffix would move the slot out
of the final component and must also be rejected. Prefixes may contain fixed
parent components, but every component and the whole path must be canonical
under the production ZIP normalizer. `slide01` is not the canonical expansion
of index `1`; a caller must not be able to smuggle leading-zero spellings into
the plan.

The comparison matrix must include all of these without enumerating an index
range:

* literal versus literal, including exact and ASCII-folded equality;
* literal versus family, including a literal equal to one family member;
* a shorter literal that is an ancestor of a family, and a family that is an
  ancestor of a longer literal;
* family versus family with equal depth, overlapping numeric intervals, equal
  and unequal fixed suffixes, and differing fixed prefixes;
* family versus family with unequal depth where a generated final filename of
  the shorter family equals a fixed parent component of the deeper family;
* the same unequal-depth case with more than one intervening fixed component;
* component-prefix lookalikes such as `a/b` versus `a/bc`, which are not
  ancestors;
* prefix/suffix lookalikes such as `slide{n}.xml` versus
  `slide{n}.xml.rels`, which are distinct at equal depth, and a fixed parent
  such as `slide1` versus the family `slide{n}/child`;
* ASCII-case variants at every one of those boundaries; and
* zero, one, digit-length transitions, `u64::MAX`, and ranges whose final
  value would overflow.

For ancestor comparison, compare complete path components. A shorter family
`/a/{n}` conflicts with `/a/1/x/{m}` whenever `1` lies in the first family's
checked interval, even though the variable component is the filename of the
shorter family and a fixed parent component of the deeper family. Checking
only equal-depth families or only the immediate parent misses this conflict.
The converse family/literal direction needs the same component-prefix test.

## Normalization and layer boundaries

The lower ZIP plan should use the exact production member-name normalization,
then require the caller spelling to already be canonical. It must account for
backslashes, repeated or leading slashes, dot segments, parent segments,
drive/colon prefixes, and trailing directory separators; otherwise two
different inputs can pass the symbolic proof and collide when the writer
normalizes them. Since the initial proof language is ASCII, non-ASCII input
should be refused rather than relying on an unexamined Unicode equivalence.
The ZIP plan must not import OPC grammar constants.

The current validator manually encodes those canonicality checks instead of
calling or sharing `ZipFilePath` normalization. Under the current production
implementation the accepted forms are equivalent for separators, dot/parent
segments, colon prefixes, and trailing separators, but ASCII controls remain
accepted because the ZIP normalizer leaves them unchanged. That is coherent
with ordinary ZIP behavior; if the contract means printable ASCII, the plan
must explicitly reject controls and add a regression test. In either case,
add a differential test against the production normalizer so a future path
normalization change cannot silently invalidate the symbolic proof.

The OPC wrapper must separately validate every literal and both endpoints of
every indexed family with `PackURI::new`. Endpoint validation is enough for
the restricted one-slot language only after the implementation proves that
all changing bytes are decimal digits and all syntax-sensitive bytes are
fixed. This includes leading slash, no empty segments, no backslashes, no
dot segments or trailing dots, no spaces/control characters, and valid
percent-encoding. The wrapper must preserve the OPC leading slash when it
translates to a ZIP member name, and must apply ASCII folding to the same
canonical bytes that `PartNameSet` uses. Do not use Unicode lowercase or
percent-decoding as a substitute for OPC's ASCII equivalence.

The proof should compare both exact ZIP names and ASCII-folded OPC names.
Examples that must be rejected include `/A/x` with `/a/x`, a family with a
folded-equivalent literal, and a family whose fixed parent is only different
by ASCII case. `%41` and `A` remain distinct under the current OPC folding;
the review expects the wrapper to follow `PackURI` and `PartNameSet`, not an
invented URI equivalence.

## Budgets and allocation

`max_patterns` must bound descriptors, and `max_pattern_bytes` must bound the
sum of owned literal/prefix/suffix bytes with checked arithmetic. `max_entries`
must charge the product of indexed `count` and pattern count with a checked
`u64` multiplication and checked accumulation. The `first + count - 1`
endpoint calculation must be checked before any endpoint is built. Zero-count
families need a defined policy and must not accidentally validate or consume
one member; nonempty families with no patterns must be refused unless the API
explicitly defines them as no-ops.

The builder must reserve only for descriptors actually supplied. It must not
allocate `max_patterns`, `max_entries`, or any expanded name vector up front.
`finish` and symbolic validation must inspect descriptors and intervals only;
no loop may run from `first` through `last` to validate uniqueness. A
`max_name_bytes` computation must use checked prefix/suffix plus decimal-width
arithmetic and remain finite for `u64::MAX`.

At runtime, deriving the current name may use one reusable bounded buffer or
one active owned name. It must not append completed names to a plan, cursor,
OPC `PackURI` history, or duplicate set. The active name and central record
are allowed to exist for one in-flight member; they must be released after
successful finalization. Any `String`/`Vec` operation whose capacity depends
on the emitted member count is a blocker.

## Cursor consumption and failure state

Every low-level name-taking route must pass through the same generated cursor:
stored and deflated sized writes, reader/stream writes, and owned entry
writers. A route must check the exact generated raw spelling (and the ordinary
normalized spelling) before writing its local header, and advance exactly once
only after the member's central record has been successfully published. A
preflight name mismatch must consume neither the plan nor the output budget. A
payload, sink, central-spool, or finalization failure after output starts must
poison the writer and must not make a later route able to skip or replay the
failed ordinal.

An unfinished owned entry must not advance the plan. Dropping it must leave the
forward-only output state unusable in the same way as the existing writer.
Repeated, skipped, reordered, and foreign names must be typed refusals. The
archive finish operation must reject an unexhausted plan before emitting an
EOCD/central-directory tail; a successful finish must require exact cursor
exhaustion. Conversely, a failed finalization must not silently mark the
ordinal as consumed. Test both zero entries and a plan with one remaining
entry, including a failure on the final record.

At the OPC layer, generated mode must not retain `PartNameSet`, prepared
ancestor vectors, or completed `PackURI` values. All borrowed and owned part
routes, stored and deflated routes, and the root-part refusal must enforce the
same lower cursor. Ordinary constructors must retain their existing exact,
ASCII-equivalent, and ancestor/descendant checks unchanged.

## Review status

Static source review found no remaining concrete symbolic acceptance bug after
the indexed cursor was split into folded proof matching and exact runtime
matching. The matrix and byte-level tests remain evidence obligations,
especially the unequal-depth variable-to-fixed-parent case, failed-finish
consumption, and the PPTX plan's exact member order: a correct lower proof does
not prove that PPTX supplied every fixed member and every layout, slide, and
relationship pair in the actual emission order.

## Current implementation observations

The OPC adapter now translates prefixes as borrowed input slices and lets the
lower builder copy them into its bounded descriptors; the earlier temporary
`Vec<String>` borrow hazard is resolved. The adapter's separate counters
preflight descriptor bytes and entry count before endpoint construction, which
is the right failure ordering for its public owner-local API.

The cursor transition and root-level component-depth issues identified during
the first pass have since been addressed: transitions initialize a following
indexed descriptor from its `first` endpoint, and an empty parent path counts
as zero components. Keep regression tests for a literal followed by a nonzero
indexed family, two indexed families with different `first` values, and a
root-level family versus a deeper fixed-parent family.

The PPTX plan's descriptor order currently appears to match the existing
emission order (content types, package relationships, static members and
layout pairs, theme/notes members, presentation and its relationships, then
slide/slide-relationship pairs). That correspondence needs an exact byte
parity test, since the plan only proves names and cannot detect a missing or
duplicated static write. The source should also test that a sink remains
untouched when plan construction or endpoint validation fails.

The generated route intentionally imposes a stricter raw-name contract than
the ordinary writer: `/expected.bin` and `expected.bin/` must be refused even
though ordinary admission normalizes them to `expected.bin`. Cross-route tests
should retain those cases and assert that the generated cursor is unchanged;
ordinary tests should separately preserve the existing normalization behavior.
Likewise, the generated constructor is required to preflight its complete
plan against the transport entry/name budgets, so a zero or insufficient
`max_entries` limit should fail construction before output. This stricter
preflight must remain local to the generated constructor unless the ordinary
API is explicitly changed.

The raw-name mismatch identified in the first pass is resolved: generated
admission checks the caller spelling before trimming/normalization, while the
ordinary policy still checks the normalized name. Keep the leading-slash and
trailing-slash refusal tests because they protect this stricter generated-mode
contract.

The indexed builder now computes and checks the aggregate static byte budget
before composing endpoint names; the OPC adapter does the same before its
`PackURI` endpoint allocations. This resolves the first-pass temporary-name
allocation issue. Endpoint scratch still includes at most the bounded decimal
width in addition to admitted static bytes, and no allocation scales with the
emitted range.

The indexed cursor admission blocker is resolved: `check_next` now uses
byte-exact prefix/suffix matching through a private exact helper, while the
symbolic conflict and ancestor proof retains ASCII folding. The focused cursor
regression rejects both prefix and suffix case variants, and the public
cross-route matrix covers indexed and literal plans across all routes with no
output or cursor consumption on rejection.

The independent proof-test `INTERNAL_DIGITS` fixture now keeps digits away from
the numeric-slot boundaries, with a separate rejection assertion for a
digit-leading suffix. That earlier test failure was a fixture defect rather
than evidence of an unsound symbolic acceptance.

## Scratch-harness follow-up

The PPTX and ZIP benchmark lanes now propagate `create_new`, writer/finish,
metadata, and quota errors before attempting unlink. Their tests preserve
pre-existing oracle, warmup, and sample paths, verify successful cleanup, and
retain a failed quota oracle path; the ZIP pre-existing-path test also checks
`io::ErrorKind::AlreadyExists` rather than accepting any nonempty message. The
ZIP documentation now states the same exclusive scratch-directory contract as
the PPTX lane. These changes resolve the earlier destructive-cleanup finding;
the prior warning in `measurement-review.md` should be refreshed by its owner.

Cleanup now follows the final output/parity checks in both lanes, so a
post-operation mismatch also retains its scratch artifact. The ZIP lane's
index-based timed filenames remain an intentional independently reviewed
choice: the deduplicated corpus index and method/repeat/sample dimensions make
them unique, and no ownership or output invariant depends on embedding the
member count in that filename.

## OPC percent-encoding follow-up

The OPC adapter now rejects an indexed prefix ending in `%` before endpoint
composition or lower-plan mutation. This closes the dynamic partial-triplet
case: `/%` plus a decimal slot can form `%00` through `%09`, and a suffix can
complete the triplet, so endpoint-only checks are insufficient. The rejection
is atomic because the adapter's budget counters are only committed after the
lower plan accepts the translated descriptor; endpoint and translation
failures likewise leave the inner builder unchanged.

The guard is complete for the remaining partial-triplet forms in the current
language. A prefix ending `%A` or `%a` is safe: the first inserted decimal byte
produces `%A0` through `%A9`, all outside the unreserved range that
`PackURI::validate_percent_encoding` rejects. A prefix ending `%4` (or any
numeric hex nibble) is already refused by the ZIP numeric-boundary rule before
OPC translation, so `%40`/`%41` variability cannot enter the accepted language.
Percent sequences wholly in fixed prefix/suffix text are checked by both
endpoints, and all other `PackURI` predicates (slashes, controls, spaces,
queries/fragments, dot segments, trailing dots, encoded separators, and root)
depend only on fixed bytes or are made safer by the inserted decimal digit.
The direct `%1A`/`%9A` accepted cases, `%4A` rejected case, partial-prefix
refusal, and fixed `%20` family regression cover the boundary; an accepted
`/%A` family would be a useful additional guard against over-rejection.

## Final validation binding

The final source manifest is
`38d17523810931fd9e41d3ea6566f2542869b32e32a09148b58fa35f9fd3f080`.
The ZIP/OPC/PPTX release test gate passes 1,887 tests, with five ignored,
including the generated-name proof, all name-taking route checks, exact PPTX
parity and scratch failures. The harness gate passes 387 tests, with one
ignored, including scratch ownership and full-member CRC failures. The final
24-process measurement matrix also passes exact-byte, complete physical-member
and semantic reopen checks. These receipts close the test obligations above
for the implemented restricted language; they do not broaden it to arbitrary
names, existing-document append, or general editing.
