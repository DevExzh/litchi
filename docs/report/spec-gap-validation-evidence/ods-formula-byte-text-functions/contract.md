# ODF 1.4 byte-position text-function evaluator contract

This contract defines the seven byte-position text functions in OpenFormula
1.4 Part 4 §6.7: `FINDB`, `LEFTB`, `LENB`, `MIDB`, `REPLACEB`, `RIGHTB`, and
`SEARCHB`. Part 4 deliberately leaves the meaning of `ByteLength` and
`BytePosition` implementation-dependent. This evaluator selects one explicit
profile: byte positions count UTF-8 octets in the semantic Text value. The
selection makes the functions deterministic and testable; it does not claim
that every ODF host uses UTF-8 or that byte-position formulas are portable
between hosts.

The contract is for the next bounded evaluator batch. It records the selected
semantic and resource boundary; it does not claim that production dispatch,
validation, native compatibility, or an independent oracle has landed. No
ambient locale, process code page, external provider, or native DBCS mode is
consulted by this profile.

The normative source is the repository-local ODF distribution:

| Source | SHA-256 |
| --- | --- |
| archive `3rdparty/specs/OpenDocument-v1.4-os.zip` | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| member `part4-formula/OpenDocument-v1.4-os-part4-formula.html` | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |

The exact normative entries are §§6.7.1–6.7.8. Their ordinary-function
counterparts are `FIND`, `LEFT`, `LEN`, `MID`, `REPLACE`, `RIGHT`, and
`SEARCH` in §§6.20.9, 6.20.12–6.20.13, 6.20.15, 6.20.17, 6.20.19, and
6.20.20. Scalar conversion, non-scalar iteration, references, errors,
normalization, and resource handling follow the repository's selected text
evaluator profile unless this document gives a byte-specific rule.

## Exact normative scope and signatures

The signatures below preserve the Part 4 pseudotypes. A semicolon separates
arguments, square brackets mark optional arguments, and the defaults in the
last column are inherited from the ordinary counterpart. `ByteLength` and
`BytePosition` are integer-valued pseudotypes; this profile gives them the
UTF-8 meaning defined below.

| Function | Part 4 signature | Return | Domain and default |
| --- | --- | --- | --- |
| `FINDB` | `FINDB(Text Search; Text T [; BytePosition Start])` | `BytePosition` | `Start ≥ 1`; omitted `Start` is 1. Case-sensitive literal search, as `FIND`. |
| `LEFTB` | `LEFTB(Text T [; ByteLength Length])` | `Text` | `Length ≥ 0`; omitted `Length` is 1. Selects a UTF-8-byte-limited prefix, as `LEFT`. |
| `LENB` | `LENB(Text T)` | `ByteLength` | No additional constraint. Counts UTF-8 octets. |
| `MIDB` | `MIDB(Text T; BytePosition Start; ByteLength Length)` | `Text` | `Start ≥ 1`, `Length ≥ 0`; extracts complete scalars in the selected byte span, as `MID`. |
| `REPLACEB` | `REPLACEB(Text T; BytePosition Start; ByteLength Len; Text New)` | `Text` | `Start ≥ 1`, `Len ≥ 0`; removes complete scalars in the selected byte span and inserts `New`, as `REPLACE`. |
| `RIGHTB` | `RIGHTB(Text T [; ByteLength Length])` | `Text` | `Length ≥ 0`; omitted `Length` is 1. Selects a UTF-8-byte-limited suffix, as `RIGHT`. |
| `SEARCHB` | `SEARCHB(Text Search; Text T [; BytePosition Start])` | `BytePosition` | `Start ≥ 1`; omitted `Start` is 1. Case-insensitive literal search, as `SEARCH`, with the fold profile below. |

These seven names are the complete §6.7 scope. The registry must not add
`DBCS`, `LEFTB`/`RIGHTB` aliases for another encoding, or Microsoft/host-only
names. A host-specific byte family is a separate qualified or compatibility
profile and must not change the meaning of these names in the base profile.

Part 4's short entries say “as” the ordinary function and do not repeat every
ordinary constraint. This contract makes the inherited rules concrete:

* `FINDB` and `SEARCHB` use the ordinary search defaults and return `#VALUE!`
  when no match exists. `FINDB` is case-sensitive and literal. `SEARCHB` is
  case-insensitive and literal in the base profile; host regular-expression
  and wildcard properties are not consulted.
* `LEFTB` and `RIGHTB` default a missing optional length to 1. Zero returns
  empty Text, and a length greater than the available UTF-8 bytes returns the
  complete Text value.
* `MIDB` requires both `Start` and `Length`. Start beyond the end of the Text
  returns empty Text, as `MID`; zero length also returns empty Text.
* `REPLACEB` inherits the selected ordinary `REPLACE` clamp profile. A zero
  length inserts `New`; a start beyond the end appends at the end after the
  selected integerization and boundary mapping. Its `Start` and `Len` are
  integer-valued byte arguments even though the ordinary function's source
  spelling calls them `Number`.

Required arity is strict. A zero-argument call, a missing required argument,
or a supplied `Missing` value in a required slot produces generated
`#VALUE!`. An omitted optional AST slot receives the documented default. A
supplied value is converted and validated even when it is zero, false, empty,
or otherwise falsy.

## Text and byte-argument conversion

The byte functions operate on semantic Text, not on an underlying storage
encoding. The common conversion profile is:

* Text is retained without Unicode normalization. Number values use the
  invariant shortest round-trip decimal spelling with a period. Logical values
  use `TRUE` and `FALSE`; Empty values use empty Text. A Formula Error
  propagates unchanged. Complex values are not silently formatted as ordinary
  Text and produce `#VALUE!` unless a producer has already supplied Text.
* A Number argument for `ByteLength` or `BytePosition` must be finite. Logical
  values convert to 0 or 1. Text uses the profile's locale-independent decimal
  grammar; malformed Text is `#VALUE!`, and a parsed non-finite value is
  `#NUM!`. An Empty reference converts to numeric zero under the common
  reference conversion rule.
* Generic integer conversion truncates toward zero after Number conversion.
  The ordinary counterpart's explicit `INT` rule is retained where it is
  normative: `LEFTB` and `RIGHTB` use `INT(Length)`, and `MIDB` uses
  `INT(Start)` and `INT(Length)`, just as `LEFT`, `RIGHT`, and `MID` do.
  `FINDB` and `SEARCHB` use the generic integer conversion for `Start`.
  `REPLACEB` follows the selected ordinary `REPLACE` profile: validate finite
  raw `Start ≥ 1` and `Len ≥ 0`, then truncate each toward zero before byte
  boundary mapping.
* A finite negative or otherwise domain-invalid length/position is generated
  `#VALUE!`. A non-finite numeric argument or non-finite intermediate is
  generated `#NUM!`. These generated errors are distinct from typed resource,
  cancellation, source, provider, and allocation failures.

The conversion profile intentionally does not use a process locale or native
code page. A provider that stores text in UTF-16, Latin-1, Shift-JIS, or any
other representation first supplies semantic Unicode Text; the byte operation
then encodes that Text as UTF-8 for counting and matching.

## Scalar and matrix argument context

All seven functions return one scalar Text or integer-valued Number per
evaluation position. None has a `ForceArray` or sequence pseudotype. The
ordinary §3.3 context applies:

* In scalar mode, an inline Array uses its `[0,0]` element and a multi-cell
  Reference uses implied intersection. A known ReferenceList passed where a
  scalar Text, ByteLength, BytePosition, or `Scalar` is required is rejected
  as `#VALUE!` before reading cells from that refused descriptor. Eager
  evaluation of another admitted argument may still read its own descriptor;
  the per-argument refusal is not a whole-call zero-read guarantee.
* In matrix mode, each scalar parameter lifts position-by-position. Scalars,
  singleton arrays, one-row arrays, and one-column arrays broadcast under the
  ordinary rectangular rules. This includes `Search`, `T`, `Start`, `Length`,
  `Len`, and `New`; an out-of-range two-dimensional input contributes the
  ordinary `#N/A` output at that position. The result remains scalar at each
  output position.
* A function returning an Array elsewhere is not collapsed by the byte
  consumer. The §3.3.2.2.1 rule selecting an array producer's own scalar
  input, such as a size argument to `MUNIT`, applies to the producer input;
  it does not make a produced Array opaque to a later `LENB`, `MIDB`, or search
  operation. The byte functions themselves remain position-sensitive scalar
  consumers.

Formula evaluation is eager unless a separate function explicitly says
otherwise. Formula Errors in admitted inputs are retained and propagated by
the ordinary leftmost-error rule. A typed `Unsupported`, `ResourceLimit`,
`Allocation`, `Cancelled`, `SourceChanged`, or
`SourceVersionAvailabilityChanged` failure remains an `EvaluationFailure`; it
is never converted to a formula Error and is not caught by `IFERROR` or
`IFNA`. If a call retains an ordinary formula Error while continuing an
admitted scan, a later typed failure supersedes that retained formula Error.

## UTF-8 byte model

Let `B(T)` be the UTF-8 encoding of the semantic Text `T`, excluding any BOM
not present in `T`, and let `n = |B(T)|`. UTF-8 is always well-formed because
semantic Text is a sequence of Unicode scalar values. No normalization or
transcoding to a host code page occurs before this encoding.

`LENB(T)` returns `n`. A byte position is one-based: a zero-based byte offset
`o` is reported as position `o + 1`. Scalar starts and the end sentinel are
valid boundaries. Thus a Text with `n` bytes has valid boundary positions from
1 through `n + 1`, with `n + 1` denoting the position immediately after the
last byte. `LENB("")` is zero and the empty Text has only boundary position 1.

An interior position is a one-based position whose zero-based offset falls
inside a multi-byte UTF-8 encoding rather than at its first byte. The selected
profile snaps an interior **start** backward to the first byte of the scalar
that contains it. It never emits a partial UTF-8 sequence. This choice is
deliberate because ODF leaves the behavior implementation-dependent; callers
that need exact scalar boundaries can use the boundary positions returned by
`LENB`-based reasoning. A start at `n + 1` is the end sentinel and is not
snapped.

Byte lengths count octets, but returned Text always contains complete scalars:

* `LEFTB(T; L)` starts at the first byte and returns the longest prefix whose
  complete scalar encodings use at most `L` bytes. If the next scalar would
  cross the limit, that scalar is omitted. For example, for
  `T = "Aé界🙂"` whose UTF-8 byte lengths are `1+2+3+4`, `LEFTB(T; 3)` is
  `"Aé"`, not an invalid prefix of `界`.
* `RIGHTB(T; L)` returns the longest suffix of complete scalars whose encoded
  byte length is at most `L`. A scalar wider than the limit is omitted rather
  than split. This is the suffix counterpart of `LEFTB` and can therefore
  return empty Text for a positive length smaller than the final scalar.
* `MIDB(T; Start; L)` maps `Start` to its scalar-start boundary as described
  above, then consumes complete scalar encodings until adding the next scalar
  would exceed `L`. A mapped start at or after the end returns empty Text;
  zero `L` returns empty Text. An interior start may therefore select the
  containing scalar, while a length too small for that scalar selects none.
* `REPLACEB(T; Start; Len; New)` applies the same start snapping and complete-
  scalar byte budget as `MIDB` to determine the removed span, then inserts
  `New` at the mapped boundary. If `Len` does not cover the containing scalar,
  no complete scalar is removed but `New` is still inserted at that boundary.
  Start beyond the end clamps to the end sentinel under the selected ordinary
  `REPLACE` profile, so the operation appends there.

The byte limit applies to the selected input span, not to the inserted `New`
text. The resulting Text must still fit the caller's checked text/output
budget; it is never truncated or emitted as malformed UTF-8.

## Search and byte-position mapping

`FINDB` and `SEARCHB` search only at original Unicode scalar boundaries. The
returned position is the one-based UTF-8 byte position of the first source
scalar at which the match begins. It is never an offset into a case-folded
temporary string.

`FINDB` compares the original scalar sequence literally and case-sensitively,
with no wildcard or regular-expression instructions. `SEARCHB` uses the
ordinary text profile's Unicode default full case fold, without normalization;
the selected Unicode data version is part of the evaluator/cache profile.
Case-fold expansions such as `ß → ss` and ligature expansions participate in
matching. A candidate starts at a source scalar boundary and succeeds only
when the folded match ends at another original source-scalar boundary. The
reported byte position always maps back to the original source string.

The optional Start is converted and validated before searching. If it falls
inside a scalar, it snaps backward to that scalar's first-byte boundary. A
non-empty search beginning at the end sentinel has no match and returns
`#VALUE!`. An empty Search matches at the normalized Start when Start is in
the range 1 through `n + 1`; at the end sentinel it returns `n + 1`. A Start
below 1 or beyond `n + 1` is `#VALUE!`. The search is leftmost in source
scalar order, even when folding expands or contracts the query/source text.

Examples pin both byte units and boundary mapping:

| Expression | Result |
| --- | --- |
| `LENB("Aé界🙂")` | `10` |
| `LEFTB("Aé界🙂"; 3)` | `"Aé"` |
| `RIGHTB("Aé界🙂"; 1)` | empty Text, because `🙂` uses four bytes |
| `MIDB("Aé界"; 3; 2)` | `"é"`; position 3 is inside `é` and snaps to position 2 |
| `REPLACEB("Aé界"; 3; 2; "X")` | `"AX界"`; the snapped `é` span is replaced |
| `FINDB("界"; "Aé界🙂")` | `4`, the first byte of `界` |
| `SEARCHB("ss"; "aß")` | `2`, mapping the folded match back to `ß`'s first byte |
| `FINDB(""; "é"; 3)` | `3`, the end sentinel after two UTF-8 bytes |

The final example uses the requested position 3, which is the end sentinel,
not an interior position. An interior empty-search start is first snapped to
the containing scalar's first byte.

## Errors, resources, and source safety

The byte profile uses the bounded evaluator's resource and source rules:

* Text references are resolved with the normal per-cell read/work/cancellation
  charges. A retained or borrowed Text value may be inspected as UTF-8 bytes
  without cloning. A ReferenceList rejected by pseudotype or shape checking is
  not read; other eagerly admitted arguments can still be read.
* UTF-8 byte counting, scalar-boundary decoding, case-fold work, and output
  construction are checked operations. Work is charged for examined input and
  fold/output units. `LENB` and successful searches need no output Text
  allocation; slicing and replacement reserve checked output bytes before
  constructing the result. A failed reservation drops temporary state before
  returning its typed resource/allocation failure.
* Matrix output shape, per-position output Text, cumulative reference reads,
  and cumulative Text/output bytes are bounded. Implementations must not
  materialize an unbounded range or build an unbounded decoded-scalar vector
  merely to count bytes or find a match. A streaming boundary scan or bounded
  output builder is sufficient.
* Evaluation checks cancellation and the source/version fence before reads,
  after the operation, and before publishing a result. Typed resolver,
  cancellation, resource, allocation, and source-version failures bubble out
  unchanged. They supersede retained formula Errors according to the common
  precedence rule.
* A no-match result, invalid position/length, malformed numeric conversion,
  and wrong pseudotype are generated formula Errors (`#VALUE!` or `#NUM!` as
  specified above). They are never used to disguise a provider or resource
  failure. A malformed UTF-8 byte sequence cannot be a semantic Text value;
  an external provider that cannot produce valid Text returns its typed source
  or provider failure before this profile runs.

The byte encoding choice, Unicode case-fold data, normalization choice,
conversion profile, and matrix shape rules are part of the demand-cache
identity. A cache entry computed under a DBCS or different Unicode profile
must not satisfy a UTF-8-profile lookup. If demand caching is introduced for
these functions, an invariant Text/query may be reused only with complete
argument identity; position-sensitive Starts and matrix outputs remain
position-sensitive.

## Native DBCS and interoperability boundary

The normative text explicitly warns that byte positions depend on the
implementation's text representation. The selected profile measures UTF-8
octets of semantic Text:

* ASCII scalars use one byte, `é` uses two, `界` uses three, and `🙂` uses four.
  A native DBCS/code-page host may count the same values as one or two bytes,
  or may use a locale-specific lead/trail-byte rule. Such a host will produce
  different `LENB` values, search positions, and slice boundaries.
* This evaluator does not approximate DBCS as “one byte for Latin-1 and two
  bytes for everything else,” does not use Shift-JIS/Windows-932 tables, and
  does not select a code page from the ambient locale. A native compatibility
  profile must name its encoding, boundary behavior, search mapping, invalid
  sequence policy, and cache identity separately.
* Byte positions refer to the UTF-8 encoding of the semantic value, not to
  XML, UTF-16, provider, or storage offsets. Re-encoding the same semantic
  Text under a native profile is therefore an intentional compatibility
  difference, not an evaluator bug.

## Validation and unresolved profile points

An independent gate should cover all seven functions in Scalar and Matrix
contexts, including:

* exact signatures, optional defaults, required-argument and `Missing` errors,
  Boolean/Number/Text/Empty conversions, fractional integer conversion,
  negative and non-finite positions/lengths, formula-error propagation, and
  typed resolver/resource/cancellation/source failures;
* ASCII parity with the ordinary functions, empty Text, zero and oversized
  lengths, one-based positions, end-sentinel searches, no-match errors, and
  ordinary `REPLACE` append/clamp behavior;
* UTF-8 lengths for one-, two-, three-, and four-byte scalars; every interior
  start class; scalar-safe prefix/suffix/mid/replacement clipping; and the
  explicit `Aé界🙂` examples above;
* literal/case-sensitive `FINDB`, Unicode-fold `SEARCHB`, case-fold expansion
  and contraction, leftmost matching, source-boundary end checks, and mapping
  results back to original UTF-8 byte offsets;
* scalar projection, matrix lifting/broadcasting, ReferenceList refusal before
  descriptor reads, checked output reservation, cumulative work/byte limits,
  and source-version/cancellation fences; and
* a native divergence fixture showing that UTF-8 and a DBCS/code-page profile
  intentionally return different lengths and positions for non-ASCII Text.

The following are implementation-defined by ODF and are selected here rather
than left implicit: UTF-8 as the byte unit, one-based boundary positions,
backward snapping of interior starts, complete-scalar clipping, generic versus
explicit-`INT` argument conversion, Unicode full-fold matching for `SEARCHB`,
and the refusal of host wildcard/regex behavior in the base profile. These are
the points that require explicit reviewer acceptance before a production
implementation can claim this contract. Native DBCS interoperability remains
a separate profile and is not a blocker for the UTF-8 profile.
