# 0794 source review — shared checked XML attribute state

Review scope: the five dependency-isolated copies of the fail-fast checked
attribute iterator (`litchi-opc`, `litchi-ole-common`, `litchi-sign`,
`litchi-xldm`, and `xml-minifier`). The `litchi-formula` OMML module has a
different lenient `first_wins` contract and is intentionally outside this
review. The `after` archive contains the candidate described below; the source
review records corrected test issues and the remaining value-scan tradeoff.
This is source reasoning, not a measured malformed-input latency result.

## Source facts that constrain the patch

The baseline checked iterator delegates the first 32 successful attributes to
quick-xml's checked `Attributes`. On the first successful attribute quick-xml
pushes a raw `Range<usize>` into its `IterState.keys` vector. At the 33rd
attribute the wrapper disables quick-xml duplicate checking, rebuilds the first
32 names into its own `BTreeMap<Name<'a>, usize>`, and checks the remainder in
that ordered map. The candidate can remove the first small-tag allocation only
if it uses `with_checks(false)` before consuming the first attribute and owns
an allocation-free bounded prefix for the first names. Merely adding a second
map around `tag.attributes()` leaves the quick-xml vector allocation in place.

The five copies must keep the same code after the existing documentation,
visibility, and test-module-path normalization performed by
`every_copy_of_the_module_is_the_canonical_code`. The formula helper is not a
copy. No manifest edit is in the frozen six-file source allowance; therefore a
new `smallvec` or `arrayvec` dependency cannot be assumed in the four copies
that do not already declare it. A fixed array (or equivalent dependency-free
representation) is the viable inline shape. A bounded vector spill through
the existing 32-name threshold is compatible with the source contract, but
its growth and its 4-to-5 transition must be measured explicitly.

The existing external tests remain authoritative. In particular, the 32-name
distinct case expects zero `Name::cmp` counts, while the 33-name case expects
the ordered fallback to perform counted comparisons. Inline duplicate checks
should compare raw byte slices directly; using `Name == Name` calls the
instrumented `Ord` implementation and changes that unchanged test oracle.
The fallback must continue to use `Name` keys so the large-tag comparison and
quadratic-regression tests still observe the ordered path.

## Required semantic equivalence

The public result is the exact sequence yielded by quick-xml's checked
iterator through its first error. The candidate must preserve all of these
details:

* Raw QName bytes are the duplicate key. Comparison is exact and
  case-sensitive; prefixes are not namespace-resolved or normalized here.
  Namespace-expanded duplicate rules remain the responsibility of callers.
* A key is checked after quick-xml has found its `=` and before its value is
  parsed. Thus a duplicate whose value is unquoted, missing, or missing its
  closing quote must yield `AttrError::Duplicated(new_position,
  first_position)`, rather than `UnquotedValue`, `ExpectedValue`, or
  `ExpectedQuote`. `ExpectedEq` is different: a key without an `=` has not
  reached duplicate checking and must retain that syntax error.
* `name_at` must start at the byte immediately after the last successful
  value, including adjacent attributes (`a="1"a="2"`) and arbitrary XML
  whitespace. The error and previous positions are byte offsets into the
  original `BytesStart` contents, not character offsets. Empty names and
  unusual raw key bytes follow quick-xml's parser and must not be normalized.
* Any lexical error ends the checked iterator. Repeated calls after an error
  must return `None`; no item after the first error may be exposed. This
  includes errors from the unchecked underlying iterator, and also a
  duplicate reconstructed from a malformed value.
* Successful `Attribute` values and keys must be the original borrowed
  quick-xml slices. The duplicate index may retain borrowed key slices and
  positions, but it must not copy values, unescape names, or let a slice
  outlive its `BytesStart`.
* `CheckedAttributes` is `Clone`, `Debug`, and `FusedIterator`. Cloning at any
  point must clone the parser cursor, inline entries, ordered fallback, and
  `next_start` together. The clone must not share mutable duplicate state with
  the original.

The safest implementation obtains items from one `Attributes` iterator with
duplicate checking disabled and records each successful raw key. For an
`Err`, it applies the existing `duplicate_before_value` rule against the
recorded names, then marks the phase `Done`. For an `Ok`, it checks the key,
updates `next_start` only after accepting the item, and returns that exact
`Attribute`. An implementation that calls the unchecked iterator only after
the value has been fully parsed still needs the malformed-duplicate repair;
checking only successful items changes error precedence.

At the candidate's inline boundary, the first four accepted names remain in
the fixed array; the fifth accepted name may create the bounded vector spill.
The 32nd accepted name remains in the vector, and the 33rd operation must make
all 32 names available to the ordered map before checking a new name. A
duplicate at either spill point, including a duplicate with a malformed
value, must report the first position from the prior store. The transition
must not lose positions, insert the duplicate, or emit it as `Ok`.

## Security and cost constraints

Inline comparisons have a bounded constant maximum only when the capacity is
fixed. The ordered fallback must remain the existing `BTreeMap` path; an
unkeyed `HashMap`/`HashSet` or a hash prefilter reintroduces attacker-chosen
collision behavior and violates the reason this substrate exists. If a chosen
inline capacity causes the ordered tree to appear before the existing 32-name
threshold, it needs direct measurement and must not be treated as equivalent
to the current threshold. A capacity above 32 needs an explicit worst-case
and stack-cost justification.

An inline array of borrowed names increases every iterator value's stack/enum
size. The candidate review must record its size or otherwise inspect the
generated layout, and the public allocation lane must guard the 0-attribute,
small-tag, and >32 fallback cases. A fixed array must remain bounded even when
input has arbitrarily many attributes; only the already-existing ordered map
may grow with the tag.

The `with_checks(false)` call must stay private to this duplicate-checking
substrate. Callers must continue to use `checked_attributes()` for strict
reads and `unchecked_attributes()` or `first_wins` only where their existing
contract permits it. No limit, entity handling, namespace handling, or
caller-side decoded-byte accounting may move into this optimization.

## Required review evidence before retention

Before the candidate can be considered semantically valid, the focused tests
must compare candidate output with quick-xml 0.41 through the first error for
0, 1, 2, 4, 8, 16, 31, 32, 33, and 64 attributes. Include duplicates at the
first, middle, 32nd, 33rd, and last positions; adjacent attributes; all four
XML whitespace bytes; namespace-like keys; empty/long/non-ASCII raw bytes;
`ExpectedEq`, `ExpectedValue`, `UnquotedValue`, and `ExpectedQuote`; and each
malformed-value duplicate variant. Include `Start` and `Empty` event-backed
`BytesStart` values, clone/resume checks, and the exact-capacity/spill cases.
The existing external differential and copy-synchronization tests must remain
unchanged and pass.

The direct allocation lane must construct the tag before entering its timed
region and prove that the candidate's lower allocation count is from the
duplicate store. Public package-level counts alone cannot attribute this
vector. The frozen 0794 profile reconciliation compares owner cost with
`allocation_calls` alone because successful reallocations are already included
in that counter; it must not sum the two fields.

Disposition remains pending until the candidate passes the semantic suite,
all five-copy synchronization, no-new-hash/no-unbounded-inline checks, and the
frozen public resource and latency gates. A source-only micro-lane gain cannot
justify adoption if small ordinary tags regress or the spill path increases
resource use.

## Review of the current candidate snapshot

The candidate now uses `Attributes` with checks disabled, four inline
`Name` entries, a `Vec<Name>` through 32 entries, and the existing ordered map
after that. The raw-byte comparisons in the inline and vector phases preserve
the unchanged comparison-count oracle, and the map's `Borrow<[u8]>` lookup is
consistent with its byte ordering. The five candidate bodies are otherwise
aligned. I found no output or state-machine mismatch in the current source
review; the following concrete items still need to be carried into root
validation and the final report.

The final source snapshot was checked against the archived files and the live
worktree after the retained visibility and test-binding corrections. The
normalized body hash (owner visibility and external test-module path removed)
is `692eae80ee27c4fcd2137625b5f07c576114b6d50a1c7af3c15cbe4f1a271f2e` for
all five copies. The exact after-source hashes, which also match the live
files, are:

| owner | after/live SHA-256 | exported surface |
| --- | --- | --- |
| `litchi-opc` | `692eae80ee27c4fcd2137625b5f07c576114b6d50a1c7af3c15cbe4f1a271f2e` | `pub` trait and iterator |
| `litchi-ole-common` | `1306cb95c9a913c44199e16b744073078165e330d6b42f078c096d240878eaa8` | `pub` trait and iterator |
| `litchi-sign` | `c2a57eac31f344ce12b3ed4fafd101cf2341ec48cf7217a44b177de52c249e1b` | `pub(crate)` trait and iterator |
| `litchi-xldm` | `c2a57eac31f344ce12b3ed4fafd101cf2341ec48cf7217a44b177de52c249e1b` | `pub(crate)` trait and iterator |
| `xml-minifier` | `c2a57eac31f344ce12b3ed4fafd101cf2341ec48cf7217a44b177de52c249e1b` | `pub(crate)` trait and iterator |

The packet records base commit `008bea923a150612213d869ed31e94017abeea4d`
and candidate patch SHA-256
`a234f9c0f13c81a9acc9bb1087d212e1fdba5a1b40d6331367c9112ee289e62d`.
`litchi-formula` is absent from this table and remains the separate lenient
helper described above.

1. An earlier candidate snapshot had a clone test that constructed
   `names(33)` and then appended `n3="again"`, making its all-`Ok` assertions
   false. The current snapshot has corrected this by using a 40-name valid
   tag for the all-`Ok` clone path and a separate failed tag for the duplicate
   path. Keep both cases when the candidate is archived.

2. An earlier snapshot changed the import to `use std::borrow::{Borrow, Cow};`.
   The current snapshot has restored the exact `use std::borrow::Cow;` source
   anchor and uses a fully-qualified `std::borrow::Borrow` implementation, so
   the unchanged external copy-synchronization test can still locate the
   canonical body. Do not regroup that import during later edits.

3. The candidate initially had the adjacent-attribute test construction bug
   described above. The current test uses `format!("e {}", ...)`, so the
   element name has an explicit one-byte separator and the empty-separator
   case now exercises adjacent attributes. Keep this construction in the
   archived candidate.

4. The candidate's unchecked lexical parser scans a quoted duplicate value
   before `SeenNames::check` sees the key. The baseline quick-xml checked
   iterator reports a duplicate after the key and `=` and before scanning the
   value. For a long quoted or unterminated duplicate value, the candidate can
   therefore scan more bytes before returning the same typed error. An unquoted
   value is rejected at its first non-whitespace byte; its entire tail is not
   scanned. Whitespace after `=` can also be scanned before duplicate recovery. This is a fail-fast and hostile-input cost change even though the
   returned `AttrError` and positions can match. The final root refinement covers 4096-byte quoted, unquoted, unterminated,
   and whitespace-only duplicate values at 1, 4, 5, 32, and 33 prior names.
   These tests establish result parity, not equal scanning work. The report must disclose the changed
   timing/work or preserve name-before-value detection. The operation's
   existing input/element limits must bound that scan, and no claim of
   equivalent fail-fast behavior should be made from short cases alone.

5. Four inline names intentionally move the first vector allocation from the
   first attribute to the fifth. That can be a reasonable measured capacity,
   but it is not a replacement for the old allocation-free 1–32 comparison
   path: the vector still grows at 8, 16, and 32 and the candidate's
   `CheckedAttributes` value is larger on every call. Resource and layout
   evidence must cover 0–4, 5–8, 9–16, 17–32, and the 33-name map transition;
   a package-level reduction alone does not establish that this tradeoff is
   safe.

6. The public iterator documentation says that the cost is quick-xml's
   linear check for the first 32 names. Read with the module-level statement
   that `with_checks(false)` disables quick-xml's duplicate check, this is a
   comparative-cost description of the same linear phase; the implementation
   actually performs those comparisons in its raw-byte inline/vector store.
   The sentence is therefore ambiguous about ownership, rather than an
   algorithm mismatch. Item 4 still records why matching returned precedence
   does not imply identical duplicate scan timing. No algorithm change is
   warranted from this wording alone.

7. `candidate-review.md` now agrees with `candidate.patch` and
   `manifest.json` on patch SHA-256
   `a234f9c0f13c81a9acc9bb1087d212e1fdba5a1b40d6331367c9112ee289e62d` and
   records the long-value scan qualification from item 4.

The first applied candidate failed the all-feature check because its OLE common
copy had inadvertently made the originally public trait and iterator
crate-private. Root restored exactly those two public visibilities before any
candidate timing. The failed source, patch, and compiler receipts are retained
under `candidate/compile-failure-1`, `quality-1`, and `candidate-fix.json`.
Canonical-body normalization alone does not prove each owner's export surface.

Quality attempt2 caught test-local `tag` constructor shadowing (E0618). Root retained the exact failed archive and logs and renamed the two local bindings in each copy before candidate performance measurements.
