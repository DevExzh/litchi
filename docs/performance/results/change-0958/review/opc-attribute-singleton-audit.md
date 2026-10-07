# 0958 OPC checked-attribute singleton audit

## Scope and evidence boundary

This is a read-only source audit at
`e29699f49af54e435145cb6051668f6929e7e16f`. It covers
`crates/litchi-opc/src/xml_attributes.rs`, the quick-xml 0.41.0 source pinned
by the workspace, all production uses of `CheckedAttributes` in litchi-opc,
and the retained 0796–0806 XML attribute experiments. No source, build, or
workload was changed or run.

The current canonical helper is SHA-256
`6c7e5aeb66f1ca420f5e11af0bfa307ae1b5fd6aa29b4d3b5af57d6b9ec6fb56`; its
shared tests are
`6c7d024783556aab85af7c142953fe0a09a848e9f6afde814f0f83523b29f098`.

## Where the allocation occurs

quick-xml's `Attributes` contains an `IterState` with an empty
`Vec<Range<usize>>` for duplicate keys (`quick-xml-0.41.0/src/events/attributes.rs:1018-1048`).
On the first key that has an equals sign, `check_for_duplicates` compares the
empty prefix and pushes that key (`:1123-1153`). That first push normally causes
the key vector's first heap allocation. A key-only XML error returns through
`key_only` without checking or pushing a key; a valid key followed by a malformed
value has already been pushed before the value error is returned. The hash set is
created only after the 32-name threshold (`:1156-1191`).

The OPC helper starts with quick-xml's checked iterator and lets it handle the
first 32 successful names (`crates/litchi-opc/src/xml_attributes.rs:81-119,
222-245`). It then disables quick-xml checking before the 33rd name and builds
its own `BTreeMap` (`:121-176`). Thus a singleton's first valid attribute can
pay for quick-xml's key vector even though no prior name exists; the late path
already has a bounded `O(n log n)` fallback.

`BytesStart::attributes_raw` exposes the undecoded bytes after the element name
(`quick-xml-0.41.0/src/events/mod.rs:296-301`). For a successfully parsed first
attribute, `end_of` in the OPC helper computes the byte immediately after the
closing quote (`xml_attributes.rs:253-260`). An exact `trailing_whitespace`
check from that offset, using quick-xml's four whitespace bytes, can prove that
no second item exists without parsing the tail. The probe must treat every
non-whitespace byte as a possible next item, including a malformed key, missing
equals sign, or missing separator.

## Feasible minimal shape and its limits

The smallest semantically defensible experiment is the 0797 shape:

1. Construct the first `Attributes` iterator with duplicate checks disabled.
2. On the first `next`, return the unchecked first result. A first lexical error
   is already the checked result because no earlier key exists.
3. If the first successful value is followed only by tag whitespace, mark the
   iterator done. Its first result and all later `None` results match the checked
   iterator, without creating quick-xml's key vector.
4. If any non-whitespace byte remains, defer to a fresh canonical checked
   iterator. Consume and discard its first successful item to seed the duplicate
   state, then return the second item on the following call. This replay occurs
   only on the multi-item path.

This preserves laziness: the first call does not parse or report an error from a
second item. On the second call, canonical quick-xml checks the second key before
reading its value. A duplicate followed by a long quoted or unterminated value
therefore remains `AttrError::Duplicated`, rather than scanning the value first.
The fallback also retains the OPC helper's 32-name handoff and ordered-map bound.

The extra proof is a linear scan of the bytes after the first value. It does not
decode the first value and is bounded by the already admitted tag bytes. It must
not become a parser that advances unchecked through a hostile tag: any non-
whitespace tail must enter the canonical checked/replay path. A dual iterator
that avoids replay would enlarge the public iterator's state and has no evidence
here; the earlier candidates that carried more state also exposed new boundary
costs.

The fast path does not bypass caller limits. The first `Attribute` still reaches
the caller, which charges its raw/decoded value bytes and increments any
per-element count. A second or malformed tail must still be visited on the next
call, so the following remain unchanged:

* `MAX_SOURCE_XML_ATTRIBUTES_PER_ELEMENT` and
  `MAX_SOURCE_RELATIONSHIP_ATTRIBUTES_PER_ELEMENT` (4,096 each);
* `ReadLimits::max_xml_attribute_bytes`, input/part/XML event budgets, and
  execution-context work accounting;
* archive member-name limits, including the custom relationship-owner name
  checks; and
* typed allocation and malformed-XML errors.

The fallback must preserve the current hostile-input guarantee: no unchecked
linear duplicate state may grow with the tag. A tag with more than 32 names
must still reach the OPC ordered map, and its first duplicate or malformed-value
error must have the same position and precedence.

## Exactness obligations

The existing contract is stronger than “returns the same values.” The helper
promises quick-xml's first error, byte positions, and no item after that error
(`xml_attributes.rs:45-63`). A candidate therefore needs differential cases for:

* zero attributes (`e`, `e `), one valid attribute with spaces/tabs/CRLF, and
  self-closing tags;
* one valid attribute followed by a distinct valid attribute, a key-only flag,
  missing `=`, unquoted value, empty value, unterminated quote, and a non-space
  missing separator;
* a second duplicate whose value is valid, unquoted, empty, quoted, or
  unterminated, proving duplicate-before-value precedence;
* unusual raw keys beginning with `=` and whitespace around the key and `=`,
  including the corrected 0800 behavior;
* first errors in the unchecked first item, where no duplicate may be reported;
* the 32/33 transition, a late duplicate, and a late malformed duplicate;
* `Clone` before first advance, after a singleton first result, after a first
  result with a pending replay, after each error, and after terminal `None`; and
* `FusedIterator` behavior after empty, error, and natural exhaustion paths.

The shared test file already covers broad random/differential inputs, error
precedence at the 32-name switch, borrowed values, clone and fused behavior,
and the ordered-map comparison bound (`crates/litchi-opc/src/xml_attributes/tests.rs:150-468`).
The singleton candidate must add the new phase boundaries rather than weaken
those tests. In particular, the long duplicate-value cases from 0796 are
mandatory: the first rejected candidate parsed the duplicate value before its
own check and was about 24 times slower on that tail.

`BytesStartExt`, `CheckedAttributes`, `Clone`, `Debug`, and the checked/unchecked
methods are public at the OPC boundary. The canonical-copy test requires the
same implementation body in five helper copies: litchi-opc, litchi-ole-common,
litchi-sign, litchi-xldm, and xml-minifier. The 0806 quality run also showed that
the OLE helper's public visibility is consumed by litchi-crypto; a visibility
narrowing is not an acceptable optimization.

## Production call-site surface

There are 16 production `checked_attributes()` call sites in litchi-opc (tests
and the helper's internal unchecked path excluded):

| source locations | consumption shape |
| --- | --- |
| `xml_splice/owned.rs:249`, `:700`; `part.rs:573`; `source_backed.rs:745,866`; `xml_splice.rs:1105,1183`; `pkgreader.rs:2337`; `content_type.rs:555`; `relationships.rs:1601` | Drain the iterator unless an error or an enclosing failure stops the loop. |
| `xml_splice/owned.rs:607`; `package/relationships.rs:976` | Return/break as soon as the sought attribute is found, otherwise drain. |
| `package/content_types.rs:407` | Drain the `Override` attributes before deciding whether its `PartName` is removable. |
| `package/relationships.rs:870`; `source_backed/content_types_plan.rs:792` | Return as soon as the sought `Id`/`Extension` is found. |
| `source_backed/content_types_plan.rs:771` | Drain to find and validate the required `PartName`. |

The table describes source control flow, not an attribute-count measurement.
The early-return sites can observe the first attribute and drop without a
second `next`; the drain sites still exercise the replay path for every
multi-attribute tag. Caller-owned `HashSet`, `HashMap`, `SmallVec`, and decoded
`String` allocations remain even if quick-xml's duplicate-key vector disappears.

The best existing shape census is 0798's caller-thread PPTX capture. It saw
295,600 `CheckedAttributes` instances and 438,800 yielded attributes across two
repeats; every observed iterator reached `None`. In the large capture, 30,374
instances (49.528%) had one attribute and 30,816 (50.249%) had two, for 99.777%
combined. That census covered the OPC helper as re-exported through OOXML common
PPTX callers; it excluded copied helpers, unchecked/lenient iterators, skipped
tails, other threads, and other workflows. It is evidence that singleton work can
be common, but it is not a current source-backed OPC-stage distribution.

## Historical decision boundary

The exact idea has already been tried and must not be revived from an allocation
count alone:

* 0797 (`docs/performance/0797-xml-first-attribute-preflight.md`) implemented
  the unchecked first item, whitespace proof, and canonical replay fallback.
  One-attribute consumption improved 35.768%, but distinct-two, distinct-three,
  and distinct-four consumption regressed 42.987%, 33.387%, and 27.033%; 20 of
  33 consumption rows triggered the frozen regression rule. It was not advanced.
* 0799 removed more replay by carrying two raw keys. One- and two-attribute
  consumption improved 34.898% and 16.716%, but six protected errors directly
  after the two-item prefix regressed roughly 50–56%; the candidate was rejected.
* 0800 repaired the leading-`=` duplicate-position/error-precedence defect found
  while preparing the no-replay candidate. Any new source must include that
  corrected baseline.
* 0801 carried duplicate state farther without value replay. The protected
  empty-tag path regressed 7.542%, and twelve other rows regressed as much as
  68.034%; it was rejected.
* 0805 combined the no-replay prefix and bounded own state and passed its direct
  micro gates (one attribute −38.646%, two −19.080%), but still had diagnostic
  regressions at four attributes (+22.649%), duplicate-after-four (+10.726%),
  and syntax tails after four (+28% range). It was eligible only for workflow
  trials, not adoption.
* 0806 ran those workflow/resource/cross-format gates. No capture or lifecycle
  row met the 3% benefit rule; the largest eligible row, large lifecycle, was
  3.105% slower. The resource lane did show a large allocation reduction, and
  nested duplicate-check calls fell from 61,342 to zero in the profile lane, but
  that did not establish a workflow benefit and production was restored.

The new heap profile can therefore rank duplicate checking against the Number
clone candidate and identify which phase deserves attention. It cannot erase the
0797–0806 semantic and public-workflow evidence. If a fresh experiment is ever
justified, the only narrow candidate supported by the prior boundary work is the
singleton-only 0797 shape with canonical checked fallback, rebased on 0800 and
with fresh workflow/resource gates. It should remain an evidence candidate until
the current profile shows material stage attribution and the new gates pass.
