# 0794 candidate review — bounded checked-attribute duplicate state

This is a source-only candidate for the allocation hypothesis recorded in
`docs/performance/results/change-0793/source-review.md`. It is archived under
`docs/performance/results/change-0794/candidate/`; no production source was
changed in this turn, and no build, native run, heap trace, or timing result is
claimed here.

The source-review decision is to investigate the checked XML attribute
iterator, whose small-tag duplicate state currently comes from quick-xml's
heap-backed `Vec<Range<usize>>`. The candidate changes the shared checked
attribute substrate in the five exact-copy owners:

* `crates/litchi-opc/src/xml_attributes.rs`
* `crates/litchi-ole-common/src/xml_attributes.rs`
* `crates/litchi-sign/src/xml_attributes.rs`
* `crates/litchi-xldm/src/xml_attributes.rs`
* `crates/xml-minifier/src/xml_attributes.rs`

`crates/litchi-formula/src/omml/xml_attributes.rs` is intentionally excluded.
It implements the separate lenient `first_wins` helper rather than the
checked-attribute substrate, and the canonical-copy oracle does not include
it.

The candidate keeps quick-xml only as the lexical parser by constructing its
iterator with duplicate checks disabled. `SeenNames` records borrowed raw
name slices in a four-entry fixed array. The fifth distinct name spills those
entries into a vector with capacity eight, retaining the existing vector-like
growth through 32 names. The 33rd distinct name moves the vector into the
existing ordered `BTreeMap<Name, usize>` fallback. Thus the candidate removes
the first small-tag heap allocation without moving the 9–32-name allocation
shape to an early tree, and the ordered path still supplies the bounded
`O(n log n)` hostile-input behavior.

The iterator starts `next_start` at `tag.name().len()`. Each successful
attribute advances it from the borrowed value span. On an unchecked lexical
error, the existing `duplicate_before_value` recovery examines the name at
that offset before returning the lexical error. This preserves quick-xml's
returned duplicate-before-malformed-value precedence, including unquoted,
missing, and unclosed values. Because the duplicate check is disabled inside
quick-xml, the candidate may scan a long or unterminated duplicate value
before emitting that same duplicate error; the output and positions match,
but the scan work is not identical to quick-xml's checked fail-fast path.
`ExpectedEq` remains a name/syntax error because quick-xml has not reached the
duplicate check at that point. Errors and end-of-input transition to `Done`,
preserving fail-fast and fused behavior. The inline and spill states derive
`Clone` with the existing iterator, so a clone at each state observes the same
remaining sequence and an exhausted or failed clone stays exhausted.

The unchecked parser scans each attribute value once, and the recovery name
scan visits only the whitespace and name at the current offset. For a long or
unterminated duplicate, that value scan occurs before duplicate reporting, so
the candidate can do more fail-fast work than the baseline; it adds no second
value scan and leaves the tag-byte and hostile-tag asymptotic bound unchanged.
Duplicate comparison is linear only within the fixed prefix and the at-most-
32-name spill; the ordered phase remains logarithmic per name.

The candidate does not alter callers, parser limits, namespace handling,
relationship projection, decoded-byte accounting, or any public API. It uses
no unsafe code, copied key bytes, hash table, retained cache, or format-
specific shortcut. `Name` implements `Borrow<[u8]>` only so the ordered map
can look up the malformed-error name without allocating or changing its byte
ordering.

Each after-source contains three focused inline tests, copied identically into
the five owners while leaving the existing external `xml_attributes/tests.rs`
oracle byte-identical:

* differential output at 0, 1, 2, 3, 4, 5, 8, 16, 31, 32, 33, and 64 names,
  including adjacent and whitespace-separated forms;
* duplicate and malformed-value precedence at the inline, spill, and ordered
  transitions, malformed names, whitespace, and UTF-8 names/values; and
* iterator clones before each transition, after normal exhaustion, and after
  the first error.

The archived `before/` files are the source state at head `008bea923a`; the
`after/` files are the candidate sources, formatted with rustfmt using
`skip_children=true` because their external test module path is relative to
the production crate. `manifest.json` records the flat before/after archive
mapping, every source hash, and the patch hash
`a234f9c0f13c81a9acc9bb1087d212e1fdba5a1b40d6331367c9112ee289e62d`.
`candidate.patch` is a unified patch labeled with the five production paths.
The canonical-copy normalization (the existing test's
`use std::borrow::Cow;` anchor, visibility normalization, and external-test
path removal) is byte-identical across all five after-sources.

The relevant accepted constraints are ADR 0002's downward crate topology,
ADR 0005's requirement for representative measured evidence before an
optimization claim, ADR 0006's fail-closed validation and malformed-input
preservation, ADR 0010/0011's ownership boundaries, and ADR 0032's limits on
retained derived state. The candidate is therefore ready for the root agent
to apply in an isolated build and run the predeclared differential, quality,
allocation, and native gates. It has no adoption or speedup conclusion until
those gates terminate.

Before application, root refined the long-value test to avoid shadowing its
`tag` constructor and to compare quoted, unquoted, unterminated, and
whitespace-only duplicate values at 1/4/5/32/33 prior names. The production
algorithm is unchanged by this test refinement; the final manifest and patch
bind the refined archive.

The first applied candidate failed the all-feature check because its OLE common
copy had inadvertently made the originally public trait and iterator
crate-private. Root restored exactly those two public visibilities before any
candidate timing. The failed source, patch, and compiler receipts are retained
under `candidate/compile-failure-1`, `quality-1`, and `candidate-fix.json`.
Canonical-body normalization alone does not prove each owner's export surface.

Quality attempt2 caught test-local `tag` constructor shadowing (E0618). Root retained the exact failed archive and logs and renamed the two local bindings in each copy before candidate performance measurements.
