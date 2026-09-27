# 0793 — source review for the small XML duplicate-state experiment

This is a bounded source and return-on-investment review for the next OOXML/OLE2
performance action. It does not modify production code, build a candidate, run
native measurements, or claim a speedup. The only artifact in this review is
the source-backed experiment design.

## Scope and inherited evidence

The review starts from the current branch head `1d17b5b0e7`, which retains the
0792 exact-empty-tail change and its namespace test correction. The accepted
ADR index and the 35 hash-bound goal, taxonomy, and ADR inputs recorded by
0792 are treated as unchanged. The relevant contracts are ADR 0002's
dependency direction, ADR 0005's measured resource evidence and bounded
scratch, ADR 0006's preservation and malformed-input defenses, ADR 0008's
verification gates, ADR 0010/0011's archive and OPC ownership, and ADR 0032's
limits on retained derived state. No contract authorizes weakening duplicate
attribute validation or the existing XML limits.

The retained evidence establishes a useful source seam:

* The 0791 large Callgrind region entered `inspect_element` 181,678 times and
  constructed the checked attribute iterator 273,263 times. The checked
  iterator had 13,856,850 self instructions and 48,295,031 instructions
  inclusive from the inspector. The nested quick-xml `IterState` path was also
  material, at 38,981,486 self instructions in the independent function
  census. These are one-process guest instruction counts, not removable
  fractions or native phase shares.
* The 0791 frame-pointer samples saw the checked iterator under the exact
  capture owner 186 and 148 times in the two repeats, and the quick-xml
  iterator 134 and 120 times. These are observed stack counts with unresolved
  interior frames; they are not population fractions.
* The retained 0792 exact-empty branch improved the generated large capture by
  5.211% and lifecycle by 6.937% under its frozen paired-p50 rule. Its large
  capture allocation lane remained exactly 72,106 allocation calls and
  4,767,939 allocated bytes in both legs, with identical net-live and
  peak-above-entry values. That result is consistent with the branch avoiding
  iterator construction for empty tails: quick-xml's empty iterator owns a
  zero-capacity `Vec` and has not allocated any key storage yet.

The allocation equality does not rule out the next hypothesis. It establishes
no whole-operation allocation reduction for the exact-empty change; the source
also shows why its skipped iterator had no key allocation to remove. The
candidate below targets nonempty tags, where duplicate bookkeeping records the
first key and can grow a heap buffer.

## Decision

A small, shared duplicate-state experiment is warranted, but only as a fresh
source-bound candidate with a narrow implementation and a direct allocation
lane. The source evidence is stronger than the presentation-index alternative:
the per-slide checked iterator has hundreds of thousands of calls and a large
inclusive instruction cost, while the presentation duplicate scan in
`notes/package.rs` appears under the owner only 8 and 5 times in the two
0791 profile repeats. The latter remains a possible later investigation, but
it is lower ROI and has no independent allocation attribution.

The experiment belongs at the existing XML attribute substrate. The canonical
implementation is `litchi-opc::xml_attributes`; `litchi-ooxml-common` reexports
it to the OOXML format owners, and dependency-isolated copies exist in
`litchi-ole-common`, `litchi-sign`, `litchi-xldm`, and `xml-minifier`. A claim
about a shared substrate must either update and test all of those copies or be
explicitly scoped to the canonical OPC/PPTX owner. A PPTX-only patch that
silently leaves the copies divergent would fail the existing canonical-copy
test and would not be a shared optimization.

## Source seam and allocation hypothesis

### The PPTX caller

`crates/litchi-pptx/src/notes/codec.rs:447` performs the existing element
namespace resolution, local-name UTF-8 validation, and root validation before
the 0792 exact-empty check. Attribute-bearing and whitespace-tail elements
then enter `element.checked_attributes()` at line 477. Every item is still
checked for malformed syntax, namespace resolution, UTF-8, entity unescaping,
relationship projection, and the notes attribute and decoded-byte limits.

The 0792 branch is therefore already the correct boundary for an empty tail.
The next candidate must not move that branch, bypass those checks, or turn the
notes caller into a lenient reader. It should change only the bookkeeping that
`CheckedAttributes` uses for a nonempty tag.

### What quick-xml 0.41 allocates

`Cargo.lock` pins quick-xml 0.41.0. Its `Attributes` iterator contains an
`IterState` with these relevant fields:

* `state`, `html`, and `check_duplicates` are scalar parser state;
* `keys: Vec<Range<usize>>` stores every previously seen raw key so the
  duplicate error can report the first position; and
* `key_hashes: Option<HashSet<u64, ...>>` is created only after the 32-name
  threshold in the upstream iterator.

`IterState::new` uses `Vec::new`, so creating an iterator or advancing it over
an empty tail does not allocate key storage. On the first valid attribute,
`check_for_duplicates` compares against the prior ranges and pushes the new
range. The zero-capacity vector consequently grows on the first attribute and
may grow again as a tag approaches 32 attributes. The upstream threshold is
32; above it quick-xml seeds a hash prefilter and then scans earlier names on a
hash hit.

The Litchi `CheckedAttributes` wrapper deliberately intercepts before that
upstream hashed phase. Its `next` method delegates the first 32 results to
quick-xml. When the 33rd result is requested, it disables quick-xml checks,
re-reads the first 32 names through `unchecked_attributes()`, and constructs
the existing `BTreeMap<Name, usize>` in `OwnCheck`. This preserves the bounded
`O(n log n)` fallback, but the ordinary one-to-32 path still inherits
quick-xml's heap-backed `keys` vector.

The source-backed hypothesis is therefore:

> For nonempty tags with a small number of attributes, the duplicate-check
> `Vec<Range<usize>>` is a repeated temporary allocation that is not visible
> in the exact-empty 0792 allocation result. Keeping a bounded prefix of key
> ranges inline, while retaining the current ordered fallback, can remove or
> defer that allocation without changing the parsed bytes.

This is an allocation hypothesis, not an allocation claim. The 72,106 whole
operation calls include package construction, semantic values, relationship
vectors, strings, and output verification. They cannot identify this vector's
share.

### Candidate shape

The safe design to test is an internal checked iterator that obtains lexical
items from `unchecked_attributes()` and performs the same fail-fast duplicate
check in an inline name store. The inline store should hold raw key ranges and
their first positions, not copied names. Once its fixed prefix is full, it
should spill to the existing ordered `BTreeMap<Name, usize>`; the map remains
the path for tags above the bounded prefix and retains the current worst-case
bound.

The prefix capacity must be measured rather than assumed. On a 64-bit host a
`Range<usize>` is 16 bytes, so an inline capacity of 32 consumes about 512
bytes before iterator and enum state. That could save every small-tag heap
allocation while increasing stack frames and cache traffic. Capacities 4, 8,
and 16 are plausible bounded controls; 32 is a diagnostic upper control, not
an automatic production choice. A `SmallVec`-style implementation is safe
Rust and avoids `unsafe` `MaybeUninit` storage, but it requires synchronized
dependency updates in the copy modules. A hand-written array with unsafe
initialization is out of scope because the workspace forbids unsafe code.

If a capacity below 32 is selected, the spill must remain a vector-like
temporary until the existing 33-name takeover, or the experiment must disclose
that 9–32-name tags now allocate a tree earlier. Converting to a B-tree at 8
would risk more allocations than the current one vector. The first experiment
should therefore compare the actual spill shape, not infer it from the
container's nominal capacity.

An alternative one-attribute peek-and-replay design could avoid allocation for
the single-attribute case, but it parses a multi-attribute tag once and then
replays it through the checked iterator. That extra scan makes it a separate
candidate. It should not be mixed into the inline-state result because a
public latency change would then have no single source explanation.

## Semantic and security constraints

The shared iterator is a fail-fast checked path. A candidate must satisfy all
of these constraints before any timing result is considered:

1. `AttrError` variants and positions must match the current checked iterator
   at the first error. In particular, a duplicate name is detected after the
   name and `=` have been read but before its value is accepted. A duplicate
   followed by an unquoted value, missing value, or missing closing quote must
   still become `AttrError::Duplicated` through the existing
   `duplicate_before_value` rule. `ExpectedEq` remains a syntax error because
   quick-xml has not reached the duplicate check at that point.
2. The iterator must stop after the first error exactly as the current
   `CheckedAttributes` implementation does. Its post-error recovery behavior
   cannot be used as an excuse to change the lenient `first_wins` iterator,
   which intentionally reads on and has a different contract.
3. Raw QName duplicate checking remains byte-exact and separate from
   namespace-expanded duplicate checks performed by individual format owners.
   The candidate must not replace the ordered fallback with an unkeyed hash
   table on untrusted input. The fixed inline prefix may do a bounded linear
   byte comparison; the existing ordered map must continue to bound larger
   tags.
4. The candidate must retain lexical malformed-name/value behavior, entity
   errors, namespace declarations, and borrowed key/value lifetimes. No raw
   slice may outlive its `BytesStart` event, and no value may be copied merely
   to implement duplicate detection.
5. Per-element attribute, decoded-byte, XML-node, depth, decompression, and
   aggregate resource limits remain at their current callers. The iterator
   optimization cannot count less, defer a limit, or turn a typed refusal into
   a partial result. The notes scanner's existing exact-empty limit tests and
   buffered oracle remain authoritative.
6. The internal use of `unchecked_attributes()` must be confined to the
   duplicate-checking substrate. Call sites must continue to use
   `BytesStartExt::checked_attributes()` or their existing explicitly lenient
   helper, so the workspace's `clippy::disallowed_methods` security boundary
   remains meaningful.
7. The canonical OPC module and every dependency-isolated source copy must
   compile the same focused differential tests. No upward crate dependency,
   public API, cache, retained state, executor, I/O behavior, or publication
   behavior is needed or permitted.

The existing `xml_attributes/tests.rs` already compares well-formed and
malformed cases against quick-xml, including duplicate positions around the
32-name transition. The candidate needs an additional direct allocation-safe
equivalence layer rather than weakening those tests. The independent test
oracle should collect the exact sequence up to the first error, compare the
error discriminant and positions, and assert that the candidate yields no
item after an error.

## Bounded measurement design

No measurement is performed in this source-review turn. A future candidate
packet should use fresh before/after builds from the current head, preserve the
0792 source and fixture identities, and predeclare the following lanes.

### 1. Focused iterator differential

Run the canonical helper and all source copies over prebuilt `BytesStart`
values with 0, 1, 2, 4, 8, 16, 32, 33, and 64 attributes. Include both Start
and Empty event construction, distinct names, duplicates at the first, middle,
and last positions, namespace declarations, long names, whitespace variants,
and malformed values (`a=x`, `a=`, an unclosed quote, and a duplicate whose
value is malformed). Compare with quick-xml 0.41's checked output through its
first error, including positions and error precedence. The existing 0792 notes
oracle should additionally cover root/refusal and node/attribute limit cases;
the candidate must remain byte- and error-identical there.

The test must exercise the >32 ordered fallback and prove that its key
positions are unchanged. It must also exercise a tag containing exactly the
inline capacity and one containing capacity plus one. This catches both the
spill transition and accidental early tree allocation.

### 2. Direct allocation micro-lane

Prebuild the tag bytes and `BytesStart` outside the measured region. In a fresh
process, repeatedly consume only `checked_attributes()` (retaining a black-box
count or digest so the loop cannot be removed) and record the
operation-scoped counting allocator fields already used by 0792:
allocation calls, allocated/deallocated bytes, net live bytes, and peak above
entry. Use the same fixed iteration count for every shape and enough repeated
operations to make one small allocation visible. Keep the zero-attribute case,
the one-to-eight cases, a 16/32 case, and a 33/64 case in separate reports;
pooling them would hide the spill behavior.

This lane identifies whether the candidate actually removes the hypothesized
heap activity. It cannot by itself establish a public workflow gain, and its
allocation totals must not be mixed with the 0792 package-level totals.

### 3. Public PPTX lane

Reuse the 0792 public capture, staged commit, and lifecycle protocol: tiny,
medium, large, ASCII-vendor, and Unicode-vendor controls; six alternating
blocks; 30 measured samples after three warmups on the fixed CPU; and a
separate two-block, three-sample allocation lane. Add a large attribute-heavy
control rather than relying only on the existing vendor case: the retained
vendor fixture is medium (12 × 8, 96 text tags) and carries six attributes per
text tag, while the 100 × 100 large control has no injected vendor attributes.
The extra control should scale the same six-attribute injection to the large
shape and keep the source/output/semantic identity checks.

Use ordinary control and candidate binaries for the latency gate. A profiling
wrapper or a heaptrack process is not an interchangeable latency control. Keep
all spread flags and paired samples, and do not turn the 0792 5.211%/6.937%
result into a cumulative claim.

### 4. Heaptrack ownership attribution

`/usr/bin/heaptrack` and `heaptrack_print` are available on this host and are
feasible for a descriptive attribution lane. Run one fresh direct iterator
process per arm with its input constructed before the target loop, then filter
the decoded stack histogram for:

* `litchi_opc::xml_attributes::CheckedAttributes` and its inline-name helper;
* `quick_xml::events::attributes::IterState::check_for_duplicates` and
  `IterState::next`; and
* the allocator growth frames beneath those owners.

Use a frame-pointer/debug-symbol build if normal release inlining hides the
owner, and retain the raw trace, print output, command, binary identity, and
parser identity. Heaptrack observes the whole process and changes allocator
behavior, so it is source attribution only. It cannot serve as an operation
local allocation gate or a physical peak-RSS claim. The operation-scoped
counting allocator remains the authoritative before/after resource lane.

For the full PPTX process, heaptrack's setup and verification allocations must
be labeled separately from the clocked `opened_presentation` operation. A
standalone iterator lane is the stronger ownership evidence; a whole PPTX
histogram may corroborate it but cannot prove that the 72,106 calls belong to
the duplicate-state vector.

## Disposition gates

The inline-state candidate is worth retaining only if all of the following are
true:

* the differential and repository malformed-input suites preserve exact
  checked errors, limits, and semantic/output identities;
* the direct micro-lane shows the predicted allocation reduction for its
  chosen capacity, with no increase for the zero-attribute or >32 fallback
  controls;
* the public allocation lane has no increase in calls, allocated bytes, net
  live bytes, or peak above entry across the full 0792 matrix; and
* the ordinary public latency lane meets the existing useful-benefit policy
  (at least one capture or lifecycle case improving 3% with its paired-p50
  bootstrap upper bound below one), with no frozen latency rejection.

If only the micro-lane improves, retain the result as diagnostic evidence and
do not claim a production optimization. If a capacity reduces allocations but
slows ordinary one-attribute tags or increases the 9–32 spill path, reject
that capacity and keep the source unchanged. If the canonical helper improves
but the synchronized OLE2/OOXML copies cannot pass their equivalence and
boundary gates, reject the shared adoption rather than leaving a split
substrate.

This action follows the goal's order of removing unnecessary duplicate
bookkeeping and allocation before lower-level tuning. It does not expand CRUD
coverage, cold/range/concurrent evidence, source-backed I/O, or the order of
magnitude program target. OLE2 and OOXML remain active, ODF remains deferred,
and iWork remains outside this work.
