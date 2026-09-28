# 0799 checked-first-two-attribute candidate

This directory is a source-only candidate archive based on commit
`049e242195`. It contains no production edit. The five helper owners are
archived because the workspace keeps the canonical implementation in
`litchi-opc` and exact source copies in the four crates that cannot depend on
it.

The candidate reads the first two attributes through quick-xml's iterator with
duplicate checks disabled. After the first successful item it keeps the raw
offset immediately after that item's borrowed value. On the next call it
mirrors quick-xml's key-prefix parser, including its consumption of the first
non-whitespace byte. If the second key is followed by `=`, it compares that
key with the first key and returns `AttrError::Duplicated` before asking
quick-xml to scan the second value. A repeated second key without `=` remains
`ExpectedEq`, because quick-xml reports that lexical error before it performs
duplicate checking. Nonduplicate second items use the unchecked iterator, so
their value and recovery errors retain quick-xml's behavior.

If the second successful value is followed only by XML whitespace, the
iterator becomes `Done` without creating quick-xml's duplicate-name state. If
there may be a third item, a cold transition creates a fresh checked iterator,
consumes the first two items as internal seeds, and resumes at
`QuickXml(2)`. The existing first-32 quick-xml phase, 33rd-item handoff, and
ordered-map bounded check are unchanged after this transition. No item is
returned twice.

The `Second(usize)` phase carries only the raw offset already needed for the
next parser position. `ReplayFirstTwo` is payload-free. Thus the largest phase
payload remains one `usize`/`Box`, and the candidate is intended to retain the
baseline iterator layout; the root-owned direct size probe must measure this
instead of assuming it. The candidate adds no `CheckedAttributes` field,
dependency, unsafe code, ambient state, or public API.

The source archive adds differential coverage for the two-to-three handoff,
clones at every short-prefix boundary, empty values, XML whitespace around
names and equals signs, unusual names whose first byte is `=`, repeated names
without an equals sign, and long malformed values on a repeated second name.
The existing canonical random, duplicate-position, malformed-tail, recovery,
fused, and bounded-comparison tests remain in the shared test copy.

The candidate deliberately trades two short-tag costs: it scans the tail after
each successful prefix item and, for tags with three or more items, reparses
the first two items to seed quick-xml. The second duplicate path avoids
parsing a hostile value before the duplicate error. Long tags keep the
existing bounded handoff and ordered-map behavior. Root owns compilation,
quality gates, direct layout and semantic preflight, workload captures,
resource measurements, and disposition; this archive makes no adoption claim.
