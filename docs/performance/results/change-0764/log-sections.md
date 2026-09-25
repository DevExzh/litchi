# Log sections for change 0764

Ready-to-paste paragraphs for the coordinator, one per shared log. This change
does not edit `HOTSPOTS.md`, `REPORT.md` or `GOAL_AUDIT.md` itself; the numbers
are the record's.

---

## For `HOTSPOTS.md`

## 0764 — Hostile attribute lists no longer cost quadratic time

[0764](0764-xml-attribute-dos-hardening.md) bounds the work of start tags with
many attributes, repeated names or many namespace declarations. quick-xml keeps
iterating after reporting a duplicate and scans the tag's earlier names for
each one, so readers that skip attribute errors (docx styles, run properties,
images, bookmarks; pptx backgrounds, custom shows, handout; drawingml chart
`get_attr`; OMML; 40 sites) paid quadratic time on a tag of repeated
names; they now read through `xml::attributes::first_wins` (first occurrence
wins, `O(n log n)`), which also stops quick-xml's recovery defect from reading
names inside a duplicate's value. DOM readers that checked each attribute
against a list (xlsx, xldm, docx, pptx, drawingml, OPC) use ordered or keyed
sets. The MCE processor and stream resolve prefixes through an index after a
walk of at most 64 declarations, build hoisted declarations once per frame and
match directive targets by lookup. The audit, the MCE processor and the stream
refuse a tag over 1,024 attributes with a typed limit. On hostile inputs: a
DOCX style tag of 40,000 + 40,000 repeated names 2,174.7 → 8.95 ms; a
shadowed MCE declaration chain 1,408 → 22.5 ms; MCE hoisting 172.9 → 15.8 ms.
Benign real parts: MCE −1.3% to −2.3% instructions. Left: quick-xml's unkeyed
large-tag pre-filter at the ~560 fail-fast sites outside the audit and MCE
(bounded only by tag size), per-event resolver clones and per-node binding
clones (namespace-scope costs). After an independent review, the MCE
processor gives each in-scope namespace URI an identity (hashed once with a
per-scope key, facts computed at declaration), so name checks no longer read
the URI: the review's 1,000 attributes under a 4 MiB URI 612.3 → 19.3 ms,
1,409.8 → 16.9 ms when that namespace is ignorable; the benign worksheet 14.06
→ 13.06 ms. Also left: the MCE stream still copies and hashes the URI per
attribute (3.72 s on the review's 1.08 MB input), which needs its public event
names to share their namespace (a breaking change).

---

## For `REPORT.md`

## 0764 — Hostile attribute lists no longer cost quadratic time

[0764](0764-xml-attribute-dos-hardening.md) is a security and correctness
change (`performance_claim: none`). Hostile inputs, before → after, ABBA on core
16: a DOCX style tag of 40,000 distinct names and 40,000 repeats read by the
style reader 2,174.7 → 8.95 ms (ratio 0.0041); an MCE part with 32 nested
elements re-declaring 1,000 prefixes around 50 attribute-heavy elements 1,408.3
→ 22.5 ms (0.016); hoisting declarations onto 2,000 children 172.9 → 15.8 ms
(0.092); a 20,000-declaration tag refused in 0.31 ms instead of after 163.5 ms;
a 200,000-attribute tag refused by the source audit at attribute 1,025.
Benign: real Excel and Word parts through MCE 0.976–0.984 in time with fewer
instructions; the audit +0.5% instructions. Two harness controls exceed 5% in
wall time with unchanged timed instructions (`docx_semantic_full_text/medium`
1.055, same instructions; `xlsx_first_cell/tiny` 1.052, +0.55% instructions
from inlining of untouched code). A differential over 12,881 real XML members
and 20,000 generated MCE documents gives byte-identical reports. After an
independent review, namespace identities in the MCE processor make its name
checks independent of the URI's length: the review's input (1,000 attributes
under a 4 MiB namespace URI) 612.3 → 19.3 ms, the same with the namespace
ignorable 1,409.8 → 16.9 ms, the real worksheet 14.06 → 13.06 ms, the
differential again byte-identical; the MCE stream's cost on long URIs is
unchanged (3.72 s on the review's 1.08 MB input) and listed as remaining.

---

## For `GOAL_AUDIT.md`

## 0764 — Hostile attribute lists no longer cost quadratic time

[0764](0764-xml-attribute-dos-hardening.md) strengthens GOAL.md rule 12 (XML
protections) under ADR 0006 and ADR 0005: a typed per-element attribute limit
(`xml_minifier::audit::Resource::ElementAttributes`, default 1,024, ceiling
4,096; `mce::Limits::max_attributes_per_element`, a breaking public field under
0652 trade-off 1), checked as attributes are read; lenient readers keep their
first-occurrence semantics at bounded cost; quadratic list scans replaced by
ordered or keyed sets with the same typed errors; no `unsafe`, no new
dependency, no ambient behaviour; output byte-identical on 12,881 real members
and 20,000 generated MCE documents. New refusals: tags over the limit (none
among 7,529 real members; the largest element has 43 attributes). Reported
residual exposure: quick-xml's unkeyed large-tag duplicate pre-filter at
fail-fast sites the limit does not cover, and quick-xml's duplicate recovery
defect for lenient readers outside the survey (upstream reports drafted in the
record). After an independent review: a hard ceiling of 4,096 attributes per
element for the MCE processor and stream (`mce::ATTRIBUTES_PER_ELEMENT_CEILING`,
ADR 0005), the SVG blip duplicate-declaration refusal restored for an empty
prefix (`xmlns:` beside `xmlns`, now refused in both orders), the complex-field
marker's refusal of a malformed attribute restored, and per-scope namespace
identities in the MCE processor; the MCE stream's per-attribute cost on long
namespace URIs is a further reported residual (bounded by its 1 MiB name
limit).
