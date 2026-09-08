# Next implementation from current measured evidence

Test borrowing events in `crates/litchi-xlsx/src/raw/compact.rs::changed`
instead of unconditionally calling `into_owned()`. Treat it as a short
measured experiment: the explicit ownership marker is only 5.03% of the
compaction subtree, and other inlined allocation effects remain unquantified.
Do not let an immaterial micro-optimization displace larger pass-reuse work.
The fresh profile places
compaction at 13.27% of inclusive exact commit-context weight. Retain its full
XML pass, start-tag normalization, namespace handling, inherited xml:space,
semantic text, CDATA, comments, declarations, PI, general references and error
mapping. Do not combine this with broad compaction/snapshot fusion.

Add differential tests against the existing ownership path for valid mixed
markup and malformed inputs, including preservation-sensitive whitespace,
qualified attributes, entities, duplicate attributes and malformed tails.
Read and comply with the unchanged accepted ADRs before editing production.

If borrowing alone shows useful representative benefit, measure normal ABBA
one-cell and one-percent save on tiny/medium/dense shapes using matched clean
build paths, plus whole-process or explicitly operation-scoped allocation
observations and full guard review. Qualify only the accepted scenario/metric
scope; do not infer improvement from fewer source-level allocations alone.

The broader measured leads remain eager parsing, snapshot scanning and the
web-publishing metadata pass. Source review records their error-order and
ownership risks. The Store handoff remains bounded at 4,096 cells / 1 MiB;
no full Store retention increase, skipped publication audit, blanket XML
validation removal or redundant pre-sort scan is authorized by this profile.
