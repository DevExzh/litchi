# 0911 v2 first-child admission failure

The v2 first child is retained as a failed attribution attempt. Its structured
binding and raw observations are in
[`first-child-failure-v2.json`](first-child-failure-v2.json). No remaining v2
legs were admitted from this child.

The v2 command used the retained 0910 binary
`f4ad7ee158061cc80354c3c8674c9fe6a8100ed76a0c8f743653efcb06bb2e8d`, CPU 12,
the materialized route, zero warmups, and one sample. The workload completed
its independent semantic, source, output, and sink checks. The Callgrind
options changed the reset and dump boundary to `timed_iteration` while
retaining the exact transaction method as the collection toggle.

The parsed positive raw part is triggered after `timed_iteration` and has
summary and totals **1,127,351,444 Ir**. The exact owner record conserves its
self cost (**106 Ir**) plus immediate direct-child inclusive cost
(**1,127,351,338 Ir**), and all function self costs reconstruct the summary.
The terminal part is zero. These grammar and conservation checks pass.

The owner nevertheless has two positive incoming edges in the same raw graph:

| caller | calls | inclusive Ir | interpretation |
| --- | ---: | ---: | --- |
| `litchi_perf_baseline::pptx_owned_benchmark::timed_iteration` | 1 | 1,000,676,539 | required measured edge |
| `litchi_perf_baseline::build_pptx_source_edit_corpus` | 1 | 126,674,905 | setup owner edge retained in the dump |

The v2 protocol requires exactly one positive owner edge from
`timed_iteration`. The raw evidence therefore rejects this child. The setup
edge is retained as a separate observation; it is not subtracted from the
summary, and the measured edge is not promoted by selecting the larger or
later-looking row. The failure does not assert why Callgrind retained that
edge beyond the raw graph fact.

The next v3 protocol explicitly admits the two known positive contexts by
their exact incoming callers and reports setup and measured instruction costs
as separate, non-additive rows. This v2 receipt remains unchanged as an audit
record.
