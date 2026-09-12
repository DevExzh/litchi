# 0528: prioritize the shared XML auditor from XLSX publication profiles

Fresh source-bound profiles identify OPC overlay XML validation as the largest
instruction owner inside XLSX publication. The runtime source is unchanged;
this batch adds reproducible attribution and selects the next investigation.
It does not adopt a production optimization or claim a speedup.

The prior scanner lead was first bounded against retained 0525 profiles.
The entire scanner-to-namespace-resolver edge is 31,526,144 Ir, 4.0274% of
commit and 7.5613% of scanner instructions. Start/Empty lookups remain required,
so only part of that edge could be removed. It is not an End-only measurement
or a bound on every possible code-generation effect. The rejected 0522 scanner
fusion and 0527 row arena remain rejected.

The 0527 baseline's 200-sample primary rows put publication at roughly
30–32% of total edit/save time, planning at roughly 32%, and commit at roughly
36%. These shares use sums of matched per-sample phases, never sums of medians;
excluded reopen time is not included. They motivate profiling another phase,
not converting instruction fractions directly into native speedups.

Four fresh Callgrind children (medium/dense-sparse, two repeats) isolate
`SourceBackedEditor::publish_multi_commit_to_stream`. Each retains five
lifecycle dumps, one uniquely identified measured dump and the zero-cost
termination dump. All output/lifecycle oracles pass. The method excludes
destruction of the returned snapshot, unlike the native publication timer.

| Measured owner | Aggregate Ir | Share of publication Ir |
| --- | ---: | ---: |
| Publication method | 535,830,169 | 100% |
| OPC overlay XML validation | 301,274,428 | 56.2257% |
| Physical preservation writer | 231,144,888 | 43.1377% |
| Attribute inspection, nested within XML validation | 62,205,865 | 11.6093% |

XML validation ranges from 53.8232% to 57.5571% across individual captures.
The nested attribute cost must not be added to the validation total. Every
reported decomposition is checked against raw edges and self-plus-child
annotation equations. The physical writer includes regenerated-entry
preparation and preserved-byte transfer, not just compression.

Next, investigate a resettable duplicate-attribute checker within the shared
authored-XML auditor. Both slice and streaming paths currently create fresh
checked attribute iterators; there is no existing reusable duplicate checker.
A narrower reviewed design probes at most two unchecked results and uses a
fast path only for zero/one valid attribute; any error or second result
replays the original checked path before audit-state mutation. The profile establishes the owner cost,
not a removable allocation count. A fresh native/allocator candidate campaign,
malformed-input differential coverage and exact error/limit preservation are
required before retention. Both original and replacement OPC audits must stay;
XLSX semantic validation is not a substitute for their different contract.
Planning remains a separate sizeable opportunity without a new profile here.

The complete source manifest matches the 0527 final source whose 12 quality
gates and 1,297 test executions passed. The fresh release build and four profile
oracles pass; two new analyzer regression tests cover bracketed Rust symbols.
See the [evidence bundle](../results/change-0528/README.md) for raw reports,
receipts, source/dependency bindings, replay, cleanup and sealing. OLE2/OOXML
optimization remains active. ODF is deferred and iWork is excluded.
