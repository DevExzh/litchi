# 0733 independent finish attribution review

Status: pass for retrospective instruction attribution, with the scope limits
described below. This review is read-only with respect to production code and
the two ancestor packets.

I independently reparsed the three retained
`change-0731/captures/callgrind-{0,1,2}.callgrind` files and compared the
result with the 0731 analysis and the 0733 replay. The profile totals are
59,578,167, 59,578,131, and 59,578,178 Ir. The embedded-finish incoming edge
is 3,087,582, 3,087,556, and 3,087,601 Ir respectively. The selected edge
costs are invariant across all three profiles:

| edge | collected Ir | denominator |
| --- | ---: | --- |
| `finish -> write_package` | 1,970,376 | finish inclusive Ir |
| `finish -> validate_rewrite` | 769,284 | finish inclusive Ir |
| `finish -> append_checked` | 312,235 | finish inclusive Ir |
| `write_package -> OleWriter::write_to` | 1,434,432 | write_package inclusive Ir |
| `write_package -> OleWriter::create_stream` | 383,349 | write_package inclusive Ir |
| `write_package -> adopt_source_layout` | 147,056 | write_package inclusive Ir |
| `write_to -> ReusePlan::validate` | 731,935 | write_to inclusive Ir |
| `write_to -> ReusePlan::emit` | 520,891 | write_to inclusive Ir |
| `write_to -> plan_sector_layout` | 180,562 | write_to inclusive Ir |
| `validate_rewrite -> open_stream` | 671,509 | validate_rewrite inclusive Ir |

These are inclusive costs of direct edges. A nested row is already contained
in its parent row and must not be added to the parent again. For example,
`ReusePlan::validate` and `ReusePlan::emit` are parts of the
1,434,432-Ir `write_to` edge; the three values must not be added to the
1,970,376-Ir `write_package` edge as extra work.

The seven replayed nodes (`finish`, `write_package`, `validate_rewrite`,
`write_to`, `create_stream`, `ReusePlan::validate`, and `ReusePlan::emit`) each
have one positive-cost incoming caller. For every one, incoming inclusive Ir
equals its self Ir plus all positive outgoing edge Ir. This is sufficient for
the reported subtree boundaries. Generic descendants such as
`OleFile::open_stream`, `memcpy`, allocation helpers, and `finish_grow` have
multiple positive callers in the complete profile. Their edge costs may be
shown in the caller context, but they must not be promoted to a globally owned
finish or writer fraction. In particular, the 671,509-Ir validation
`open_stream` edge and the 593,163-Ir reuse-plan `open_stream` edge are separate
contexts; combining either with global `open_stream` totals would double-count
or misattribute work.

The three `calls=3` fields on finish edges are not a per-owner invocation
denominator. The measured public wrapper has one owner-to-public-edit edge,
while the standalone safe-Rust witness records one owner-to-work call and
three work-to-leaf calls (5,009 and 5,008 Ir on those edges, with a 5,010-Ir
owner total) across all three fixed runs. The retained PPT profiles therefore
support process-level collected Ir partitions only. They do not establish
three measured finishes, per-call Ir, or native latency fractions.

The source bridge and report are consistent with this reading. The 0732 source
census differs from 0731 only in the diagnostic Cargo feature and
`slide_order.rs`; the archived ordinary `Transaction::commit` method is byte
identical. The nominated `create_stream` handoff is a plausible ownership
investigation because the current finish path passes borrowed stream slices to
an existing copying API while the writer also exposes `create_stream_owned`.
The retained evidence supports the candidate as a measurement target, not as
a claimed optimization. A future pilot must preserve source-layout adoption,
stream topology, Reuse-plan validation, final rewrite validation, limits,
failure ordering, semantic reopen, payload checks, and patch/digest behavior.

I found no attribution or source-custody defect in the final 0733 analyzer,
report, or ten negative parser/context/custody controls. The broader goal
remains active; this batch is ready for packet sealing after the root-owned
cleanup and final census.
