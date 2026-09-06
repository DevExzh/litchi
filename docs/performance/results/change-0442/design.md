# 0442: Share auxiliary staging XML traversal

The sealed 0441 candidate profile attributes 5.860% / 9.176% / 7.060% of
weighted sampled periods to settings / declarations / page metadata under
transaction. Source-fragment scanning is a further 7.360%; unclassified
transaction work is 0.922%. These are observed call-chain groups including
warmups, with incomplete symbolization, not isolated phase wall times.

All three auxiliary readers use the same slice input and trim_text(false).
Extract their unchanged event state machines, then drive them with one shared
namespace-aware event traversal. Keep the individual public parse functions.
Retain their original complete loops as test-only independent reference parsers.

Historical priority is complete settings parse/finish, then declarations,
then page metadata. Run settings first on each shared event and immediately
return its errors; defer lower-priority errors while continuing settings.
Finish in historical order. XML tokenization failures belong to settings,
which historically traverses the complete input first. Preserve each parser's
own byte gate, state transitions, lazy attributes, limits, final validation,
exact diagnostics and owned outputs. Lower-priority output may be accumulated
before a later settings failure; finite existing input/model limits still
bound this additional error-path work. Do not claim invalid-input performance.

Freeze A1/B1/B2/A2 before source edits: 24 reports, 720 samples, normal/allocator,
64/4096/8192 slides, CPU 2, one worker, 30 samples and three warmups. Require a
5% medium/large gain in both repeats for normal p50 OR calls/requested bytes/
peak above entry. Peak is explicitly included up front. Review every >5%
adverse normal/RSS/allocator metric and every >5% repeat change. Four fresh
whole-process profiles follow. Root CPU jobs are serialized; switch source
only between confirmed terminal jobs. Preserve all attempts and frozen drivers.

The accepted ADR tree remains c950b6c8be822561b498d7bbe87c460873dcbf49 from the
prior complete read. ADR 0002/0023/0024 keep grammar in ODP; 0003 preserves
isolated edits and exact source authority; 0005 requires measured benefit and
bounded ownership; 0006/0008 require fail-closed diagnostics, preservation and
publication readback. No public API, dependency, unsafe code, executor or
ambient I/O is added. The full non-iWork goal remains active.
