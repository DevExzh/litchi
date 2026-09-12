# Next bounded OLE2 investigation

Keep the 0534 paired role/FAT prefix loop rejected. Its measured physical
exclusive Ir reduction did not translate into native improvement, and all
eight primary XLS p50 rows regressed. Do not revive that loop rewrite, the
rejected visited-bit fusion, or the rejected freshness-session mechanism.

The next bounded investigation should inspect the existing
`SectorChainScratch::collect_exact` path inside
`validate_stream_allocations`. It is the largest named exclusive owner in
the retained profiles: 5,601,140 Ir in each XLS-owned profile and 5,571,945
Ir in each CFB few-large profile, while the parent totals are 10.99–11.32M
and 9.93–10.26M Ir respectively. Its self Ir and 2,730 aggregate calls are
unchanged between 0534 stages. The candidate's unchanged stream and chain
rows, together with the physical-loop rejection, make this a better bounded
place to inspect than another physical-layout spelling change.

Start with source and final-binary code-shape review of the already existing
collector: `reset`, visited-map preparation and clearing, checked sector
lookup, `contains`/`insert`, exact vector push, and end-marker branches. The
review should identify which work repeats across the root, MiniFAT and
regular-stream chains and whether any safe layout opportunity exists. This
is an investigation of the current collector, not a proposal to retain
visited state across independent chains. Preserve collect-then-claim order,
cycle and bounds checks, marker/error precedence, fallible capacity labels,
and failure reset behavior. Do not merge collection with ownership
publication, add a freshness cache, or change the allocation vectors.

No runtime patch or new capture is part of this handoff. If the code-shape
review produces one narrow candidate later, bind its exact source and final
binary, retain positive incoming timed-constructor versus CFB-setup
classification, and rerun the complete two-repeat ABBA protocol before any
adoption decision. The same four primary p50 gates, parent/exclusive
instruction rows, allocation guards, RSS review, correctness matrix and
actual final-binary inspection remain required. A local Ir reduction or a
smaller symbol is not sufficient evidence of a speedup.

Keep OLE2 and OOXML active. Defer ODF until that optimization goal completes;
iWork remains excluded.
