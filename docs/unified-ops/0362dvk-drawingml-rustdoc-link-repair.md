# DrawingML rustdoc link repair

The fresh all-68 capture at `29f072f19c53503395164fcb3a618c44fc22cf75` stops after nine successful packages. `litchi-drawingml` rustdoc exits 101 because the new `LimitKind` documentation links `Error`, which names both the error enum and the imported derive macro. The workspace denies broken intra-doc links.

The link now explicitly selects `enum@Error`. This changes documentation only. The original capture remains incomplete and immutable; no inventory or dispositions were generated from its nine JSONs. [Failure receipt](results/0362dvk/capture-failure.json) preserves the command, source, 4100 lock and external log identities. Peak RSS was 649,797,632 bytes and the target ended at 1,509,922,247 bytes, so this was a documentation defect rather than memory exhaustion.

Focused rustdoc at `642f2e45f6bda1b642890128408cc243e29aa953` passes with all features, hidden items, JSON output, `--offline --locked`, one Cargo job and the 20G cap. The [receipt](results/0362dvk/focused-rustdoc.json) binds the command, 4100 lock and external log; elapsed time is 20.671 seconds and the existing target ends at 6,083,810,704 bytes. A renewed full source-bound capture remains required. Nothing is pushed.
