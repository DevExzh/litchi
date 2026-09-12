# Rejection and restored-source review

The original candidate passed its four native primary rows. Correcting test expectations and a pre-existing test lint led to a final release rebuild with the same runtime source but a different binary hash and ELF text-section size. The build difference is observed; its cause is not established. The initial binary was not substituted for the rebuilt binary.

The complete final native matrix was frozen before capture: 44 children and 2,680 measured samples in baseline–final–final–baseline order, with the original primary and guard counts and thresholds. All runtime oracles passed. Dense-sparse repeat 2 achieved only 0.8936597% total p50 reduction and 0.6056070% total mean reduction, below both 1% requirements. This fails admission even though the other three rows and all four planning instruction comparisons improve sufficiently. No additional timing repetitions are used to rescue this candidate.

The production codec is restored exactly to the baseline blob. The only retained Rust changes are nine public MCE regression tests and the equivalent `repeat_n` call inside an existing XLSX test helper. The corrected tests also passed separately against unchanged baseline production. The restored source receives the same eight quality commands; its receipts live under `restored/`, separate from the two measured candidate source states.

All original and final timing flags remain visible. Initial allocation, eager-read, reopen-confirmation, and hardware evidence concerns the original candidate and carries no retained optimization claim. Final profiles are diagnostic; lower planning instructions do not override the failed end-to-end gate. No broad OOXML, cold-source, scaling, memory, or OLE2 speedup is claimed.

The next performance work remains OLE2/OOXML. Fresh CFB owner attribution is proposed in `next-ole2-review.md`; ODF stays deferred and iWork is excluded. This rejected batch advances the evidence and regression coverage but does not complete the optimization goal.
