# 0555 profile consumer contract

`analyze_profiles.py` is a read-only consumer for the matched OLE2 physical-marker
Callgrind lane. It reuses the immutable raw parser retained by change 0554 and
binds every child to the frozen 0555 plan, output-stage source manifest, live
execution-stage source manifest, workspace-lock copy, driver, binary, command,
and receipt artifact hashes.

The profile matrix has two repeats of one owned-source XLS constructor and one
CFB constructor for each of `tiny`, `many-small`, and `few-large`. Every XLS
job must retain five positive timed dumps; every CFB job must retain one setup
dump plus five positive timed dumps. The role of each dump comes from the
selected owner's positive incoming edge and positive ancestry to either the
benchmark runner or an allowed CFB setup caller. The dump suffix is checked for
identity and continuity only. Baseline repeat two is output in the baseline
folder but must carry candidate execution-stage identity.

Each timed dump retains selected-owner inclusive, self, direct, and call
counts. It also retains separate attribution for `load_fat`, `claim_sector`,
`validate_stream_allocations`, `collect_exact`, and
`validate_physical_sector_layout`. Inclusive values are never added to self or
to another target. If a target is absent, inlined, or has no positive incoming
edge, the report records a nullable value and
`indeterminate_assembly_required`; it never substitutes zero. For an inline
`claim_sector`, possible caller context is retained separately and is never
counted as `claim_sector` work.

The comparison exposes all moved-work rows and reports two mechanism checks:
the total XLS owner inclusive Ir must decrease in both repeats, and the
physical reconciliation leaf self Ir in the targeted XLS-owned constructor
lane must decrease in both repeats only when the leaf is positively attributed
in both stages. CFB shape rows remain retained controls and moved-work
diagnostics. Otherwise the leaf check is explicitly indeterminate pending
same-binary assembly mapping. These
profiles do not authorize a native, allocation, physical-I/O, or production
adoption claim.

The CLI may write only the three canonical reports:

* `baseline/profile-analysis.json`;
* `candidate/profile-analysis.json`; and
* `profile-comparison.json`.

An existing canonical report must be byte-identical on replay. No annotation,
temporary, or alternate output is created by this consumer.
