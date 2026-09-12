# Measured CFB claim layout

The candidate preserves the source-level ownership checks and moves only error
formatting into private cold functions. The final measured binary is
`d7d935fca5066d647f67a2ef1b18e56fff2162b2f0726b5fd22094fd9910b819`;
its source manifest includes the two new private contract tests. All 14
quality gates passed against that exact source before the release build.

[Assembly analysis](assembly-analysis.json) verifies six baseline out-of-line
claim variants, each 363 bytes, and no candidate out-of-line claim variant.
The baseline stream-validation selection contains 24 explicit claim call
sites and the load-FAT selection 51; both candidate selections contain zero.
The candidate retains a 132-byte bounds-error function and a 171-byte
conflict-error function. Every `u32` fits `usize` on this x86-64 host, so the
conversion-error function is absent from the generated binary; the checked
conversion and its exact failure path remain in portable source.

Absence of a symbol alone is insufficient proof of inlining. In
[candidate stream-validation variant 0](candidate/assembly-validate_stream_allocations-0.stdout),
the root-mini-stream loop at `0x2862820` and regular-stream loop at
`0x2862c70` each contain the sector load, role-map bounds comparison and
out-of-bounds branch, existing-role load and refusal branch, then the role
store. The two stores use the existing MiniStream and RegularStream enum
values. There is no call on either successful loop path. The error branches
remain outside that loop and perform calls. Source review and contract tests
separately verify their exact typed errors and mutation behavior.

The [matched profiles](profile-comparison.json) retain positive incoming
constructor edges for 40 timed dumps per stage, excluding six CFB setup dumps
per stage. XLS owned one-cell constructor Ir falls by **18.9967% / 19.0280%**
for the two repeats. CFB few-large falls by **21.4860%** in both repeats;
many-small and tiny fall by **0.3504%** and **0.2325%**. The old claim symbol's
zero candidate attribution does not mean ownership checks became free: their
remaining work is now in callers. Parent totals, not subtraction of overlapping
inclusive owners, establish the instruction reduction.

The chain collector and full physical-reconciliation pass retain their
matched instruction totals. The latter's twelve recorded variants remain
275 bytes each. This candidate does not merge collection with claiming or
remove the final physical scan. Rejected visited-bit fusion and freshness
session proposals remain rejected.

Inlining grows the selected stream-validation code from 25,832 to 28,373
bytes and selected load-FAT code from 80,175 to 85,913 bytes. The complete
normal executable grows from 60,132,200 to 60,140,440 bytes (**8,240 bytes,
0.0137%**). These are symbol-selection and executable-size observations,
not a complete instruction-cache or working-set measurement. The native
scenario and RSS guards assess the practical captured effects.

Hardware captures have 100% grouped-event coverage, but encompass fixture
setup, query/oracle work, drop and report generation as well as constructors.
They cannot establish operation-local cycles, IPC, branch prediction, or
throughput. Native elapsed time and operation allocation vectors remain
separate evidence. The standalone harness has its own Cargo workspace with
no explicit release LTO profile; root-workspace LTO is not applicable.

Whole-child grouped counters show cycles **-3.8477% / -3.1821%**, instructions
**-9.2474% / -9.1767%**, branches **-9.3843% / -9.3087%**, and branch misses
**-2.4075% / -4.3496%**. IPC decreases from **1.6539 / 1.6674** to
**1.5610 / 1.5641**, and the branch-miss ratio increases from
**0.2305% / 0.2339%** to **0.2482% / 0.2467%**. The reduced instruction and
branch denominators matter to these ratios. No IPC or branch-prediction
improvement is claimed, and these whole-child values do not replace the
constructor clock or parent-constructor Ir evidence.
