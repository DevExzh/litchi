# 0801 — carry duplicate state forward without value replay

The candidate is rejected: empty-tag consumption breaches the frozen protected
regression guard. Production remains at the corrected 0800 baseline.

This fresh preflight rebases the deferred 0800 candidate on the corrected
`54aa0f3fb4` production baseline. It tests whether carrying the first two raw key
positions forward can retain short-tag savings without the third-item replay
that rejected 0799. Production remains unchanged throughout this packet.

The first two attributes use quick-xml with duplicate checks disabled. Before
the second and third value, the candidate compares raw key prefixes with the
previous keys. After a successful third attribute with a non-whitespace tail,
it seeds a boxed ordered map from those borrowed names. Later items precheck
the map before asking quick-xml to scan the value. No earlier value is replayed.
The corrected leading-equals key behavior and first-error contract remain in
the archive and regression matrix. The cost tradeoff is earlier ordered-map
allocation and checking on longer tags, plus a measured 128-byte iterator
instead of the 120-byte baseline.

## Protocol and verification

The [packet](results/change-0801/README.md) retains the 0799 literal 39-case
matrix and advancement rules unchanged. Two helper legs share one release
binary. Six alternating native blocks use 30 samples, three warmups and 4,096
iterations per process on CPU 12. The fresh bootstrap seed is 801080. Two
separate Callgrind repeats retain instructions and branch counters; these are
mechanism diagnostics and are not pooled with native timing.

Both one- and two-attribute consume cases must improve at least 3%, with ratio
confidence intervals wholly below 1. Any of the 18 protected consume cases
regressing over 5%, with its interval wholly above 1, vetoes advancement.
All other regressions remain review triggers. Passing preflight allows only
fresh public-workflow, resource and cross-format trials, not adoption.

Exact five-copy helper tests in isolated mirror crates pass: 70 baseline and
95 candidate tests, zero failures or ignores, with warnings-denied Clippy for
both legs. The release probe passes formatting, build, check, Clippy, its
39-case literal fixture audit, full first-error/error-position parity and
clone checks at advances 0, 1, 2, 3, 4, 5, 32 and 33. These checks do not claim
full production-crate or workspace verification.

Two initial Clippy failures and their exact inputs remain retained. The helper
attempt required equivalent Option simplification and removal of a test
identity map. The direct harness required a narrow dead-code allowance for
the unused unchecked API in its candidate module. The final helper source is
byte-identical in its archive, mirror tests and probe. No failed-attempt native
measurements were taken.

## Measured result and disposition

Both required dominant-class benefits pass, but the protected empty case vetoes
advancement. Ratios are candidate/baseline; values below 1 are faster. These are
fresh direct-helper comparisons, not public-workflow speedups.

| Consume case | Paired p50 ratio | Change | 95% bootstrap ratio interval |
|---|---:|---:|---:|
| `distinct-0` | 1.075420 | +7.542% | [1.062726, 1.077057] |
| `distinct-1` | 0.610504 | -38.950% | [0.594645, 0.714536] |
| `distinct-2` | 0.806294 | -19.371% | [0.790086, 0.806685] |
| `distinct-3` | 0.908208 | -9.179% | [0.894137, 0.922479] |
| `distinct-4` | 1.611177 | +61.118% | [1.594146, 1.638939] |
| `distinct-5` | 1.327504 | +32.750% | [1.296225, 1.410613] |
| `distinct-16` | 1.680343 | +68.034% | [1.667178, 1.724234] |
| `distinct-32` | 1.599407 | +59.941% | [1.582182, 1.667345] |
| `distinct-33` | 0.845074 | -15.493% | [0.836047, 0.850544] |
| `distinct-64` | 1.001356 | +0.136% | [0.994029, 1.010791] |
| `duplicate-long-quoted-after-33` | 0.552602 | -44.740% | [0.484738, 0.558279] |
| `syntax-flag-after-2` | 0.961742 | -3.826% | [0.952555, 0.981642] |

There are 13 significant consume regressions in total. Besides the empty case,
they cover distinct counts 4, 5, 8, 9, 16, 17 and 32; valid duplicates after 4
and 5; and all three syntax tails after 4. The empty case is the sole protected
veto. The other twelve remain material review triggers, including regressions
up to 68.034%; fixing only the empty case would not establish suitability for
public workflows. All 78 case/mode rows, samples and confidence intervals are
retained in the [full summary](results/change-0801/summary.md).

The fresh candidate avoids the 0799 third-item replay regression: three distinct
attributes improve 9.179% against the current baseline. No historical samples
are pooled to compute that number. Starting the ordered map early introduces
substantial costs on the next boundary. This is a useful design constraint for
future work, not evidence to waive the failed guard.

Native spread exceeds 5% in 50 of 156 case/mode/leg groups. All are retained,
including variation in construction rows and one-attribute consumption. The
one-attribute benefit still has its full interval below 1. No tail-latency or
resource improvement is claimed.

All 312 Callgrind owners qualify; each has one positive dump and an empty
termination dump. Independent scalar conservation agrees for all 624 dumps
and all five events. In both repeats, empty consumption rises from 170 to 184
guest instructions while opaque construction falls from 70 to 69. Four-attribute
consumption rises from 1,702 to 3,090 instructions; sixteen attributes rise from
9,211 to 17,300. These observations locate additional work, but neither iterator
size nor instruction counts alone establish a native-time cause. In particular,
three-attribute guest instructions rise from 1,286 to 1,463 while native time
improves. Counter and native results remain separate.

The complete capture contains 936 native reports with 28,080 samples and 312
counter reports with one measured sample each. Independent input, error,
checksum and paired-statistic audits agree with all 78 analysis rows. Post-cleanup
replay passes. The owned target removal frees 180,293,462 logical bytes, with the
final and failed-build executable identities retained. All 9,196 production
source hashes, 35 architecture inputs, unrelated files and other worktrees remain
unchanged. The candidate and failed attempts are evidence archives only.

No public-workflow/resource/cross-format trial is authorized for this candidate.
A future design must retain early duplicate refusal and avoid both empty-path
regression and the early ordered-map penalty, under a newly frozen experiment.
Existing production performance claims and CRUD coverage remain unchanged;
iWork remains excluded and the broader GOAL work remains incomplete.
