# 0783 — current PPTX lifecycle phase attribution

Status: diagnostic evidence; production code is unchanged. On the generated
100-slide × 100-shape corpus, initial opened-presentation capture accounts for
about 69% of the instrumented lifecycle, serialization 23%, and staged commit
4%. The next qualified CPU profile should target initial capture. This does
not attribute the historical 0780 lifecycle regression or establish a speedup.

The 0780 change improved staged commit while slightly slowing its large full
workflow. The current diagnostic retains that probe's exact corpus generator,
marker, public calls and semantic readback. Package ingress remains outside
the clock. Consecutive timestamps split capture, edit staging, commit,
publication and serialization. Each measured sample's five durations sum
exactly to the total. Timing includes intervening clock overhead and temporary
drops; publication includes dropping the returned snapshot. Report-map
allocation and readback are outside the clock.

A feature-off control retains the original clock sequence. Six alternating
blocks cover tiny (3×4), medium (12×8), and large (100×100) generated
presentations, thirty samples after three warmups per process, pinned to CPU
12. This is current-source in-memory evidence, with 36 processes and 1,080
measured outputs. Every output reopens with exact expected text and matches
the inherited 0780 source/output/semantic identities. No process is excluded
or retried.

## Results and interpretation

The [generated table](results/change-0783/phase-summary.md) and
[analysis](results/change-0783/analysis.json) retain process-level summaries,
paired instrumentation ratios, uncertainty and flags. Phase shares below are
medians of six process medians of per-sample fractions; they are not ratios of
independently aggregated durations and need not sum to exactly 100%. Process
p50 uses nearest rank. Paired total ratios use the median across the six
matched process pairs; 95% bootstrap intervals use 10,000 resamples with seed
783078 and describe this capture rather than population guarantees.

| Corpus | Capture | Stage edit | Commit | Apply | Serialize |
| --- | ---: | ---: | ---: | ---: | ---: |
| Tiny | 17.233% | 4.399% | 15.233% | 2.806% | 60.115% |
| Medium | 23.908% | 4.840% | 14.778% | 2.749% | 53.454% |
| Large | 69.051% | 2.671% | 4.290% | 0.631% | 23.320% |

Median paired phase/control total-latency changes are −0.949%, −0.943%, and
−3.653% for tiny, medium and large. In particular, the large instrumented build
is consistently faster. This difference can include code-generation/layout and
measurement effects; it is neither negative clock overhead nor a production
improvement. Use these shares to select a profile, not as unperturbed exact
fractions or a causal explanation of the earlier before/after result. The
frozen one-sided perturbation guard does not turn a faster diagnostic binary
into proof of measurement equivalence. The 95% ratio intervals are
[0.989173, 0.996081], [0.989369, 0.992642], and [0.958465, 0.964104].
No total-p50 or phase-p50 spread exceeds 5%; three whole-process RSS spreads
do (tiny control, tiny phases, and medium phases), and all remain retained.

For tiny and medium presentations serialization is the largest phase. For the
large shape-heavy corpus capture is dominant; further constructor-only or
staged-commit tuning has a limited whole-workflow ceiling on this case. A
qualified operation-local capture profile should separate slide/text parsing,
MCE handling, notes graph validation, package fingerprinting and snapshot
construction. Source paths are hypotheses until directly profiled. Keep
required validation and exact error order under the accepted ADRs. Broader
corpora, native producers, source-backed/provider intersections and scaling
remain separate requirements. The borrowed `Package::from_bytes` ingress used
here has no retained source archive, so serialization follows the full writer
route. These fractions do not describe the owned source-preserving route.
The [source review](results/change-0783/source-review.md) records exact future
profiling seams and this route distinction.

## Validation and reproducibility

Both release feature variants and probe formatting pass. They retain sixteen
inherited unused allocator/constructor helper warnings. Production gates were
not rerun; no production source changed. All 35 architecture and goal input
hashes remain unchanged. The packet retains source/probe/lock/build identities,
raw reports, RSS, commands, and the frozen plan. RSS is a whole-process maximum
including setup and verification; it is not phase-local retained memory.

Replay uses `python3 -B docs/performance/results/change-0783/validate.py` and
`python3 -B docs/performance/results/change-0783/analyze.py --check`.
See the [packet README](results/change-0783/README.md) for scope and fresh-run
instructions. The seal covers 144 payload files. Both executable identities
were verified before removing the owned target (696,897,828 file bytes).
Complete replay passes after removal; analyzer byte-for-byte replay also passes
from a temporary relocated packet, which was then removed. This evidence-only
batch created no worktree or branch. Preexisting worktrees and the three
unrelated local-file hashes remain unchanged. OLE2/OOXML remain active, ODF
deferred and iWork excluded; the comprehensive goal remains open.
