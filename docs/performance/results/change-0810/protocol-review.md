# 0810 pre-measurement protocol review

Current base is `3677e31be5c9d5582a1f6d531ebb4d54db5a0acc`. The prior turn
is progress: 0809 repairs the independent baseline Clippy failures and passes
all six PPTX quality gates. All 35 previously read architecture/taxonomy inputs
and the unrelated three working-tree files retain their recorded hashes.
Accepted ADR ownership, bounded validation, preservation, and publication
constraints remain binding; no ADR change or iWork work is proposed.

The candidate Rust bytes and one-file production patch are exact copies of
0808's reviewed direct-event candidate. The notes codec consumes Reader events
while delegating namespace scope and resolution to the same public owner.
Pending pops precede reads; namespace pushes precede existing node/depth checks;
namespace-error conversion, root/name order, strict/transitional behavior,
attribute checks, limits, and both buffered oracle functions remain unchanged.
The three focused differential tests are retained. No independent manual
namespace parser, validation bypass, public API, dependency, or unsafe change.

Root builds and verifies three before binaries, runs the 36-test probe quality
lane, and accepts all eighteen baseline reports against sealed source/output/
full semantic and extension oracles before application. Candidate application
must change only `crates/litchi-pptx/src/notes/codec.rs` within the 9,196-file
source census. All six candidate PPTX quality gates then pass before after
measurement builds and probe quality. Root's Cargo lockfile and rustfmt config
are captured before any build. Prior 0808/0809 tests are not substituted for
candidate quality, and no historical timing is pooled.

The matrix is six shapes crossed with capture, commit, and lifecycle. Native
sampling uses CPU 12, six paired blocks AB BA AB BA BA AB, thirty measured
samples, and three warmups per process (216 reports/6,480 samples). Allocation
uses two paired AB/BA blocks, three samples, no warmup (72/216). Before-only
qualification contributes 18/18. Four owner-scoped large-capture Callgrind
processes run AB then BA with one sample and no warmup. Totals: 310 retained
reports and 6,718 measured samples; warmups are excluded.

The policy is frozen before builds: at least one capture/lifecycle row needs
at least 3% improvement with bootstrap high endpoint below 1. Any of eighteen
paired process-p50 ratios above 1.05 with low endpoint above 1 vetoes adoption.
No paired allocation-block median may increase calls, allocated bytes, net live
bytes, or peak above entry. Allocation reductions alone are insufficient.
Bootstrap uses Python Random seed 810810, 10,000 resamples of six paired
process ratios, median statistic, sorted endpoints 250/9749; process quantiles
use nearest rank ceil(n*p)-1. Tail latency, RSS, and spread flags remain visible
for review rather than being hidden in an aggregate.

Callgrind is diagnostic guest-Ir attribution only, with exact owner
`namespace_uri_probe::capture_region_0793`, one positive numbered publication,
empty termination, one owner call, and exact self/edge conservation. The target
mechanism is removal or restructuring of the scanner-to-process_event edge;
other NsReader users may remain. No global zero-call assertion or extra
instruction-savings adoption threshold is justified. No native phase-fraction,
causal cycle, profile-latency, cross-format, perf, or heaptrack claim is made.

Root alone runs all Rust commands and captures, serially; independent agents
prepare and review source/drivers/readers. Heavy offline replay begins only
after all captures are terminal. Final adoption or exact restoration, binary
identity checks, owned-target cleanup, result replay, exact staged/committed
blob audits, and a truthful report precede batch completion. The broad goal
remains incomplete regardless of this candidate's outcome.
