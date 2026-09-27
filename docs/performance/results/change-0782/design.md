# 0782 optional borrowed text slice

The previous turn made progress: 0781 repaired performance CI and rejected a
measured Cow candidate, retaining all observations and restoring production
source. This experiment starts at `f1df64d15a8f2711e554f58acb413b06d56ef1ca`.
Its production crate census is unchanged from the 0781 baseline. All 35
previously read architecture inputs and the three unrelated main-worktree
files were rehashed unchanged before work began.

The current construction graph needs no owned variant in the private plain
`UserShapeData.text` field. Source writer strings live through synchronous
serialization. Rich paragraphs and centered fallback continue to own their
copies because their conversion can mutate paragraph or smart-tag state.
The new candidate uses `Option<&str>` and propagates its source lifetime
through the same private group and codec types. The public writer retains
its existing owned strings. No public API, dependency, global cache, unsafe
code, source limit, or existing-document preservation policy changes.

The 0781 large-text benefits and short-text regressions motivate this narrower
representation. They do not identify the cause of the regression or predict
this candidate's result. A layout or branch explanation remains a hypothesis
until supported by separate evidence.

The exact corrected 0781 probe source, manifest template and lockfile are
inherited unchanged, including its tool identifier and authored fixture
identities. New builds and captures bind new paths, source censuses and binary
hashes. The current run compares its own baseline with its own candidate;
historical timing values are not pooled or treated as a third paired leg.

The frozen protocol retains ten cases: tiny, many short text boxes, large
ASCII payload, large Unicode payload, and rich text, each in public write and
fixture-to-publication lifecycle modes. Six alternating native blocks use 30
samples after three warmups. Two separate allocation blocks use three samples
without warmup. CPU 12 and all exact orders are retained in `plan.json`.
Before application, qualify all ten baseline cases and capture payload/write
allocation ancestry. Capture the complete pair before any adoption decision;
do not resample based on observed results.

Correctness requires identical source/output identities, public-reader text
and direct untrimmed text-atom checks for every case. The strict corrected
Unicode oracle is inherited intact. Candidate quality requires formatting,
all-feature/all-target compilation, all-feature PPT tests, warning-denied
library Clippy and rustdoc, and repository crate boundaries. Standalone probe
tests run in their verified synthetic-counter and real-allocator configurations.

Adoption requires a useful public-case improvement, no new retained-memory
cost, and no persistent control regression: a paired p50 median above +5%
whose 95% interval remains above zero counts against adoption in any of the
ten cases. Report all tails, RSS, allocation metrics, spread and paired flags
regardless of that guard. Lower allocation counts alone do not justify a
common-workflow slowdown. No physical-copy, cold-cache, concurrent, device,
native Office interoperability, or complete CRUD claim follows.

ADRs 0001/0002/0024 keep correctness and ownership boundaries; 0003/0004 keep
the public model unchanged; 0005 permits removing transient work with measured
evidence; 0006 preserves output/refusal contracts; 0008 requires final gates.
No derived-value memo under 0032 is introduced. OLE2/OOXML remain the active
priority, ODF is deferred under the existing owner decision, and iWork is
excluded. The broader goal remains open.
