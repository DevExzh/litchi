# 0362f — Return PPT replay for performance repair

- Date: 2026-10-21.
- Status: returned to W8. The candidate is preserved on
  `uo2/ppt-0362f-candidate` at `be72b3a163f4fc1bc58a54c4d9a18b7fb8eba0b6`;
  its production changes are not merged into trunk.
- Contract: [0351](0351-api-contract.md); required first-commit follow-up to
  [0357b](0357b-ppt-block-roots-and-durable-structure-edits.md).

The candidate removes the original first-commit regression, but forward and
inverse replay exceed the matched SipHash reference in four of five paired
blocks. This is required repair, not an accepted performance exception. The
original reference's larger improvements do not waive that result.

## Reviewed implementation and correctness repair

The candidate uses a process-keyed 128-bit SipHash digest for each 512-byte
block, with SHA-256 over the image length and ordered block digests. Compact
spans retain changed blocks. Patches remain opaque, process-local and bound
to exact content. Native scoped editors reuse the canonical owner's admitted
presentation and retained CFB index. Signature discovery excludes the root
entry and unreachable slots, preserves source fences, and bounds traversal
when the retained event policy requires it.

Changed publication seeds an unsigned memo only after authenticating and
checking the source. Its closed typed publishers cannot introduce a reachable
signature entry. Source freshness is checked before using the memo; exact
signed no-op publication remains legal.

Coordinator review found that `Base::publish` constructed compact target
spans before admitting the candidate's length. The source's admission alone
did not prove the target bound. The preserved candidate admits the target
against the retained input-byte ceiling first, while preserving source-change
and signed-source error precedence. Tests cover exact/one-over limits,
`usize::MAX` without a large allocation, actual oversized publication,
signed input and a changed positional source. This finding is fixed; the
performance finding below remains open.

## Validation scope

Broad CFB/PPT library and integration runs plus added boundary checks pass
2,107 distinct tests, with three existing ignored tests. The broad runs
precede the final early-length admission fix; after that fix, all 15 scoped
tests, two-crate all-target/all-feature Clippy with `-D warnings`, and
workspace formatting pass. This is targeted validation, not the final
integrated gate. [Validation receipts](results/0362f/validation.json) retain
the commands and log identities.

An initial coordinator test incorrectly expected a tight-policy package to
open before checking its signature projection. Its failure demonstrated the
earlier admission boundary. The corrected expectation and a separate large
admitted-directory test pass; the initial failure remains recorded.

## Native measurements

Five completed ABBA blocks are retained per reference. CPU 29 is pinned;
every block uses common A/B/B/A counts. First-commit and replay metrics are
phase-total/count means. Steady commit uses the probe's per-operation median;
the complete commit loop, which includes serialization and patch retention,
is reported separately in the machine-readable evidence.

| Metric, median candidate/reference | Original `e673d4addb` | Matched SipHash `e4bdab6066` |
|---|---:|---:|
| First commit | 0.9675 | 0.9155 |
| Steady commit | 0.8810 | 0.9104 |
| Forward replay | 0.7692 | **1.0077** |
| Inverse replay | 0.9872 | **1.0466** |
| Process peak RSS | 0.1550 | 1.0003 |
| Logical retained patch bytes | 0.1565 | 1.0000 |

The final candidate retains 15,384 logical bytes per patch versus the
original's 98,304; this measure includes compact block spans but excludes
allocator metadata. RSS is measured by `/usr/bin/time -v`, not inferred from
that logical count.

Every accepted phase must span at least 150 ms. Calibration targets 500 ms
with headroom. The final original and matched-Sip comparisons have minimum
phase spans of 559.4 and 549.2 ms. The steady qualifier uses the rigorous
lower bound `floor(N/2) * (printed_median_ns - 0.5)`, accounting for rounding;
it does not substitute the wider commit-loop timer. The probe verifies every
inverse artifact byte-for-byte outside timing. Functional tests cover exact
forward replay; the runner also checks endpoint extents.

Inverse paired ratios against matched SipHash are 1.2240, 1.1041, 0.9527,
1.0466 and 1.0310. They vary substantially, but their median and four slower
blocks require investigation. No noise or code-placement exception is
accepted. Source review identifies a redundant replay copy: `Base::capture`
reads a full `Vec`, converts it to a copied `Arc<[u8]>`, then rebuilds a
second full target. A repair can authenticate the local vector and borrow
the package during publication. It must retain encryption/signature policy,
recipient limits, no-op behavior, exact restoration and source freshness.

## Provenance and excluded attempts

[Measurements](results/0362f/measurements.json) contain every completed
block, both build manifests, commands, source and binary identities, and
inventories of external raw artifacts. [Dependency parity](results/0362f/dependency-parity.json)
confirms all 518 registry package identities/checksums and release profiles
match across the original, rebuilt SipHash and final candidate. The
historical SipHash binary used different dependency versions and is excluded.

The measured final source is base `00673b9006` plus patch SHA-256
`3d725e367f95ffa732be299a9c4f27590e1594bad71eeb81904cca4226e0b7b6`, now preserved
as `be72b3a163`. The final binary is 1,764,192 bytes, SHA-256
`ca2721c456769f99056a7541dc2abe02fe14153e35876fa010a22a7ea69123fe`.
Build with `cargo build --release --offline --locked -p litchi-ppt --example
scoped_patch_probe`, using the recorded normalized lock (registry versions
from `8f84379e31`) and two capped Cargo jobs. The source checkout's ordinary
lock is not the measurement lock.

The 1,363,443-byte external `source-digests.json` has SHA-256
`e2904c09c7d4ed9b36e29291c32275e132600341d10a13beca48592d0cffbab3`.
[Candidate artifacts](results/0362f/candidate-artifacts.json) also records
the external metadata, source patch and build log with regeneration commands.
No compiled executable or large capture is committed.

The [runner](results/0362f/ppt_phase_abba.py) checks the host before, every
100 ms during, and after a child, preserving raw output before rejecting
overlap. Initial startup failures and partially completed attempts remain
recorded. The original final comparison combines three completed blocks
from `original-final-v1` and two from `original-final-v2`; a bookkeeping-shell
false positive interrupted the former. The guard was corrected to match
actual executables or interpreter scripts rather than strings embedded in
shell commands. The matched-Sip comparison completed all five blocks in
`sip-final-v1`. Earlier `8f` measurements are separate groups because the
production source changed; they are not pooled with the final candidate.

Regenerate the reduction with `python3
docs/unified-ops/results/0362f/reduce.py --coord (deleted-coord-artifact) --out OUT` using the
retained raw directories. The runner's complete invocations appear in each
run's measurement record; replace its external path with the committed
runner and choose a fresh output directory for a new attempt.

W8 owns the replay repair and repeat native qualification. W4/W4b/W5/W8
integration, remaining shared/raw audits, glossary/inventory closure and the
full workspace gate remain required.
