# 0741: unfinished experiment, superseded before integration

Disposition recorded on 2026-09-27 at `f8f69d5caf`. **Do not integrate the
candidate files or use this packet for a current performance comparison.**

The experiment started at `009d515bef` after the 0740 profile. It proposed
transient verified compressed-image imports for owned PPTX slide copies.
While this work was interrupted, [0742](../../0742-pptx-owned-cross-copy-media-transfer.md)
implemented that optimization and [0751](../../0751-pptx-cross-copy-apply-digest-reuse.md)
changed the subsequent proof/digest path. Current patches use LPCP0004;
this unfinished candidate explored compatibility within LPCP0003.

What was actually completed:

- Two source-bound baseline harness builds and four one-sample qualifications
  against the full 0739 oracle, before any candidate integration.
- A pre-change Store-image fixture, genuine LPCP0003 forward/inverse patches,
  and an exact expected target. The generator verified forward application and
  exact inverse restoration. Its first failed attempt is retained separately.
- Isolated, uncompiled OPC/PPTX candidate copies and draft measurement scripts.
  No candidate tests, candidate build, matched measurements, or production
  integration were completed. Syntax/format checks do not establish correctness.

`supersession.json` records the source drift: 835 entries of the original
7,281-entry source census changed or disappeared, four captured constraint files
changed, and ADR 0032 was added. The original manifests are deliberately retained
unchanged. The old guards therefore refuse the current checkout.

`compare.py` has partial validation improvements beyond `compare-notes.md`;
neither is a qualified final measurement contract. The candidate build and
qualification scripts were syntax-checked only. There is no formal capture
matrix and no speedup, allocation, or compatibility claim from this experiment.

Before cleanup, all three original binary hashes/sizes and every fixture and
generator hash in `legacy-capture.json` were reverified. `cleanup.json` records
removal of the owned baseline target (2,156,373,494 logical file bytes) and packet
Python cache. No current build directory or unrelated workspace file was removed.
`artifact-manifest.json` seals the archived packet, excluding itself.

The baseline fixture remains historical evidence only. Rebuilding it requires
the recorded baseline revision and locks; running its generator against today's
owners would not reproduce the original compatibility contract.
