# 0973 — fresh instruction-profile pilot stops on event loss

**Status: pilot failed; custody and failed-pilot audit complete.** This record
reports the first-child result of the fresh 0973
`FloatingPictures/e1/default` diagnostic.
The retained 0967 target was used without a production source edit or target
rebuild. The runner stopped at the declared gate and did not retry.

## Binding and controls

The final-v2 plan requested one serial 12-child cohort on CPU 31, with five
warmups, 80 measured samples, `instructions:u` period 1,000,000, and the
ASLR-safe corrected decoder. The final-v2 helper seal, ASLR translation v4,
and end-to-end control checks passed before the root GO. The plan and helper
were sealed for custody; these controls do not waive the per-child loss gate.

## Pilot result

Child 01 exited natively with rc0 and a valid report, but its packet records
`lost_events=8`. The raw stream contains two `PERF_RECORD_LOST` records, each
reporting four lost samples. The independent audit confirms that the
`PERF_RECORD_LOST_SAMPLES` trailer also reports eight: it is an end
`ID|LOST` counter representation of the same loss, not an additional eight
samples. The corrected decode reports 6,931 sample records and
6,931,000,000 total period, with 947 marker samples carrying 947,000,000
period and zero empty callchains. The stock decode has 4,540 empty callchains
for the same captured child. These decoder counts are descriptive; they do not
qualify the profile or establish a causal owner.

The runner exited 2 with `per_child_marker_gate_failed` because lost events
were nonzero. Children 02 through 12 were not started, `no_retry` is true, and
there is no all-12 reducer result. The retained raw `perf.data` is 228,856,640
bytes with SHA-256
`4140b417724ea9606b6637ae6aade3a874a682bd9baf84d6fa0f6fdb716acbaa`.

## Interpretation and next step

Event loss is a hard validity failure for this pilot. The observed marker
count and corrected-versus-stock decode difference remain useful for the loss
investigation, but they are not a timing, phase, speedup, adoption, or source
attribution result. The audit establishes the loss accounting and packet
identity, but does not establish whether the loss arose from capture transport,
event admission, collector capacity, or another mechanism.

The next step is a separately reviewed capture-loss design with explicit loss
controls and custody checks. The bounded [next-design note](results/change-0973/analysis/loss-mitigation-next-design-v1.md)
(SHA-256 `adf038de8f3c8943e9e8f41b2d93259e2759023c9cb112779bd8fd01d5075d91`)
records an unproven hypothesis to test an explicit `-m16M` buffer while holding
the event, stack, CPU, and gate settings constant; it does not establish a
cause, identify a default buffer size, or authorize a new capture. The loss
mechanism must be addressed before another cohort is considered; no retry or
relaxed gate follows from this result.
No production/runtime change, iWork workload, or target build was performed.

## Evidence

| Evidence | SHA-256 |
| --- | --- |
| [final-v2 plan](results/change-0973/plans/plan-0973-final-v2.json) | `12bd803923586cf73c0881bfd3ef7e3cbc7a1e4853867c91d422b6333fa2e3b4` |
| [final-v2 helper seal](results/change-0973/receipts/helper-seal-0973-final-v2.json) | `d7a3bc9807b1339d3ecc8b3de2fffb25cdd07f43b1eefe6d0d151b15f535c2a5` |
| [ASLR translation v4](results/change-0973/receipts/aslr-translation-0973-final-v4.json) | `89a48f57bf35cb556d890bcfd7476a3a6e27601fd030a643c6e3275384bfc74b` |
| [root terminal](results/change-0973/receipts/root-terminal-0973-v1.json) | `07fded0dcafa601a4fea20fb73f8eeb343f997fda4b390aee1b1f80a378f9711` |
| [failed capture manifest](results/change-0973/capture/capture-manifest-v2.json) | `7c0b0e159c510e63f61fd45ca1af516fa95b72ae21bb7b4c8a080dfc1f2cd6e3` |
| [author seal](results/change-0973/receipts/author-seal-0973-final-v2.json) | `e7bab50fefb2b374ac99e67c6963b370c24bc44252ca7ed5f9e2f33f876a1433` |
| [pre-capture independent review](results/change-0973/receipts/independent-review-0973-final-v2.json) | `24533a04dcbba04f47d9a298fbb71bc6934316b9cddf66f60ef982bd415d76de` |
| [failed-pilot analyzer audit](results/change-0973/receipts/failed-pilot-audit-0973-v1.json) | `872fd7051b098901910bc200da08b9aa68399f7e8ca470a570d82b7143038a67` |
| [independent failed-pilot audit](results/change-0973/receipts/independent-failed-pilot-audit-0973-v2.json) | `977cee099b640a094425d8e7c3641cc71fe658d3c68209c9e8a4209a29e641ea` |
| [temporary cleanup receipt](results/change-0973/receipts/temporary-cleanup-0973-v1.json) | `5f632cbd85b5832fa14e1c081aeab66c45f865d97c212d2ee13c6ad01f2c5c40` |

The independent audit found all 21 bound artifacts, raw bounds, identities, and
reparse results consistent. **Performance claim:** none. **Production change:**
none. **Pilot disposition:** failed on nonzero event loss; no all-12
qualification follows.
