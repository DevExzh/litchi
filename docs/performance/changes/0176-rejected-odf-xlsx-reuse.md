# Change 0176: rejected ODF/XLSX retained-readback experiments

Date: 2026-08-17

Status: production experiments reverted

## ODS authenticated `content.xml` reuse

The experiment retained owner-bound source version/length/size/SHA-256 proof
with the ODS source snapshot so changed publication could avoid a second
`content.xml` payload read and decompression. Security review found and fixed
stale unencrypted `manifest:size` and encrypted-plaintext-size edge cases;
focused common/ODS tests and strict warning/deprecation gates passed.

The clean A/B/B/A measurement used the existing media-rich one-cell and 21-cell
selectors, CPU 2, one logical CPU, 20 warmups and 200 samples. Output hashes
were exact in all legs. Both paired directions regressed source-backed p50:

| Selector | A1 -> B1 candidate regression | B2 -> A2 candidate regression |
|---|---:|---:|
| one existing cell | 1.84% | 2.70% |
| 21 existing cells (`ceil(1%)`) | 1.63% | 2.83% |

The second proof hash offset the avoided small payload read on this corpus.
Commits `c775989a0` and `0188c0850` record the experiment and exact revert. No
ODF production behavior from the experiment remains.

## XLSX conditional-formatting readback reuse

The experiment passed the rewrite's already parsed conditional-formatting
readback into `SourceEdit::commit`, eliminating one duplicate worksheet parse.
Focused tests proved one parse for a changed commit and zero for exact no-op;
the complete XLSX unit suite and strict checks passed.

The clean A/B/B/A measurement used the existing eager/source-backed selector,
CPU 2, one logical CPU, 20 warmups and 200 samples. Output hashes were exact.
The source-backed paired directions disagreed: A1 -> B1 regressed 4.81%, while
B2 -> A2 improved 1.99%. This fails the usefulness and repeatability gates.
Commits `153c3ff24` and `1ff922739` record the experiment and exact revert. No
XLSX production behavior from the experiment remains.

## Decision

Do not revive either narrow handoff without a materially different design or
corpus attribution. No latency, I/O, allocation/RSS, memory, cold-cache,
producer, or CRUD-coverage claim is accepted from these rejected experiments.
The retained raw reports document why the tempting duplicate-work removals
were not merged as active production behavior.

Artifacts:

- [summary](../results/rejected-reuse-0176-summary.json)
- [manifest](../results/rejected-reuse-0176-manifest.json)
- raw ODS and XLSX A1/B1/B2/A2 reports listed in the manifest
