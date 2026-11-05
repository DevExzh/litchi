# 0362g — Return the remaining PPT inverse-replay regression

- Date: 2026-10-21.
- Status: returned for repair; no production PPT changes merged.
- Candidate: `e3aa14d255573f5ca51c4cd9f96a698eb8c1b1cb`, following
  the returned [0362f](0362f-ppt-replay-performance-review.md) candidate.
- Next source repair: `c0f346607bb28559ae21fa780c835dc70e87fb0b`, pending
  its own validation and measurements. Its results cannot be pooled here.

The candidate removes replay's temporary conversion of the exact source
`Vec<u8>` into `Arc<[u8]>`. It preserves content authentication, independent-copy
replay, positional source fences, signed-source refusal, receiver options and
candidate-length admission before reconstruction. No-op replay returns the
unchanged receiver with its patch and diagnostics.

The [source review](results/0362g/source-review.json) found no correctness
blocker. The isolated final source passes 2,109 distinct PPT/CFB tests, with
three existing ignored tests, and both all-target/all-feature Clippy commands.
The [build receipt](results/0362g/candidate-artifacts.json) records the six
commands, logs, hashes, matched registry lock, probe, fixture and protected
binary. This is targeted validation, not the final workspace gate.

## Native comparison

The unchanged workload and calibration protocol from 0362f run five complete
ABBA blocks against each reference. Every observation is retained in the
[measurement receipt](results/0362g/measurements.json). Ratios are
candidate/reference; values above one are slower or larger.

| Metric | Original reference | Matched SipHash reference |
|---|---:|---:|
| First commit | 0.959715 | 0.948226 |
| Steady commit | 0.879488 | 0.906921 |
| Commit plus publication loop | 0.876164 | 0.907096 |
| Forward replay | 0.762028 | 0.968509 |
| Inverse replay | **1.012124** | 0.914126 |
| Process peak RSS | 0.154982 | 1.001782 |
| Logical retained bytes per patch | 0.156494 | 1.000000 |

Original-reference inverse ratios are `1.012124, 1.232539, 0.940150,
1.083973, 0.945870`: three of five blocks are slower. The matched-Sip inverse
ratios are `0.889542, 0.828820, 0.997379, 0.914126, 1.115805`; its slower fifth
block remains included. The matched-Sip improvement does not waive the
remaining original-reference regression. The candidate remains returned.

The matched-Sip comparison uses `N=7400, FIRST=3200, REPLAY=35400`; the
original comparison uses `N=7400, FIRST=3300, REPLAY=35800`. Each calibrates
against both endpoints. Minimum measured phase spans are respectively
554.56 ms and 599.68 ms. The receipt records every phase duration and the
minimum-span check; the steady duration uses the rigorous median lower bound,
not the wider publication-loop timer. Each inverse is checked byte-for-byte.
The references and their source/dependency provenance are those retained in
0362f; the new candidate binary has SHA-256
`3fae0cb2b27ec4ef8b21263d12824906108a616ce6a56d0b98496bb12409b8bb`
and size 1,770,856 bytes.

The [runner](results/0362g/ppt_phase_abba.py) records all exact regeneration
commands in the receipt. Its guard recognizes the added clone/probe/test
process names. An initial attempt stopped before the first calibration when
the expanded guard matched persistent installed `/usr/local/bin/peculiar_shell`
cluster services. The subsequent runner distinguishes those installed
services from build-tree test processes; all process rows remain recorded.
Both completed runs reject competing compilers, profilers and known probes
before, every 100 ms during, and after each child. The
[attempt inventory](results/0362g/attempts.json) retains that initial stop and
all completed raw files by name, size and SHA-256. No unfavorable observation
was excluded on its value.

## Heap attribution and next repair

The matched-Sip RSS increase is small, but is not dismissed as a performance
exception. A separate `N=100, FIRST=10, REPLAY=100` heaptrack capture reports:

| Endpoint | Allocation calls | Reported peak heap |
|---|---:|---:|
| Original | 368,599 | 10.51 MB |
| Matched SipHash | 350,101 | 2.27 MB |
| `e3aa14d255` | 323,171 | 2.20 MB |

The [heap receipt](results/0362g/heap.json) retains commands, raw capture
hashes, analysis and host snapshots. Peak values use heaptrack's rounded
display; the sampled Massif maximum is not substituted for the true peak.
These captures support allocation/heap attribution only. After an initial
capture filename assumption and an external-build guard stop, the two
remaining captures ran with concurrent builds explicitly recorded. No
instrumented latency or RSS is used for acceptance.

Source review identified a remaining whole-source copy: replay owns the
source buffer, but `BytePatch::rebuild` clones it before resizing and writing
the changed ranges. The next candidate derives the digest recipe first and
then consumes that buffer. Its admission, source freshness, signature,
exact-byte, recipient-policy and allocation-reuse checks must pass, followed
by new matched native measurements. W4/W4b/W5/W8 integration and the complete
final gate remain required.
