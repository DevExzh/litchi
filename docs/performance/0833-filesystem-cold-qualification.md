# 0833 — cold filesystem qualification exposes two harness blockers

The planned large-package baseline is **not admitted**. Fresh qualification
finds a cold-only PPTX source-replay classification failure and a cold-only OPC
save output-parity failure. No formal measurements or optimization claims
follow. Production and Rust harness source remain unchanged at `eefaca16e3`.

## What ran

The [frozen plan](results/change-0833/measurement-plan.json) proposed six
counterbalanced blocks, six cases, two cache states, 30 samples and three
warmups: 72 reports / 2,160 formal samples. All six one-sample qualification
commands had to pass first. Five succeeded; the source-backed PPTX command
failed. Three separately labelled diagnostics then isolated the failures.

| Execution | Outcome |
| --- | --- |
| OPC eager/source open, each warm plus cold-verified | Both commands pass |
| OPC eager/source one-Part atomic save, each warm plus cold-verified | Both commands pass independently; cold output hashes differ |
| PPTX eager open plus selected-slide lifecycle, warm plus cold-verified | Pass |
| PPTX source-backed selected-slide lifecycle, warm plus cold-verified | Payload-range classification fails; no JSON report |
| Diagnostic: source-backed PPTX warm only | Pass |
| Diagnostic: source-backed PPTX cold only | Same classification failure; no JSON report |
| Diagnostic: both OPC save routes together, cold only | Existing eager/source output-parity gate rejects the result |

There are six retained JSON reports with eleven qualification/diagnostic
samples, nine workload commands, and three terminal workload failures. Partial
work inside failed commands is not counted as retained samples. There are zero
formal reports. Qualification timings remain raw evidence and are not promoted
to p50, throughput, RSS or speedup claims.

The fresh offline/locked release build passes on the recorded AMD EPYC 9R45
Linux x86-64 host, using Rust 1.95.0 and two Cargo jobs, with incremental
compilation disabled. Workloads use CPU 12 and a private root on ext4. The
[prepare receipt](results/change-0833/prepare.json) binds compiler, host,
filesystem, environment, fincore executable and 9,388 source/normative inputs.
The native executable and exact build invocation are bound by the
[build receipt](results/change-0833/build.json).

## OPC: the aligned archive changes publication framing

The OPC corpus has four incompressible 4 MiB logical members. Its archive is
16,783,632 bytes, SHA-256
`a0c1af9e2c7a19148b44fc2a8c594c7a274131d74f9f042d55b487d5337cd1e6`.
Cold verification creates a 16,785,408-byte private archive by adding a
1,776-byte EOCD comment. It does not change logical members.

| Save route/state | Output bytes | SHA-256 |
| --- | ---: | --- |
| Eager warm or cold; source-backed warm | 16,783,632 | `f4bbe4de18853444cc6cd093cf561249decaa81f776afcf5de122667f5dd7009` |
| Source-backed cold | 16,785,408 | `3f66abdd5fbdc94eaf4961089cd7193deb9eaffff9004871e8f4e47544a55a66` |

The current harness explicitly calculates a route-specific expected digest for
the aligned source. Its source-backed publisher preserves that comment; the
borrowed eager reconstruction returns the unpadded length. The separate route
commands pass their existing output checks. The combined route command then
hits `eager and source-backed OPC filesystem save samples differ` because its
additional cross-route check demands identical whole-output hashes.

The source-backed OPC open qualification report also records 1,064 logical bytes in
warm state and 66,600 with alignment, a difference of 65,536 bytes. These
are observed request results, not decompressed-byte counts or physical I/O.

This establishes incompatible harness expectations for the aligned case. It
does not establish an adopted production optimization or, by itself, a new
production preservation defect. A correction must keep exact per-route output
checks and independently prove the permitted framing difference and logical
member preservation. Silently dropping comments or removing output checks is
not an acceptable repair.

## PPTX: cold-only replay rejection needs a bounded range proof

The generated PPTX has 200 slides, eight text boxes per slide, and eight 2 MiB
media parts. Its 17,017,139-byte archive SHA-256 is
`61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757`.
The warm-only source replay fully covers the selected slide's 522 compressed
bytes, with zero unselected-slide or media overlap. Its semantic digest is
`f5f7db181150c00a4323a48c142721ead73aca3ad7c3b3594e8b1a18a686b257`.

The cold-only command fails with `PPTX source replay violated
pptx_file_source_open_selected_slide_lifecycle payload-range classification`.
This check belongs to the untimed replay, not the native operation's elapsed
clock. The generic child error does not identify whether aligned-source
priming or the verified-cold child failed; it does not prove that a timed cold
sample completed. No failed request ranges are retained, so the exact
offending overlap is not proven by this packet.

EOCD alignment is the leading source-supported explanation: a nonempty archive
comment requires a broader tail search than the no-comment shortcut, and raw
metadata reads can overlap compressed payload ranges without decompressing
those payloads. The existing DOCX aligned-source replay separates this bounded
metadata probe from semantic reads. The next PPTX repair must first retain and
verify exact failing ranges, prove the alignment transform and bounded tail
probe, and retain all raw overlap counts. It must continue rejecting additional
unselected-slide/media reads; a blanket allowance would weaken the evidence.

## Verification, limitations and next action

The eight sealed 0832 after-source quality gates are reused after checking all
9,388 current inputs against that exact source, including the four final XLSX
hashes. They include the 555-pass/one-ignored harness suite, 2,090 XLSX tests,
three pinned oracle tests and associated quality gates. They are not fresh
0833 tests. The [reuse reader](results/change-0833/reuse_quality.py) binds all
receipts and logs. The fresh release build and qualification are separate.

The independent [failure audit](results/change-0833/audit.py) validates source
custody, command chronology, report/binary/log identities, cold proofs, sample
vectors, output differences, retained failures and absence of formal captures.
Its mutation checks admit the valid control and reject twelve corruptions of
cold, source, output, binary, sample and configuration evidence.

All five successful cold qualifications report zero pre-operation resident,
dirty and writeback bytes, consistent fincore provenance, successful post probes
and positive process `read_bytes`. This is page-cache/procfs evidence only;
physical-media state is unproven. OPC positional counters are timed; PPTX
positional counters come from an untimed replay; eager logical counts are
unavailable. OPC open also has different package-drop boundaries. No route
latency comparison, allocation, hardware-counter, copy-volume or concurrency
claim follows.

The [diagnosis](results/change-0833/diagnosis.md) and
[protocol review](results/change-0833/protocol-review.md) define the next harness
work. ADRs 0003/0005/0006/0008 retain publication, evidence, preservation and
admission requirements; 0010/0011/0024 retain package ownership. All 35 previously
read normative hashes are unchanged. iWork remains excluded. The broader goal
is active.

Offline replay:

```sh
python3 -B docs/performance/results/change-0833/reuse_quality.py
python3 -B docs/performance/results/change-0833/reader.py
python3 -B docs/performance/results/change-0833/audit.py
python3 -B docs/performance/results/change-0833/test_audit.py
```

The reader correctly returns `qualification_valid: false`; successful offline
replay confirms the failure record, not baseline admission. Root removed both
owned temporary roots (1,783 files / 1,247,257,569 logical bytes) after binding
the executable. A separate receipt records Python cache cleanup. Post-cleanup
audits pass, and the final packet seal binds the owned files. The three
unrelated workspace files remain unchanged.
