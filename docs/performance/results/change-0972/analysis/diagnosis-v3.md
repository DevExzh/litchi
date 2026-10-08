# 0972 — offline recovery of the 0971 callchain decode

**Status:** final v6-bound diagnostic record with the author seal and the
PF_X raw-scan correction incorporated. CPU review approved the other claims.
The external v6 replay is the sole authority for the recovered counts. The
earlier v5 replay remains retained as historical prototype evidence and is not
used for the final counts.

This record is an offline decode and forensic correction. It did not launch the
retained executable, run an Office workload, rebuild a target, modify production
source, mutate the sealed `perf.data`, or perform a new capture.

## Finding

The 0971 scanner used the target DSO runtime base `0x64d1e4e5c000`. The verified
ELF load bias is `0x64d1e4e5b000`, one 4 KiB page lower. The historical scanner
also used a truncated executable upper bound. The sealed v2 scanner derives
its executable interval from the retained ELF `PF_X PT_LOAD` program header.
The earlier statement that the raw stack contained zero marker words is
therefore retracted. The corrected raw scan finds potential address words, but
those words are not live return ancestry and are not phase evidence.

The authoritative v6 replay applies a narrow, externally compiled diagnostic
interposer to the exact target DSO only. It forwards the complete six-argument
`dwfl_report_elf` ABI, changes its fifth-argument base from
`0x64d1e4e5c000` to `0x64d1e4e5b000`, and changes the returned module start only
for direct `/usr/bin/perf` callers from the corrected ELF base back to perf's
expected comparator start `0x64d1e4e5c000`. v6 captures the direct caller PC
before classification and restores `errno` after each real call. Non-target
modules and non-perf callers receive no adjustment. The stock and no-op controls
are byte-identical, and the immutable raw input hash is unchanged.

This coupled correction recovers symbolized marker ancestry from the existing
capture. It supports a map-metadata/decoder interaction in this replay. It does
not identify the upstream perf/libdw layer that produced the original symptom,
prove a stack-depth or unwind cause, or establish what a new live capture would
do.

## Bound input and raw capture facts

- Retained capture: `/home/zhuhe/litchi-performance-0971/capture/instruction-profile-v2/child-01/child-01.perf.data`, 23,129,880 bytes, SHA-256
  `499e3008e6496b01fe09c9cf7f6f2aedbf46c1b24f7697e98289c09460143746`.
- Retained executable: `/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile`, 47,909,224 bytes, SHA-256
  `0db34357a1e725dc397a4fd895532433db1214c5862b7bc96bf23f21c0396ecb`, Build ID
  `e5c457384eeebd057932c444efda77618f137399`.
- The capture contains 1,387 sample records, zero lost events, 1,377 user-filter
  records and 10 kernel-IP records. Every retained dynamic-stack and
  `STACK_USER` payload is 16,384 bytes; every raw callchain count is zero and
  the record size is 16,624 bytes. These facts do not support a claim that the
  payload was omitted or that stack depth caused the original decode failure.
- The decoder is `/usr/bin/perf` version 7.0.14, SHA-256
  `174229fb48b8557aafa821dd828d883dce47c35c31067e1f31e4a01a91c6c98b`, with
  the retained DWARF/libdw capability. The replay event is `instructions:u`;
  the retained sample metadata identifies CPU 31 and PID 3683783.

## Replay result

All three modes saw the same 1,387 headers and total event period
6,935,000,000, with zero lost events.

| mode | empty decoded chains | decoded frames | marker chains |
|---|---:|---:|---:|
| stock | 906 | 1,945 | 0 |
| no-op interposer | 906 | 1,945 | 0 |
| v6 corrected | 0 | 11,154 | 78 |

The stock and no-op output are byte-identical. In the corrected decode, 10,543
frames resolve and 611 frame lines remain unresolved; the unresolved lines are
outside the marker-containing chains. A zero empty-chain result does not mean
that every outer caller frame is available, and no physical stack truncation is
claimed.

The corrected decode recovers these exact qualified symbol tags:

- `doc_retention_probe_0942_cfb_profile::profiled_measured_case`: 78 chains,
  390,000,000 instructions of event period.
- `doc_retention_probe_0942_cfb_profile::measure_one`: 41 chains,
  205,000,000 instructions of event period.
- `doc_retention_probe_0942_cfb_profile::run_case`: present in all 78 recovered
  marker chains, with 390,000,000 instructions of inclusive event period.

The 78 marker chains are the exact marker count. Forty-one include the outer
`measure_one` frame and 37 do not. The retained binary does not resolve
`stage_timed_commit` or `timed_iteration`, so no claim is made for either name.
The counts and period values are inclusive event tags: they overlap, are not
wall-time measurements, and must not be added to form a phase or owner
fraction. The descriptive marker-period share is approximately 5.6236%; it is
not a phase fraction or speedup estimate.

## Correction to the raw-word result

With the corrected `0x64d1e4e5b000` base, the sealed v2 raw `STACK_USER`
scan derives the executable interval from the `PF_X PT_LOAD` range and scans
all 1,387 sample records. The target-leaf subset contains 908 samples with a
total period of 4,540,000,000. Within that subset it reports 53 samples
containing a potential marker-range word, 39 containing a potential
marker-return word, 17 containing a potential `measure_one` return word, and
54 containing a potential `run_case` range word. These values are
address-word matches only. They do not prove that a frame was on a live call
chain, do not establish phase ownership, and do not replace the v6 decoded
ancestry. The prior zero-marker raw conclusion is withdrawn.

## Qualification boundary

The original pilot required at least 100 marker samples. The authoritative v6
replay recovers 78, so the 0971 pilot remains below gate. There is no qualified
12-child cohort, no exact phase owner fraction, no latency result, no speedup
claim, and no production adoption decision. This artifact is a diagnostic
recovery from one retained child profile, not a native performance result.

## Evidence bindings

The machine-readable binding is
`/home/zhuhe/litchi-performance-0972/receipts/diagnosis-v3-receipt.json`.
The authoritative v6 files are:

- Replay receipt: `receipts/dwfl-replay-receipt-v6.json`, SHA-256
  `452ffec504e387b5beea157d755b817a1ce1e282c0a8acc379ddf6bafa429522`.
- Decode summary: `analysis/dwfl-corrected-decode-v6.json`, SHA-256
  `3b75fdabf79ca7104f24e5deb2e389ecbd94c7a439579cc16a5a8d3c7b42e7db`.
- v6 shim source: `tools/dwfl_base_shim_v6.c`, SHA-256
  `391ecf6192425ebb7e05e5ce3de4f4fe9be3195e22b9f05b111dad7de6b05311`.
- v6 shim shared object: `replay-dwfl-v6/dwfl_base_shim_v6.so`, SHA-256
  `de1348268dd283086680973102dccb2e5584db7eb98efc5a87b88fc8a0eec39f`.
- v6 version map: `tools/dwfl_base_shim_v6.map`, SHA-256
  `d0a7914c5cb02770073db94dc4d2b207a95ae513b436c661d8a9c9a206aefef6`.
- Reproduction commands: `commands-dwfl-replay-v6.txt`, SHA-256
  `47c07fef0a6a7d072fd817d10354b5d6d852d457fcad749db2716738ca9de893`.
- Corrected script output: `replay-dwfl-v6/corrected.script-no-inline.txt`,
  2,013,005 bytes, SHA-256
  `fa27bd237b3716cfbdc1ff052882114e77d2228b0caa87f81f54948285c8401d`.
- Corrected decode log: `replay-dwfl-v6/corrected.dwfl.jsonl`, 18,165,996
  bytes, SHA-256
  `433ba37c3a220f35c8c02e46f141c13a34ce11ced50f2b8782b2ca512c0ca6d2`.
- Corrected raw scan: `analysis/corrected-raw-stack-scan-0972-v2.json`, 443,936
  bytes, SHA-256
  `6ef76bc01cc7e298c924eda95cfd29dc0cc070ac3e3de4cbe9d87a7cf1b93a9b`.
- Corrected raw-scan source: `analysis/corrected-raw-stack-scan-0972-v2.py`,
  22,507 bytes, SHA-256
  `b60e2638fc6af5432a3dfa1402721678f6b58f8b3f454ee99611e67fb80f3987`.
- Author seal: `receipts/dwfl-replay-author-seal-v6.json`, SHA-256
  `a8be68fb63e7460d839ca23ab2457b2415762da76d8a92d26e8640281a7bde1c`.

The v5 source and shared object remain retained as historical lineage only;
v6 is the authoritative replay for this record. No production source or
repository evidence files were changed.
