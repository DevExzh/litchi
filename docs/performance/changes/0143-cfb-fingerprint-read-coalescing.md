# Change 0143: CFB fingerprint read coalescing

Date: 2026-08-15

Status: accepted for the named CFB/OLE2 same-length atomic-save case. Both
balanced release directions improve p50, p95 and mean in warm and
`cold-requested` states. The exact logical request-count reduction is accepted;
logical `ReadAt` calls are not physical-I/O calls.

## Change and retained invariants

Commit `3ef285be2efd4d4c443380674c58738e2c04dcec` changes only the private
`litchi-cfb` overlay implementation and its tests. Complete source-fingerprint
passes now use a fallibly allocated, right-sized window capped at 1 MiB. The
64 KiB comparison and publication windows are unchanged and are explicitly
dropped before a fingerprint window is allocated. Zero- and small-length
sources therefore allocate no more than their source length.

No fingerprint stage was removed. Planning still brackets validation with two
complete source/target fingerprints. Direct `write_to` still performs its
initial preflight, 64 KiB emission hash and post-emission preflight. Atomic
`save` still performs initial preflight, emission hashing and the final
post-flush/fsync pre-rename preflight. Source version/stable-token checks,
malicious positional-source refusal, typed partial-output progress, complete
candidate reopen, selected-stream readback, sibling-temp cleanup and atomic
destination replacement are unchanged.

The final current-tree gates were:

- 18/18 focused overlay tests;
- 174 CFB unit tests and six integration tests;
- strict all-target Clippy with warnings and deprecations denied;
- rustdoc with warnings denied, rustfmt and diff checks;
- strict library Clippy across `litchi-ole-common`, `litchi-xls`, `litchi-doc`
  and `litchi-ppt`;
- two independent final reviews, including a fix that releases the publication
  buffer before post-emission fingerprinting.

The broad dependent test run linked and passed the DOC/PPT/OLE and almost all
XLS targets, then hit the unrelated reproducible
`xls_writer_encryption::all_profiles_round_trip_and_emit_exact_filepass_families`
assertion (`225` versus `66`). That test does not call the changed overlay
planning/publication functions. Its isolated rerun failed identically; it is
recorded as an external test blocker rather than hidden as a green gate.

## Clean balanced release evidence

The control is the phase-attribution commit
`3c23ab40dd9a2fd01875ef501f9723771cf06a7e`; its clean release binary is
39,382,224 bytes with SHA-256
`c56d190f6a70c2bfcce4969f50268c89ccdbb6cd1954b0f8b0a35c0006c286e0`.
The candidate clean release binary is 39,399,400 bytes with SHA-256
`8cafbfacb40346be66951e63c7ed6583e41d49966e9043953af7ffd2fc8e9f0c`.
Both reports record the expected revision and `git_worktree_dirty: false`.

The CPU-2 order was `A1 control, B1 candidate, B2 candidate, A2 control`, with
20 warm-ups and 200 fresh-child samples for each warm and advisory-cold state
in every leg:

```sh
taskset -c 2 litchi-perf-baseline \
  --warmup 20 --samples 200 \
  --filesystem-cache warm,cold-requested \
  --case cfb_file_same_length_overlay_atomic_save \
  --json <leg>.json
```

An initial set of runs was rejected before analysis because the candidate was
built with the dirty root checkout's manifest directory embedded. Rebuilding
from the clean detached candidate checkout made the one-sample probe and both
candidate legs report `git_worktree_dirty: false`. Rejected captures are not
present in the committed artifact.

The deterministic five-entry incompressible CFB is 16,913,408 bytes with
SHA-256
`7ffbd37c3e472a21b382bcbb02e430a62164e58d2270bbee0deaa584ff47a94d`.
Every one of the 1,600 retained samples changed one 36-byte MiniFAT stream,
reported one physical span, published 16,913,408 bytes and produced SHA-256
`7994759e1b2e3e520c0f0df5efb1586e34c6bc0f5744a7f4b989733cfd2830fc`.

| Leg / state | p50 | p95 | p99 | Mean |
|---|---:|---:|---:|---:|
| A1 control warm | 105,840,020 ns | 110,751,696 ns | 114,574,340 ns | 106,308,180 ns |
| B1 candidate warm | 102,312,742 ns | 107,400,487 ns | 109,860,349 ns | 102,487,478 ns |
| B2 candidate warm | 102,893,024 ns | 107,029,869 ns | 109,584,416 ns | 103,189,001 ns |
| A2 control warm | 104,265,459 ns | 108,791,793 ns | 110,557,158 ns | 104,337,590 ns |
| A1 control cold-requested | 118,451,198 ns | 128,508,710 ns | 401,252,054 ns | 129,446,493 ns |
| B1 candidate cold-requested | 105,696,527 ns | 110,631,594 ns | 112,239,841 ns | 105,737,844 ns |
| B2 candidate cold-requested | 105,497,592 ns | 109,781,819 ns | 110,992,895 ns | 105,927,738 ns |
| A2 control cold-requested | 116,525,718 ns | 120,683,765 ns | 124,292,564 ns | 116,627,519 ns |

Candidate improvement, computed separately in each direction, is:

| Direction / state | p50 | p95 | Mean |
|---|---:|---:|---:|
| A1 -> B1 warm | 3.3327% | 3.0259% | 3.5940% |
| B2 -> A2 warm | 1.3163% | 1.6195% | 1.1008% |
| A1 -> B1 cold-requested | 10.7679% | 13.9112% | 18.3154% |
| B2 -> A2 cold-requested | 9.4641% | 9.0335% | 9.1743% |

`cold-requested` means the kernel accepted `posix_fadvise(DONTNEED)` advice;
it is not proof that the storage device served physically cold reads.

## Exact work and resource boundary

All samples in a variant have the same logical source work:

| Interval | Control calls | Candidate calls | Logical bytes in both |
|---|---:|---:|---:|
| open | 264 | 264 | 135,680 |
| plan and candidate validation | 784 | 300 | 33,962,596 |
| atomic publication | 777 | 293 | 50,740,224 |
| **operation** | **1,825** | **857** | **84,838,500** |

The candidate removes 968 logical requests, or 53.0411%, without changing a
logical byte. In both ABBA directions, the advisory-cold plan p50 falls by
20.27%/20.61%; planning contains one coalesced complete fingerprint and atomic
publication contains the final coalesced pre-rename fingerprint. Phase
percentiles remain independently computed and are not additive.

The code-local fingerprint window increases from 65,536 to at most 1,048,576
bytes, a 983,040-byte increase for this corpus. It never overlaps the 65,536-byte
publication window. A clean `/usr/bin/time -v` ABBA boundary (three warm-ups,
30 warm samples per leg) recorded whole-process maximum RSS of 111,640 and
111,508 KiB for control and 111,508/111,508 KiB for candidate. This shows no
visible process-level RSS regression in the named run; it is not an
operation-only peak-allocation measurement. Heaptrack on the ordinary harness
coordinator did not follow the exec'd timed child and was therefore rejected
as allocation evidence.

## Artifacts and claim boundary

The [compact summary](../results/cfb-fingerprint-abba-0143-summary.json) contains
all leg statistics, phase medians, exact counters, resource bounds, binary
identities and hashes; its SHA-256 is
`aeea4b8cc9e50a4a55b0c3311fa7bcd1d91e8eb86d73d6c33792394a9cd35e68`.
The [compressed raw artifact](../results/cfb-fingerprint-abba-0143.json.zst)
has SHA-256
`f3956ea9d52ffa24867704afb8f4c22bcf0c2480b1308777039e2309eb2e946e`.
It expands to 53,189,198 bytes with SHA-256
`018617c68ddb65037e376f17e1c88cebe7bf2eaec266fbac4fd426b1e4083357`.

This accepts one generic CFB/OLE2 same-length atomic-save optimization. It does
not establish physical I/O, CPU-cycle, decompression/recompression,
operation-only allocation, guaranteed-cold-storage, high-latency remote source,
or broad DOC/XLS/PPT semantic CRUD performance.
