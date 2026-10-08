# 0977 combined-profile qualification packet

**Status: sealed bounded failed-pilot evidence packet.** The authoritative
manifest is [`packet-manifest-0977-v1.json`](packet-manifest-0977-v1.json);
its file list excludes the manifest itself and binds every copied file.

This packet records the frozen `docfloat/e1/default` protocol, the harmless
control, and the registered Office launch through its fail-closed prefix. The
plan requested twelve serial children with five warmups and 80 measured calls
per child. The independent control audit and pre-capture Office readiness
review passed their bounded scopes. The control completed with 4,633 samples,
raw loss zero, terminal `LOST_SAMPLES` zero, and sampled parent CPU 30 and
target CPU 31 observations. Those observations establish control custody and
affinity evidence only.

The Office launch returned 2 after child 2 failed the retained raw
`STACK_USER` scanner. Child 1 passed its local gate with 946 marker samples,
empty-callchain count zero, raw loss zero, and terminal loss zero. Child 2
returned zero natively and had terminal loss zero, but decoder admission
failed. Children 3--12 were never launched; there was no retry or replacement.
The reducer was invoked once against the failed manifest, rejected it with
return code 2, and produced no reduction output, partial cohort, or synthetic
substitution. Therefore this packet contains no twelve-child Office
qualification, performance, attribution, causal, production, or adoption
claim.

The final child-2 diagnostic and independent failed-cohort audit are the
authoritative explanations. One retained `PERF_RECORD_SAMPLE` is 72 bytes,
with type 9, kernel misc `0x1`, CPU 31, IP `0xffffffff920000a7`,
`callchain_nr=0`, and user stack size zero. Its raw perf-data offset is
`0x26f5b0`; the retained perf text gives stack offset `0x40`. The other 7,200
sample records are 33,008 bytes. This is an unsupported variable sample shape
for the frozen fixed-offset scanner. The audit observed no stack-bound
violation and makes no PMU-skid or other causal inference. The diagnostic is
offline evidence only and did not rerun the workload, build, or mutate the
decoder.

## Copied evidence

The packet includes the sealed v2 plan and helper receipt, design and
pre-capture reviews, control and Office terminal metadata, bounded child
summaries, offline requalification and synthetic/loader-gate receipts, the
fail-closed reducer receipt, the final child-2 diagnostic, the final
independent audit, and their replay helpers. The v9 integration is an offline
synthetic duplicate-packet custody/arithmetic check; it is not a native
Office cohort. The current-perf requalification is offline over retained
inputs and preserves the `Invalid HEADER_EVENT_DESC` warning descriptively;
`evlist -v` remains the event authority.

Large raw captures, mmap/task dumps, retained perf inputs, decoded logs,
synthetic packets, target executables, shared objects, and other compiled
artifacts remain outside the repository. Their paths, sizes, SHA-256 values,
roles, and regeneration references are recorded in the sealed
[`external-artifact-descriptor-final-v1.json`](provenance/external-artifact-descriptor-final-v1.json).
The descriptor’s repository list excludes itself; the packet manifest binds
that descriptor as a copied file.

Every copied file is below 1 MiB. No `binaries/` directory, compiled
executable, shared object, `perf.data`, or mmap dump is copied. The cleanup
receipts record removal of only unbound temporary cache/process artifacts;
the bound source, raw evidence, and replay inputs were preserved.

## Reproduction and custody

[`provenance/reproduction-v1.md`](provenance/reproduction-v1.md) contains
bounded offline checks and the historical Office launch argv. Commands shown
there are recipes derived from sealed receipts. Output-producing helpers must
use a fresh external root or versioned output paths because they refuse to
overwrite existing sealed outputs. Replaying them is not a request to rerun
the failed native cohort, and no command in this packet authorizes a decoder
change or adoption decision.

The authoritative final records are:

- `analysis/cpu-review-control-audit-0977-v2.json`
- `analysis/cpu-review-office-readiness-0977-v2.json`
- `analysis/cpu-review-office-failed-cohort-0977-v3.json`
- `receipts/child02-rawstack-diagnostic-0977-v3.json`
- `analysis/reduce-office-failed-0977-v1.json` and
  `receipts/reduce-office-failed-0977-v1.json`
- `receipts/temporary-cleanup-0977-v1.json` and
  `receipts/temporary-cleanup-0977-supplement-v1.json`
