# 0730 probe design: bounded DOC validated-render handoff

This packet's measurement harness is deliberately fixed before the production
handoff candidate is edited. It is a paired baseline/candidate probe, not a
new format oracle and not a cache benchmark.

## Scope

The measured operation is the public `litchi_doc::body_text` lifecycle:

1. open the source fixture into a public DOC snapshot;
2. construct an edit and replace paragraph 0 with the fixed 45-UTF-16-unit
   string `litchi copy-through baseline replacement text`;
3. commit the edit; and
4. copy the committed snapshot bytes into the returned output vector.

The timing binary measures this whole lifecycle with `Instant`, and the
allocator binary measures the same lifecycle with a process-local counting
system allocator. The output remains live after the measured region so the
post-region validation sees the exact bytes returned by the lifecycle.

The accepted matrix has two DOC fixtures, `FloatingPictures.doc` and
`NoHeadFoot.doc`, using their 0728 paths and source identities. The copied
harness still exposes the inherited PPT and common-container operations so
its oracle implementation stays identical to the qualified 0728 probe, but
0730 qualification invokes only `--operation format` for the two DOC cases.
No iWork input or candidate-only API appears in this probe.

## Interval boundaries

The expected output is made before warmups and samples. It is an independent
public edit used as the semantic and byte identity oracle; its work is outside
all timing and allocation regions. For timing, `timed_format` starts its clock
before `public_format_edit` and stops after the returned `Vec` is available.
Oracle inventory, CFB validation, semantic reopening, and negative controls run
after that clock stops. Warmups run the same operation but are dropped without
being reported.

For allocation, `allocation_format` opens the allocation region immediately
before `public_format_edit` and closes it on return. The returned vector is
stored outside the region and retained through validation. `peak_live_bytes`
is relative to live bytes at region entry; `retained_bytes` is the live-byte
delta at region exit. These are allocator-region measures, not RSS and not a
sum of independent phase peaks.

## Oracle contract

The inherited 0728 build-5 oracle remains the acceptance contract. It checks:

* complete, structurally valid source, expected, and measured CFB files;
* exact source/expected/output stream paths and expected output stream bytes;
* unchanged source stream bytes and an explicit allowed changed-stream set;
* root and storage CLSID preservation and semantic directory metadata;
* raw directory images after the documented allocation-field normalization;
* a logical length-changing stream proof plus the DOC-specific UTF-16 length
  and paragraph witness; and
* direct paragraph, story, table-cell, field, revision, and embedded-object
  collection comparisons where the public projection is available, with an
  explicit unavailable state where it is not.

The expected identity is generated with the same public operation as the
measured route. Every sample must match it byte-for-byte at stream level and
pass the semantic witness. The inherited controls deliberately mutate or
remove data and must remain rejected. Their outputs are evidence that the
oracle cannot accept a structurally valid but semantically wrong handoff.

## Candidate comparison

The probe is compiled unchanged against the baseline and candidate workspace
states. A candidate may add a bounded validated-render handoff internally,
but this probe never calls that API directly. Root qualification compares
the two runs by source and fixture identities, expected-output SHA-256,
replacement digest, direct semantic witness, inventory, oracle controls, and
the captured output. Any identity drift or oracle failure is a gate failure,
regardless of elapsed time or allocator counters.

The probe reports one whole lifecycle timing value for the public DOC route and
one whole allocation region. It does not attribute callback costs, split
phases, retained-state policy, or a production speedup. Those are separate
design and review questions; this harness only supplies the paired observable
needed to evaluate the bounded pilot.
