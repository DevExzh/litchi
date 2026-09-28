# 0814 next-step and source review

This review is evidence-only. It does not select or implement a production
candidate. It covers the retained 0813 source at `5eb8629254` and the
current-source native-attribution packet being prepared in this directory.
iWork is outside scope.

## What the current evidence establishes

The 0813 change is retained: direct matching of `Reader::read_event()` with
`Ok(Event::...)` and an explicit error arm removed the three pre-dispatch event
payload-copy chains in both ordinary and profile release assembly. The public
matrix covered six generated shapes and eighteen capture/commit/lifecycle
workflow rows; six rows met the frozen benefit rule. Large capture improved by
10.544% and large lifecycle by 7.269%. The paired allocation calls, allocated
bytes, net live bytes, and peak-above-entry values were equal in all 144
comparisons. This is useful evidence for the one retained codec change and
those six generated shapes; it does not establish broad producer or format
coverage.

The attempt in 0812 to reconstruct the 0811 historical frame-pointer binary
failed its exact hash check, so the 0811 offsets were correctly refused. 0812
then captured the actual rebuilt executable and validly joined its fresh
samples to that executable's own DWARF and assembly, localizing the old
`map_err` copy chain. No historical offsets were relabeled or reused. The next
attribution must use the exact 0813 current-source binaries while they exist;
historical samples and timing values must remain outside any new aggregate.

The current scanner path is visible at `notes/codec.rs`: MCE processing is
followed by `scan_processed_xml`, which performs the bounded `Reader` event
loop, namespace push/pop, root/depth/node checks, and `inspect_element`. The
inspector resolves the element and every checked attribute, unescapes every
non-namespace attribute value, and charges attribute limits before recording
the value. Only values in the relationship namespace are retained in its
relationship, notes-master, and slide-ID vectors. The opened-presentation
capture also performs slide identity/name checks, package
fingerprinting, MCE retention, and notes-graph loading. The source already
contains the accepted slide-root and payload-digest memos; a new memo is not
an assumed explanation for the remaining scan cost.

## Static review of the 0814 assembly driver

`assembly.py` has the intended custody checks, and it fails closed when its
short inspector selector is ambiguous.

* It requires each immutable `control`, `profile`, and `fp` executable from
  the build receipt and checks its artifact digest before reading it.
* It runs one `nm -S --defined-only` result per binary, requires exactly one
  `18scan_processed_xml17h` and one `15inspect_element17h` symbol, and records
  the exact address, size, type, and mangled symbol line for each match. The
  inspector suffix is ambiguous in this binary because OPC and PPTX both
  export it, so the original run correctly stopped before completing the
  control inspector.
* It invokes `objdump` with the recorded symbol range, demangles only for
  inspection, and retains one bounded assembly artifact per symbol. The
  receipt binds every row to the build binary and source descriptors.
* The aggregate validator requires the six rows in deterministic variant/name
  order, checks every instruction address against its recorded symbol range,
  and requires assembly to start after the native, perf, and decode lanes.
  Cleanup then verifies all three executable identities before removing the
  owned target.

The recovery driver applies a fully qualified PPTX mangled prefix and the
exact symbol/range checks, removing the 0812 reconstruction failure mode for
the retained packet. The failed suffix attempt remains recorded beside the
recovered rows. Assembly does not itself prove a runtime cost, phase share, or
semantic equivalence. No workload or profiler command was run for this
review.

The pre-build metadata refresh is now verified: `README.md`, `plan.json`,
`origin.json`, `inheritance.json`, `build.py`, cleanup, and validation use the
0814 schema, base `5eb8629254fbf88009b988734af59eb90f8c6271`, target
`litchi-target-0814`, and the retained 0813 source/probe references. The
quality-reuse witness binds the sealed 0813 after source and its 1,241 passed,
zero failed, three ignored, 85-suite production result. The explicit 0811
driver reference in `inheritance.json` is historical lineage only; its policy
is limited to reusing the current-production custody/build/probe contract.
The recorded build freeze kept these identities and the 35 architecture and
three unrelated-file witnesses unchanged through capture, decode, assembly,
independent review, and cleanup. This is packet custody, not a production
change.

## Terminal 0814 result

The fresh packet answers the attribution gate without authorizing a source
change. The main analysis, independent raw audit, and exact-offset replay all
pass: 56 reports and 1,820 samples, with 54 native reports at thirty samples
per process and two large-only perf reports at 100 samples each. Every source,
output, semantic, raw-perf, gzip, and exact-binary identity check passes. The
0814 timing summaries contain no historical timing pool and the packet has no
allocation, Callgrind, cross-format, concurrency, or adoption lane.

The native p50 wrapper/control ratios are 1.002120 for tiny, 1.012834 for
medium, and 1.040429 for large; the fp/wrapper ratios are 1.005104, 0.999433,
and 1.006262. The large wrapper interval is [1.034036, 1.044374], so the
profiling wrapper materially perturbs that shape. The fp interval crosses one.
These are build perturbations. They cannot be read as a source speedup or an
ordinary-build phase share.

The sampled stack lane is large-only. Exact-owner samples are 851 and 853 out
of 2,867 and 2,851 whole-process samples. Scanner self leaves are 150 and 142;
hardware SHA is 91 and 84; `Reader::read_event_impl` is 69 and 67;
`inspect_element` is 39 and 59; checked-attribute iterator leaves are 25 and
14. Scanner and inspector inclusive counts are 729/716 and 239/284, so they
overlap and must not be summed. Lost-event lines are zero in both reports;
qualified stacks with an unknown interior are retained at 2 and 1. Unresolved
samples and frame occurrences remain in the reports, so this is descriptive
attribution rather than a phase-fraction claim. Tiny and medium have timing
controls and exact assembly only; they have no sampled-stack counts.

The exact fp scanner assembly maps the two dominant scanner leaves to
`movups 0x18(%rax),%xmm1` at offsets `0x25b` and `0x33b`, with 66/62 and 59/52
samples across the two repeats. The ordinary and profiling-wrapper scanners
have the analogous two arm-local moves. The old three pre-dispatch copy chain
is absent. The offset reader confirms instruction boundaries and conserves
851/853 exact-owner samples as 150/142 scanner leaves, 39/59 inspector
leaves, and 662/652 other-owner leaves. Sampling skid and the frame-pointer build still prevent causal cycle
claims.

Root preserved the first assembly driver and its failed suffix-selector
attempt: the short inspector suffix matched both the OPC and PPTX symbols and
stopped after the control scanner. The recovery driver selected the fully
qualified PPTX mangled prefixes, regenerated all six bounded ranges from the
unchanged control/profile/fp binaries, and retained the partial failed files.
`assembly-recovery.json` records the original and recovered driver hashes,
binary witnesses, and `workload_rerun: false`; the recovery and offset
conservation validators pass. This repairs packet attribution without
re-running a workload.

The source therefore supports one narrow hypothesis. In `scan_processed_xml`,
the `Start` and `Empty` arms use the event only through `&BytesStart` calls to
`resolver.push` and `inspect_element`, and neither reference escapes the arm.
Changing those patterns to `Ok(Event::Start(ref element))` and
`Ok(Event::Empty(ref element))`, with the existing calls borrowing `element`,
could remove the two owned 32-byte arm-local payload moves while preserving
event drop at arm exit. Resolver push/pop timing, node/depth/attribute limits,
typed error mapping, and the `Err(error)` arm remain in the same source order.
This is a bounded candidate hypothesis, not a semantic or timing result; the
next batch must prove the assembly change and run the existing differential,
refusal, and public workflow gates.

## What the fresh profiles leave unresolved

The evidence sets the following limits before another source edit is
considered:

1. The scanner remains a material large exact-owner leaf, but the two moves may
   be event-arm lowering or another compiler representation. Only a fresh
   ordinary release assembly can establish whether `ref` changes them.
2. `inspect_element`, checked attributes, resolver work, and package
   fingerprinting remain separate mechanisms. Nested stack counts and the
   hardware SHA leaf cannot be converted into removable scanner fractions.
3. Unknown/unresolved frames and owner selection remain visible evidence. Any
   future packet must retain them and refuse a phase-share claim when they
   prevent one.
4. A source candidate still needs semantic differential/refusal coverage,
   resource-limit coverage, exact ordinary/profile/fp assembly, and the full
   eighteen-row public workflow comparison. A static instruction reduction is
   insufficient.

The terminal packet is therefore an attribution gate, not an adoption gate.

## Bounded next investigations

The decision after the profile should follow the evidence:

* **Continue the PPTX scanner line for one bounded candidate only if** the
  current exact-owner samples still show a material scanner/inspector leaf and
  the assembly identifies a local mechanism whose error ordering and resource
  checks can remain unchanged. The candidate must be private to the PPTX
  owner. Namespace discovery and checked-attribute handling must remain
  separate unless a new proof reproduces their different duplicate and refusal
  ordering; the 0812 review explicitly rejects an unproved fusion.
* **Close this scanner line if** the scanner is no longer a material leaf, the
  only apparent saving is a build/profile artifact, or the next change would
  require changing namespace timing, checked-attribute semantics, MCE
  ownership, or refusal precedence. Static instruction reductions without a
  public workflow benefit are insufficient.
* **Prioritize the broader measured queue next** when the scanner gate closes.
  The current hotspot inventory names PPTX per-capture notes-graph validation,
  XLSX commit/save deflate and its remaining audit, DOCX commit audits and
  multi-window proofs, the CFB structural reparse map clearing, legacy DOC
  output zero-fill, and the incompressible OPC mutation save. The XLSX deflate
  path must use the explicit execution-context budgets; no hidden global pool
  is acceptable. Each choice still needs current-source attribution rather than
  relying on the older wave-wide ordering.

There is also a program-level reason to stop iterating on generated PPTX
fixtures indefinitely. The goal still lacks representative native-producer,
cold-cache, caller-supplied range-source, non-seek output, concurrent/scaling,
and broad CRUD evidence. The CRUD coverage record explicitly says that the
0813 matrix adds no capability or cross-format coverage. After this attribution
round, a high-ROI cross-format measurement should be preferred unless a fresh
profile supplies a stronger PPTX mechanism and a bounded end-to-end question.

## Architecture and semantic limits

Any follow-up remains subject to the accepted architecture:

* ADRs 0001, 0002, 0010, 0011, and 0024 keep the change in its format owner,
  preserve dependency direction, and prohibit physical archive types or raw
  locks in ordinary CRUD APIs.
* ADR 0003 requires immutable snapshots, exact source/revision checks, atomic
  publication, and reversible edits. ADR 0006 requires preservation,
  deterministic refusal and error ordering, validation without mutation, and
  fail-closed handling of unsupported or malformed content.
* ADR 0005 and ADR 0031 require measured representative evidence, finite
  resource accounting, explicit execution context for parallel work, and no
  ambient I/O or hidden worker pool. A count of removed instructions or
  allocations is not an end-to-end claim.
* ADR 0008 requires current-source, reproducible, independently auditable
  evidence before retention. ADR 0030 keeps lazy OPC decoding behind fallible
  accessors. ADR 0032 already admits the bounded slide-root memo under strict
  ownership and reservation conditions; it does not authorize an unbounded
  parsed-value cache.

The broad non-iWork performance goal therefore remains open. Continued PPTX
scanner attribution is justified as a final bounded evidence step because the
scanner was repeatedly measured and 0813 produced a retained workflow gain.
It is not sufficient evidence to defer the unresolved OLE2/OOXML coverage and
source/I/O/scaling work after this step.
