# Change 0474: PPTX fresh streaming operation memory

`performance_claim: none; descriptive fresh-creation memory baseline`

`claim_authorized: false`

The opt-in `pptx_streaming_create` selector measures the existing public
StreamingPresentationWriter over 8, 256 and 8,192 slides. Earlier small PPTX
creation cases build a materialized Package and output Vec, which cannot
establish the streaming writer's allocation behavior. This harness addition
brings selectable cases to 441 while retaining the 37-case default matrix.
Production authoring and the allocator observer are unchanged. Implementation
revision: `60a1a4300a2844f372c0c916b5be69c50ce38668`.

The frozen protocol runs two reversed repeats of normal and allocator targets,
three shapes, fresh processes on CPU 2, one worker, three warmups and thirty
samples. Normal and allocator timings remain separate. There is no production
before/after or new registered latency claim.

## Measured results

All twelve formal lanes pass: 360 samples, plus two excluded one-sample pilots.

| Slides | Members | PPTX bytes | Normal p50 ms R1 / R2 | Operation peak above entry |
| ---: | ---: | ---: | ---: | ---: |
| 8 | 53 | 36,259 | 1.035935 / 1.032350 | 435,541 bytes |
| 256 | 549 | 274,398 | 8.565925 / 8.493936 | 681,659 bytes |
| 8,192 | 16,421 | 7,940,406 | 259.277908 / 258.604302 | 8,875,092 bytes |

Every one of the sixty allocator samples per shape has the listed incremental
peak, zero live-byte change at exit and zero failed allocation calls. The large
operation peak is 20.377 times the tiny peak across 1,024 times as many slides.
This rejects constant total operation memory for the tested public path.

Requested allocation calls are 954 / 10,145 / 310,993, requested bytes are
21,964,902 / 227,629,482 / 6,809,604,013, and reallocations are
138 / 1,385 / 48,271 per operation. These quantities are identical in all
allocator samples of each shape. Allocation work and retained peak are distinct:
the large operation requests about 6.81 GB over its lifetime while peaking at
about 8.88 MB above entry. The current writer starts a new Deflate encoder for
every member; stack attribution and a matched test are needed before claiming
which owner causes the allocation work or retaining an optimization.

All normal mean/p50/p95/p99 repeat changes are below one percent in absolute
value (maximum 0.9624%, tiny p95). Timings remain descriptive; this is not a
production speedup comparison or registered latency claim. The full vectors,
sample order, dispersion and Student-t mean intervals are retained and checked.
Normal whole-process RSS is 82,604–82,732 KiB; allocator captures are
82,660–82,736 KiB. Nearly flat process RSS does not contradict the measured
operation heap growth, because these scopes differ and RSS includes setup.

The emitted maximum slide XML lengths are 889 / 889 / 890 bytes under the
same 16,614-byte per-slide policy ceiling. That ceiling is not retained heap.
Presentation target-part lengths are 848 / 8,944 / 285,476 bytes.

A materialized preflight verifies exact ordered 37+2N package members, every
slide's text/shape count/geometry and layout ownership, presentation references,
and typed master/layout/notes relationships. It releases the artifact, physical
reader, semantic package and target XML before returning the scalar/digest corpus.
The public notes graph caps presentations at 4,096 slides; the largest shape
uses typed OPC ownership checks while smaller shapes also invoke that API.
This is not a general validator for all secondary/stale relationship records,
full notes features or native Office round trips.

The timed operation constructs finite limits, generates one deterministic UTF-8
text box per slide, runs the public writer and finalizes/destroys its ZIP/OPC
state. It validates the writer's slide and accepted-text counters. Output goes
to a non-seek hashing discard sink, then its complete digest/length is checked
against preflight outside timing. Sink construction/digest extraction and
observer endpoint reads are outside the timer. `authored_part_bytes` is the
preflight identity of ppt/presentation.xml, not a total slide-XML measurement.

The max_slide_xml_bytes limit is a scalar emitted-byte policy counter, not a
retained-memory reservation. The active slide writes XML fragments directly;
formatting integers uses a 20-byte stack buffer. Total operation heap also
contains active Deflate state, OPC folded/original name and ancestor maps, ZIP
normalized-name validation, FileHeader vectors and raw member-name storage.
Those structures remain live until package finalization and grow with members.
Zero retained sink output cannot establish constant total memory.

Allocator values describe callback-order logical requested heap, including
other process threads between endpoints. They exclude allocator-internal
reallocation overlap and physical copies. Whole-process RSS includes untimed
setup/reopen and cannot isolate writer memory causality. Source inspection alone
does not attribute measured peak bytes to any particular map/vector.

See the [evidence bundle](../results/change-0474/README.md) for source/build/capture
bindings, deterministic corpus identities, full sample vectors, strict report
arithmetic and portable replay. Its source review records accepted ADR scope
and the valid-layout-retarget oracle fix. The initial formatting failure and
subsequent exact-source final gates are retained.

Fresh creation is separate from logical append to an existing structure,
package Part addition, and modification followed by repackaging. Native
breadth, source variants, worker scaling and the full non-iWork goal remain
open. A measured metadata-growth result must guide the next transport work;
it cannot be renamed an explicit-window total-memory bound.

Validation passes 368 release harness/allocator-target tests (one ignored),
548 PPTX unit tests, 13 streaming integration tests and 11 Python evidence
tests. Final formatting, warning-denied Clippy and rustdoc pass against all
7,034 authenticated source hashes. Crate boundaries and the strict ten-claim registry also pass. Independent
source and evidence review confirm the scoped oracles and exact summary replay.
Final portable results are retained with the completed bundle.

Sealed live verification, post-cleanup fresh-copy portable replay and output
mutation rejection pass. Temporary binaries/checkouts/bytecode/copies are
removed; shared Cargo caches and both user-owned files are preserved. The
frozen protocol and all twelve formal captures remain unchanged.
