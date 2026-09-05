# Share proven staged media without skipping validation

The 0423 media lifecycle stages eight 2 MiB image payloads during public
planning, then stages the same payloads again when publication reruns
preparation. The retained 0424 trace identifies each eight-allocation,
16,777,216-byte clone stack under the actual lifecycle runner. The experiment
protocols were committed at `5b176dc9a` before production edits.

Prepared images and charts now own independent `Arc<Vec<u8>>` staging storage.
Initial preparation still reads and clones the source payload under its normal
limits. Publication supplies its immutable plan as an optional reuse hint.
Preparation still reads, checks and validates the current source and graph.
Only an exact match of source URI, allocated target URI, content type, declared
size and decoded bytes can share the staged Arc. Byte comparison checks
execution in 64 KiB chunks. A mismatch follows the original checked allocation
path; the existing full `Prepared::matches` comparison remains the stale-plan
decision. The subsequent candidate reread and byte comparison are unchanged.

The OPC topology plan already stored payloads in `Arc<Vec<u8>>`. Its new
lower-level `try_add_part_shared` accepts that ownership directly. The Vec and
shared entry points use the same private checks and fallible vector reservation.
A lazy payload constructor preserves the old Vec entry point's ordering:
URI, count, duplicate and content-type checks and reservation precede Arc
construction. Neither entry point authorizes exact-source passthrough.

The plan does not capture a managed source-cache allocation. It owns independent
staging bytes; `PartData::into_arc` is not used. Callers retaining an OPC Arc
handle cannot mutate the plan's bytes through safe Rust; `Arc::make_mut` detaches
when another owner exists. The original PPTX plan remains borrowed throughout
publication and keeps its reservation alive. Both original and current full
staging reservations and the candidate reread reservation remain charged.
This conservative accounting does not promise a smaller managed-budget
requirement or exact physical heap accounting.

| Accepted ADR constraint | Preservation in this change |
| --- | --- |
| 0001 / 0002 / 0024 public layers and owners | The new sharing method stays in low-level OPC. The ordinary PPTX plan remains opaque; no archive type, cache handle or lock is exposed. |
| 0003 source-checked atomic edits | Full preparation, byte-semantic plan equality, lineage/version checks and candidate verification remain before output. |
| 0005 finite budgets and explicit execution | Existing staging/reread charges remain; comparisons check execution by bounded chunks; no ambient I/O, cache or parallelism is added. |
| 0006 preservation and refusal | Changed sources and unsupported graphs still refuse; destination raw ZIP preservation, exact media/XML and deterministic output remain independent gates. |
| 0010 / 0011 physical OPC ownership | OPC owns topology publication and physical authorization. Shared payload bytes supply no new source or archive authority. |

There is no new dependency, unsafe code, global cache or executor. Normal and
allocator measurements use the unchanged 0423 harness and exact common inputs.
The expected benefit is one fewer selected-payload clone during publication;
repeated source reads, decompression, graph validation and candidate rereads
remain. Region peaks, allocation request volume, live endpoints, whole-process
RSS and normal timing are reported separately. Explicit whole-lifecycle object
drop measurements, near-limit workload profiles and broader native/range/scaling
coverage remain open.
