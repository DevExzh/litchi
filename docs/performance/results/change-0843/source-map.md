# Current-source publication boundaries

Read-only independent source audit at 03fa423d73. These are possible attribution
boundaries, not measured costs. Absence from a symbol table can mean inlining;
line/source/stack evidence must establish any attribution.

| Boundary | Current owner |
| --- | --- |
| Public selected operation | source_backed.rs:9872 write_topology_to_stream |
| Topology validation/effective graph | source_backed.rs:7989 prepare_topology; 8028 prepare_topology_state |
| Physical folded-name lookup | source_backed.rs:11560 build_physical_member_lookup |
| Metadata/limit scans, two calls | source_backed.rs:10959 validate_topology_limits; callers 8667 and 9400 |
| Effective Part vector | source_backed.rs:9579–9719 |
| Prepared dispatch | source_backed.rs:2745 PreparedTopology::publish_to_stream; 2816 publish_prepared_topology |
| Shared preservation publication | source_backed.rs:11721 write_changed_overlays_with_appended_inner |
| Raw preservation index | soapberry-zip office.rs:3769 preservation_index_with_limits → preserve.rs:594 new_with_limits_and_policy |
| Index metadata ownership/validation | preserve.rs:695–876; validate_local_span:2654 |
| OPC lookup/count/action vectors | source_backed.rs:11768–11935 |
| Low-level action checks | preserve.rs:1084 prepare_with |
| Offset/ZIP64 layout proof | preserve.rs:1196 validate_output_layout |
| Local/central/tail emission | preserve.rs:1006 write_prepared; 2690 write_prepared_local; 1461 write_prepared_central; 1545 write_prepared_tail |
| One changed Deflate member | preserve.rs:2457 generated_entry; compressor 2539–2564 |
| Untouched local spans | preserve.rs:2766 copy_range |

The publish clock excludes from_vec/source catalog construction and
SourceTopologyPlan::try_replace_part. Lifecycle includes them. Both clocks
include prepare_topology and its full applicable validation, plus publication.
The replacement-only topology path skips content-types/relationship XML
regeneration because Part names/types and relationships are unchanged and the
open-time source catalog already validated them. It still builds physical and
canonical name lookups, validates limits and constructs effective_parts.

The unmanaged from_vec fixture has no write session or transfer sources, so
publication uses the serial index.write_to branch and keeps SourceCheckedSink.
No explicit-budget, parallel, borrowed-transfer or SourceXmlPart case is tested.

Future changes must preserve source-freshness fences, exact archive-end checks,
case-folded ambiguity and exactly-one-target checks, signed/encrypted/non-Part
refusals, central/local consistency, overlap/truncation and ZIP64 validation,
complete action coverage, and full output-layout proof before sink writes.
Default durability is outside this in-memory publication diagnostic.
