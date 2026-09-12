# XLSX publication instruction attribution

All four source-bound publication profiles passed their output/lifecycle
oracles. Production and harness source are unchanged. The profile owner is
`SourceBackedEditor::publish_multi_commit_to_stream`; it excludes destruction
of the returned `MultiSnapshot`, which the native publication timer includes.
These instruction shares must not be converted into latency predictions.

| Shape / repeat | Publication Ir | Overlay XML audit | Physical writer | Attribute inspection (nested in audit) |
| --- | ---: | ---: | ---: | ---: |
| medium / 1 | 95,539,471 | 53.8239% | 45.4031% | 11.0877% |
| dense-sparse / 1 | 172,376,919 | 57.5571% | 41.8823% | 11.8983% |
| medium / 2 | 95,536,593 | 53.8232% | 45.4039% | 11.0876% |
| dense-sparse / 2 | 172,377,186 | 57.5571% | 41.8816% | 11.8984% |

The aggregate publication owner is 535,830,169 Ir. Its OPC topology-write
child dominates. Within that child, `validate_overlay_xml` accounts for
301,274,428 Ir (56.2257% of publication) and the physical writer accounts for
231,144,888 Ir (43.1377%). These are disjoint immediate children. Attribute
inspection is 62,205,865 Ir (11.6093%) and is nested within the audit; adding it
to the audit share would double count. Detailed raw-edge/self equations are in
`publication-analysis.json` and the retained annotations.

The audit delegates through `verify_authored` to `verify_with_policy` in
`crates/xml-minifier/src/audit.rs`. Its major work is XML event reading and
attribute inspection. The physical writer primarily prepares regenerated
entries and writes prepared local records, including preserved source bytes.
The profile does not establish that compression alone dominates publication.

Each fresh child has six numbered owner dumps: one exact no-op lifecycle
publication, then foreign-source refusal and normal publication for each of
clear and remove, followed by the unique measured owner call. All five earlier
dumps have positive incoming edges from `run_xlsx_cell_value_lifecycle_gates`;
the sixth has exactly one positive edge and one call from
`run_xlsx_cell_values_edit_save`. Every selected summary equals that incoming
edge and the owner's annotation total. The program-termination dump is zero.
The vendor-extension partial-sink check is not part of these two shapes.

The demangled method name covers five compiled writer specializations. The
symbol observation is retained, and positive raw caller edges identify the
measured specialization. Call metadata elsewhere can include collection-off
work; it is not an allocation count. All raw dumps are retained. Valgrind's
brk-segment-overflow note remains in stderr; every child exits successfully,
produces its complete report and passes the captured oracles. This is not an
allocator/OOM assurance or native hardware-counter result.

A local analysis correction was required: the retained display-name helper
split Rust slice names at their `[T]` rather than at the final executable
suffix. Only the 0528 adapter changed. Two focused regression tests now pass,
and every selected child cost is checked against raw Callgrind edges. No raw
capture or sealed prior helper was changed.

The next useful candidate is within the shared authored-XML auditor, preserving
both original and replacement validation. Both slice and streaming paths create a fresh checked attribute iterator per
tag; streaming reuses event/capture buffers, not duplicate-key scratch. A
resettable duplicate checker would be new work, not reuse of an existing
auditor checker. The profile
shows an attribute-inspection owner, not a removable-allocation count; a fresh
allocator/native baseline and exact error/limit differential tests are needed
before adopting any implementation. Duplicate-before-malformed-value precedence must match the existing checked
iterator; a simple unchecked-iterator post-pass is not automatically safe.

Do not remove OPC audits based on XLSX semantic readback: they enforce different
contracts, including compact authored XML, whitespace policy and limits.
Source/version, cancellation, atomic pre-output failures and partial-output
classification must remain. Planning also consumes about one third of the
prior native total and remains a separate unprofiled opportunity. No production
change or speedup is adopted here. OLE2/OOXML remains first; ODF is deferred.
