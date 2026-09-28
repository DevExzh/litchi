# 0817 source and scope review

Production and the tracked performance harness are unchanged at base
`953866d382`. Root revalidated all 35 previously read normative documents;
their hashes match 0816. No accepted ADR is amended or reinterpreted.

The reviewed path is `tools/perf-baseline/src/ordinary_save.rs`: `build_corpus`
opens each caller-selected real archive, derives the first sheet/A1 or first
admitted slide/shape target, freezes the edit outcome, and checks two fresh
publications and repeated saves for deterministic output. `run_case` separates
lifecycle, semantic edit, atomic publication and counting publication scopes.
Owner destruction and digest verification occur outside the measured region.
Default atomic save includes file and parent-directory synchronization.
Counting publication for PPTX calls `to_bytes`, so it is not evidence of a
streaming PPTX serializer or bounded streaming memory.

DOCX appends the harness marker in an ordinary paragraph. XLSX replaces first
sheet A1 with that marker; PPTX replaces text at the first admitted position.
A typed refusal remains an outcome and the harness may then publish an unedited
package. Therefore report determinism and a successful process exit are not
sufficient edit admission. The independent artifact audit must bind source
bytes, actual target, output bytes and preservation before the timing matrix.
Its structural XML comparison is an additional local oracle, not a fresh
LibreOffice or Microsoft Office certification.

The exporter uses the same corpus/editor/publication implementations and
retains source plus default/full/file-only/no-sync/stream outputs. The extra
three generated corpora are exporter controls, not native timing cases. All
five outputs per corpus must match its reference; the default/full policies
must retain identical durability semantics. Reduced-durability artifacts are
untimed controls and do not substitute for the default native save.

The ordinary timing binary has no allocator or procfs instrumentation. The
observer binary enables both existing features. Allocation regions bracket
the timed operation before owner destruction; retained live bytes are not a
leak measurement. Procfs windows overlap probe activity and retain 32 adjacent
empty controls without subtraction. Procfs CPU ticks may quantize short
operations to zero; read/write accounting is not physical device attribution.
The external time wrapper reports whole-child RSS, including setup and corpus
qualification, not an operation-local allocation peak.

The harness lock is a distinct existing tracked dependency graph. Its exact
external differences from the workspace lock are preserved. Fresh harness
checks use that graph, and this record will not pool old timing or claim a
production speedup. Source, locks, normative inputs, corpus bytes, binary
identities and unrelated workspace files are checked around driver work.
