# 0915 retained-binary DOCX section-layout profile

This packet records a diagnostic Callgrind profile of the retained 0914
release binary. It identifies an implementable source-local hypothesis:
`ensure_source_section_inventory_xml` clones `quick_xml`'s namespace resolver
for each XML event. The profile supports investigating a borrowed resolver;
it does not adopt an optimization or make a speedup claim.

The frozen executable is
`/home/zhuhe/litchi-performance-0915/frozen/litchi-perf-baseline`, 69,940,744
bytes, SHA-256
`fe56e5192c1198e831583e870beb5e1ba4731fb2077faaffdfc46044bb0ed2db`.
The frozen protocol is
`/home/zhuhe/litchi-performance-0915/frozen/0914-protocol-v1.json`, 53,363
bytes, SHA-256
`88bca0028d1b658053a31bd9a395f6c282e53ed0bd219cbeda01c21ee4acdfd7`.
The source archive is
`/home/zhuhe/litchi-performance-0915/frozen/source-HEAD-code.tar`,
171,120,640 bytes, SHA-256
`7b77f1875adcdc3c822351b11dcf621f4ebff501fe501cab80fd4fb3f9a59571`, at
source revision `a4ca85e28cdea151176534b7f8dc912253a1ca1e`.

The retained source member
`litchi/crates/litchi-docx/src/source_backed.rs` and its frozen extracted copy
both place the clone at line 2701:

```rust
let resolver = reader.resolver().clone();
```

The extracted copy used for the line proof is
`/home/zhuhe/litchi-performance-0915/frozen/litchi/crates/litchi-docx/src/source_backed.rs`,
145,276 bytes, SHA-256
`f12638008f7742a9e0af7295414d0acd54b79823bebdd459d9e1f2a5cf6e127e`.

The capture used CPU 12, zero warmups, one harness sample, and Callgrind's
`Ir` event. Its command was:

```sh
taskset -c 12 setarch x86_64 -R valgrind --tool=callgrind --quiet \
  --callgrind-out-file=/home/zhuhe/litchi-performance-0915/callgrind/query-dump2.out \
  --separate-callers=2 --collect-jumps=yes \
  --zero-before='litchi_docx::source_backed::Package::main_section_layout_snapshot' \
  --dump-after='litchi_docx::source_backed::Package::main_section_layout_snapshot' \
  /home/zhuhe/litchi-performance-0915/frozen/litchi-perf-baseline \
  --warmup 0 --samples 1 \
  --case docx_source_backed_existing_section_layout_edit_save \
  --json /home/zhuhe/litchi-performance-0915/raw/query-dump2.json
```

The trigger bounds each selected dump to one invocation of
`main_section_layout_snapshot`. Part `.16` is reached through the public
`section_layout_snapshot` query. Part `.17` is reached through
`publish_section_layout_commit_to_stream`, so it is the current-snapshot
subcall inside publication. Part `.17` excludes patch application, output
writing, and output hashing from the whole publication phase. Parts `.1-.15`
are setup/control calls and `.18-.19` are post-timer verification or inverse
publication; they are excluded from the selected-owner attribution.

The caller records in the retained raw files prove the scope. In `.16`, the
profile contains the chain
`main_section_layout_snapshot'...section_layout_snapshot'...run_docx...`,
with one `Snapshot::from_shared_xml` call, one
`ensure_source_section_inventory_xml` call, and 2,140,918 Ir for the latter.
In `.17`, the chain is
`main_section_layout_snapshot'...publish_section_layout_commit_to_stream'...run_docx...`,
again with one snapshot and one ensure call, with 2,141,278 Ir for ensure.
The exact dump sizes, hashes, caller values, and repeat values are in
`profile-reduced.json`.

The selected owner totals are 118,144,825 Ir for the query snapshot and
117,957,908 Ir for the publication current-snapshot subcall. Within those
selected snapshots, `SourceArtifact::fingerprint` consumes 94.45% and 94.60%
respectively; software guest `sha2::sha256::compress256` consumes 92.60% and
92.75%. The resolver validation function is 1.81% and 1.82% respectively.
These percentages describe the selected snapshot owners, not whole
publication-phase latency.

Callgrind associates the inlined resolver clone site with 1,838 observed
loop/process-event iterations in the publication dump. It records 3,676
clone-associated allocation calls (143,364 Ir), 3,674 deallocation calls
(124,916 Ir), and 1,838 memcpy calls (55,132 Ir). The repeat dump reproduces
these counts. They are profiler call records, not XML byte counts,
allocation-byte counts, or native cycles. `NamespaceResolver` in quick-xml
0.42.0 derives `Clone` over a `String` buffer and `Vec` bindings, while
`resolve_event` borrows `self`; this makes borrowing a reviewable source-local
hypothesis without changing resolver state or introducing a cache.

The host denied hardware perf access (`perf_event_paranoid=4`). Callgrind's
`Ir` values are Valgrind guest instruction simulation, and the dominant SHA
owner is software guest execution. This packet therefore makes no native SHA
cost claim and no end-to-end performance claim.

The retained external artifacts and their custody hashes are listed in
`profile-reduced.json`. Replay the custody and caller proof with:

```sh
python3 docs/performance/results/change-0915/profile-replay.py
```

To regenerate a new diagnostic capture, reuse the command above with new
`--callgrind-out-file` and `--json` paths under
`/home/zhuhe/litchi-performance-0915`; never overwrite `frozen/*` or the
retained dump files.
