# OLE paired native correctness

The normal release baseline at `0ce4bb8b` and assembled candidate at `29f072f19c53503395164fcb3a618c44fc22cf75` build successfully with the pinned role-specific locks, `--offline --locked`, one Cargo job and the 20G cap. Their targets end at 373,070,878 and 375,340,930 bytes. Both actual correctness-oracle executions pass and produce exactly eighteen artifacts each; all eighteen match byte-for-byte across endpoints.

The oracle verifies OLE target-only replacement and retained stream inventory, DOC and XLS expected edits, no-op output, retained histories, properties and opaque XLS bytes. Independent review checks 66 unique source, lock, probe, binary, log and artifact identities, endpoint cleanliness, commands and proof contents. Its 146 checks pass. The coordinator also rehashes every oracle artifact and confirms both sets match.

The [native diagnostic receipt](results/0362ul10j/native-correctness.json) is 19,308 bytes, SHA-256 `c1d21a96a7f68a3b8d6933be92ebea22db065e2f412faa9565a49016503e0de9`; the external original is `/home/zhuhe/.cache/litchi-api-contract/ole-owner-native-29f/diagnostic-receipt.json`. It records exact build/oracle commands, binaries and evidence identities. Preparation remains external at `/home/zhuhe/.cache/litchi-api-contract/ole-owner-preparation-29f.json`, 59,007 bytes, SHA-256 `1fada236dd406750c20a6930e442963793dcebaf1532d5e07ae9ee89cd4ec02b`.

The clone and timing probes were built but were not executed. The normal oracle's allocator marker does not constitute allocator qualification. Calibration, allocator evidence and quiet paired performance qualification remain required. Nothing is pushed.
