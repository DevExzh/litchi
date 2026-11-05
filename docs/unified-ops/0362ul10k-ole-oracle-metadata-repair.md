# OLE oracle metadata and payload proof repair

Both actual instrumented release endpoints build with the pinned role locks, `--offline --locked`, one job and the 20G cap. Their targets end at 373,194,780 and 375,466,326 bytes. Both allocator oracles execute successfully, produce eighteen artifacts each and emit all six required allocator rows. Common correctness artifacts match exactly across endpoints.

Receipt validation then rejects the real metadata. The runner and independent validator expect a 19-byte replacement, but `ole-w5-replacement` is eighteen bytes. Coordinator review also finds that both Python consumers expect `ole-w5_edit_target.bin`, while both native probes use `ole_w5_edit_target.bin`. The synthetic fixture copied both mistakes and therefore failed to expose them.

This change migrates both Python consumers and their fixture together. It also checks each stream's recorded size against its payload, so a synthetic inventory cannot silently pair an eighteen-byte payload with a nineteen-byte size. The regression test accepts the real target/extent and rejects the wrong extent, wrong target and a mismatched inventory size in both validators.

The [original allocator diagnostic](results/0362ul10k/allocator-diagnostic.json), 49,094 bytes, SHA-256 `0657212732c9fd2f5f9fb56cd9ecd014de6d05552dfeaae885686579a2efefca`, retains the failed metadata check unchanged. Its external original is `/home/zhuhe/.cache/litchi-api-contract/ole-owner-allocator-29f/diagnostic-receipt.json`. Exact source, lock, probe, binary and command identities remain there; no binary is committed.

After repair, both proof validators accept all four retained normal/allocator endpoint outputs. Seventy-two artifact hashes are rechecked, both six-route allocator parsers pass conservation and 129-package/128-patch history checks, and all sixty-six combined OLE, legacy and shared runner tests pass. The [repair receipt](results/0362ul10k/repaired-proof-validation.json) binds the exact tested source files and external log. No Cargo rebuild is needed for these Python-only consumer repairs; the native outputs remain unchanged.

Independent review confirms the finite consumer change, all eighty allocator receipt identities, all actual artifact hashes and the conservation/history checks. It accepts the source/protocol correctness evidence.

This establishes correctness and proof parsing, not final performance acceptance. Allocated calls/bytes do not increase and peak bytes are lower in these six observations, while peak live allocation counts increase by ten for XLS cells/staging and by 129 for XLS history; retained XLS history also holds 129 more live allocations. Those observations remain visible for final qualification review. Clone execution, calibration, quiet paired timing and full allocator qualification remain required. Nothing is pushed.
