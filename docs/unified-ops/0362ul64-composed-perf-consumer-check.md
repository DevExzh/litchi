# 0362ul64 — composed standalone performance consumer check

The standalone performance harness compiles on clean source
`805c08facf9ff7c92c7f05b66bde1e2bdb8524f2`, including the XLS,
SpreadsheetDrawing and DOCX migrations and the six-transform historical
DrawingML adapter. This closes the compiler check queued in records
0362ul46 and 0362ul42. The receipt predates the later PPT raw-record merge,
crypto retention repair, font dependency bridge and new PPTX raw returns;
it does not validate those changes.

The source is a new detached capture at
`.coord/final-composed-perf-805c08/source`. Historical RTF sources remain
unchanged. All 1,962 tracked test-data files were hydrated before the check.
The command, from `tools/perf-baseline`, was:

```sh
cargo check --offline --locked --all-targets
```

It passes with exit 0 in 175.3352015018463 seconds under pinned Rust 1.95.0,
one build job, debug information and incremental compilation disabled, the
`20G` process cap and the actual f5 ownership/live-target guards. The target
ends at 368,735,177 bytes; all 10,406 guard samples remain within the limit
and no guard error occurs. Terminal owned-process lists are empty. Root
independently verifies the receipt, command log and Cargo trace, unchanged
locks, source identity and clean source state. Compiler runtime is not a
native timing measurement.

The external receipt is
`.coord/final-composed-perf-805c08/perf-baseline-explicit-check/receipt.json`
(17,408 bytes), SHA-256
`1c034d333f5db83c95e0e61251c7915116c25aa2a1d4ce57d911ca2f7b611a77`.
The command log is `00-perf-baseline-check-attempt-805-explicit.log`
(40,217 bytes), SHA-256
`f0b9c4c6f6f6f8e3200a20f22a605343c1937f97e61420d4d3a13a073748fb7b`.
The Cargo trace is 118 bytes, SHA-256
`25a851bf24468d12147a25878cdb6a6333e519afad343d56c77b64a692793fc5`.

The root lock remains
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`;
the standalone performance lock remains
`eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c`.
The approved runner SHA-256 is
`1904177040fc7c523582433028d9bba56efc30fe0d9d96beb554a2fab343691e`,
and its adjacent Cargo shim is
`495e24c6f60b238378386d86ac4158cba904992ce8870a9494f2c4a0eedd256d`.

To reproduce, create a fresh checkout at the source commit, materialize its
tracked inputs, restore these preserved locks and invoke the recorded command
through the same guarded runner and environment. Keep the resulting log,
trace and receipt outside the repository. Slot 3 is released after the
verified terminal snapshots; actual performance qualification and the final
integrated gate remain required.
