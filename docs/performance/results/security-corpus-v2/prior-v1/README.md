# Bounded security corpus catalog v1

This additive sidecar catalogs checked-in security and malformed-input fixtures
by exact archive SHA-256. It has 22 corpora, 22 opt-in bindings, and 10
correctness oracles. The default performance catalog remains unchanged at 201
bindings. No row in this sidecar is a timed benchmark, and no result here is a
latency, RSS, or allocator claim.

The catalog hashes ZIP member payloads after bounded decompression. CFB/OLE2
files remain opaque archive bytes in this sidecar; stream-level parsing belongs
to the CFB/XLS security harness. Each row is limited by
`profile.json` (64 MiB input/materialized/output, 4,096 ZIP members, 64 MiB per
member, and 8,192 relationship slots). The validator reads every listed ZIP
member and checks the archive and member identities without opening a host
application.

The opt-in rows cover these guarded operations:

- signed OOXML signature inspection, exact no-op preservation, and refusal of a
  changed source;
- document protection no-op preservation and changed-publication refusal;
- encrypted CFB password-required, wrong-password, and authenticated semantic
  readback checks;
- encrypted OOXML profile and integrity-status readback, including Standard
  unauthenticated profiles and Agile authenticated profiles;
- the narrow third-party zero `EncryptionBlockSize` DataSpaces compatibility
  case, where strict graph inspection still refuses and the explicit reader
  compatibility accepts;
- inert VBA inventory and source-preserving no-op/refusal behavior;
- external relationship inventory without resolving or fetching a target;
- malformed OPC/XLSX and CFB boundary inputs under bounded read policy.

Macro-bearing and external-link files are inventory and preservation inputs.
They are never executed, activated, refreshed, or sent to a network. The
oracles set `host_code_execution`, `external_reference_resolution`, and
`timed` to false for every row.

Provenance is deliberately conservative. Apache POI paths identify a test-data
snapshot, not the application that authored the Office bytes. The encrypted
OOXML paths under `crates/litchi-crypto/tests/data/ooxml/` are component-built,
msoffcrypto-derived, or litchi-authored interoperability vectors; their
`document_producer` is explicitly unknown and they are not native Microsoft
Office evidence. Native LibreOffice QA and Apache POI checkout fixtures remain
separate ignored tests because those external `3rdparty` trees are not tracked
by this candidate. Their presence does not promote any synthetic fixture to a
native-producer claim.

Generate and validate the catalog from the candidate root with:

```sh
python3 tools/generate_security_corpus_catalog.py \
  --revision 0201c716e6005ad26712895db409b807cdbca5b0 \
  --worktree-dirty
python3 tools/validate_security_corpus_catalog.py
```

The Rust security correctness harness is independently opt-in:

```sh
taskset -c 8-31 cargo +1.95.0 test \
  --manifest-path tools/perf-baseline/Cargo.toml \
  --lib security_corpus -- --ignored
```

That harness is not part of the default 201-selector run. Timed corpus runs
are intentionally absent until a separate CPU reservation and measurement
protocol are approved.

`prior-receipts/xls-embedded-payload-v1-receipt.json` retains the exact
historical XLS embedded-payload v1 receipt and its Rust 1.95.0 commands. That
receipt is historical audit evidence: v1 was later blocked for its odd
`cbFmla=5`/missing `PtgTbl` and embed-info wire shape. It is retained alongside
this catalog and must not be read as the corrected v2 result.
