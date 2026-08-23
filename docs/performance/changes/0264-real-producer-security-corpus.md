# Change 0264: real-producer security correctness corpus

## Status

Landed in `be46f6bf6b6491f451167a9adda6ec3fbcfa1c52`. This is an ignored,
locked CI correctness gate over checked-in producer artifacts. It adds no
selector, default-record, latency, allocation, RSS, physical-I/O, or resource
speedup claim.

## Scope

The `real_producer_security_corpus` test in
`tools/perf-baseline/src/security_corpus.rs` covers eight pinned public
fixtures:

| Fixture | Repository path | Source SHA-256 |
|---|---|---|
| POI Office 2010 signed DOCX | `test-data/poi/test-data/xmldsign/ms-office-2010-signed.docx` | `bc55c0362722818823a6dd95f8e0ca9869e179ace972a0915241feb4677bde5f` |
| POI Office 2010 signed XLSX | `test-data/poi/test-data/xmldsign/ms-office-2010-signed.xlsx` | `4cbd8cbe613f036b7a0c779ffaaec7c5838710896c6ef26b3f27410d25d5ce45` |
| POI Office 2010 signed PPTX | `test-data/poi/test-data/xmldsign/ms-office-2010-signed.pptx` | `4d925d282dcca86e62b6716647a458246f8b9ea0eae0ec6664bbbf5a3f91bce1` |
| Read-only protected DOCX | `test-data/ooxml/docx/documentProtection_readonly_no_password.docx` | `5d4c919f2e06b84fbe35cfaaa4012e8f469b811e1f643deb3e660b798bfe4544` |
| POI password CryptoAPI DOC | `test-data/poi/test-data/document/password_password_cryptoapi.doc` | `f2d0dc59ad7ec2356695ad5dc550057052a4017d5f1eb46e887297f5089896fb` |
| POI password binary-RC4 DOC | `test-data/poi/test-data/document/password_tika_binaryrc4.doc` | `9231e724bb17a2e5f74815728d90b06e15684cf5fb2443a6fa24deebd33be952` |
| POI `SimpleMacro.xls` | `test-data/poi/test-data/spreadsheet/SimpleMacro.xls` | `0e92c9bb018abd8a5f9121d65827c9e3bd280777219cb77a2efd70635143c00a` |
| OOXML startup external-link XLSX | `test-data/ooxml/xlsx/external-link-path-startup.xlsx` | `e06155747da482bfb7c1ac5f0ab3a80cbe5b510e664926709c356ba6b59e9bc4` |

The test remains outside the performance selector and is run explicitly with:

```sh
cargo test --locked --manifest-path tools/perf-baseline/Cargo.toml \
  --lib security_corpus -- --ignored
```

## Correctness and security gates

Each signed OOXML package must report valid integrity and signature status.
An exact no-op publication must reproduce the source bytes, while a signed
mutation must refuse with `SignedSourceRequiresExplicitPolicy` before writing
any output. The protected DOCX must have an exact no-op and must refuse a
semantic edit with `UnsafeEdit`, also with zero output.

The two encrypted DOC fixtures must distinguish `PasswordRequired` from
`InvalidPassword`. The accepted passwords must produce non-empty semantic
text with the pinned SHA-256 digests
`6dd4273bea0a8f70f4b6d8448e0ea1cb22b54713ad78d177394bd8b496e0aea6`
(CryptoAPI) and
`5c5c945257fcd1569b5161722e15a9d73283daf786aca1daf04d0029d3736b78`
(binary RC4), and a second correct-password read must reproduce each digest.

The macro fixture is inspected as inert data only: VBA metadata, project
storage, structural completeness, and a module containing `Sub ` are checked;
code is never run or authored. An exact no-op preserves the digest of every
`_VBA_PROJECT_CUR` CFB stream, and a source-backed row insertion refuses with
the existing typed `UnsupportedFeature` error without changing the source.
The external-link fixture is reduced to an exact relationship inventory with
digest `16a2466d394b25d4a465c4db740fca842b9627e98a91c37163312b9585b80beb`;
the target is not resolved or fetched, and no `personal.xls` member is
present in the package.

The bounded-ingress probe rejects an input limit one byte below the fixture
size. A managed output budget rejects publication with a typed
`OutputBytes` resource-limit error and an empty sink; after the artifact,
package, and source are dropped, retained `Memory`, `Objects`, and
`OutputBytes` charges must be zero. Input and work are cumulative counters and
are not required to return to zero.

## Boundaries

This is a bounded admission, no-op, refusal, decryption-readback, inert-VBA,
external-inventory, and RAII-release matrix for these eight checked-in
artifacts. It does not provide signature creation or authorized signed
mutation, protected-document authoring, encrypted publication, VBA execution
or authoring, external-link resolution, arbitrary producer breadth, repair, or
general security scanning. It is correctness-only evidence; no latency,
allocation, RSS, physical-I/O, or resource-performance claim follows, and the
current selectable matrix and default 36-case / 198-record matrix are
unchanged.
