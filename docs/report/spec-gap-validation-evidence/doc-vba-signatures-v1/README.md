# DOC VBA signature metadata validation

This batch adds a bounded deferred reader for Word `StwUser` signature variables (`Sign`, `SigAgile`, `SigV3`) and DOC/common property-set APIs for PIDDSI `DigitalSignature`. The shared blob owner already existed at the recorded base. The reader preserves exact nested bytes and caches both successful reads and deferred errors. The property-set transaction preserves exact no-ops and supports source-checked inverse restoration.

The golden wire test verifies one generic VT_BLOB size field followed directly by the raw DigSigBlob. These APIs leave PKCS#7, ASN.1 content, cryptographic validity, certificate trust and macro execution opaque; they do not write StwUser or activate a VBA project.

Fresh dedicated-target Rust 1.95 gates passed 1,451 unit/integration tests and 14 doctests across DOC and OLE common, with three test and twelve doctest ignores. All-target/all-feature warning-denied Clippy, warning-denied rustdoc, formatting and whitespace checks passed. Independent final local-spec review cleared the fourteen source paths. These are synthetic structural/preservation checks, not native producer or trust acceptance evidence.

`receipt.json` binds source hashes, commands and log hashes. To reproduce from the feature commit, restore the supplied workspace lock with `gzip -dc docs/report/spec-gap-validation-evidence/doc-vba-signatures-v1/Cargo.lock.gz > Cargo.lock`, then run the recorded commands with an empty dedicated Cargo target directory. The absolute target and temporary-directory locations in the receipt identify the captured run and may be replaced locally. Earlier shared-target exploratory logs are superseded by this dedicated-target run.
