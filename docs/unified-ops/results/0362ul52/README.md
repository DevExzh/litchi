# 0362ul52 — bounded crypto raw codecs

This lane makes the raw OOXML crypto, DataSpaces, and sensitivity-label APIs
require an explicit resource profile. Its original base is integration
`267ea715266ffc3bd754bfbfa4896b139adad9fc`. No native or profiler run was
performed; the bounded Cargo validation packet is recorded below.

## API contract

* `litchi_crypto::ooxml::{open, inspect, load, encrypt, rekey}` are the
  canonical profile-taking entry points. The duplicate `*_with` wrappers were
  removed; integrity-policy variants remain explicit `*_with_policy` APIs.
* `litchi_crypto::spaces::{inspect, inspect_bytes, parse_*, write_*}` require
  `&spaces::Limits`, covering per-stream, aggregate, entry, component, string,
  XML-depth/node, and output ceilings.
* `litchi_crypto::labels::{parse, write}` require `&labels::Limits`, covering
  XML bytes, labels, extensions, depth/nodes, attributes, and output.
* DataSpaces `Snapshot` and `Patch` retain the validated profile. OOXML's
  StrongEncryption container derives its DataSpaces profile from the caller's
  OOXML encryption profile; inert IRM licenses and sensitivity-label XML remain
  source data and are never activated.
* `spaces::Error` and `labels::Error` report typed invalid-profile,
  limit-exceeded, and allocation outcomes. OOXML maps those outcomes to its
  typed container limits, DOC maps them to its core resource/format errors,
  and ops preserves the corresponding limit, invalid-argument, or internal
  classification.
* DataSpaces directory walks visit bounded entry references, source capture
  checks declared stream and aggregate sizes before materializing payloads,
  and `Patch::write_to` bounds both stream admission and the final CFB sink.

## Focused source tests

The existing crypto unit suites retain their round-trip, native-compatibility,
refusal, and inert-payload coverage. New focused checks cover inclusive exact
stream/output limits, undersized stream/output rejection, entry-count limits,
aggregate DataSpaces bytes, and nested inert XML depth for both DataSpaces and
LabelInfo. The checks also cover typed source/output limits and the OOXML
container mapping. Test call sites and DOCX/PPTX/XLSX encryption consumers were migrated
to the canonical signatures in the same change.

The DOC package's CFB owner retains the `OpenOptions::data_spaces_limits`
profile on `raw::Reader`; both `data_spaces` and `data_spaces_snapshot` pass
that profile to `spaces::inspect` and `Snapshot::from_ole`. Focused DOC
coverage proves that a selected stream ceiling affects both projections.

## Atomic handoff and validation packet

The reviewed source checkpoint was `3c843b9e9f18592565f79ee0254ab6c01e9ef839`.
The historical handoff source is `1a3e88336a001bfe52df4a794358f59ffa55ab9f`,
with a clean worktree and the same reviewed production API. The delta after the reviewed
checkpoint is `21f9be6fd647637d3dfe82fd1853e3420118485f` (all remaining DOCX,
PPTX, and unified-ops encryption consumers) followed by
`1a3e88336a001bfe52df4a794358f59ffa55ab9f` (the corresponding test-lint
qualification). That historical integration commit has parent
`267ea715266ffc3bd754bfbfa4896b139adad9fc`.

All receipts below were produced by the approved runner
`1904177040fc7c523582433028d9bba56efc30fe0d9d96beb554a2fab343691e` with the
approved Cargo wrapper
`1b22a69fd625d57565041bf6c634d8ebd1a82be8b3d0006390472c86324230d8`, Rust
1.95.0, offline and locked resolution, one build job, zero incremental/debug
build settings, and the 20 GB target guard. The bound workspace lock is
`Cargo.lock` SHA-256
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`.

Final-source receipts, all with exit 0 and source equality to the validated
source checkpoint:

* Ops all-target check: `(deleted-coord-artifact)/00-ops-consumer-check-final.receipt.json`, receipt SHA-256 `ed789000549c02559185a132e12ddb36750ffff9b978e699be86d83838e33d65`, trace SHA-256 `8646683c8ba5f2b6c91039c69a3b9ff38d1874f23720bb1b684b1957b261f1b0`, target 668,917,628 bytes.
* Ops all-target clippy: `(deleted-coord-artifact)/01-ops-consumer-clippy-final.receipt.json`, receipt SHA-256 `c24832c7244eb302d0d0da516638d7c54d23ad64726e52499279bf15810e16b8`, trace SHA-256 `ca7b2bc9b32adda478694b6a90110f3da59b37190b0bb10d2d1fd42c6a9f843c`, target 988,678,185 bytes.
* DOCX `read_limits` encryption test: `(deleted-coord-artifact)/00-docx-read-limits-encryption.receipt.json`, receipt SHA-256 `7296e2223604e70b636634b39f7605ccb18c6f4d349dc062e8590e411007ae60`, trace SHA-256 `7c7bce78d7c588f43c0c5fc3c00d1c518365d8e87ea6191f54015819f4125a8c`, target 557,131,740 bytes.
* DOCX `source_backed_managed` encryption test: `(deleted-coord-artifact)/01-docx-source-backed-managed-encryption.receipt.json`, receipt SHA-256 `546380287ac90e398585fd3721136f84b029d69bdfe71a9d7a6ac43988b1b2d0`, trace SHA-256 `211f20611dbd42575e9a68ff71a4229ca2dedf788b998b505d9a4d61bd9be8e9`, target 578,576,124 bytes.
* PPTX `package_encryption` test: `(deleted-coord-artifact)/02-pptx-package-encryption.receipt.json`, receipt SHA-256 `f3661273312c61d450f58a2b37107a2e8215edaa5b6d19f018fcbda3b7b95b9a`, trace SHA-256 `870dad0c9ac3441f0a02cfa682de5437759c578700d538144e9dc421952bb1b4`, target 701,498,395 bytes.

The post-review unlock-profile repair is validated separately from the
historical handoff above. Commit `354e38969991af8395b613ea4ace62971643a279`
passes the selected OOXML encryption profile from `unlock_ooxml_with_origin`
to the DOCX, PPTX, and XLSX owning constructors; its PPTX `from_decrypted`
route now receives the same owning options shape as the other two formats.
Commit `106e48ea98d65bd2a7acd9565308b638e767b54f` adds the DOCX classifier
delegation that preserves the crypto layer's typed output-limit error.

* Unified unlock-to-reencrypt regression, source `106e48ea98d65bd2a7acd9565308b638e767b54f`: `(deleted-coord-artifact)/receipt.json`, receipt SHA-256 `9456df26c187172cfcb045d91698ace179312797b96ec4bf4a8e2590a8185e43`, log SHA-256 `5b34dbfc24cb268b13494c4826d5255667023f5c04294cf64d2c5a082dad91a9`, trace SHA-256 `f86277d3a1d62aea47a9a253b678b30c55419e43f75d9f2be1be4adeb8b890ef`, target 2,598,241,308 bytes, exit 0. The test exercises DOCX, PPTX, and XLSX with `max_output_bytes: 512`; every post-exit ownership sample is empty.
* Affected all-target check, source `37626e75f05774f253eb2849db5fe61ae56b4057`: `(deleted-coord-artifact)/affected-consumer-check-receipt.json`, receipt SHA-256 `53c2dcaeec332135149da1c962a4ed19ecd834967d63f465ce40b0c8435c1b83`, trace SHA-256 `26fb1b4985df3a9100b004091b543ddaa4c9518d08de1460a025f33b8f723261`, target 710,864,977 bytes, exit 0.
* Affected all-target clippy (`-D warnings`), source `37626e75f05774f253eb2849db5fe61ae56b4057`: `(deleted-coord-artifact)/affected-consumer-clippy-receipt.json`, receipt SHA-256 `da467e05aa9617d7556970b700532e6a838ebead3a757d09838a4be149d9fe48`, trace SHA-256 `09bcbed8933d7dcf51ec2dde0b2d457fd06d14467b29dddf5cbdd2cb8b1d60fb`, target 1,032,412,773 bytes, exit 0. Both affected commands use the approved runner, offline/locked resolution, one job, and the 20 GB target guard; both ended with no owned processes.

The first focused attempt is retained as a diagnostic failure rather than
silently replaced: `(deleted-coord-artifact)/receipt.json`, receipt
SHA-256 `7ce627afd644c14951521c10f21c3a300c5de6e012c90f4b23920d755fed7130`,
failed because the test supplied an already-created empty destination and the
facade correctly classified it as a corrupt existing package. The destination
fixture was changed to a non-existent extension-bearing path before the passing
rerun. The `37626e75f0` validation commit only fixes three unnecessary
qualification warnings in the existing XLSX encryption tests; it does not
change production behavior.

After the final atomic squash, the external source-equality proof is
`(deleted-coord-artifact)`; the corresponding
post-exit ownership proof is
`(deleted-coord-artifact)`. Their SHA-256 values
are supplied with the coordinator handoff because the proof files remain
outside Git with the bounded build artifacts.

The earlier reviewed checkpoint retains its passing crypto and DOC evidence:

* Crypto tests passed at `3c843b9e9f18592565f79ee0254ab6c01e9ef839`; log `(deleted-coord-artifact)/01-crypto-tests-attempt-clean-1.log`, SHA-256 `866dedc28dcee7077458064f1feaea53c90c7964041143a4d61b3754ced0e596`.
* Crypto clippy passed; log `(deleted-coord-artifact)/02-crypto-clippy-attempt-clean-2.log`, SHA-256 `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`.
* DOC tests passed; log `(deleted-coord-artifact)/03-doc-tests-attempt-clean-3.log`, SHA-256 `448587635df67ea2d23333d290444e1e231e22c7d5f0947f57f06381cb7d00e2`.
* DOC clippy passed; log `(deleted-coord-artifact)/04-doc-clippy-attempt-clean-4.log`, SHA-256 `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`.

The full ops test attempt is preserved honestly as resource-bound, not as a
code pass: `(deleted-coord-artifact)/receipt.json` (receipt
SHA-256 `94310a0135d90b448dfe77b5c0a4c94a15b9d06b1fe0423965d2b9f909e63551`)
is bound to `21f9be6fd647637d3dfe82fd1853e3420118485f`, was terminated with
exit `-15` by the 19 GB live guard at 19,101,829,380 sampled bytes, and ended
at 18,863,904,748 target bytes with no owned processes remaining. The earlier
`3c843` full-ops attempt remains at
`(deleted-coord-artifact)/05-ops-tests.receipt.json`; it failed before
linking on the then-stale consumer calls (log SHA-256
`22986bdda62145f206473a465179b9c93e1884df2a28bd98f24d1c2f3c336a27`).

Every final receipt's post-exit owned-process sample is `owned=[]`. The
external proof is `(deleted-coord-artifact)/crypto-raw-1a3e88336a-owned-process-proof.txt`, SHA-256
`3b6ccc0087edf2798e0f1bd37195e8d1c09b8753793431f2ea0157788c7ddbdf`; slot 4
is released because no process owned by this lane remains.
