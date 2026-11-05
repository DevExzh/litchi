# 0095 — XLSX native creation and publication

This begins the XLSX write lifecycle required by M2. `create(Xlsx)` adopts
`Workbook::create` directly, without serializing and reopening the package.
The unified `to_bytes`, `write_to`, `save`, `save_as` and `save_with` routes
now publish XLSX through its retained owner. Default creation is an ordinary
Transitional workbook. Templates, Strict conformance and initial metadata
remain explicit option-level refusals listed in capabilities.

Source owners stream the retained OPC source artifact without promotion or
worksheet decoding. Eager owners delegate to the native workbook writer.
Source freshness brackets publication and is checked before atomic replacement.
Filesystem publication uses the existing guarded OPC atomic writer, including
no-clobber and durability policies. An existing leaf execution context is
forwarded when present. Public `OpenOptions::execution` still promises only
open-boundary cancellation (record 0008); these opens do not yet construct
budget-managed owners. This checkpoint does not claim managed publication
through that option. Path opens
retain their stable origin across eager/source routes and encryption unlock.
Byte-opened and created handles require an explicit destination.

Signed publication preserves the signature infrastructure; explicit `unsign`
uses the native removal path. Encrypted owners refuse ordinary plaintext
publication; explicit re-encryption uses the native encryption profile and
stages ciphertext through the atomic writer. This does not add encryption or
signature processing to the facade.

## Remaining work

XLSX cell edits, structural verbs, transactions and batch coalescing remain
pending. Managed open/publication requires a separate leaf policy integration
covering path, bytes and prepared-package routes. The existing editor seam's shared-owner eager conversion needs exact
source preservation and a bounded, memoized promotion before cell writes are
integrated. The narrow source cell editor does not provide an immutable owning
result and must not be used as a serialize/reopen workaround. DOCX text-output
parity, the other format write waves, Python, CLI and full performance and
acceptance gates remain open. No performance improvement is claimed here.

## Validation

Generated workbook fixtures exercise source/eager publication, exact no-op
bytes and unrelated member retention, default creation/reopen, path origins,
atomic overwrite refusal/replacement, changed-source refusal, signature
preservation/removal and encryption policy. Capability goldens and the generated
OfficeCLI delta expose creation option gaps and remaining mutations.

Encryption tests exposed a preexisting XLSX error-classification gap: ordinary
output refusing encryption removal fell through to `corrupt_file`. The facade
now retains that typed leaf cause and reports `document_protected` with an
explicit re-encryption suggestion. Native encryption errors use the existing
crypto classifier. Eager encrypted path fallback also retains its origin even
when encryption hides the inner format until unlock.

Checks include:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features xlsx,encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,encryption
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features docx,xlsx,encryption --all-targets -- -D warnings
python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check
python3 -m unittest discover -s tools -p test_officecli_capability_delta.py
```

Results: 418 minimal-feature tests/doctests, 558 XLSX+encryption, 814 combined
DOCX/XLSX+encryption, and 14 delta-generator tests passed. Final capability
goldens cover XLSX-only and combined encrypted configurations. Clippy passed
with warnings denied. The read-only review found the managed-open limitation
described above; it remains an explicit capability gap rather than a support
claim. No native crate or iWork source changed.

The rewritten base remains an ancestor. Staged evidence/binary checks print
nothing. Commit staging remains blocked by read-only shared Git metadata
(`index.lock`). Temporary logs, generated root lockfile and build artifacts
are cleaned after checks; no temporary worktree was created.
