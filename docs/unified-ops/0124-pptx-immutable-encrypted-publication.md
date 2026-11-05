# 0124 — PPTX immutable encrypted publication

## Native ownership and publication

Native `Package::write_reencrypted` publishes an already-opened package through
shared immutable access. It checks graph freshness, refuses pending mutable
authoring, requires retained encryption provenance and uses the existing crypto
implementation with explicit crypto limits. It neither clones the package nor
serializes and reopens an intermediate owner. Existing authoring APIs retain
their behavior.

A fallible writer bounds clear ZIP staging before appending bytes or reserving
memory. Requested capacity growth is clamped to the plaintext limit; allocator
rounding and simultaneous crypto buffers are not claimed as managed memory.
Serialization-failure staging uses a zeroizing buffer. Encryption then takes
ownership through the existing crypto API and follows that API's memory policy.
The caller's sink receives ciphertext only after encryption succeeds. A failing
caller sink can receive a ciphertext prefix and returns its original typed I/O
error, including flush failures; transactional publication remains the atomic save layer's responsibility.

`SaveOptions::reencrypt` now delegates PPTX publication to this native method.
It retains the input encryption profile while accepting an explicit output
password. All three common durability policies and both overwrite policies use
the existing guarded atomic writer, including its final source/execution fence.
Clear owners reject re-encryption as an invalid argument. Ordinary plaintext
publication of encrypted owners remains protected before and after a save.

Signed encrypted packages preserve their inner signature bytes during unchanged
publication. Explicit signature removal still refuses on encrypted owners,
because the native encrypted editor cannot implement that mutation; re-encryption
must not silently ignore an authorized removal. The remaining editing and managed
signature-removal gaps stay visible. Wrapped PPTX crypto failures now delegate
to native classification, preserving invalid-password and resource-limit codes.

## Verification

Native and public regressions cover retained Standard/Agile profiles, byte-exact
clear package preservation, unchanged encryption provenance, plaintext refusal,
authoring/clear-owner rejection, bounded staging, crypto and sink failures,
save-to-origin, eager/auto ingress, all durability policies, no-clobber,
unchanged signed content and encrypted signature-removal refusal. Capability
reports remove the implemented re-encryption gap and the OfficeCLI delta is
regenerated from the report. The full PPTX/encryption facade suite passed 531 tests and doctests; the
native encryption integration suite passed 13 tests. Fourteen delta-generator
tests passed. Final native/public regressions also check sink flush failures: all 13 native
encryption and 10 public write tests passed. All-target native/facade Clippy
passed with warnings denied, as did 17 native doctests (two existing ignores).
The PPTX-without-encryption facade build passed. The PPTX glossary profile
passed with 68 banned names and zero transaction families checked.

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption --test package_encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features pptx,encryption
CARGO_NET_OFFLINE=true cargo check -p litchi-ops --no-default-features --features pptx
CARGO_NET_OFFLINE=true cargo clippy -p litchi-pptx -p litchi-ops --no-default-features --features litchi-ops/pptx,litchi-ops/encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo test -p litchi-pptx --features encryption --doc
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check
python3 -m unittest discover -s tools -p test_officecli_capability_delta.py
```

The full goal remains active. Broader PPTX writes, the remaining format tiers,
bindings and paired performance acceptance remain required. This checkpoint
makes no performance parity claim.

## Commit and cleanup

The rewritten base remains an ancestor of HEAD, so no rebase replay is needed.
Staging failed with exit 128 because Git cannot create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the read-only
Git metadata filesystem. No commit was created. Required staged guards and the
changed/untracked artifact audit reported no forbidden evidence or binaries.
No temporary worktree was created; review was read-only.

After every build/check process terminated, this checkpoint's generated root
`Cargo.lock`, `target`, Python cache and `/tmp/litchi-unified-0124-*` scratch files
were removed. Tracked tool lockfiles and inherited changes were preserved.
