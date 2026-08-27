# Change 0312: OPC compact preservation provenance

Status: implemented and validated

`performance_claim: none`

## Compact provenance model

Preservation provenance now stores compact, exact canonical relationship
baselines instead of repeating the recognized-member name in every provenance
entry. The raw `PreservationIndex` identity remains the authoritative identity
for preserved members, so compaction does not merge distinct source members or
change their publication order.

Unknown relationship members and non-UTF-8 names remain preserved as their
original raw representation. They are not normalized into lossy replacement
strings merely to fit the compact recognized-member baseline.

`SourcePart` payload ownership continues to retain its existing `Arc` values.
The source-checked publication boundary, candidate-package atomicity, and
serialized part/relationship semantics are unchanged.

Hash-derived identities are rejected. A hash is not accepted as a substitute
for the raw preservation identity or as evidence that two source members may be
coalesced.

## Regression scope

Focused regressions confirm:

- empty and non-empty canonical relationship baselines use the compact sentinel
  and exact owned-byte representations;
- package and part relationship member names are derived from the raw
  preservation index without changing exact publication identity;
- an explicit empty part relationship member is raw-copied after an unrelated
  part mutation rather than silently omitted or regenerated;
- ordinary-part, generated relationship-member, exact-name, and
  ancestor/descendant topology conflicts refuse preservation before output;
- ZIP sources whose local and central filenames differ still copy exactly, but
  any append from that ambiguous topology is refused before sink output.

This record does not claim complete coverage of unknown members, non-UTF-8
names, duplicate relationships, or every publication failure path.

## Validation

- `cargo test -p soapberry-zip --lib`: 268 passed.
- `cargo test -p litchi-opc --lib`: 251 passed.
- Strict all-feature library Clippy passed for `soapberry-zip` and `litchi-opc`
  with warnings denied.
- All-feature rustdoc passed for `soapberry-zip` and `litchi-opc` with warnings
  denied and without dependency documentation.
- All Cargo commands used `CARGO_BUILD_JOBS=1` in one isolated target and ran
  sequentially.

## Measurement boundary

No total-memory, RSS, peak-memory, or OOM improvement is claimed. This change
also makes no throughput claim and does not alter `SourcePart` payload
retention, eager final residency, or publication behavior.
