# 0362ul118 — repaired compiler pass and branch consolidation

Private integration `a8d32513b8e427c6d234ddd17fb836b0e8d2477b` passes all eight
joint compiler/consumer checks. Record116 corrected one metadata helper's
formatting; record117 removed the unused Web Extensions attribute import
exposed by minimal PPTX compilation. Neither return changes behavior or
relaxes a lint. Root verifies both failed captures and the repaired pass.

The passing packet is external at
`(deleted-coord-artifact)` under the
worktree parent. It covers 508 changed Rust files, minimal PPTX, strict
all-target Clippy for six OOXML leaves, ops and facade consumers, the current
performance/allocator harness, native resave, and minimal common/OPC.
All eight rows pass. The final target is 1,147,745,347 bytes; post-exit samples
have no owned or unrelated builders.

| Artifact | Bytes | SHA-256 |
|---|---:|---|
| `rows.json` | 62,773 | `e04f75a8226c57ca377218fc2dc44fde6b64c0dfa690657112165628ebf834cd` |
| `result.json` | 66,024 | `befa5ff358cbe9e6f42b41f3095357abb05729b2af31fb8710dc2be76f160d7a` |

An earlier read-only preflight left a stale target/tmp binding failure marker
in the packet. Its timestamp predates GO; root distinguished it from the live
capture. The exact 171-byte marker is preserved externally under
`(deleted-coord-artifact)`, SHA-256
`7322e8a5880676e3d969c3697293105a9160908397da6b64ad3423ff95f9cced`.
It is not a failed Cargo row and is not relabeled as a pass. The final result
above is the independently verified terminal capture of actual execution.

Fresh 34-row runtime and 13-profile/26-step glossary packets now bind a8d32513.
Runtime, glossary, the full workspace gate including isolated wsdoc, final
API inventory, fidelity and performance qualification remain open. The source
stays frozen while these run; coordinator-only progress documentation can
advance independently. Nothing is pushed or accepted as final.

## Local consolidation after record103

Cleanup92 removes eight obsolete composed-source snapshots after preserving
their locks. No allocated pre-removal size was captured, so they contribute
no space claim. Its ninth, dirty PPTX snapshot is reviewed separately: the
eleven working files are identical, superseded, or stale intermediate
variants that would regress quote/limit handling or retain unnecessary state.
Root reapplies the archived patch to the retained Git base and obtains all
eleven exact preserved files. The source checkout is removed after approval;
the patch and full files remain outside the repository.

Cleanup93 removes nineteen retired branch worktrees. Fifteen are clean;
four have reviewed dirty payloads. The XLSX and ODG edits are formatting-only;
Numbers limit propagation is represented by the later typed/canonical API;
PPTX variants are already represented or superseded. A truncated performance
JSON is preserved as unusable historical evidence rather than merged. Root
runs the independent 311-check verifier, replays all four dirty patches, and
additionally matches every one of 355,047 deleted paths, modes and blob IDs
to the actual source tree. All nineteen refs are then deleted atomically by
exact tip. The old receipt's minimum-blob-size prose says five; the actual
manifest correctly contains 39,412 empty blobs and a minimum of zero.

Cleanup94 removes seven old gate/API source checkouts plus the historical
fidelity source. Root verifies/restores all 115 archived payload files,
reapplies the fidelity patch, and verifies the preserved receipt hashes.
The two differing historical fidelity-report variants remain distinct.
Their existence supplies no final fidelity acceptance.

Known foreign read-only filesystem scanners are left running. Their open-file
exceptions are recorded; no project writer or builder is present in removed
checkouts. No universal zero-FD claim is made. Previously recorded cleanup83,
89 and 90 evidence/guard limitations remain unchanged.

The notes benchmark worktree is cleanly detached at measured baseline
`c5d2b30740186e2b4beb4a4d84e929e00a1c5718`. Its previous `009f88d14e` branch
differs by one provenance document only. Root archives that document and all
three role locks, verifies the eight identical harness blobs against a8d32513,
and deletes the redundant branch. The baseline checkout remains available;
its production source is never merged. The old inert performance packet still
requires a candidate/lock rebind and execution.

The following files are relative to the external cleanup directory's
`trunk-consolidation/`:

| Evidence | SHA-256 |
|---|---|
| `cleanup92-root-verification.json` | `f890ae1293c9808ed24aa85679fd0a73c7d4d87491564ab3fa8a1a54cbc10f3f` |
| `cleanup93-root-actual-verification.json` | `54b22dd97c49c6f9e1823ec786359fef02ccacbcf0bb697dc6e33ecc95c2d787` |
| `cleanup93-root-supplemental-verification.json` | `52a04bc12ef92d2508a8ff097d47c0d7910f07ee3641bd827d7c05d4f51e93a2` |
| `cleanup93-nineteen-ref-delete-receipt.json` | `2dad2a7ce63aadc600ce160f2c5489beb4cbb90a3072269d5e170674033b5b8d` |
| `cleanup94-root-verification.json` | `c3dfe64a1c42dfddd5a6201853bc3bedf62e37add492c7a209c31348c3d7543e` |
| `cleanup96-notes-baseline-detach-receipt.json` | `9c98810f03ca18cc08c7b6c0efd2f00fd4e48484ce1c57c9bf726b6a69085275` |

Cumulative removals are 422 refs and 231 worktrees. The measured allocation
ledger is 676,827,754,496 bytes: cleanup92's dirty PPTX adds 7,565,844,480,
cleanup93 adds 119,901,904,896, and cleanup94 adds 60,457,906,176. The same
unmeasured-wave and supplemental-artifact exclusions apply; cleanup92's eight
clean snapshots are also excluded. Dirty-fidelity pre-allocation is explicitly
reconstructed from the measured source after lock move-out plus the measured
lock allocation. These are allocation-accounting figures, not a filesystem
free-space delta.

There are 73 local branches: main and office-format-completeness, 68 independent
office siblings, unified trunk, private integration and the active gate runner.
Other unified worktrees are detached measurement controls or still-reviewed
historical endpoints. Cleanup95's next seven candidate removals are authorized
but are not included in these completed totals. Root remains checked out on
`feat/unified-ops-api`; the protected constructor probe is unchanged.
