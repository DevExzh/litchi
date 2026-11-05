# 0362ul89ap — Reviewed PPTX limit transcript regeneration

The actual four-row capture on `e10039df1c54e5815c6d02805a0f5921e6841a67`
passes formatting (2.839 seconds), minimal all-target compilation (65.684),
strict all-feature/all-target Clippy (83.288), and the explicit ignored golden
generator (97.686). The generator itself reports one passed test, zero failed,
and 1,386 filtered out in 4.52 seconds. Final target size is 798,501,049 bytes.
Root verifies every receipt, source/three-lock fence, command, log/trace,
resource guard and terminal ownership; the persistent session exits 0.

Capture: `(deleted-coord-artifact)`.

| Artifact | Bytes | SHA-256 |
|---|---:|---|
| `rows.json` | 40,065 | `408fc399b0fe9ab715d1dc911268a467d77e8d8d8a70f6c36b1011e7cf8cbf40` |
| `result.json` | 42,164 | `2a5a6723aaa4902daf2f5cc5bd0bee09a595dcc6fb81cc11f3aab9cf99271fdb` |
| `golden-artifacts/digest_reuse_lptx0003.stdout.log` | 17,190 | `afb5e595f4befb3e8ee5095a6a9d4835479097a1cd9c194e6ca9696859b4b13c` |
| `golden-artifacts/digest_reuse_lptx0003.transcript.txt` | 16,989 | `3a2ad886bc5b3ef04efb77ebc73dfd8c2cd58ae1d4d2e4658a37e8ecf49f9c63` |

The command is the explicitly requested filtered library generator with
`--ignored --nocapture`. Blessing variables are unset. Raw stdout remains
outside Git. Root accounts for the harness prefix on the first row and verifies
the terminal test summary before extracting the complete transcript.

Independent review verifies all 192 unique row keys and their unchanged order,
all 140 digest values, all ten unchanged bounds, all 32 refusal rows and the
legacy non-wire/non-limit invariant. Exactly 18 rows change; the other 174,
including every wire, revision, publication and untouched-state value, remain
identical. The changes express measured runtime charges as typed limits:

- Media plan: `PatchBytes`, observed 137,457.
- Media patch parsing: `PatchBytes`, observed 82,578.
- Plain patch/plan: `PatchBytes`, observed 31,375.
- Plain source-exact planning: `Codec(ArchiveBytes)`, observed 31,404,
  maximum 31,375.

Each maximum is unchanged. These replace stale invalid-limit or string-resource
expectations without changing the operation, bound, output or replay behavior.
The legacy constant remains frozen. This separate artifact commit copies the
reviewed generated transcript byte-for-byte into the versioned golden; it
changes no production code or test assertions.

Fresh full/minimal runtime helpers, glossary, downstream and performance
qualification remain required on the new source. This four-row capture is
scoped evidence and does not discharge the retained 23-row proposal or final
workspace acceptance.
