# 0362ul165ua — Clippy profiles pass; fuzz lock inputs are missing

The current-7c diagnostic passes all 29 Clippy feature profiles at row 40
in 782.88 seconds. Root rehashes the log and target trace, verifies clean
source and four unchanged locks, and confirms that the resource guards and
post-exit process checks pass. Target usage is approximately 6.14 GB.
Together with row 39, the diagnostic has passed 31 build profiles and 29
Clippy profiles. Neither row substitutes for the fresh final-candidate gate.

## Fuzz prerequisite

Row 41 exits with status 101 before compiling its first fuzz workspace.
`crates/litchi-core/fuzz/Cargo.lock` is missing, and the gate's Cargo wrapper
correctly enforces `--offline --locked`. The diagnostic source is unchanged,
and the run continues to golden checks. This is a missing preparation input,
not a demonstrated fuzz-code failure or an accepted exception.

All four standalone fuzz locks are absent from this diagnostic source:
core, RTF, ops and the facade. Ignored copies of the previously approved
core/ops/facade locks exist in the current unified worktree and match the
recorded byte identities from 0362ul44a. They are candidate seeds, not proof
that their dependency graphs still satisfy the current manifests. Before a
fresh fuzz run, preparation must verify them with offline locked metadata.

The older RTF seed belongs to a different source/root-lock combination and
is not reusable as current evidence. Preparation must resolve a fresh RTF
lock offline against a clean source with an explicit seed policy and record
its identity. A separate preparation handoff is being built. It must not
mutate the active diagnostic source or reuse the failed row as a pass.

## Protocol review

The v14 static composition closes the missing-output return: all 24
overlays match virtual tree `c1d7345b`, the previous 34 base files remain
present, and the two 52-file output trees pass 20 composition and 20 P2
controls. The shared gate retains both ODF lock tracking and DOCX tests.

Independent review returns two remaining problems. Some current-looking
metadata still names historical `9feb`/`172c` source identities. The copied
producer recipe also accepts only frozen C33 and supplies no concrete
rebinding mechanism for the later selected candidate. A successor must
distinguish historical metadata from pending final identity and provide
verified, coherent rebinding of the recipe and its pin inputs. Static
output checks do not establish producer readiness.

## Evidence

Paths are relative to `(deleted-coord-artifact)`.

| Artifact | SHA-256 |
| --- | --- |
| `remaining20-feature-matrix-clippy-root-verification-20261026/receipt.json` | `c98673f34c10443df8725ffa39c500b221f3a6fc6fa2dab31e9bf54669b90082` |
| `remaining20-fuzz-root-verification-20261026/receipt.json` | `52aab20d7e7e2790a570ea6ede77efbed2360635bd9383d9a4815c25236ed9c8` |
| `protocol-consumer-composition-C2a-20261026-v14/receipt.json` | `88153f7083315038053d72e2571c7bab15225be8e39443ef68e87e9b3fcf162a` |
| `protocol-consumer-composition-C2a-20261026-v14-independent-review-20261026-fresh/review-receipt.json` | `61546928d08eea0f9c6773be901310dea790076fb938cbfea7062f20cfefbe85` |

The DOCX release builds remain queued behind the live diagnostic. Actual
final-source binding, all 29 native groups and final all-65 remain required.
Nothing is pushed.
