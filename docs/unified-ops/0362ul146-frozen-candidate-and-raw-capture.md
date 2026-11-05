# 0362ul146 — frozen candidate and guarded raw-census execution

Candidate C is `f8cd398a770f7363f628be40edb3ce4219566c16`, tree
`7d57ecfc2b5f46f395f2eb49164ce82bc0c7eaba`, with sole parent S
`035615e0de91575af2fe8872f502a7cea32e0565`. Integration remains on
`uo2/coordinator-replay-integration`; root remains on `feat/unified-ops-api`.
C contains the reviewed implementation and protocol consumers plus the 145a
matrix binding. Root verifies clean tracked contents, all 461,401 tracked
paths hydrated, only the admitted ignored root lock, and unchanged three role
locks. Matrix SHA-256 is
`e928944a955721deb06da88003ac1d1032ae859dfd1e560179ce3dd2512ffc27`.

The reusable raw-census rebinder's first version retained C780 metadata from
its inherited plan. Returned v2 rewrites and validates C/tree/parents/branch,
including a real emitted-plan regression; eight source-only tests and
independent review pass. Root executes only its source-binding preparation,
then checks the emitted packet against C, S, all role locks, 84 profiles and
68 packages. No historical capture is relabeled.

The fresh packet is external at
`(deleted-coord-artifact)` under the worktree parent:

| Artifact | Bytes | SHA-256 |
|---|---:|---|
| `capture-manifest-Cf8cd398a770f.json` | 14,882 | `44b740d65feed93adbf54200e934ff811f6790acf11893a1e7a8cb27f92042dd` |
| `capture-binding-Cf8cd398a770f.json` | 3,248 | `e163cd9fe4f3b10c88778cf500b57611d27830d6a71486321e55878903013c0f` |
| `plan-v9-source-Cf8cd398a770f.json` | 224,766 | `edf68a9c339bc55d0223d44d62e31000cc8eca6c910a3bea15a424fa4061ba97` |
| `source-C-preflight-f8cd398a770f.json` | 5,477 | `8d486819c489e2b56e4aca815bf22a924282fbfe2a1a2aa242fe51af56c40908` |

Root starts the separately reviewed v9 adapter with the following command
from the parent worktree directory (all paths remain external to source):

```sh
PYTHONDONTWRITEBYTECODE=1 python3 -B \
  (deleted-coord-artifact) --execute --go \
  --binding (deleted-coord-artifact) \
  --packet-root (deleted-coord-artifact)
```

The first two guarded rustc identity steps pass. The 68-package inventory,
84 raw profiles and diagnostic validation are the remaining capture work;
terminal verification, callable review and distinct strict finalization are
still required. The runner preserves the 18/19/20 GB target thresholds,
source/lock/tool/binding guards, process ownership and isolated build slot.
This record reports launch, not a completed census or acceptance result.

Full53 plus seven compiler and five runtime supplements remain separate. Its
v4 rebinder is returned for producing only a plan, missing the enforced
OfficeCLI row37-preservation-before-row38 boundary, and incomplete fresh-path
guards. The executable v5 packet is being repaired and reviewed before GO.

A sparse protocol worktree `uo2-native-final-protocol`, branch
`uo2/native-final-protocol`, starts from C so qualification tooling can be
completed without mutating C. Its source is initially clean and its root
lock is copied from the admitted C lock. The native-current recipe finds a
real command defect: Cargo commands omit the external target directory whose
binaries the executor expects. Return135g must fix commands and consumers
before native execution. Signed and DrawingML binding recipes are also under
review. The signed external v1 packet was accidentally overwritten; its v2
recovery disclosure records the original hashes as unrecoverable bytes rather
than claiming preservation. No native or performance pass is inferred.

All prior leaf branches are retired. The new active protocol branch will be
merged and removed after qualification. Required detached baselines and the
already-merged runner remain until used. Main/office and their independent
lineage remain untouched. Final gates, census review, fidelity/performance,
trunk integration and final cleanup remain open. Nothing is pushed.
