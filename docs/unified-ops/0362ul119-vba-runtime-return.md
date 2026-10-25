# 0362ul119 — VBA runtime return and detached-checkout cleanup

The fresh 34-row runtime packet on frozen integration
`a8d32513b8e427c6d234ddd17fb836b0e8d2477b` stops in its first common-library
group: 529 tests pass and two fail. The preceding eight compiler/consumer
checks remain passing evidence for this source; they do not establish runtime
acceptance.

The failures are returned for repair as 119a:

- `edit_rejects_tighter_structural_limits_before_publication` expects the
  aggregate relationship XML size to equal the root relationship XML size;
  the actual values are 2,529 and 450 bytes.
- `replay_projects_serialized_source_namespace_after_unrelated_remove_add`
  receives `PreservationUnavailable` while serializing its source fixture.

The return must determine whether the defect is in production or the fixture,
preserving structural bounds, publication fidelity and performance. No
assertion or preservation rule is waived. The original failure remains intact;
the remaining runtime groups will run as an explicitly partial diagnostic
continuation on the same source while the repair is reviewed separately.

The external packet is
`.coord/ooxml-pptx-joint-a8d32513b8-f5-runtime` under the worktree parent.
Root verifies the terminal result, source and three unchanged locks, log
hashes, target accounting and post-exit process cleanup. The first group
takes 61.584 seconds; its final target is 387,989,123 bytes.

| Artifact | Bytes | SHA-256 |
|---|---:|---|
| `rows.json` | 4,337 | `f1fdc514833e0da3f1f33402a9017a63ad55789e902f33cb66948680d2f90c8d` |
| `result.json` | 5,076 | `57854f1d40fdfabe1f5b1b4290380db7d66bbfad1e7930c0f1f0d21093edc15d` |
| `common-tests/0000.log` | 213,792 | `0438ec51f52c79499095fce2169dccda90353ee463e9b12428fe323a96e97721` |

Regeneration is the packet's source-bound `python3 -B run_f5.py --go`
command with fresh target/tmp directories and the approved isolated runner.
The failed packet is preserved; regeneration must use a new packet directory.

Cleanup95's first seven obsolete gate/API/tooling source worktrees are removed
without force. Root independently verifies their Git tips remain reachable,
their paths and registrations are absent, and all 21 preserved payload files
match their recorded bytes, modes and hashes after an actual restore. The
payload is 1,293,606 bytes. Its manifest hash is
`2077af5ca8e4bda7effbe69a3ba33cd2dbc8881476ec7feb76665fd145a65d36`;
the removal receipt hash is
`497139b575d0cc7e5992cb861032f87279530750e37ca43f705527e784a5eca8`.
Both are external under
`.coord/worktree-cleanup-2026-10-25/trunk-consolidation/quarantine/wave95-first7/`.
The root verification receipt is `cleanup95-first7-root-verification.json`
in `trunk-consolidation/`, SHA-256
`27003e60fc5af2f4dc65e2c18c66a6e7859dab6e34f62194a8aa5425ff6206ca`.

Pre-removal allocated size was not captured despite the cleanup instruction.
The allocation addendum records this omission; no reclaim amount is inferred
from apparent size. Cumulative completed cleanup is 422 deleted refs and 238
removed worktrees, including the previously disclosed evidence limitations.
The measured allocation ledger remains 676,827,754,496 bytes. Historical
benchmark controls and the protected constructor probe remain intact.

Fresh glossary, full-workspace/wsdoc and notes performance packets are inert.
They supply no new acceptance result. Final runtime, API inventory, fidelity,
performance qualification and integration into trunk remain open. Nothing is
pushed.
