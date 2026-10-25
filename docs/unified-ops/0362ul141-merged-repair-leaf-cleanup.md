# 0362ul141 — retire the merged readonly and MCE leaves

After record140's composition, root removes the two merged repair worktrees
without force and deletes their refs with exact-tip compare-and-swap:

| Branch | Retained commit in integration |
|---|---|
| `uo2/pptx-readonly-noop-0362ul134a` | `4078b4350b7348e080c6fec4c6a6a6770760fffb` |
| `uo2/docx-mce-limit-identity-0362ul138` | `d39e0adc0bb88d71dedc6405a47a185ce60a3236` |

The fresh root preflight passes fourteen checks: exact worktree/ref/tip,
ancestry, clean tracked state, all five source blobs matching integration,
protected probe, and process-holder policy. No nonregenerable payload exists
in either leaf. Seven ignored Python bytecode files are individually checked
by path, mode, inode, device, size and hash before disposal. Source, locks,
root probe, retained baselines and other worktrees are unchanged. The holder
policy has no visible holders, unclassified processes or unexpected errors;
the previously documented process-visibility limits remain explicit.

Root verifies absent paths, registrations and refs after removal. The
registered worktree count is 57. Cumulative consolidation reaches 270 removed
worktrees and 432 deleted refs. This wave removes 15,136,047,104 allocated
bytes and retains 425,984 allocated bytes of evidence, a net reduction of
15,135,621,120 bytes. The conservative allocation ledger is 800,318,455,808
bytes under all prior exclusions and evidence limitations; it is not an
observed filesystem free-space change.

Evidence is under the worktree parent's
`.coord/worktree-cleanup-2026-10-25/trunk-consolidation/quarantine/wave103-root-removal/`:

| File | Bytes | SHA-256 |
|---|---:|---|
| `preflight.json` | 185,540 | `127b3e98576639a280bc885ff50ca6b80a4f8968f0639f6d30f0542213e91d30` |
| `result.json` | 1,379 | `7182016a8883d4373616e9bdf10adef015535f44e9eff5b44f3cad93d6245bbd` |
| `allocation.json` | 913 | `c19becc903ce0848d174ac356884e3ccdcc9ace6362deade531d8c8ebabcc677` |

The retained `cleanup103-root-remove.py` records the exact verification and
non-force removal commands. It is a one-shot mutation record, not a command
to rerun against absent worktrees. All merged commits remain reachable from
integration `e7841dd222f17a00abe79c5ae034ad18cde3f45d`.

The branch census also confirms that the 68 independent office-lineage refs
remain outside the requested unified-derived sweep. The new detached a8
checkout is retained as the required 126 performance baseline, as corrected
in the cleanup102 addendum. Three active native protocol leaves and the
integration/runner branches still have required work. Final qualification
and trunk integration remain open; nothing is pushed.
