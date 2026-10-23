# Worktree cleanup — 2026-10-23

The inactive, merged leaf `uo2-xlsb-provenance-review` was removed with
`git worktree remove` without `--force`.

- Path: `/home/zhuhe/code/litchi-worktrees/uo2-xlsb-provenance-review`
- HEAD: `7b3fd27e2ec9f6be3103807932828f183981b4fb`
- Branch: `uo2/xlsb-provenance-review` (retained)
- Ancestry: integration `a814baa1de`; not root `bcc6366b74`
- Size: 37,389 bytes before; path absent after removal
- Proof: clean index/worktree, no nonignored or ignored files, no build or
  evidence directories, no large files, and no active process or open file
  descriptor referenced the path

Named leaves containing ignored files, tracked deletions, build targets, or
active/reference work were retained. The full before/after enumeration and
proof are in `.coord/worktree-cleanup-2026-10-23/receipt.json` outside the
checkout.
