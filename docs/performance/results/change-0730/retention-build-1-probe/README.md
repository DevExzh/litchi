# 0730 DOC retained-render probe

This is a candidate-only memory probe for the bounded DOC validated-render
handoff. It is isolated from the frozen 0730 timing probe and must be run only
after the candidate `body_text` API is present. It has two binaries:
`doc_retention_probe` uses the normal allocator and
`doc_retention_probe_alloc` installs the same counting system allocator used by
the main 0730 probe.

The probe accepts one of the two inherited DOC fixtures, a one-operation or
two-operation public `Edit` route, and one of three retention controls:

* `zero`: `TransactionLimits` has a zero retained-render ceiling;
* `default`: the pilot's explicit 8 MiB ceiling is used; and
* `release`: the same 8 MiB ceiling is used, then the retained render is
  explicitly released after each successful staged replacement.

The first replacement is the exact 0728/0730 text
`litchi copy-through baseline replacement text` (45 UTF-16 units). The second
replacement, used only by the two-operation route, is a different fixed text.
Both operations target paragraph zero in the same public `Edit`, so the second
stage exercises invalidation or replacement of the first handoff.

Each stage records `Edit::retained_render_bytes()` before and after the
replacement and after an explicit release when the selected control requests
one. The value is the owned `Vec` capacity, not its serialized length. The
probe also records the final retained report immediately before commit.

The allocator region starts after the source file has been read and ends when
the final committed output `Vec` is returned. It covers public snapshot open,
edit construction, one or two replacements, commit, and final output copy.
The source input is held before region entry. Direct reference output is made
after the region and compared with `Vec` equality, so reference work and its
allocations do not enter the reported region. The reported region is therefore
an ownership/allocator observation; it makes no timing, RSS, or speedup claim.

Every selected control publishes the final output hash and a direct byte
equality result against an independently recomputed zero-retention reference
for the same fixture and route. Root compares these hashes across zero,
default, and release processes and retains the JSON evidence. The direct
comparison must remain true even when retention is over budget or explicitly
released; those cases are recomputation fallbacks, not operation refusals.

Example shape (root supplies the executable and exact process schedule):

```text
--case docnohf --input test-data/ole/doc/NoHeadFoot.doc \
--edits two --retention default
```

No iWork fixture belongs in this probe. Root owns compilation, process
ordering, native cleanup, and interpretation.
