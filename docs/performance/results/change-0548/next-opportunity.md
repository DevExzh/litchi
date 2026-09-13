# Next OLE2 opportunity: keep exact cycle detection, reduce repeated bitset work

This is an unimplemented hypothesis, not an admitted improvement. The 0548
checkpoint replacement increases dynamic collector instructions despite fewer
static instructions. Before another checkpoint design, consider a smaller
candidate that keeps the current exact first-duplicate detection and combines
membership testing and marking into one checked private bitset operation.

The restored source performs `contains(slot)` followed by `insert(slot)` after
checking the FAT/MiniFAT slot. Both source helpers compute the word and mask
and check logical/backing bounds. The measured collector already inlines these
helpers, so source-level duplication is not proof of remaining machine work.
Fresh disassembly and instruction attribution must establish which checks,
loads or mask operations a fused operation actually removes.

A possible private operation would return whether the bit was already present,
checking logical length and backing word safely and setting a previously absent
bit without allocation. Its caller must still report the current exact cycle
error before appending a duplicate sector, preserve invalid-index precedence,
and keep allocation labels/order, zero-fill, scratch reset/reuse, FAT/MiniFAT
separation and collect-before-claim semantics unchanged. The existing standalone
bitset API and other callers need not change. No unchecked indexing or unsafe
code is justified.

If compiler evidence shows no removable work, do not integrate this source
rewrite merely for appearance. A cheaper checkpoint representation remains a
separate possibility, with its arithmetic and replay proof still required.
Any viable candidate must pass fresh matched native, allocation, profile,
public malformed guard and correctness gates; 0548 gains cannot be reused as
its evidence. OLE2/OOXML remain active and ODF remains deferred.
