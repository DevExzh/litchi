# 0726 empty-slot setup hypothesis

The 0725 combined candidate was rejected. This fresh candidate changes only
replay_indexed_cell: resolve the existing slot slice after worksheet lookup;
for an empty slice retain the trailing execution and freshness fences and
return None before path-vector/resolver/hint setup. Stored replay and the
224-byte index, scanner, budget admission and public API are unchanged.

Use the full 0725 matrix afresh against 3ad29e42da. Keep 5% p50/mean native and
repeat limits, the native warm 10 ns exception, all source/outcome/custody gates.
Require owned 54016-missing native q8 and repeat p50/mean benefits >=10% in both
pairs. q2 and all nonmissing counters must be exact; missing positive-budget
q3/q8 must remove exactly 16 allocated/deallocated/peak bytes and one allocation/
deallocation call, with zero retained-byte change. Every budget-fence source
metric must now remain exact because admission is unchanged. No checkpoint
benefit claim is made. All rules precede main measurement. Retain only if all
gates pass; otherwise archive and restore. Separate missing-query profiles and
allocator deltas supplement prior route evidence; no instrumented trace build
is needed for this local setup deletion. The non-iWork goal remains active.
