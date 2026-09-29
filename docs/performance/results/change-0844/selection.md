# 0844 — stage only missing Part-name ancestors

Previous goal turn: progress. Commit 1b6072fc32 retained 0843's current-source
many-Part publication measurements and removed owned temporary roots. Its
4,096-Part profile attributes 20.39–20.45% of owner guest instructions to
PartNameIndex::insert; allocation stacks identify 4,096 temporary ancestor Vecs
and 4,096 ancestor Strings despite every name sharing /custom.

Test one small change: borrow-lookup existing ancestor keys, stage only absent
keys, then perform all fallible reservations before index mutation. Preserve
first-descendant representatives and duplicate/equivalent/derived conflict
ordering. Keep randomized hashing, source validation, typed errors, resource
bounds, exact output, and owner boundaries. This adds no memo, cache, worker,
public API, durability policy, codec or unsafe code. Accepted ADR 0001–0006,
0010/0011/0024,0030–0032 remain constraints; their recorded hashes match the
previously read normative set. ODF remains deferred and iWork excluded.

Four shapes separate benefits from guards: shared /custom/pNNNNN.bin;
root /pNNNNN.bin with no ancestor; deep-shared /a/b/c/d/e/f/g/h/pNNNNN.bin;
deep-unique /uNNNNN/b/c/d/e/f/g/h/pNNNNN.bin. Each has 4/64/1024/4096
256-byte opaque Parts and two metadata members. Only the last Part changes.
Publish and lifecycle retain 0843's public route and clock boundaries.

Both release legs use one probe, compiler, features, dependency versions and
CPU 12. Sixty-four qualification outputs must be byte-identical across legs and
modes per shape/count, with independent ZIP and exact logical checks. Six
matched native blocks alternate case and leg order:384 processes × 30 samples after
3 warmups. Heaptrack separately captures publish once per child, two repeats per
shape/count/leg:64 processes. Formal total 512 reports / 11,648 samples. Native timing,
process RSS / faults, and owner-ancestry allocation events / requested bytes remain
separate populations. Requested allocation bytes are not peak live memory or
copy volume; whole-child RSS includes fixture construction and validation.

Flag paired p50/RSS medians above 1.05 with 95% bootstrap lower endpoint above 1;
review every p95/p99 point above 1.05, allocation growth above 1.05, and block-p50
max/min above 1.20. Bootstrap 10,000 draws, seed 844084, six matched block units.
No native speedup is inferred from allocation or guest-instruction shares.
Do not adopt a shared-ancestor win that materially damages unique/deep/root
common cases. Reject or revise if the controlled comparison fails; preserve
all initial results rather than selectively replacing workloads.

Production baseline exactly matches 0843/0842, so the sealed previous gates are
replayed. Fresh baseline probe/OPC tests precede candidate application; fresh
candidate probe, workspace library/tests check, workspace library Clippy,
affected-owner all-target check/doc/tests, and crate boundaries must pass.
Previously confirmed iWork example/rustdoc failures remain excluded. No broad
semantic Office, native producer, cold-cache, filesystem, remote or scaling
claim follows from these synthetic substrate workloads.
