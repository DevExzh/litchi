# 0704 candidate work notes

This candidate is prepared from baseline `d48523eec2` in the isolated
worktree `/home/zhuhe/code/litchi-0704-mce-candidate`. The root worktree was
never used for candidate editing or Cargo execution. The production patch is
`candidate-production.patch`; it intentionally excludes the separate policy
field and intersection patch.

The candidate adds a private, immutable default-profile slide MCE table. It
keeps the exact second `Part::blob()` observation as the `(pointer, length)`
source witness, reuses only a retained processed `Arc<Vec<u8>>`, and holds the
raw owner for the ABA proof. Fresh owned outputs are admitted only when the
checked aggregate charge fits; marker-free borrowed results and every cache
error/miss continue through the existing path. The final table is built only
from raw owners already proven by `PartDigests`, so foreign non-aliasing parts
fall back without a new typed refusal or an extra payload observation.

The cache is used only by the opened capture's context-aware
`SlidePart::from_part_with_name` route. Root, producer-name, notes-proof,
relationship, slide identity, and raw/MCE limit validation still run for each
capture. The ordinary `from_part`, semantic readers, notes readers, and public
part methods are unchanged. Snapshot, transaction, and commit owners expose a
charged-byte counter and value-preserving release operation. Snapshot rebind
projects the table through the existing digest-owner projection and drops
entries whose raw allocation is no longer owned.

The candidate uses `Arc::new` to move an already allocated processed `Vec`
without copying. As elsewhere in the library, global allocator failure during
that infallible primitive remains process-fatal; explicit table reservations
and arithmetic use fallible checks and silently abandon optional retention on
failure.

Root must apply the policy patch first or alongside this candidate. The policy
patch owns `Limits::max_retained_mce_bytes()` and its default/intersection
plumbing; those hunks are deliberately absent from the candidate patch.

No Cargo build, test, benchmark, or performance claim is made by this source
handoff. Root owns compilation, focused tests, and the baseline/candidate
measurement lane.
