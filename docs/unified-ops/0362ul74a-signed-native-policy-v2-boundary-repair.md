# 0362ul74a — signed-native v2 boundary repair

This source-only follow-up repairs the three independent review findings for
0362ul74.  The candidate source binding remains pending and the owner’s
signed-batch selection is `immediate_refusal`; no historical ul24 receipt is
changed or relabelled.

## Materialized vendor tree

`_vendor_tree_identity` now walks each reviewed vendor root with
`followlinks=False` in addition to enumerating `git ls-tree`.  Every
materialized regular file must be one of the tracked tree paths.  Missing
sparse files, ignored/untracked overlays, symlinked files/directories, special
files, modified blobs, and unreviewed vendor roots all fail before the vendor
digest is accepted.  The focused tests include an ignored overlay and an
untracked symlink.

## Disjoint source, archive, and output paths

The reviewed path guard compares resolved paths pairwise and rejects equality
and both ancestor/descendant relationships.  A future preparation therefore
cannot place an archive inside the source checkout, put output below the
source or archive parent, or use an external source capture as an output
ancestor.  Symlink components are rejected before resolution.  The check runs
before candidate identity validation or any materialization.

## Archive-to-tree proof

An archive digest alone is not source provenance.  A bound archive must use
`git-tree-tar-v1`, declare its prefix, selected Git tree id, canonical
tracked-tree digest, and the separately bound ignored workspace lock policy.
The verifier opens the tar, rejects links and non-regular members, checks the
complete regular-file set and modes against the selected Git tree, and
recomputes every Git blob id from archive bytes.  Only the separately bound
ignored lock files may be extra members.  The archive’s own byte count and
SHA-256 are also recomputed.

The source-only tests prove a correct synthetic archive and then construct a
different archive whose receipt hash is updated to the actual bytes; the
Git-blob comparison still rejects it against the selected commit/tree.

Run the focused suite:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest tools.test_signed_native_policy_v2
```

No Cargo, compiler, rustdoc, native executable, profiler, or timing workload
is invoked by this repair.

Coordinator integration registers `tools.test_signed_native_policy_v2` in the
tracked and external tooling rows (37 to 38 modules; 53 gate steps remain).
The approved runner used by active scoped validations remains unchanged.
All 13 source-only protocol tests pass in the composed tree. Candidate binding
and the signed-batch contract is now an owner-confirmed immediate refusal;
this is not an executed qualification result.
