# Independent source review, 0775

Read-only review of 6154a06398 against 92dfbe0bd5 found migrated consumers consistent, retained URI-length checks and bounded scope identities. No Cargo/native work was run by the reviewer.

The matching registered-extension path still constructs a full `Name` and copies the URI in both stream and tree processors. Integration must not claim that all extension checks are independent of URI length. Fresh matching-extension controls are required; that remaining work is explicitly outside this optimization. Namespace declarations still hash URI text once per declaration, so many aliases remain declaration-linear. Public NamespaceUri text hashing also remains proportional to URI length.

Existing alias generation records separate processor results; it does not itself establish tree/stream parity or valid alias-pair semantics. Fresh valid alias/direct pairs and explicit differential triage supplement existing unit tests. Coverage cannot be called exhaustive.

The old patterns module comment and historical 0764 description refer to the pre-0771 stream representation; current integration documentation must distinguish them.

Final differential review: all 1,017 generated changed keys belong to the tree processor, with no changed stream/fixture/ordinary-generated keys. The 20 pair differences repeat the reserved XML alias correction. The pre-ID source bdea6e8b86 is the last text-based MCE implementation before c0d5b94428. Root's identical-probe capture matches all 64,541 candidate outcomes exactly, not merely normalized hashes.

The pair-specific codec oracle reparses output under empty capabilities; it does not independently compare raw codec serialization or opaque extension output. The six pair templates contain no extension elements, and 64 repetitions are disclosed as six shapes. The independent reviewer confirmed the existing expanded-duplicate gap is recorded rather than asserted to pass.

Final independent acceptance review found no blocker: 18 cases/108 timed processes, 64,541 differential comparisons (63,504 unchanged and 1,037 changed), exact pre-ID match, and all three regression flags recompute. No reviewer Cargo/native work or file edits were performed. Root verified the entire candidate/pre-ID map is identical, beyond the changed-key comparison.
