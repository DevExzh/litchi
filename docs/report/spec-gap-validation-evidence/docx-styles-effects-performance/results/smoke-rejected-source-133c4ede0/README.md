# Rejected DOCX current-source smoke attempt

These 166 raw files are retained for audit only. This run is **not approved**.
The frozen harness at `133c4ede020fd221e97722d4c5c73f26394eeb68` emitted
historical source `d1f299d00e0dd5cc5cd8ddf9811c4b1ad21d1119` in every receipt,
while the current verifier required
`8702fd4db8723acceb7deb51bcb40ff66604bf10`. Verification failed, so there is
no passing verification receipt. Successful process exits do not override
that provenance failure.

`retained-files.json` records each original file's length and SHA-256. The
files are copied byte-for-byte, not repaired or relabeled. The subsequent
[fresh corrected run](../smoke-current-637082e31/) passed at a new frozen
harness commit. Neither directory is timing-matrix evidence.
