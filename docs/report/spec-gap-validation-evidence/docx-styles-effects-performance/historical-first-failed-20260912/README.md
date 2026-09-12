# First frozen-source smoke failure

This directory is a durable, source-pinned record of the first bounded
`stylesWithEffects` smoke attempt. It is historical correctness evidence only;
its timing and RSS values are not a performance result.

The clean source worktree was at Git commit
`995820ee18bf9b8647a6f79f9ad0efbb938ffffd` and the harness production pin was
`d687e38349e4348506a56dc4ae996298844d4091`. The captured source manifest hash was
`443df85cce80960c848a3c115a07d39782923560ddee185a80a77e107e12dc65`; Cargo metadata before and after
were both `4c262809b136d1c1c049e18c46b69b4faec89d03e2dfb639ad483eb43d0b1d26`. The build used the
binary hash recorded in `binary.sha256` and `binary-after.sha256`; those files
match byte-for-byte. The original command/environment record is in
`commands.txt`, and the original build output is in `build.log`.

The preserved raw bundle contained 43 lane JSON receipts, 43
RSS sidecars, and 43 stderr sidecars. The complete external
bundle remains at `/var/tmp/litchi-docx-styles-effects-first-failed-raw-20260912-b`.
`raw-bundle-sha256.txt` retains a byte-exact copy of the original per-file hash
manifest. The archived filenames in this directory intentionally
drop the original `smoke-` prefix; the copied raw-bundle manifest retains the
original names so each selected file can be matched to the external capture.
The original executable target was cleaned after the run; its before/after
hashes are retained. Cargo metadata and the clean source worktree remain
external. This directory records selected failure evidence and its
source/provenance hashes. Re-execution requires those source inputs and a new
build; this slice is not a standalone replay bundle.

The failing nested existing-owner control was
`cap_total_part_bytes`. Its exact-fit replacement projected aggregate part
bytes from **110,027** to
**110,151**. The one-unit-under limit was
**110,150**; publication returned the concrete typed error
`DocxError::Opc::ReadLimit` with resource `TotalPartBytes`, actual
`110151`, and maximum `110150`. The receipt records
`cap_existing.commit_stage_checked=false` and
`cap_existing.commit_refusal=null`: the publication guard
refused, but the transaction commit-stage refusal was not observed. The
fail-closed verifier therefore rejected this smoke attempt. This is the
pre-fix behavior that the later production commit
`d1f299d00e0dd5cc5cd8ddf9811c4b1ad21d1119` was intended to correct.

The exact machine receipts are retained as `cap_total_part_bytes-p1.json`,
`cap_total_part_bytes-p1.time.txt`, and the empty stderr sidecar (corresponding
to the original `smoke-cap_total_part_bytes-p1.*` names). The JSON also
records the source-less aggregate-cap control, allocator counters, package
metrics, physical/metadata readback, and typed refusal fields. No claim is made
about a speedup or a managed memory limit from this historical attempt.
