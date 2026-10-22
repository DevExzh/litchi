# 0360em — XLSB replay content identity and copy admission follow-up

- Status: source and focused validation complete on
  `uo2/w4b-xlsb-raw-policy`; native performance qualification remains open.
- Parent record: 0360el.

## Content identity

Replay is authorized by the recorded package content. A package reopened from
the same bytes is an eligible recipient even when its owner allocation or
source identity differs. The XLSB leaf returns `PatchConflict` only when the
recipient's current serialized image differs from the patch `before` image.
The unified facade may reject a changed foreign owner at its transaction
boundary, but its guard accepts equal content; this is not an owner-identity
exception.

## Bounded exact-image handoff

Structural and inverse patches retain their before/after images as
`Arc<Vec<u8>>`. A borrowed `PhysPkgReader::new_with_limits` preflight checks
the recipient input and ZIP archive/member envelope before ownership is
taken. The accepted image then enters `OpcPackage::from_shared_vec_with_limits`
through the same Arc, retaining exact-source publication without cloning the
whole archive. The generated replay path applies the same preflight and
single-Arc handoff. BIFF12 raw limits remain enforced by `project_candidate`
and complete worksheet validation before publication.

## Validation

The existing 0360el leaf and facade suites remain green after this follow-up;
the all-test compile also migrated the three exhaustive external-link error
matches for the typed `PatchConflict` variant. No native timing, profiler,
binary, or large evidence file is committed.
