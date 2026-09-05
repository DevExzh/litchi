# 0426 facade correction source review

I reviewed the facade-only corrections against baseline `f22917bbe` and the
reproduced baseline failures at `340cc91ae`, with ADR 0006 (content-derived
format ownership and typed incompatible-format errors) and ADR 0009 (ODF
detection ownership) as the applicable contracts. This is a static review;
no Cargo command, build, test, profiler, or CPU workload was run here.

No blocker was found in the reviewed corrections. They change test fixtures
and stale expectations; they do not weaken production validation or archive
limits.

## ODT arbitration

The owned and filesystem malformed-catalog regressions now expect the
malformed `[Content_Types].xml` member to fall back to the ODT owner and still
read the ODT text (`crates/litchi/src/document/doc.rs:2230` and
`:3210`). That follows the existing content-derived probe policy: a malformed
OOXML probe is not a valid OOXML owner, while the valid ODT package remains
readable. The filesystem case also retains `DocumentImpl::OdtSource`, so the
source-backed ownership boundary is checked rather than merely checking a
successful parse.

The renamed `.DOCX` ordinary-ODT case now asserts native ODT behavior, while
the valid ODT/OOXML polyglot loop still exercises both catalog spellings and
requires the typed `ResourceLimit::InputBytes` result for owned bytes and a
path (`crates/litchi/tests/catalog_detection_arbitration.rs:93`). Thus the
correction respects ADR 0006's filename independence without dropping the
valid-polyglot preflight limit proof.

## XLSX deferred payload proof

The byte-backed selected-read fixture now flips a byte inside the stored
`sheet2.xml` payload while preserving the local header, central record,
declared sizes, and CRC metadata (`crates/litchi/src/sheet/adapters/xlsx.rs:777`).
This is the right corruption boundary for the strict archive-layout proof:
catalog construction can succeed, `First!A1` is checked as the concrete
`Int(7)` value, and reading the damaged second worksheet is still required to
fail (`:906`). The pre-existing central-CRC mutation remains confined to the
separate filesystem test that explicitly verifies deferred CRC failure, so
the two tests do not conflate structural proof with selected payload
verification.

## XLSB wrong-facade proof

The minimal XLSX fixture now has a valid OPC content-types default for
relationships and a root `officeDocument` relationship targeting
`xl/workbook.xml` (`crates/litchi/src/sheet/workbook.rs:1241`). The workbook
catalog remains an empty, valid XLSX workbook. This lets content-derived
detection finish as XLSX before `open_xlsb_workbook_dyn_with_limits` rejects
the incompatible facade; the test continues to require the exact typed
`Error::NotOfficeFile` (`:1268`). It does not infer the answer from the temp
file's suffix or accept an arbitrary parse error.

The unrelated legacy XLS Formula-record failure remains outside this review
and is owned by the separate format investigation. No result is claimed for
that case.
