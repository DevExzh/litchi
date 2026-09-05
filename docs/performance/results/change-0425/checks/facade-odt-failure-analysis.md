# ODT/polyglot facade failure analysis

This note diagnoses the three ODT-related failures in
[`facade-test.log`](facade-test.log.gz) at baseline `340cc91ae`. The inspection was
read-only; no Cargo, build, test, or profiling command was run for this note.

## Observed failures

The failing assertions are:

- [`doc.rs:2230`](../../../../../crates/litchi/src/document/doc.rs:2230),
  `owned_odt_bytes_do_not_hide_malformed_ooxml_catalog`, expects
  `Document::from_bytes(bytes).is_err()`.
- [`doc.rs:3204`](../../../../../crates/litchi/src/document/doc.rs:3204),
  `filesystem_odt_does_not_hide_malformed_ooxml_catalog`, expects
  `Document::open(path).is_err()`.
- [`catalog_detection_arbitration.rs:93`](../../../../../crates/litchi/tests/catalog_detection_arbitration.rs:93),
  `ordinary_odt_uses_native_policy_but_polyglot_honors_docx_input_limit`,
  expects a `.DOCX`-suffixed ordinary ODT to fail with the DOCX input limit.

The first two fixtures start with a valid ODT and append a canonical
`[Content_Types].xml` member containing malformed XML. The third fixture keeps
the ordinary ODT package unchanged and only changes the temporary pathname
suffix to `.DOCX`.

## Production behavior

The failures match the current content-derived arbitration policy rather than
a parser regression:

1. `litchi_odf_common::detect::packaged_has_ooxml_catalog*` is a bounded
   metadata probe. It reports a canonical catalog member, while malformed
   catalog XML is handled by the subsequent OPC probe.
2. Both `detect_docx_source_bytes` and
   `detect_docx_from_odt_source_candidate_with_limits` classify
   `InvalidContentTypesManifest` through `missing_ooxml_content_types_error`.
   The former returns the original bytes to the normal detector; the latter
   returns `Ok(None)` so the ODT candidate remains the owner.
3. The normal detector and its existing unit coverage explicitly accept an ODF
   package with a malformed catalog as ODT/ODS after the bounded OOXML probe
   (`packaged_odf_detection_handles_ordinary_and_malformed_catalogs` and
   `ods_handoff_handles_ordinary_and_malformed_ooxml_catalogs`). Thus a valid
   ODT with an inert malformed OOXML-looking extra member reaches the ODT
   facade and opens successfully.
4. The source ODT probe deliberately does not consult the filename suffix. Its
   comment documents that extensionless and incorrectly suffixed ODT files are
   accepted when their package MIME identifies ODT. An ordinary ODT therefore
   bypasses DOCX limits even when the path ends in `.DOCX`; only a package with
   a valid OOXML catalog enters the DOCX limit path.

This is consistent with ADR 0006's content-derived format rule and ADR 0009's
ODF detection ownership. A valid OOXML/ODF polyglot with a usable OOXML catalog
still takes OOXML precedence, and the later canonical and case-variant
polyglot cases in the arbitration test retain the meaningful DOCX input-limit
check.

## Classification and next work

All three assertions are stale against the current policy. No production fix is
indicated by this log or source inspection.

The focused test follow-up should:

- change the owned malformed-catalog case to assert successful ODT ownership
  and an ODT semantic read;
- change the filesystem malformed-catalog case to assert successful
  source-backed ODT ownership and a semantic read;
- rename or split the arbitration case so an ordinary ODT with a `.DOCX`
  suffix asserts the native ODT policy, while retaining the input-limit
  assertions for the two valid OOXML-catalog polyglots.

If the intended contract is instead to reject every malformed canonical
`[Content_Types].xml` member, that is a separate policy change: the detector
would need to distinguish malformed catalog XML from an absent catalog and
the existing malformed-catalog detector tests would need to change together.
The current source and tests encode fallback to the ODF owner, so changing only
these three assertions is the coherent baseline repair.
