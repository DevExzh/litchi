# 0362ul126 — enforce the recipient event budget during PPTX read-only replay

The frozen a8 runtime qualification stopped at
`pptx_readonly_recommended::replay_rechecks_readonly_output_with_the_recipient_profile`.
The source package was opened with `max_events = before_events + 1`, while the
read-only edit adds an extension branch.  Replay nevertheless succeeded, so
the assertion at line 169 reported that the expanded owner was not refused by
a typed bound.

The captured log is
`(deleted-coord-artifact)/24-pptx-full-chunk-31-attempt-actual-f5-24.log`
(227,285 bytes, SHA-256
`4809895faad74af3e2eea5540514e52a80e7579445392cc0087b1e5a4b6d20ca`).
The frozen source is `a8d32513b8e427c6d234ddd17fb836b0e8d2477b`.

`presentation_properties::Limits::from_source_limits` used the recipient
event ceiling only as a semantic node/extension ceiling.  Its parser did not
charge every physical quick-XML event, so the added opaque extension could
fit even though the recipient's source profile could not.  Replay validation
now applies the shared PPTX physical XML-limit scan to the changed
presentation-properties part before the strict vocabulary parse.  The scan
returns the existing typed `Error::Opc(OpcError::Limit { kind: Events, .. })`
with the measured event count and recipient maximum.  Read-only, math, and
metadata replay callers all pass the retained recipient profile through the
same helper; unchanged parts and no-op patches remain unscanned.

The extra physical pass is intentional and bounded to the changed
presentation-properties payload: the replay callback runs once per changed
payload, then the existing strict properties parser performs its semantic
pass.  `part_edit`'s generic source admission checks parts and byte totals but
does not inspect XML events, and replay installs changed blobs through
`set_blob_shared`/`try_add_part`, so the OPC publication audit is not available
at this point.  The strict properties parser counts only retained
start/empty nodes after MCE processing; it cannot represent the raw event
budget when an MCE branch is discarded.  Reusing the existing
`source_limits::validate_xml_limits` scan is therefore the smallest exact
physical check and does not rescan unrelated parts.  Runtime qualification
must still verify this write-set-only scan does not regress replay throughput
or bounded-memory behavior.

Performance acceptance is pending a paired ABBA measurement on the composed
tip.  It must compare the pre-repair and repaired read-only replay paths for
the same changed `presProps.xml` payload, with a large extension payload and
the recipient event ceiling, recording wall time, instructions, peak heap,
and the refusal result.  The run must also include no-op and
relationship-only controls to show that the new scan is not reached there;
the expected count is one bounded physical scan plus the existing semantic
parse for one changed properties payload, with no package-wide scan.  The
measurement source, commit IDs, and small summary belong in the external
qualification evidence; this source-only record does not claim that
performance acceptance.

The presentation-properties replay unit test now also asserts the typed event
limit.  This is a source-only repair: no Cargo, build, rustdoc, or metadata
command was run.  Fresh compiler and runtime qualification on the composed
tip remains required.
