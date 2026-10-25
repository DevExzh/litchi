# 0362ul89as — PPTX Slide Synchronization runtime repair

Date: 2026-10-25. This is a source-only correctness repair; it does not claim
a new compiler, runtime, performance, or qualification result.

The frozen source base is
`ef9470f7da08a131f31600041f9d562203d62bfa`. The bounded full-feature test log
is outside Git at
`/home/zhuhe/code/litchi-worktrees/.coord/pptx-combined-0362ul89-f5-qualification-ef9470f7da/fullfeatures-tests/0000.log`
(548,761 bytes, SHA-256
`cfe5aa3f9c4082d67019cfc8285a121768d0a5e3063f8e0c064bfcbda60d3fef`). The
owned failure was
`presentation_properties::metadata::slide_sync::codec::tests::escaped_server_id_is_counted_before_output_admission`
at codec line 464; notes and presence failures remain with their assigned
owners.

The preceding delimiter repair added the quote-bearing fragments between the
three attributes, but the root fragment still ended at `serverSldId=`. Thus
the escaped ID was emitted without an opening quote even though the following
fragment supplied the closing quote. The root fragment now emits
`serverSldId="` as well. The complete writer pair is therefore checked before
allocation and read back by `Properties::parse`: root opening quote, escaped
server ID, server-time close/open pair, client-time close/open pair, and the
suffix close. `escaped_xml_len` and `escape_xml` continue to use the same
entity byte lengths, and the selected output limit, fallible reservation, and
failure-atomic refusal remain unchanged.

No public API, selected policy, fixture, or assertion was broadened. No Cargo,
rustdoc, build, runtime, or performance command was run. The external log
identity and source base above are diagnostic provenance only; repository diff,
evidence-size, binary, and path guards are required before commit.
