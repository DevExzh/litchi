# 0362ul121a — Restore selected-closure replay scope

The first 0362ul121 repair added a package-wide part-count guard to make the
stale-replay test fail.  That rejected a valid patch recipient with an
unrelated part, contrary to the Ribbon contract: replay is checked against the
recorded selected Ribbon closure and must preserve unrelated package content.
The guard is removed.  The runtime coverage now accepts and verifies an
unrelated XML part while mutating the selected Ribbon part to exercise the
actual `PatchConflict` boundary.

The content-types provenance repair remains narrow.  The shared OPC token
replacement now distinguishes “same bytes with the same valid retained
token” from “same bytes with no retained token,” and installs the recorded
token in the latter case through the existing checked API.  Ribbon destination
validation still accepts an authored-to-retained transition only when the
retained `[Content_Types].xml` bytes exactly equal the recorded target bytes.
Forward removal replay now covers this missing-token case; inverse restoration
asserts both the semantic manifest bytes and retained token bytes, alongside
the restored binary image and Ribbon XML.

This is a source-and-regression-test follow-up to the frozen 121 runtime
capture.  No Cargo, build, rustdoc, runtime, native, metadata, or performance
command was run; the coordinator must rerun the scoped Ribbon/common gate.
