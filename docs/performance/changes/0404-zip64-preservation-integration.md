# Change 0404: ZIP64 preservation integration

Date: 2026-09-04

`performance_claim: none`

`claim_authorized: false`

## Problem and mechanism

At control revision `9c6742c5212dd0e7ff2367da585abe357aae8975`, the
ZIP reader retained ZIP64 field origins and the preservation writer could
rewrite validated ZIP64 framing under an internal policy. Public preservation
constructors still rejected ZIP64. Consequently an OPC edit could read a valid
ZIP64 source but could not publish a preserving changed artifact.

The integration uses that existing validated writer through the public
preservation constructors. The physical owner retains unchanged local member
spans, including compressed bytes and descriptors, and original central
records except the local-offset fields that must move. Existing ZIP64 end
records retain their representation and extensible bytes while the affected
counts, sizes, and offsets are updated. OPC continues to own package planning
and semantic validation; ZIP framing remains in `soapberry-zip`.

ODF source-backed content publication now obtains declared sizes from the
validated ZIP preservation entry instead of treating fixed 32-bit sentinel
fields as payload lengths. Its canonical `mimetype`, local-name, descriptor,
source-version, and no-padding guards remain in force. The ODF regression
fixture promotes ordinary central records and the tail while retaining the
required canonical `mimetype` record.

Archive-level ZIP64 context also permits both 32-bit and 64-bit descriptors
when central sizes have no ZIP64 sentinels. Required ZIP64 size fields still
require 64-bit descriptors; ambiguous matching interpretations refuse. The
reader bounds candidates by the central-directory boundary and retains its
checksum/size error ordering.

This is a necessary capability for the goal's unchanged-member ZIP64
passthrough requirement. The control refuses the changed operation, so there
is no comparable successful control latency and no speedup claim. Large-file
latency, allocation/RSS, actual device I/O, and representative producer
measurements remain follow-up work.

## ADR compliance

| Requirement | Implementation boundary |
| --- | --- |
| ADR 0001 / 0002 / 0010 / 0011 / 0024 | No archive implementation type moves into a semantic facade; framing stays in the existing ZIP owner. |
| ADR 0003 | Source snapshots remain unchanged; package publication retains its existing validation and conflict contracts. This change does not add a new semantic patch format. |
| ADR 0005 | Output uses the existing sequential caller sink and bounded raw-copy window; metadata remains subject to the preservation index limits. |
| ADR 0006 | Unsupported or ambiguous framing still refuses before emission; retained unknown metadata is copied from the source. |
| ADR 0008 | Public integration needs focused preservation and downstream tests; low-level self-roundtrips alone do not certify native Office workflows. |

## Remaining boundaries

The writer does not promote generated member sizes or offsets to ZIP64.
A generated member whose representation needs that promotion still fails
before the sink is written. ZIP32 sources are not automatically upgraded to
ZIP64. Multi-disk, prefixed, unresolved, overlapping, truncated, and ambiguous
layouts retain their typed refusals. This batch does not certify native Office
interoperability or complete the large-offset corpus and cross-platform gates.
ODF’s additional framing guards still refuse local-header ZIP64 sentinel/extra
layouts and unsigned descriptors whose CRC equals the descriptor signature;
these compatibility cases remain open and do not silently normalize.

## Validation

The all-feature test command covers libraries, integration tests, and doctests
for `soapberry-zip`, `litchi-opc`, and `litchi-odf-common`. ZIP has 362 passing
library tests, OPC has 316 including the topology test,
and ODF has 287. The full run passes 1,227 tests with two ignored doctests. The ODF integration includes the new ZIP64 sequential-sink
publication test, all 14 streaming package tests, and all 5 encryption-authoring
tests. ZIP/OPC warning-denied Clippy and the three-crate formatter check pass.

The combined Clippy command reaches an existing `large_enum_variant` failure
in `litchi-odf-common/src/package/model.rs:230`, outside this diff. The check is
recorded as failed; no allocation change or lint suppression is introduced to
hide it. This batch also repairs the ODF root-versus-manifest MIME validation
boundary and a source-text test that mistook an error message for unsafe Rust.

The [correctness evidence](../results/change-0404/validation.json) records exact
commands, source hashes, results, and compressed logs. These checks do not
measure performance or certify native-producer compatibility.
