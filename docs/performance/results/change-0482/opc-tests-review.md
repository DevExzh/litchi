# OPC splice integration-test review

Status: focused integration coverage added; execution is intentionally left to
the root coordinator because this batch shares the workspace build and
measurement locks. The test contract has been aligned with the current API:
changed signed sources refuse during preparation, and exact no-op proofs must
authenticate their source and candidate hashes before the physical copy.

`crates/litchi-opc/tests/source_part_splice.rs` exercises the public
`SourceBackedPackage::prepare_source_part_splice` boundary for both Store and
Deflate members.  The fixture also carries an untyped `scratch.bin` member so
the changed local record can be compared with an untouched physical member.

The test contract covers:

* exact no-op reproduction of malformed XML and signed physical packages;
* changed signed-source refusal before sink output;
* independent rejection of malformed source and candidate XML before plan
  publication;
* source, fragment, candidate, offset, length, and digest proof rejection;
* source-version mutation, managed cancellation, and finite candidate-limit
  refusal;
* no materialized Part-cache load while preparing and publishing;
* Store/Deflate replay, raw opaque-member preservation, partial sink progress,
  flush errors, and exact inverse authorization;
* rejection of a foreign candidate artifact by the inverse authorization;
* inverse publication freshness fences when the current candidate changes on a
  sink write or final flush, plus managed cancellation at both points;
* large binary-fragment Work admission and cooperative cancellation while the
  fragment-only phase advances in bounded chunks;
* inverse output-budget refusal from the current candidate context, independent
  of the context that authorized the original publication;
* shared-context inverse accounting, requiring one output-budget charge for the
  retained original archive length;
* typed XML, preservation, and replay workspace quotas, exact XML quota
  boundary behavior, and managed reservation release after successful plans or
  rejected quota checks.
* binary `.bin` Part replay without applying XML grammar rules to a
  non-XML content type.

The malformed-source/candidate cases are intentional contract probes. The
OPC path now exposes the independent source and complete-candidate audit
seam, so both checks must remain before the preservation writer touches the
sink. A consistent proof hash alone is not an XML validation authority.
The binary case keeps the same decoded hash/length and physical replay
contract while requiring XML audit classification by both part name and
content type.

The accounting case asserts exactly two complete decoded reads for the target
member during replay: one for the low-level measurement pass and one for the
emission pass. Store uses `stored_payload_bytes_read`; Deflate uses
`deflate_bytes_produced`. This assertion is deliberately separate from the
preparation passes, whose accounting report is not supplied by the caller.

The sink assertions accept the established `IncompleteOutput` wrapper or the
direct typed I/O error while requiring accepted-prefix progress.  The final
root run should tighten this to the exact public precedence once the replay
and source-freshness fixes settle, then retain the command receipt and test
output with the rest of the 0482 evidence.
