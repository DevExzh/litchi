# 0362ul165b — ODF raw-codec protocol review return

Independent review returns isolated leaf
`48d68bcc29fb56afd3c20c54f6e518374a5e3d46`. It is not merged or
execution-ready. Seven Python tests, AST checks and rustfmt pass, but they do
not establish valid preparation, Rust behavior or native qualification.
No Cargo, archive, native or measurement command has run for this leaf.

The review identifies eleven findings:

1. The manifest pins production C2a while preparation expects the leaf's
   endpoint, which contains the new consumer. These identities cannot pass
   the current checks. Keep C2a frozen and bind the separate consumer source
   explicitly rather than weakening or advancing the production pin.
2. Preparation records a Cargo manifest, lock and probe overlay; the runner
   accepts only lock and probe. The exact composition must agree everywhere.
3. The namespace negative case assumes eager validation and panics before
   routes run. It must drive the reader and check the actual typed failure.
4. The memory writer proof requires a deterministic no-op output to differ
   from its input. It must prove the intended exact no-op or an explicit edit.
5. Facade publication likewise inverts the no-op invariant. Require exact
   bytes and reopen the published result.
6. Manifest and namespace-reader timing boundaries omit or include operations
   contrary to their descriptors.
7. Splice, facade-reopen and writer descriptors also disagree with the timed
   work. Setup, copies, proofs and measured operations need explicit alignment.
8. Splice preservation slots use literal `true`, and facade proof does not
   validate the published package. Compare actual opaque/untouched members
   and metadata and validate publication.
9. API/fixture descriptions do not match the probe; CLI preparation-to-runner
   composition has no positive test. Bind the actual fixture and API variants
   and exercise the complete handoff, with endpoint negatives.
10. Manifest validation admits mutations to timing and proof metadata.
    Pin or validate the full canonical protocol and reject such mutations.
11. Preparation changes baseline files before validating output freshness
    and location, and receipt publication is not exclusive/atomic. Validate
    first and publish the complete receipt safely.

The author is repairing these findings on the same isolated branch. Every
affected manifest, preparer, runner, probe and test must migrate together.
The original commit and failed review remain available. Source-only checks
may continue while the existing full65 run owns the build slot; Rust builds
and actual native qualification remain pending.

The external review is
`(deleted-coord-artifact)/odf-raw-native-final-48d68-independent-review-return-20261026.json`,
12,428 bytes, SHA-256
`f0added9c5380415486344803e089c1f0c7e8ca39a55b06444c956320ebc50ce`.
It records exact file locations, scope, checks and required repairs.

The full 25-row qualification scope in 165 remains required. No capability,
fidelity or performance exception is accepted, and nothing is pushed.
