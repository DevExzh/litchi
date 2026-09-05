# Change 0418 source and compatibility review

The prerequisite revision `39e8dca8a` adds two opt-in owned PPTX cross-copy lifecycle
selectors. Both use the established plain/media corpora and exact output
oracles. Their timers include owned source/destination ingress, snapshots,
planning, atomic application and sequential publication. Corpus clones, sink
reservation, verification, reopen and teardown remain outside. Allocation
regions have that same lifecycle boundary; allocator elapsed is excluded from
latency comparisons. The default selector matrix is unchanged.

The final measured control is `79dfee5025276676d433a80b0b5475ae03f09db8`;
the candidate is `f8f9e6667284ae28fdbf9b9313b2ff9583ec74fc`. Both include the
same corrected 427-selector registry assertion. Their source diff contains
three production files and the candidate's new adversarial integration tests.
Both revisions remain reachable in the branch history.

Two unrelated failures were reproduced at the original `147b2f6da` control:
an outdated raw-XML test expected a malformed slide-list child to survive open,
and the reader rejected a missing optional slide-ID list. The control includes
the corrected refusal expectation and accepts zero or one direct slide list.
Duplicate, nested, orphan and malformed lists remain errors. Four new cases
cover missing/empty lists in transitional and strict namespaces. Microsoft's
[official SDK schema](https://raw.githubusercontent.com/dotnet/Open-XML-SDK/main/generated/DocumentFormat.OpenXml/DocumentFormat.OpenXml.Generator/DocumentFormat.OpenXmlGenerator/schemas_openxmlformats_org_presentationml_2006_main.g.cs)
models `SlideIdList` with occurrence bounds zero and one. Initial failure logs
and successful control verification are retained under `checks/`.

The production candidate retains the freshly constructed, serialized and
reopened package from application-time planning. Exact descriptor equality,
source/destination graph and physical fingerprints, patch preconditions and
postconditions, main-root checks, limits, target revision and final physical
fingerprint remain mandatory. Assignment occurs only after those checks.
An inverse retains its detached restored candidate after the existing fresh
forward-plan proof.

Independent review found that unconditional replacement by the reopened
candidate would lose caller-defined `Part` implementations and save options.
The corrected path uses the read-only OPC unmodified-owned-source predicate.
For destinations with revoked authorization it retains the previous
clone-and-apply behavior. This does not restore authorization to a changed
package: the reusable candidate owns the new artifact through ordinary owned
ingress. Custom-part sentinel behavior and nondefault font-save preferences
are tested through plan, durable forward and serialized inverse application.

The retained candidate also retains its complete archive and decoded blobs.
Plan/patch blobs can have different owners. Lower allocation volume, lower
peak RSS and lower end-of-timer live bytes must therefore be assessed
separately; none follows automatically from removing repeated compression.
The protocol requires explicit review of adverse allocation/RSS movement.

The borrowed real-producer fixture test establishes a typed provenance refusal.
It is not evidence of native PowerPoint/LibreOffice output acceptance.

The full candidate PPTX and OPC all-feature suites pass, including doctests.
The lifecycle test and both selectable-registry tests pass. All 291 final
Python comparison/catalog/index/package/claim tests pass. The comparator binds
the fixed PPTX corpora and validates sample permutations and sink observations
before removing timing-dependent values from cross-leg identity. Scoped
formatting and the crate-boundary
check pass. The boundary tool requires `RUSTUP_TOOLCHAIN=1.98.1` here because
the pinned 1.95 installation has no Cargo component.

Unqualified Clippy is not green on the control. Rust 1.98 reports
`chunks_exact_to_as_chunks` in unchanged core code and, after that exemption,
`clone_on_copy` and `needless_lifetimes` in unchanged PPTX slide-order code.
The latter two were reproduced on the matched prerequisite control. The
candidate passes all-target/all-feature Clippy with exactly those three
command-scoped exemptions. No repository lint policy was weakened.
