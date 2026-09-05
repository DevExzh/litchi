# 0424 source review

Review basis: frozen worktree diff against `5b176dc9a`, the 0424 staged-payload
reuse protocol baseline. This is a source and test review only; no benchmark,
Cargo, or profiler command was run by this reviewer.

The low-level API boundary is appropriate. `SourceTopologyPlan` already owns
`Arc<Vec<u8>>` payloads internally, and `litchi-opc` already exposes shared
immutable Part payloads through `Part::blob_arc` and `PartFactory::load_shared`.
The new `try_add_part_shared` remains below the PPTX facade, keeps the plan
opaque, and factors the existing URI, operation-count, duplicate, content-type,
and fallible vector-reservation checks through one helper. It does not transfer
source authority, ZIP state, or a managed reservation into the topology plan.

The PPTX reuse predicate checks source URI, target URI, lexical content type,
declared size, and decoded bytes before cloning the staged Arc. Source/version/
lineage checks, the full `Prepared::matches` value comparison, candidate reread,
writer preflight, cancellation fences, and both staging reservations remain in
place. `bytes_equal_checked` restores the old chunk-level cancellation checks
for the new equality pass. The original plan remains borrowed throughout
publication, so its payload handles and `_memory_reservation` cannot be dropped
before the shared topology is consumed; retaining both logical reservations is
conservative and should remain documented as such.

The initial focused PPTX compile reported an unused
`PreparedChart::payload_identity` helper; the coder has since used it in the
chart identity and declared-size fallback cases. That historical failure is
resolved. The final focused PPTX reuse test passed, including exact image and
chart identity/fallback and cancellation behavior. The final public
source-backed image/chart and adversarial suite also passed 59 tests, while the
OPC topology suite passed its 21 tests.

Coverage is now meaningful across both payload kinds. The OPC unit test proves
shared Arc identity, duplicate/root/content-type/operation-bound checks,
delayed payload construction, and Arc copy-on-write ownership. The focused PPTX
test covers all four image metadata fields, same-length byte fallback, chart
identity and declared-size fallback, and cancellation. The 59 public
image/chart tests exercise publication output, and five independent harness
cross-copy output oracles provide the corresponding generated-media check. The
dedicated OPC test remains an ownership/refusal test rather than a second
serialization mirror; the public image/chart tests and independent oracles
cover shared publication bytes.

No public-layer, source-authority, partial-sink, or reservation-lifetime
blocker was found. The low-level shared API
should continue to describe payloads as caller-owned immutable Arc handles;
callers that need a handle after the call must pass a clone. No managed-budget
reduction or general zero-copy claim follows from the unchanged full logical
reservation charges.

The final source-scope evidence reports `verify_candidate`, `reserve_memory`,
and `Prepared::matches` byte-identical to the control. Strict lint findings
remain confined to the four pre-existing locations, with the established
command-local exemptions, and rustdoc with warnings denied passed.
