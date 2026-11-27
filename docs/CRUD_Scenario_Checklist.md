# CRUD Scenario Checklist

This checklist is the support-certification gate referenced by ADR-0008. It is
a reusable evidence template, not a statement that every row already passes.
Each format-specific feature matrix must link to tests or other reproducible
evidence for every applicable supported direction. Marking a row `N/A` requires
a format- or feature-specific rationale.

The checklist applies the API, preservation, security, and resource contracts
from ADR-0003 through ADR-0008. A package capability does not establish a
semantic capability, and fresh authoring does not establish editing of an
opened document.

## Claim and API gate

| Check | Required evidence |
|---|---|
| [ ] Scope is explicit | The matrix names the exact format, package flavor, standards revision, and supported semantic subset. |
| [ ] Read and write are independent | Read-only, fresh-write, opened-document edit, and opaque preservation are not conflated. |
| [ ] Public API is semantic | Focused modules use short names, checked value types, typed selectors, and typed errors; native IDs and raw wire fields stay at the codec boundary. |
| [ ] Unsupported states are explicit | Unsupported, ambiguous, noncanonical, protected, stale, and over-limit inputs fail without silently dropping data. |
| [ ] Feature isolation is explicit | The required Cargo feature set is documented and builds without unrelated format families. |

## Read and snapshot scenarios

| Check | Required evidence |
|---|---|
| [ ] Open valid input | At least one representative producer file opens through the public facade. |
| [ ] Reject wrong family | MIME, container, relationship, root-element, or record-family mismatches return a typed error. |
| [ ] Snapshot is immutable | Cloneable snapshots share immutable source state; ordinary public methods cannot mutate attached document state. |
| [ ] Selection is checked | Name, index, address, relationship, and object selectors distinguish absent, ambiguous, and invalid values without panicking. |
| [ ] Unknown data is bounded | Unknown XML, records, streams, and extensions are retained within explicit limits or rejected before allocation. |
| [ ] Malformed input is covered | Truncation, duplicate ownership, invalid ordering/cardinality, bad lexical values, and hostile nesting have focused negative tests. |

## Create scenarios

| Check | Required evidence |
|---|---|
| [ ] Minimal valid artifact | The smallest public builder output reopens through Litchi and a relevant independent producer or consumer when available. |
| [ ] Deterministic output | Identical semantic inputs and explicit options produce identical bytes; ambient time, randomness, locale, or network state is not consulted. |
| [ ] Compact XML output | Every generated XML part is byte-minimal: no pretty-print indentation, non-semantic formatting whitespace, indentation-only nodes, whitespace between tags, or optional space before `/>` or `>`. Required separators remain intact, and all semantic character data is preserved byte-for-byte. |
| [ ] Referenced XML is compact | Relationships, content types, manifests, metadata, styles, settings, and every other generated or rewritten referenced XML part follow the same compact rule. |
| [ ] Package topology is valid | Content types, manifests, relationships, owner cardinalities, part names, and orphan rules are checked before publication. |
| [ ] Authoring is bounded | Builders validate semantic counts, strings, opaque payloads, aggregate bytes, and final output size before committing allocations or I/O. |

## Opened-document CRUD scenarios

| Check | Required evidence |
|---|---|
| [ ] Create semantic owner | Add the first owner and another owner at every supported placement boundary. |
| [ ] Read after create | Reopen the serialized artifact and compare the complete supported semantic state. |
| [ ] Update each field family | Replace scalars, optional values, lists, references, and opaque extensions independently and in composition. |
| [ ] Delete first/middle/last/only | Remove owners across boundary positions and remove the final owner/container where the format permits it. |
| [ ] Reorder or move | Stable semantic identity, references, positional metadata, and package ownership remain correct. |
| [ ] Exact no-op | An unchanged commit returns byte-identical output and preserves signatures and opaque content. |
| [ ] Atomic failure | Any validation, allocation, protection, or I/O failure leaves the published snapshot and destination unchanged. |
| [ ] Reversible patch | Forward application checks its source, inverse application restores exact source bytes, and application to another artifact reports conflict. |
| [ ] Concurrent composition | Independent edits join deterministically or return structured conflicts; no last-writer-wins mutation occurs. |

## Preservation and compatibility scenarios

| Check | Required evidence |
|---|---|
| [ ] Unknown content survives | Unmodeled attributes, children, records, streams, namespace bindings, and package members survive supported mutations byte-exactly where promised. |
| [ ] Lossless-or-refuse is enforced | A writer refuses mutation when it cannot prove preservation of adjacent opaque or producer-specific content. |
| [ ] Strict/transitional variants | Applicable namespace, relationship, package, and lexical variants are tested without prefix-sensitive matching. |
| [ ] Signature policy is explicit | Exact no-ops preserve signatures; changed writes either resign under explicit caller policy or remove/refuse stale signatures. |
| [ ] Encryption policy is explicit | Password opening, re-encryption, unsupported profiles, and protected mutation have separate typed outcomes. |
| [ ] Independent producer evidence | Real files and save/reopen round trips identify producer versions and the exact scenario certified. |

## Validation, limits, and performance scenarios

| Check | Required evidence |
|---|---|
| [ ] Caller-selected limits | Input, output, node/record count, depth, string, opaque payload, relationship, and aggregate ceilings are configurable where untrusted data is accepted. |
| [ ] Exact boundaries | Each important limit has below, exact, and above-boundary tests, including decoded/entity-expanded values where applicable. |
| [ ] Pre-allocation checks | Encoded lengths and counts are rejected before proportional allocation, decoding, decompression, or cloning. |
| [ ] Iterative hostile traversal | Attacker-controlled recursion uses bounded iterative traversal or proves a safe depth ceiling. |
| [ ] Hot paths are indexed | Graph lookup, overlap detection, ownership checks, and repeated edits avoid accidental quadratic behavior. |
| [ ] Allocation failures are atomic | Fallible reservation and serialization failures do not leave queued or partially published mutations. |
| [ ] Performance claims are measured | Throughput, allocation, zero-copy, lazy, streaming, or memory-efficiency claims cite a reproducible benchmark or are omitted. |

## Permanent non-execution boundary

| Check | Required evidence |
|---|---|
| [ ] Macros and VBA remain inert | Code may be detected, inventoried, validated, preserved, or edited as data but is never compiled, interpreted, or executed. |
| [ ] Controls and actions remain inert | ActiveX, form controls, actions, event bindings, add-ins, and embedded code are never activated or dispatched. |
| [ ] External targets remain inert | DDE, database commands, mail merge, links, schemas, media, and remote relationships are never fetched, refreshed, followed, or executed. |
| [ ] Formula behavior is capability-bound | Parsing and serialization do not imply evaluation; any evaluator accepts only explicit caller capabilities and a documented function subset. |
| [ ] Embedded payloads remain data | OLE objects, media, scripts, and nested documents are not opened by host applications or handed to execution runtimes. |

## Workspace release gate

| Check | Required evidence |
|---|---|
| [ ] Formatting and lints pass | Workspace formatting plus warning-denied Clippy gates pass for the certified feature combinations. |
| [ ] Tests pass in isolation | No-default, required-feature, all-target, doc-test, and relevant all-feature configurations pass. |
| [ ] Dependency boundaries pass | `tools/check_crate_boundaries.py` reports no unclassified, forbidden, or stale edges. |
| [ ] Metadata and examples pass | Cargo metadata resolves and consumer-facing examples use only public facade paths with correct `required-features`. |
| [ ] Matrices match evidence | Every supported row names its boundary; known gaps remain `🟡` or `❌` rather than being inferred from shared crates. |
| [ ] Diff and generated XML are audited | The integrated diff is inspected, compact XML fixtures pass the repository auditor, and no iWork result is used as evidence for non-iWork certification. |
