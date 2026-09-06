# Source, semantic and reservation review

OPC owns the opaque retained capture. Consuming an authorized token splits its
lifetime: the shared capture and decoded reservations remain, and writer staging
is released. Each `authorize_for_publication` call reserves `C + F` independently,
checks lineage/version/context before and after admission, and returns a consumed
publication token. Destination-name reservations remain per token. No physical
ZIP implementation type leaks into PPTX or a public semantic API.

Original authorization still admits a total `2*C + F`, with `F >= 4096`, now as capture, writer-payload and fixed-overhead reservations,
with the same checked aggregate overflow refusal. Failure releases every prior reservation. Retention drops writer payload staging,
leaving `C + F` plus managed decoded `U`. `F` also covers long source URI/content
type metadata and token storage; normal fixture names retain the 4096-byte floor. The fixed charge bounds retained
token metadata even for empty captures; repeated promotion cannot create
unlimited uncharged retained handles.
Clones share those reservations; N live publication tokens reserve N independent
`C + F` writer/overhead budgets in addition to the retained handle. Exact-budget refusal is retryable after another token drops. Tests cover
concurrent complete writes, no source read/decode work during reauthorization,
source-change precedence over cancellation, and final memory/object release.

PPTX retains captures beside its existing staged decoded-byte allocations; this
batch does not remove those copies or relax the existing aggregate destination
staging reservation. First preparation uses combined capture/read. Publication
reruns preparation, checks all metadata and bytes, and can reuse a capture only
when the existing helper returns the exact previously staged allocation. Prepared
payload equality compares semantic identity and bytes, independently of this
optional physical optimization. The touched digest, candidate source rereads,
chart XML validation, dependency closure, source-checked sink and partial-output
classification remain unchanged. Only memory refusals allow decoded fallback.

The public destination editor remains consuming and bound to one opened OPC
lineage; a fresh editor is not a way to publish the same public plan again.
Reusable retained OPC captures are needed because the planner rerun borrows the
original plan. No mutable one-shot token or Clone editor is introduced. OPC tests
exercise concurrent independent publications of one retained capture directly.

Source freshness relies on the existing immutable ReadAt/version contract.
The retained bytes are immutable and already physically verified; reauthorization
does not reread compressed source bytes. Candidate validation still compares the
current decoded source view with staged bytes and checks source version/lineage.

No dependency, unsafe policy, pool or ambient I/O change. Accepted ADR tree remains
c950b6c8be822561b498d7bbe87c460873dcbf49 (0003/0005/0006 lifetime, budgets,
preservation; 0002/0010/0011/0024 ownership). Native/default CRUD coverage counts
are not promoted by synthetic provider measurements.
