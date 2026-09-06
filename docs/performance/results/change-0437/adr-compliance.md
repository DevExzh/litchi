# 0437 ADR compliance

| Constraint | Implementation and evidence |
| --- | --- |
| 0001 API layers and priorities | Opt-in format-owned fresh plain-slide API; existing Builder and rich-slide authoring contracts remain available. Correctness and bounded refusal precede performance. |
| 0002 / 0024 dependency direction | ODP owns page/frame/text grammar; common owns generated XML validation and package publication; archive implementation stays behind common. Boundary receipt covers the integrated source. |
| 0003 immutable edits and atomic commits | Fresh caller-sink publication introduces no snapshot edit, patch, join, source identity, or commit behavior. Partial sink output is explicitly invalid on failure and must be discarded. |
| 0005 explicit I/O and execution | Caller supplies the sequential Write sink and ExecutionContext. Finite input/object/Work/output budgets and modeled memory reservation are checked; there is no ambient filesystem, networking, executor, or global pool. |
| 0006 preservation and validation | New API accepts only explicit title/body strings. It cannot silently flatten rich slide fields. Fixed default page/frame/style/meta grammar follows the existing Builder and is independently checked. Common strict envelope behavior remains unchanged; the prelude extension is opt-in and audited. |
| 0008 verification status | Source-only review, root serialized checks, retained failed attempts, independent fixture/report gates and exact source/binary custody distinguish implementation from measured claims. Native runtime is unavailable and visual rendering is unverified. |
| 0010 / 0011 archive ownership | Publication uses PackageWriter typed authored/generated XML paths. The provider does not concatenate raw ZIP records or expose physical archive identities. Stored first mimetype and manifest/member contracts are checked by the benchmark gates. |

The API reserves a modeled provider window, two fixed shell copies and bounded
metadata. It does not claim that this reservation includes every ZIP/auditor
allocation or caller-owned input. Operation allocator measurements and
whole-executable RSS are reported separately. No unsafe code is introduced;
existing forbid(unsafe_code) remains in force. No production XML audit ceiling
is relaxed to admit the rejected 32,768-title baseline proposal.
