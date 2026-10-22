# 0359sc — public DOC body replay conflicts

The 0359sb correction described `litchi_doc::body_text::Error::Conflict` as a
private body-editor detail. The module and error are public, and the canonical
DOC package edit, body, transfer, and durable replay routes expose that error
through their public results. This correction separates stale replay from
three-way composition:

- `body_text::Error::PatchConflict` now reports a package, plan, transfer, or
  semantic precondition whose recorded source no longer matches the receiver.
- `body_text::Error::Conflict` remains the result of committing a three-way
  plan whose branches write the same semantic facet differently.
- The DOC package edit and body documentation, internal consumers, and
  `litchi-ops` classification tests use the split vocabulary. Stale public
  `Package::apply` errors are checked through the facade and classify as
  `invalid_input`.

Parser and structural failures remain `Invalid`/`Corrupted`, and the source
package remains unchanged after a rejected replay. The source-only change was
formatted and audited without borrowing a build target; the coordinator should
run the focused DOC and facade tests on the warmed integration cache.
