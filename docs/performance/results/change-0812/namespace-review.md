# 0812 namespace/attribute traversal feasibility review

This is a read-only design review for the event-local shared namespace and
attribute walk suggested after the 0811 capture. It does not establish a
performance benefit, and it does not authorize a production change.

## Current ordering

The current notes scanner reads a `Start` or `Empty` event, updates the
namespace scope, and then inspects the element. In
[`scan_processed_xml`](../../../../crates/litchi-pptx/src/notes/codec.rs#L358),
`NamespaceResolver::push` is called before depth/node checks and before
[`inspect_element`](../../../../crates/litchi-pptx/src/notes/codec.rs#L465).
The inspector resolves the element QName and root conformance first, then
iterates `checked_attributes`, skips namespace declarations, counts ordinary
attributes, decodes their values, and resolves their expanded names.

This ordering is observable. A declaration on the same start tag must be in
scope when resolving that element. Namespace-resolver errors therefore happen
before element/root errors, and those happen before an attribute iteration or
attribute-value error. The resolver also has to be in the new scope before
resolving attributes. The empty-element path marks the scope for a pop on the
next event; the start-element path keeps it until the matching end event.

## What the pinned APIs actually provide

The workspace pins quick-xml 0.41.0 in
[`Cargo.toml`](../../../../Cargo.toml#L168), with the resolved package recorded
in [`Cargo.lock`](../../../../Cargo.lock#L3809). Its public
[`NamespaceResolver::push`](https://docs.rs/quick-xml/0.41.0/quick_xml/name/struct.NamespaceResolver.html#method.push)
does its own `start.attributes().with_checks(false)` pass. It increments the
scope, adds every namespace declaration until a lexical attribute error, and
enforces the resolver's per-element declaration limit. The public
`add`, `set_level`, `level`, and `max_declarations_per_element` methods expose
pieces of that behavior, but there is no public operation that accepts an
already-running attribute iterator or combines `push` with caller-side
attribute validation.

The project-side
[`BytesStartExt::checked_attributes`](../../../../crates/litchi-opc/src/xml_attributes.rs#L46)
is deliberately a different iterator. It preserves quick-xml's first 32
checks, switches to the project's ordered duplicate check, and stops after the
first error. Its documented behavior and transition are in
[`CheckedAttributes`](../../../../crates/litchi-opc/src/xml_attributes.rs#L81).
This is the iterator used by notes through the re-exported
[`xml::attributes` extension](../../../../crates/litchi-ooxml-common/src/xml/attributes.rs#L15).

Consequently, the present implementation necessarily has two logical walks of
the start tag's attributes: quick-xml's resolver pass and the project's
checked-attribute pass. They are not interchangeable:

* The resolver pass is unchecked for duplicate names and applies later
  namespace declarations after a duplicate. It stops only at its first
  lexical iterator error.
* The checked pass reports the first duplicate or lexical error and stops.
  It is also the pass that enforces the notes scanner's aggregate attribute
  count and byte budget.
* The element QName cannot be resolved until all declarations on the same tag
  have been installed, while the current error order resolves that QName
  before the checked attribute pass reports an error.

## Can one walk be written with the existing APIs?

Preserving the ordering contract would require additional state and
reimplementation of part of quick-xml; there is no demonstrated drop-in
one-walk implementation under the stated constraints.

A hand-written pass could call `set_level(level + 1)`, inspect each
`unchecked_attributes` item, call `NamespaceResolver::add` for declarations,
and implement the bounded duplicate check itself. That would have to duplicate
all of the following behavior:

1. the resolver's declaration-count limit (because `add` does not enforce the
   per-element limit);
2. reserved-prefix and reserved-URI errors, including partial scope mutation;
3. the checked iterator's first-error position and recovery rules;
4. the resolver pass's treatment of duplicate namespace declarations; and
5. the empty-element/start-element scope timing.

Even after duplicating those details, a streaming pass sees attributes before
it can resolve the element QName. Deferring the QName check until the pass
finishes lets an attribute error win where the current code lets a bad element
QName win. Resolving before the pass gives the wrong answer when the element's
prefix is declared on that same tag. Staging the first attribute error while
continuing to install declarations can reproduce some cases, but it changes
the checked iterator's stop behavior and still requires special handling for
lexical errors and duplicate declarations.

For a concrete semantic divergence, consider a same-tag prefix declaration
followed by a duplicate declaration of that prefix. The existing resolver pass
applies both declarations before the inspector resolves the element; the
checked pass then reports the duplicate. A combined iterator that stops at the
duplicate sees only the first binding, so it can produce a different root
namespace result or a different error. If a later declaration has a reserved
prefix, the existing resolver can report that resolver error before the earlier
duplicate is reported by the checked pass; a combined checked iterator reports
the duplicate first.

The alternatives are therefore:

* retain the two walks;
* add a quick-xml API that lets the resolver consume a caller-controlled
  declaration stream while the checked iterator reports attributes; or
* introduce and thoroughly test a project-owned resolver/attribute state
  machine that explicitly preserves the current error and scope contract.

The latter two are dependency or substantial parser-state changes, not a
local optimization. They would require buffered differential oracles covering
namespace timing, duplicate and malformed attributes, reserved bindings,
root/refusal ordering, limits, and empty-element scope before any measurement.

## Review decision

Reject the shared namespace/attribute traversal as the next local candidate.
The public pinned API does not provide the required fusion point, and the
obvious manual fusion either duplicates security-sensitive parser behavior or
changes refusal ordering. The 0811 samples can motivate localization, but
they do not make this semantic risk acceptable and provide no causal savings
claim. Keep the two walks while the current scanner/event-handling
investigation proceeds; revisit only with an explicit API/state-machine design
and differential evidence.
