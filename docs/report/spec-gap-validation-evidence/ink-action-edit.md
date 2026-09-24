# Detached InkAction construction and editing

The shared `litchi_drawingml::ink::actions` API supports bounded construction
and source-preserving structural edits of the strict action profile. The
builder and edit types are re-exported from `actions`; `actions::edit` is also
public.

## Usage and ownership

Use `Draft` with `ActionDraft`, `ActionGroupDraft`, `ActionDataDraft`, and
property/data child drafts to construct a new value. `finish()` produces a
`Prepared` value after bounded readback.

For an existing fragment, call `actions::read_profile(bytes)`, construct
`Edit::new(profile)` (or `Edit::with_limits`), and select source actions with
`ActionSelector`. Stage scalar or structural operations, then call `finish()`
to obtain a `Commit`. The commit retains the prepared result and a
source-checked `Patch`; `patch().inverse()` reverses it against the exact
forward result. Selectors refer to the captured source, not shifting indexes
after each staged insertion. Equivalent selectors coalesce scalar updates;
sequential setters within one edit retain their final value.

An unchanged edit replays its exact source. Changed output preserves untouched
source spans, including comments, namespace spelling, and opaque InkML bytes.
Removal and `clear` account for IDs inside removed descendants and refuse
retained references that would become dangling. Removed nodes release final
node budget before additions are admitted.

## Payload and limit contract

`OpaquePayload` retains one bounded element for definitions, transforms, or
trace data. Source edits require payload-local namespace declarations.
Detached construction additionally admits the canonical `iact` and `inkml`
prefixes supplied by the generated root. Explicit empty or wrong bindings
are not interpreted as inherited canonical bindings.

Authored payload staging validates element and attribute qualified names,
namespace resolution, duplicate expanded attributes, XML characters,
references, comments, and CDATA. Legal `<?` bytes in comments or CDATA remain
opaque data. The complete candidate is also read back before publication.
`Limits` bounds source/output bytes, payload bytes, nodes, depth, action/group
counts, and scalar lengths below the shared reader's hard ceilings. Custom
action types include the empty string permitted by the schema's string branch.

## Evidence and remaining scope

The focused gates are reproducible with:

```sh
cargo test --locked --offline -p litchi-drawingml --lib --test ink_action_edit --test ink_action_id_boundaries
cargo clippy --locked --offline -p litchi-drawingml --lib --test ink_action_edit --test ink_action_id_boundaries -- -D warnings
```

At integration, these cover 171 library tests, 21 edit tests, and 12 boundary
tests. The boundary tests include escaped local references, IDs inside opaque
payloads, exact no-ops/inverses, typed payload roots, depth, XML grammar,
namespace staging, and `clear` closure/node accounting.

This owner does not discover or author a PPTX package relationship, infer an
InkAction Part identity, execute actions, recognize handwriting, or implement
the full InkML semantic model. Native-host compatibility is not claimed.
The separate `ink-action-edit-performance` harness remains subject to its
freeze and verification gates; correctness tests are not performance results.
