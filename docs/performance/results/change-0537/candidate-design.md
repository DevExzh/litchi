# Draft: borrow transient raw worksheet attributes

The unapplied candidate changes only raw/worksheet/codec.rs. It retains the
Cow returned by quick-xml for the transient reference and numeric fields.
The retained PendingCell cell_type still calls into_owned at its existing
storage boundary. No parser event, namespace, attribute or validation pass is
skipped; no raw model or public API changes.

Pinned quick-xml returns Cow tied to Attribute's input lifetime, not to the
short borrow of Attribute itself. Borrowed input may remain borrowed; owned
attribute bytes and normalized replacement text return owned values. Dropping
the local Attribute therefore does not invalidate the returned value. The
same normalization method and XmlError mapping run at the same points.

The complete checked scan still decodes r when encountered and defers s/cm/vm/t
until their existing semantic order. Duplicate, malformed trailing attribute,
invalid reference, style/metadata bounds and cell-type diagnostics keep their
precedence. The candidate does not change infallible String allocation into a
new error policy; normalized values may still allocate as before.

Required future guards compare old/new normalized values and exact errors for
plain, escaped, whitespace-normalized, non-UTF8 and owned Attribute inputs;
cover malformed/duplicate attributes and semantic error precedence; verify
persisted cell type survives subsequent events; and retain public readback,
unknown-byte preservation, no-op and reversible edit coverage. Compilation and
all correctness/performance gates remain unexecuted for this draft.

First add planning allocator observations to the harness, with normal-build
Unavailable status and aligned samples. Freeze baseline source with those
tests/metrics, then measure this production-only delta through fresh matched
native, allocation, profile and eager-read guards. Historical allocation-child
Ir does not establish how many allocations the candidate would eliminate.
The draft applies cleanly and is formatted with repository rustfmt.toml; it is
not an accepted optimization and is not applied to production.
