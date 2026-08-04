# ODI legacy migration

The files beside this document are the source-preserving relocation of the
former `litchi-odf::odi` implementation. They are not declared by the active
crate because the legacy document reader still depends on the umbrella's
generic package and shared root types.

The eventual mapping is `codec` for XML parsing, `model` for image semantics,
`package` for archive ownership, and `authoring` for edits. No compatibility
forwarding layer is added. There are currently no external ODI tests to wire;
the reader's unit tests remain with this deferred source.
