# 0362ul101c — ribbon test import qualification

The joint compiler pass from integration `6964022294` reached the common
ribbon integration tests and reported one unnecessary qualification:
`litchi_opc::SourceLimits::default()` at `tests/ribbon.rs:33`, although
`SourceLimits` is already imported by the test module. The call now uses the
existing import; behavior and the required source-limit policy are unchanged.

This is a source-only lint repair. No Cargo, build, rustdoc, runtime, native,
or metadata command was run on this leaf; the coordinator should rerun the
joint compiler gate against the committed tip.
