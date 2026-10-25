# 0362ul89aq — PPTX runtime notes repair

This fresh repair leaf is based on `ef9470f7da08a131f31600041f9d562203d62bfa`.

The borrowed notes-scanner oracle now applies the accepted selected notes
policy to owning `presentation` roots, preserving parity with the production
scanner; only the related `sld` root retains its independent 16 MiB and
one-million-node admission. Slide-root proof reuse now receives the selected
semantic notes profile separately from the source-constrained scan limits, so
the physical 64 KiB attribute ceiling cannot accidentally invalidate an
otherwise exact default proof, while tighter semantic profiles still force a
bounded reclassification.

Opened-capture regressions retain the legacy capture-local comparator for
well-formed relationship, root, notes, and deferred-name behavior. Public
opened capture tests now explicitly assert the documented content-free
`PPTX XML is malformed` refusal when duplicate attributes are rejected by the
physical pre-MCE admission pass. No physical admission guard is weakened.

Source-only rustfmt and diff checks are required before handoff. Cargo,
native/runtime qualification, rustdoc, and the combined gate remain
coordinator work. No generated evidence, fixture, executable, or `3rdparty/`
path is included.
