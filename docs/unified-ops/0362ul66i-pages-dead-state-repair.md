# 0362ul66i — Pages dead-state repair

Status: source READY for the next scoped F5 retry.

The fourth F5 run reached the Pages library and stopped on `-D warnings`:
`State::semantic` was stored during package construction but had no consumer.
The selected semantic limits are already applied during projection, so the
unused retained copy was removed from the immutable state and its initializer.

This is storage cleanup only; parsing, projection limits, and public APIs are
unchanged.
