# RTF native adapter composed review

The reviewed leaf `2b794c52702eefef8c96f7eecc6256c3884f41e1` migrates both the standalone historical/current probe and the live current crate example. Timed replay now measures native apply only; serialization and expected replacement, remainder, reopen, inverse, exact no-op and tail-append checks run outside that interval. All eight routes remain.

Independent review verifies the recorded locked compile/check-only logs for baseline, candidate and live current consumers. Direct route diagnostics contain positive medians, but they are not quiet paired ABBA acceptance. Record [0362ul8b](0362ul8b-rtf-native-adapter.md) preserves the original failed current compile and all later evidence.

The merge is clean onto DrawingML integration `04058bc7e13bf227d6adaf2165fb67c64ca40a40`. All 42 composed tooling tests pass at index `98fb42576fc886c47107a11422506c032a53e205`; [receipt](results/0362ul8c/composed-tooling.json) records the exact command and external log identity. Final assembled workspace and native qualification remain required. No performance acceptance or push is claimed.
