# Final seal corrections

The first final-seal attempt rejected a changed `final-gates.json`: refreshing
its helper receipt had changed an input bound by the frozen capture provenance.
The original manifest was restored byte-for-byte to its recorded SHA-256;
`post-capture-gates.json` retains the later helper manifest separately.

The next attempt rejected the completed ordinal-191 capture directory as if
it were still interrupted scratch. A successful resume intentionally reuses
that capture path. The seal now requires its regular directory and absent
private scratch; the existing formal collector, archive hashes, and resume
chronology still authenticate the completed child and the separate interrupted
raw files. A focused regression covers the reused directory, remaining private
scratch, missing capture, and symlink refusal. `final-helper-tests2.json`
records the passing 82-test suite with unchanged helper hashes during execution.
Neither correction changes the frozen protocol, production candidate, raw
samples, analysis, or performance scope.

The interrupted-start check also expected the full file-descriptor shape for
its protocol reference, while the capture schema stores only path and SHA-256.
It now compares that exact reference shape; full protocol size/hash custody
remains independently checked through the archived input descriptors. The
resume, profile, source, ADR, gate, and review bindings pass independently.
`final-helper-tests3.json` records the final passing 82-test helper suite.
