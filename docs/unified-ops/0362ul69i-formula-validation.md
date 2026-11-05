# 0362ul69i — Formula raw-policy validation receipt

Date: 2026-10-25. Parent: 0362ul69h. Status: source and validation complete;
the coordinator still owns integration and independent review.

Final source is `c937149a3b03a0acabc0ca9a4cef7372116e81a8`. The approved
runner is `1904177040fc7c523582433028d9bba56efc30fe0d9d96beb554a2fab343691e`.
The hydrated root lock is the exact 4100 lock, SHA-256
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`; the
original tool locks remained byte exact (`tools/native-resave` SHA-256
`51c25d3d6f36e20b50f0b19e39b1aa9c7b49777154122ecc95d3acd05c041378`,
`tools/perf-baseline` SHA-256
`eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c`). No
lockfile was generated.

The final Formula all-target and focused boundary lane passed under offline,
locked, one-job, one-test-thread, debug-zero, incremental-zero, decimal
19,000,000,000-byte live and 20,000,000,000-byte hard target guards:

    (deleted-coord-artifact)/10-formula-final-validation-attempt-formula-final-validation.log
    33241 bytes, SHA-256 902029b946632a680f38f205239469aa19d4436c61a15a68697033e132599f17

The all-target command reported 348 library tests and nine LaTeX integration
tests; the focused reruns reported 62 MTEF writer tests, three MTEF converter
boundary tests, and one OMML split Text/CDATA test. Every result was passing.
The target trace is 364 bytes, SHA-256
`76b071a7933f834d871860a5cf2457f09bc975a26fe13ae902cbde8c8e1c43eb`.

The affected DOC/DOCX, ops, and facade consumer lane also passed. It ran on
the production-equivalent `8447db0b4bb52b609ae7fa051bb347dce0292295` source
before the final Clippy-only source/test cleanup: the production cleanup
removes an identical MTEF charge branch and the test cleanup replaces two
single-element clones with `from_ref`.

    (deleted-coord-artifact)/06-formula-affected-consumers-attempt-affected-consumers.log
    612157 bytes, SHA-256 e2edffbd960a553dd8186c615374a222c9df119b477ea7de05116c9c58f9385e

The lane emitted 687 passing test-result blocks and no compiler or test
failure. Its target trace is 546 bytes, SHA-256
`8753daf49148a4ff4bc23e3052ee694da0aff0f7ccde9de94f81ad4674c3060a`.

The final strict-Clippy lane passed all seven scoped commands: Formula,
DOC, DOCX, ops `doc,formula`, ops `docx,formula`, facade `doc,formula`, and
facade `docx,formula`, each with `--all-targets -D warnings`:

    (deleted-coord-artifact)/09-formula-affected-clippy-rerun2-attempt-affected-clippy-rerun2.log
    0 bytes, SHA-256 e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855

The Clippy target trace is 651 bytes, SHA-256
`a395e7de708a8729ad21a7a39152b1249f65ba26e995ab8a9da046487f2ab58d`. The
earlier strict-Clippy blockers and their repairs are recorded in 69g and 69h;
no warnings remain in the scoped lane.

The final runner target reached 5,429,565,120 bytes and retained 378,564,894,720
free bytes. No native, rustdoc, broad-ops, or performance command was run in
this validation.

The earlier invocations above were recorded only as child logs and target
traces; no contemporaneous JSON rows existed, and none has been reconstructed.
To close that evidence gap, a fresh slot-1 run persisted its manifest and each
`run_step` return while reusing the compiled target:

    (deleted-coord-artifact)/manifest.json
    2768 bytes, SHA-256 ca0524a1551a3aea830cb20e9d048499f6e9dcdd5856d35676e18847d35e8300
    (deleted-coord-artifact)/rows.json
    21236 bytes, SHA-256 c6d289bb01b4aa1ed175dc1b8690d27eed400c7b9c2bbaf59acf013b735860b6

The persisted rows are `row-00.json` (4718 bytes, SHA-256
`12bdf427f91f823984e4d80f7f7157cb72df4f303ab0f980b3a4107799fa3834`),
`row-01.json` (5270 bytes, SHA-256
`bdbbf621243f2694093bf054d457ff20bd9480bd2ff531f3e082e4407927fda1`) and
`row-02.json` (10422 bytes, SHA-256
`df902fbbd7bf658f20488dfe0abdab0d73e2acaa4cfdf54eb324a3e6d8e7f93c`). They
all report exit zero, the final source commit above, and the exact three lock
hashes.

The fresh row-00 Formula log is 33125 bytes (SHA-256
`c3e7d258b6a53af5753228959088e9db230643f22dacecaa0a32c916899febc2`) with a
364-byte target trace (SHA-256
`1bff0345898a4fdd478d96eb45acd1a9aeef936b7f01aba6f47c19e5d391b782`). Row-01
consumer testing is 600471 bytes (SHA-256
`29140aba12192509b49940db9fcc6247795d46438772dffbae69d43010a9f1a4`) with a
546-byte trace (SHA-256
`09c6cef42e80000720768d19d082ef43f0b4470fa01a576ba9dce195697b7b72`).
Row-02 strict Clippy is an empty 0-byte log (SHA-256
`e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`) with a
651-byte trace (SHA-256
`0d553ea4adafe7214244ce981c98ff427fc088e21a1ccbee2a31c385f91ea0da`). The
manifest records decimal 18,000,000,000-byte preflight, 19,000,000,000-byte
live and 20,000,000,000-byte hard target guards, plus slot-1 ownership and
release after all three rows reached terminal success.
