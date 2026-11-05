# 0362ul50a — Composed PPT correctness consumer return

The fresh composed check of `b8897649d6364c5d4b90300bfb97f36748975fc9`
passes all-target check and Clippy for fonts, crypto, DOC and DOCX, then fails
while checking PPT. The outside-clock `scoped_patch_correctness` example
still calls `slide_order::Snapshot::from_bytes`, which is now crate-private.
This consumer existed on integration beyond the earlier PPT leaf's base;
separate leaf checks did not validate the final composition. The composed
result is failed, and downstream checks were stopped.

Both source and target snapshot reads now call the public
`from_bytes_with_options` with `OpenOptions::new()`. This preserves the
previous finite default policy, exact byte inputs, parsing and receipt
semantics. No private constructor is republished and no timed probe or
production operation changes. The corrected source requires a new guarded
PPT and downstream consumer run before composed acceptance.

The original detached capture and failed evidence remain intact under
`(deleted-coord-artifact)`. Its `f5-validation/result.json`
is 94,117 bytes, SHA-256
`282c1f878f14e5c97d9e9f7dfa8f6811067f502415a3e0c72a913a852297abe9`.
The failed `08-check-litchi-ppt-attempt-b889-f5-08.log` has SHA-256
`b947d3e757879cae5206e05cf6e5cb61afde54815de50fc4bfe012827965c2f0`.
Root verifies all nine log/trace identities, exact clean source, unchanged
root/perf/native locks, target guards and empty terminal owned-process lists.
The final target is 496,368,443 bytes. This is compiler evidence only;
actual final gate and native/performance qualification remain required.
