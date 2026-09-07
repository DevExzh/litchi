# Change 0461 source review

## Verdict

ACK. The frozen `validation.rs` split is semantically sound and I found no
source-level blocker. `ElementAttrs::get` now uses a pure cached key predicate
and calls `decode_value` only after a namespace/local-name hit. The production
diff is limited to that lookup path and its focused tests; the drawing-attribute
harvest path is unchanged.

## Semantic checks

- `matches` keeps namespace comparison before local-name comparison and uses
  the cached resolver outcome. Unknown and unbound namespaces therefore remain
  misses, while qualified names and namespace shadowing retain their existing
  behavior.
- Cached-prefix hits decode lazily and return the same formatted value error.
  A scanned matching attribute is decoded before it is appended to `parsed`,
  preserving the existing behavior in which a fresh decode failure is not
  cached. A cached decode failure remains replayable.
- Non-matching attributes are never value-decoded. Iterator advancement,
  document order, first-match behavior, and the stored malformed raw-attribute
  error path remain unchanged.
- `drawing_attributes` still uses the existing cached namespace snapshots and
  harvest order, including modeled-attribute skips, foreign/unknown prefixes,
  lazy values, and shape-specific error messages.

## Focused coverage

The added test uses an independent one-shot lookup helper and exercises both a
cached invalid value and an invalid value first reached by the shared iterator.
It verifies matching valid values skip unrelated invalid values, fresh error
messages are preserved, cached decode errors replay, and the newly reached
decode failure retains the historical no-cache/iterator-advance behavior.

The reviewer did not edit Rust, run builds, run tests, or run CPU jobs. Owner
build, strict Clippy, focused/full tests, and the candidate performance gates
remain the authoritative validation for this batch.

## Generated-code comparison

The authenticated release disassembly confirms the intended shape. In the
baseline, `ElementAttrs::get` calls `ElementAttrs::lookup` for each cached
attribute (baseline `get` line 37) and again for each newly scanned attribute
(line 109); the helper returns the `Result<Option<String>>` transport even for
misses. In the candidate, the `lookup` symbol has no body in
`candidate-lookup.asm.txt`, and `get` performs the cached namespace/local-name
checks directly at lines 33--51 and the newly scanned checks at lines 111--129.
Those miss branches contain no `decode_value` call. The candidate calls
`decode_value` only on the cached hit path (line 54) or scanned hit path (line
201), after both comparisons succeed.

The candidate `get` stack frame is 0x128 bytes versus 0x148 bytes in the
baseline, consistent with removing the helper's result transport from the
miss loop. Apart from normal address/layout shifts and the expected helper
elimination, the retained assembly shows no extra parser or value-decoding
work on misses. This is generated-code evidence for the mechanism only; it is
not a latency claim. The matrix and allocation measurements remain required.
