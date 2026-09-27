# 0796 source and protocol review

Status: source/protocol review passed with one documented coverage limitation.
The production tree is unchanged and root has completed the frozen probe build
and semantic preflight before native capture.

## What is sound

The packet freezes 31 inputs in two timed modes.  The cases cover 11 distinct
attribute-count boundaries (0, 1, 4, 5, 8, 9, 16, 17, 32, 33, 64), five
valid duplicate tails at the inline/vector/map boundaries, six 4,096-byte
duplicate tails (quoted, unterminated, and unquoted at prefixes 1 and 33),
and nine syntax tails (`flag`, `tail=x`, and `=\"1\"` after prefixes 0, 4,
and 33).  This is 31 cases, and construct/consume gives the intended 62
case-mode combinations.  The native schedule has six alternating before/after
blocks with 30 samples after three warmups and 4,096 iterations; the profile
schedule has two alternating repeats with one sample and one iteration.

`baseline.rs` is byte-for-byte equal to the inherited OPC helper and
`candidate.rs` is byte-for-byte equal to the retained 0794 candidate archive.
The current `source.json` also matches the tracked source tree at the packet
base revision; the production helper files have not changed.  The independent
manifest uses quick-xml 0.41 directly and keeps both helpers as local modules,
so before and after are selected in one identical release binary.  Capture
starts a fresh process per case/mode/leg/block, pins CPU 12, and uses the same
preconstructed `BytesStart` for both legs.

Case generation, the semantic oracle, clone checks, and result validation are
outside the timed owner.  Construction black-boxes an iterator reference and
drops it without advancing it; consumption black-boxes each yielded item,
checks lengths and the first-error marker, and stops at the first error.  The
runner is selected before the clock, with only the same one-call function
pointer dispatch and timer read included in every sample.  The profile owner
names correspond to the four `#[inline(never)]` leg/mode functions.  The
construction checksum is leg-independent by design; full borrowed key/value
bytes, error variants, positions, and clone behavior are checked outside the
clock by the oracle.

The raw quick-xml reference is wrapped in a cloneable `FailFast` iterator.  It
stops after the first error so it models `checked_attributes`'s documented
contract rather than quick-xml's recovery behavior.  `collect_trace` and the
timed drain also stop at that error, and the post-trace `None` calls check the
actual fused behavior.  This fixes the initial review finding that raw
quick-xml recovery could otherwise make the oracle compare different
contracts.

## Closure status

1. **Case content is independently frozen.**  `fixtures.json` contains the
   literal 31 inputs, categories, counts, byte lengths, and SHA-256 values.
   `fixture_check.py` reconstructs each generated source value and compares
   every field and hash in order.  `build.py` freezes both files in
   `build/inputs.json`, and `validate.py` runs the same check before accepting
   captures.  A probe edit cannot silently replace the measured inputs while
   preserving only the case IDs.

2. **The standalone manifest and lock are frozen.**  `probe-src/Cargo.toml`
   and its generated lock are retained.  `lock-generation.json` records the
   offline generation and hash, `build/inputs.json` freezes the generation
   receipt and probe lock, and build/check/clippy all used `--offline --locked`.
   The workspace lock remains separate from this independent manifest.

3. **The cfg(test) boundary is explicit.**  The probe README records that the
   copied helper modules retain external cfg(test) declarations, that this
   standalone binary excludes cfg(test), and that cargo test/--all-targets are
   not used.  Fresh semantic evidence comes from the independent oracle;
   helper unit-test evidence is correctly identified as inherited from 0794.

4. **Recovery-tail coverage is a declared limitation.**  The cloneable
   fail-fast adapter and first-error drains are correct, but all frozen timed
   errors end at the error, so a valid attribute after a raw quick-xml error is
   not independently exercised.  The probe README records this limitation;
   no broader recovery claim should be made in the final report.

These closures are evidence-protocol controls, not production adoption
requests. 0796 must leave the production candidate rejected and unchanged
regardless of the diagnostic timing or Callgrind result.
