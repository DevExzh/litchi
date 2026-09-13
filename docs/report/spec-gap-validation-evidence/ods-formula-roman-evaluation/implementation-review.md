# ROMAN/ARABIC implementation review

Verdict: no correctness or bounded-resource blocker remains in the reviewed
draft. This is an independent review of the ODF 1.4 Part 4 §6.19.2 and
§6.19.17 implementation; it does not replace the crate gates or integration
tests.

The checked-in ODF archive has SHA-256
`9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`, and its
formula entry has SHA-256
`ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`.
The numerical check reads that retained ZIP and hashes the exact formula HTML
entry; it has no dependency on temporary extracted text. The reviewed source
digests are:

* `crates/litchi-ods/src/codec/formula/evaluation/roman.rs`:
  `3441233624f68e5efc72157953d3b7137bbb86f6c9835fd809eee40674342da1`.
* `crates/litchi-ods/src/codec/formula/evaluation.rs`:
  `5b2320a6204b96185b6e91eaa27d3f166ac19555e71b307c79a0600329975cf4`.
* `crates/litchi-ods/tests/ods_formula_roman_evaluation.rs`:
  `4bcc00fd7278787c1ae77a4bd58118677a743a92bf023f6f725309bd6b8877ff`.

## Findings

`ARABIC` follows the specified indirect-subtraction rule. The right-to-left
maximum fold subtracts a symbol exactly when a strictly larger symbol occurs
to its right, and the byte matcher accepts only the seven listed ASCII Roman
letters in either case. Empty text returns zero. The independent checks cover
`IIX = 8`, `IXX = 19`, `IC = 99`, `MIM = 1999`, lowercase input, and rejection
of whitespace, signs, Unicode Roman numeral characters, and other letters.

Formats 0 through 3 use the literal table permissions: format 0 permits only
`I`, `X`, and `C` subtraction; format 1 adds `V` and `L` with the ten-times
ratio cap; format 2 adds `L` but excludes `V` and removes that cap; format 3
adds both `V` and `L` without the cap. The finite greedy chunk model was
checked for all 4000 values. Every spelling round-trips through indirect
subtraction and satisfies the direct-pair plus trailing-symbol condition. The
literal ODF reading intentionally gives `ID` for 499 in formats 2 and 3 and
`XLV` for 45 in format 2; Excel compatibility spellings such as `XDIX` and
`VDIV` are separate profile choices.

For format 4, the implementation enumerates all 64 choices of the two
minimal representatives of each residue modulo `[5, 2, 5, 2, 5, 2]`. The
normalization argument is sound: transferring one adjacent radix from a
coefficient whose magnitude is at least that radix decreases the L1 digit
count by at least one, including when the next coefficient has the opposite
sign. Thus a minimum has only the nonnegative residue or its negative carry at
each lower denomination. Negative coefficients are emitted in ascending
denomination order before positive coefficients in descending order. The
highest negative denomination has a larger positive symbol to its right, so
the ODF indirect rule evaluates the emitted text to the signed coefficient
sum. The top `M` coefficient is nonnegative for the constrained input range.

The independent integer BFS in
`numerical-review/check_roman.py` searches every signed denomination at every
depth through 15, retaining the complete prefix-sum interval
`[-15000, 15000]`. Fifteen is a conservative upper bound because the classic
representation of every value below 4000 has at most 15 symbols. This gives a
global lower bound even for signed combinations that are not valid Roman
output. A separate BFS tracks the highest negative and positive denominations
and emits canonical negative-then-positive witnesses. Results for the full
`0..3999` scope are:

* all 4000 unrestricted targets are reached, with maximum shortest distance
  11;
* all 4000 canonical-sign targets are reached;
* canonical and unrestricted minimum lengths match for all 4000 values;
* all 4000 canonical witnesses round-trip through the independent ARABIC
  fold;
* the source-shaped residue construction round-trips and reaches the
  unrestricted distance for all 4000 values.

The BFS is independent of the implementation's residue recursion, so the
minimum claim does not rest only on the six-level candidate enumeration. Tie
spelling is deterministic by first-minimum mask order but is not prescribed
by §6.19.17; equal-length spellings such as `IMM` and `MIM` for 1999 are both
valid.

## Resource and failure review

The Roman output is constructed in fixed 64-byte arrays. Formats 0 through 3
make bounded progress, and format 4 has exactly 64 masks and six levels per
mask. Each candidate/level or greedy denomination check charges existing Work
and therefore checks cancellation. `ARABIC` charges the input byte length and
checks cancellation every bounded 4096-byte scan interval. No value-indexed
allocation or ambient I/O is introduced.

Output text checks the configured text limit before charging Work or reserving
Memory. The fallible `String` reservation is held together with the returned
`TextValue`; error paths drop the partially constructed output and reservation
without leaving live storage. The input `TextValue` remains live through the
ARABIC scan and is dropped before the result is pushed. Numeric format text
coercions use the evaluator's existing bounded conversion path. Formula errors
are propagated before conversion, while unsupported references remain typed
evaluator refusals.

The string scanner's cancellation check now runs once per outer 4096-byte
window; a doubled quote may consume one byte into the following window, so
quote-pair recognition remains unchanged. The focused concatenation test
covers boundaries at 4094, 4095, 4096, and 8191 bytes, including a multibyte
literal. The numerical receipt was generated without Cargo or Rust execution:
`numerical-review/receipt.json` records the exact command, script digest
`6e70059c7c9220ca72c5837d06ccf6751608df95137d8210e145844291703e36`, source
digests, full-domain counts, and sample witnesses. The receipt itself has
SHA-256 `5392283e7dbcd377108fcdd6d9e9a6d2c790df586b8ab729020ccdb33ae7e4a5`.
