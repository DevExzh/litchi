# ODF 1.4 Part 4 function-catalog review

This review is bounded to standard function-name recognition. It does not
claim expression-grammar, parameter-arity, evaluation, recalculation, or
host-defined-function conformance.

The normative source is the bundled ODF 1.4 Part 4 formula document from
baseline `21ad91183`:

| artifact | SHA-256 |
| --- | --- |
| `3rdparty/specs/OpenDocument-v1.4-os.zip` | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| `part4-formula/OpenDocument-v1.4-os-part4-formula.html` | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |
| `part4-formula/OpenDocument-v1.4-os-part4-formula.pdf` | `c49a0cf2cd8c57606875807aeb4b6a6fd8e0666754c937019375c142699baabd` |

The reproducible command is:

```text
python3 docs/report/spec-gap-validation-evidence/ods-formula-functions/extractor.py \
  3rdparty/specs/OpenDocument-v1.4-os.zip --output /tmp/normative-functions.tsv
```

The extractor examines chapter 6 function-family `h3` headings in the HTML
body. There are 408 headings in sections 6.5 through 6.20, including 15
`General` prose headings (section 6.17 has no separate `General` heading).
Removing those prose headings leaves exactly 393 function headings. Every one has a `Syntax:` paragraph,
has a unique section anchor, and has a name matching the Part 4
`FunctionName` shape after removing presentation whitespace. The section
counts are:

```text
6.5  5   6.6  5   6.7  7   6.8 26   6.9 12   6.10 24
6.11 2   6.12 54  6.13 33  6.14 11  6.15  9  6.16 69
6.17 8   6.18 86  6.19 16  6.20 26                         = 393
```

The candidate chapter-6 set in `crates/litchi-ods/src/codec/formula/functions.rs`
contains the same 393 unique names: no catalog row is missing and no name is
invented. The previous 161-name set is a strict subset. The HTML has two
small presentation/anchor defects that are not additional functions:

* `6.8.22 IMSEC` is split as `I` and `MSEC` in adjacent text nodes. The
  extractor removes that presentation whitespace and retains
  `#a_6_8_22_IMSEC` (and `#IMSEC`).
* `6.12.8 COUPNCD` has the section anchor
  `#a_6_12_8_COUPNCD` but no separate name anchor; an anchor for that name is
  adjacent to the preceding `COUPDAYSNC` heading. The section anchor is used
  for this row. It does not create a duplicate function.

The PDF body confirms the same function sections and signatures. Its table of
contents contains a stale `6.10.8 EDATE` entry, while the body has the
normative sequence `6.10.8 EASTERSUNDAY` and `6.10.9 EDATE`; the HTML body and
Appendix A agree on `EASTERSUNDAY`.

Appendix A (PDF p. 213) lists one new function, `EASTERSUNDAY` (6.10.8), and
these twelve changed definitions:

```text
CONVERT COUNTA INDEX ISBLANK ISFORMULA ISLOGICAL ISNONTEXT ISNUMBER ISREF
ISTEXT NPER PMT
```

The other 380 rows are not listed as changed by Appendix A. The eight
`LEGACY.*` names are exact standard chapter-6 names:

```text
LEGACY.CHIDIST LEGACY.CHIINV LEGACY.CHITEST LEGACY.FDIST LEGACY.FINV
LEGACY.NORMSDIST LEGACY.NORMSINV LEGACY.TDIST
```

Neither chapter 6 nor Appendix A specifies aliases for these names (or for
any other row). The TSV therefore records `normative_aliases=none-listed` and
does not invent unprefixed aliases such as `CHIDIST` for
`LEGACY.CHIDIST`. The legacy prefix is part of each standard function name.

The lexical boundary is set by Part 4 §5.6 (PDF p. 40): `FunctionName` starts
with `LetterXML` and continues with `LetterXML`, `DigitXML`, `_`, `.`, or
`CombiningCharXML`; function names are case-insensitive; a call has
parentheses, and the parameter list may be empty. No particular Unicode case
folding algorithm is specified. The catalog is consequently stored in its
published uppercase ASCII spelling. Standard-name lookup should normalize
ASCII case while keeping the broader XML identifier grammar separate; Unicode
characters that happen to uppercase to ASCII must not become invented standard
aliases. Part 4 §5.7 (PDF p. 40) separately governs nonstandard names and
their origin prefixes.

Part 4 §5.8 (PDF pp. 41–42) makes references bracketed (`[` begins a
reference), which keeps references distinct from function names. `LOG10`
(6.16.41) and `BIN2DEC` (6.19.4) therefore remain ordinary catalog names;
digits are permitted after the initial letter. A token followed by `(` is a
function invocation, while a bare token may be handled by the broader
identifier/named-expression grammar. This review does not expand that parser
grammar.

The principal HTML anchors used for spot checks are
`#a_6_10_8_EASTERSUNDAY`, `#a_6_8_22_IMSEC`, `#a_6_16_41_LOG10`,
`#a_6_19_4_BIN2DEC`, and `#a_6_18_11_LEGACY_CHIDIST`. The chapter-6 body
starts on PDF p. 55; §5.6/§5.7 are on p. 40, §5.8 starts on p. 41, §5.14 is
on p. 46, and Appendix A is on pp. 212–213.

Disposition: no pseudo-heading, alias, or catalog-completeness blocker was
found in this bounded review. The implementation preserves the prior lookup helper's Unicode-uppercase
behavior through bounded normalization. Those compatibility spellings are not
additional catalog entries. The tokenizer's broader XML identifier grammar
remains incomplete and is not certified by this catalog review.
