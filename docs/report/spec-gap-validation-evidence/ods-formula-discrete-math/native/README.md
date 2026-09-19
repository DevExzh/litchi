# Upstream discrete-math observations

`cached-results.json` retains 77 selected numeric observations across the 11
requested functions: `COMBIN`, `COMBINA`, `FACT`, `FACTDOUBLE`, `GCD`, `LCM`,
`MULTINOMIAL`, `EVEN`, `ODD`, `DELTA`, and `GESTEP`.  The observations come
from the LibreOffice mathematical and add-in FODS fixtures at commit
`d804d6aff49054bad1719ec3c2d136b545bbc7e7`.  `provenance.json` records the
raw URL, every exact input SHA-256, selected coordinates, and explicit profile
exclusions.

The extractor reads only the 11 listed source files and bounds every reference
closure before emitting it.  Each referenced cell is retained as a literal
number, text, logical, or empty cell.  A reference to an upstream formula cell
is refused, so the Rust test cannot evaluate a source formula or silently
import its cached result as an input.

The retained receipt was generated from a temporary source tree populated with
the exact raw files at the pinned commit.  Reproduce it, including the download
and byte checks, with:

```text
python3 reproduce.py
```

`reproduce.py` downloads all 11 raw files into a bounded temporary source tree,
verifies each SHA-256 against `provenance.json`, regenerates both JSON
receipts, compares them with the retained files, and removes the temporary
tree on exit.  It does not modify a checkout.  The repository's auxiliary
LibreOffice tree, if present, is not assumed to have the pinned bytes.

The exclusions preserve source evidence without promoting host behavior to a
normative result.  They include the `MULTINOMIAL` formula-cell dependency,
fractional `LCM` operands accepted by the native cache while the resolved
contract rejects them, domain-error `COMBIN`/`COMBINA` cases, text/error
GESTEP and DELTA cases, and the `COMBINA(0;0)` host variance.  Fractional
`MULTINOMIAL` row 19 is retained: its cached value is 10 and agrees with the
resolved raw-sum-before-floor contract, although that particular pair does not
distinguish raw-sum conversion from per-argument flooring.  The pinned GESTEP
fixture has no logical-valued input, so logical conversion is explicitly
unobserved rather than inferred.

The public integration test compares finite caches with relative tolerance
`1e-13`, requiring exact zero.  This is bounded cached-value corroboration,
not LibreOffice application execution, file resave or acceptance, recalculation,
full-fixture evaluation, or complete OpenFormula conformance.  Upstream fixture
data is distributed under the retained MPL 2.0 license.
