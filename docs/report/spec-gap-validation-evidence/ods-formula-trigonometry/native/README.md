# Upstream cached-result observations

`cached-results.json` retains 115 scalar formulas and their numeric caches
from 24 LibreOffice mathematical-function fixtures. `provenance.json` records
the upstream commit, each complete input SHA-256, and each observation's sheet,
row and column. The extraction script accepts a LibreOffice source checkout;
run it with a separate output directory to compare a reproduction without
overwriting these receipts.

All 24 local input hashes were checked against the raw files at the pinned
upstream commit. The extractor refuses different input hashes.

The source is [LibreOffice/core](https://github.com/LibreOffice/core/tree/d804d6aff49054bad1719ec3c2d136b545bbc7e7/sc/qa/unit/data/functions/mathematical/fods) at commit
`d804d6aff49054bad1719ec3c2d136b545bbc7e7`, under
`sc/qa/unit/data/functions/mathematical/fods/`. The extracted test data retains
the upstream MPL-2.0 license (see `LICENSE-MPL-2.0.txt`).

Selection is explicit: the first called function must match the fixture name,
all calls must belong to the implemented trigonometric family or TRUE/FALSE,
the expression must contain no reference, and the cell must have a numeric
cache. Duplicated upstream observations remain duplicated. One selected
`ATAN2(0;0)` observation is recorded separately as a profile variance:
OpenFormula permits either zero or an error; the evaluator selects an error.
Other unsupported-function, reference, text/error-cache and metadata cells
are outside this numeric comparison.

The public test evaluates the retained formulas and compares the numeric
caches with relative tolerance `1e-13`, reflecting their decimal serialization;
cached zero requires exact zero. This is limited cached-value corroboration,
not execution of LibreOffice, native acceptance of generated files, full
fixture evaluation, or complete OpenFormula conformance.
