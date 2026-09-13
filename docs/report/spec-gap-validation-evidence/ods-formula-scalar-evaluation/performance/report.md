# Scalar evaluation profile

This profile uses the frozen source vector recorded by the [final gate receipt](../gates/results.json): base checkout `9072e969e34263be73d5b8d2eb2700478d182101` plus the gate’s recorded dirty source vector, with 879 tests across 52 targets, clippy, rustdoc, doctests, and formatting all passed. Builds used one sparse checkout and one release target with `--locked --offline`, `TMPDIR=/var/tmp/ods-scalar-tmp`, and `taskset -c 2`; each lane used three warmups and 15 measured iterations. The [evaluation candidate raw data](evaluation-candidate/raw.csv) has 105 rows (35 cases × 3 phases), all status 0. The [AST candidate raw data](ast-candidate/raw.csv) has 24 rows, all status 0.

The final evaluator ELF is `acf4e57d11c301cbaad0cc318dabb70638d9686be631abd51d299b6c99d580b8` (1,226,672 bytes); the AST regression ELF is `2d3ae0bdd1e174d6aaa47b8fbceb428a567c8f61f875094041e7336d6f90fab1` (1,074,616 bytes). Their [source manifests](evaluation-candidate/source-sha256.json), [harness manifest](evaluation-candidate/harness-sha256.txt), and [binary receipts](evaluation-candidate/binary-provenance.json) bind these captures to the gate vector.

The three evaluation phases have separate meanings:

- `parse` times `Expression::parse` for each repeat.
- `evaluate` parses once before timing and times repeated `evaluate_scalar` calls on that retained expression.
- `parse-evaluate` times parsing and evaluation for every repeat.

Caller execution-context and budget setup is outside each timed region. The evaluation harness consumes each returned scalar with a full byte checksum inside the timed region, so the evaluation figures include result consumption and should not be read as pure evaluator instruction latency. The [harness source](evaluation-harness/src/main.rs) performs exact numeric/text value checks and typed failure checks in an untimed preflight; stable aggregate checksums alone are not the semantic proof.

The candidate-only evaluation profile gives these representative p50 times per process (the `repeat` value is shown because each process batches that many operations):

| Case and phase | Repeat | p50 ns | Requested bytes | Peak live delta | Result |
| --- | ---: | ---: | ---: | ---: | --- |
| flat 4096, parse | 2 | 139,461 | 3,423,424 | 860,160 | 2 successes |
| flat 4096, evaluate | 2 | 790,284 | 787,040 | 196,960 | 2 successes |
| flat 4096, parse-evaluate | 2 | 909,635 | 4,210,464 | 1,057,120 | 2 successes |
| UTF-8 text/number-left 4096, evaluate | 2 | 23,730 | 17,380 | 8,641 | 2 successes |
| escaped UTF-8 text/text 4096, evaluate | 2 | 407,872 | 99,296 | 37,312 | 2 successes |
| numeric coercion 4096, evaluate | 2 | 769,544 | 1,573,472 | 393,568 | 2 successes |
| reference, evaluate | 128 | 28,190 | 28,672 | 224 | 128 typed refusals |
| work limit, evaluate | 128 | 48,481 | 46,080 | 312 | 128 `Resource::Work` refusals |
| text limit, evaluate | 128 | 35,770 | 33,792 | 264 | 128 `Resource::Memory` refusals |
| cancelled, evaluate | 128 | 850 | 0 | 0 | 128 cancellations |

The [AST baseline](baseline/raw.csv) and [AST candidate](ast-candidate/raw.csv) use the same 24 tokenizer/expression cases and harness. The initial candidate parser p50 changed from baseline by −5.65% to +11.07%. The increases above 5% were `expr-array-1024` (+5.73%, 466,372 → 493,102 ns) and `expr-reference-1024` (+11.07%, 142,480 → 158,250 ns). `expr-reference-64` decreased 5.65% (146,281 → 138,021 ns). A four-round [balanced baseline/candidate A/B capture](ast-abab/raw.csv) reduced those two apparent increases to +1.82% for `expr-array-1024` (upper medians 475,032 → 483,692 ns) and +2.04% for `expr-reference-1024` (147,790 → 150,810 ns); all 16 rows succeeded. Allocation calls and peak-live deltas were identical for every AST lane; RSS varied between process startups and host state. These figures do not establish a general parser speedup or regression.

Three candidate-only [perf-stat receipts](perf-stat/) captured the evaluator's flat 4096, UTF-8 number-left 4096, and numeric-coercion 4096 lanes. The counter samples were successful against the final evaluator ELF. The flat lane reported 71,384,711 cycles and 215,535,482 instructions; number-left reported 5,047,743 cycles and 8,319,192 instructions; coercion reported 69,755,140 cycles and 212,186,596 instructions. Exact commands, stderr, stdout, status, binary digest, and capture timestamps are retained beside each receipt.

The earlier [ASCII-name experiment](ascii-candidate-abab/raw.csv) remains separate evidence. It was rejected after improving name-only lanes while regressing ordinary string/reference lanes; its patch is not part of the frozen candidate and is not used for the AST comparison.
