# 0903 — XML audit hot-loop attribution

The offline reduction of four retained 0901 publication profiles passes. It
identifies repeated attribute scanning as a bounded future hypothesis, while
leaving production unchanged after the rejected 0902 candidate. No new build,
workload, profiler, native timing, allocation or RSS measurement was run.

## Evidence and accounting

The frozen [plan](results/change-0903/plan.json) binds the four measured part-6
dumps, retained raw reader, prior analysis/review, source inventory and 0902
rejection. The [reducer](results/change-0903/reduce.py) retains every direct and
incoming edge for five selected functions, including zero-cost records.
Self plus direct-child costs equal incoming costs for every selected function;
the publication owner also equals each dump summary and total exclusive self.
Nested inclusive rows must not be added together.

First repeats, guest instructions (Ir):

| Function | Medium self | Medium inclusive | Dense-sparse self | Dense-sparse inclusive |
| --- | ---: | ---: | ---: | ---: |
| verify_source_replacement | 6,315,810 | 61,141,568 | 12,181,556 | 117,269,638 |
| verify_with_policy | 4,901,950 | 29,862,502 | 9,349,850 | 56,926,357 |
| check_start | 5,364,083 | 7,852,178 | 10,419,019 | 15,146,973 |
| inspect_attributes | 4,113,321 | 13,011,481 | 7,871,479 | 25,027,738 |
| scan_qname | 2,469,853 | 2,469,853 | 4,710,273 | 4,710,273 |

XML-source-attributed exclusive self is 25,096,782 Ir for both medium repeats
and 48,167,770 Ir for both dense-sparse repeats. The disjoint non-XML remainder
is respectively 79,300,628 / 79,300,210 and 141,333,928 / 141,333,085 Ir.
These source-file buckets differ from inclusive XML audit ownership: reader,
allocator and other callees can contribute to audit work without belonging to
the XML source-file bucket. Shared callee self is not assigned to a callpath.

## Interpretation limits

For the first medium dump, reader edges from verify_source_replacement record
46,502 calls / 13,448,131 Ir, while verify_with_policy records 116,296 calls /
13,448,055 Ir. Similarly, inspect_attributes edges record 18,642 calls /
6,504,477 Ir versus 46,631 calls / 6,507,004 Ir. The cause of these count
differences at nearly equal instruction costs is unproven. Recorded calls are
retained evidence, not a tag census, branch coverage or cost-per-call basis.

The retained reader leaves relative source positions encoded. No independently
validated callsite mapping distinguishes original and replacement audit sides
here. Inherited Valgrind brk warnings, demangled specialization aggregation and
debug attribution limits remain applicable. Instruction shares do not establish
native latency shares or predict an optimization's benefit.

## Next qualification

The [source audit](results/change-0903/source-audit.md) shows check_start
validating attribute layout before inspect_attributes walks checked attributes
again. A future zero/one-attribute summary might reuse that completed validation.
First establish cardinality coverage independently and qualify exact parser
acceptance/error ordering; the retained graph cannot supply that census.

Any candidate must preserve original-first failures, aggregate and per-element
limits at their existing logical points, namespace and QName checks, xml:space
validation, streaming/window behavior and malformed-input offsets. A zero-count
summary must not introduce a new limit check where the old loop performed none.
Higher cardinalities and xml:space can retain the existing path. Returning a
summary may change ABI/code layout, so less source work alone is insufficient.
Earlier rejected iterator experiments remain constraints, not reusable benefit
claims. Semantic parity and warning-denied quality checks must precede fresh
paired native gates; the 0902 primary threshold remains a median ratio at most
0.98 with bootstrap 95% upper endpoint below 1 for both protected shapes.

## Validation and disposition

[Canonical analysis](results/change-0903/analysis.json) and
[verification](results/change-0903/verification.json) record successful
conservation and byte-identical offline replay. All 10,771 selected current
source files match the retained 0894 manifest. No production code or tests were
changed, so no new compilation/testing is claimed. The duplicate replay output
was removed; no worktrees or build roots were created. Unrelated untracked files
and other worktrees were preserved.

The Office performance goal remains open: OLE2/OOXML first, ODF deferred, iWork
excluded. This batch adopts no optimization and makes no performance claim.
