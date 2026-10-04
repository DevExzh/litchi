# Change 0903 evidence

This packet reduces the four retained 0901 measured Callgrind dumps to
function-level XML-audit evidence. It is an offline reanalysis of existing
files. The packet does not launch a workload, build, profiler, annotation
tool, or worktree, and it makes no native-time, allocation, RSS, speedup, or
production-adoption claim.

`plan.json` is frozen and binds the four exact `callgrind.out.6` files, the
0845 raw reader, the 0901 analysis/plan/review, the 0894 source inventory of
10,771 files, and the rejected 0902 decision. `reduce.py` verifies every
descriptor before parsing. It imports the retained `reader_v2.py` grammar and
does not modify that reader.

Run the canonical reduction once from the repository root with a new output
path:

```sh
python3 -B docs/performance/results/change-0903/reduce.py \
  --output docs/performance/results/change-0903/analysis.json
```

The output is create-only: an existing output is an error. A successful report
has schema `litchi.xml-audit-hotloops-analysis-0903.v1` and status `pass`.
The report retains each selected function's exclusive self Ir, every raw
direct edge (including calls, inclusive Ir, file context, raw position token,
and serialized reader record line), including zero-cost records, and every incoming edge for the five
target functions. Positive self-plus-direct-child costs are reconciled with
positive incoming costs for each target; the publication owner is additionally
reconciled to its single measured runner edge and the full dump summary.

For each dump, the report lists every function whose retained source file is
under `crates/xml-minifier/`. Their exclusive self costs form a disjoint XML
bucket. The non-XML exclusive self remainder is checked against the full
publication-owner inclusive Ir. Nested inclusive rows are never added to this
partition, and no shared callee self cost is assigned to a callpath.

The report preserves the raw outgoing edge evidence from
`verify_source_replacement` to `verify_with_policy`, but deliberately leaves
original-versus-replacement attribution undetermined. The retained reader's
edge position is a raw Callgrind token; relative positions are not decoded,
absolute positions are not resolved, and no source-line claim is made.
Recorded call counts remain Callgrind edge fields. They are not XML token
counts, branch coverage, or a basis for interpreting cost per call.
