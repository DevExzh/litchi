# Coordinator application binding

The coordinator selected `candidate/selected-model.patch`, which is
byte-identical to `candidate/model.patch`, and verified it with
`git apply --check` before applying it. Formatting produced
`candidate/applied-model.patch` and the complete `candidate/applied/` source
archive. The three source differences between `candidate/files/` and
`candidate/applied/` are formatter-only line wrapping/blank-line changes in
`core/tests.rs`, `escher/semantic.rs`, and `escher/tests/groups.rs`; the
changed tokens remain identical.

| applied artifact | bytes | SHA-256 |
| --- | ---: | --- |
| `selected-model.patch` | 18824 | `35a7af255660310db93214333b2b58247bcf2cf85fc9b6a85bfa3c4476d9f4bf` |
| `applied-model.patch` | 18824 | `b9ea9b31f2daff2513aeb9e735324d6aab81671f1a7385a9a011abb929136d59` |

The final applied source hashes are bound by `build-after/source.json`; the
original archive receipt remains unchanged. All six quality gates passed.
The final measured candidate was rejected and production source restored;
see `disposition.json` and the final report.
