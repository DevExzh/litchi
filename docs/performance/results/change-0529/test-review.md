# 0529 shared XML attribute-audit test review

This handoff contains focused tests only. The 0529 production candidate is
the shared `xml-minifier` attribute probe in `audit.rs`; it has not been
changed by this task. OLE2/OOXML remains the active priority and ODF stays
deferred.

## Bound source and artifacts

The tests are based on revision
`3f4c7be06159dc3d742d9c815800dc55bf7c2db6`.

| Item | SHA-256 | Git blob / detail |
| --- | --- | --- |
| `crates/xml-minifier/src/audit.rs` base | `40c7dfea199aa8f2e1951d0d75cd9a47e84f13bf49f2abcd1232230b64b355ac` | `fb1ee700a1b218ef832d582dbd946be2fc78a529` |
| `crates/xml-minifier/tests/stream_audit.rs` base | `d0152fcffb883c0611c40c73933b52c0ddfc47223705202964dc8a0d933b5f2f` | `cee29f50820cce6c663dcda6550523135ea9c034` |
| frozen 0529 `production.patch` | `ca0970a03ab76e0bd6ebffd4ced9501fadb656ac4bb55acb515a2371d471e509` | candidate source below |
| `tests.patch` | `ca6f6b88d4f28658aebaa14f225ca1dd7efbdadf2eda3ad99bc12412f08e949f` | public baseline-compatible patch |
| `candidate-only-tests.patch` | `06cb0c1a1b2c07b5faf14f16968ac63553daa1e41cde7bf32d07ab52388755bc` | candidate-only private oracle |
| candidate `audit.rs` source | `507feadd1d48e4a5a1c26a009137f8ec46ff7a9d922d0ea5a5b6c1f11229ca89` | `/home/zhuhe/litchi-goal-0529-draft/.../audit.rs` |

## Patch coverage

[`tests.patch`](tests.patch) changes only the existing public
`crates/xml-minifier/tests/stream_audit.rs` module. Its four tests exercise:

- zero, one, many, and namespace attributes, with exact aggregate counts;
- `xml:space` preserve/default transitions, entity-normalized values, and
  the multi-attribute checked path;
- duplicate `xml:space` before an invalid value, invalid `xml:space` values,
  and slice/stream parity for those error results; and
- inclusive per-document attribute limits, aggregate counting across tags,
  and attribute-limit precedence over a later whitespace defect.

Every new public case calls `verify_authored` and compares it with
`verify_authored_reader` over five chunk schedules. The public fixtures keep
lexical tag layout valid so the test reaches the attribute inspector; the
existing lexical-layout suite remains the oracle for `check_start` failures.

[`candidate-only-tests.patch`](candidate-only-tests.patch) appends one unit
test to the existing private `stream_tests` module. After
`production.patch` is applied, it compares `inspect_attributes` directly
with `inspect_attributes_checked` for:

- zero/one/many attributes and namespace declarations;
- `xml:space` state overrides and malformed values;
- duplicate-before-invalid-value ordering;
- invalid attribute syntax as the first and second attribute;
- zero and one attribute-limit boundaries; and
- a generated 40-attribute tag crossing quick-xml's checked duplicate
  checker threshold.

The comparison checks both the returned `Result<Space, Error>` and every
`State` field, so a probe cannot increment counters, alter inherited state,
or leave scratch state behind before proof. The checked helper remains the
behavioral oracle; no parser is copied into the test.

## Application and validation boundary

Apply `production.patch`, then `tests.patch`, then
`candidate-only-tests.patch`. The first patch remains the source candidate;
the second compiles against the base public API; the third intentionally
requires the candidate-only `inspect_attributes_checked` helper and must be
applied only after the production patch.

Only static checks were performed: `git diff --check` and `git apply --check`
against minimal source copies. No build, test, benchmark, capture, or
production-source edit was run for this handoff. Root owns all baseline and
candidate builds and test execution.
