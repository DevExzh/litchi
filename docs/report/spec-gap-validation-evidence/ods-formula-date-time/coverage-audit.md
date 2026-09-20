# ODS date/time coverage audit

This audit records the final retained evidence for the 24-function date/time
batch. The coverage manifest is PASS: every one of the 24 function requirement
sets and all 20 cross-cutting requirements has exact, hash-checked evidence
bindings. The verifier validates each focused test against its defining Rust
test and passing gate line, each structured receipt against its typed case
fields, and each source proof against the frozen source map.

## Scope and custody

`coverage-scope.json` is the immutable projection of the contract's 24 functions
and 20 cross-cutting requirements. It is SHA-256
`28a754ded86ad7524d0d94c5856757f775f52dee2f52dd603f8f66353fdbd114`.
The contract is SHA-256
`cc77d41f487993b3438f817dd62a359ba4ec4b3ca2359a893b31aecc79bc2c7f`.
`coverage-requirements.json` is the mutable evidence binding and is currently
SHA-256 `c4c2cd219182d4e40b5a11ecaa67ffc877a632c501b7ce0ccc4bee9a1f50a4ae`.
It contains 223 function bindings and 54 cross-cutting bindings. Every binding
is tied to a current source or retained receipt hash; identifiers were reviewed
against the defining Rust test, structured receipt fields, or oracle/native
case IDs. Broad family names are not used as substitutes for a function case.

The final source freeze is `gates/freeze.json`, SHA-256
`461a76708e36b2a716cd622df45f014e009186223bd1db511c3e0ff0f7fa3561`.
The retained gate bundle is internally verified: `gates/results.json` is
`fa5dfab0bef43d1403b22204185ae7e2e8fffe31e753c07860fb4768c228286a`,
`gates/ods-tests.log` is
`026363768e8ef5a9a289523f29357ead967bcfeaa015d14f2cd17ca2b36149bb`, and
`gates/verification.json` is
`bdeb941ccf550e43a341df081dcd54e1ba9e70779744ecb80e785cca71739243`.
The seven gate commands completed with 1,745 passed tests, zero failures, zero
ignored tests, and 130 successful test summaries.

The selected source fingerprints used by the bindings include:

| source | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation/value/date_time.rs` | `1da16aae011dd8c94fafc39fdf81cc5d9f8019024ab3860bdc00b4d28962e159` |
| `crates/litchi-ods/src/codec/formula/evaluation/value.rs` | `e2f9e5511e27e4a165ebdb7b0fba4ea8721231c610773a98f54b846c232c61f3` |
| `crates/litchi-ods/src/codec/formula/evaluation.rs` | `6db733541c9ab49e53628aab8e0fc823af8411b3a74ae47f88c50ba977446a1b` |
| `crates/litchi-ods/src/codec/formula/evaluation/timestamp.rs` | `c8ec12c46a116f48a60387e61e339f362770831b07b1c8e03c84958f134c60e7` |
| `crates/litchi-ods/src/codec/formula/evaluation/date_time/kernel.rs` | `2bc377db1a01e28dbc4940c99c61fb43a308919b315bfc37a92ce86a67e5d0cc` |
| `crates/litchi-ods/tests/ods_formula_date_time_evaluation.rs` | `d9644432fd5dcac9272dcd3f942e4b669865db22fb6fe866a706ebd05bf04ea1` |
| `crates/litchi-ods/tests/ods_formula_date_time_limits.rs` | `fbf6ba638d7b80ebcda36d1091641ef45e42be07f9bc20a987583a3987fed411` |

The independent oracle execution receipt is
`oracle-execution.json`, SHA-256
`184609c7c7b40c51702947fd2549ab4190158032549f4cc91de594e857bc1dda`; it
replays 132 rows over all 24 functions. The expected corpus is separately
hashed as `b33011089974b18b1fcc0b984adab839600acc98f09dd850e02522348316259f`.
Native results contain 132 rows and are SHA-256
`8936eda768c113a39fcfaa04e071cf6d01c1cca722b8ff6e25537c8f6c3c024a`, with
provenance SHA-256
`374716756025a74eb35b432a402091667e41dc43de0e4ab5302ae777e3bdb079`.
The native records retain explicit divergence dispositions for host-clock
observations and therefore corroborate the profile without replacing it.

Both substantive reviews are accepted in the retained receipt
`review-receipt.json` (SHA-256
`af38cc90788e881cc36032c1bd7462f395441e698f02d064d0c7e4e04adbf4e5`). The
semantic report is SHA-256
`6b29d4eb13b8f726952005e84c5d8f29baab530ac02bd160c49dbc3a12513171`; the
resource report is SHA-256
`0df60bf7a2933924db6449e4e2d5ba1cb850213f26c399cee32258c9bda0b32e`.
The source-proof receipts are hash-bound to the same freeze, including the
adapter, resolver-authority, kernel, profile, timestamp, WORKDAY offset
conversion, DATEDIF DateParam, and ISOWEEKNUM DateParam/domain proofs. The
DATEDIF receipt is SHA-256
`61e62be7e8a1bec7e3f11640c4678db110465dda16792c5136f4820d8e2fabcb`, and the
ISOWEEKNUM receipt is SHA-256
`e1ad0e5027e5de80519524b49216bd0ebe1a676e2c6b992af03b604149e21e98`; both
select one independently reviewed proof ID from the shared date-parameter
report, SHA-256
`836c5fbfa24ec9aa7a6d04ba69bbe23f2eb7d5eadf440e06698fab9d344f7126`.

The retained performance bundle has two captures and 4,620 rows. Its verified
report is SHA-256
`6a2876f11ea81206ceadab02d886cfa675ebbd9d0e3fc24df290e2b9497bd3fe`, and the
root audit is SHA-256
`70a8d706c5a1f414ea35a7556182ee56cbdfbed306d4e8d356bd3590de1cda0d`.
The independent review accepts the capture descriptively. It records a median
latency trigger of +5.228% in the 4x4 SIN parse/evaluate lane and 22 median RSS
triggers ranging from +184 to +240 KiB; all matched accounting sets remain
exact. These observations support no causal regression or speedup claim.

## Function mapping audit

The manifest uses the following direct anchors. Generic arity and argument-error
bindings point only to tests that enumerate the complete family. Optional-slot
bindings point to `required_missing_and_optional_empty_slots_keep_distinct_profiles`.
Formula-error bindings point to the argument-error test that enumerates all
argument-taking names. Function-specific rows use the matching oracle IDs.

| function | direct semantic/resource anchors | oracle anchors used for function-specific behavior |
| --- | --- | --- |
| `DATE` | `date_bounds_fractional_floor_and_zero_offsets_follow_the_profile`; `component_and_month_shift_edges_use_checked_profile_bounds` | rollover, leap/1900, fractional truncation, profile bounds, invalid month |
| `DATEDIF` | `datedif_formats_trim_case_and_apply_signed_month_remainders`; selected `date_param.datedif` source proof | Y/M/D, MD/YM/YD, reversed interval, invalid format |
| `DATEVALUE` | fixed parser and numeric fallback tests; borrowed-text and parser-limit tests | ISO, datetime, en_US/month-name, pivot, numeric fallback, fractional fallback, invalid calendar day |
| `DAY` | date-bound and component-bound tests; family arity/error test | flooring, text, out-of-domain |
| `DAYS` | `days_converts_logicals_and_single_cell_references`; family formula-error test | retained fractions, reversed interval, text/number conversion |
| `DAYS360` | optional-slot, date-bound, and family formula-error tests | US February, European sign/31st, reversed US |
| `EASTERSUNDAY` | timestamp and family arity/error tests | explicit years, timestamp-relative year, invalid year, missing clock |
| `EDATE` | date-bound, component-bound, and family formula-error tests | month-end clamping, negative/fractional month |
| `EOMONTH` | date-bound, component-bound, and family formula-error tests | leap/previous month and invalid domain |
| `HOUR` | finite time-domain and family arity/error tests | midday, negative time, final second |
| `ISOWEEKNUM` | ISO/week boundary and component-bound tests; selected `date_param.isoweeknum` source proof | first-Thursday and adjacent week-year boundaries |
| `MINUTE` | finite time-domain, leap-second, and family formula-error tests | half-second, near-half, negative, end-of-day wrap |
| `MONTH` | date-bound, component-bound, and family arity/error tests | flooring, named text, invalid text |
| `NETWORKDAYS` | projected complete-sequence, invalid-holiday, retained-error, and sequence-limit tests | default/reversed/custom week, duplicate holiday, all-off, formula error, list refusal |
| `NOW` | explicit timestamp and family arity tests | explicit and missing timestamp |
| `SECOND` | finite time-domain, leap-second, and family formula-error tests | half-second, boundary, negative, final-day fraction |
| `TIME` | finite time-domain and family formula-error tests | direct, negative, multi-day, overflow |
| `TIMEVALUE` | fixed parser, numeric fallback, borrowed-text, and leap-second tests | clock/datetime, numeric/fraction fallback, 24:00/date-only/leap-second refusal |
| `TODAY` | explicit timestamp and family arity tests | explicit and missing timestamp |
| `WEEKDAY` | optional-slot, component, and matrix tests | all accepted types, omitted type, invalid type |
| `WEEKNUM` | optional-slot and component tests | accepted modes, aliases, ISO modes, omitted/non-integer mode |
| `WORKDAY` | sequence/cancellation/limit tests; checked WORKDAY offset source proof | forward/backward stepping, exact zero offset, holiday/custom/all-off |
| `YEAR` | date-bound, component-bound, and family arity/error tests | pivot, datetime, profile boundary, invalid text |
| `YEARFRAC` | optional-slot, basis, floor, and 30/31st tests | all bases, reversed nonnegative result, invalid basis |

The WORKDAY mapping was corrected after review: the projected complete-sequence
semantic test is NETWORKDAYS-specific and is not used as WORKDAY evidence.
WORKDAY sequence behavior is linked to
`workday_consumes_holiday_and_workweek_sequences_before_zero_and_all_off_results`
and `workday_cancellation_follows_complete_sequences_and_work_is_bounded`.
Its toward-zero finite offset conversion is linked to the selected
`workday.offset_conversion` source proof in
`gates/source-review-workday-conversion.json` (SHA-256
`ebe94398855df389c1c64dbd8ca33a1a64d38e74e978734a72c7eecd14317bee`).
DAYS360's European reversed-sign requirement uses
`days360.european_swaps_with_sign`; NETWORKDAYS's custom week, duplicate holiday,
and all-off requirements use their corresponding named rows rather than the
default-week row.

The exact `NOW`/`TODAY` ambient-state requirement is bound by the supplementary
`clock.no_ambient_state` source proof in
`gates/source-review-clock.json` (SHA-256
`7dea33a3fa3de83510eefe43d222f68706731a3c284aaa5e47315b178ee7dc8c`), whose
frozen source is `evaluation/date_time.rs` at SHA-256
`29f1495a55aac7effee9f6509deef38e5d5f747ba2b87c65a188d8759c1265e3`. The
shared `profile.no_ambient_state` proof remains bound to the cross-cutting
deterministic-profile requirement; the two proofs have distinct requirement
strings and are not substituted for one another.

A binding is an exact evidence link, not a claim that the test suite forms an
unbounded Cartesian product of every conversion and boundary. The shared
DateParam cases that do not have separate function-level vectors are closed by
the selected, function-specific `date_param.datedif` and
`date_param.isoweeknum` source proofs. Their proof requirements are split in
the receipts so DATEDIF evidence cannot accidentally certify ISOWEEKNUM, or
vice versa. Valid ISO-date oracle rows are not used to certify ISOWEEKNUM error
handling; valid named-text rows are not used to certify MONTH error handling.


## Cross-cutting evidence

The cross-cutting bindings cover arity and error precedence, borrowed-text
parsing and bounded scratch, scalar/reference shape gates, matrix publication,
complete sequence streaming, per-cell work/read/cancellation checks,
source-version and final-cancellation fences, checked calendar arithmetic,
bounded reservations and drop order, typed failure precedence, demand-cache
identity and volatile timestamp eligibility, position-sensitive MUNIT handling,
read-only authority, independent oracle/native corroboration, and gate/performance
custody. Resource and authority requirements are linked both to their focused
limits tests and to the corresponding hash-bound source proof. The source-proof
receipts are multi-proof receipts; the verifier selects the requested proof IDs
and requires their exact mapped requirement union, while still validating every
receipt proof's schema and source identity.

The final command

```text
python3 docs/report/spec-gap-validation-evidence/ods-formula-date-time/verify.py
```

passes contract, scope, source closure, coverage, gate, review, oracle, native,
and performance custody checks with `verified: true`; the coverage check reports
223 function bindings and 54 cross-cutting bindings. The negative-case suite
and four source-proof selection regressions also pass. No production, test,
contract, freeze, gate, oracle, native, or performance input was changed by
this audit pass.

Earlier gate archives remain diagnostic history only: the provenance-typo bundle
was invalidated by its corpus archive hash, and the harness-checksum bundle was
superseded after a harness-only capture failure. The current root bundle listed
above is the only custody referenced for final review.
