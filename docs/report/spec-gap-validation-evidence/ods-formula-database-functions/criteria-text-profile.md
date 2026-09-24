# Criterion text profile correction

The shared private matcher now follows the selected ODF 1.4 criterion profile:

- a bare numeric-looking Text criterion such as `"3"` remains Text and
  matches a Text candidate containing `3`;
- an operator-prefixed criterion such as `"=3"` uses numeric comparator
  semantics and matches a Number candidate containing `3`;
- the rule is shared by the twelve database functions and the six conditional
  aggregates.

The focused regression is
[`ods_formula_criterion_text_profile.rs`](../../../../crates/litchi-ods/tests/ods_formula_criterion_text_profile.rs).
It exercises both `DSUM` and `SUMIF` with Text and Number candidates. The
validation command is:

```text
cargo test -p litchi-ods --test ods_formula_criterion_text_profile
```

The earlier database candidate receipts remain historical records of the
pre-correction profile. They are intentionally not rewritten; this note and
the implementation profile describe the current behavior.
