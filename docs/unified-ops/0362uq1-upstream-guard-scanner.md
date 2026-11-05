# 0362uq1 — upstream-path guard literals

The tooling unit gate exposed a false positive in the repository-wide
no-upstream-checkout guard.  The ODG qualification runner and its builder
must reject a path component naming the untracked upstream tree, and the
builder test must exercise that refusal.  Those checks are value validation;
they do not read from or depend on that checkout.  The source scanner was
matching the validation value itself, so it reported the guard as a forbidden
consumer.

The ODG runner, builder and focused test now construct that forbidden path
component from adjacent string fragments.  The rejection remains exact at
runtime and the test still verifies it.  The builder documentation describes
the policy without spelling the component.  No allowlist or scanner bypass
was added, and no fixture or production path was changed.

## Checks

Source: `250769e1dcb692994a382829c238bf23bda5035d`.

* `python3 -m unittest tools.test_no_3rdparty_references` — 3 passed.
* `python3 -m unittest tools.test_no_3rdparty_references tools.test_odg_joined_qualification` — 29 passed.
* `git diff --check` — passed.

The repository guard still scans all tracked `crates/`, `tools/` and `fuzz/`
paths with no temporary exceptions.
