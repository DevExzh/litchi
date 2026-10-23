# 0362ul — OPC external consumer migration

The required OPC constructor ingress now takes an explicit `SourceLimits`
profile, and source-backed ingress additionally takes `SourceCacheLimits` and
`SourceReadPolicy`. The external OPC tests and examples are migrated together
on top of `7858571ad0`; every use retains the existing default or selected
policy and no assertion or fixture behavior changes.

The migration also supplies the missing limit and policy names in the focused
source tests, updates the one lazy-owned-package constructor that still used
the removed one-argument form, and removes denied redundant qualifications
where the test already imports the canonical policy type. No public alias or
compatibility constructor was added.

This source-only commit was checked with rustfmt, `git diff --check`, and the
repository artifact guards. Cargo validation is deferred to the assembled
integration check because the OPC stream admission repair is still being
merged; the v12 log identified the import and constructor diagnostics this
commit addresses.
