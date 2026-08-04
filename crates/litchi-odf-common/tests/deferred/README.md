# Deferred common ODF coverage

The `flat/` and `cross_family/` sources preserve the former umbrella's
cross-family regressions. They are outside direct Cargo integration-test
discovery because the old wrappers depend on the monolithic API; each flat or
family-specific path will be wired by its owning crate after shared scanners
and package primitives are isolated.
