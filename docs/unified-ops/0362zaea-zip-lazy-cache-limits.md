# 0362zaea — canonical lazy ZIP cache limits

- Date: 2026-10-24.
- Parent lane: `0362zae` (generic ZIP transport limits).
- Base: `e35cede32a`.
- Scope: the public lazy ZIP reader/cache policy boundary and two consumers
  missed by the first source migration.

## Canonical shape

`LazyArchiveReader::new(data, archive_limits, cache_limits)` is the one public
lazy-reader constructor. The former `new_with_cache_limits` spelling is gone;
all 32 former default-cache calls pass `LazyArchiveCacheLimits::default()`, and
all eight custom-cache calls pass their existing profile as the third argument.

`LazyArchiveCacheLimits::new` now requires all four resource values:

```text
new(max_bytes, max_entries, max_active_flights, max_flight_key_bytes)
```

The three compatibility aliases `new_with_flight_limits`,
`with_flight_limits`, and `max_active_key_bytes` are removed. The primary
active-flight construction spellings are also canonicalized: the consuming
setter is `active_flight_limits(max_active_flights, max_flight_key_bytes)`;
the former `new_with_active_flight_limits` and `with_active_flight_limits`
names are removed. The four policy getters are
`max_bytes_value()`, `max_entries_value()`, `max_active_flights_value()`, and
`max_flight_key_bytes_value()`.

The old two-value cache constructor's default flight values are preserved at
callers with `DEFAULT_MAX_ACTIVE_FLIGHTS` and
`DEFAULT_MAX_FLIGHT_KEY_BYTES`. Custom four-value profiles retain their exact
values. Cache-policy configuration remains distinct from the raw ZIP reader's
`ArchiveLimits`; no archive entry, payload, validation, error, or cache
algorithm changed.

## Consumer corrections

The ODF `SourceBackedPackage::preservation_index` helper now passes its
retained `self.limits.archive().into_zip_limits()` profile instead of a fresh
ZIP default. A regression test tightens the retained profile after open and
verifies that preservation rejects the metadata budget, then restores the
profile and verifies success. This binds the source-backed owner to its
format-selected transport profile without changing the normal open path.

The ZIP `parse_zip` fuzz target now passes its existing `fuzz_limits` to the
replay preservation-index call. No fixture or upstream checkout is referenced.

## Gate coverage

`tools/api_glossary.json` registers exact bans for `new_with_cache_limits`,
both removed cache constructor aliases, both removed active-flight setter
spellings, and `max_active_key_bytes`, in addition to the generic ZIP
`_with_limits`/`with_limits` bans. `docs/API_GLOSSARY.md` records the required
four-value cache profile and the three-argument reader constructor.

Validation after the follow-up source correction used the same coordinator-
approved environment: the build shim was first in `PATH`, `LITCHI_CAPPED` and
`LITCHI_BLESS` were unset, `LITCHI_MEM_MAX=20G`,
`CARGO_BUILD_JOBS=1`, `RUST_TEST_THREADS=1`, `CARGO_INCREMENTAL=0`, and the
external target was
`(deleted-coord-artifact)`. The
coordinator-pinned ignored `Cargo.lock` remained SHA-256
`52ed0cacac15b5e440349bfc0f4aca266bcf50d59d4321624e596bb410783c6`.

The corrected focused package command was:

```text
cargo test --offline --locked -p soapberry-zip --all-targets
```

It passed 506 unit tests and 125 integration tests (631 passed total), with
one bounded multi-gigabyte test ignored; all four examples compiled and had
zero tests. The focused ODF/OPC consumer check was:

```text
cargo check --offline --locked -p litchi-odf-common -p litchi-opc --all-targets
```

It completed successfully (`Finished dev profile`, exit 0). The targeted ODF
regression was then run with:

```text
cargo test --offline --locked -p litchi-odf-common --lib source_backed_preservation_index_uses_retained_zip_limits
```

It passed 1 test, with 372 filtered out. The first consumer-check attempt
failed only because the new assertion used `unwrap_err()` on a successful type
without `Debug`; the assertion was changed to an explicit match and the
corrected command passed. The canonical setter coverage then passed with:

```text
cargo test --offline --locked -p soapberry-zip --lib lazy_cache_active_flight_limits_reconfigure_and_validate
```

The setter test passed 1 test, with 506 filtered out. No command-log artifact
was retained, so there are no log hashes to claim. Rust formatting was checked
on the 307 changed Rust files only; no rustdoc, dependency-graph comparison,
or performance measurement is claimed.
