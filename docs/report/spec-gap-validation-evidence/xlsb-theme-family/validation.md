# Validation scope

The retained gate runner binds its results to the production and test source
hashes before and after execution. The downstream caller has a separate owned
build target and records the same production source manifest, its own public-API
source, and generated XML hashes. Neither run substitutes for independent review.

## Behavioral coverage

The shared Theme tests exercise direct owner selection, both admitted extension
identifiers, token whitespace, inherited namespace closure, prefix collisions,
unknown content, absent-owner add/remove, exact no-op and inverse behavior,
malformed recognized content, XML lexical constraints, reserved namespace rules,
MCE ownership refusal, and finite input/output limits. Native complete Theme
bytes are retained as a provenance-documented fixture.

XLSB tests exercise native and Strict profiles, combined base/family edits,
stale-source refusal, signature policy, optional Theme lifecycle, writer
creation, save/reopen, and source-backed projection without reading unrelated
media. The external caller independently uses the public API for native
read/update/remove/add and newly authored workbooks, with exact inverse checks
across actual save/reopen.

Offline validation uses the vendored ECMA complete Theme schema and the adapted
Microsoft family schema. The adaptation is documented in `validate_schema.py`;
no network or native Office application is involved. Native fixtures and the
four caller-generated outputs have separate schema reports.

## Interpretation

The gate receipt is authoritative for final test counts and command results.
Existing ignored tests are reported rather than counted as passing. Performance
results are instrumented, scoped workloads; they do not establish general
workbook throughput, a before/after speedup, or native application acceptance.
The broader audit remains open.
