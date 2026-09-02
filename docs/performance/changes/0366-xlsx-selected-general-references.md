# Change 0366: XLSX selected general-reference decoding

- `performance_claim`: none

## Summary

The selected worksheet single/range scanner now decodes bounded XML general
references in scalar payloads through the canonical decode helper. Predefined
`amp`, `lt`, `gt`, `quot`, and `apos` references are eligible in formula, value,
and inline payloads. Decimal and hexadecimal numeric references are eligible in
formula and value payloads only when the spelling is ASCII and the full token is
at most 12 bytes.

Numeric references in inline payloads, overlong or non-ASCII numeric spellings,
and numeric scalar values outside the XML 1.0 `Char` production return
`NotEligible` and use the verified eager fallback. Malformed, custom, and
out-of-range references remain MCE/typed errors.

The scanner still drains XML/MCE/x14ac and the OPC reader to verified EOF before
publishing a result or invoking a callback. The eligible cold `cell`, `cells`,
and `visit_cells` paths retain no `Store`, `PartData`, or semantic caches. There
is no API change or public accepted-input change. The pre-existing
eager/shared-string XML-legality residual is explicitly out of scope.

## Validation

Focused validation passed `9/9`; full `litchi-xlsx` library validation passed
`892/892`; and scoped Clippy passed with `-D warnings`, with only the unrelated
`clippy::useless-asref` issue allowed. This is correctness and boundary evidence
only; no latency, RSS, fixed-memory, or OOM claim is made.
