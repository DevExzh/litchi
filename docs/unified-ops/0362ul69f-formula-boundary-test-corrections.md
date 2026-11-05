# 0362ul69f — Formula raw-policy boundary-test corrections

Date: 2026-10-25. Parent: 0362ul69e. Status: source repair complete; the
coordinator still owns the next guarded validation attempt.

The hydrated all-target Formula run compiled and executed 348 tests, then
reported three focused failures. Two were test-boundary mistakes: the root
LINE plus two ruled wrappers contributes five encoded levels, so the
double-rule acceptance boundary is max depth five and the refusal reports
actual five. The runtime `write_ruled` and preflight `depth_ruled` both have
the same two enters per wrapper; this correction changes only the regression
expectation and does not claim a production depth fix.

The third failure was the cumulative split Text/CDATA boundary assertion. Its
diagnostic showed `StringBytes` actual seven from the `m:oMath` root-name
encoding (the lexical `<m:oMath>` spelling is nine bytes): the test's two-byte
string ceiling could never reach text admission. The witness now allows the
seven-byte root-name charge with a max-nine ceiling and splits a ten-byte text
payload, so the exact refusal is cumulative `StringBytes` actual ten. The
parser implementation remains unchanged in this return.

The external f5 receipt is:

    (deleted-coord-artifact)/02-formula-all-targets-tests-attempt-formula-all-targets-fixtures-fixed.log
    28403 bytes, SHA-256 4e3f79eb537fb124d8cdbed90aff6c7fe61b17c9289eb47a3857f2735ae44700

The command exited 101 after 345 passing tests, three failures, and nine
passing LaTeX integration tests; no resource or provenance guard failed. Exact
4100 and original tool locks remain byte-checked. No lock generation, native,
rustdoc, broad ops, or performance command is part of this return.
