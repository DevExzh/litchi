# 0360ab — Return metadata scope composition for capability repair

Candidate `00e6add236` repairs the earlier blanket refusal to compose a
scoped no-op with a changed patch. Independent review still reproduces two
valid operations that it refuses. It is returned for repair, not accepted
as an exception or integrated into the W4/File/XLSB assembly.

The coordinator assembled `528e04e6d5`, `9d68463b1c` and `00e6add236` over
`933ce43de3`. The scope implementation, XLSX adapter and existing tests match
the worker candidate before adding the
[two regression assertions](results/0360ab/regression.patch). The assembly
also retains W4b's unchanged deferred snapshot helpers; the receipt verifies
that these are the only difference from the worker's common source file.
The [receipt](results/0360ab/review.json) records source hashes, commands and
external log sizes/hashes. The corrected test filter runs two tests and
both fail. An initial incorrect filter selected zero tests; that attempt
is retained and is not counted as validation.

1. On an empty OPC package, compose a no-op with creation of core
   properties, publish it, then apply its inverse. The inverse refuses the
   application-properties scope of its own publication. Every absent
   family initially records the same free relationship ID. Creation uses
   that ID, but the after scope still requires it to be free for the
   other absent families.
2. Compose a changed core-property patch with its inverse, then compose
   that cancellation with an independent application-properties change.
   The second composition refuses because the application family is
   absent from a supposedly complete scope. Cancellation constructed a
   partial scope, which the composition code treats as complete.

The repair must preserve compatible composition and inverse replay while
authenticating recorded before and after states. It needs distinct handling
for complete and partial scopes, actual free-ID transitions, and tests for
creation, removal, independently captured no-ops and disjoint changes.
Signed exact no-op behavior and recipient policy must remain intact.
The W4 worker owns this repair; DOCX/PPTX metadata consumers must integrate
against the repaired shared contract. Performance qualification remains
required after correctness is restored.

Registering this receipt also exposed that the benchmark evidence scanner
silently skipped record suffixes longer than one letter. It now recognizes
these follow-up records, and a fixture verifies that unregistered evidence
fails coverage. Eleven existing validation/native-only evidence registrations
are copied from the validated assembly; none claims a leaf-versus-facade
measurement. Matrix regeneration/checking and all nine matrix tests pass.
