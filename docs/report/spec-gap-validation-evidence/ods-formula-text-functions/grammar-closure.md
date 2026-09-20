# Root follow-up on formatter grammar findings

The independent reviewer identified three blockers in the first frozen
candidate; its original report and sources remain under
`diagnostics/pre-grammar-fix/`. Root completed this follow-up after the
reviewer and oracle agents reached their usage limit. This is a root review,
not a claim of a second independent-agent approval.

Current formatter SHA-256:
`129b4323c1f577f048e6bc3431653d4c2d979cc04f2fa60305d3bcf84d923c06`.

The source diff closes the reported cases:

- Numeric parsing rejects tabs and allows only one mixed-fraction separating
  space; date parsing admits its explicit unquoted token/separator set.
- Section parsing counts unquoted percent markers across prefixes, core and
  suffix, refusing duplicates only for numeric sections. Quoted, escaped and
  text-section percent literals retain their behavior.
- Integer, fractional and rational rendering emit admitted percent markers
  in source order. Suppressed zero fields scan for percent literals without
  restoring omitted numeric slots, slash or mixed separator.
- Exponents retain contiguous E/e, optional sign and placeholder slots;
  percent cannot split that token. The contract now states this explicitly.

Every added scan checkpoints before inspecting a byte. Percent output uses
the existing Sink admission, work and typed-failure paths. The diff adds no
allocation or retained vector. Both counting and writing passes share the
same rendering logic; formula-error publication retains execution fences.

Validation includes 244 exact oracle observations across 26 functions and
six direct scalar regressions in the existing formatter test. Those pass.
The five newly added oracle cases specifically cover optional fractional
zero, improper zero, mixed zero, required whole zero and required numerator
zero with a percent in the suppressed denominator. Final isolated gates and
performance evidence are separate completion requirements.

The corrected isolated freeze now passes all seven gates. The independent
gate verifier confirms 1,592 passed, zero failed/ignored, exact focused
13/12/2/26 counts and unchanged source manifests. Performance capture and
its separate verification remain pending.
