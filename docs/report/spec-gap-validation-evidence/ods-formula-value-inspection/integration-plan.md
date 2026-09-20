# Value-inspection integration groundwork

Status: implementation integrated; independent source review and final isolated
validation remain pending. The settled semantics are in `contract.md`. This
document records the integration decisions, not a validation or performance
acceptance claim.

This batch covers 16 pure value-inspection/conversion names from ODF 1.4
section 6.13, listed in baseline.json. Five counting functions already exist.
The remaining reference/host-metadata functions require separate explicit
integration; this batch does not claim completion of section 6.13.

Root inspected the local normative sections for ISBLANK, NUMBERVALUE, TYPE and
VALUE. ISBLANK distinguishes Empty from empty Text and does not propagate
formula errors. TYPE accepts Any and reports Array separately, so it cannot
blindly use ordinary scalar projection. NUMBERVALUE has an ordered separator,
whitespace and percent transform followed by xsd:float syntax. VALUE requires
ISO dates, times, datetimes and mixed fractions as well as numeric input;
English locale support also entails the required currency/grouping/date forms.
These semantics must be implemented rather than reduced to a decimal parser.

The scalar WorkingValue representation remains unchanged. A private inspection
Input retains Empty and omitted arguments before ordinary scalar coercion. The
value mapper preserves reference descriptors and reads one selected cell per
output coordinate. TYPE scans complete references with fixed state and reports
64 for arrays; N uses scalar reference intersection and the first inline array
element. Ordinary inspection functions lift over matrix coordinates. Explicit
reference lists are refused before reads, and known invalid NUMBERVALUE
separators produce shape-preserving errors without source reads.

TYPE and N evaluate complete Any arguments in the enclosing calculation mode,
including computed arrays under projected lazy branches. This uses the same
context boundary as TRANSPOSE while preserving the current output position.
N explicitly supplies scalar intersection when needed, including
`TYPE(MUNIT(N(range)))`; MUNIT's existing first-element matrix parameter behavior
is unchanged. Direct 3-D matrix
arguments select the current sheet plane, with metadata-only shape planning.

TEXT and VALUE share private Gregorian calendar helpers. VALUE uses the fixed
en_US profile, epoch 1899-12-30 without a fictitious 1900 leap day, and the
1930–2029 two-digit-year window. No ambient locale or clock is consulted and no
public options were added. Long grouped numbers use budgeted scratch storage
and exact decimal-to-binary conversion; references and borrowed text are never
collected into an input cell vector.
