# Planning mechanism and whole-child counters

This review covers the initial candidate binary. The rebuilt candidate subsequently failed native admission and the production change was reverted; see `final-native-comparison.json`.

All four planning profiles pass the frozen 1% instruction reduction gate:
medium improves3.274% and5.041%; dense-sparse improves5.107% and5.047%.
Each selected dump has exactly one measured-runner caller, separate from
three retained lifecycle dumps and the zero-instruction termination dump.
Source/binary/output identities and self/direct equations are checked.

The medium MCE processor edge falls from6,466,820 to122,148 instructions in
both repeats. Total planning improves less in repeat1 because its separate
XML validation edge rises by2,252,081 instructions. Repeat2 validation is
nearly unchanged. This variation is retained; the evidence does not establish
its cause. Nested MCE costs are not additive to worksheet parsing or planning.

Hardware counters cover the entire child, including corpus construction,
setup, validation and reopen oracles. They have mixed directions between
repeats: first-repeat cycles rise2.727%/4.591% and instructions rise0.938%/0.802%
for medium/dense-sparse; second-repeat cycles fall1.719%/3.996% and instructions
fall2.449%/4.331%. These are diagnostics, not an isolated planning hardware
speedup or a reason to substitute counter ratios for the native timing gates.
See the exact coverage, multiplexing, raw metrics and limitations in
`hardware-analysis.json` and the profile comparison.
