# Final MUNIT boundary review

The source-only `ods_reviewer` review covers `value.rs` SHA-256
`d03de9971049cee8ae1d31de442224931ee68dd646994958aae3b30147a34756`
and `value/matrix.rs` SHA-256
`5b88197b08b759cd811c65e665c78118280f941a194c0cc921dededb0a7921d2`.
These match candidate 06 and the implementation committed in `ab199671f`.
No blocker was found within the MUNIT parameter/context/ownership scope.

- Argument evaluation preserves the enclosing mode and clears inherited output
  projection. Scalar outer evaluation retains implicit intersection; Matrix
  mode and projected lazy branches consume the first row-major parameter cell,
  matching Part 4 §3.3 rule 2.2.1.
- Conversion preserves formula errors, applies documented Logical/Empty rules,
  and reads only the first cell of a single-area reference through the shared
  checked and work-charged provider path.
- Moving the selected array value releases the unused tail and metadata.
  Probe restoration reinstates suspended planner, VM and argument context state;
  temporary buffers drop before their corresponding reservations.
- Shape discovery uses the same isolated probe and conversion. A failed probe's
  1×1 hint does not hide the selected evaluation's typed or formula failure.

Multi-area/3D reference projection remains an explicit unsupported profile
boundary. This review does not establish broader OpenFormula completeness,
numerical stability for arbitrary matrices, or performance acceptance. Runtime
validation is separately retained in candidate-06, with 34 focused tests and
all five ODS gates passing.
