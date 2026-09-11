# InkAction structural profile

The shared DrawingML owner now exposes `ink::actions::read_profile`, `write_profile`, and `write_profile_to`. A profile provides ordered typed action groups, actions, properties, and data records with retained source ranges. Writers replay the validated source exactly. Existing `actions::read` remains the compatible opaque read path.

The profile checks the action-container grammar from vendored [MS-ODRAWXML] section 2.21 and schema appendix 5.19: admitted expanded names, required attributes, ordering, cardinality, units, decimal whitespace handling, and duplicate expanded attributes. Unknown direct container children and attributes are refused. Definitions, transforms, traces, and trace views retain opaque payload subtrees.

This is a prerequisite for host integration. It does not provide action authoring mutations, host package attachment, runtime action execution, full InkML payload XSD validation, or semantic validation of reserved property combinations. References remain bounded lexical anyURI values. Native Office acceptance and performance improvements have not been established.

Source, attribute, node, depth, action, and group caps bound retained work. Namespace comparison can perform repeated decoding across attribute pairs; the attribute limits bound this cost, but no linear-time or benchmark claim is made.

`verification.json` binds the production and test sources to retained compressed logs. The broad DrawingML tests passed, including 166 library tests. After a test-only Clippy correction, the 16 public profile tests, all-target Clippy with warnings denied, and rustdoc with warnings denied passed. The original lint failure is retained alongside the successful final logs.
