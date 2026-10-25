# 0362ul94rb — properties source-boundary coverage

Review of `0362ul94ra` found two remaining provenance gaps. The extended-
properties scanner did not inspect start-tag attributes and only checked text
after accumulating a recognized scalar, so controls in namespace attributes,
unknown subtrees, CDATA, or entity references could survive as retained source
and later reach a valid splice. The core reader similarly skipped namespace
attribute values and the attributes/text of unknown or duplicate subtrees.

The readers now validate every decoded source attribute value, including
`xmlns` declarations, and every decoded text, CDATA, and general reference
before the event is admitted or skipped. Application validation is fused with
event decoding, avoiding a second scan of a completed scalar. Adversarial
tests cover namespace attributes, unknown subtrees, entity references, CDATA,
and the existing recognized-property case; source failures remain nested
package errors while authored controls remain invalid values.

Core authored-target validation now checks its fields directly. It no longer
constructs a detached `Core` and runs a duplicate preflight before the actual
merge and output-length admission.

This is a source-only repair. No Cargo, build, rustdoc, runtime, native, or
performance command was run; the coordinator must rerun the affected gates.
