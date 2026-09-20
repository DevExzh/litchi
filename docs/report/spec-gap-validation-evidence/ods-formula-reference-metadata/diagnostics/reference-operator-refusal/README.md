# Descriptor refusal integration regression

The private probe outcome initially merged unknown shape and an internal
reference-handler array refusal. Full integration testing caught that ordinary
reference operators then accepted array-valued IF/error-handler operands under
projection. The correction keeps a distinct private refusal outcome: metadata
shape discovery may defer it, while ordinary reference operators return their
existing typed Unsupported(ReferenceOperator). No timing capture used this
intermediate source. The complete gate receipts are retained after termination.
