# 0362dvhb — PPTX subordinate-reader limit direction

This return packet follows `0362dvha` and corrects the PPTX shape facade's
subordinate-reader policy. `Common::transform` and `Common::text_body` both
admit their source fragments under the retained scene `input_bytes` budget;
the scene's `output_bytes` budget remains the owner-level processed-output
bound enforced while the scene is indexed.

The mapping is covered by asymmetric input/output limit regressions in the
shape model so a future change cannot silently route a source reader through
the output budget. No compatibility API or alias was added.

Validation is coordinator-owned until a Cargo slot is available. This lane
contains source and test changes only; no native workload or ABBA run was
performed.
