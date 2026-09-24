# OPC source-preserving relationship removal

Removing an SVG attachment can require deleting a relationship from a native,
formatted `.rels` member. The previous topology publisher refused this operation.
The new removal-only path uses parser-derived element spans and retains the XML
declaration, root metadata, comments, processing instructions, whitespace, trailing
content, and remaining relationships. Removing the last relationship retains the
empty member. Candidate validation completes before publication.

The parser validates root attribute QNames, bound prefixes, expanded-name
uniqueness, and reserved namespace bindings. It releases its first range map
before reparsing the candidate. Existing cancellation and memory admission remain
in force.

Noncanonical members support append-only or removal-only operations. Mixed edits
and relationship replacement require canonical source. Encoded namespace URI
declarations are safely refused by this lexical profile; that refusal does not
mean such declarations are invalid XML.

`opc-review.json` binds independent scoped approval and the root's nine passing
focused tests to exact source and log hashes. The author separately reported all
387 OPC library tests, check, and strict clippy passing. This is a completed OPC
prerequisite, not approval of the still-in-progress PPTX SVG lifecycle or a
performance claim.
