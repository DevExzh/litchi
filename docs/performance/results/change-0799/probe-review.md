# 0799 source-only probe review

This packet is a source-only extension of the 0797 direct checked-attribute
probe. The release binary remains an independent project over `quick-xml`
0.41.0. The production baseline and candidate helper modules are intentionally
omitted from this handoff; the root build owns their later materialization and
must bind each archive by hash before compiling or measuring.

The probe keeps the 0797 construction and consumption owners, checksum
protocol, timing boundaries, semantic oracle, and fail-fast iterator behavior.
Only the packet identity and case catalog changed. The catalog has 39 cases:
the 33 retained 0797 inputs plus valid duplicates after the second attribute,
long quoted and unterminated duplicates after the second attribute, and the
three syntax tails after the second attribute. `fixtures.json` is an
independently frozen literal catalog; `fixture_check.py` requires exact
identifier order, source bytes, lengths, and SHA-256 values.

Clone oracle transitions now include advances 0, 1, 2, 3, 4, 5, 32, and 33.
This checks the new second-attribute boundary while retaining the existing
32-name transition checks. The probe source itself has not been built, tested,
captured, or used to claim production adoption. Root-owned gates must run
rustfmt, the locked offline build/check/Clippy sequence, the self-check, and
the 39-case fixture audit after supplying the two helper archives.
