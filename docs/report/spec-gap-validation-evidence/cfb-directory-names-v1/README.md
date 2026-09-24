# CFB directory-name validation

`directory_names_equal` and `validate_directory_name` expose the existing
CFB directory sorting and authoring validation to format owners. Name
validation stops after 32 UTF-16 code units when the 31-unit payload limit is
exceeded, before retaining the name or scanning the rest for forbidden
characters. Invalid names never compare equal. Oversized inputs now report
an at-least length diagnostic before any embedded forbidden-character error.

The shared comparator now leaves supplementary scalars unchanged, as required
by [MS-CFB] section 2.6.4's UTF-16 surrogate rule, and handles the 27 BMP Greek
simple-uppercase mappings whose full uppercase form expands. Other expanding
full-uppercase mappings remain unchanged. This corrects both public helpers
and existing directory reader/writer comparisons.

Rust 1.95.0 validation passed 321 unit/integration tests and fourteen doctests
(one doctest ignored), strict all-target/all-feature Clippy, formatting, and
whitespace checks. Independent review checked the local Markdown specification,
Unicode edge cases, bounded validation, and existing call sites. The receipt
binds both source files and compressed/raw logs. This is correctness evidence;
no measured latency or memory improvement is claimed.
