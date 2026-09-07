# Standalone harness strict-lint repair

The attempted all-target/all-feature strict harness gate found 50 existing style
errors across library modules and binary targets. Root applied mechanical repairs after the draft agent was
interrupted, and removed its incomplete proposed patch artifact. Full before/
after source copies and the failed strict command remain in this bundle.

Replacements use equivalent short-circuit let chains, constant/nonzero-divisor
is_multiple_of, direct same-type values/errors, contains, is_none_or, and complete
struct initializers. Case-fold lookup visits the same fixed-size outcome rows
with iter_mut rather than repeated indexing. The single CFB structural range
remains a Vec containing one Range; it is not expanded into all offset values.
The vendor payload checked size is still evaluated eagerly before the shape
branch, preserving the prior then_some error ordering. Rustfmt also normalizes
existing formatting in the touched harness files. No lint is suppressed.

Both compared builds use the exact same final harness source manifest. These
style repairs are not attributed as PPTX optimization benefits. Existing full
harness tests and strict checks validate the final source epoch.

After the initial 29 library warnings, Cargo exposed two XLS/XLSB binary warnings,
seven DOCX selected-paragraph warnings, seven XLSX ABBA warnings, and five
allocator-test comparisons. All were repaired without suppressions. Windows
staged-file cleanup still retries after clearing readonly only when the first
remove fails; the non-Windows path still makes one best-effort remove. The XLSX
ABBA test module moves unchanged to the file end. No Windows runtime claim is
made from the Linux checks. `harness-lint-repair-r5` passes all targets/features.
