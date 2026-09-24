# ODS data-style model corrections

Omitted decimal replacement now resolves through explicit/inherited decimal places; only a present empty replacement implies zero minimum decimal places. Typed text follows the XML 1.0 character domain, accepting tab, LF and CR while refusing illegal characters. The model also exposes a distinct read-only styles.xml automatic owner and preflights aggregate metadata patches before cloning or mutation.

Independent model-only review approved the source recorded in source.json. Root validated the exact isolated model with 271 library tests, strict Clippy for library/tests, and warning-denied rustdoc. Commands, exit codes, losslessly compressed logs and the resolved Cargo.lock are retained. TMPDIR=/var/tmp and a dedicated target were used; inherited Rust warning suppressions were removed.

The tested and committed file omits only the in-progress source-module export from the shared model. Source scanner and package-owner integration are not approved by this batch and remain separate work. No performance result is claimed.
