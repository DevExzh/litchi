# 0362ul68b — XLDM same-size rewrite preflight

Status: final source-review return for 0362ul68. Cargo, native
qualification and rustdoc remain deferred to the coordinator's build slot.

`Xldm140Closure::rename_table_name` now computes the escaped metadata
replacement length and the unchanged outer output length before constructing
the replacement `Vec`. A source whose retained `max_output_bytes` is already
below that output is refused at the identity boundary; the codec remains a
second admission check before any output copy. Exact name no-ops still return
a borrowed patch before this preflight and therefore do not allocate or copy
the source.

The restrictive-cap identity regression covers both paths: a changed
same-size rename reports the retained output refusal before replacement
materialization, while an exact no-op remains borrowed. Formatting, diff,
repository-size, no-binary, and staged guards pass. No Cargo command or
runtime qualification is claimed.
