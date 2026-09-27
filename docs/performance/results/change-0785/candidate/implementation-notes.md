# 0785 archived known-URI namespace candidate

This is an archive-only candidate against base commit
`4d89ccf28d6ef6c61d7d2f35b716cd1fcfab7985`. The live `litchi-pptx` source was
not edited. `candidate/model.patch` is a git-compatible patch relative to the
base, `candidate/before/` records the complete baseline replacement file, and
`candidate/files/` and `candidate/applied/` record the complete candidate
replacement file. The candidate and applied archives are byte-identical; the
applied name records the coordinator's future application slot and does not
claim that production currently contains the change.

The change keeps `resolved`'s public and crate-private signature, all existing
constants, error variants, error text, and resolver lifetime contract. Bound
values are dispatched by their byte length to the six existing exact constants
(`P`, `PS`, `A`, `AS`, `R`, and `RS`) and compared byte-for-byte. An exact match
returns the corresponding static string. Every other bound value follows the
existing `std::str::from_utf8(value).map_err(xml_error)` path, so valid vendor
namespaces remain borrowed from the resolver and invalid UTF-8 still returns
the same `Error::Xml`. `Unbound` and `Unknown` arms are unchanged. The helper
has no allocation, cache, unsafe code, dependency, API, or XML-validation
policy change.

Focused tests are inline under `#[cfg(test)]` in the archived replacement.
They use an independent copy of the pre-change resolver as an oracle and cover
all six exact constants, every replacement byte at every position (skipping
only the original byte), pointer-preserving fallback for every valid
substitution, exact invalid-UTF-8 parity, prefix/suffix/vendor namespaces,
Unicode namespaces sharing each known byte length, empty bound namespaces,
`Unbound`, and both ordinary and non-UTF-8 `Unknown` prefix messages. The
existing notes scanner and end-to-end suites remain the broader differential
and publication gates; they are unchanged and are required before any live
application or performance claim.

The archive was checked with `git apply --check` and the strict whitespace
variant. No Cargo, formatter, native, profiler, or benchmark command was run
for this archive. The candidate is a measured hypothesis only; it carries no
speedup claim and must remain archive-only until the coordinator's baseline,
quality, differential, and paired measurement gates complete.
