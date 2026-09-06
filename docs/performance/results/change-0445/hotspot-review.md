# Plain-source attribution and next candidate

Plain normal p50 is 1.074–1.079 / 5.230–5.280 / 18.869–18.886 ms for
64/1024/4096 existing Parts. Relative to the matched observed mode, medium
falls 34.970–35.524% and large 69.576–69.635%. These differences measure source
observer overhead. Operation allocation calls/requested bytes/above-entry peaks
are identical between modes. The one repeat flag is observed-allocator tiny p99,
R2 versus R1 -7.928%; it is instrumented tail drift, not a production regression.

The observed run-frame subset attributes 68.989% self period to the instrumented
reader. Plain has no such frame. Its leading self rows are diffuse: SHA-256
5.521%, memcmp 5.348%, PackURI hashing 4.562%, then URI scanning/validation and
Part-name indexing. Inclusive plain run-frame rows expose the lifecycle split:
publication 59.290%, opening 40.552%, source catalog 32.200%, and content-type
map parsing 29.291%. Inclusive rows overlap. The run frame contains source setup,
endpoint probes and warmups as well as timed work; its percentages must not be
used directly as exact timed-region or Amdahl speedup estimates.

At the captured source, `ContentTypeMap::from_xml` is used during mandatory
catalog opening. `SourceBackedPackage::write_topology_to_stream` re-reads and
parses the exact source content-types manifest after freshness checks because
managed opening deliberately does not retain an uncharged large manifest.
`content_types_with_overrides` validates the generated candidate by parsing it.
These ownership, bounds and publication checks remain requirements. A naive
cache or skipped parse would not be an authorized optimization.

A smaller allocation candidate is visible in `content_type.rs::inspect_element`:
`required_attributes` returns an owned Part-name String, then `PackURI::new(&partname)`
passes it by reference into an `Into<String>` constructor. The Part-name variable
is not used afterward. Moving the owned String could remove one clone per parsed
Override without changing validation or ownership. The repeated map parsing makes
that handoff worth measuring against the plain lifecycle and allocation baseline.
No production edit or performance acceptance follows from this inspection alone.

Next batch: capture before/after against this matched plain-source case, implement
the smallest ownership transfer, run the existing content-type/URI/topology/error
and resource gates, and keep it only if the frozen practical gate is met. Avoid
selecting SIMD or changing map hash security before removing measured redundant
work. Broader semantic CRUD, native, cold/range and scaling obligations remain.
