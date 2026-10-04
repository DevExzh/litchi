# 0911 v1 first-child admission failure

The first 0911 Callgrind child is retained as a failed protocol attempt. It
must not contribute a snapshot instruction result. The corresponding structured
binding and raw observations are in
[`first-child-failure.json`](first-child-failure.json).

The child used the retained 0910 binary
`f4ad7ee158061cc80354c3c8674c9fe6a8100ed76a0c8f743653efcb06bb2e8d`, CPU 12,
the `materialized` route, zero warmups, and one requested sample. Its command
and workload gates completed successfully: the source is 17,017,139 bytes with
SHA-256
`61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757`, and the
output is 17,017,145 bytes with SHA-256
`23d6d7b8dd433ff453307ee56485efadc34d6b18b455a6858d9f565c4b1b6cd9`. Those
facts establish workload custody only; they do not admit the profile scope.

The v1 options toggled, zeroed, and dumped the exact owner
`litchi_pptx::package::model::Package::opened_presentation_transaction`. The
raw reader found two positive owner-scoped parts:

| part | summary Ir | positive owner caller | result |
| ---: | ---: | --- | --- |
| 1 | 126,617,815 | `litchi_perf_baseline::build_pptx_source_edit_corpus` | rejected setup interval |
| 2 | 1,000,613,578 | `litchi_perf_baseline::pptx_owned_benchmark::timed_iteration` | required measured edge, but not admissible alone |

Both parts have one owner call and conserve their owner self plus immediate
direct-child inclusive costs. The termination part is zero. The failure is the
extra positive owner interval and the wrong caller on part 1, not malformed
Callgrind grammar or a workload failure. The raw line tables are `???` in the
debug-info-free binary, so the caller names and graph edges are the authority;
no source line attribution is inferred from them.

The retained source explains the extra interval: corpus construction occurs
before `timed_iteration` and invokes the same transaction owner. With
`--zero-before` and `--dump-after` both bound to that owner, Callgrind emits one
positive part per dynamic owner call. Choosing part 2 or deleting part 1 would
discard contradictory raw evidence and violate the v1 admission gate.

The later v2 attempt must bind a single positive owner edge from
`timed_iteration`, retain a zero termination part, and preserve this failed v1
packet unchanged for audit history.
