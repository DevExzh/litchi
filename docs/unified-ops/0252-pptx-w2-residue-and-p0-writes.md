# 0252 — PPTX ADR 0034 W2 residue and P0 writes

Status: in progress on branch `uo2/p7c-pptx` (plan 0248 §6 steps 5-pptx and 7c).

## 1. Ingress options and patch conflict (W2 residue, part 1)

- `litchi_pptx::OpenOptions` (`source_limits`, `cache_limits`,
  `execution_context`, feature `encryption`: `encryption_limits`) is the one
  configured-open envelope. `OpenPolicy` names the lazy-only policies.
- Owning `Package`: `open_with_options`, `from_reader_with_options`,
  `from_bytes_with_options`, `open_with_password_and_options`,
  `from_reader_with_password_and_options`,
  `from_bytes_with_password_and_options`. Cache/execution policies are refused
  before input is read with `Error::UnsupportedOpenPolicy` (ops code
  `invalid_argument`).
- Deleted without aliases: `Package::{open_with_limits,
  from_reader_with_limits, from_bytes_with_limits, from_vec,
  from_vec_with_limits, from_vec_with_password,
  from_vec_with_password_and_limits, open_with_password_and_limits,
  from_reader_with_password_and_limits, from_bytes_with_password_and_limits}`
  and the six `open_with_*` combinations on `SourceBackedPresentation` and
  `SourceBackedPresentationEditor` (now `open_with_options`).
- Consumers formerly calling `from_vec(bytes)` now call
  `from_reader(Cursor::new(bytes))`, which keeps physical-source provenance
  (cross-slide copy authorization) at the cost of one bounded copy.
  `from_bytes` keeps its borrowed, provenance-free ingress. Affected
  perf-baseline legs (`pptx_*` routes in `tools/perf-baseline`) therefore
  include that copy; no performance claim about them is made here.
- `Error::StaleSource` is renamed `Error::PatchConflict` (ADR 0034 §6); ops
  still classifies it as `io_error` (the source changed under the handle).
- Glossary: both PPTX profiles ban the removed names; pptx-full also gains the
  primary-transaction bans pptx-minimal already had (plan 0248 §2 hole).

Not done here: positional `from_path_*`/`from_read_at_*` constructor
combinations on the source owners still exist (pending, as for DOCX in 0038).
