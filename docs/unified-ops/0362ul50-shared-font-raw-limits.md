# 0362ul50 — Shared font raw limits

This packet makes the shared font preparation and PowerPoint EOT boundaries
follow the required-limit rule from 0351 §2.5.  It is based on source
checkpoint `267ea715266ffc3bd754bfbfa4896b139adad9fc` and does not change the
fixed OOXML obfuscation transform.

## Canonical APIs

`embedding::PreparationLimits` is the required policy for both preparation
entry points.  Its dimensions are `max_requests`,
`max_glyphs_per_request`, `max_family_name_bytes`, `max_program_bytes`, and
`max_total_program_bytes`, and `max_subset_working_bytes`.  The first five
fields describe source admission and retained output; the last is an
independent envelope for temporary memory used by the optional allsorts
subsetter.  The standard default for the last field is the finite 1 GiB
policy, subject to bounded corpus validation; owners can raise or narrow it
through their retained preparation options.

```text
embedding::prepare(used_glyphs, mode, limits)
embedding::prepare_with(resolver, used_glyphs, mode, limits)
```

`embedding::powerpoint::Limits` is required at every EOT raw boundary.  Its
dimensions are `max_input_bytes`, `max_output_bytes`, `max_font_bytes`,
`max_sfnt_tables`, `max_name_records`, and `max_name_bytes`.

```text
embedding::powerpoint::View::parse(bytes, limits)
embedding::powerpoint::restore(font, eot, limits)
embedding::powerpoint::encode(font, intent, limits)

subset::glyph_ids(font_data, codepoints, limits)
Subsetter::subset(&self, font, glyph_ids, limits)
```

The former no-limit `data`, `View::parse_with`, and `restore_with` entry points
are removed; no aliases or deprecated twins are retained.  `prepare_with`
remains the resolver-supplied entry point, while automatic `prepare` retains
the system-resolver capability and now receives the same explicit policy.

## Owner policy and consumers

DOCX retains preparation policy in `SemanticLimits` (`fonts` and
`fonts_value()`), with `Builder::font_preparation_limits` selecting it before
the package-owned automatic embedding path.  PPTX retains preparation and EOT
policies on its authoring `Builder`; `build` passes them through publication to
automatic preparation and explicit `encode(Intent::PreviewPrint, limits)`.
The existing PPT owner already carries its EOT policy and now calls canonical
`restore`/`View::parse` with it.  Defaults remain the owners' standard defaults,
not per-operation fallback arguments.  DOCX maps a shared preparation limit
failure to its typed `FontLimit(FontPreparation)` error instead of flattening
it to an unclassified `Other` error.

All live consumers in `litchi-fonts`, DOCX, PPTX, and PPT, including the
embedded-font tests, were migrated in this packet.  A tracked-source search of
the standalone/performance Rust harnesses found no direct shared-font raw API
consumer; the few tracked manifests that mention `litchi-fonts` only declare a
workspace dependency.  Historical receipts and source pins were left intact.

The dedicated `fonts-minimal` and `fonts-full` glossary profiles register the
required signatures with their exact nominal paths, argument names, and typed
limits.  The minimal profile covers resolver-supplied preparation and EOT;
the full `automatic` profile additionally covers system-resolver `prepare`.
The DOCX/PPTX/PPT profiles consume these APIs but do not claim ownership of
the shared crate's public paths.

The existing font tests continue to cover preparation request/program limits,
EOT input/output/name/table bounds, intent handling, round-trip restoration,
and byte-fidelity checks.  Subset preparation now checks source bytes plus the
requested glyph-ID staging bytes before mapping, with exact and one-byte-over
tests.  The consumer enters one pinned Allsorts budget scope before glyph
mapping and keeps it active through source parsing, provider creation, CFF/glyf
expansion, and output writing.  Its working cap is
`max_subset_working_bytes`; its retained-output cap is the checked minimum of
`max_program_bytes` and the remaining `max_total_program_bytes`, so a nested
subset helper cannot reset either cumulative counter.  The dependency packet
owns the bounded parser/index/writer implementation; this consumer does not
invent a source-size multiplier.  The source bytes and exact requested
glyph-ID staging capacity are charged into that same scope before mapping;
they are not a separate allowance from the adapter's cumulative requests.
Mapping converts the `Glyphs` `u64` count to `usize` with a typed overflow
refusal and reserves the requested-codepoint count plus the retained `.notdef`
slot.  The raw `Subsetter::subset` route charges exactly the supplied
`&[u16]` length; it does not invent another `.notdef` slot.
The public raw mapping and pluggable-subsetter entry points require the
caller-selected `subset::Limits` as well: `subset::glyph_ids` and
`Subsetter::subset` create a bounded raw session from that required value.
Preparation uses crate-private scoped helpers with one shared session, so
its mapping and writing remain cumulatively charged without exposing an
ambient or unbounded raw route.
Allsorts parser, writer, CFF, and subset errors become a typed font limit only
when the active scope exposes an actual `BudgetFailure` snapshot; structural
`LimitExceeded` values without that snapshot remain ordinary subsetting
failures.  The snapshot preserves working bytes, output bytes, and work units
with their rejected requested amount, so the consumer never infers a policy
failure from an error string.
System-font admission uses the same typed `LimitExceeded` error for a selected
program cap; it is not flattened to the legacy `ProgramTooLarge` shape.  The
subset tests exercise the real `prepare_with(..., Mode::Subset, ...)` route
against the tracked 4,488-byte Noto Sans Lycian fixture, covering staging
admission, output admission, and a successful reduced program.  The executable
subset example also uses this canonical route rather than invoking the raw
subsetter without its selected scope.
Initial source-only validation for this packet was
`rustfmt --edition 2024` on all changed Rust files and `git diff --check`.
Cargo, native, and performance runs were intentionally deferred until the
coordinator granted a bounded validation slot; no performance claim is made
by that source-only checkpoint.

## Final bounded validation

The complete font lane was then run at pre-squash source commit
`e69d10883064b5d2a3a5251d81ac0a640f8ff9eb` with the coordinator's guarded
`run_safe_gate.py` (`1904177040fc7c523582433028d9bba56efc30fe0d9d96beb554a2fab343691e`).
All eleven rows passed: the vendor budget compile and tests, `litchi-fonts`
tests and Clippy, DOCX/PPTX/PPT checks and Clippy, the three focused consumer
font tests, and both font glossary profiles.  The external receipt is
`/home/zhuhe/code/litchi-worktrees/.coord/shared-font-raw-validation-e69d10-full-20261025/state/driver-result.json`
(100,030 bytes, SHA-256
`4ab3c781f5319af72f730d53fd0998decc6805fb408be1aae85c6abe143edd62`).
It records the unchanged lock hashes: root
`efff87af6c339a44e267987c5c79ba3fb1910cef734f0a9f212f08f865dac811`,
native-resave
`989da70f10deab4460a77e4afa3a1b840b4697bbea2c07e134e22d9b4c5084ae`, and
perf-baseline
`de1f8ae53b3455a4246da0b9340896881cd72ca0295b8e52444e0d55e6c9c4ea`.
The reused guarded target ended at 5,493,966,365 bytes under the 20 GiB cap;
no validation binaries or large evidence are committed.
