# 0478 implementation contract (coordinator)

The batch targets the public PPTX explicit-scratch fresh writer through all
three retained-name owners. This is an implementation plan, not a completion
or performance claim. All previously read accepted ADR hashes are unchanged.

ZIP adds `generated_names::{GeneratedNamePlanLimits, GeneratedNamePlanBuilder,
GeneratedNamePlan}`. Limits fields: `max_patterns: usize`,
`max_pattern_bytes: usize`, `max_entries: u64`. Builder `new(limits) -> Result`,
`push_literal(&str) -> Result<()>`, `push_indexed(first: u64, count: u64,
patterns: &[(&str, &str)]) -> Result<()>`, `finish() -> Result<GeneratedNamePlan>`.
Each indexed step emits patterns in supplied order for each index in the checked
range; patterns supply prefix/suffix around a canonical decimal run in the
final component. A plan has public `entry_count() -> u64` and
`max_name_bytes() -> usize`; crate-private `check_next(&str) -> Result<()>`,
`advance() -> Result<()>`, and `is_complete() -> bool`. Public callers cannot
advance or manufacture checked cursors. Builder/plan allocation and comparison
work depend on finite descriptor budgets, never expanded member count.

The initial proof language accepts canonical ASCII paths, restricts numeric
slots to final components and unambiguous non-digit stems/suffix boundaries,
and rejects ambiguous/unsupported patterns. Validate canonical production ZIP
normalization and prove exact and ASCII-folded uniqueness plus whole-component
ancestor disjointness. Check all literal/family, family/family and unequal-depth
cases symbolically, including a variable filename matching another family's
fixed parent component. Do not expand ranges or trust caller assertions.
No OPC/PPTX grammar constants enter ZIP. Syntax, overflow, budgets and
allocation failures remain typed refusals, before output.

ZIP Office writer replaces private name-set ownership with a private policy
(ordinary set or boxed generated cursor), carried through owned entries.
New constructor `with_writer_and_limits_and_spool_and_name_plan(writer, limits,
spool, spool_limits, plan) -> Result<Self, Error>`. Every name-taking route
must validate the exact next name in generated mode before publication;
reservation and recording keep no growing set. Advance only after successful
entry finalization. Finish refuses an unexhausted plan. Existing counters,
limits, poisoning, output progress and arbitrary-name behavior remain intact.
No public skip-validation flag or unchecked name token is introduced.

OPC exposes owner-local wrappers in phys_pkg:
`GeneratedPartNamePlanLimits` (same fields), `GeneratedPartNamePlanBuilder`
(same methods, absolute OPC names/prefixes), `GeneratedPartNamePlan`.
No ZIP types leak from these wrappers. The builder additionally validates
production PackURI syntax for literals and indexed endpoints; the ZIP proof
language guarantees interior decimal expansions preserve syntax. New
`PhysPkgWriter::with_writer_and_generated_plan_and_metadata_spool(writer,
plan, spool, maximum_spool_bytes, buffer_bytes) -> Result<Self>` installs the
checked lower plan. Ordinary PartNameSet stays unchanged. Generated mode
retains no PartNameSet/PreparedPartName or completed PackURI values; every
part call still must match the lower checked sequence and root is refused.
All named borrowed/owned routes and finalization must enforce the plan.

PPTX adds `StreamingPresentationScratchLimits { max_bytes: u64,
buffer_bytes: usize }`, plus `with_options_and_metadata_spool(writer,
slide_count, options, limits, spool, scratch_limits) -> Result<Self>` and
convenience `with_metadata_spool(writer, slide_count, spool, scratch_limits)`.
Provider bounds match OPC (Read + Write + Seek + Send + Sync + 'static).
Build and validate the exact current emission sequence from fixed members,
layout pairs, presentation members and slide pairs before touching the sink.
Reuse existing serialization functions, default API, semantic limits and
output bytes. No ZIP types or raw part-plan types enter ordinary signatures.

Root alone runs heavy commands under `/tmp/litchi-goal-0478/cpu.lock`, with
RUSTUP_TOOLCHAIN=1.98.1, and commits. Agents own assigned files only and must
coordinate API changes before editing shared files. No iWork edits. Preserve
shared Cargo caches and user-owned docs/GOAL.md and spec-gap-audit.md.
