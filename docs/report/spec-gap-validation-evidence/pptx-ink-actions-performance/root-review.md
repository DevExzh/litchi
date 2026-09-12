# Scaffold review gate

Independent design review approved this bounded plan. Root verified the
23-recipe/42-lane inventory, complete recipe use, and pinned helper hash.
The unchecked items below remain implementation review and execution gates;
the scaffold API is present, but design approval is not proof that those gates
have passed. No release profile or timing capture is authorized by this
document alone.

- [ ] The full source pin is
      `cf6fdb8e91dd232d7d762596763d2e9d8a5b9dbd`, or a reviewed replacement
      pin is recorded with its full source manifest.
- [ ] The manifest has exactly 23 recipes and 42 lanes; no Cartesian or
      generator-selected point is added without a plan revision; every recipe
      is used and every lane resolves to a real recipe.
- [ ] The future adapter calls only public `Package`/`Presentation` and
      source-backed snapshot/edit/patch APIs, with fresh mutable setup through
      `Package::from_opc_package`.
- [ ] Archive reopen uses `Package::from_vec_with_limits` with serialized
      retained OPC `ReadLimits`, separately from owner `ink_actions::Limits`.
- [ ] Synthetic fixtures are complete OPC owner closures derived from the
      committed helper and its recorded SHA-256; no neutral fragment, guessed
      graph, or native PowerPoint acceptance claim is admitted.
- [ ] Shared, distinct, case-equivalent, and Strict target recipes charge
      unique target bytes and retain all inbound graph edges.
- [ ] Anchor, target-byte, aggregate-byte, and graph-edge boundaries record
      exact/one-under/one-over nonzero outcomes with actual typed errors.
- [ ] MCE inactive choices, fallback bytes, opaque action descendants,
      internal unknown and external outbound diagnostics, and lexical
      relationship/source members have exact preservation gates.
- [ ] No-op, inverse, stale, signed, save, and reopen behavior is checked
      outside the timer and cannot publish partial state; stale mutations are
      applied to a separate raw OPC graph before public wrapping, and the
      content-type refusal is recorded as `Error::ContentType`.
- [ ] Setup, operation, validation, drop, retained-baseline, and post-drop
      allocation boundaries verify the stated equations; incremental peak-live
      bytes and whole-process RSS remain separate.
- [ ] Before any later build, the implementation materializes and hashes its
      isolated `harness/Cargo.lock`; the later runner captures source/toolchain/
      host/binary/fixture hashes and rejects root-lockfile substitution.
- [ ] The fixed later budget is three fresh processes, two warm-ups, and
      twenty measured samples per lane: 126 launches, 252 warm-ups, 2,520
      measured calls, and 2,772 calls total.
- [ ] Timing remains disabled until this checklist and the adapter are
      reviewed; no native acceptance or speedup claim is added.
