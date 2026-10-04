# 0911 result review

**Status: pass.** The v3 packet is admitted for bounded same-binary Callgrind
attribution of the PPTX owning cold snapshot. This review was read-only: it ran
no build, workload, profiler, or production change. The final report was
reviewed at SHA-256
`9adf2b9de5b30a261ae85dc96ac78638978207ece5ec3e1f85f8362b5d759967`.

The retained raw profiles were independently re-parsed with the retained 0845
packet-local Callgrind grammar. All four legs and eight positive parts passed
these checks:

- the event set is exactly `Ir`, with part 1 as setup and part 2 as measured;
- each part has exactly one positive incoming edge to
  `litchi_pptx::package::model::Package::opened_presentation_transaction`,
  with one call and an inclusive count equal to its summary;
- setup edges come only from
  `litchi_perf_baseline::build_pptx_source_edit_corpus`, and measured edges
  come only from
  `litchi_perf_baseline::pptx_owned_benchmark::timed_iteration`;
- positive summary equals `totals`, owner self plus direct children equals the
  summary, and all function self counts equal the summary; and
- every part 3 terminal dump is an empty zero-Ir profile.

The independently checked measured owner counts are 1,000,666,812 and
1,000,594,047 Ir for materialized A/B, and 1,099,868,747 and 1,099,912,679
Ir for owned A/B. The two route medians recompute to 1,000,630,429.5 and
1,099,890,713 Ir, matching the +99,260,283.5 Ir (+9.92%) descriptive
guest-instruction comparison recorded in the candidate assessment. Setup
counts remain retained but excluded from that comparison.

All four semantic receipts independently bind the same source archive
(`61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757`) and
output (`23d6d7b8dd433ff453307ee56485efadc34d6b18b455a6858d9f565c4b1b6cd9`),
and each raw admission and semantic receipt is `pass`. The reused 0910 custody
receipt reports 9,355 reconstructed source files, 72 capture artifacts and 24
native legs without a build or workload run. Re-reading the 24 retained 0910
reports gives 240 snapshot samples per route and the report's lower pooled
p50 values: 12,349,500 ns materialized and 15,865,198 ns owned.

The frozen analyzer was also tested fail-closed using only deep-copied
in-memory parse trees. It rejected an unknown owner caller, an extra positive
owner edge, a summary/totals conservation mismatch, and a nonzero terminal
dump. The mutation receipt is
`analyzer-checks.json`; retained packet files were not modified by those tests.

Replay is byte exact before and after cleanup. Cleanup removed the three
marked temporary capture roots (195 logical bytes), created no worktree or
build tree, and left the retained raw profiles and archives in place.

The admission is limited to this generated 200-slide, media-rich fixture and
guest instruction counts. It does not establish native SHA dispatch, native
hash speed, elapsed-time attribution, allocation/RSS, physical bytes, or a
production optimization. The software SHA path and deferred-payload path are
documented hypotheses requiring native dispatch confirmation and a later
end-to-end measurement. No production source was changed.

Hash bindings for the reviewed packet are:

| artifact | SHA-256 |
| --- | --- |
| final report | `9adf2b9de5b30a261ae85dc96ac78638978207ece5ec3e1f85f8362b5d759967` |
| canonical analysis | `ac6fa98aed88d9594fe2bafef7c38cb06f08dc6b84eab574b919fd11bdedf76f` |
| frozen analyzer | `d52b4f1e5ccfd651c4256e89a2fdd94e2db2d32541abed85a8349cf3b317bc38` |
| mutation checker | `fb3a03780c67f000e1e6ecc2f3bd608a1969d01f23e3041ff95c5e1a84516665` |
| mutation receipt | `98fe1bbbfa42120d7775b9f93f89778b875b49a1f6884421b741250628add321` |
| v3 protocol review | `8c255b6f03424b73f58bc6d819c33b9ebcacfeacd0447d385b5168d80ea4c901` |
| candidate assessment | `39f91c22c4250342f9c879f76746c33169618b45784f74b77d86a28b6b832084` |
| plan v3 | `28366047e447183badd9d7a91f10e82bc6e17a5de82529a4ef8da859b4034089` |
| replay before/after | `5820a24017c6ab7c53a85e6edccca296d44fff73e271a370a0661c164e9e50da` |
| cleanup receipt | `21dc1d563bbe1a32be5010dba7b81d46d1781b9800900b7db3f3f7b2e2dd484b` |
