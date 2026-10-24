# 0362ugw — IWA native fixture provenance binding

The build-free native-qualification protocol at source `2dfc4f134d94195dd9b41f2d806887b38ef10c58`
(`uo2/budget-validation`) failed its fixture identity check because the
recorded `footnote-fixture.pages` source row still described the pre-lazy-entry
Pages test file. This correction updates only that tracked provenance row:

| Tracked source | Previous identity | Current identity |
| --- | --- | --- |
| `crates/litchi-pages/tests/body_footnote.rs` | 54,653 bytes; `56b7062bae16a99a63bae1a7d17b39e8a089fdd67ef7977fe2ec5fe5ee19f8e0` | 54,780 bytes; `3c6a522976632130d4e32a792fbf3c7baaa725f152f423115552b4d32f7c988f` |

The byte change is the reviewed `5b8339b4ddbcf2a060f2a2ca7b1dbe2be0a6847b`
deferred-entry migration: test consumers now use the fallible `Edit::before()`
accessor and exercise deferred missing-selector errors. The generator includes
this test file, so its complete source identity must remain bound. The actual
`build_fixture()` function used by the generator is unchanged: comparing lines
253–373 before and after that migration produced identical 3,687-byte slices,
SHA-256 `95ec152d3e1209491dadef55e1ef41629159c03c6b84062d6c0e65f94c19ace9`.
No fixture bytes were replaced or regenerated.

The generator inputs remain bound to the same files. The generator example is
982 bytes, SHA-256
`8bf4e6704cedf230a8b0e94a796f956757d1c2f07ced255716b68060788a36fb`;
the external wrapper is 2,425 bytes, SHA-256
`4ccacda3726db5afc4e4e651b202c0b517350e02810277175d988ea1cd190f29`.
After the correction, `fixture-provenance.json` is 2,326 bytes, SHA-256
`72e067ae44f5017f84cf8c883db72c0d4ec97ba79b9048f4c9e74b5d4e2ea86f`.

## Build-free validation

The exact protocol test was run without Cargo or native builds:

```text
python3 -m unittest discover -s docs/unified-ops/results/0361fda -p test_iwa_native_qualification.py
```

Result: 20 tests passed, 0 failed, 0 skipped. The external log is
`/home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/iwa-native-protocol-tests-2df-fixture-binding.log`,
119 bytes, SHA-256
`f2e68ff5eaca5112c86f24c4ad0c9d0d26001ccbe752852d8d57a9abf2564f99`.

This record validates the provenance protocol and fixture binding only. It
does not claim native timing, Cargo validation, or final acceptance. The
unrelated `docs/unified-ops/results/0361fda/constructor-probe.rs` was not
modified.
