# 0546 integrated XLSX candidate source note

This directory contains the complete five-file candidate patch for the
current baseline revision `6b866732489413366b40c94da7aa77bd37e719c3`. The
patch is `candidate.patch` and applies cleanly with `git apply --check`.
Production sources were not edited while preparing this bundle.

The candidate starts from the sealed 0544 candidate source snapshots. The
0544 snapshot manifest is
`docs/performance/results/change-0544/candidate/source-manifest.json`
(SHA-256 `fada10910fbb27b802baeac9404afd47af31bab4d0a182fffddb2eb1d8def1dc`).
The copied files are byte-identical to those snapshots before the two 0546
adaptations below. The 0544 source patch itself has SHA-256
`374aa71b6acf4aad1c51e4f22859ba9aa79dd2379c48e7251455828d9ce4612a`.

| Candidate path | Current baseline SHA-256 | Candidate SHA-256 | Source custody |
| --- | --- | --- | --- |
| `crates/litchi-xlsx/src/cell_values/snapshot.rs` | `2f3f839adc91f0da204aefc83ebe2bb605cc02d269346759155b7bda54abc0e9` | `c684c62aa523cc202027c733c92ad7cba3c91e456b4705c3f7dd9a7876cb53c2` | sealed 0544 snapshot, unchanged |
| `crates/litchi-xlsx/src/cell_values/validation.rs` | `52e7d2d18f59e716c686f1c4c59b5835632fe6981dc6789ef7b06b555d136c0b` | `19b4cb00420f895416debe3879f993ad292b1f79973b285b9cc440d6fae11522` | sealed 0544 snapshot, unchanged |
| `crates/litchi-xlsx/src/cell_values/shared_traversal_tests.rs` | absent | `5c56a86bd67d014c28823c9c315d6abbed8c614faf6f458795e3015ecfde3b95` | sealed 0544 tests plus 0546 oracle additions |
| `crates/litchi-xlsx/src/raw/worksheet/codec.rs` | `69719bc7a0aa0754ab4745f49baba077303ca7909f4e9adbc5243189e85fb9a7` | `98a5cf4db40e316cdd58a6904c80bdd11c06f86bd360f0d293651ca521648119` | sealed 0544 snapshot, unchanged |
| `crates/litchi-xlsx/src/raw/worksheet/mod.rs` | `7ed9276c713b8fcf8c84722bc62e58696912a11a1f02360e011059b830cf9b73` | `8aedf9ab085f5dc627e76d8e13841f211a1a53d4dd02f67cf389a6ffe70cd6e1` | sealed 0544 snapshot plus counted-bound adaptation |

The resulting patch has SHA-256
`bd9a0e25a5ce1a7dde78d568e4c1c998dd4c7103318269da31f201a125a9d539`.

The raw worksheet bound retains the exact 0544 formula:

```text
1 + initial_nonmarker + count('<' or '&')
  + count('>' or ';' followed by a nonmarker)
```

The first marker term uses the measured 0546 `counted.rs` algorithm: probe at
most 16 marker hits with `memchr2`, then count `<` and `&` independently in
disjoint 64 KiB chunks. The existing checked addition helper is parameterized
by the subtotal. The 16-hit and 64 KiB values are private descriptive
implementation constants. The diagnostic `inline(never)` boundary is not
carried into production. No unsafe code, dependency, public API, event-limit,
source-limit, or validation-policy change is introduced.

The remaining 0544 behavior is copied exactly: source UTF-8/MCE/x14ac
eligibility, the shared validator and ordinary parser transition, authoritative
fallback and historical x14ac error retry, parser error ordering, source
ownership, and value-only publication semantics. The test module remains
`cfg(test)` and is included by the existing validation test path.

The direct `NsReader` oracle now uses the pinned diagnostic configuration:
`allow_dangling_amp = false`, `allow_unmatched_ends = false`,
`check_comments = false`, `check_end_names = true`,
`expand_empty_elements = false`, `trim_markup_names_in_closing_tags = true`,
and `trim_text(false)`. It adds exact-cap, cap-plus-one, final-byte, and
4,095/4,096/8,191/8,192/65,535/65,536 successor-pair cases while reusing the
existing independent lexical bound, direct reader prefix oracle, and
source-preservation/error-order tests. The additions do not duplicate the
0544 workflow controls.

ADR custody is bound to
`docs/performance/results/change-0546/adr-manifest.json`, SHA-256
`e815a51327a39956a3e5e3adfd3fc76b7942b1c65124c66189ab7b98f8bf6893`; its 30
accepted ADR/index hashes were unchanged when checked. The relevant accepted
constraints are the API/dependency topology and priority records (0001,
0002), immutable edit/patch behavior (0003), bounded performance evidence
(0005), validation/security/preservation (0006), migration verification
(0008), archive ownership (0010, 0011), and current topology (0024).

Preparation intentionally performed no Rust build, test, or commit. The
isolated 0546 scanner evidence justifies this fresh integrated campaign but is
not integrated performance proof. Its clustered control has an approximately
174% relative scanner regression (about 30 microseconds); the integrated
campaign must retain and review the valid sparse 1/2 controls before any
runtime retention decision.
