# 0978 offline reproduction note

This note covers only the selected variable-layout scanner, lifecycle helper,
scoped reducer replay, and retained-child diagnostics. The external root is
`/home/zhuhe/litchi-performance-0978`; the 0977 input and decoder dependencies
are the sealed files cited by `dependency-0977-v1.json`. Use a separate
external copy or fresh output paths. Do not overwrite sealed receipts or place
generated output in this repository packet.

## Frozen source bindings

- policy/design review: `review/independent-design-review-v3.json`, SHA-256
  `7181d28ff5f18dc21ec60ba30472328676eafb9da0ece36f4d5c08bc97c6fb39`;
- scanner: `tools/raw_stack_0978.py`, SHA-256
  `5f844b82a40426662ffaed5ba21dd0687fefa3f98ecc233dc572eeb3041017b7`;
- lifecycle helper: `tools/lifecycle_ledger_0978.py`, SHA-256
  `6296e2dd20f911b2d989f1a8c7b9e337c9a34efc5c0f04831a50111192b9e7d0`;
- final reducer: `tools/reduce_0978_v4.py`, SHA-256
  `1c40193d0debcea9feb696d82f029cc058b10e39e20ceb2183e2444223fc756c`;
- validated reducer predecessor: `tools/reduce_0978_v2.py`, SHA-256
  `e925755639ccdca3b75ed5b6e64265f711f506b6e6478bcee9a94990ce08521f`;
- original decoder used for the retained fresh result:
  `tools/focuseddecode_0978.py`, SHA-256
  `866723324881a57ce7c4282cdfd2dabe497397b1612e8d7c4f975cb7e9fe2134`;
  the copied `analysis/focuseddecode_0978-pre-full-lifetime-guard-866723.py`
  is a byte-identical snapshot for custody and is not the self-path used by
  the original plan;
- final decoder guard: `tools/focuseddecode_0978_v3.py`, SHA-256
  `9596e246c7c627fb3810e7c6c6d3dda12af58b62477f1f1233654c4256e2dac9`;
- final decoder guard receipt: `receipts/focused-decoder-full-lifetime-guard-0978-v3.json`, SHA-256
  `689dc4ea2c1fbb85fee7732d038db4c7b46f6f4236327263122b3930bc51dd31`;
- final decoder guard review: `review/focused-decoder-lifetime-guard-review-v2.json`, SHA-256
  `01d4ae822ef4e69e39e7d4a442b697d24cf0d9e85c6e5e994f4ab97eadb8df5c`;
- fresh offline result: `analysis/fresh-offline-child02-result-v1.json`, SHA-256
  `cde5d9f72d3d8c7f4c08641746efde82b174cd2f5949bbb4265de58197bdb76`;
- final fresh packet validation: `analysis/fresh-offline-child02-packet-validation-v4.json`, SHA-256
  `a970f7309ab455c74a73b80c9ff7d99fcd2859b8ba05801aaf8c6f3a59643aab`;
- native-boundary runner source: `tools/run_0978_v1.py`, SHA-256
  `c24da0749dd6d5dd1464e23dffbb9a3650634f0a803ac3a372a6d365cfc9837a`;
- scanner controls: `receipts/raw-stack-scanner-controls-0978-v2.json`,
  SHA-256 `7b9445c1a3c95bd8bb69366e8fdcc9aaa3a1a1c87055809acbd0caa2a51769d3`;
- lifecycle review v4: `review/lifecycle-ledger-review-v4.json`, SHA-256
  `9fbd3c6d833f7267aaee5f75e74139efd41f5fcf29bf7cc8b37a6f4d3a1226e2`;
- historical native-boundary review v3 (superseded):
  `review/independent-native-boundary-review-v3.json`, SHA-256
  `b793c8c393cb29b8ba79dfc0cef1ecaaac420a4cf5f4c8efdd1f59a784c1aaea`;
- final native-boundary review v4:
  `review/independent-native-boundary-review-v4.json`, SHA-256
  `a3e29e8a9512ac43aaeb7bf736ac709549d67682cfc059cde63428de6b1a041a`;
- final readiness review v5: `review/final-readiness-review-v5.json`, SHA-256
  `f424f6ebac813616656abc82ac595de3c1426a53058cc17d35c2496106347a7f`;
- control equivalence v5: `receipts/control-boundary-equivalence-0978-v5.json`,
  SHA-256 `c4b0c814ce93cb3a3f2bc3d30c92a91c0f8ca814fbc9003d38200a162a91432d`;
- reducer scoped approval v3: `review/independent-reducer-scoped-approval-v3.json`,
  SHA-256 `ef96bd4601f339ad6ac282cd94cdca8591051d1f05fb2eded7617281c032458f`;
- reducer v4 final review: `review/independent-reducer-v4-final-review-v1.json`,
  SHA-256 `2b713bc35b0a46e447e9b41eb475f65b8f2f23b7bba15537c354e987fa3fde75`.
- final pre-capture plan: `plans/plan-0978-v3.json`, SHA-256
  `9849dfd914463db998f02df7505db85ff02eb84215143d4ac1e67af267fa6424`;
- final helper seal: `receipts/helper-seal-0978-v3.json`, SHA-256
  `a9735d56b4a73fab0ef8e6f14a2e05b19813e686c99e850283c06f4c2c8281fe`.

## Offline controls

From the external 0978 root, with bytecode disabled:

```sh
cd /home/zhuhe/litchi-performance-0978
export PYTHONDONTWRITEBYTECODE=1
python3 -B tools/test_raw_stack_0978.py
python3 -B tools/test_lifecycle_ledger_0978.py
python3 -B tests/independent_raw_stack_tests_v1.py
python3 -B tests/independent_reducer_mutation_tests_v2.py
python3 -B tools/test_full_lifetime_evlist_guard_0978_v1.py
python3 -B tools/test_focuseddecode_command_boundary_0978_v1.py
python3 -B tools/reduce_0978_v4.py --self-test
python3 tools/replay_raw_stack_0978.py --child 1 --output /tmp/replay-0978-child01-new.json
python3 tools/replay_raw_stack_0978.py --child 2 --output /tmp/replay-0978-child02-new.json
```

The reducer mutation command and the external-root `tools/test_reduce_0978.py`
are retained v1 lineage controls; the final v4 proof is the packet validation
below. All commands write only fresh external outputs. Do not reuse sealed
receipt paths. The offline controls do not make a native cohort or production
integration complete.

The final packet validation command, inherited by the v4 metadata review,
reads the retained fresh child and writes a new receipt path:

```sh
/usr/bin/python3 -B tools/validate_packet_integration_0978_v4.py \
  --integration-root /home/zhuhe/litchi-performance-0978/fresh-offline-child02-v1 \
  --child child-02 \
  --output /tmp/litchi-0978-packet-validation-<unique>.json
```

The checked-in `analysis/fresh-offline-child02-result-v1.json` was produced
with the original 866 decoder against retained inputs. The v3 decoder source
is bound by the guard review; no new perf or decoder run is implied by this
packet.

The latest offline child-2 reducer replay used the equivalent command shape
recorded in `receipts/replay-receipt-v6.json`:

```sh
/usr/bin/python3 -B tools/replay_attribution_0978_v1.py \
  --output <fresh-replay-receipt>.json \
  --ledger-output <fresh-derived-ledger>.json \
  --lifecycle-stream <bound-lifecycle-stream-v2.json> \
  --stream-generation-receipt <bound-child02-build-ledger-receipt-v3.json>
```

This is one retained child from the 0977 capture, offline only. The full
ledger and lifecycle stream remain external. `reparse_lifecycle_0978_v1.py`
provides the corresponding one-pass raw-dump reparse and requires explicit
`--dump`, `--output`, `--receipt`, `--pid`, and `--cpu` paths; use fresh output
files. The final retained reparse receipt is
`receipts/lifecycle-stream-v2-receipt.json`; it binds the reparse helper and
script hashes and records the external stream hash. No perf subprocess is
started by either offline replay.

## Diagnostic lineage and native boundary

The direct corrected perf-script run has return code 0 and is separate from
the earlier focused-decoder process that ended with SIGTERM before natural
completion. `focuseddecode-run-correction-v1.json` withdraws that earlier
interpretation. The child-2 diagnostic and reducer replay preserve unknown
periods and one-child scope; they do not assign source ownership or causality.
The final decoder and reducer receipts remain offline proof for retained data.

`run_0978_v1.py`, `plan_binding_0978_v1.py`, and `prepare_plan_0978_v1.py` are
copied for source custody. The final plan v3 and helper seal v3 are copied
pre-capture outputs; re-generation must use a copied external root and fresh
exclusive output paths. The normal helper-generation CLI is:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 -B tools/prepare_plan_0978_v1.py \
  --seal-output <fresh-helper-seal-v3.json> \
  --decoder tools/focuseddecode_0978_v3.py \
  --reducer tools/reduce_0978_v4.py
PYTHONDONTWRITEBYTECODE=1 python3 -B tools/prepare_plan_0978_v1.py \
  --helper-seal <fresh-helper-seal-v3.json> \
  --decoder tools/focuseddecode_0978_v3.py \
  --reducer tools/reduce_0978_v4.py \
  --output <fresh-plan-v3.json>
```

The runner's `--run` path was invoked exactly once after the final readiness
review and root GO. The historical command was:

```sh
/usr/bin/python3.14 -B /home/zhuhe/litchi-performance-0978/tools/run_0978_v1.py \
  --run \
  --plan /home/zhuhe/litchi-performance-0978/plans/plan-0978-v3.json \
  --go /home/zhuhe/litchi-performance-0978/receipts/root-go-0978-office-v1.json
```

It stopped serial execution at child 3 after the loss gate observed four
terminal lost samples. Children 4--12 were never started; no retry or
selective replacement was attempted. Do not rerun sealed paths, retry child 3,
or treat the partial capture as a complete cohort. The external descriptor
binds all 92 capture files and the compact custody summary records the exact
failure disposition.
