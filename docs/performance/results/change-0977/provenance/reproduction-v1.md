# 0977 reproduction note

This note describes bounded checks for the external root
`/home/zhuhe/litchi-performance-0977`. The repository packet is a source and
receipt copy. Packet assembly did not invoke perf, start a workload, build a
binary, or alter the external helpers.

## Frozen bindings

- plan: `plans/plan-0977-v2.json`, SHA-256
  `12abb6239a7f14a82a84e87306deffca4231c83377bf19a123cfb503a3684a1c`;
- helper seal: `receipts/helper-seal-0977-v2.json`, SHA-256
  `d51ea07cdca58e700a82ff8f4789fffc3dd40e4f284b086026458308fbd78e2e`;
- control audit: `analysis/cpu-review-control-audit-0977-v2.json`, SHA-256
  `2d2f70cb800239201a941d9f78d55129be2344bb6887f8fe34fc6dc2187d7927`;
- Office readiness: `analysis/cpu-review-office-readiness-0977-v2.json`, SHA-256
  `979bea121bdb5e4db5bdcd7a1d02ce212cc20390bfbca083f73bde555621032f`;
- Office failed-cohort audit: `analysis/cpu-review-office-failed-cohort-0977-v3.json`, SHA-256
  `050b5c6006122ce5ab5fb780971069196bcc03cc7a32ced93f262231d50c1e3d`;
- child-2 diagnostic: `receipts/child02-rawstack-diagnostic-0977-v3.json`, SHA-256
  `dbfb076ec8010d82bb81125569dcc94907b8fcadd67035572a4e4b5532631a1b`;
- Office root GO: `receipts/root-go-0977-office-v1.json`, SHA-256
  `ae2b92cc846d50c6e57c508cd2b0ea74421515b5310c516421e2b382a1e86193`.

The GO authorizes the registered Office protocol only. It does not turn the
harmless control or offline synthetic fixtures into Office evidence.

## Bounded offline checks

From a clean external root, set a fresh output directory for every helper
that writes outputs. These command shapes are retained by the sealed
receipts; they were not rerun while assembling this packet:

```sh
cd /home/zhuhe/litchi-performance-0977
export PYTHONDONTWRITEBYTECODE=1
/usr/bin/python3.14 tools/run_0977_v2.py --check-plan plans/plan-0977-v2.json
/usr/bin/python3.14 tools/run_0977_v2.py --self-test
python3 tools/test_control_replay_0977_v2.py
python3 tools/test_focuseddecode_0977.py
python3 tools/requalify_current_perf_0977_v1.py
python3 tools/posthoc_requalify_current_perf_0977_v2.py
python3 tools/reduce_0977.py --self-test
```

The requalification and posthoc commands consume retained external inputs
and write their sealed output directory; use a new versioned directory and
update the helper’s configured root before a replay. The control and decoder
checks are offline. The reducer self-test exercises assertions only. A
replay must not point an output-producing helper at an existing sealed path;
these helpers intentionally refuse to overwrite it. The commands above are
historical reproduction recipes, not a claim that this packet can be
replayed in place.

The synthetic integration and loader-gate helpers are also offline fixtures.
Their copied receipts bind the exact inputs and output manifests. Earlier
failed helper variants remain external historical lineage and are not
authoritative packet inputs.

## Recorded Office prefix

The historical registered launch argv, as sealed in
`receipts/root-office-command-0977-v1.json`, was:

```sh
/usr/bin/python3.14 -B tools/run_0977_v2.py --run \
  --plan plans/plan-0977-v2.json \
  --go receipts/root-go-0977-office-v1.json
```

It used a fresh registered output root, stopped at child 2, returned 2, and
was not retried. The bounded `office/capture-manifest-v2.json` records one
completed child out of twelve and a child-2 `RunnerError`; children 3--12
were never launched. Child 1 passed with 946 marker samples,
empty-callchain count zero, raw loss zero, terminal loss zero, and native
return code zero. Child 2 returned zero natively and reported terminal loss
zero, but its retained raw stack was rejected before decoder admission.

Child 1’s raw perf, mmap, packet, and corrected decoder logs remain external.
Child 2’s raw perf and mmap remain external together with bounded launch and
terminal metadata; it produced no packet or corrected decoder logs. The
final diagnostic records one kernel-shaped 72-byte sample with
`callchain_nr=0`, user stack size zero, raw offset `0x26f5b0`, and perf-text
stack offset `0x40`. The fixed scanner therefore rejects this variable
sample shape. The diagnostic and audit establish gate classification only;
they do not establish a causal mechanism or permit decoder mutation.

The reducer was invoked once against the failed manifest and returned 2 with
`rejected_as_expected`. It wrote no reduction output and made no partial or
synthetic substitution. There is no all-twelve-child reduction.

## Scope limits

The harmless control has 4,633 samples, raw and terminal loss zero, parent
CPU 30 and target CPU 31 sampled observations, and return code zero. These
facts support control custody only. The v9 integration is a synthetic
fixture, and its all-child arithmetic is not a native cohort. Current perf
requalification is offline and retains the `Invalid HEADER_EVENT_DESC`
warning; `evlist -v` supplies the exact event contract. No result in this
note supports a speedup, phase owner, attribution, causal explanation,
production change, or adoption decision.
