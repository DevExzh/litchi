# 0704 longer native/refusal follow-up

The initial run raised 73 native metric/phase >5% review triggers and two
refusal p99 triggers. Real no-op total p50 was +6.43% in one candidate leg and
+0.21% in the other. The follow-up therefore retains the full 13-workflow
matrix and ten refusal cases, with four ABBA legs, 300 samples and ten warmups.
It uses the already frozen binaries; no candidate/probe source changes apply.

Run serially after other Cargo/performance processes finish:

```sh
python3 docs/performance/results/change-0704/measure-followup.py
python3 docs/performance/results/change-0704/audit_followup.py
```

The driver refuses to overwrite `followup/`, checks binary hashes, preserves
all 52 native process outputs and four refusal matrix outputs, and computes
statistics for every phase. This gives 15,600 native workflows and 12,000
refusal captures. Initial samples remain separate. The independent audit
recomputes statistics, pair changes, semantic identities, source hashes and
binary bindings from retained raw files. Every >5% trigger remains visible;
no sample is excluded or called noise without evidence.
