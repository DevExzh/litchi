# 0968 replay references

The canonical CPU record is an offline reduction of retained 0942 perf-script
callchains. It has no 0968 build, workload, latency, allocation, or adoption
claim. Raw perf-script files, the retained 0942 binary, formal reduction, and
capture manifest remain outside the repository; their paths and hashes are in
`../provenance/retained-cpu-binding-0968-v1.json`.

The analyzer is copied at `../tools/analyze_cfb_commit_v1.py`. Its fixed output
guard requires a fresh external directory. A replay reference is:

```text
PYTHONDONTWRITEBYTECODE=1 /usr/bin/python3 /home/zhuhe/litchi-performance-0968/analyze_cfb_commit_v1.py --output-dir /tmp/litchi-0968-cpu-replay-<unique>
```

The output directory must be new. The independent review records a successful
read-only replay in `/home/zhuhe/litchi-performance-0968/replay-independent-v1`;
the replay JSON matches the sealed v2 output after removing `created_utc`, and
the Markdown is byte-identical.

The decision utility replay is recorded separately in
`../replay/decision-replay-receipt-v1.json`. The utility source remains in the
main checkout and is bound by `../provenance/utility-binding-0968-v1.json`.
