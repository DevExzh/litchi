# Change 0954 evidence

See [the report](../../0954-xlsx-replacement-task-census.md).

Recompute into a new output path from the repository root:

```sh
python3 -B docs/performance/results/change-0954/census.py \
  /home/zhuhe/litchi-performance-0947/inputs/artifacts \
  --output /tmp/litchi-0954-census.json
```

`census.json` binds the retained manifest and each source/output ZIP's filename,
size and SHA-256. XML sidecar identities are verified against that manifest.
The ZIP files exceed 1 MiB and remain outside this checkout. To regenerate the
fixture family, build and run the exporter described in
[0904's reproduction instructions](../change-0904/README.md); then verify its
manifest and artifact identities before comparing this census. The committed
0904 manifest and XML sidecars also retain the identities independently of the
external ZIP files. No compiled binaries or raw ZIPs are added here.

`inputs.json` records HEAD, prior-turn progress and unchanged normative hashes.
`validation.json` binds the reducer, preservation policy source and existing
managed harness, and records byte-identical replay. Calculated widths assume
zero per-task floor and adequate explicit budgets. They do not measure an
executed preservation plan or a native performance improvement.
