# Exploratory evidence compaction

The three executable payloads were hash-verified before removal with the
recorded `binary.sha256` files. The two retained valid source archives were
also verified with their `source-snapshot.sha256` files. Raw JSON receipts,
`/usr/bin/time -v` output, build commands, build provenance, source manifests,
valid source preimages, and valid source archives remain.

The invalid build directories retain their failure marker, captured source
manifests (before/after where available), manifest/hash records, build log, and
process/source status. Their
copied source-preimage trees and Cargo metadata were redundant after the
manifest records were captured and were removed. No invalid executable was
retained.

The retained exploratory evidence is one fresh process and one measured
sample per lane. It remains excluded from acceptance or final performance
claims.

Verification before compaction:

```sh
for f in \
  results/exploratory-before/binary.sha256 \
  results/exploratory-detach-before/binary.sha256 \
  results/exploratory-namespace-before/binary.sha256; do
  sha256sum -c "$f"
done
for f in \
  results/exploratory-detach-before/source-snapshot.sha256 \
  results/exploratory-namespace-before/source-snapshot.sha256; do
  sha256sum -c "$f"
done
```

The recorded pre-removal checks returned `OK` for all five files:

```text
exploratory-before/bin/xlsx-svg-lifecycle-profile
  66d1a25f7499bb212db61a7e9ae03af17888e7c749356c2f15344c8c5735d8de
exploratory-detach-before/bin/xlsx-svg-lifecycle-profile
  1cdb1564fba8dbb173cef453910dc91f984ce96ec0940209a384f0a822601691
exploratory-namespace-before/bin/xlsx-svg-lifecycle-profile
  7b8d6c8535cce624ab721f89940fedc3e8361f635a1242cfa2075f13ade19cce
exploratory-detach-before/source-snapshot.tar.gz
  eb5228d7c3b7da45d6fa253e5e9c2c3cf7c55aa05dd326d2484e722c4735d211
exploratory-namespace-before/source-snapshot.tar.gz
  c992d93ef4723b8d923f47531e6893410c7c9430315f7674b5ffa6fd187b7266
```

The executable files were then removed. Their hash records, build commands,
raw receipts, and source provenance remain for audit; these bundles are no
longer directly runnable without rebuilding from a retained source snapshot.
