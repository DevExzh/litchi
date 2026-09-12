# Replaying the rejected checkout

Verify the retained source bundle against a repository containing its
prerequisite, then fetch the failed checkout object from the bundle:

```bash
git bundle verify captured-source-16cf1e3c6.bundle
git fetch captured-source-16cf1e3c6.bundle \
  901994438a3202fae2c3a13535f538c222a08240:refs/heads/docx-styles-effects-projection-reuse-901994438
git worktree add /var/tmp/docx-styles-effects-projection-reuse-901994438 \
  refs/heads/docx-styles-effects-projection-reuse-901994438
git -C /var/tmp/docx-styles-effects-projection-reuse-901994438 rev-parse HEAD
```

The final command must print
`901994438a3202fae2c3a13535f538c222a08240`. The bundle's advertised head is
the later repaired source `16cf1e3c6`; the failed checkout is an ancestor in
the same bundle. The bundle requires
`d000d977b99e03f8542c7dae74acf767a91b1feb`.

The raw files came from this external results path:

`/var/tmp/litchi-docx-styles-effects-projection-reuse-smoke-results-901994438-20260912-a`

The runner used this disjoint target path:

`/var/tmp/litchi-docx-styles-effects-projection-reuse-smoke-target-901994438-20260912-a`

The target no longer exists because the runner's cleanup trap removed it after
the verifier rejected the stale corpus pin. The raw receipts preserve those
absolute paths as historical provenance; replaying source does not recreate
them or turn this failed run into an approved capture.
