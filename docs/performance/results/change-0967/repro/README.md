# 0967 candidate reproduction

This external note binds the candidate-only patch to baseline `HEAD 893072657ad8b1d726bfce351b5ec2e98f084e94`. It compares exactly these three paths:

- `crates/litchi-doc/src/sprm.rs`
- `crates/litchi-doc/src/tracked_revision/codec.rs`
- `crates/litchi-doc/src/tracked_revision/package.rs`

The candidate source is `/home/zhuhe/litchi-performance-0967/source-candidate`. The Git-relative patch is `diagnostic-candidate-0967.patch` (23017 bytes, SHA-256 `10feea837f1f95d4f55cb4220d8a2192d91a89107931cbb9e0a716c433403de2`). It was generated from `/usr/bin/git show HEAD:<path>` and the corresponding candidate files.

Reproduction uses `/usr/bin/patch` with an explicit temporary scratch target:

```text
/usr/bin/patch --dry-run --batch --forward --fuzz=0 -p1 --directory <temporary-scratch> < /home/zhuhe/litchi-performance-0967/repro/diagnostic-candidate-0967.patch
/usr/bin/patch --batch --forward --fuzz=0 -p1 --directory <temporary-scratch> < /home/zhuhe/litchi-performance-0967/repro/diagnostic-candidate-0967.patch
```

The three patched scratch hashes exactly match `source-candidate`; the scratch was removed. The main checkout hashes were unchanged before and after. No build, workload, binary copy, or repository edit was performed. The machine-readable receipt is `reproduction-0967-v1.json`.
