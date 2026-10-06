# 0924 source audit: owned-reader requested-byte gap

Status: read-only diagnosis. No production source, harness, tests, builds, or captures were changed. The checkout is HEAD 679252bad65004085eb62a28b056b1a16b3f7468; existing unrelated files remain untouched.

## Scope and identity

I read docs/GOAL.md, docs/adr/README.md, every accepted ADR's original/non-iWork material, and the rejected 0779 input-growth record. I excluded iWork-only sections and ADRs 0028/0029 under the user's instruction.

The frozen 0923 candidate production source is still present:

| file | current SHA-256 | frozen 0923 candidate SHA-256 |
| --- | --- | --- |
| crates/litchi-opc/src/source_backed.rs | d28843caf8245bd6f0ff0f953b9b5545c4d9ad7a4e11f362d29ce226cc251cc2 | same |
| crates/litchi-opc/src/phys_pkg.rs | 3dfa9ec92c0d090b0bc753ecacc63f263975a470c267b79b03ae1c5898ac3ab8 | same |
| crates/litchi-docx/src/source_backed.rs | ebcf3af27a60bccc71a4a4e9eee30e4d4d7d4cf88633e25f1d143be15229ce73 | same |

The frozen 0923 production patch is external at /tmp/litchi-performance-0923/candidate-source-0923-v3.patch, SHA-256 77c2a9952335d5f16c24fbfc8d87e120301808258f70e990324fa21b8ea8e75d. The current benchmark harness differs from the frozen harness only in formatting; this audit makes no harness claim.

## Attribution

Retained Heaptrack JSON reports:

- /tmp/litchi-performance-0923/memory-audit-v3v4/parsed/009-owned_vec-pair-01-leg-1-baseline.json: requested bytes 151,803,020, peak live 24,234,029.
- /tmp/litchi-performance-0923/memory-audit-v3v4/parsed/017-owned_reader-pair-01-leg-1-baseline.json: requested bytes 423,388,539, peak live 24,234,083.

The difference is 271,585,519 requested bytes. The fixture is 2,101,427 bytes. At the 8 KiB buffer in crates/litchi-opc/src/phys_pkg.rs:1789-1817, it yields 257 nonzero reads. Because the fixture is far below the default 512 MiB input ceiling, the 258th read is the ordinary EOF read returning zero at lines 1826-1827; it is not the exact-limit one-byte sentinel path at lines 1791-1808.

The initial reservation plus the 256 actual Vec growth reallocations request capacities totalling 271,585,459 bytes. Heaptrack stack filtering confirms 256 growth allocations under litchi_opc::phys_pkg::read_limited in the reader profile and no such stack in the vector profile. The remaining 60 requested bytes are unassigned by this arithmetic; no allocator-overhead attribution is claimed.

The owned vector route moves its existing Vec through SourceBackedPackage::from_vec (source-backed lines 6798-6822). The reader route moves its input into OwnedIngressReader, then from_reader invokes read_limited and builds a second Vec (benchmark lines 50408-50430 and 50539-50595; source-backed lines 6783-6795). The avoidable mechanism is repeated destination growth/reallocation while preserving the same returned bytes, not additional ZIP payload materialization.

## Compatibility boundaries

read_limited must retain:

- max-input accounting and the exact over-limit sentinel/error;
- Interrupted retries, short reads, EOF, invalid read-count rejection, and typed allocation errors;
- input-limit-before-ZIP-validation ordering, covered by crates/litchi-opc/src/package.rs:3779-3803;
- deferred ordinary payload behavior after source-backed catalog admission.

ReadLimits::default().max_input_bytes() is 512 MiB (crates/litchi-opc/src/limits.rs:306-333). Sequential from_reader has no ExecutionContext; cancellation checks belong to explicit managed ReadAt constructors (crates/litchi-opc/src/source_backed.rs:6854-6892). A generic Read has no trustworthy length or seek contract.

The 0779 geometric-growth candidate already reduced cumulative requests but increased retained memory/RSS and was rejected. Reapplying that policy is out of scope.

## Disposition

Do not add a speculative length-hinted public API from this synthetic Cursor alone, and do not change generic from_reader yet. Prioritize an end-to-end source-provider intersection review: identify real callers that already possess owned bytes or a stable positional/seekable source, then compare their existing from_vec/ReadAt route against from_reader on representative DOCX/OPC workloads. A new source API is admissible only after such a caller and its source-freshness, limits, cancellation, and error-order contracts are demonstrated. No 0924 candidate is selected or implemented.
