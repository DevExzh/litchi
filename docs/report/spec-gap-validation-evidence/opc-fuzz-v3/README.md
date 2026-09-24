Bounded OPC fuzz validation
===========================

This extends the existing OPC fuzz target. V3 adds meaningful internal-relationship/main-document resolution, topology write/reopen checks and 22 deterministic exact/over/under limit cases through source-backed, eager and probe ingress. ArchiveTotalEntries and XmlDepth are not isolated by the current fixture; no all-limit claim is made. XML replacement/addition success is checked, while reopened XML byte equality remains outside this receipt.

The retained AddressSanitizer/libFuzzer campaign completed 10,000 executions without crash/timeout, using the actual 99-package locked fuzz workspace. One MiB is the input ceiling, not the largest exercised fixture. Independent review verified instrumentation, source/binary/lock bindings, corpus hashes and primary preservation. Root repeated the locked smoke successfully and verified the compiled OPC/core/ZIP/package source is unchanged from the capture base to integration.

The four earlier pending fuzz paths are superseded with exact preimages retained. The unrelated primary OPC README stays pending. Historical setup logs and earlier campaign attempts are excluded. Raw final logs are gzip-compressed without content changes; the corpus archive contains all 352 final files and replay seeds. Extract fuzz-Cargo.lock.gz to the fuzz workspace Cargo.lock before locked replay; author README-v3.md and receipt-v3.json retain the exact commands and original paths/toolchain versions. This is bounded safety evidence, not a completeness or performance claim.
