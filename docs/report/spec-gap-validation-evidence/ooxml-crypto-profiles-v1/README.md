# OOXML encryption profile validation

The receipt identifies the baseline, exact candidate sources, fixture hashes and commands. The isolated all-feature gate passed 74 unit/integration tests and 3 doctests; the three external-corpus tests were then run explicitly and passed. Strict all-target Clippy and changed-file formatting checks passed.

The native tests cover Office 2007 Standard AES128/SHA1, Office 2010 Agile AES128/SHA1, Office 2013 Agile AES256/SHA512, LibreOffice Standard AES128/SHA1, and the Apache POI 60320 mixed wrapper fixture. They assert mode and integrity status, reject wrong passwords, decompress every ZIP entry and verify CRCs. LibreOffice requires the explicit missing-DataSpaces compatibility flag; default refusal is tested. Standard Encryption has no package authentication field; the Agile results are authenticated.

Synthetic third-party component vectors separately cover Standard AES192/256 and Agile SHA512/mixed key sizes. Their manifests distinguish synthetic provenance from native producer evidence. Native coverage does not establish every implemented key/hash combination. The allocation regression bounds growth across cryptographic segments; it is not a latency, total-memory or improvement measurement.

Payloads remain inert. No IRM/license handling, PKI trust or universal producer compatibility is claimed. Logs use deterministic gzip headers.
