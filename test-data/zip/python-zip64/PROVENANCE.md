# Python ZIP64 interoperability archives

Small archives written by CPython's standard-library `zipfile` with
deterministic names, timestamps and payloads, including `force_zip64` output:
ZIP64 local headers with ordinary central sizes, signed and unsigned data
descriptors, and a member whose CRC equals the descriptor signature
`0x08074b50`. The `-central-local-` and `-zip32-tail-` variants were derived
from that output by rewriting the central directory framing.

They are project-generated and carry no third-party content. The
`soapberry-zip` and `litchi-opc` ZIP64 tests compile them in. They moved here
from `docs/performance/results/change-0416/corpus/`.
