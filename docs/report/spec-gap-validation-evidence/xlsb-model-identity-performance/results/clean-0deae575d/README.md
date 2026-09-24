# Clean 0deae575d XLSB identity capture

This directory is an exact-byte retained copy of the approved external run
from `/var/tmp/litchi-xlsb-identity-results.UwOSJT`. The source checkout was
`/var/tmp/litchi-xlsb-identity-clean.vRFKKx` at commit
`0deae575dbecd3d5b351eadd51b87c39a6dc11df`; the runner used Cargo 1.95.0,
fresh source manifests, and a fresh private target.

The committed runner completed 11 synthetic smoke lanes and the full
correctness matrix: 120 coordinates, 110 successful points, and 10 typed
`graph work` limit refusals at `T=64,R=256`. The matrix receipt SHA-256 is
`f427ddd13a8e928faed90c3e6b07ff2158dfb526397b189762d1926c749a798c`.
The binary SHA-256 was identical before and after execution:
`4648585a5f9f45f454d879e3cbdf5fe69de6c727755dd407f3b23ee4407b62cb`.
The source manifest SHA-256 was identical before and after:
`a7c4e779ebc3ffe4562820ceb8dd07236c5e248ede5fca53d00316e7b92ebbf4`.

This is correctness and provenance evidence only. It makes no performance,
allocator/RSS, or native XLSB acceptance claim. The separate preflight archive
at `/var/tmp/litchi-xlsb-identity-preflight.o5vZsp` used checkout-local default
results and is not part of this retained capture. Root verification independently passed the complete verifier on this retained
copy and compared all 46 original files byte-for-byte with the external capture.
`root-verification.json` records those file hashes and results. The older
`clean-2a726ded6` receipt remains separate and is not reused here.
