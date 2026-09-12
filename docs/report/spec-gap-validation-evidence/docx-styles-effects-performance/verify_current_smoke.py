#!/usr/bin/env python3
"""Verify a fresh 52-lane smoke against the current 8702 source pin.

The historical verifier remains bound to d1f299d00 and is not modified.
"""

import verify

CURRENT_SOURCE_COMMIT = "8702fd4db8723acceb7deb51bcb40ff66604bf10"
verify.SOURCE_COMMIT = CURRENT_SOURCE_COMMIT

if __name__ == "__main__":
    verify.main()
