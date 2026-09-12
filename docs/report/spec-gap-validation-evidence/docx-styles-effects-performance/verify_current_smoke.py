#!/usr/bin/env python3
"""Verify a fresh 52-lane smoke against the current 8702 source pin.

The historical verifier remains bound to d1f299d00 and is not modified.
"""

import verify

CURRENT_SOURCE_COMMIT = "d000d977b99e03f8542c7dae74acf767a91b1feb"
verify.SOURCE_COMMIT = CURRENT_SOURCE_COMMIT

if __name__ == "__main__":
    verify.main()
