#!/usr/bin/env python3
"""Pinned source arms for matched InkAction profile captures."""

from __future__ import annotations

import argparse


APPROVED_BASE_COMMIT = "079cbbcbfc00c8d2412a38586a9688ae7eb0009e"
ARMS = ("baseline", "candidate")

# The candidate pin is an isolated child of the clean baseline pin with only
# c8d60d563's reviewed 28-line InkAction source change cherry-picked.  Guard
# commits may be applied on top; the runner separately checks this pin is an
# ancestor and checks these exact production source hashes.
SOURCE_PINS = {
    "baseline": "f1cb119361af9ea2227d27050e41915a9a92ae04",
    "candidate": "ab94a4d7a02053765bf4c70b4af5022273b821e5",
}

SOURCE_HASHES = {
    "baseline": {
        "crates/litchi-drawingml/src/ink/mod.rs": "a9a55cf0c44b59afa00a7b7c472c0f010eef5c23e62a816d4bb0f27aa6f67ff7",
        "crates/litchi-drawingml/src/ink/actions.rs": "4e4731d59ab95679f205d424567dcf4d8e02f27c9318ed8ac955a8c163910628",
        "crates/litchi-drawingml/src/ink/actions_edit.rs": "b037eb7fae01b2e62c3c3a1050d9044466ad8d8e1a1402bb13533b2ba8bf9850",
        "crates/litchi-drawingml/tests/ink_action_edit.rs": "af92fe9d2923ac1197e4f152a5e34350217b5c782fa00365b82c141f69787640",
        "crates/litchi-drawingml/tests/ink_action_id_boundaries.rs": "f07b423989443d68ccba070aef5fed610bc39acb2c8b5a6adf14de8074afbf06",
    },
    "candidate": {
        "crates/litchi-drawingml/src/ink/mod.rs": "a9a55cf0c44b59afa00a7b7c472c0f010eef5c23e62a816d4bb0f27aa6f67ff7",
        "crates/litchi-drawingml/src/ink/actions.rs": "4d15cfd25456115bb750622096573118c7f1a652bc499be093db1fd93863a724",
        "crates/litchi-drawingml/src/ink/actions_edit.rs": "8262d2701e497b85c49024640bcaf4d8aaa1b3f856fa2d0956efa2a6f80534bb",
        "crates/litchi-drawingml/tests/ink_action_edit.rs": "dc565061ccdce79f8eac369ed092f37b608763bd35b91a8079089fb2b8bd2404",
        "crates/litchi-drawingml/tests/ink_action_id_boundaries.rs": "f07b423989443d68ccba070aef5fed610bc39acb2c8b5a6adf14de8074afbf06",
    },
}


def profile(arm: str) -> tuple[str, dict[str, str]]:
    if arm not in ARMS:
        raise ValueError(f"unknown profile arm: {arm}")
    return SOURCE_PINS[arm], dict(SOURCE_HASHES[arm])


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--arm", choices=ARMS, required=True)
    args = parser.parse_args()
    pin, hashes = profile(args.arm)
    print(f"base\t{APPROVED_BASE_COMMIT}")
    print(f"pin\t{pin}")
    for relative, digest in sorted(hashes.items()):
        print(f"source\t{relative}\t{digest}")


if __name__ == "__main__":
    main()
