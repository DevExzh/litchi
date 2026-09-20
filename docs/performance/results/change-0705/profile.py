"""Capture current-head XLSX publication-owner Callgrind profiles."""

import json

from capture import BINARY, HERE, capture


def main() -> None:
    profile = json.loads((HERE / "profile-plan.json").read_text())
    native = json.loads((HERE / "plan.json").read_text())
    assert profile["shapes"] == native["shapes"]
    assert profile["repeats"] == native["native_repeats"] == 2
    assert profile["warmup"] == 0 and profile["samples"] == 1
    owner = profile["owner"]
    for repeat in range(1, profile["repeats"] + 1):
        # Reverse the second pass to avoid coupling shape with acquisition order.
        shapes = profile["shapes"] if repeat == 1 else list(reversed(profile["shapes"]))
        for shape in shapes:
            name = f"profile-r{repeat}-{shape}"
            command = [
                "taskset",
                "-c",
                str(native["cpu"]),
                "valgrind",
                *profile["options"],
                "--callgrind-out-file=" + str(HERE / (name + ".callgrind")),
                str(BINARY),
                "--warmup",
                str(profile["warmup"]),
                "--samples",
                str(profile["samples"]),
                "--case",
                native["case"],
                "--xlsx-cell-crud-shape",
                shape,
                "--json",
                str(HERE / (name + ".json")),
            ]
            assert f"--toggle-collect={owner}" in command
            capture(name, command)


if __name__ == "__main__":
    main()
