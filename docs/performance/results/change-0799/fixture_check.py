"""Compare the binary-generated catalog with independently frozen literal inputs."""

import hashlib

import custody as c


def check():
    expected = c.read(c.P / "fixtures.json")
    actual = c.read(c.P / "cases.json")
    assert len(expected) == len(actual) == 39
    expected_ids = [entry["id"] for entry in expected]
    actual_ids = [entry["id"] for entry in actual]
    assert len(set(expected_ids)) == len(expected_ids) == 39
    assert actual_ids == expected_ids
    for expected_entry, actual_entry in zip(expected, actual):
        assert all(
            expected_entry[key] == actual_entry[key]
            for key in ["id", "category", "attribute_count"]
        )
        source = actual_entry["source"]
        raw = (
            source["value"].encode()
            if source["encoding"] == "utf8"
            else bytes.fromhex(source["value"])
        )
        assert raw == expected_entry["input"].encode()
        assert len(raw) == source["bytes"] == expected_entry["bytes"]
        assert hashlib.sha256(raw).hexdigest() == expected_entry["sha256"]


if __name__ == "__main__":
    check()
