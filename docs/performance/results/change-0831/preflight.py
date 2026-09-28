"""Run the existing allocation-reader regression suite without workloads."""
import json
import sys
import unittest

import driver as d

sys.path.insert(0, str(d.ROOT))
from tools import test_perf_allocation_schema


def main():
    suite = unittest.defaultTestLoader.loadTestsFromModule(test_perf_allocation_schema)
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    assert result.wasSuccessful()
    print(json.dumps({"status": "pass", "tests": result.testsRun,
        "schema_source_sha256": d.sha(d.ROOT / "tools/perf_allocation_schema.py"),
        "test_source_sha256": d.sha(d.ROOT / "tools/test_perf_allocation_schema.py")}, sort_keys=True))


if __name__ == "__main__":
    main()
