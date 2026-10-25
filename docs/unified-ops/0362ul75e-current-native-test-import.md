# 0362ul75e — standalone protocol test import

The coordinator's first standalone invocation of
`python3 -m unittest tools.test_current_native_policy_v2` fails at import with
`ModuleNotFoundError: No module named 'native_qualification'`. The test relied
on a caller-provided Python path or an earlier test module adding `tools` to
the import path. Its new gate registration must also work independently.

The test now derives the repository root from its own file and inserts the
tracked `tools` directory before importing the three policy modules, matching
the other qualification suites. Policy, trusted bindings, production source,
runner and validation behavior are unchanged. This corrects the prematurely
recorded composed-pass sentence in 0362ul75d; the initial import failure is
preserved here rather than treated as an executed test pass.

After the repair, the standalone current-native suite and the signed-native,
DrawingML-v2 and gate-wiring suites pass all 47 tests without `PYTHONPATH`.
The composed runner has 39 tooling modules and all 53 steps. Relative to
0362ul76, only the three additional test-module names change its argv. Its
52,787 bytes have SHA-256
`e88893d497f7e7ef70bde4314c3e577d47f20997f2ea9d82b4d4595edbbe8458`.
The synchronized external gate is 6,249 bytes with SHA-256
`23170aedf5583e637b48eacb79f7976eef65c83e3fbe9f3c3749ddd807e1671c`.
