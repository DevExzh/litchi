"""``python/litchi_py/__init__.pyi`` against the runtime module.

The goal asks for complete stubs "kept in sync by a test". These tests compare
the stub with ``dir()``/``inspect`` of the extension: public names, ``__all__``,
member kinds (method, staticmethod, property, enum attribute) and parameter
names, so keyword calls such as ``Document.open(path=...)`` stay valid.

They tolerate exactly the known drift recorded in ``_support`` (D2, D3, D5)
and fail on any new gap, in either direction. Adding API means adding it to
the runtime and the stub together. The today/intended pairs for the known
drift live in test_known_drift.py.

Stub classes whose name starts with ``_`` are typing helpers that exist only
in the stub (the TypedDicts of the dicts the unified ``File`` API returns);
``test_typing_only_declarations`` checks that they stay private and
stub-only, and test_file_api.py checks their keys against real results.
"""

from __future__ import annotations

import unittest

from _support import (
    KNOWN_README_ONLY_API,
    KNOWN_RUNTIME_ONLY_MEMBERS,
    KNOWN_STUB_BASE_DRIFT,
    L,
    parse_stub,
    public_members,
    readme_api_references,
    resolve_stub_base,
    runtime_has,
    runtime_member_kind,
    runtime_params,
)

STUB = parse_stub()


class StubSyncTests(unittest.TestCase):
    def test_module_level_names(self):
        runtime = {name for name in dir(L) if not name.startswith("_")}
        declared = set(STUB.classes) | set(STUB.functions)
        self.assertEqual(runtime - declared, set(), "runtime names missing from the stub")
        self.assertEqual(declared - runtime, set(), "stub names missing from the runtime")

    def test_dunder_all(self):
        self.assertEqual(sorted(L.__all__), sorted(STUB.all))
        self.assertEqual(set(STUB.all), set(STUB.classes) | set(STUB.functions))

    def test_class_members_modulo_known_drift(self):
        for name, stub_class in STUB.classes.items():
            runtime = public_members(getattr(L, name))
            declared = {member for member in stub_class.members if not member.startswith("_")}
            with self.subTest(cls=name):
                self.assertEqual(declared - runtime, set(), "stub declares members the runtime lacks")
                unexpected = runtime - declared - set(KNOWN_RUNTIME_ONLY_MEMBERS.get(name, ()))
                self.assertEqual(unexpected, set(), "runtime members missing from the stub")

    def test_member_kinds(self):
        for name, stub_class in STUB.classes.items():
            runtime_class = getattr(L, name)
            for member in stub_class.members.values():
                if member.kind == "constructor" or not hasattr(runtime_class, member.name):
                    continue
                with self.subTest(member=f"{name}.{member.name}"):
                    self.assertEqual(runtime_member_kind(runtime_class, member.name), member.kind)

    def test_parameter_names(self):
        for name, params in STUB.functions.items():
            with self.subTest(function=name):
                self.assertEqual(runtime_params(getattr(L, name), drop_self=False), params)
        for name, stub_class in STUB.classes.items():
            runtime_class = getattr(L, name)
            for member in stub_class.members.values():
                with self.subTest(member=f"{name}.{member.name}"):
                    if member.kind == "constructor":
                        actual = runtime_params(runtime_class, drop_self=False)
                    elif member.kind == "staticmethod":
                        actual = runtime_params(getattr(runtime_class, member.name), drop_self=False)
                    elif member.kind == "method":
                        actual = runtime_params(getattr(runtime_class, member.name), drop_self=True)
                    else:
                        continue
                    self.assertEqual(actual, member.params)

    def test_typing_only_declarations(self):
        self.assertTrue(STUB.typing_only, "the stub declares the shapes of returned dicts")
        for name, stub_class in STUB.typing_only.items():
            with self.subTest(cls=name):
                self.assertTrue(name.startswith("_"))
                self.assertIn("TypedDict", stub_class.bases)
                self.assertFalse(hasattr(L, name), "a typing-only name exists at runtime")
                self.assertNotIn(name, STUB.all)

    def test_declared_bases_modulo_known_drift(self):
        for name, stub_class in STUB.classes.items():
            for base in stub_class.bases:
                if (name, base) in KNOWN_STUB_BASE_DRIFT:
                    continue
                with self.subTest(cls=name, base=base):
                    self.assertTrue(issubclass(getattr(L, name), resolve_stub_base(base)))


class ReadmeSyncTests(unittest.TestCase):
    def test_readme_api_reference_modulo_known_drift(self):
        references = readme_api_references()
        self.assertTrue(references, "README API Reference section not found")
        missing = {ref for ref in references if not runtime_has(*ref)}
        self.assertEqual(missing - KNOWN_README_ONLY_API, set(), "README documents missing API")


if __name__ == "__main__":
    unittest.main()
