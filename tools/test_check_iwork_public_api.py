from __future__ import annotations

import unittest

from tools import check_iwork_public_api as public_api


def document(*, leak: str | None = None, argument: str = "selector") -> dict:
    paths = {
        "0": {"path": ["litchi"], "kind": "module"},
        "1": {"path": ["litchi", "iwork"], "kind": "module"},
        "2": {"path": ["litchi", "iwork", "Document"], "kind": "struct"},
        "3": {"path": ["litchi", "iwork", "Document", "open"], "kind": "function"},
    }
    signature_id = "2"
    if leak is not None:
        paths["9"] = {"path": leak.split("::"), "kind": "struct"}
        signature_id = "9"
    return {
        "root": 0,
        "index": {
            "0": {
                "name": "litchi",
                "inner": {"module": {"items": [1]}},
            },
            "1": {
                "name": "iwork",
                "inner": {"module": {"items": [2]}},
            },
            "2": {
                "name": "Document",
                "inner": {"struct": {"fields": [], "impls": [3]}},
            },
            "3": {
                "name": "open",
                "inner": {
                    "function": {
                        "sig": {
                            "inputs": [[argument, {"resolved_path": {"id": signature_id}}]],
                            "output": {"resolved_path": {"id": "2"}},
                        }
                    }
                },
            },
        },
        "paths": paths,
    }


class IworkPublicApiGateTests(unittest.TestCase):
    def test_command_builds_only_the_root_iwork_feature(self) -> None:
        self.assertEqual(
            public_api.rustdoc_command(),
            (
                "cargo",
                "rustdoc",
                "--package",
                "litchi",
                "--no-default-features",
                "--features",
                "iwork",
                "--",
                "-Zunstable-options",
                "--output-format",
                "json",
            ),
        )
        self.assertEqual(
            public_api.isolation_rustdoc_command(),
            (
                "cargo",
                "rustdoc",
                "--package",
                "litchi",
                "--no-default-features",
                "--features",
                "pages,keynote,numbers",
                "--",
                "-Zunstable-options",
                "--output-format",
                "json",
            ),
        )

    def test_accepts_facade_owned_semantic_surface(self) -> None:
        self.assertEqual(public_api.violations(document()), [])

    def test_rejects_each_forbidden_dependency_family(self) -> None:
        self.assertEqual(
            public_api.FORBIDDEN_CRATES,
            {
                "buffa",
                "litchi_iwa",
                "litchi_iwa_archive",
                "litchi_iwa_common",
                "litchi_iwa_core",
                "litchi_iwa_detect",
                "litchi_iwa_protos",
                "litchi_iwa_structured",
                "litchi_iwa_text",
                "litchi_iwa_text_wire",
                "litchi_keynote",
                "litchi_numbers",
                "litchi_numbers_wire",
                "litchi_pages",
                "prost",
                "prost_types",
            },
        )
        for crate_name in sorted(public_api.FORBIDDEN_CRATES):
            with self.subTest(crate_name=crate_name):
                failures = public_api.violations(
                    document(leak=f"{crate_name}::private::WireValue")
                )
                self.assertEqual(len(failures), 1)
                self.assertIn(f"`{crate_name}::private::WireValue`", failures[0])

    def test_rejects_raw_id_argument_and_type_names(self) -> None:
        for argument in ("id", "identifier", "native_id", "object_id"):
            with self.subTest(argument=argument):
                argument_failure = public_api.violations(document(argument=argument))
                self.assertEqual(len(argument_failure), 1)
                self.assertIn(f"raw identifier as `{argument}`", argument_failure[0])

        raw_type = document()
        raw_type["index"]["2"]["name"] = "MessageId"
        type_failure = public_api.violations(raw_type)
        self.assertEqual(len(type_failure), 1)
        self.assertIn("raw identifier as `MessageId`", type_failure[0])

        referenced_type = document(leak="semantic_common::NodeId")
        path_failure = public_api.violations(referenced_type)
        self.assertEqual(len(path_failure), 2)
        self.assertTrue(
            any(
                "raw identifier type `semantic_common::NodeId`" in failure
                for failure in path_failure
            )
        )

        public_field = document()
        public_field["index"]["2"]["inner"]["struct"]["fields"] = ["4"]
        public_field["index"]["4"] = {
            "name": "object_id",
            "inner": {"struct_field": {"primitive": "u64"}},
        }
        field_failure = public_api.violations(public_field)
        self.assertEqual(len(field_failure), 1)
        self.assertIn("raw identifier as `object_id`", field_failure[0])

    def test_rejects_physical_capability_names(self) -> None:
        names = (
            "raw",
            "raw_message",
            "archive",
            "source_catalog",
            "components",
            "prepared",
            "PreparedSource",
            "source_bytes",
        )
        for name in names:
            with self.subTest(name=name):
                failures = public_api.violations(document(argument=name))
                self.assertEqual(len(failures), 1)
                self.assertIn(f"as `{name}`", failures[0])

    def test_external_types_are_closed_to_a_standard_library_allowlist(self) -> None:
        self.assertEqual(public_api.ALLOWED_EXTERNAL_CRATES, {"alloc", "core", "std"})
        for crate_name in sorted(public_api.ALLOWED_EXTERNAL_CRATES):
            with self.subTest(crate_name=crate_name):
                self.assertEqual(
                    public_api.violations(document(leak=f"{crate_name}::path::Path")),
                    [],
                )

        failures = public_api.violations(document(leak="serde::private::Value"))
        self.assertEqual(len(failures), 1)
        self.assertIn("non-allowlisted crate `serde::private::Value`", failures[0])

        local_failure = public_api.violations(document(leak="litchi::pages::Package"))
        self.assertEqual(len(local_failure), 1)
        self.assertIn(
            "type outside the iWork facade `litchi::pages::Package`",
            local_failure[0],
        )

    def test_leaf_features_must_not_publish_the_aggregate_module(self) -> None:
        isolated = document()
        isolated["paths"].pop("1")
        self.assertEqual(public_api.isolation_violations(isolated), [])
        self.assertEqual(
            public_api.isolation_violations(document()),
            ["public module `litchi::iwork` is available without the `iwork` feature"],
        )

    def test_requires_the_public_iwork_namespace(self) -> None:
        missing = document()
        missing["paths"].pop("1")
        self.assertEqual(
            public_api.violations(missing),
            ["missing public module `litchi::iwork`"],
        )

    def test_ignores_dependency_blankets_but_checks_explicit_local_impls(self) -> None:
        value = document()
        value["index"]["0"]["crate_id"] = 0
        value["index"]["2"]["inner"]["struct"]["impls"] = ["7", "8"]
        value["index"]["7"] = {
            "crate_id": 0,
            "name": None,
            "inner": {
                "impl": {
                    "blanket_impl": {"generic": "T"},
                    "trait": {"resolved_path": {"id": "70"}},
                    "for": {"resolved_path": {"id": "2"}},
                    "items": [],
                }
            },
        }
        value["index"]["8"] = {
            "crate_id": 0,
            "name": None,
            "inner": {
                "impl": {
                    "blanket_impl": None,
                    "trait": {"resolved_path": {"id": "80"}},
                    "for": {"resolved_path": {"id": "2"}},
                    "items": [],
                }
            },
        }
        value["paths"]["70"] = {
            "path": ["unrelated_dependency", "Blanket"],
            "kind": "trait",
        }
        value["paths"]["80"] = {
            "path": ["litchi_iwa_structured", "Leak"],
            "kind": "trait",
        }

        failures = public_api.violations(value)
        self.assertEqual(len(failures), 1)
        self.assertIn("litchi_iwa_structured::Leak", failures[0])


if __name__ == "__main__":
    unittest.main()
