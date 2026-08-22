from __future__ import annotations

import unittest

from tools import check_iwork_leaf_public_api as public_api


def document(
    crate: str = "litchi_numbers",
    *,
    leak: str | None = None,
    leak_name: str | None = None,
    argument: str = "selector",
    semantic: bool = False,
) -> dict:
    """Build a small rustdoc document with one public document reader."""
    paths: dict[str, dict[str, object]] = {
        "0": {"path": [crate], "kind": "module"},
        "1": {"path": [crate, "document"], "kind": "module"},
        "2": {"path": [crate, "document", "Document"], "kind": "struct"},
        "3": {
            "path": [crate, "document", "Document", "open"],
            "kind": "function",
        },
        "4": {"path": [crate, "Document"], "kind": "struct"},
    }
    index: dict[str, dict[str, object]] = {
        "0": {
            "name": crate,
            "crate_id": 0,
            "inner": {"module": {"items": ["1", "4"]}},
        },
        "1": {
            "name": "document",
            "crate_id": 0,
            "inner": {"module": {"items": ["2", "3"]}},
        },
        "2": {
            "name": "Document",
            "crate_id": 0,
            "inner": {"struct": {"fields": [], "impls": ["3"]}},
        },
        "3": {
            "name": "open",
            "crate_id": 0,
            "inner": {
                "function": {
                    "sig": {
                        "inputs": [
                            [
                                argument,
                                {
                                    "resolved_path": {
                                        "id": "2",
                                    }
                                },
                            ]
                        ],
                        "output": {"resolved_path": {"id": "2"}},
                    }
                }
            },
        },
        "4": {
            "name": "Document",
            "crate_id": 0,
            "inner": {
                "use": {
                    "source": "crate::document::Document",
                    "name": "Document",
                    "id": "2",
                    "is_glob": False,
                }
            },
        },
    }
    if leak_name is not None:
        paths["5"] = {"path": [crate, "document", leak_name], "kind": "struct"}
        index["5"] = {
            "name": leak_name,
            "crate_id": 0,
            "inner": {"struct": {"fields": [], "impls": []}},
        }
        index["3"]["inner"]["function"]["sig"]["output"] = {
            "resolved_path": {"id": "5"}
        }
    if leak is not None:
        leak_parts = leak.split("::")
        paths["9"] = {"path": leak_parts, "kind": "struct"}
        index["9"] = {"name": leak_parts[-1], "crate_id": 9, "inner": {}}
        index["3"]["inner"]["function"]["sig"]["output"] = {
            "resolved_path": {"id": "9"}
        }
    if semantic:
        for item_id, name in (("5", "DocumentIdentifier"), ("6", "NodeId")):
            paths[item_id] = {
                "path": [crate, "document", name],
                "kind": "struct",
            }
            index[item_id] = {
                "name": name,
                "crate_id": 0,
                "inner": {"struct": {"fields": [], "impls": []}},
            }
        index["3"]["inner"]["function"]["sig"]["output"] = {
            "tuple": [
                {"resolved_path": {"id": "5"}},
                {"resolved_path": {"id": "6"}},
            ]
        }
    return {"root": 0, "index": index, "paths": paths}


def private_document(crate: str = "litchi_pages") -> dict:
    """Build rustdoc for a private document module re-exported at the root."""
    value = document(crate)
    value["index"]["0"]["inner"]["module"]["items"] = ["4"]
    value["paths"].pop("1")
    value["paths"].pop("4")
    value["index"].pop("1")
    return value


def keynote_selector_error_bridge() -> dict:
    """Build the generated selector-error bridge reached from a Keynote reader."""
    value = document(crate="litchi_keynote")
    value["paths"]["5"] = {
        "path": ["litchi_keynote", "selector", "SlideSelectorError"],
        "kind": "enum",
    }
    value["index"]["5"] = {
        "name": "SlideSelectorError",
        "crate_id": 0,
        "inner": {"enum": {"variants": [], "impls": ["7"]}},
    }
    value["paths"]["6"] = {
        "path": ["litchi_keynote", "package", "edit", "EditError"],
        "kind": "enum",
    }
    value["index"]["6"] = {
        "name": "EditError",
        "crate_id": 0,
        "inner": {"enum": {"variants": [], "impls": []}},
    }
    value["index"]["7"] = {
        "name": None,
        "crate_id": 0,
        "attrs": ["automatically_derived"],
        "inner": {
            "impl": {
                "trait": {
                    "path": "From",
                    "args": {
                        "angle_bracketed": {
                            "args": [
                                {
                                    "type": {
                                        "resolved_path": {"id": "5"},
                                    }
                                }
                            ]
                        }
                    },
                },
                "for": {"resolved_path": {"id": "6"}},
            }
        },
    }
    value["index"]["3"]["inner"]["function"]["sig"]["output"] = {
        "resolved_path": {"id": "5"}
    }
    return value


class IworkLeafPublicApiGateTests(unittest.TestCase):
    def test_commands_build_each_leaf_without_aggregate_features(self) -> None:
        commands = public_api.rustdoc_commands()
        self.assertEqual(len(commands), 3)
        self.assertEqual(
            commands[0],
            (
                "cargo",
                "rustdoc",
                "--package",
                "litchi-numbers",
                "--no-default-features",
                "--lib",
                "--",
                "-Zunstable-options",
                "--output-format",
                "json",
            ),
        )
        self.assertEqual(public_api.rustdoc_command("litchi-pages")[3], "litchi-pages")
        self.assertNotIn("iwork", " ".join(commands[0]))

    def test_accepts_semantic_identifier_names(self) -> None:
        self.assertEqual(
            public_api.violations(document(semantic=True), "litchi_numbers"), []
        )
        value = document(argument="semantic_identifier")
        value["index"]["3"]["name"] = "identifier"
        self.assertEqual(public_api.violations(value, "litchi_numbers"), [])

    def test_follows_scalar_ids_nested_in_public_tuple_fields(self) -> None:
        value = document()
        value["paths"]["5"] = {
            "path": ["litchi_numbers", "document", "WireView"],
            "kind": "struct",
        }
        value["index"]["5"] = {
            "name": "WireView",
            "crate_id": 0,
            "inner": {"struct": {"fields": [], "impls": []}},
        }
        value["paths"]["8"] = {
            "path": ["litchi_numbers", "document", "Document", "0"],
            "kind": "struct_field",
        }
        value["index"]["8"] = {
            "name": "0",
            "crate_id": 0,
            "inner": {"struct_field": {"tuple": ["5"]}},
        }
        value["index"]["2"]["inner"]["struct"]["fields"] = ["8"]

        failures = public_api.violations(value, "litchi_numbers")

        self.assertTrue(
            any("litchi_numbers::document::WireView" in failure for failure in failures),
            failures,
        )

    def test_all_forbidden_dependency_families_are_rejected(self) -> None:
        for crate_name in sorted(public_api.FORBIDDEN_CRATES):
            with self.subTest(crate_name=crate_name):
                failures = public_api.violations(
                    document(leak=f"{crate_name}::private::WireValue"),
                    "litchi_numbers",
                )
                self.assertEqual(len(failures), 1)
                self.assertIn(f"`{crate_name}::private::WireValue`", failures[0])

    def test_semantic_text_and_core_types_are_allowlisted(self) -> None:
        for path in (
            "litchi_iwa_text::storage::Storage",
            "litchi_core::Metadata",
            "serde::Serialize",
            "std::path::Path",
        ):
            with self.subTest(path=path):
                self.assertEqual(
                    public_api.violations(document(leak=path), "litchi_numbers"), []
                )

        failures = public_api.violations(
            document(leak="serde_json::Value"), "litchi_numbers"
        )
        self.assertEqual(len(failures), 1)
        self.assertIn("non-allowlisted crate", failures[0])

    def test_local_generated_wire_archive_and_raw_names_are_rejected(self) -> None:
        for name, expected in (
            ("GeneratedSettings", "generated type"),
            ("WireView", "implementation type"),
            ("ArchiveObject", "implementation type"),
            ("MessageId", "implementation type"),
        ):
            with self.subTest(name=name):
                failures = public_api.violations(
                    document(leak_name=name), "litchi_numbers"
                )
                self.assertTrue(any(expected in failure for failure in failures))

        failures = public_api.violations(
            document(argument="object_id"), "litchi_numbers"
        )
        self.assertEqual(len(failures), 1)
        self.assertIn("raw identifier", failures[0])

    def test_physical_package_paths_are_outside_the_document_graph(self) -> None:
        failures = public_api.violations(
            document(leak="litchi_numbers::package::Package"), "litchi_numbers"
        )
        self.assertEqual(len(failures), 1)
        self.assertIn("physical package type", failures[0])

    def test_missing_document_module_is_reported(self) -> None:
        value = document()
        value["paths"].pop("1")
        value["index"]["0"]["inner"]["module"]["items"].remove("4")
        self.assertEqual(
            public_api.violations(value, "litchi_numbers"),
            ["missing public module `litchi_numbers::document`"],
        )

    def test_private_pages_document_module_is_reached_through_root_reexport(
        self,
    ) -> None:
        self.assertEqual(
            public_api.violations(private_document(), "litchi_pages"), []
        )

    def test_private_pages_document_reexport_does_not_hide_package_leaks(self) -> None:
        value = private_document()
        value["paths"]["9"] = {
            "path": ["litchi_pages", "package", "Package"],
            "kind": "struct",
        }
        value["index"]["9"] = {
            "name": "Package",
            "crate_id": 0,
            "inner": {},
        }
        value["index"]["3"]["inner"]["function"]["sig"]["output"] = {
            "resolved_path": {"id": "9"}
        }
        failures = public_api.violations(value, "litchi_pages")
        self.assertEqual(len(failures), 1)
        self.assertIn("physical package type", failures[0])

    def test_public_keynote_document_module_remains_supported(self) -> None:
        self.assertEqual(
            public_api.violations(document(crate="litchi_keynote"), "litchi_keynote"),
            [],
        )

    def test_blanket_dependency_impls_are_ignored(self) -> None:
        value = document()
        value["index"]["2"]["inner"]["struct"]["impls"] = ["7"]
        value["index"]["7"] = {
            "crate_id": 9,
            "name": None,
            "inner": {"impl": {"blanket_impl": {"generic": "T"}}},
        }
        value["paths"]["70"] = {
            "path": ["prost", "Message"],
            "kind": "trait",
        }
        self.assertEqual(public_api.violations(value, "litchi_numbers"), [])

    def test_only_generated_keynote_selector_error_bridge_is_ignored(self) -> None:
        value = keynote_selector_error_bridge()
        self.assertEqual(public_api.violations(value, "litchi_keynote"), [])

        non_derived = keynote_selector_error_bridge()
        non_derived["index"]["7"]["attrs"] = []
        failures = public_api.violations(non_derived, "litchi_keynote")
        self.assertEqual(len(failures), 1)
        self.assertIn("physical package type", failures[0])

        direct_signature = keynote_selector_error_bridge()
        direct_signature["index"]["3"]["inner"]["function"]["sig"][
            "output"
        ] = {"resolved_path": {"id": "6"}}
        failures = public_api.violations(direct_signature, "litchi_keynote")
        self.assertEqual(len(failures), 1)
        self.assertIn("physical package type", failures[0])


if __name__ == "__main__":
    unittest.main()
