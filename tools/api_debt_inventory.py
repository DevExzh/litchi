#!/usr/bin/env python3
"""List the record 0351 API debt (owner decision 0350 items 14-19) per crate.

Reads all-features rustdoc JSON (one file per crate) and walks it with the
glossary gate's own Surface class, so it sees exactly the public surface the
gate sees. Each row names one public item that a migration wave must change:

- mut_owner: a `&mut self` method on a Package-like owner (item 14);
- mut_draft: a `&mut self` method on a Mutable*/Flat*/Reader/projection type;
- commit_type: a commit type missing a triad accessor or still publishing a
  snapshot/workbook (item 15); edit_type: a staging type still named
  *Transaction or exposing rollback/is_changed/snapshot (items 15-16);
- commit_snapshot, into_workbook, is_changed, rollback, transaction_type;
- edit_output: artifact output on an Edit instead of its committed Package;
- commit_return: Edit::commit does not return Result of its family's Commit;
- replay_return: an owner replay has no verified Result<Commit> return rule;
- unsafe_edit: an `UnsafeEdit` error variant (item 17);
- exhaustive_error: a public Error enum missing #[non_exhaustive] (item 17);
- with_limits: a public `with_limits` or `*_with_limits` entry point (items 18-19).

Build the input with, per crate (sequentially; RUSTC_BOOTSTRAP=1):
  cargo rustdoc -p <crate> --lib --all-features -- -Zunstable-options \
      --output-format json --document-hidden-items
then run: tools/api_debt_inventory.py target/doc/litchi_docx.json ... > out.tsv
A wave is done for a category when its crate has no rows left in it. A
`Commit` that aliases `litchi_core::edit::Commit` has the triad accessors
(record 0356 P4), and a type under a `builder` module is an authoring draft,
whose `&mut self` methods are allowed (0351 §2.1 rule 7; record 0356b).
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import check_api_glossary as gate  # noqa: E402

OWNERS = {"Package", "Workbook", "Document", "Presentation", "Drawing", "Chart",
          "Spreadsheet", "Image", "Database", "Formula", "FlatDrawing",
          "FlatPresentation", "FlatDocument", "FlatSpreadsheet", "MutableDocument", "File"}
CORE_COMMIT = ("litchi_core", "edit", "Commit")
CORE_COMMIT_METHODS = frozenset({"package", "into_package", "patch", "diagnostics", "into_parts"})
DRAFT_NAMES = {"Reader", "Workbook", "Spreadsheet", "Presentation", "Document"}
OUTPUT_PREFIXES = ("save", "to_", "write_to")
SIGNATURE = {"sign", "sign_with", "resign", "resign_with", "unsign"}
VERBS = ("apply_", "put_", "remove_", "set_", "clear_", "add_", "replace_", "insert_",
         "move_", "order_", "rename_", "update_", "edit_", "store_", "take_",
         "publish_", "undo_", "redo_")
SUFFIXES = ("_with_and_limits", "_with_limits", "_with_policy", "_slot_patch", "_patch",
            "_commit", "_plan", "_with", "_by_name", "_owner")
ALIASES = {
    "durable_document": "document", "durable_ink": "ink", "durable_task_panes": "task_panes",
    "ink_actions": "ink", "ink_edit": "ink", "conflicts": "conflict",
    "content_controls": "content_control", "tracked_revisions": "tracked_revision",
    "readonly_recommended": "read_only_recommended", "glossary_graph": "glossary",
    "document_variables": "document_variable", "merged_cell_ranges": "merged_cell_range",
    "mail_merge_recipients": "mail_merge", "alt": "alt_chunk",
    "bibliography_source": "bibliography", "font_embedding": "fonts",
    "slide_layout": "master_layout", "slide_master": "master_layout",
    "story_hyperlink_texts": "story_hyperlinks", "attached_template_uri": "attached_template",
}


class Inventory(gate.Surface):
    """The gate's surface, plus the shortest public path of every item."""

    def __init__(self, document, crate):
        self.items: dict[str, tuple[str, ...]] = {}
        super().__init__(document, crate)

    def walk(self, item_id, path, ancestors, *, exported):
        if item_id in self.index:
            current = self.items.get(item_id)
            if current is None or (len(path), path) < (len(current), current):
                self.items[item_id] = path
        return super().walk(item_id, path, ancestors, exported=exported)


def target_family(method: str) -> str:
    """The Edit family a package mutator moves to (a naming heuristic)."""
    if method.startswith(OUTPUT_PREFIXES):
        return "output: takes &self"
    if method in SIGNATURE:
        return "signature verb: &self -> Result<Package>"
    if method == "custom_props_mut":
        return "delete: metadata::Edit::set_custom_props (0342)"
    if method == "edit_opc":
        return "raw escape: edit_opc(&self) -> opc::Edit committing a Package"
    if method in ("undo_document", "redo_document", "publish_document_commit_with_history"):
        return "history: &self -> Result<Commit>"
    name = method
    for verb in VERBS:
        if name.startswith(verb):
            name = name[len(verb):]
            break
    for _ in range(2):
        for suffix in SUFFIXES:
            if name.endswith(suffix):
                name = name[: -len(suffix)]
    name = ALIASES.get(name, name)
    return f"{name}::Edit via Package::edit_{name}() / apply_{name}(&Patch)"


def rows(document, crate: str, result_rules=None):
    surface = Inventory(document, crate)
    return_types = gate.ReturnTypes(surface)
    if result_rules is None:
        result_rules = [rule for profile in gate.checked_policy(gate.read_json(gate.POLICY))
                        if profile["package"].replace("-", "_") == crate
                        for rule in profile.get("result_methods", [])]
    verified_results = set()
    for rule in result_rules:
        path = tuple(rule["method"].split("::"))
        args = [return_types.policy_type(arg, path)
                for arg in rule.get("type_args", [])]
        if (return_types.is_commit(tuple(rule["ok"].split("::")))
                and not any(arg is None for arg in args)
                and return_types.matches_result(path, tuple(rule["ok"].split("::")), args)
                and return_types.matches_owner(path, tuple(rule["ok"].split("::")), args)):
            verified_results.add(path)
    methods: dict[tuple[str, ...], set[str]] = {}
    for path in surface.items.values():
        if len(path) > 1:
            methods.setdefault(path[:-1], set()).add(path[-1])
    replay_paths = {}
    for path, kind in surface.exports.items():
        if (kind == "function" and len(path) > 1 and path[-2] in OWNERS
                and (path[-1] == "apply" or path[-1].startswith("apply_"))
                and "builder" not in path[:-1]):
            item_id = surface.item_ids[path]
            previous = replay_paths.get(item_id)
            if previous is None or (len(path), path) < (len(previous), previous):
                replay_paths[item_id] = path
    for item_id, path in sorted(surface.items.items(), key=lambda item: item[1]):
        item = surface.index[item_id]
        kind = next(iter(item.get("inner") or {"none": None}))
        name, text = path[-1], "::".join(path)
        owner = path[-2] if len(path) > 1 else ""
        if kind == "function":
            receiver = gate.self_receiver(item["inner"]["function"]) or "fn"
            draft = "builder" in path[:-1]
            if receiver == "&mut self" and owner in OWNERS and not draft:
                if owner == "MutableDocument":
                    target = "primary Edit (opened) / Builder (fresh)"
                else:
                    target = target_family(name)
                yield "mut_owner", text, receiver, target
            elif receiver == "&mut self" and crate != "litchi_iwa" and not draft and (
                    owner.startswith(("Mutable", "Flat")) or owner in DRAFT_NAMES):
                yield "mut_draft", text, receiver, "Builder draft or Edit verb; no Mutable* type"
            if name == "with_limits" or name.endswith("_with_limits"):
                yield "with_limits", text, receiver, "OpenOptions / EditOptions, or required limits (raw)"
            if name == "rollback":
                yield "rollback", text, receiver, "delete: drop the Edit; Edit::reset in place"
            if name in ("snapshot", "into_snapshot") and "Commit" in owner:
                yield "commit_snapshot", text, receiver, "delete: Commit::package/into_package"
            if name == "into_workbook":
                yield "into_workbook", text, receiver, "delete: Package owns; Workbook is read-only"
            if name == "is_changed":
                yield "is_changed", text, receiver, "delete: diagnostics().changed() / patch.is_noop()"
            if (not draft and owner.endswith(("Edit", "Transaction", "Editor"))
                    and gate.is_edit_output(name)):
                yield "edit_output", text, receiver, "commit().package() owns artifact output"
            if not draft and owner == "Edit" and name == "commit":
                expected = path[:-2] + ("Commit",)
                if (surface.exports.get(expected) in gate.TYPE_KINDS
                        and (not return_types.is_commit(expected)
                             or not return_types.matches_result(path, expected))):
                    yield "commit_return", text, receiver, "Result<" + "::".join(expected) + ">"
            replay_path = replay_paths.get(item_id)
            if replay_path is not None and replay_path not in verified_results:
                payload = return_types.result_payload(return_types.method_output(replay_path))
                detail = ("returns owner" if payload is not None
                          and payload == return_types.type_at(replay_path[:-1])
                          else "return contract unverified")
                yield "replay_return", "::".join(replay_path), detail, "register and satisfy Result<family::Commit> in result_methods"
        elif kind in ("struct", "enum", "type_alias"):
            if (kind == "enum" and name == "Error"
                    and not {"non_exhaustive", "#[non_exhaustive]"}.intersection(item.get("attrs", []))):
                yield "exhaustive_error", text, kind, "add #[non_exhaustive]; migrate exhaustive consumers"
            members = methods.get(path, set())
            if surface.foreign.get(path) == CORE_COMMIT:
                # An alias of the shared commit (record 0356 P4): rustdoc lists
                # its accessors in litchi-core, whose unit tests check them.
                members = members | CORE_COMMIT_METHODS
            if name.endswith("Commit"):
                missing = [m for m in ("package", "into_package", "patch", "diagnostics", "into_parts")
                           if m not in members]
                extra = [m for m in ("snapshot", "into_snapshot", "workbook", "into_workbook")
                         if m in members]
                if missing or extra:
                    detail = ",".join(sorted(members & {"package", "into_package", "patch",
                                                        "diagnostics", "into_parts", *extra}))
                    target = "; ".join(filter(None, (
                        "add " + ",".join(missing) if missing else "",
                        "delete " + ",".join(extra) if extra else "")))
                    yield "commit_type", text, detail, target
            if name.endswith(("Edit", "Transaction")):
                bad = sorted(members & {"rollback", "is_changed", "snapshot"})
                renamed = name.endswith("Transaction") or name != "Edit"
                if bad or renamed:
                    target = "; ".join(filter(None, (
                        "rename to the focused module's Edit" if renamed else "",
                        "delete " + ",".join(bad) if bad else "")))
                    yield "edit_type", text, ",".join(bad), target
            if name.endswith("Transaction"):
                yield "transaction_type", text, kind, "rename Edit (focused module)"
        elif kind == "variant" and name == "UnsafeEdit":
            yield "unsafe_edit", text, "variant", "Refused(litchi_core::edit::Refusal)"


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("json", nargs="+", type=Path, help="rustdoc JSON, one per crate")
    args = parser.parse_args(argv)
    print("crate\tcategory\tpath\tdetail\ttarget")
    for path in args.json:
        document = gate.read_json(path)
        # Captures may be renamed by profile or source revision. The crate
        # identity belongs to the rustdoc graph, not the artifact filename.
        crate = document["index"][str(document["root"])]["name"]
        for row in rows(document, crate):
            print(crate.replace("_", "-"), *row, sep="\t")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
