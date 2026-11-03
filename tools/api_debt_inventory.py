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
Two more shapes are not debt (record 0366): a `raw::Reader`, the mutable
codec model 0351 §2.1 rule 6 names, and a type that rustdoc places in a
build-script output file (the generated IWA wire schema keeps the file
format's message names, such as `tswp::UndoTransaction`).
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
FACADE_COMMIT_PATHS = gate.FACADE_COMMIT_PATHS
FACADE_COMMIT_METHODS = frozenset({
    "file", "into_file", "patch", "diagnostics", "value", "into_value", "into_parts",
})
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


class InventoryError(ValueError):
    """The rustdoc surface cannot prove a policy requirement."""


def policy_immutable_owners(crate: str, profiles=None) -> frozenset[tuple[str, ...]]:
    """Return policy-declared immutable owners for one rustdoc crate.

    A package can have more than one glossary profile (for example minimal and
    full feature sets), so the paths are deliberately de-duplicated here. The
    inventory's historical name heuristic remains the fast path; these policy
    paths close the gap for owners whose names are not in ``OWNERS``.
    """
    if profiles is None:
        profiles = gate.checked_policy(gate.read_json(gate.POLICY))
    return frozenset(
        tuple(owner.split("::"))
        for profile in profiles
        if profile["package"].replace("-", "_") == crate
        for owner in profile.get("immutable_owners", [])
    )


def policy_delegated_globs(crate: str, profiles=None) -> frozenset[str]:
    """Return all explicitly delegated foreign-glob crate names for ``crate``.

    A facade may have more than one checked profile, and an all-features
    rustdoc graph can contain a glob that is present only in one feature
    profile.  Use the union so the inventory does not accidentally depend on
    which profile happened to mention the glob.  ``Surface`` still refuses
    every foreign glob outside this explicit set; delegation only transfers
    the check to the delegated crate's own rustdoc inventory.
    """
    if profiles is None:
        profiles = gate.checked_policy(gate.read_json(gate.POLICY))
    return frozenset(
        delegated
        for profile in profiles
        if profile["package"].replace("-", "_") == crate
        for delegated in profile.get("delegated_globs", [])
    )


def missing_delegated_inputs(crates, profiles=None) -> dict[str, tuple[str, ...]]:
    """Return policy-delegated crates absent from a multi-crate input set.

    A facade's delegated glob is checked by the delegated crate's own
    rustdoc.  Requiring that rustdoc input alongside the facade prevents a
    partial command from presenting a facade as complete while silently
    omitting the delegated surface.
    """
    available = {str(crate).replace("-", "_") for crate in crates}
    if profiles is None:
        profiles = gate.checked_policy(gate.read_json(gate.POLICY))
    missing = {}
    for crate in sorted(available):
        delegated = policy_delegated_globs(crate, profiles)
        absent = tuple(sorted(delegated - available))
        if absent:
            missing[crate] = absent
    return missing


def policy_owner_paths(surface: Inventory, crate: str,
                       immutable_owners=None) -> frozenset[tuple[str, ...]]:
    """Resolve policy owners and their public aliases by nominal identity.

    ``Surface`` retains every public spelling while ``ReturnTypes`` resolves
    local aliases and reexports to rustdoc's nominal item identity. Matching
    identities makes an alias count as the same immutable owner without
    guessing from a spelling or emitting duplicate rows.
    """
    if immutable_owners is None:
        immutable_owners = policy_immutable_owners(crate)
    owner_paths = {
        tuple(path.split("::")) if isinstance(path, str) else tuple(path)
        for path in immutable_owners
    }
    return_types = gate.ReturnTypes(surface)
    identities = set()
    for path in sorted(owner_paths):
        kind = surface.exports.get(path)
        if kind not in gate.TYPE_KINDS:
            state = "missing" if kind is None else f"has non-nominal kind {kind!r}"
            raise InventoryError(
                f"{crate}: policy immutable owner {'::'.join(path)} {state}"
            )
        identity = return_types.nominal_identity(path)
        if identity is None:
            raise InventoryError(
                f"{crate}: policy immutable owner {'::'.join(path)} has a "
                "missing or ambiguous nominal identity"
            )
        identities.add(identity)
    for path, kind in surface.exports.items():
        if (kind in gate.TYPE_KINDS
                and return_types.nominal_identity(path) in identities):
            owner_paths.add(path)
    return frozenset(owner_paths)


class Inventory(gate.Surface):
    """The gate's surface, plus the shortest public path of every item."""

    def __init__(self, document, crate, delegated_globs=frozenset()):
        self.items: dict[str, tuple[str, ...]] = {}
        super().__init__(document, crate, delegated_globs)

    def walk(self, item_id, path, ancestors, *, exported):
        if item_id in self.index:
            current = self.items.get(item_id)
            if current is None or (len(path), path) < (len(current), current):
                self.items[item_id] = path
        return super().walk(item_id, path, ancestors, exported=exported)


def auditable_commit_nominals(document, crate: str, delegated_globs=frozenset()):
    """Return canonical Commit paths whose defining graph audits the triad."""
    surface = Inventory(document, crate, delegated_globs)
    methods: dict[tuple[str, ...], set[str]] = {}
    for path in surface.items.values():
        if len(path) > 1:
            methods.setdefault(path[:-1], set()).add(path[-1])
    covered = set()
    for path, kind in surface.exports.items():
        item_id = surface.item_ids.get(path)
        item = surface.index.get(str(item_id)) if item_id is not None else None
        definition = surface.paths.get(str(item_id))
        identity = (surface._nominal_identity(item_id)
                    if item_id is not None else None)
        if (kind not in gate.TYPE_KINDS or not path or path[0] != crate
                or not isinstance(item, dict)
                or item.get("crate_id") != surface.crate_id
                or not isinstance(definition, dict)
                or definition.get("crate_id") != surface.crate_id
                or identity is None or identity[0] != surface.crate_id):
            continue
        required = (FACADE_COMMIT_METHODS if path in FACADE_COMMIT_PATHS
                    else CORE_COMMIT_METHODS)
        if path[-1] in {"Commit", "BatchCommit"} and required <= methods.get(path, set()):
            covered.add(path)
    return frozenset(covered)


def is_generated(item) -> bool:
    """Whether rustdoc places ``item`` in a build-script output file.

    Generated wire schemas (``litchi-iwa-protos`` includes prost/buffa code
    from ``OUT_DIR``) keep the message names of the file format, such as
    ``tswp::UndoTransaction``; those names are not Litchi staging types.
    """
    span = item.get("span") if isinstance(item, dict) else None
    filename = span.get("filename", "") if isinstance(span, dict) else ""
    parts = Path(filename).parts
    return any(parts[i] == "build" and parts[i + 2] == "out" for i in range(len(parts) - 2))


def target_family(method: str) -> str:
    """The Edit family a package mutator moves to (a naming heuristic)."""
    if method.startswith(OUTPUT_PREFIXES):
        return "output: takes &self"
    if method in SIGNATURE:
        return "signature verb: &self -> Result<Package>"
    if method == "custom_props_mut":
        return "delete: metadata::Edit::{set_custom_props, remove_custom_props} (0362ac)"
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


def _rows(document, crate: str, result_rules=None, immutable_owners=None,
          delegated_globs=None, available_crates=None, covered_nominals=None):
    # A single-graph caller still represents an inventory with only its own
    # defining crate. Foreign Commit aliases therefore fail closed unless the
    # caller supplies an exact canonical coverage manifest. The core Commit is
    # handled as the explicit shared-contract exemption below.
    if available_crates is None:
        available_crates = frozenset({crate})
    if covered_nominals is None:
        covered_nominals = frozenset()
    profiles = None
    if result_rules is None or immutable_owners is None or delegated_globs is None:
        profiles = gate.checked_policy(gate.read_json(gate.POLICY))
    if delegated_globs is None:
        delegated_globs = policy_delegated_globs(crate, profiles)
    surface = Inventory(document, crate, delegated_globs)
    return_types = gate.ReturnTypes(surface)
    if immutable_owners is None:
        immutable_owners = policy_immutable_owners(crate, profiles)
    owners = policy_owner_paths(surface, crate, immutable_owners)
    if result_rules is None:
        result_rules = [rule for profile in profiles
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
        if (kind == "function" and len(path) > 1
                and (path[-2] in OWNERS or path[:-1] in owners)
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
            if (receiver == "&mut self"
                    and (owner in OWNERS or path[:-1] in owners) and not draft):
                if owner == "MutableDocument":
                    target = "primary Edit (opened) / Builder (fresh)"
                else:
                    target = target_family(name)
                yield "mut_owner", text, receiver, target
            elif receiver == "&mut self" and not draft and (
                    owner.startswith(("Mutable", "Flat")) or owner in DRAFT_NAMES) and not (
                    owner == "Reader" and "raw" in path[:-2]):
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
            generated = is_generated(item)
            if (kind == "enum" and name == "Error"
                    and not {"non_exhaustive", "#[non_exhaustive]"}.intersection(item.get("attrs", []))):
                yield "exhaustive_error", text, kind, "add #[non_exhaustive]; migrate exhaustive consumers"
            members = methods.get(path, set())
            if surface.foreign.get(path) == CORE_COMMIT:
                # An alias of the shared commit (record 0356 P4): rustdoc lists
                # its accessors in litchi-core, whose unit tests check them.
                members = members | CORE_COMMIT_METHODS
            if name.endswith("Commit"):
                # A foreign nominal type is audited in its defining crate. In
                # particular, ODF metadata::Commit is a local generic alias of
                # litchi_odf_common::metadata::Commit; inspecting the alias as
                # if it owned methods produces nine false debt rows.
                foreign = surface.foreign.get(path)
                if foreign is not None and foreign != CORE_COMMIT:
                    defining_crate = foreign[0]
                    covered = (covered_nominals is not None
                               and foreign in covered_nominals)
                    if (available_crates is not None
                            and (defining_crate not in available_crates or not covered)):
                        yield (
                            "commit_type", text,
                            f"foreign Commit from {'::'.join(foreign)}; defining crate "
                            "rustdoc target is not auditable in this inventory",
                            f"include {defining_crate} rustdoc with the exact Commit triad",
                        )
                    continue
                required = (FACADE_COMMIT_METHODS if path in FACADE_COMMIT_PATHS
                            else CORE_COMMIT_METHODS)
                missing = [m for m in sorted(required) if m not in members]
                extra = [m for m in ("snapshot", "into_snapshot", "workbook", "into_workbook")
                         if m in members]
                if missing or extra:
                    detail = ",".join(sorted(members & {"package", "into_package", "file",
                                                        "into_file", "patch", "diagnostics",
                                                        "value", "into_value", "into_parts",
                                                        *extra}))
                    target = "; ".join(filter(None, (
                        "add " + ",".join(missing) if missing else "",
                        "delete " + ",".join(extra) if extra else "")))
                    yield "commit_type", text, detail, target
            if name.endswith(("Edit", "Transaction")) and not generated:
                bad = sorted(members & {"rollback", "is_changed", "snapshot"})
                renamed = name.endswith("Transaction") or name != "Edit"
                if bad or renamed:
                    target = "; ".join(filter(None, (
                        "rename to the focused module's Edit" if renamed else "",
                        "delete " + ",".join(bad) if bad else "")))
                    yield "edit_type", text, ",".join(bad), target
            if name.endswith("Transaction") and not generated:
                yield "transaction_type", text, kind, "rename Edit (focused module)"
        elif kind == "variant" and name == "UnsafeEdit":
            yield "unsafe_edit", text, "variant", "Refused(litchi_core::edit::Refusal)"


def rows(document, crate: str, result_rules=None, immutable_owners=None,
         delegated_globs=None, available_crates=None, covered_nominals=None):
    """Yield unique debt rows while preserving distinct public paths.

    Rustdoc can expose the same nominal item through both a public spelling
    and an alias.  Those paths remain separate rows, but repeated emission of
    the exact same tuple is not a second debt item and must not make the
    inventory structurally ambiguous.
    """
    seen = set()
    for row in _rows(document, crate, result_rules, immutable_owners,
                     delegated_globs, available_crates, covered_nominals):
        if row not in seen:
            seen.add(row)
            yield row


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("json", nargs="+", type=Path, help="rustdoc JSON, one per crate")
    args = parser.parse_args(argv)
    print("crate\tcategory\tpath\tdetail\ttarget")
    try:
        profiles = gate.checked_policy(gate.read_json(gate.POLICY))
        # Parse each graph to obtain its root identity, retaining only those
        # names before the sequential row pass.  This prepass still parses
        # each full graph; the second pass remains sequential so the
        # inventory does not retain all 1+ GiB of all-features JSON.
        crates = []
        seen_crates = set()
        covered_nominals = set()
        for path in args.json:
            document = gate.read_json(path)
            root = str(document["root"])
            crate = document["index"][root]["name"]
            if crate in seen_crates:
                raise InventoryError(f"duplicate crate rustdoc input: {crate}")
            seen_crates.add(crate)
            crates.append(crate)
            covered_nominals.update(
                auditable_commit_nominals(
                    document, crate, policy_delegated_globs(crate, profiles)
                )
            )
        del document
        missing = missing_delegated_inputs(crates, profiles)
        if missing:
            details = "; ".join(
                f"{crate}: {', '.join(delegated)}"
                for crate, delegated in sorted(missing.items())
            )
            raise InventoryError(
                "delegated crate rustdoc input(s) missing; refusing partial inventory: "
                + details
            )
        for path in args.json:
            document = gate.read_json(path)
            # Captures may be renamed by profile or source revision. The crate
            # identity belongs to the rustdoc graph, not the artifact filename.
            crate = document["index"][str(document["root"])]["name"]
            for row in rows(
                    document, crate,
                    available_crates=frozenset(crates),
                    covered_nominals=frozenset(covered_nominals)):
                print(crate.replace("_", "-"), *row, sep="\t")
    except InventoryError as error:
        print(f"inventory error: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
