#!/usr/bin/env python3
"""Re-pin the checked DOC fresh-writer corpus identity after the 0759 merge.

usage: repin_doc_fresh_writer.py REPO PREFLIGHT_REPORT PREFLIGHT_CATALOG [--write]

The merged Dop2002 conformance fix (`581711a4c4`) makes every fresh DOC
write the 594-byte DOP that nFibNew 0x0101 requires, so the three
`doc_fresh_write_to` corpora (`doc-tiny`, `doc-large`, `doc-payload-heavy`)
have new bytes. This follows 0508's `promote.py`:

* The identity comes from a one-sample, zero-warmup preflight of the default
  matrix, run from a clean worktree.
* Every other corpus, the case-to-corpus mapping and the identity
  configuration must be unchanged. Otherwise nothing is written.
* The Python catalog generator must reproduce the Rust catalog the harness
  emitted, including its build revision and dirty flag.
* Derived pins follow: the V1 `result_keys_sha256`, the policy's expected
  keys (with the `policy_id` bumped, as each earlier key change did), and
  both coverage indexes' checked-catalog hashes and DOC corpus IDs.

Each rewritten file must first round-trip through the same serializer, so
every diff is semantic. Without `--write` it only verifies and reports.
"""
import copy
import json
import sys
from pathlib import Path

DOC_CORPORA = ("doc-tiny", "doc-large", "doc-payload-heavy")
V1 = "docs/performance/results/perf-regression-default-manifest-v1.json"
CATALOG = "docs/performance/results/perf-corpus-manifest-v2.json"
POLICY = "docs/performance/perf-regression-policy-v1.json"
INDEXES = (
    "docs/performance/crud-coverage-index-v1.json",
    "docs/performance/crud-coverage-index-v2.json",
)
PREVIOUS_POLICY_ID = "litchi-hosted-default-matrix-v4"
NEW_POLICY_ID = "litchi-hosted-default-matrix-v5"


def dump(value):
    return json.dumps(value, indent=2, ensure_ascii=False) + "\n"


def load_checked(repo, relative):
    text = (repo / relative).read_text(encoding="utf-8")
    value = json.loads(text)
    assert dump(value) == text, f"{relative} does not round-trip"
    return value


def manifest_keys(identity):
    return [
        (case, json.dumps(identity["corpora"][name], sort_keys=True, separators=(",", ":")))
        for case in identity["default_cases"]
        for name in identity["case_corpora"][case]
    ]


def main():
    repo, report_path, catalog_path = (Path(arg) for arg in sys.argv[1:4])
    write = sys.argv[4:] == ["--write"]
    sys.path.insert(0, str(repo))
    from tools import generate_corpus_manifest_v2 as catalog_tool
    from tools import perf_compare

    report = json.loads(report_path.read_text(encoding="utf-8"))
    rust_catalog = json.loads(catalog_path.read_text(encoding="utf-8"))
    previous = load_checked(repo, V1)
    previous_catalog = load_checked(repo, CATALOG)
    policy = load_checked(repo, POLICY)
    indexes = {path: load_checked(repo, path) for path in INDEXES}

    config = report["configuration"]
    assert config["samples_per_case"] == 1 and config["warmup_iterations_per_case"] == 0
    assert config["cases"] == previous["default_cases"]
    assert rust_catalog["build"]["git_worktree_dirty"] is False
    for name, expected in previous["identity_configuration"].items():
        assert config[name] == expected, name
    rows = report["results"]
    assert len(rows) == previous["result_count"] == 213

    observed = {}
    case_corpora = {case: [] for case in config["cases"]}
    report_keys = []
    for row in rows:
        assert len(row["elapsed_ns"]["samples"]) == 1
        corpus = row["corpus"]
        name = corpus["name"]
        assert observed.get(name, corpus) == corpus, name
        observed[name] = corpus
        case_corpora[row["case"]].append(name)
        report_keys.append((row["case"], json.dumps(corpus, sort_keys=True, separators=(",", ":"))))
    assert case_corpora == previous["case_corpora"]
    assert set(observed) == set(previous["corpora"])
    changed = [name for name in previous["corpora"] if observed[name] != previous["corpora"][name]]
    assert changed == list(DOC_CORPORA), changed
    for name in DOC_CORPORA:
        old, new = previous["corpora"][name], observed[name]
        moved = sorted(field for field in old if old[field] != new[field])
        assert set(old) == set(new) and moved == ["archive_sha256", "target_payload_sha256"], (name, moved)

    assert (
        perf_compare.result_key_manifest_sha256(manifest_keys(previous))
        == previous["result_keys_sha256"]
        == policy["expected_result_keys_sha256"]
    )
    identity = copy.deepcopy(previous)
    identity["corpora"] = {name: observed[name] for name in previous["corpora"]}
    identity["result_keys_sha256"] = perf_compare.result_key_manifest_sha256(report_keys)
    assert identity["result_keys_sha256"] == perf_compare.result_key_manifest_sha256(
        manifest_keys(identity)
    )

    catalog = catalog_tool.generate(
        identity,
        rust_catalog["build"]["git_revision"],
        worktree_dirty=rust_catalog["build"]["git_worktree_dirty"],
    )
    assert catalog == rust_catalog, "Python/Rust catalog derivation differs"
    old_ids = {corpus["name"]: corpus["id"] for corpus in previous_catalog["corpora"]}
    new_ids = {corpus["name"]: corpus["id"] for corpus in catalog["corpora"]}
    assert set(old_ids) == set(new_ids)
    assert [name for name in old_ids if old_ids[name] != new_ids[name]] == [
        name for name in old_ids if name in DOC_CORPORA
    ]

    assert policy["policy_id"] == PREVIOUS_POLICY_ID
    new_policy = copy.deepcopy(policy)
    new_policy["policy_id"] = NEW_POLICY_ID
    new_policy["expected_result_keys_sha256"] = identity["result_keys_sha256"]

    by_case = {}
    for binding in catalog["case_bindings"]:
        by_case.setdefault(binding["case"], []).append(binding["corpus_id"])
    corpus_by_id = {corpus["id"]: corpus for corpus in catalog["corpora"]}
    new_indexes = {}
    for path, index in indexes.items():
        updated = copy.deepcopy(index)
        reference = updated["checked_catalog"]
        assert reference["catalog_sha256"] == previous_catalog["catalog_sha256"], path
        assert reference["content_set_sha256"] == previous_catalog["content_set_sha256"], path
        reference["catalog_sha256"] = catalog["catalog_sha256"]
        reference["content_set_sha256"] = catalog["content_set_sha256"]
        touched = []
        for category in updated["categories"]:
            for scenario in category["scenarios"]:
                corpus = scenario.get("corpus")
                if not isinstance(corpus, dict) or corpus.get("kind") != "checked-catalog":
                    continue
                ids = sorted(by_case[corpus["case"]])
                shapes = sorted({corpus_by_id[item]["legacy_v1"]["shape"] for item in ids})
                if (corpus["ids"], corpus["shapes"]) != (ids, shapes):
                    touched.append(scenario["selector"])
                    corpus["ids"] = ids
                    corpus["shapes"] = shapes
        assert touched == ["doc_fresh_write_to"], (path, touched)
        new_indexes[path] = updated

    summary = {
        "status": "pass" if write else "verified",
        "revision": rust_catalog["build"]["git_revision"],
        "changed_corpora": {
            name: {
                field: {"old": previous["corpora"][name][field], "new": observed[name][field]}
                for field in ("archive_sha256", "target_payload_sha256")
            }
            for name in DOC_CORPORA
        },
        "unchanged_corpora": len(previous["corpora"]) - len(DOC_CORPORA),
        "result_keys_sha256": {"old": previous["result_keys_sha256"], "new": identity["result_keys_sha256"]},
        "catalog_sha256": {"old": previous_catalog["catalog_sha256"], "new": catalog["catalog_sha256"]},
        "content_set_sha256": {"old": previous_catalog["content_set_sha256"], "new": catalog["content_set_sha256"]},
        "policy_id": {"old": PREVIOUS_POLICY_ID, "new": NEW_POLICY_ID},
    }
    if write:
        (repo / V1).write_text(dump(identity), encoding="utf-8")
        (repo / CATALOG).write_text(dump(catalog), encoding="utf-8")
        (repo / POLICY).write_text(dump(new_policy), encoding="utf-8")
        for path, index in new_indexes.items():
            (repo / path).write_text(dump(index), encoding="utf-8")
    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()
