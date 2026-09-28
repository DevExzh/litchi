"""Root-only application of the reviewed one-file candidate after qualification."""

import subprocess

import custody as c


P = c.P
assert not (P / "application.json").exists()
assert not (P / "build-after").exists()
qualification = c.read(P / "qualification/complete.json")
assert qualification["reports"] == 18
assert qualification["samples"] == 18
audit_path = P / "qualification-audit.json"
assert audit_path.is_file(), "independent accepted qualification audit is required"
audit = c.read(audit_path)
assert audit["schema"] == "litchi.performance.0810.qualification-audit.v1"
assert audit["passed"] is True
assert audit["accepted_before_application"] is True
assert audit["reports"] == 18 and audit["samples"] == 18
before = c.read(P / "build-before/source.json")
assert c.source() == before

manifest_path = P / "candidate/manifest.json"
patch_path = P / "candidate/candidate.patch"
manifest = c.read(manifest_path)
allowlist = set(c.read(P / "plan.json")["source_allowlist"])
assert manifest.get("production_source_changed", False) is False
if "files" in manifest:
    rows = list(manifest["files"].values())
    assert {row["production_path"] for row in rows} == allowlist
else:
    assert manifest["production_path"] in allowlist
    rows = [
        {
            "production_path": manifest["production_path"],
            "before": manifest["before"],
            "after": manifest["after"],
        }
    ]
assert len(rows) == 1
expected = dict(before["files"])
for row in rows:
    before_path = P / row["before"]["path"]
    after_path = P / row["after"]["path"]
    assert before_path.stat().st_size == row["before"]["bytes"]
    assert c.sha(before_path) == row["before"]["sha256"]
    assert after_path.stat().st_size == row["after"]["bytes"]
    assert c.sha(after_path) == row["after"]["sha256"]
    path = c.ROOT / row["production_path"]
    assert c.sha(path) == row["before"]["sha256"]
    expected[row["production_path"]] = row["after"]["sha256"]

assert patch_path.stat().st_size == manifest["patch"]["bytes"]
assert c.sha(patch_path) == manifest["patch"]["sha256"]
subprocess.run(["git", "apply", "--check", str(patch_path)], cwd=c.ROOT, check=True)
subprocess.run(["git", "apply", str(patch_path)], cwd=c.ROOT, check=True)
after = c.source()
assert after["revision"] == before["revision"]
assert c.changed_files(before, after) == allowlist
assert after["files"] == expected
c.write(
    P / "application.json",
    {
        "schema": "litchi.performance.0810.application.v1",
        "manifest": c.artifact(manifest_path),
        "patch": c.artifact(patch_path),
        "source_before": before,
        "source": after,
        "allowlist": sorted(allowlist),
    },
)
print("0810 one-file candidate applied after 18 before qualification", flush=True)
