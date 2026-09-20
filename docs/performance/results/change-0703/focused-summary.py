#!/usr/bin/env python3
"""Independently isolate each transaction, excluding setup and verification."""
import hashlib
import json
from collections import Counter
from pathlib import Path
P = Path(__file__).resolve().parent
KEYS = ("raw_sha256", "raw_len", "output_sha256", "output_len", "ownership", "profile",
        "limits_input", "limits_output", "limits_depth", "limits_bindings", "limits_directives", "limits_choices")
STAGES = ("open", "capture", "clone", "edit", "commit", "apply", "verify")

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def summarize():
    runs = json.loads((P / "runs.json").read_text())
    assert len(runs) == 12
    assert {(r["case"],r["workflow"],r["repeat"]) for r in runs} == {
        (c,w,n) for c in ("real","generated") for w in ("noop","one","two") for n in range(2)}
    rows = []
    for run in runs:
        assert run["exit_code"] == 0
        for stream in ("stdout", "stderr"):
            assert sha(P / run[stream]) == run[stream + "_sha256"]
        records = []
        boundaries = []
        last_phase = "unset"
        for line in (P / run["stderr"]).read_text().splitlines():
            if line.startswith("LITCHI0703_BOUNDARY "):
                fields = dict(token.split("=",1) for token in line.split()[1:])
                last_phase = fields["phase"]
                boundaries.append(last_phase)
            if line.startswith("LITCHI0703_MCE "):
                record = dict(token.split("=",1) for token in line.split()[1:])
                assert int(record["call"]) == len(records)
                assert record["phase"] == last_phase
                assert record["profile"] == "default-ooxml"
                assert record["status"] == "ok"
                if record["ownership"] == "borrowed":
                    assert record["raw_sha256"] == record["output_sha256"]
                    assert record["raw_len"] == record["output_len"]
                    assert record["raw_ptr"] == record["output_ptr"]
                records.append(record)
        prefix = run["workflow"] + ".i0."
        selected = [r for r in records if r["phase"].startswith(prefix)]
        assert [b[len(prefix):] for b in boundaries if b.startswith(prefix)] == list(STAGES)
        phases = {stage:[r for r in selected if r["phase"] == prefix+stage] for stage in STAGES}
        key = lambda r: tuple(r[k] for k in KEYS)
        capture = phases["capture"]
        commit = phases["commit"]
        first = {}
        for r in capture:
            k = key(r)
            if k in first:
                assert first[k]["output_capacity"] == r["output_capacity"]
            first.setdefault(k,r)
        hits = [r for r in commit if key(r) in first]
        hit_keys = {key(r) for r in hits}
        owned = [r for r in first.values() if r["ownership"] == "owned"]
        results = [line for line in (P/run["stdout"]).read_text().splitlines() if line.startswith("result\t")]
        assert len(results) == 1
        result = dict(field.split("=",1) for field in results[0].split("\t")[1:])
        assert result["iteration"] == "0" and result["workflow"] == run["workflow"]
        assert result["changed"] == result["commit_changed"] == ("false" if run["workflow"]=="noop" else "true")
        if run["workflow"] == "noop":
            assert len(commit) == 0
        row = dict(name=run["name"],case=run["case"],workflow=run["workflow"],repeat=run["repeat"],
            all_trace_calls=len(records), excluded_setup_and_verify_calls=len(records)-sum(len(phases[s]) for s in STAGES if s!="verify"),
            phase_calls={stage:len(value) for stage,value in phases.items()},
            capture_unique_exact_pairs=len(first),capture_duplicate_calls=len(capture)-len(first),
            capture_owned_unique_pairs=len(owned),
            capture_unique_raw_payload_bytes=sum(int(r["raw_len"]) for r in first.values()),
            capture_owned_output_length_bytes=sum(int(r["output_len"]) for r in owned),
            capture_owned_output_capacity_bytes=sum(int(r["output_capacity"]) for r in owned),
            commit_calls_matching_capture_pairs=len(hits),commit_unique_pairs_matching_capture=len(hit_keys),
            commit_calls_not_matching_capture_pairs=len(commit)-len(hits),
            matched_owned_output_capacity_bytes=sum(int(first[k]["output_capacity"]) for k in hit_keys if first[k]["ownership"]=="owned"),
            published_revision=result["revision"],
            capture_pairs=[{k:r[k] for k in KEYS+("output_capacity",)} for r in first.values()],
            commit_pairs=[{k:r[k] for k in KEYS+("output_capacity",)} for r in commit])
        rows.append(row)
    for case in ("real","generated"):
        for workflow in ("noop","one","two"):
            a,b = [r for r in rows if r["case"]==case and r["workflow"]==workflow]
            for key in set(a)-{"name","repeat"}:
                assert a[key] == b[key], (case,workflow,key)
    return dict(performance_claim="none",diagnostic_only=True,
        scope="One transaction per fresh process; setup and verify excluded from reuse counts; default process_ooxml only",
        retention_caveat="Distinct raw/output/profile pairs and observed Vec capacities; excludes metadata, source-owner overhead, allocator rounding and policy admission. This is not a cache implementation, identity proof, peak allocation or RSS measurement.",
        fresh_repeat_equality=True,rows=rows)

def main():
    (P / "focused-summary.json").write_text(json.dumps(summarize(),indent=2)+"\n")
    print("PASS: 12 traces, phase boundaries, successful publication and fresh-repeat equality")

if __name__ == "__main__":
    main()
