"""Measure source-extracted private type declarations with real public types.

This diagnostic does not replace the compiled crate or enter timing results.
The private declarations are copied byte-for-byte from each frozen archive;
Width, Outline and ColumnIndex come from the actual baseline library metadata.
"""
import json
from pathlib import Path
import subprocess
import sys
import time

import driver as d

P = d.P / "layout"
SOURCE = "crates/litchi-xlsx/src/column.rs"


def declaration(source, marker, derive=True):
    start = source.index(marker)
    if derive:
        start = source.rfind("#[derive(", 0, start)
        assert start >= 0
    opened = source.index("{", start)
    depth = 1
    end = opened + 1
    while depth:
        depth += (source[end] == "{") - (source[end] == "}")
        end += 1
    return source[start:end]


def generated_source():
    chunks = ["#![allow(dead_code)]\n"]
    types = ("Properties", "Node<Properties>", "Assigned<Properties>",
             "Assignments<Properties>", "Node<usize>", "Assigned<usize>",
             "Assignments<usize>")
    for leg in ("before", "after"):
        source = d.candidate_path(leg, SOURCE).read_text()
        chunks.append(f"mod {leg} {{\nuse bitflags::bitflags;\n"
                      "use litchi_xlsx::{ColumnIndex as Index, Width, Outline};\n")
        chunks.append(declaration(source, "bitflags! {", False))
        for marker in ("pub(crate) struct Properties {", "enum Node<T> {",
                       "pub(crate) struct Assignments<T> {",
                       "pub(crate) struct Assigned<T> {"):
            chunks.append(declaration(source, marker))
        chunks.append("pub(super) fn print() {\n")
        for name in types:
            chunks.append(f'println!("{leg}\\t{name}\\t{{}}\\t{{}}", '
                          f'std::mem::size_of::<{name}>(), std::mem::align_of::<{name}>());\n')
        chunks.append("}\n}\n")
    chunks.append("fn main() { before::print(); after::print(); }\n")
    return "\n".join(chunks)


def descriptor(path):
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": d.sha(path)}


def capture():
    d.check("after")
    assert not P.exists()
    P.mkdir()
    source = P / "declarations.rs"
    source.write_text(generated_source())
    dependencies = d.TARGET / "release/release/deps"
    libraries = {}
    for name in ("litchi_xlsx", "bitflags"):
        matches = list(dependencies.glob("lib" + name + "-*.rlib"))
        assert len(matches) == 1, (name, matches)
        libraries[name] = descriptor(matches[0])
    binary = d.TARGET / "layout-probe"
    assert not binary.exists()
    argv = ["rustc", "--edition", "2024", "-C", "opt-level=3",
            "-C", "debuginfo=0", "-L", "dependency=" + str(dependencies)]
    for name, row in libraries.items():
        argv += ["--extern", name + "=" + row["path"]]
    argv += [str(source), "-o", str(binary)]
    row = {"argv": argv, "cwd": str(d.ROOT), "started_unix": time.time(),
           "source": descriptor(source), "libraries": libraries,
           "script_sha256": d.sha(Path(__file__)),
           "input_inventory_sha256": d.sha(d.P / "inputs.json"),
           "baseline_build_sha256": d.sha(d.P / "build-before.json"),
           "rustc": d.output(["rustc", "-Vv"])}
    d.write(P / "compile.started.json", row)
    with (P / "compile.log").open("xb") as log:
        child = subprocess.run(argv, cwd=d.ROOT, env=d.env("capture"),
                               stdout=log, stderr=subprocess.STDOUT)
    row.update(exit_code=child.returncode, finished_unix=time.time(),
               log_sha256=d.sha(P / "compile.log"))
    d.write(P / "compile.json", row)
    assert child.returncode == 0
    assert all(d.sha(Path(value["path"])) == value["sha256"] for value in libraries.values())
    executable = descriptor(binary)
    run = {"argv": [str(binary)], "cwd": str(d.ROOT), "binary": executable,
           "started_unix": time.time()}
    d.write(P / "run.started.json", run)
    with (P / "output.tsv").open("xb") as output:
        child = subprocess.run([str(binary)], cwd=d.ROOT, stdout=output, stderr=subprocess.STDOUT)
    run.update(exit_code=child.returncode, finished_unix=time.time(),
               output_sha256=d.sha(P / "output.tsv"))
    d.write(P / "run.json", run)
    assert child.returncode == 0
    d.check("after")
    print((P / "output.tsv").read_text(), end="")


def check():
    compile_receipt = d.read(P / "compile.json")
    run = d.read(P / "run.json")
    assert compile_receipt["exit_code"] == run["exit_code"] == 0
    assert (P / "declarations.rs").read_text() == generated_source()
    assert d.sha(P / "declarations.rs") == compile_receipt["source"]["sha256"]
    assert d.sha(Path(__file__)) == compile_receipt["script_sha256"]
    assert d.sha(P / "compile.log") == compile_receipt["log_sha256"]
    assert d.sha(P / "output.tsv") == run["output_sha256"]
    rows = {}
    for line in (P / "output.tsv").read_text().splitlines():
        leg, name, size, alignment = line.split("\t")
        assert (leg, name) not in rows
        rows[leg, name] = (int(size), int(alignment))
    assert len(rows) == 14
    assert rows["before", "Node<Properties>"][0] * 32768 == 1048576
    assert rows["before", "Node<usize>"][0] * 32768 == 524288
    for name in ("Properties", "Node<Properties>", "Assigned<Properties>",
                 "Node<usize>", "Assigned<usize>"):
        assert rows["before", name] == rows["after", name]
    print("0832 declaration-layout replay PASS; no cache-miss or runtime claim")


if __name__ == "__main__":
    if sys.argv[1:] == ["--check"]:
        check()
    else:
        assert not sys.argv[1:]
        capture()
        check()
