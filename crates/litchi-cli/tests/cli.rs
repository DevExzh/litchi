//! End-to-end tests of the `litchi` binary (record 0256).
//!
//! Each test runs the built binary over repository fixtures or files it
//! creates, and checks two things: the `--json` output is exactly what the
//! in-process `litchi-ops` call serializes, and the exit status is the one
//! `src/exit.rs` documents for the error code.

use std::path::{Path, PathBuf};
use std::process::{Command, Output, Stdio};

use litchi_ops::node::ViewOptions;
use litchi_ops::view::Mode;
use serde_json::{Value, json};

const BIN: &str = env!("CARGO_BIN_EXE_litchi");

fn fixture(relative: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative)
}

struct Run {
    status: i32,
    stdout: String,
    stderr: String,
}

impl Run {
    fn json(&self) -> Value {
        serde_json::from_str(&self.stdout)
            .unwrap_or_else(|error| panic!("stdout is not JSON ({error}): {}", self.stdout))
    }
}

fn finish(output: Output) -> Run {
    Run {
        status: output.status.code().expect("the binary exits normally"),
        stdout: String::from_utf8(output.stdout).expect("stdout is UTF-8"),
        stderr: String::from_utf8(output.stderr).expect("stderr is UTF-8"),
    }
}

fn litchi<I, S>(args: I) -> Run
where
    I: IntoIterator<Item = S>,
    S: AsRef<std::ffi::OsStr>,
{
    finish(
        Command::new(BIN)
            .args(args)
            .stdin(Stdio::null())
            .output()
            .expect("run litchi"),
    )
}

/// Runs and asserts the exit status, reporting stderr on a mismatch.
fn expect<I, S>(status: i32, args: I) -> Run
where
    I: IntoIterator<Item = S>,
    S: AsRef<std::ffi::OsStr>,
{
    let run = litchi(args);
    assert_eq!(
        run.status, status,
        "unexpected status; stdout: {}\nstderr: {}",
        run.stdout, run.stderr
    );
    run
}

fn arg(path: &Path) -> &str {
    path.to_str().expect("UTF-8 path")
}

/// One fixture per format with a read adapter: paths to `get` and a
/// selector to `query`.
const READS: &[(&str, &[&str], &str)] = &[
    (
        "ooxml/docx/drawing.docx",
        &["/", "/body", "/body/p[2]"],
        "p",
    ),
    (
        "ooxml/xlsx/SimpleMultiCell.xlsx",
        &["/", "/Sheet1", "/sheet[1]/row[1]", "/Sheet1/A1"],
        "cell",
    ),
    (
        "ooxml/pptx/sample.pptx",
        &["/", "/slide[1]", "/slide[1]/shape[1]"],
        "shape",
    ),
    ("ooxml/xlsb/Simple.xlsb", &["/", "/sheet[1]"], "cell"),
    ("ole/doc/Lists.doc", &["/", "/body", "/body/p[3]"], "p"),
    ("ole/xls/Simple.xls", &["/", "/sheet[1]"], "cell"),
    (
        "ole/ppt/SampleShow.ppt",
        &["/", "/slide[1]", "/slide[2]"],
        "slide",
    ),
    ("rtf/column-break.rtf", &["/", "/body/p[1]"], "p"),
];

#[test]
fn get_json_is_the_node_view() {
    for &(relative, paths, _) in READS {
        let path = fixture(relative);
        let file = litchi_ops::open(&path).expect("open in process");
        for &node_path in paths {
            for depth in [0_u32, 1, 3] {
                let node = file.get(node_path).expect("get in process");
                let expected =
                    serde_json::to_value(node.view_with(ViewOptions::default().depth(depth)))
                        .expect("serialize in process");
                let depth = depth.to_string();
                let run = expect(
                    0,
                    ["get", arg(&path), node_path, "--depth", &depth, "--json"],
                );
                assert_eq!(run.json(), expected, "{relative} {node_path} depth {depth}");
            }
        }
        // The default depth is the view's default, one level.
        let expected = serde_json::to_value(file.root().expect("root").view()).expect("root view");
        assert_eq!(expect(0, ["get", arg(&path), "--json"]).json(), expected);
    }
}

#[test]
fn get_json_keeps_officecli_member_order() {
    let run = expect(
        0,
        [
            "get",
            arg(&fixture("ole/doc/Lists.doc")),
            "/body/p[1]",
            "--json",
        ],
    );
    let order: Vec<usize> = [
        "\"path\"",
        "\"type\"",
        "\"text\"",
        "\"childCount\"",
        "\"format\"",
    ]
    .iter()
    .map(|member| run.stdout.find(member).expect("member present"))
    .collect();
    assert!(
        order.windows(2).all(|pair| pair[0] < pair[1]),
        "{}",
        run.stdout
    );
}

#[test]
fn query_json_is_officecli_nodes_result_of_node_views() {
    for &(relative, _, selector) in READS {
        let path = fixture(relative);
        let file = litchi_ops::open(&path).expect("open in process");
        let views: Vec<Value> = file
            .query(selector)
            .expect("query in process")
            .map(|node| serde_json::to_value(node.expect("match").view()).expect("view"))
            .collect();
        let run = expect(0, ["query", arg(&path), selector, "--json"]);
        assert_eq!(
            run.json(),
            json!({ "matches": views.len(), "results": views }),
            "{relative} {selector}"
        );
        let limited = expect(0, ["query", arg(&path), selector, "--limit", "1", "--json"]);
        assert_eq!(
            limited.json(),
            json!({ "matches": views.len().min(1), "results": views.iter().take(1).collect::<Vec<_>>() })
        );
    }
}

#[test]
fn view_output_is_the_rendered_view() {
    for &(relative, _, _) in READS {
        let path = fixture(relative);
        let file = litchi_ops::open(&path).expect("open in process");
        for mode in [Mode::Text, Mode::Annotated, Mode::Outline, Mode::Stats] {
            let expected = serde_json::to_value(file.view_json(mode).expect("view_json"))
                .expect("serialize view");
            let run = expect(0, ["view", arg(&path), "--mode", mode.name(), "--json"]);
            assert_eq!(run.json(), expected, "{relative} {}", mode.name());
            let text = file.view(mode).expect("view");
            let run = expect(0, ["view", arg(&path), mode.name()]);
            assert_eq!(
                run.stdout.trim_end_matches('\n'),
                text.trim_end_matches('\n')
            );
        }
        let run = expect(5, ["view", arg(&path), "html"]);
        assert!(
            run.stderr.contains("error[unsupported_capability]"),
            "{}",
            run.stderr
        );
    }
}

#[test]
fn capabilities_json_is_the_capability_report() {
    for &(relative, _, _) in READS {
        let path = fixture(relative);
        let file = litchi_ops::open(&path).expect("open in process");
        let expected = serde_json::to_value(file.capabilities()).expect("serialize");
        assert_eq!(
            expect(0, ["capabilities", arg(&path), "--json"]).json(),
            expected
        );
    }
    for format in [
        "docx", "xlsx", "pptx", "xlsb", "doc", "xls", "ppt", "rtf", "odt", "pages",
    ] {
        let parsed = match format {
            "docx" => litchi_core::FileFormat::Docx,
            "xlsx" => litchi_core::FileFormat::Xlsx,
            "pptx" => litchi_core::FileFormat::Pptx,
            "xlsb" => litchi_core::FileFormat::Xlsb,
            "doc" => litchi_core::FileFormat::Doc,
            "xls" => litchi_core::FileFormat::Xls,
            "ppt" => litchi_core::FileFormat::Ppt,
            "rtf" => litchi_core::FileFormat::Rtf,
            "odt" => litchi_core::FileFormat::Odt,
            _ => litchi_core::FileFormat::Pages,
        };
        let expected = serde_json::to_value(litchi_ops::capabilities::Capabilities::of(parsed))
            .expect("serialize");
        assert_eq!(
            expect(0, ["capabilities", "--format", format, "--json"]).json(),
            expected,
            "{format}"
        );
        let text = expect(0, ["capabilities", "--format", format]);
        assert!(
            text.stdout.starts_with(&format!("format: {format}\n")),
            "{}",
            text.stdout
        );
    }
    expect(2, ["capabilities"]);
    expect(4, ["capabilities", "--format", "txt"]);
}

#[test]
fn schema_json_is_the_registry_export() {
    for (format, parsed) in [
        ("docx", litchi_core::FileFormat::Docx),
        ("xlsx", litchi_core::FileFormat::Xlsx),
        ("doc", litchi_core::FileFormat::Doc),
        ("rtf", litchi_core::FileFormat::Rtf),
        ("pages", litchi_core::FileFormat::Pages),
    ] {
        let expected = serde_json::to_value(litchi_ops::schema(parsed)).expect("serialize");
        assert_eq!(
            expect(0, ["schema", format, "--json"]).json(),
            expected,
            "{format}"
        );
        let text = expect(0, ["schema", format]);
        assert!(
            text.stdout.starts_with(&format!("format: {format}\n")),
            "{}",
            text.stdout
        );
    }
    let xlsx = expect(0, ["schema", "xlsx"]);
    assert!(
        xlsx.stdout
            .contains("  outline int -s- aliases=group,outlinelevel\n"),
        "{}",
        xlsx.stdout
    );
    expect(2, ["schema"]);
    expect(4, ["schema", "txt"]);
}

#[test]
fn exit_statuses_follow_the_code_families() {
    let docx = fixture("ooxml/docx/drawing.docx");
    let docx = arg(&docx);
    let dir = tempfile::tempdir().expect("tempdir");

    // 2: usage.
    expect(2, ["get"]);
    expect(2, ["frobnicate"]);
    expect(
        2,
        [
            "add",
            docx,
            "/body",
            "--type",
            "paragraph",
            "--index",
            "0",
            "--after",
            "/body/p[1]",
        ],
    );
    // 3: path.
    let run = expect(3, ["get", docx, "/body/p[100000]"]);
    assert!(
        run.stderr.starts_with("error[path_not_found]: "),
        "{}",
        run.stderr
    );
    let run = expect(3, ["get", docx, "/body/p[", "--json"]);
    let error: Value = serde_json::from_str(&run.stderr).expect("JSON error on stderr");
    assert_eq!(error["code"], "invalid_path");
    assert!(
        error["error"]
            .as_str()
            .is_some_and(|message| !message.is_empty())
    );
    assert!(run.stdout.is_empty());
    expect(3, ["query", docx, "p[[["]);
    // 4: property or argument.
    let copy = dir.path().join("copy.docx");
    std::fs::copy(docx, &copy).expect("copy");
    let before = std::fs::read(&copy).expect("read");
    let run = expect(
        4,
        [
            "set",
            arg(&copy),
            "/body/p[1]",
            "--prop",
            "bolld=true",
            "--json",
        ],
    );
    let error: Value = serde_json::from_str(&run.stderr).expect("JSON error");
    assert_eq!(error["code"], "unsupported_property");
    assert_eq!(error["property"], "bolld");
    assert_eq!(error["suggestion"], "did you mean 'bold'?");
    assert_eq!(
        std::fs::read(&copy).expect("read"),
        before,
        "a failed set never writes"
    );
    expect(4, ["view", docx, "--mode", "nonsense"]);
    expect(4, ["create", arg(&dir.path().join("x")), "--type", "txt"]);
    // 5: unsupported.
    expect(5, ["view", docx, "pdf"]);
    let text = dir.path().join("plain.txt");
    std::fs::write(&text, "hello").expect("write");
    expect(5, ["get", arg(&text)]);
    expect(5, ["create", arg(&dir.path().join("x.pages"))]);
    // 6: I/O.
    expect(6, ["get", arg(&dir.path().join("missing.docx"))]);
    expect(6, ["create", arg(&copy)]);
    expect(
        6,
        [
            "batch",
            arg(&copy),
            "--input",
            arg(&dir.path().join("missing.json")),
        ],
    );
    // 7: corrupt input.
    let corrupt = dir.path().join("corrupt.docx");
    std::fs::write(&corrupt, b"PK\x03\x04garbage").expect("write");
    let run = expect(7, ["get", arg(&corrupt)]);
    assert!(
        run.stderr.starts_with("error[corrupt_file]: "),
        "{}",
        run.stderr
    );
    // 8: protected.
    expect(8, ["get", arg(&fixture("ole/doc/PasswordProtected.doc"))]);
    let signed = dir.path().join("signed.docx");
    std::fs::copy(
        fixture("poi/test-data/xmldsign/hello-world-signed.docx"),
        &signed,
    )
    .expect("copy");
    let run = expect(8, ["set", arg(&signed), "/body/p[1]", "--prop", "text=x"]);
    assert!(run.stderr.contains("suggestion: "), "{}", run.stderr);
}

/// Copies a fixture or creates a blank file of `name`'s type in `dir`.
fn scratch(dir: &Path, name: &str, source: Option<&str>) -> PathBuf {
    let path = dir.join(name);
    match source {
        Some(relative) => {
            std::fs::copy(fixture(relative), &path).expect("copy fixture");
        },
        None => {
            let run = expect(0, ["create", arg(&path), "--json"]);
            assert_eq!(run.json()["path"], json!(path.display().to_string()));
        },
    }
    path
}

/// The node's `text`, or its `preview` when it has no text (sheets).
fn text_of(path: &Path, node: &str) -> Value {
    let view = expect(0, ["get", arg(path), node, "--json"]).json();
    match view.get("text") {
        Some(text) => text.clone(),
        None => view["preview"].clone(),
    }
}

/// The single-item batch rendering of a successful write.
fn applied(run: &Run) -> Value {
    let report = run.json();
    assert_eq!(report["summary"]["total"], 1);
    assert_eq!(report["summary"]["succeeded"], 1);
    assert_eq!(report["results"][0]["success"], true);
    report
}

#[test]
fn create_writes_every_creatable_format_and_refuses_overwrite() {
    let dir = tempfile::tempdir().expect("tempdir");
    for format in ["docx", "xlsx", "pptx", "xlsb", "doc", "xls", "ppt", "rtf"] {
        let path = scratch(dir.path(), &format!("blank.{format}"), None);
        let file = litchi_ops::open(&path).expect("reopen in process");
        assert_eq!(
            file.format()
                .map(|format| format!("{format:?}").to_ascii_lowercase()),
            Some(format.to_owned())
        );
        expect(6, ["create", arg(&path)]);
        expect(0, ["create", arg(&path), "--force"]);
    }
    // A path without an extension gains the type's, as OfficeCLI does.
    expect(
        0,
        ["create", arg(&dir.path().join("noext")), "--type", "docx"],
    );
    assert!(dir.path().join("noext.docx").is_file());
    expect(4, ["create", arg(&dir.path().join("noext2"))]);
}

#[test]
fn docx_writes_save_in_place() {
    let dir = tempfile::tempdir().expect("tempdir");
    let path = scratch(dir.path(), "w.docx", None);
    let file = arg(&path);
    let run = expect(
        0,
        [
            "add",
            file,
            "/body",
            "--type",
            "paragraph",
            "--prop",
            "text=First",
            "--json",
        ],
    );
    let report = applied(&run);
    assert_eq!(report["results"][0]["path"], "/body/p[1]");
    assert_eq!(report["summary"]["changed"], true);
    assert_eq!(
        report["summary"]["saved"],
        json!(path.display().to_string())
    );
    expect(
        0,
        [
            "add",
            file,
            "/body",
            "--type",
            "paragraph",
            "--prop",
            "text=Second",
        ],
    );
    expect(
        0,
        [
            "add",
            file,
            "/body",
            "--type",
            "paragraph",
            "--prop",
            "text=Zeroth",
            "--at",
            "0",
        ],
    );
    expect(
        0,
        [
            "add",
            file,
            "/body",
            "--type",
            "paragraph",
            "--prop",
            "text=Between",
            "--after",
            "/body/p[2]",
        ],
    );
    assert_eq!(text_of(&path, "/body/p[1]"), "Zeroth");
    assert_eq!(text_of(&path, "/body/p[3]"), "Between");
    expect(0, ["set", file, "/body/p[2]/r[1]", "--prop", "bold=true"]);
    let bold = expect(0, ["get", file, "/body/p[2]/r[1]", "--json"]).json();
    assert_eq!(bold["format"]["bold"], true);
    // An exact no-op reports no change and leaves the file untouched.
    let before = std::fs::read(&path).expect("read");
    let run = expect(
        0,
        [
            "set",
            file,
            "/body/p[2]/r[1]",
            "--prop",
            "bold=true",
            "--json",
        ],
    );
    assert_eq!(run.json()["results"][0]["changed"], false);
    assert_eq!(run.json()["summary"].get("saved"), None);
    assert_eq!(std::fs::read(&path).expect("read"), before);
    expect(0, ["swap", file, "/body/p[1]", "/body/p[2]"]);
    assert_eq!(text_of(&path, "/body/p[1]"), "First");
    expect(0, ["move", file, "/body/p[2]", "--index", "0"]);
    assert_eq!(text_of(&path, "/body/p[1]"), "Zeroth");
    expect(0, ["remove", file, "/body/p[1]"]);
    assert_eq!(text_of(&path, "/body/p[1]"), "First");
    // `--out` publishes elsewhere and leaves the input as it was.
    let before = std::fs::read(&path).expect("read");
    let out = dir.path().join("out.docx");
    let run = expect(
        0,
        ["remove", file, "/body/p[1]", "--out", arg(&out), "--json"],
    );
    assert_eq!(
        run.json()["summary"]["saved"],
        json!(out.display().to_string())
    );
    assert_eq!(std::fs::read(&path).expect("read"), before);
    assert_eq!(text_of(&out, "/body/p[1]"), "Between");
}

#[test]
fn spreadsheet_writes_save_in_place() {
    let dir = tempfile::tempdir().expect("tempdir");
    for format in ["xlsx", "xlsb"] {
        let path = scratch(dir.path(), &format!("w.{format}"), None);
        let file = arg(&path);
        expect(0, ["set", file, "/Sheet1/A1", "--prop", "value=42"]);
        assert_eq!(text_of(&path, "/Sheet1/A1"), "42", "{format}");
        let run = expect(
            0,
            [
                "add",
                file,
                "/",
                "--type",
                "sheet",
                "--prop",
                "name=Data",
                "--json",
            ],
        );
        assert_eq!(applied(&run)["results"][0]["path"], "/sheet[2]", "{format}");
        expect(0, ["swap", file, "/sheet[1]", "/sheet[2]"]);
        expect(0, ["move", file, "/Sheet1", "--index", "0"]);
        assert_eq!(text_of(&path, "/sheet[1]"), "Sheet1", "{format}");
        expect(0, ["remove", file, "/Data"]);
        expect(3, ["get", file, "/Data"]);
    }
    let path = scratch(dir.path(), "w.xls", None);
    let file = arg(&path);
    expect(0, ["set", file, "/Sheet1/A1", "--prop", "value=3"]);
    assert_eq!(text_of(&path, "/Sheet1/A1"), "3");
    expect(0, ["remove", file, "/Sheet1/A1"]);
}

#[test]
fn presentation_writes_save_in_place() {
    let dir = tempfile::tempdir().expect("tempdir");
    let path = scratch(dir.path(), "w.pptx", None);
    let file = arg(&path);
    expect(
        0,
        [
            "add",
            file,
            "/",
            "--type",
            "slide",
            "--prop",
            "layout=Title Slide",
        ],
    );
    expect(
        0,
        [
            "add",
            file,
            "/",
            "--type",
            "slide",
            "--prop",
            "layout=Title and Content",
        ],
    );
    let run = expect(
        0,
        [
            "add",
            file,
            "/slide[1]",
            "--type",
            "textbox",
            "--prop",
            "text=Hello",
            "--prop",
            "x=1cm",
            "--prop",
            "y=1cm",
            "--prop",
            "width=4cm",
            "--prop",
            "height=2cm",
            "--json",
        ],
    );
    let added = applied(&run)["results"][0]["path"]
        .as_str()
        .expect("path")
        .to_owned();
    expect(0, ["set", file, &added, "--prop", "text=Changed"]);
    assert_eq!(text_of(&path, &added), "Changed");
    expect(0, ["swap", file, "/slide[1]", "/slide[2]"]);
    expect(0, ["move", file, "/slide[2]", "--index", "0"]);
    expect(0, ["remove", file, "/slide[2]"]);
    assert_eq!(
        expect(0, ["get", file, "/", "--json"]).json()["children"]
            .as_array()
            .map(|c| c.iter().filter(|n| n["type"] == "slide").count()),
        Some(1)
    );

    let path = scratch(dir.path(), "w.ppt", Some("ole/ppt/SampleShow.ppt"));
    let file = arg(&path);
    let first = text_of(&path, "/slide[1]");
    expect(0, ["swap", file, "/slide[1]", "/slide[2]"]);
    expect(0, ["move", file, "/slide[2]", "--index", "0"]);
    assert_eq!(text_of(&path, "/slide[1]"), first);
    expect(0, ["remove", file, "/slide[2]"]);
    expect(3, ["get", file, "/slide[2]"]);
}

#[test]
fn document_writes_save_in_place() {
    let dir = tempfile::tempdir().expect("tempdir");
    let path = scratch(dir.path(), "w.rtf", None);
    let file = arg(&path);
    expect(
        0,
        [
            "add",
            file,
            "/body",
            "--type",
            "paragraph",
            "--prop",
            "text=One",
        ],
    );
    expect(
        0,
        [
            "add",
            file,
            "/body",
            "--type",
            "paragraph",
            "--prop",
            "text=Two",
        ],
    );
    expect(
        0,
        [
            "add",
            file,
            "/body",
            "--type",
            "paragraph",
            "--prop",
            "text=Three",
        ],
    );
    expect(0, ["swap", file, "/body/p[1]", "/body/p[2]"]);
    expect(0, ["move", file, "/body/p[2]", "--index", "0"]);
    assert_eq!(text_of(&path, "/body/p[1]"), "One");
    expect(0, ["remove", file, "/body/p[3]"]);
    expect(0, ["set", file, "/body/p[1]/r[1]", "--prop", "bold=true"]);
    assert_eq!(
        expect(0, ["get", file, "/body/p[1]/r[1]", "--json"]).json()["format"]["bold"],
        true
    );

    let path = scratch(dir.path(), "w.doc", None);
    expect(0, ["set", arg(&path), "/body/p[1]", "--prop", "text=Hello"]);
    assert_eq!(text_of(&path, "/body/p[1]"), "Hello");
}

const ITEMS: &str = r#"[
  {"command":"add","parent":"/body","type":"paragraph","props":{"text":"B1"}},
  {"command":"set","path":"/body/p[1]/r[1]","props":{"bold":true}},
  {"command":"set","path":"/body/p[99]","props":{"bold":true}},
  {"command":"set","path":"/body/p[1]","props":{"bolld":true}}
]"#;

#[test]
fn batch_is_atomic_by_default_and_best_effort_on_request() {
    let dir = tempfile::tempdir().expect("tempdir");
    let path = scratch(dir.path(), "b.docx", None);
    let file = arg(&path);
    let input = dir.path().join("items.json");
    std::fs::write(&input, ITEMS).expect("write items");
    let before = std::fs::read(&path).expect("read");

    for flags in [&[][..], &["--atomic"][..]] {
        let mut args = vec!["batch", file, "--input", arg(&input), "--json"];
        args.extend_from_slice(flags);
        let run = expect(3, args);
        let report = run.json();
        assert_eq!(report["results"][0]["index"], 2);
        assert_eq!(report["results"][0]["code"], "path_not_found");
        assert_eq!(report["results"][0]["item"]["path"], "/body/p[99]");
        assert_eq!(report["summary"]["atomicRolledBack"], true);
        assert_eq!(report["summary"]["skipped"], 1);
        assert_eq!(
            std::fs::read(&path).expect("read"),
            before,
            "atomic failure never writes"
        );
    }

    let run = expect(
        9,
        [
            "batch",
            file,
            "--input",
            arg(&input),
            "--best-effort",
            "--json",
        ],
    );
    let report = run.json();
    let results = report["results"].as_array().expect("results");
    assert_eq!(results.len(), 4);
    assert_eq!(results[0]["path"], "/body/p[1]");
    assert_eq!(results[1]["success"], true);
    assert_eq!(results[2]["code"], "path_not_found");
    assert_eq!(results[3]["code"], "unsupported_property");
    assert_eq!(results[3]["property"], "bolld");
    assert_eq!(report["summary"]["failed"], 2);
    assert_eq!(report["summary"]["changed"], true);
    assert_eq!(text_of(&path, "/body/p[1]"), "B1");

    // Inline and stdin payloads; a clean batch exits 0.
    let run = expect(
        0,
        [
            "batch",
            file,
            "--commands",
            r#"[{"op":"add","parent":"/body","type":"paragraph","props":{"text":"B2"}}]"#,
        ],
    );
    assert!(
        run.stdout.starts_with("batch /body/p[2]: ok\n"),
        "{}",
        run.stdout
    );
    let mut child = Command::new(BIN)
        .args(["batch", file, "--input", "-", "--json"])
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .expect("spawn");
    {
        use std::io::Write as _;
        let mut stdin = child.stdin.take().expect("stdin");
        stdin
            .write_all(br#"[{"command":"remove","path":"/body/p[2]"}]"#)
            .expect("write stdin");
    }
    let run = finish(child.wait_with_output().expect("wait"));
    assert_eq!(run.status, 0, "{}", run.stderr);
    assert_eq!(run.json()["summary"]["succeeded"], 1);

    // Malformed payloads are refused before the document is touched.
    expect(4, ["batch", file, "--commands", "[{\"command\":"]);
    expect(2, ["batch", file]);
    expect(
        2,
        [
            "batch",
            file,
            "--atomic",
            "--best-effort",
            "--commands",
            "[]",
        ],
    );
}

#[test]
fn officecli_smoke_script_replays() {
    let script = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/officecli_smoke.sh");
    let dir = tempfile::tempdir().expect("tempdir");
    let output = Command::new("bash")
        .arg(&script)
        .arg(BIN)
        .arg(dir.path())
        .output();
    let Ok(output) = output else {
        eprintln!("bash is unavailable; skipping the OfficeCLI smoke script");
        return;
    };
    let run = finish(output);
    assert_eq!(
        run.status, 0,
        "stdout: {}\nstderr: {}",
        run.stdout, run.stderr
    );
}

/// OfficeCLI's `add --from`, `set --find/--replace` and `remove --prop`
/// reach the command channel, which refuses them with typed errors instead
/// of a usage error (record 0265).
#[test]
fn officecli_option_spellings_reach_the_command_channel() {
    let dir = tempfile::tempdir().expect("tempdir");
    let path = scratch(dir.path(), "o.docx", None);
    expect(
        0,
        [
            "add",
            arg(&path),
            "/body",
            "--type",
            "paragraph",
            "--prop",
            "text=fox",
        ],
    );
    let before = std::fs::read(&path).expect("read");

    let run = expect(
        5,
        ["add", arg(&path), "/body", "--from", "/body/p[1]", "--json"],
    );
    let error: Value = serde_json::from_str(&run.stderr).expect("JSON error");
    assert_eq!(error["code"], "unsupported_capability");
    // Exactly one of --type and --from.
    expect(2, ["add", arg(&path), "/body"]);
    expect(
        2,
        [
            "add",
            arg(&path),
            "/body",
            "--type",
            "paragraph",
            "--from",
            "/body/p[1]",
        ],
    );

    // `--find T` is `--prop find=T`; a find needs `replace` (or a format
    // property), as OfficeCLI's does (record 0273).
    let run = expect(
        4,
        ["set", arg(&path), "/body/p[1]", "--find", "fox", "--json"],
    );
    let error: Value = serde_json::from_str(&run.stderr).expect("JSON error");
    assert_eq!(error["code"], "missing_property");
    assert_eq!(error["property"], "replace");
    let run = expect(
        4,
        [
            "set",
            arg(&path),
            "/body/p[1]",
            "--replace",
            "cat",
            "--json",
        ],
    );
    let error: Value = serde_json::from_str(&run.stderr).expect("JSON error");
    assert_eq!(error["property"], "replace");
    // `set` still needs something to set.
    expect(2, ["set", arg(&path), "/body/p[1]"]);

    let run = expect(
        4,
        [
            "remove",
            arg(&path),
            "/body/p[1]",
            "--prop",
            "revision.author=Bob",
            "--json",
        ],
    );
    let error: Value = serde_json::from_str(&run.stderr).expect("JSON error");
    assert_eq!(error["code"], "invalid_argument");

    assert_eq!(
        std::fs::read(&path).expect("read"),
        before,
        "refusals write nothing"
    );
}

/// `set --find T --replace R` rewrites the matches and reports their count
/// as OfficeCLI's `matched` (record 0273); zero matches write nothing.
#[test]
fn find_and_replace_report_the_match_count() {
    let dir = tempfile::tempdir().expect("tempdir");
    let path = scratch(dir.path(), "f.docx", None);
    expect(
        0,
        [
            "add",
            arg(&path),
            "/body",
            "--type",
            "paragraph",
            "--prop",
            "text=the fox and the fox",
        ],
    );
    let run = expect(
        0,
        [
            "set",
            arg(&path),
            "/",
            "--find",
            "fox",
            "--replace",
            "cat",
            "--json",
        ],
    );
    let report = applied(&run);
    assert_eq!(report["results"][0]["matched"], 2);
    assert_eq!(report["results"][0]["changed"], true);
    assert_eq!(text_of(&path, "/body/p[1]"), json!("the cat and the cat"));
    let run = expect(
        0,
        ["set", arg(&path), "/", "--find", "fox", "--replace", "cat"],
    );
    assert!(run.stdout.contains("(0 matched)"), "{}", run.stdout);
    let before = std::fs::read(&path).expect("read");
    let run = expect(
        0,
        [
            "set",
            arg(&path),
            "/",
            "--prop",
            "find=absent",
            "--prop",
            "replace=x",
            "--json",
        ],
    );
    assert_eq!(applied(&run)["results"][0]["matched"], 0);
    assert_eq!(std::fs::read(&path).expect("read"), before);
}

#[test]
fn unsupported_warn_applies_the_rest_and_exits_partial() {
    let dir = tempfile::tempdir().expect("tempdir");
    let path = scratch(dir.path(), "w.xlsx", None);
    let before = std::fs::read(&path).expect("read");
    // The default refuses the whole command and writes nothing.
    let run = expect(
        4,
        [
            "set",
            arg(&path),
            "/Sheet1/A1",
            "--prop",
            "value=Total",
            "--prop",
            "sparkle=yes",
        ],
    );
    assert!(run.stderr.contains("property: sparkle"), "{}", run.stderr);
    assert_eq!(std::fs::read(&path).expect("read"), before);

    // `--unsupported warn` applies the value, warns and exits 9.
    let run = expect(
        9,
        [
            "set",
            arg(&path),
            "/Sheet1/A1",
            "--prop",
            "value=Total",
            "--prop",
            "sparkle=yes",
            "--unsupported",
            "warn",
            "--json",
        ],
    );
    let report = run.json();
    assert_eq!(report["results"][0]["success"], true);
    assert_eq!(report["results"][0]["changed"], true);
    let warnings = report["results"][0]["warnings"]
        .as_array()
        .expect("warnings");
    assert_eq!(warnings.len(), 1);
    assert_eq!(warnings[0]["code"], "unsupported_property");
    assert_eq!(warnings[0]["property"], "sparkle");
    assert_eq!(text_of(&path, "/Sheet1/A1"), json!("Total"));

    // Text output names the skipped property on stderr and the report line.
    let run = expect(
        9,
        [
            "set",
            arg(&path),
            "/Sheet1/A2",
            "--prop",
            "value=More",
            "--prop",
            "sparkle=yes",
            "--unsupported",
            "warn",
        ],
    );
    assert!(
        run.stderr.starts_with("warning[unsupported_property]:"),
        "{}",
        run.stderr
    );
    assert!(run.stderr.contains("  property: sparkle"), "{}", run.stderr);
    assert!(run.stdout.contains("(skipped: sparkle)"), "{}", run.stdout);

    // A clean call under warn exits 0; an invalid value still fails.
    expect(
        0,
        [
            "set",
            arg(&path),
            "/Sheet1/A3",
            "--prop",
            "value=3",
            "--unsupported",
            "warn",
        ],
    );
    let after = std::fs::read(&path).expect("read");
    expect(
        4,
        [
            "set",
            arg(&path),
            "/Sheet1/A3",
            "--prop",
            "font.bold=maybe",
            "--prop",
            "sparkle=yes",
            "--unsupported",
            "warn",
        ],
    );
    assert_eq!(std::fs::read(&path).expect("read"), after);

    // Batches take the policy per item and exit 9 when any item warned.
    let run = expect(
        9,
        [
            "batch",
            arg(&path),
            "--unsupported",
            "warn",
            "--json",
            "--commands",
            r#"[{"command":"set","path":"/Sheet1/B1","props":{"value":"x","sparkle":"y"}},
                {"command":"set","path":"/Sheet1/B2","props":{"value":"z"}}]"#,
        ],
    );
    let report = run.json();
    assert_eq!(report["summary"]["succeeded"], 2);
    assert_eq!(report["results"][0]["warnings"][0]["item"], 0);
    assert!(report["results"][1].get("warnings").is_none());
    assert_eq!(text_of(&path, "/Sheet1/B2"), json!("z"));
}

/// A `--unsupported warn` command that still fails after skipping prints the
/// skipped properties with its error, never the bare reduced failure
/// (record 0271).
#[test]
fn unsupported_warn_failures_name_the_skipped_properties() {
    let dir = tempfile::tempdir().expect("tempdir");
    let path = scratch(dir.path(), "w.pptx", None);
    expect(
        0,
        [
            "add",
            arg(&path),
            "/",
            "--type",
            "slide",
            "--prop",
            "layout=Blank",
        ],
    );
    let before = std::fs::read(&path).expect("read");
    let base = vec![
        "add",
        arg(&path),
        "/slide[1]",
        "--type",
        "chart",
        "--prop",
        "chartType=line",
        "--prop",
        "series1.trendline=linear",
        "--prop",
        "series1.categories=Sheet1!A2:A3",
        "--unsupported",
        "warn",
    ];
    let add = |extra: &[&'static str]| {
        let mut args = base.clone();
        args.extend_from_slice(extra);
        args
    };
    // Text: each skipped property as a warning, then the error with its own
    // property; the exit status is the failure's (4, property family).
    let run = expect(4, add(&[]));
    let error = run.stderr.find("error[missing_property]").expect("error");
    let warnings = &run.stderr[..error];
    assert!(
        warnings.starts_with("warning[unsupported_property]:"),
        "{}",
        run.stderr
    );
    assert!(warnings.contains("  property: series1.trendline"));
    assert!(warnings.contains("  property: series1.categories"));
    assert!(
        run.stderr[error..].contains("  property: data"),
        "{}",
        run.stderr
    );
    assert_eq!(std::fs::read(&path).expect("read"), before);

    // JSON: OfficeCLI's ErrorResult with the skipped properties as warnings.
    let run = expect(4, add(&["--json"]));
    let error: Value = serde_json::from_str(run.stderr.trim()).expect("ErrorResult");
    assert_eq!(error["code"], "missing_property");
    assert_eq!(error["property"], "data");
    let warnings = error["warnings"].as_array().expect("warnings");
    assert_eq!(warnings.len(), 2);
    assert_eq!(warnings[0]["code"], "unsupported_property");
    assert_eq!(warnings[0]["property"], "series1.trendline");
    assert_eq!(warnings[1]["property"], "series1.categories");

    // A failure with nothing skipped carries no warnings.
    let run = expect(
        4,
        [
            "add",
            arg(&path),
            "/slide[1]",
            "--type",
            "chart",
            "--prop",
            "chartType=line",
            "--unsupported",
            "warn",
            "--json",
        ],
    );
    let error: Value = serde_json::from_str(run.stderr.trim()).expect("ErrorResult");
    assert!(error.get("warnings").is_none(), "{error}");

    // The corpus form of the call now applies (exit 0, no warnings).
    expect(
        0,
        [
            "add",
            arg(&path),
            "/slide[1]",
            "--type",
            "chart",
            "--prop",
            "chartType=area",
            "--prop",
            "series1.name=Web",
            "--prop",
            "series1.values=50,60,70",
            "--prop",
            "series1.color=4472C4",
            "--unsupported",
            "warn",
        ],
    );

    // An atomic batch item that fails after skipping reports them too.
    let run = expect(
        4,
        [
            "batch",
            arg(&path),
            "--unsupported",
            "warn",
            "--json",
            "--commands",
            r#"[{"command":"add","parent":"/slide[1]","type":"chart",
                 "props":{"chartType":"line","series1.trendline":"linear"}}]"#,
        ],
    );
    let report = run.json();
    assert_eq!(report["results"][0]["code"], "missing_property");
    assert_eq!(
        report["results"][0]["warnings"][0]["property"],
        "series1.trendline"
    );
}

/// `import`, `evaluate` and `recalculate` (record 0266).
#[test]
fn import_evaluate_and_recalculate() {
    let dir = tempfile::tempdir().expect("tempdir");
    let workbook = dir.path().join("book.xlsx");
    expect(0, ["create", arg(&workbook)]);
    let source = dir.path().join("data.tsv");
    std::fs::write(&source, "name\tscore\nAda\t36\n").expect("write source");
    let run = expect(
        0,
        [
            "--json",
            "import",
            arg(&workbook),
            "/Sheet1",
            arg(&source),
            "--header",
            "--start-cell",
            "B2",
        ],
    );
    let summary = run.json();
    assert_eq!(summary["origin"], "B2");
    assert_eq!(summary["rows"], 2);
    assert_eq!(summary["numbers"], 1);
    assert_eq!(summary["changed"], true);
    let file = litchi_ops::open(&workbook).expect("reopen");
    let cell = file.get("/Sheet1/C3").expect("imported cell");
    assert_eq!(cell.view().text().expect("text").as_deref(), Some("36"));

    // Empty input is a no-op that does not rewrite the file.
    let empty = dir.path().join("empty.csv");
    std::fs::write(&empty, "").expect("write empty source");
    let before = std::fs::read(&workbook).expect("read");
    let run = expect(0, ["import", arg(&workbook), "/Sheet1", arg(&empty)]);
    assert_eq!(run.stdout.trim(), "No data to import");
    assert_eq!(std::fs::read(&workbook).expect("read"), before);

    // Formulas are written only on request, and evaluate computes them.
    let formulas = dir.path().join("formulas.csv");
    std::fs::write(&formulas, "2,=A10*21\n").expect("write formulas");
    expect(
        0,
        [
            "import",
            arg(&workbook),
            "/Sheet1/A10",
            arg(&formulas),
            "--formulas",
        ],
    );
    let run = expect(0, ["--json", "evaluate", arg(&workbook), "/Sheet1/B10"]);
    assert_eq!(
        run.json(),
        json!({"path": "/Sheet1/B10", "value": 42.0, "origin": "computed"})
    );
    let run = expect(0, ["--json", "recalculate", arg(&workbook), "--all"]);
    let report = run.json();
    assert_eq!(report["formulas"], 1);
    assert_eq!(report["uncached"], 1);
    assert_eq!(report["cells"][0]["computedValue"], 42.0);

    // Usage and typed refusals.
    expect(
        2,
        ["import", arg(&workbook), "/Sheet1", arg(&source), "--stdin"],
    );
    let run = expect(
        4,
        [
            "import",
            arg(&workbook),
            "/Sheet1",
            arg(&source),
            "--delimiter",
            ";;",
        ],
    );
    assert!(run.stderr.contains("single character"), "{}", run.stderr);
}
