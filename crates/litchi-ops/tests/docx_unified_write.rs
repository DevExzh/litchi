//! Immutable DOCX `File::set` and publication coverage.
//!
//! The fixture deliberately contains more than the direct paragraph/run
//! surface.  The handler must preserve the source for the supported edits,
//! while refusing table, nested-table and markup-compatibility paths until it
//! owns the corresponding ordinal proof.
#![cfg(feature = "docx")]

use std::io::{self, Cursor, Read as _, Write as _};

use litchi_core::FileFormat;
use litchi_ops::capabilities::Capabilities;
use litchi_ops::query::{Plan, Query};
use litchi_ops::value::{Color, Length};
use litchi_ops::{Code, File, OpenOptions, PropValue, Props, SourceMode, props};

const WORD_MAIN: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

/// Build a package around one main-document part. Stored ZIP members make
/// exact no-op publication checks deterministic.
fn package(document: &str) -> Vec<u8> {
    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{WORD_MAIN}"/></Types>"#
    );
    let root_relationships = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;

    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, contents) in [
        ("[Content_Types].xml", content_types.as_str()),
        ("_rels/.rels", root_relationships),
        ("word/document.xml", document),
    ] {
        zip.start_file(name, options).expect("start fixture member");
        zip.write_all(contents.as_bytes())
            .expect("write fixture member");
    }
    zip.finish().expect("finish fixture package").into_inner()
}

fn zip_member_contents(bytes: &[u8]) -> Vec<(String, Vec<u8>)> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("open ZIP package");
    let mut members = Vec::with_capacity(archive.len());
    for index in 0..archive.len() {
        let mut member = archive.by_index(index).expect("read ZIP member");
        let name = member.name().expect("valid ZIP member name").into_owned();
        let mut contents = Vec::new();
        member
            .read_to_end(&mut contents)
            .expect("read ZIP member contents");
        members.push((name, contents));
    }
    members
}

fn assert_exact_bytes(label: &str, actual: &[u8], expected: &[u8]) {
    if actual == expected {
        return;
    }
    let first_difference = actual
        .iter()
        .zip(expected)
        .position(|(actual, expected)| actual != expected)
        .unwrap_or(actual.len().min(expected.len()));
    panic!(
        "{label}: exact bytes differ (actual_len={}, expected_len={}, first_difference={first_difference}, actual_byte={:?}, expected_byte={:?})",
        actual.len(),
        expected.len(),
        actual.get(first_difference),
        expected.get(first_difference),
    );
}

/// The positive-edit corpus has only direct body paragraphs.  Some DOCX
/// leaf transactions reject an owner containing MCE or nested table closure,
/// even when an unrelated direct paragraph is selected; keep that policy
/// separate from the direct paragraph/run proof.
fn fixture() -> Vec<u8> {
    package(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="11111111"><w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:b w:val="0"/><w:sz w:val="22"/><w:color w:val="000000"/></w:rPr><w:t xml:space="preserve">Original </w:t></w:r><w:r><w:t>paragraph</w:t></w:r></w:p><w:p w14:paraId="22222222"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Second</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#,
    )
}

/// A separate closure corpus for table, nested-table and MCE refusal tests.
fn complex_fixture() -> Vec<u8> {
    package(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:litchi:unsupported"><w:body><w:p w14:paraId="11111111"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Plain</w:t></w:r></w:p><w:p w14:paraId="33333333"><mc:AlternateContent><mc:Choice Requires="x"><w:r><w:t>choice</w:t></w:r></mc:Choice><mc:Fallback><w:r><w:t>fallback</w:t></w:r></mc:Fallback></mc:AlternateContent></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Table cell</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Nested cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>"#,
    )
}

fn open_source() -> File {
    litchi_ops::open_bytes(fixture()).expect("source-backed fixture")
}

fn open_eager() -> File {
    litchi_ops::open_bytes_with(
        fixture(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )
    .expect("eager fixture")
}

fn open_complex_source() -> File {
    litchi_ops::open_bytes(complex_fixture()).expect("complex source-backed fixture")
}

fn snapshot(file: &File, path: &str) -> (Option<String>, Vec<(String, String)>) {
    let node = file
        .get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"));
    let text = node
        .view()
        .text()
        .unwrap_or_else(|error| panic!("{path} text: {error}"))
        .map(|value| value.into_owned());
    let format = node
        .view()
        .format()
        .unwrap_or_else(|error| panic!("{path} format: {error}"))
        .map(|property| (property.key().to_owned(), property.to_string()))
        .collect();
    (text, format)
}

fn assert_code(error: litchi_ops::Error, code: Code) {
    assert_eq!(error.code(), code, "{error}");
}

/// Every writable paragraph/run registry row gets a non-default value.  This
/// keeps the capability report and the handler table coupled: adding a
/// writable row without a working conversion makes this test fail at the
/// named property instead of silently advertising an inert setter.
fn sample_props(kind: &str, key: &'static str) -> Props<'static> {
    if kind == "paragraph" {
        return match key {
            "text" => Props::new().with(key, "live paragraph"),
            // Record 0319: `styleId` and `styleName` write the style too.
            "style" | "styleId" | "styleName" => Props::new().with(key, "Title"),
            "align" => Props::new().with(key, "right"),
            "spaceBefore" | "spaceAfter" => Props::new().with(key, Length::from_points(6)),
            "lineSpacing" => Props::new().with(key, "1.15x"),
            "lineRule" => Props::new().with(key, "exact"),
            "spaceBeforeLines" | "spaceAfterLines" => Props::new().with(key, 1),
            "spaceBeforeAuto" | "spaceAfterAuto" => Props::new().with(key, true),
            "firstLineIndent" | "hangingIndent" | "indent" | "rightIndent" => {
                Props::new().with(key, Length::from_points(1))
            },
            "firstLineChars" | "hangingChars" | "leftChars" | "rightChars" => {
                Props::new().with(key, 1)
            },
            "keepNext" | "keepLines" | "pageBreakBefore" | "widowControl" => {
                Props::new().with(key, true)
            },
            "direction" => Props::new().with(key, "rtl"),
            "contextualSpacing" => Props::new().with(key, true),
            "numId" => Props::new().with(key, "9"),
            "numLevel" => Props::new().with(key, "1"),
            "outlineLvl" => Props::new().with(key, 1),
            // Record 0251e: both join the paragraph to a new list, visible
            // through its numId.
            "listStyle" => Props::new().with(key, "bullet"),
            "start" => Props::new().with(key, 3),
            _ => panic!("unhandled writable DOCX paragraph property {key}"),
        };
    }
    assert_eq!(
        kind, "run",
        "the liveness table only covers paragraph and run"
    );
    match key {
        "text" => Props::new().with(key, "live run"),
        "font.latin" | "font.ascii" | "font.hAnsi" | "font.ea" | "font.cs" => {
            Props::new().with(key, "Arial")
        },
        "font.asciiTheme" | "font.hAnsiTheme" | "font.eaTheme" | "font.csTheme" => {
            Props::new().with(key, "minorAscii")
        },
        "font.hint" => Props::new().with(key, "default"),
        "lang.latin" | "lang.ea" | "lang.cs" => Props::new().with(key, "en-US"),
        "size" | "size.cs" => Props::new().with(key, Length::from_points(14)),
        "bold" | "italic" | "bold.cs" | "italic.cs" | "strike" | "caps" | "smallcaps"
        | "dstrike" | "vanish" | "outline" | "shadow" | "emboss" | "imprint" | "noproof"
        | "superscript" | "subscript" => Props::new().with(key, true),
        "color" | "underline.color" => Props::new().with(key, Color::rgb(0x1F, 0x4E, 0x79)),
        "underline" => Props::new().with(key, "single"),
        "highlight" => Props::new().with(key, "yellow"),
        "direction" => Props::new().with(key, "rtl"),
        "position" | "charSpacing" => Props::new().with(key, Length::from_points(1)),
        "rStyle" => Props::new().with(key, "Emphasis"),
        _ => panic!("unhandled writable DOCX run property {key}"),
    }
}

#[test]
fn paragraph_and_run_sets_are_immutable_chainable_and_match_source_modes() {
    let mut results = Vec::new();
    let openers: [fn() -> File; 2] = [open_source, open_eager];
    for open in openers {
        let original = open();
        let original_paragraph = snapshot(&original, "/body/p[1]");
        let original_run = snapshot(&original, "/body/p[1]/r[1]");

        let paragraph = original
            .set(
                "/body/p[1]",
                props! {
                    "text" => "Revised paragraph",
                    "align" => "center",
                },
            )
            .expect("paragraph text and alignment");
        assert_eq!(snapshot(&original, "/body/p[1]"), original_paragraph);
        assert_eq!(
            snapshot(&paragraph, "/body/p[1]").0.as_deref(),
            Some("Revised paragraph")
        );
        assert!(
            snapshot(&paragraph, "/body/p[1]")
                .1
                .iter()
                .any(|(key, value)| key == "align" && value == "center")
        );
        let paragraph_run = snapshot(&paragraph, "/body/p[1]/r[1]");

        let edited = paragraph
            .set(
                "/body/p[1]/r[1]",
                props! {
                    "text" => "First run",
                    "bold" => true,
                    "size" => Length::from_points(14),
                    "color" => Color::rgb(0x1F, 0x4E, 0x79),
                },
            )
            .expect("run text and direct formatting");
        assert_eq!(snapshot(&original, "/body/p[1]/r[1]"), original_run);
        assert_eq!(
            snapshot(&edited, "/body/p[1]/r[1]").0.as_deref(),
            Some("First run")
        );
        let run_format = snapshot(&edited, "/body/p[1]/r[1]").1;
        for (key, wanted) in [("bold", "true"), ("size", "14pt"), ("color", "#1F4E79")] {
            assert!(
                run_format
                    .iter()
                    .any(|(actual, value)| actual == key && value == wanted),
                "missing {key}={wanted} in {run_format:?}"
            );
        }

        let chained = edited
            .set("/body/p[1]/r[1]", props! { "text" => "second run" })
            .expect("a second immutable edit");
        assert_eq!(
            snapshot(&chained, "/body/p[1]/r[1]").0.as_deref(),
            Some("second run")
        );
        assert_eq!(snapshot(&paragraph, "/body/p[1]/r[1]"), paragraph_run);
        results.push((
            snapshot(&chained, "/body/p[1]"),
            snapshot(&chained, "/body/p[1]/r[1]"),
        ));
    }

    assert_eq!(
        results[0], results[1],
        "source and eager semantic snapshots"
    );
}

#[test]
fn every_advertised_paragraph_and_run_setter_is_live_after_reopen() {
    let capabilities = Capabilities::of(FileFormat::Docx);
    let openers: [(&str, fn() -> File); 2] = [("source", open_source), ("eager", open_eager)];
    for (mode, open) in openers {
        for (kind, path) in [("paragraph", "/body/p[1]"), ("run", "/body/p[1]/r[1]")] {
            let element = capabilities
                .element(kind)
                .unwrap_or_else(|| panic!("missing DOCX {kind} capability"));
            let writable: Vec<_> = element
                .properties()
                .filter(|property| property.set())
                .collect();
            assert!(
                !writable.is_empty(),
                "DOCX {kind} has no writable properties"
            );
            for property in writable {
                let key = property.name();
                let original = open();
                let before = snapshot(&original, path);
                let updated = original
                    .set(path, sample_props(kind, key))
                    .unwrap_or_else(|error| panic!("{mode} {kind} {key}: set failed: {error}"));
                let after = snapshot(&updated, path);
                assert_ne!(before, after, "{mode} {kind} {key}: setter was inert");

                let bytes = updated.to_bytes().unwrap_or_else(|error| {
                    panic!("{mode} {kind} {key}: publication failed: {error}")
                });
                let reopened = litchi_ops::open_bytes(bytes)
                    .unwrap_or_else(|error| panic!("{mode} {kind} {key}: reopen failed: {error}"));
                assert_eq!(
                    snapshot(&reopened, path),
                    after,
                    "{mode} {kind} {key}: readback"
                );
            }
        }
    }
}

#[test]
fn borrowed_and_compiled_targets_keep_the_mutation_scope_guard() {
    let original = open_source();

    let bare_query = Query::parse("p").expect("bare query");
    let borrowed_error = original
        .set(&bare_query, props! { "align" => "center" })
        .expect_err("a borrowed unscoped query must be guarded");
    assert_code(borrowed_error, Code::BareSelectorRejected);

    let bare_query = Query::parse("p").expect("bare query for plan");
    let bare_plan = Plan::compile(&bare_query, FileFormat::Docx).expect("compile bare plan");
    let typed_error = original
        .set(bare_plan, props! { "align" => "center" })
        .expect_err("a compiled unscoped plan must be guarded");
    assert_code(typed_error, Code::BareSelectorRejected);

    let scoped_query = Query::parse("/body/p[style=Heading1]").expect("scoped query");
    let scoped_plan = Plan::compile(&scoped_query, FileFormat::Docx).expect("compile scoped plan");
    let edited = original
        .set(scoped_plan, props! { "align" => "right" })
        .expect("a scoped compiled plan is a valid mutation target");
    for path in ["/body/p[1]", "/body/p[2]"] {
        assert!(
            snapshot(&edited, path)
                .1
                .iter()
                .any(|(key, value)| key == "align" && value == "right"),
            "{path}"
        );
    }
}

#[test]
fn unknown_predicate_negation_is_diagnostic_and_mutation_is_atomic() {
    let original = open_source();
    let mut matches = original
        .query("/body/p[not(unknown=secret)]")
        .expect("unknown predicate query");
    assert_eq!(matches.diagnostics().len(), 1);
    let paths: Vec<_> = matches
        .by_ref()
        .map(|node| node.expect("query node").path().to_string())
        .collect();
    assert_eq!(
        paths,
        vec!["/body/p[1]".to_owned(), "/body/p[2]".to_owned()]
    );

    let before_paragraph = snapshot(&original, "/body/p[1]");
    let error = original
        .set(
            "/body/p[not(unknown=secret)]",
            props! { "align" => "center" },
        )
        .expect_err("mutation selectors must reject unknown predicates");
    assert_code(error, Code::InvalidSelector);
    assert_eq!(snapshot(&original, "/body/p[1]"), before_paragraph);
}

#[test]
fn mixed_paragraph_and_run_query_refuses_before_any_assignment() {
    let original = open_source();
    let mixed = original
        .query("/body/p, r")
        .expect("scoped mixed-kind query");
    let paths: Vec<_> = mixed
        .map(|node| node.expect("mixed query node").path().to_string())
        .collect();
    assert!(paths.iter().any(|path| path == "/body/p[1]"));
    assert!(paths.iter().any(|path| path == "/body/p[1]/r[1]"));

    let before_paragraph = snapshot(&original, "/body/p[1]");
    let before_run = snapshot(&original, "/body/p[1]/r[1]");
    let error = original
        .set("/body/p, r", props! { "align" => "center" })
        .expect_err("one mutation cannot mix paragraph and run targets");
    assert_code(error, Code::UnsupportedCapability);
    assert_eq!(snapshot(&original, "/body/p[1]"), before_paragraph);
    assert_eq!(snapshot(&original, "/body/p[1]/r[1]"), before_run);
}

#[test]
fn null_clears_text_and_direct_format_without_missing_model_panics() {
    let openers: [(&str, fn() -> File); 2] = [("source", open_source), ("eager", open_eager)];
    for (mode, open) in openers {
        let original = open();
        let paragraph = original
            .set("/body/p[1]", Props::new().with("text", PropValue::Null))
            .unwrap_or_else(|error| panic!("{mode} paragraph text clear: {error}"));
        assert_eq!(
            snapshot(&paragraph, "/body/p[1]").0.as_deref(),
            Some(""),
            "{mode} paragraph text clear"
        );

        let run = original
            .set(
                "/body/p[1]/r[1]",
                Props::new().with("text", PropValue::Null),
            )
            .unwrap_or_else(|error| panic!("{mode} run text clear: {error}"));
        assert_eq!(
            snapshot(&run, "/body/p[1]/r[1]").0.as_deref(),
            Some(""),
            "{mode} run text clear"
        );

        let formatted = original
            .set(
                "/body/p[1]/r[1]",
                props! {
                    "bold" => true,
                    "color" => Color::rgb(0x1F, 0x4E, 0x79),
                },
            )
            .unwrap_or_else(|error| panic!("{mode} format setup: {error}"));
        let cleared = formatted
            .set(
                "/body/p[1]/r[1]",
                Props::new()
                    .with("bold", PropValue::Null)
                    .with("color", PropValue::Null),
            )
            .unwrap_or_else(|error| panic!("{mode} format clear: {error}"));
        let format = snapshot(&cleared, "/body/p[1]/r[1]").1;
        assert!(
            !format
                .iter()
                .any(|(key, value)| key == "bold" && value == "true")
        );
        assert!(
            !format
                .iter()
                .any(|(key, value)| key == "color" && value == "#1F4E79")
        );
    }
}

#[test]
fn duplicate_assignments_are_applied_in_order_with_last_value_winning() {
    let original = open_source();
    let edited = original
        .set(
            "/body/p[1]/r[1]",
            Props::new()
                .with("bold", true)
                .with("BOLD", false)
                .with("text", "first")
                .with("TEXT", "last"),
        )
        .expect("ordered duplicate assignments");
    assert_eq!(
        snapshot(&edited, "/body/p[1]/r[1]").0.as_deref(),
        Some("last")
    );
    assert!(
        !snapshot(&edited, "/body/p[1]/r[1]")
            .1
            .iter()
            .any(|(key, value)| key == "bold" && value == "true")
    );
}

#[test]
fn scoped_query_fans_out_and_bare_selector_is_rejected() {
    let original = open_source();
    let rejected = original
        .set("p", props! { "align" => "right" })
        .expect_err("a bare mutation selector must be guarded");
    assert_code(rejected, Code::BareSelectorRejected);

    let edited = original
        .set("/body/p[style=Heading1]", props! { "align" => "center" })
        .expect("a scoped selector mutation");
    for path in ["/body/p[1]", "/body/p[2]"] {
        assert!(
            snapshot(&edited, path)
                .1
                .iter()
                .any(|(key, value)| key == "align" && value == "center"),
            "{path}"
        );
    }
    assert!(
        !snapshot(&original, "/body/p[1]")
            .1
            .iter()
            .any(|(key, value)| key == "align" && value == "center")
    );
    assert_eq!(snapshot(&edited, "/body/p[2]").0.as_deref(), Some("Second"));
}

#[test]
fn text_predicates_use_virtual_text_for_read_and_mutation() {
    let original = open_source();
    let exact = original
        .query("/body/p[text='Original paragraph']")
        .expect("exact paragraph text query");
    let exact_paths: Vec<_> = exact
        .map(|node| node.expect("exact text node").path().to_string())
        .collect();
    assert_eq!(exact_paths, vec!["/body/p[1]".to_owned()]);

    let aggregate = original
        .query("p[text='Original paragraph']")
        .expect("whole-body text query");
    let aggregate_paths: Vec<_> = aggregate
        .map(|node| node.expect("aggregate body node").path().to_string())
        .collect();
    assert_eq!(aggregate_paths, vec!["/body/p[1]".to_owned()]);
    assert_eq!(
        original.text().expect("whole-document text"),
        "Original paragraph\nSecond"
    );

    let edited = original
        .set(
            "/body/p[text='Original paragraph']",
            props! { "align" => "center" },
        )
        .expect("text predicate mutation");
    assert!(
        snapshot(&edited, "/body/p[1]")
            .1
            .iter()
            .any(|(key, value)| key == "align" && value == "center")
    );
    assert!(
        !snapshot(&edited, "/body/p[2]")
            .1
            .iter()
            .any(|(key, value)| key == "align" && value == "center")
    );
    assert_eq!(
        edited.text().expect("whole-document text after mutation"),
        "Original paragraph\nSecond"
    );
}

#[test]
fn unknown_and_invalid_properties_fail_before_any_assignment() {
    let original = open_source();
    let before = snapshot(&original, "/body/p[1]/r[1]");

    let unknown = original
        .set(
            "/body/p[1]/r[1]",
            props! {
                "bold" => true,
                "boldd" => true,
            },
        )
        .expect_err("unknown property");
    assert_code(unknown, Code::UnsupportedProperty);
    assert_eq!(snapshot(&original, "/body/p[1]/r[1]"), before);

    let invalid = original
        .set(
            "/body/p[1]/r[1]",
            props! {
                "bold" => "not a boolean",
            },
        )
        .expect_err("invalid typed value");
    assert_code(invalid, Code::InvalidValue);
    assert_eq!(snapshot(&original, "/body/p[1]/r[1]"), before);
}

#[test]
fn table_nested_and_mce_paths_refuse_without_partial_changes() {
    let original = open_complex_source();
    let before = snapshot(&original, "/body/p[1]");
    for path in [
        "/body/tbl[1]/tr[1]/tc[1]/p[1]",
        "/body/tbl[1]/tr[1]/tc[1]/tbl[1]/tr[1]/tc[1]/p[1]",
        "/body/p[2]/r[1]",
    ] {
        let error = original
            .set(path, props! { "text" => "must be refused" })
            .expect_err(path);
        assert_eq!(error.code(), Code::UnsupportedCapability, "{path}: {error}");
        assert_eq!(snapshot(&original, "/body/p[1]"), before, "{path}");
    }
}

#[test]
fn signed_unretained_and_stale_sources_are_refused_with_typed_errors() {
    let signed = litchi_ops::open(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/xmldsign/hello-world-signed.docx"
    ))
    .expect("signed fixture");
    let signed_error = signed
        .set("/body/p[1]", props! { "text" => "no" })
        .expect_err("signed documents cannot be edited");
    assert_code(signed_error, Code::DocumentProtected);

    let unretained =
        litchi_ops::open_bytes_with(fixture(), &OpenOptions::new().retain_source(false))
            .expect("unretained fixture");
    let retention_error = unretained
        .set("/body/p[1]", props! { "text" => "no" })
        .expect_err("editing needs a retained source");
    assert_code(retention_error, Code::UnsupportedCapability);

    let directory = tempfile::tempdir().expect("temporary directory");
    let path = directory.path().join("stale.docx");
    std::fs::write(&path, fixture()).expect("write source");
    let stale = litchi_ops::open(&path).expect("path-backed fixture");
    std::fs::write(&path, b"source changed after open").expect("change source");
    let stale_error = stale
        .set("/body/p[1]", props! { "text" => "no" })
        .expect_err("stale source");
    assert_code(stale_error, Code::IoError);
}

#[test]
fn empty_assignments_keep_security_policy_and_signed_noops_are_retained() {
    let original = open_source();
    let unchanged = original
        .set("/body/p[1]", Props::new())
        .expect("an empty assignment list is a no-op on a writable handle");
    assert_eq!(
        snapshot(&unchanged, "/body/p[1]"),
        snapshot(&original, "/body/p[1]")
    );

    let signed_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/xmldsign/hello-world-signed.docx"
    );
    let signed = litchi_ops::open(signed_path).expect("signed fixture");
    let signed_bytes = signed.to_bytes().expect("signed source bytes");
    let signed_noop = signed
        .set("/body/p[1]", Props::new())
        .expect("an empty assignment is a signed exact no-op");
    assert!(signed_noop.is_signed());
    assert_eq!(
        signed_noop.to_bytes().expect("signed no-op bytes"),
        signed_bytes
    );
    let invalid_target = litchi_ops::open(signed_path).expect("signed fixture for target check");
    let target_error = invalid_target
        .set("/body/p[999]", Props::new())
        .expect_err("an invalid target must not be hidden by signed no-op handling");
    assert_code(target_error, Code::PathNotFound);

    let unretained =
        litchi_ops::open_bytes_with(fixture(), &OpenOptions::new().retain_source(false))
            .expect("unretained fixture");
    let retention_error = unretained
        .set("/body/p[1]", Props::new())
        .expect_err("an empty assignment must not bypass source retention policy");
    assert_code(retention_error, Code::UnsupportedCapability);
}

#[test]
fn signed_noop_is_exact_and_explicit_unsign_removes_signature_on_write() {
    let signed_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/xmldsign/hello-world-signed.docx"
    );
    let original_bytes = std::fs::read(signed_path).expect("read signed fixture");
    let directory = tempfile::tempdir().expect("signed temporary directory");
    let path = directory.path().join("signed.docx");
    std::fs::write(&path, &original_bytes).expect("write signed fixture");
    let path_eager =
        litchi_ops::open_with(&path, &OpenOptions::new().source_mode(SourceMode::Eager))
            .expect("file-backed eager signed fixture");
    let noops = [
        (
            "bytes source",
            litchi_ops::open_bytes(original_bytes.clone()).expect("bytes source signed fixture"),
        ),
        (
            "bytes eager",
            litchi_ops::open_bytes_with(
                original_bytes.clone(),
                &OpenOptions::new().source_mode(SourceMode::Eager),
            )
            .expect("bytes eager signed fixture"),
        ),
        (
            "path source",
            litchi_ops::open(&path).expect("file-backed source signed fixture"),
        ),
        ("path eager", path_eager.clone()),
    ];
    for (label, signed) in noops {
        assert!(signed.is_signed(), "{label}");
        assert_exact_bytes(
            &format!("{label} signed no-op publication"),
            &signed.to_bytes().expect("signed no-op bytes"),
            &original_bytes,
        );
    }

    let signed = litchi_ops::open(&path).expect("signed fixture for explicit unsign");
    let unsigned = match signed {
        File::Signed(signed) => signed.unsign(),
        other => panic!("expected a signed file, got {other:?}"),
    };
    assert!(!unsigned.is_signed());
    let edited = unsigned
        .set("/body/p[1]", props! { "text" => "explicitly unsigned" })
        .expect("explicitly unsigned edit");
    let bytes = edited.to_bytes().expect("publish explicitly unsigned edit");
    assert!(
        zip_member_contents(&bytes)
            .iter()
            .all(|(name, _)| !name.starts_with("_xmlsignatures/")),
        "signature parts remain after explicit unsign"
    );
    let reopened = litchi_ops::open_bytes(bytes).expect("reopen unsigned edit");
    assert!(!reopened.is_signed());
    assert_eq!(
        snapshot(&reopened, "/body/p[1]").0.as_deref(),
        Some("explicitly unsigned")
    );

    let eager_unsigned = match path_eager {
        File::Signed(signed) => signed.unsign(),
        other => panic!("expected a file-backed eager signed file, got {other:?}"),
    };
    let eager_edit = eager_unsigned
        .set("/body/p[1]", props! { "text" => "eager unsigned" })
        .expect("file-backed eager explicitly unsigned edit");
    let eager_bytes = eager_edit
        .to_bytes()
        .expect("publish file-backed eager edit");
    let original_members = zip_member_contents(&original_bytes);
    let eager_members = zip_member_contents(&eager_bytes);
    for member in ["word/styles.xml", "word/theme/theme1.xml"] {
        let original_xml = original_members
            .iter()
            .find(|(name, _)| name == member)
            .map(|(_, contents)| contents);
        let eager_xml = eager_members
            .iter()
            .find(|(name, _)| name == member)
            .map(|(_, contents)| contents);
        assert_eq!(eager_xml, original_xml, "untouched member {member}");
    }
    let eager_reopened = litchi_ops::open_bytes(eager_bytes).expect("reopen eager unsigned edit");
    assert!(!eager_reopened.is_signed());
    assert_eq!(
        snapshot(&eager_reopened, "/body/p[1]").0.as_deref(),
        Some("eager unsigned")
    );
}

#[test]
fn publication_preserves_noop_bytes_matches_owner_modes_and_reports_sink_errors() {
    let original_bytes = fixture();
    let source = open_source();
    let eager = open_eager();
    assert_eq!(
        source.to_bytes().expect("source no-op bytes"),
        original_bytes
    );
    let eager_noop = eager.to_bytes().expect("eager no-op bytes");
    assert_eq!(
        zip_member_contents(&eager_noop),
        zip_member_contents(&original_bytes),
        "eager no-op changed ZIP member contents"
    );
    assert_exact_bytes(
        "eager no-op publication must preserve exact package bytes",
        &eager_noop,
        &original_bytes,
    );
    let eager_reopened = litchi_ops::open_bytes(eager_noop).expect("reopen eager no-op");
    assert_eq!(
        snapshot(&eager_reopened, "/body/p[1]"),
        snapshot(&eager, "/body/p[1]"),
        "eager no-op semantic preservation"
    );

    let source_edit = source
        .set(
            "/body/p[1]/r[1]",
            props! { "text" => "published", "bold" => true },
        )
        .expect("source edit");
    let eager_edit = eager
        .set(
            "/body/p[1]/r[1]",
            props! { "text" => "published", "bold" => true },
        )
        .expect("eager edit");
    let source_bytes = source_edit.to_bytes().expect("source publication");
    let eager_bytes = eager_edit.to_bytes().expect("eager publication");
    let reopened = litchi_ops::open_bytes(source_bytes).expect("reopen publication");
    let eager_reopened = litchi_ops::open_bytes(eager_bytes).expect("reopen eager publication");
    assert_eq!(
        snapshot(&reopened, "/body/p[1]/r[1]"),
        snapshot(&eager_reopened, "/body/p[1]/r[1]"),
        "source/eager semantic publication parity"
    );
    assert_eq!(
        snapshot(&reopened, "/body/p[1]/r[1]").0.as_deref(),
        Some("published")
    );

    let mut sink = Vec::new();
    let written = source_edit.write_to(&mut sink).expect("write publication");
    assert_eq!(written as usize, sink.len());

    struct FailingSink;
    impl io::Write for FailingSink {
        fn write(&mut self, _: &[u8]) -> io::Result<usize> {
            Err(io::Error::new(io::ErrorKind::BrokenPipe, "sink refused"))
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }
    let sink_error = source_edit
        .write_to(FailingSink)
        .expect_err("a failing publication sink");
    assert_code(sink_error, Code::IoError);
    assert_eq!(
        snapshot(&source_edit, "/body/p[1]/r[1]").0.as_deref(),
        Some("published")
    );
}

#[test]
fn same_value_text_and_bool_sets_preserve_exact_output_in_both_owner_modes() {
    let original_bytes = fixture();
    let openers: [(&str, fn() -> File); 2] = [("source", open_source), ("eager", open_eager)];
    for (mode, open) in openers {
        let updated = open()
            .set(
                "/body/p[1]/r[1]",
                props! {
                    "text" => "Original ",
                    "bold" => false,
                },
            )
            .unwrap_or_else(|error| panic!("{mode} same-value set: {error}"));
        let bytes = updated
            .to_bytes()
            .unwrap_or_else(|error| panic!("{mode} same-value publication: {error}"));
        assert_exact_bytes(
            &format!("{mode} same-value publication"),
            &bytes,
            &original_bytes,
        );
    }
}

#[cfg(any(unix, windows))]
#[test]
fn prepared_eager_ingress_preserves_exact_noop_and_untouched_members() {
    let bytes = fixture();
    let directory = tempfile::tempdir().expect("prepared temporary directory");
    let path = directory.path().join("prepared.docx");
    std::fs::write(&path, &bytes).expect("write prepared fixture");
    let prepared =
        litchi_ops::detect::detect_and_prepare(&path, &litchi_ops::detect::SourceLimits::default())
            .expect("prepare fixture");
    let eager = litchi_ops::open_prepared_with(
        prepared,
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )
    .expect("adopt prepared eager fixture");
    std::fs::remove_file(&path).expect("eager owner no longer needs the input file");
    assert_exact_bytes(
        "prepared eager no-op",
        &eager.to_bytes().expect("publish no-op"),
        &bytes,
    );
    let edited = eager
        .set("/body/p[1]/r[1]", props! { "text" => "prepared edit" })
        .expect("edit prepared fixture");
    let output = edited.to_bytes().expect("publish prepared edit");
    let original_members = zip_member_contents(&bytes);
    let edited_members = zip_member_contents(&output);
    for (name, payload) in original_members {
        if name != "word/document.xml" {
            assert_eq!(
                edited_members
                    .iter()
                    .find(|(member, _)| member == &name)
                    .map(|(_, bytes)| bytes),
                Some(&payload),
                "untouched member {name}"
            );
        }
    }
    assert_eq!(
        snapshot(&eager, "/body/p[1]").0.as_deref(),
        Some("Original paragraph")
    );
}

#[test]
fn inherited_word_alias_text_is_visible_to_nodes_and_query_predicates() {
    let bytes = package(
        r#"<a:document xmlns:a="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:b="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:f="urn:foreign"><a:body><a:p><a:r><b:t>visible</b:t><f:t>hidden</f:t><b:t xmlns:b="urn:foreign">shadowed</b:t></a:r></a:p></a:body></a:document>"#,
    );
    for file in [
        litchi_ops::open_bytes(bytes.as_slice()).unwrap(),
        litchi_ops::open_bytes_with(
            bytes.as_slice(),
            &OpenOptions::new().source_mode(SourceMode::Eager),
        )
        .unwrap(),
    ] {
        for path in ["/body/p[1]", "/body/p[1]/r[1]"] {
            assert_eq!(
                file.get(path).unwrap().view().text().unwrap().as_deref(),
                Some("visible")
            );
        }
        let paths: Vec<_> = file
            .query("p[text='visible']")
            .unwrap()
            .map(|node| node.unwrap().path().to_string())
            .collect();
        assert_eq!(paths, ["/body/p[1]"]);
        assert_eq!(file.query("r[text='visible']").unwrap().count(), 1);
        assert_eq!(file.query("p[text='hidden']").unwrap().count(), 0);
    }
}

#[test]
fn paragraph_text_fanout_late_leaf_refusal_keeps_the_original_archive() {
    let bytes = package(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:hyperlink w:anchor="bookmark"><w:r><w:t>complex</w:t></w:r></w:hyperlink></w:p></w:body></w:document>"#,
    );
    for mode in [SourceMode::Auto, SourceMode::Eager] {
        let file =
            litchi_ops::open_bytes_with(bytes.as_slice(), &OpenOptions::new().source_mode(mode))
                .unwrap();
        let error = file
            .set(
                Query::parse("/body/p").unwrap(),
                props! { "text" => "replacement\nline" },
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert_eq!(error.property(), Some("text"));
        assert_eq!(file.to_bytes().unwrap(), bytes);
    }
}
