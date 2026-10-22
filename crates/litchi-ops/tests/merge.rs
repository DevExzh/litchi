//! Template merge through `File::merge` (record 0274): DOCX placeholders
//! across runs, stories and `MERGEFIELD` fields, PPTX runs and XLSX cells;
//! one atomic commit, exact no-ops, round trips, typed refusals and limits.
#![cfg(all(
    feature = "docx",
    feature = "pptx",
    feature = "xlsx",
    feature = "xlsb",
    feature = "xls"
))]

use std::io::{Cursor, Read as _, Write as _};

use litchi_core::FileFormat;
use litchi_ops::capabilities::{Capabilities, Verb};
use litchi_ops::merge::{Data, Missing, Options, Record};
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{Code, File, NewNode, OpenOptions, SourceMode, props};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WORD_MAIN: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const WORD_HEADER: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";

/// A DOCX package of `body` and, when given, a default header.
fn docx(body: &str, header: Option<&str>) -> Vec<u8> {
    let header_override = if header.is_some() {
        format!(r#"<Override PartName="/word/header1.xml" ContentType="{WORD_HEADER}"/>"#)
    } else {
        String::new()
    };
    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{WORD_MAIN}"/>{header_override}</Types>"#
    );
    let root = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    let (relationships, reference) = if header.is_some() {
        (
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>"#,
            r#"<w:headerReference w:type="default" r:id="rIdH1"/>"#,
        )
    } else {
        (
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#,
            "",
        )
    };
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>{body}<w:sectPr>{reference}<w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#
    );
    let mut members = vec![
        ("[Content_Types].xml", content_types),
        ("_rels/.rels", root.to_owned()),
        ("word/_rels/document.xml.rels", relationships.to_owned()),
        ("word/document.xml", document),
    ];
    if let Some(header) = header {
        members.push((
            "word/header1.xml",
            format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="{W}" xmlns:r="{R}">{header}</w:hdr>"#),
        ));
    }
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, contents) in members {
        zip.start_file(name, options).expect("member");
        zip.write_all(contents.as_bytes()).expect("write member");
    }
    zip.finish().expect("finish").into_inner()
}

fn member(bytes: &[u8], name: &str) -> String {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("zip");
    let mut entry = archive.by_name(name).expect("member");
    let mut text = String::new();
    entry.read_to_string(&mut text).expect("UTF-8 member");
    text
}

/// Replace one member of a package, keeping the others byte for byte.
fn rewrite(bytes: &[u8], name: &str, change: impl Fn(&str) -> String) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("zip");
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).expect("entry");
        let entry_name = entry.name().to_owned();
        let mut contents = Vec::new();
        entry.read_to_end(&mut contents).expect("read entry");
        if entry_name == name {
            contents = change(std::str::from_utf8(&contents).expect("UTF-8")).into_bytes();
        }
        zip.start_file(entry_name, options).expect("member");
        zip.write_all(&contents).expect("write member");
    }
    zip.finish().expect("finish").into_inner()
}

fn open(bytes: Vec<u8>) -> File {
    litchi_ops::open_bytes(bytes).expect("open")
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .text()
        .unwrap_or_else(|error| panic!("{path} text: {error}"))
        .map(std::borrow::Cow::into_owned)
        .unwrap_or_default()
}

fn property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .format()
        .unwrap_or_else(|error| panic!("{path} format: {error}"))
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn record(pairs: &[(&str, &str)]) -> Record {
    pairs.iter().copied().collect()
}

#[test]
fn docx_placeholders_across_runs_keep_each_runs_formatting() {
    for mode in [SourceMode::Auto, SourceMode::Eager] {
        let bytes = docx(
            r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Dear {{na</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">me}}, welcome to {{ city }}.</w:t></w:r></w:p>"#,
            None,
        );
        let template =
            litchi_ops::open_bytes_with(bytes, &OpenOptions::new().source_mode(mode)).unwrap();
        let commit = template
            .merge(
                &record(&[("name", "Ada"), ("city", "London")]),
                &Options::new(),
            )
            .unwrap();
        let file = commit.file();
        assert_eq!(text(file, "/body/p[1]"), "Dear Ada, welcome to London.");
        assert_eq!(text(file, "/body/p[1]/r[1]"), "Dear Ada");
        assert_eq!(text(file, "/body/p[1]/r[2]"), ", welcome to London.");
        assert_eq!(
            property(file, "/body/p[1]/r[1]", "bold").as_deref(),
            Some("true")
        );
        assert_eq!(
            property(file, "/body/p[1]/r[2]", "italic").as_deref(),
            Some("true")
        );
        let summary = commit.value();
        assert_eq!(summary.replacements(), 2);
        assert_eq!(summary.keys(), ["city", "name"]);
        assert!(summary.unresolved().is_empty());
        assert_eq!(commit.diagnostics().commits(), 1);
        // The template is unchanged, and the result survives a round trip.
        assert_eq!(
            text(&template, "/body/p[1]"),
            "Dear {{name}}, welcome to {{ city }}."
        );
        let reopened = open(file.to_bytes().unwrap());
        assert_eq!(text(&reopened, "/body/p[1]/r[1]"), "Dear Ada");
        // The patch inverts to the template.
        let restored = commit.patch().inverse().apply(file).unwrap();
        assert_eq!(text(&restored, "/body/p[1]"), text(&template, "/body/p[1]"));
    }
}

#[test]
fn docx_merges_tables_controls_and_headers() {
    let bytes = docx(
        concat!(
            r#"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Total: {{total}}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
            r#"<w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>{{client.name}}</w:t></w:r></w:p></w:sdtContent></w:sdt>"#,
            r#"<w:p><w:r><w:t>Plain</w:t></w:r></w:p>"#,
        ),
        Some(r#"<w:p><w:r><w:t>{{client.name}} &amp; Co</w:t></w:r></w:p>"#),
    );
    let template = open(bytes);
    let data = Data::from_json(
        br#"{"total": 5200.00, "client": {"name": "Acme <EU>"}}"#,
        &SemanticLimits::default(),
    )
    .unwrap();
    let commit = template.merge(&data.records()[0], &Options::new()).unwrap();
    assert_eq!(commit.value().replacements(), 3);
    let published = commit.file().to_bytes().unwrap();
    let document = member(&published, "word/document.xml");
    assert!(document.contains("<w:t>Total: 5200.00</w:t>"), "{document}");
    assert!(
        document.contains("<w:t>Acme &lt;EU&gt;</w:t>"),
        "{document}"
    );
    let header = member(&published, "word/header1.xml");
    assert!(
        header.contains("<w:t>Acme &lt;EU&gt; &amp; Co</w:t>"),
        "{header}"
    );
    // A header-only change is a change.
    let header_only = open(docx(
        r#"<w:p><w:r><w:t>Body</w:t></w:r></w:p>"#,
        Some(r#"<w:p><w:r><w:t>{{x}}</w:t></w:r></w:p>"#),
    ));
    let commit = header_only
        .merge(&record(&[("x", "y")]), &Options::new())
        .unwrap();
    assert!(commit.diagnostics().changed());
    assert!(!commit.patch().is_noop());
    let header = member(&commit.file().to_bytes().unwrap(), "word/header1.xml");
    assert!(header.contains("<w:t>y</w:t>"), "{header}");
}

const MERGE_FIELDS: &str = concat!(
    r#"<w:p><w:r><w:t xml:space="preserve">To </w:t></w:r>"#,
    r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
    r#"<w:r><w:instrText xml:space="preserve"> MERGEFIELD First \* Upper \b "Mx " </w:instrText></w:r>"#,
    r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
    r#"<w:r><w:rPr><w:b/></w:rPr><w:t>«First»</w:t></w:r>"#,
    r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
    r#"<w:r><w:t xml:space="preserve"> </w:t></w:r>"#,
    r#"<w:fldSimple w:instr=" MERGEFIELD &quot;Last Name&quot; "><w:r><w:rPr><w:i/></w:rPr><w:t>«Last Name»</w:t></w:r></w:fldSimple>"#,
    r#"<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>"#,
);

#[test]
fn docx_merge_fields_become_formatted_runs() {
    let template = open(docx(MERGE_FIELDS, None));
    let commit = template
        .merge(
            &record(&[("first", "ada"), ("Last Name", "Lovelace")]),
            &Options::new(),
        )
        .unwrap();
    assert_eq!(commit.value().fields(), 2);
    assert_eq!(commit.value().keys(), ["Last Name", "first"]);
    let file = commit.file();
    assert_eq!(text(file, "/body/p[1]"), "To Mx ADA Lovelace1");
    let document = member(&file.to_bytes().unwrap(), "word/document.xml");
    assert!(
        document.contains(r#"<w:r><w:rPr><w:b/></w:rPr><w:t>Mx ADA</w:t></w:r>"#),
        "{document}"
    );
    assert!(
        document.contains(r#"<w:r><w:rPr><w:i/></w:rPr><w:t>Lovelace</w:t></w:r>"#),
        "{document}"
    );
    assert!(document.contains("PAGE"), "other fields stay");
    assert!(!document.contains("MERGEFIELD"));
    // Fields off: nothing changes.
    let off = template
        .merge(&record(&[("First", "x")]), &Options::new().fields(false))
        .unwrap();
    assert!(off.patch().is_noop());
}

#[test]
fn missing_keys_are_kept_or_emptied() {
    let template = open(docx(
        &format!(r#"<w:p><w:r><w:t>[{{{{gone}}}}]</w:t></w:r></w:p>{MERGE_FIELDS}"#),
        None,
    ));
    let kept = template.merge(&Record::new(), &Options::new()).unwrap();
    assert!(kept.patch().is_noop());
    assert_eq!(kept.value().unresolved(), ["First", "Last Name", "gone"]);
    let emptied = template
        .merge(&Record::new(), &Options::new().missing(Missing::Empty))
        .unwrap();
    assert_eq!(text(emptied.file(), "/body/p[1]"), "[]");
    assert_eq!(text(emptied.file(), "/body/p[2]"), "To  1");
    assert_eq!(emptied.value().unresolved(), ["First", "Last Name", "gone"]);
}

#[test]
fn a_template_without_placeholders_is_an_exact_no_op() {
    let bytes = docx(
        r#"<w:p><w:r><w:t>{not a placeholder} {{ }} {{-x}}</w:t></w:r></w:p>"#,
        None,
    );
    let template = open(bytes.clone());
    let commit = template
        .merge(&record(&[("x", "y")]), &Options::new())
        .unwrap();
    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().commits(), 0);
    assert_eq!(commit.diagnostics().operations(), 1);
    assert_eq!(
        commit.file().to_bytes().unwrap(),
        template.to_bytes().unwrap()
    );
}

#[test]
fn one_template_fills_every_record() {
    let template = open(docx(
        r#"<w:p><w:r><w:t>Hello {{name}} ({{n}})</w:t></w:r></w:p>"#,
        None,
    ));
    let data = Data::from_csv(
        b"name,n\nAda,1\nGrace,2\n\"Lin, M\",3\n",
        &litchi_ops::import::Options::csv(),
        &SemanticLimits::default(),
    )
    .unwrap();
    let texts: Vec<String> = data
        .records()
        .iter()
        .map(|record| {
            text(
                template.merge(record, &Options::new()).unwrap().file(),
                "/body/p[1]",
            )
        })
        .collect();
    assert_eq!(
        texts,
        ["Hello Ada (1)", "Hello Grace (2)", "Hello Lin, M (3)"]
    );
    assert_eq!(text(&template, "/body/p[1]"), "Hello {{name}} ({{n}})");
}

#[test]
fn merge_composes_with_other_edits() {
    let template = open(docx(
        r#"<w:p><w:r><w:t>{{greeting}}</w:t></w:r></w:p>"#,
        None,
    ));
    let commit = template
        .edit(|edit| {
            let summary = edit.merge(&record(&[("greeting", "Hi {{name}}")]), &Options::new())?;
            edit.set("/body/p[1]", props! { "align" => "center" })?;
            Ok(summary)
        })
        .unwrap();
    // A value is never matched again.
    assert_eq!(text(commit.file(), "/body/p[1]"), "Hi {{name}}");
    assert_eq!(commit.value().replacements(), 1);
    assert_eq!(commit.diagnostics().operations(), 2);
    let restored = commit.patch().inverse().apply(commit.file()).unwrap();
    assert_eq!(restored.to_bytes().unwrap(), template.to_bytes().unwrap());
}

#[test]
fn values_and_limits_are_checked_before_any_change() {
    let template = open(docx(
        r#"<w:p><w:r><w:t>{{a}} {{b}}</w:t></w:r></w:p>"#,
        None,
    ));
    let bad = template
        .merge(&record(&[("a", "\u{1}")]), &Options::new())
        .unwrap_err();
    assert_eq!(bad.code(), Code::InvalidValue);
    let limits = SemanticLimits::builder()
        .max_expanded_targets(1)
        .unwrap()
        .build()
        .unwrap();
    let error = template
        .merge_with_options(
            &record(&[("a", "1"), ("b", "2")]),
            &Options::new(),
            litchi_ops::EditOptions::new().semantic_limits(limits),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::LimitExceeded);
    let limits = SemanticLimits::builder()
        .max_staged_bytes(100)
        .unwrap()
        .build()
        .unwrap();
    let error = template
        .merge_with_options(
            &record(&[("a", "1"), ("b", "2")]),
            &Options::new(),
            litchi_ops::EditOptions::new().semantic_limits(limits),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::LimitExceeded);
    assert_eq!(text(&template, "/body/p[1]"), "{{a}} {{b}}");
    // Tabs and line breaks become Word run content.
    let commit = template
        .merge(&record(&[("a", "x\ty"), ("b", "1\r\n2")]), &Options::new())
        .unwrap();
    let document = member(&commit.file().to_bytes().unwrap(), "word/document.xml");
    assert!(
        document.contains("<w:t>x</w:t><w:tab/><w:t>y 1</w:t><w:br/><w:t>2</w:t>"),
        "{document}"
    );
}

fn pptx_template() -> File {
    let file = litchi_ops::create(FileFormat::Pptx)
        .unwrap()
        .add("/", NewNode::new("slide").prop("layout", "Blank"))
        .unwrap()
        .file
        .add(
            "/slide[1]",
            NewNode::new("shape")
                .prop("width", "2in")
                .prop("height", "1in")
                .prop("text", "Hello {{name}}!"),
        )
        .unwrap()
        .file;
    let bytes = rewrite(&file.to_bytes().unwrap(), "ppt/slides/slide1.xml", |xml| {
        xml.replace(
            r#"<a:r><a:t xml:space="preserve">Hello {{name}}!</a:t></a:r>"#,
            r#"<a:r><a:rPr lang="en-US" b="1"/><a:t>Hello {{na</a:t></a:r><a:r><a:rPr lang="en-US" i="1"/><a:t>me}}! </a:t></a:r><a:br/><a:r><a:t>{{city}}</a:t></a:r>"#,
        )
    });
    open(bytes)
}

#[test]
fn pptx_runs_keep_their_formatting() {
    let template = pptx_template();
    assert_eq!(
        text(&template, "/slide[1]/shape[1]"),
        "Hello {{name}}! \n{{city}}"
    );
    let commit = template
        .merge(
            &record(&[("name", "Ada"), ("city", "Paris & Co")]),
            &Options::new(),
        )
        .unwrap();
    let file = commit.file();
    assert_eq!(text(file, "/slide[1]/shape[1]"), "Hello Ada! \nParis & Co");
    assert_eq!(
        text(file, "/slide[1]/shape[1]/paragraph[1]/run[1]"),
        "Hello Ada"
    );
    assert_eq!(text(file, "/slide[1]/shape[1]/paragraph[1]/run[2]"), "! ");
    assert_eq!(
        property(file, "/slide[1]/shape[1]/paragraph[1]/run[1]", "bold").as_deref(),
        Some("true")
    );
    assert_eq!(
        property(file, "/slide[1]/shape[1]/paragraph[1]/run[2]", "italic").as_deref(),
        Some("true")
    );
    assert_eq!(commit.value().replacements(), 2);
    assert_eq!(commit.diagnostics().commits(), 1);
    let reopened = open(file.to_bytes().unwrap());
    assert_eq!(
        text(&reopened, "/slide[1]/shape[1]"),
        "Hello Ada! \nParis & Co"
    );
    let restored = commit.patch().inverse().apply(file).unwrap();
    assert_eq!(
        text(&restored, "/slide[1]/shape[1]"),
        text(&template, "/slide[1]/shape[1]")
    );
    // A line break cannot go into a run.
    let error = template
        .merge(&record(&[("name", "a\nb")]), &Options::new())
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
    // No match: no commit.
    let none = template
        .merge(&record(&[("other", "x")]), &Options::new())
        .unwrap();
    assert!(none.patch().is_noop());
    assert_eq!(none.value().unresolved(), ["city", "name"]);
}

#[test]
fn xlsx_text_cells_are_filled_and_formulas_kept() {
    let template = litchi_ops::create(FileFormat::Xlsx)
        .unwrap()
        .edit(|edit| {
            edit.set(
                "/Sheet1/A1",
                props! { "value" => "Invoice {{number}} for {{client}}" },
            )?;
            edit.set("/Sheet1/B2", props! { "value" => 42 })?;
            edit.set("/Sheet1/C3", props! { "formula" => "\"{{client}}\"&\"!\"" })?;
            edit.add("/", NewNode::new("sheet").prop("name", "Second"))?;
            edit.set("/Second/A1", props! { "value" => "{{client}}" })?;
            Ok(())
        })
        .unwrap()
        .into_file();
    let commit = template
        .merge(
            &record(&[("number", "0274"), ("client", "Acme")]),
            &Options::new(),
        )
        .unwrap();
    let file = commit.file();
    assert_eq!(text(file, "/Sheet1/A1"), "Invoice 0274 for Acme");
    assert_eq!(text(file, "/Second/A1"), "Acme");
    assert_eq!(text(file, "/Sheet1/B2"), "42");
    assert_eq!(
        property(file, "/Sheet1/C3", "formula").as_deref(),
        Some("\"{{client}}\"&\"!\"")
    );
    assert_eq!(commit.value().replacements(), 3);
    assert_eq!(commit.diagnostics().commits(), 1);
    let reopened = open(file.to_bytes().unwrap());
    assert_eq!(text(&reopened, "/Sheet1/A1"), "Invoice 0274 for Acme");
    let restored = commit.patch().inverse().apply(file).unwrap();
    assert_eq!(
        text(&restored, "/Sheet1/A1"),
        "Invoice {{number}} for {{client}}"
    );
}

#[test]
fn other_formats_report_the_gap() {
    #[cfg(feature = "odg")]
    {
        let format = FileFormat::Odg;
        let file = litchi_ops::create(format).unwrap();
        let error = file.merge(&Record::new(), &Options::new()).unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert_eq!(error.format(), Some(format));
        assert!(!Capabilities::of(format).verb(Verb::Merge).is_supported());
    }
    // Record 0295: the workbook formats merge too; an empty template is an
    // exact no-op.
    for format in [FileFormat::Xlsb, FileFormat::Xls] {
        let file = litchi_ops::create(format).unwrap();
        let commit = file.merge(&Record::new(), &Options::new()).unwrap();
        assert!(commit.patch().is_noop());
    }
    for format in [
        FileFormat::Docx,
        FileFormat::Pptx,
        FileFormat::Xlsx,
        FileFormat::Xlsb,
        FileFormat::Xls,
    ] {
        assert!(
            Capabilities::of(format).verb(Verb::Merge).is_supported(),
            "{format:?}"
        );
        let file = litchi_ops::create(format).unwrap();
        assert!(
            file.capabilities().verb(Verb::Merge).is_supported(),
            "{format:?}"
        );
    }
}

#[test]
fn data_errors_are_typed() {
    for (source, code) in [
        (&b"[1]"[..], Code::InvalidJson),
        (b"\"text\"", Code::InvalidJson),
        (b"{\"a\": }", Code::InvalidJson),
        (b"{\"a\": 1} x", Code::InvalidJson),
    ] {
        assert_eq!(
            Data::from_json(source, &SemanticLimits::default())
                .unwrap_err()
                .code(),
            code
        );
    }
    let deep = format!("{}1{}", "[".repeat(200), "]".repeat(200));
    let error = Data::from_json(
        format!("{{\"a\": {deep}}}").as_bytes(),
        &SemanticLimits::default(),
    )
    .unwrap_err();
    assert_eq!(error.code(), Code::LimitExceeded);
    let error = Data::from_csv(
        b"a\n\"open",
        &litchi_ops::import::Options::csv(),
        &SemanticLimits::default(),
    )
    .unwrap_err();
    assert_eq!(error.code(), Code::InvalidInput);
}

/// Record 0285 (0274 follow-ups F2, F3): table cells and speaker notes are
/// merged run by run through the text bodies find/replace reads.
#[cfg(feature = "pptx")]
#[test]
fn pptx_table_cells_and_notes_are_merged() {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    let presentation = builder.presentation().unwrap();
    let slide = presentation.add_slide().unwrap();
    slide.add_text_box("To {{name}}", 0, 0, 1_828_800, 914_400);
    slide.set_notes("Call {{name}} in {{city}}");
    let package = builder.build().unwrap();
    let mut edit = package.edit().unwrap();
    edit.add_table(
        0usize,
        &[vec!["{{city}}".to_owned(), "fixed".to_owned()]],
        (0, 1_000_000, 2_000_000, 600_000),
    )
    .unwrap();
    let bytes = edit.commit().unwrap().into_package().to_bytes().unwrap();
    for mode in [SourceMode::Eager, SourceMode::Auto] {
        let template =
            litchi_ops::open_bytes_with(bytes.clone(), &OpenOptions::new().source_mode(mode))
                .unwrap();
        let commit = template
            .merge(
                &record(&[("name", "Ada"), ("city", "Paris")]),
                &Options::new(),
            )
            .unwrap();
        let file = commit.file();
        assert_eq!(commit.value().replacements(), 4);
        assert_eq!(text(file, "/slide[1]/shape[1]"), "To Ada");
        assert_eq!(text(file, "/slide[1]/table[1]/tr[1]/tc[1]"), "Paris");
        assert_eq!(text(file, "/slide[1]/table[1]/tr[1]/tc[2]"), "fixed");
        assert_eq!(text(file, "/slide[1]/notes"), "Call Ada in Paris");
        let restored = commit.patch().inverse().apply(file).unwrap();
        assert_eq!(
            text(&restored, "/slide[1]/notes"),
            "Call {{name}} in {{city}}"
        );
    }
}
