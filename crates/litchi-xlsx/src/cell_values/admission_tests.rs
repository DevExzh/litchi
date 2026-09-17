#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::too_many_lines,
    reason = "one constructed witness per construct of the dependency rule"
)]
//! One constructed witness per construct of change 0657's dependency rule.
//!
//! Each test builds the smallest package that carries exactly one producer
//! construct and states the rule's verdict on it: **admitted** (the rewrite
//! copies it and an edit inside it changes nothing else), **adjusted** (the
//! commit rewrites it deliberately), or **refused** (its meaning depends on
//! the value the edit replaces, and the editor does not model it).

use std::sync::Arc;

use litchi_core::OwnedSource;
use litchi_opc::constants::content_type as ct;
use soapberry_zip::office::StreamingArchiveWriter;

use crate::cell_values::{CellValueEdit, SheetCellValueEdit, SourceBackedEditor};
use crate::error::EditBlock;
use crate::{Address, Error, Number, Value};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const X14AC: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";

/// A package with one worksheet and configurable relationship envelopes.
#[derive(Default)]
struct PackageShape {
    /// Extra package-root relationships, as `(Id, Type, Target, external)`.
    package_relationships: Vec<(&'static str, String, &'static str, bool)>,
    /// Extra workbook relationships, in the same shape.
    workbook_relationships: Vec<(&'static str, String, &'static str, bool)>,
    /// Worksheet relationships, in the same shape.
    worksheet_relationships: Vec<(&'static str, String, &'static str, bool)>,
    /// Extra members, as `(part name, bytes)`; no content-type override.
    members: Vec<(&'static str, Vec<u8>)>,
    /// Extra workbook children before `<sheets>`.
    workbook_body: Option<String>,
    /// Extra workbook children after `<sheets>`.
    workbook_tail: Option<String>,
}

/// The canonical relationship serialization `litchi-opc` itself writes, so
/// that a published package can replace a relationship member in place.
fn relationships(entries: &[(&'static str, String, &'static str, bool)]) -> String {
    let mut sorted = entries.to_vec();
    sorted.sort_by(|left, right| left.0.cmp(right.0));
    let mut xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"{PKG_REL}\">"
    );
    for (id, kind, target, external) in &sorted {
        xml.push_str(&format!(
            "<Relationship Id=\"{id}\" Type=\"{kind}\" Target=\"{target}\"{}/>",
            if *external {
                " TargetMode=\"External\""
            } else {
                ""
            }
        ));
    }
    xml.push_str("</Relationships>");
    xml
}

fn package(sheet: &str, shape: &PackageShape) -> Vec<u8> {
    let shared_strings_override = if shape
        .members
        .iter()
        .any(|(name, _)| *name == "xl/sharedStrings.xml")
    {
        format!(
            r#"<Override PartName="/xl/sharedStrings.xml" ContentType="{}"/>"#,
            ct::SML_SHARED_STRINGS
        )
    } else {
        String::new()
    };
    let content_types = format!(
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"bin\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"{}\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"{}\"/><Override PartName=\"/xl/calcChain.xml\" ContentType=\"{}\"/>{shared_strings_override}</Types>",
        ct::SML_SHEET_MAIN,
        ct::SML_WORKSHEET,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
    );
    let body = shape.workbook_body.clone().unwrap_or_default();
    let tail = shape.workbook_tail.clone().unwrap_or_default();
    let workbook = format!(
        "<workbook xmlns=\"{SML}\" xmlns:r=\"{REL}\">{body}<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rIdSheet\"/></sheets>{tail}</workbook>"
    );
    let mut package_relationships = vec![(
        "rIdRoot",
        format!("{REL}/officeDocument"),
        "xl/workbook.xml",
        false,
    )];
    package_relationships.extend(shape.package_relationships.iter().cloned());
    let mut workbook_relationships = vec![(
        "rIdSheet",
        format!("{REL}/worksheet"),
        "worksheets/sheet1.xml",
        false,
    )];
    workbook_relationships.extend(shape.workbook_relationships.iter().cloned());

    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .expect("content-types member");
    writer
        .write_stored(
            "_rels/.rels",
            relationships(&package_relationships).as_bytes(),
        )
        .expect("package relationships member");
    writer
        .write_stored("xl/workbook.xml", workbook.as_bytes())
        .expect("workbook member");
    writer
        .write_stored(
            "xl/_rels/workbook.xml.rels",
            relationships(&workbook_relationships).as_bytes(),
        )
        .expect("workbook relationships member");
    writer
        .write_stored("xl/worksheets/sheet1.xml", sheet.as_bytes())
        .expect("worksheet member");
    if !shape.worksheet_relationships.is_empty() {
        writer
            .write_stored(
                "xl/worksheets/_rels/sheet1.xml.rels",
                relationships(&shape.worksheet_relationships).as_bytes(),
            )
            .expect("worksheet relationships member");
    }
    for (name, bytes) in &shape.members {
        writer.write_stored(name, bytes).expect("extra member");
    }
    writer.finish_to_bytes().expect("test XLSX archive")
}

fn editor(sheet: &str, shape: &PackageShape) -> Result<SourceBackedEditor, Error> {
    let bytes = package(sheet, shape);
    SourceBackedEditor::from_read_at(Arc::new(OwnedSource::new(bytes)))
}

/// Set `A1` to 42 through the multi-sheet door and return the committed
/// worksheet bytes.
fn edit_a1(sheet: &str, shape: &PackageShape) -> Result<Vec<u8>, Error> {
    edit_cell(sheet, shape, "A1")
}

fn edit_cell(sheet: &str, shape: &PackageShape, address: &str) -> Result<Vec<u8>, Error> {
    let editor = editor(sheet, shape)?;
    let commit = editor
        .edit_many([SheetCellValueEdit::set(
            "Sheet1",
            Address::from_a1(address).expect("A1 address"),
            Value::Number(Number::new("42").expect("numeral")),
        )])?
        .commit()?;
    let snapshot = commit.snapshot();
    Ok(snapshot
        .sheets()
        .first()
        .expect("one worksheet")
        .source_xml()
        .to_vec())
}

fn message(error: &Error) -> String {
    error.to_string()
}

/// A scalar worksheet whose head and tail carry `head` and `tail` verbatim.
fn worksheet(head: &str, tail: &str) -> String {
    worksheet_with("", head, tail)
}

/// The same, with `pr` before `<dimension>` for the one child `CT_Worksheet`
/// orders ahead of it.
fn worksheet_with(pr: &str, head: &str, tail: &str) -> String {
    format!(
        "<worksheet xmlns=\"{SML}\" xmlns:r=\"{REL}\">{pr}<dimension ref=\"A1:B1\"/>{head}<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><v>2</v></c></row></sheetData>{tail}</worksheet>"
    )
}

/// Every byte outside the edited `<c>` span and the `<dimension>` tag is the
/// source's own.
fn assert_only_the_edited_cell_changed(source: &str, output: &[u8], edited: &str) {
    assert_only_the_edited_cell_changed_beside(source, output, edited, "<c r=\"B1\"><v>2</v></c>");
}

fn assert_only_the_edited_cell_changed_beside(
    source: &str,
    output: &[u8],
    edited: &str,
    neighbour: &str,
) {
    let output = std::str::from_utf8(output).expect("UTF-8 output");
    let head_end = source.find("<sheetData>").expect("sheetData in source") + "<sheetData>".len();
    assert_eq!(
        &source[..head_end],
        &output[..head_end],
        "the head before sheetData must be copied verbatim"
    );
    let tail = &source[source.find("</sheetData>").expect("sheetData close")..];
    assert!(
        output.ends_with(tail),
        "the tail from </sheetData> must be copied verbatim"
    );
    assert!(
        output.contains(&format!("<c r=\"{edited}\"")),
        "the edited cell must still be present"
    );
    assert!(
        output.contains("<v>42</v>"),
        "the new value must be written"
    );
    assert!(
        output.contains(neighbour),
        "an unedited neighbour must be copied verbatim"
    );
}

// ---------------------------------------------------------------- admitted

#[test]
fn compatibility_markers_on_the_worksheet_root_are_copied_through() {
    let source = format!(
        "<worksheet xmlns=\"{SML}\" xmlns:r=\"{REL}\" xmlns:mc=\"{MC}\" xmlns:x14ac=\"{X14AC}\" xmlns:xr=\"urn:fixture:xr\" mc:Ignorable=\"x14ac xr\" xr:uid=\"{{0000-0001}}\"><dimension ref=\"A1:B1\"/><sheetFormatPr defaultRowHeight=\"15\" x14ac:dyDescent=\"0.25\"/><sheetData><row r=\"1\" spans=\"1:2\" x14ac:dyDescent=\"0.25\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><v>2</v></c></row></sheetData></worksheet>"
    );
    let output = edit_a1(&source, &PackageShape::default()).expect("marker-bearing worksheet");
    assert_only_the_edited_cell_changed(&source, &output, "A1");
    let output = std::str::from_utf8(&output).expect("UTF-8");
    assert!(
        output.contains("x14ac:dyDescent=\"0.25\""),
        "the row descent must survive the rewrite"
    );
}

#[test]
fn out_of_sheet_data_children_are_copied_through() {
    // Every one of these is a construct whose range addresses cells by
    // coordinate. A value change does not move any of them.
    let tail = concat!(
        "<autoFilter ref=\"A1:B1\"/>",
        "<mergeCells count=\"1\"><mergeCell ref=\"A3:B3\"/></mergeCells>",
        "<conditionalFormatting sqref=\"A1:B1\"><cfRule type=\"cellIs\" dxfId=\"0\" priority=\"1\" operator=\"greaterThan\"><formula>1</formula></cfRule></conditionalFormatting>",
        "<dataValidations count=\"1\"><dataValidation type=\"whole\" sqref=\"A5\"><formula1>0</formula1></dataValidation></dataValidations>",
        "<printOptions horizontalCentered=\"1\"/>",
        "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>",
        "<pageSetup paperSize=\"9\" orientation=\"portrait\"/>",
        "<headerFooter><oddHeader>&amp;C header</oddHeader></headerFooter>",
        "<rowBreaks count=\"1\" manualBreakCount=\"1\"><brk id=\"1\" max=\"16383\" man=\"1\"/></rowBreaks>",
        "<ignoredErrors><ignoredError sqref=\"A1\" numberStoredAsText=\"1\"/></ignoredErrors>",
    );
    let head = "<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><selection activeCell=\"A1\" sqref=\"A1\"/></sheetView></sheetViews><cols><col min=\"1\" max=\"2\" width=\"9\"/></cols>";
    let source = worksheet_with(
        "<sheetPr codeName=\"Sheet1\"><tabColor rgb=\"FFFF0000\"/></sheetPr>",
        head,
        tail,
    );
    let output = edit_a1(&source, &PackageShape::default()).expect("producer-shaped worksheet");
    assert_only_the_edited_cell_changed(&source, &output, "A1");
}

#[test]
fn a_foreign_extension_payload_is_copied_through() {
    let tail = "<extLst><ext xmlns:x14=\"urn:fixture:x14\" uri=\"{FIXTURE}\"><x14:sparklineGroups><x14:sparklineGroup><x14:sparklines><x14:sparkline><xm:f xmlns:xm=\"urn:fixture:xm\">Sheet1!A1:B1</xm:f></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>";
    let source = worksheet("", tail);
    let output = edit_a1(&source, &PackageShape::default()).expect("extLst-bearing worksheet");
    assert_only_the_edited_cell_changed(&source, &output, "A1");
}

#[test]
fn worksheet_relationships_are_admitted_and_their_references_preserved() {
    let shape = PackageShape {
        worksheet_relationships: vec![(
            "rIdPrn",
            format!("{REL}/printerSettings"),
            "../printerSettings/printerSettings1.bin",
            false,
        )],
        members: vec![("xl/printerSettings/printerSettings1.bin", vec![0u8; 32])],
        ..PackageShape::default()
    };
    let source = worksheet(
        "",
        "<pageSetup paperSize=\"9\" orientation=\"portrait\" r:id=\"rIdPrn\"/>",
    );
    let output = edit_a1(&source, &shape).expect("relationship-bearing worksheet");
    assert_only_the_edited_cell_changed(&source, &output, "A1");
    let output = std::str::from_utf8(&output).expect("UTF-8");
    assert!(
        output.contains("r:id=\"rIdPrn\""),
        "the relationship reference must be preserved"
    );
}

#[test]
fn unfamiliar_package_and_workbook_relationships_are_admitted() {
    let shape = PackageShape {
        package_relationships: vec![
            (
                "rIdApp",
                format!("{REL}/extended-properties"),
                "docProps/app.xml",
                false,
            ),
            (
                "rIdCore",
                "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties".to_owned(),
                "docProps/core.xml",
                false,
            ),
        ],
        workbook_relationships: vec![(
            "rIdLink",
            format!("{REL}/externalLink"),
            "externalLinks/externalLink1.xml",
            false,
        )],
        members: vec![
            ("docProps/app.xml", b"<Properties/>".to_vec()),
            ("docProps/core.xml", b"<coreProperties/>".to_vec()),
            (
                "xl/externalLinks/externalLink1.xml",
                b"<externalLink/>".to_vec(),
            ),
        ],
        ..PackageShape::default()
    };
    let source = worksheet("", "");
    let output = edit_a1(&source, &shape).expect("docProps-bearing package");
    assert_only_the_edited_cell_changed(&source, &output, "A1");
}

#[test]
fn unfamiliar_workbook_children_are_copied_through() {
    let shape = PackageShape {
        workbook_body: Some(
            "<fileVersion appName=\"xl\" rupBuild=\"12345\" future=\"1\"/><workbookPr defaultThemeVersion=\"124226\"/><workbookProtection lockStructure=\"1\"/><bookViews><workbookView xWindow=\"0\" yWindow=\"0\"/></bookViews>"
                .to_owned(),
        ),
        ..PackageShape::default()
    };
    let source = worksheet("", "");
    edit_a1(&source, &shape).expect("workbook with unfamiliar children");

    let with_names = PackageShape {
        workbook_tail: Some(
            "<definedNames><definedName name=\"Range\">Sheet1!$A$1:$B$1</definedName></definedNames>"
                .to_owned(),
        ),
        ..PackageShape::default()
    };
    edit_a1(&source, &with_names).expect("workbook with defined names");
}

#[test]
fn inline_rich_text_is_admitted_and_reads_back_complete() {
    let source = format!(
        "<worksheet xmlns=\"{SML}\"><dimension ref=\"A1:B1\"/><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\" t=\"inlineStr\"><is><r><rPr><b/><sz val=\"11\"/></rPr><t>bo</t></r><r><t>ld</t></r><phoneticPr fontId=\"1\"/></is></c></row></sheetData></worksheet>"
    );
    let editor = editor(&source, &PackageShape::default()).expect("rich inline worksheet");
    let commit = editor
        .edit_many([SheetCellValueEdit::set(
            "Sheet1",
            Address::from_a1("A1").expect("A1"),
            Value::Number(Number::new("42").expect("numeral")),
        )])
        .expect("plan beside a rich inline cell")
        .commit()
        .expect("commit beside a rich inline cell");
    let snapshot = commit.snapshot();
    assert_eq!(
        snapshot.value(0, Address::from_a1("B1").expect("B1")),
        Some(&Value::Text("bold".into())),
        "the rich runs must read back concatenated"
    );
    let output = snapshot
        .sheets()
        .first()
        .expect("one worksheet")
        .source_xml();
    assert_only_the_edited_cell_changed_beside(
        &source,
        output,
        "A1",
        "<c r=\"B1\" t=\"inlineStr\"><is><r>",
    );
    let output = std::str::from_utf8(output).expect("UTF-8");
    assert!(
        output.contains("<rPr><b/><sz val=\"11\"/></rPr>"),
        "run properties must be copied verbatim"
    );
}

// ---------------------------------------------------------------- adjusted

#[test]
fn a_merged_range_admits_its_anchor_and_refuses_a_covered_cell() {
    let source = format!(
        "<worksheet xmlns=\"{SML}\"><dimension ref=\"A1:B1\"/><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><v>2</v></c></row></sheetData><mergeCells count=\"1\"><mergeCell ref=\"A1:B1\"/></mergeCells></worksheet>"
    );
    let output = edit_a1(&source, &PackageShape::default()).expect("the anchor is editable");
    assert_only_the_edited_cell_changed(&source, &output, "A1");

    let error = edit_cell(&source, &PackageShape::default(), "B1")
        .expect_err("a covered cell is not editable");
    match error {
        Error::EditBlocked { reason, .. } => assert_eq!(reason, EditBlock::CoveredMerge),
        other => panic!("expected a covered-merge refusal, got {other:?}"),
    }
}

#[test]
fn a_group_formula_refuses_an_edit_inside_its_range() {
    let array = format!(
        "<worksheet xmlns=\"{SML}\"><dimension ref=\"A1:B1\"/><sheetData><row r=\"1\"><c r=\"A1\"><f t=\"array\" ref=\"A1:B1\">SUM(1)</f><v>1</v></c><c r=\"B1\"><v>2</v></c></row></sheetData></worksheet>"
    );
    let error = edit_cell(&array, &PackageShape::default(), "B1")
        .expect_err("an array-formula member is not editable");
    match error {
        Error::EditBlocked { reason, .. } => assert_eq!(reason, EditBlock::GroupFormula),
        other => panic!("expected a group-formula refusal, got {other:?}"),
    }

    let shared = format!(
        "<worksheet xmlns=\"{SML}\"><dimension ref=\"A1:B1\"/><sheetData><row r=\"1\"><c r=\"A1\"><f t=\"shared\" ref=\"A1:B1\" si=\"0\">1+1</f><v>2</v></c><c r=\"B1\"><f t=\"shared\" si=\"0\"/><v>2</v></c></row></sheetData></worksheet>"
    );
    let error = edit_cell(&shared, &PackageShape::default(), "B1")
        .expect_err("a shared-formula follower is not editable");
    match error {
        Error::EditBlocked { reason, .. } => assert_eq!(reason, EditBlock::GroupFormula),
        other => panic!("expected a group-formula refusal, got {other:?}"),
    }
}

#[test]
fn the_calculation_chain_is_dropped_with_its_relationship() {
    let shape = PackageShape {
        workbook_relationships: vec![(
            "rIdChain",
            format!("{REL}/calcChain"),
            "calcChain.xml",
            false,
        )],
        members: vec![(
            "xl/calcChain.xml",
            format!("<calcChain xmlns=\"{SML}\"><c r=\"A1\" i=\"1\"/></calcChain>").into_bytes(),
        )],
        ..PackageShape::default()
    };
    let source = worksheet("", "");
    let bytes = package(&source, &shape);
    let editor = SourceBackedEditor::from_read_at(Arc::new(OwnedSource::new(bytes)))
        .expect("calcChain package");
    let commit = editor
        .edit_many([SheetCellValueEdit::set(
            "Sheet1",
            Address::from_a1("A1").expect("A1"),
            Value::Number(Number::new("42").expect("numeral")),
        )])
        .expect("plan")
        .commit()
        .expect("commit");
    assert!(commit.changed());
    let mut sink = Vec::new();
    editor
        .publish_multi_commit_to_stream(&mut sink, &commit)
        .expect("publish");
    let published = String::from_utf8_lossy(&sink).into_owned();
    assert!(
        !published.contains("calcChain.xml"),
        "the chain part and its relationship are dropped together"
    );
}

// ---------------------------------------------------------------- adjusted

#[test]
fn a_shared_string_part_is_admitted_and_preserved() {
    let shape = PackageShape {
        workbook_relationships: vec![(
            "rIdSst",
            format!("{REL}/sharedStrings"),
            "sharedStrings.xml",
            false,
        )],
        members: vec![(
            "xl/sharedStrings.xml",
            format!("<sst xmlns=\"{SML}\" count=\"1\" uniqueCount=\"1\"><si><t>a</t></si></sst>")
                .into_bytes(),
        )],
        ..PackageShape::default()
    };
    let source = format!(
        "<worksheet xmlns=\"{SML}\"><dimension ref=\"A1:B1\"/><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\" t=\"s\"><v>0</v></c></row></sheetData></worksheet>"
    );
    let output = edit_a1(&source, &shape).expect("shared strings are readable");
    assert_only_the_edited_cell_changed_beside(
        &source,
        &output,
        "A1",
        "<c r=\"B1\" t=\"s\"><v>0</v></c>",
    );

    let bytes = package(&source, &shape);
    let editor = SourceBackedEditor::from_read_at(Arc::new(OwnedSource::new(bytes)))
        .expect("shared-string package");
    let commit = editor
        .edit_many([SheetCellValueEdit::set(
            "Sheet1",
            Address::from_a1("A1").expect("A1"),
            Value::Number(Number::new("42").expect("numeral")),
        )])
        .expect("plan beside a shared-string cell")
        .commit()
        .expect("commit beside a shared-string cell");
    assert_eq!(
        commit
            .snapshot()
            .value(0, Address::from_a1("B1").expect("B1")),
        Some(&Value::Text("a".into()))
    );
    let mut published = Vec::new();
    editor
        .publish_multi_commit_to_stream(&mut published, &commit)
        .expect("publish shared-string package");
    let published = String::from_utf8_lossy(&published);
    assert!(
        published.contains("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\"><si><t>a</t></si></sst>"),
        "the shared-string part must be transferred byte-for-byte"
    );
}

#[test]
fn a_shared_string_cell_cannot_be_mutated_or_removed() {
    let shape = PackageShape {
        workbook_relationships: vec![(
            "rIdSst",
            format!("{REL}/sharedStrings"),
            "sharedStrings.xml",
            false,
        )],
        members: vec![(
            "xl/sharedStrings.xml",
            format!("<sst xmlns=\"{SML}\" count=\"1\" uniqueCount=\"1\"><si><t>a</t></si></sst>")
                .into_bytes(),
        )],
        ..PackageShape::default()
    };
    let source = format!(
        "<worksheet xmlns=\"{SML}\"><dimension ref=\"A1\"/><sheetData><row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row></sheetData></worksheet>"
    );
    for edit in [
        CellValueEdit::set(
            Address::from_a1("A1").expect("A1"),
            Value::Number(Number::new("42").expect("numeral")),
        ),
        CellValueEdit::clear(Address::from_a1("A1").expect("A1")),
        CellValueEdit::remove(Address::from_a1("A1").expect("A1")),
    ] {
        let error = match editor(&source, &shape).and_then(|editor| {
            editor.edit_many([SheetCellValueEdit {
                selector: "Sheet1".into(),
                edit,
            }])
        }) {
            Err(error) => error,
            Ok(_) => panic!("a shared-string cell mutation must be refused"),
        };
        assert_eq!(
            message(&error),
            "invalid XLSX structure: value-only edits cannot add, remove, or renumber shared strings"
        );
    }
}

#[test]
fn a_shared_string_table_is_lazy_when_no_cell_references_it() {
    let shape = PackageShape {
        workbook_relationships: vec![(
            "rIdSst",
            format!("{REL}/sharedStrings"),
            "sharedStrings.xml",
            false,
        )],
        members: vec![("xl/sharedStrings.xml", b"<sst".to_vec())],
        ..PackageShape::default()
    };
    let output = edit_a1(&worksheet("", ""), &shape)
        .expect("a malformed but unreferenced shared-string table stays lazy");
    assert_only_the_edited_cell_changed(&worksheet("", ""), &output, "A1");
}

#[test]
fn a_large_shared_string_table_stays_within_the_retained_part_bound() {
    const ITEMS: usize = 66_935;
    let mut table = String::with_capacity(ITEMS * 30);
    table.push_str(&format!(
        "<sst xmlns=\"{SML}\" count=\"{ITEMS}\" uniqueCount=\"{ITEMS}\">"
    ));
    for index in 0..ITEMS {
        table.push_str(&format!("<si><t>shared-{index}</t></si>"));
    }
    table.push_str("</sst>");
    let shape = PackageShape {
        workbook_relationships: vec![(
            "rIdSst",
            format!("{REL}/sharedStrings"),
            "sharedStrings.xml",
            false,
        )],
        members: vec![("xl/sharedStrings.xml", table.into_bytes())],
        ..PackageShape::default()
    };
    let source = format!(
        "<worksheet xmlns=\"{SML}\"><dimension ref=\"A1:B1\"/><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\" t=\"s\"><v>{}</v></c></row></sheetData></worksheet>",
        ITEMS - 1
    );
    let editor = editor(&source, &shape).expect("large shared-string package");
    let commit = editor
        .edit_many([SheetCellValueEdit::set(
            "Sheet1",
            Address::from_a1("A1").expect("A1"),
            Value::Number(Number::new("42").expect("numeral")),
        )])
        .expect("numeric edit beside the table")
        .commit()
        .expect("large shared-string table is bounded and readable");
    assert_eq!(
        commit
            .snapshot()
            .value(0, Address::from_a1("B1").expect("B1")),
        Some(&Value::Text(format!("shared-{}", ITEMS - 1).into()))
    );
}

#[test]
fn shared_string_cells_require_a_valid_table_and_index() {
    let missing_relationship = format!(
        "<worksheet xmlns=\"{SML}\"><dimension ref=\"A1\"/><sheetData><row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row></sheetData></worksheet>"
    );
    let error = edit_a1(&missing_relationship, &PackageShape::default())
        .expect_err("a shared-string cell without a relationship is invalid");
    assert_eq!(
        message(&error),
        "invalid XLSX structure: worksheet uses shared strings but the workbook has no shared-string part"
    );

    let shape = PackageShape {
        workbook_relationships: vec![(
            "rIdSst",
            format!("{REL}/sharedStrings"),
            "sharedStrings.xml",
            false,
        )],
        members: vec![(
            "xl/sharedStrings.xml",
            format!("<sst xmlns=\"{SML}\" count=\"1\" uniqueCount=\"1\"><si><t>a</t></si></sst>")
                .into_bytes(),
        )],
        ..PackageShape::default()
    };
    let out_of_range = format!(
        "<worksheet xmlns=\"{SML}\"><dimension ref=\"A1\"/><sheetData><row r=\"1\"><c r=\"A1\" t=\"s\"><v>1</v></c></row></sheetData></worksheet>"
    );
    let error =
        edit_a1(&out_of_range, &shape).expect_err("an out-of-range shared index is invalid");
    assert_eq!(
        message(&error),
        "invalid XLSX structure: shared-string index 1 exceeds table length 1"
    );

    let malformed = PackageShape {
        members: vec![(
            "xl/sharedStrings.xml",
            format!("<sst xmlns=\"{SML}\"><si><t>a</t></si>").into_bytes(),
        )],
        ..shape
    };
    let valid_reference = format!(
        "<worksheet xmlns=\"{SML}\"><dimension ref=\"A1\"/><sheetData><row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row></sheetData></worksheet>"
    );
    let error = edit_a1(&valid_reference, &malformed)
        .expect_err("a referenced malformed shared-string table is invalid");
    assert_eq!(
        message(&error),
        "invalid XLSX structure: shared strings XML has a missing or unterminated SpreadsheetML sst root"
    );
}

#[test]
fn multiple_shared_string_relationships_are_refused_before_parsing() {
    let shape = PackageShape {
        workbook_relationships: vec![
            (
                "rIdSst1",
                format!("{REL}/sharedStrings"),
                "sharedStrings.xml",
                false,
            ),
            (
                "rIdSst2",
                format!("{REL}/sharedStrings"),
                "sharedStrings.xml",
                false,
            ),
        ],
        members: vec![(
            "xl/sharedStrings.xml",
            format!("<sst xmlns=\"{SML}\" count=\"1\" uniqueCount=\"1\"><si><t>a</t></si></sst>")
                .into_bytes(),
        )],
        ..PackageShape::default()
    };
    let error = edit_a1(&worksheet("", ""), &shape)
        .expect_err("multiple shared-string relationships are ambiguous");
    assert_eq!(
        message(&error),
        "invalid XLSX structure: cell edits require worksheet relationships and at most one styles, theme, shared-string, and calculation-chain relationship"
    );
}

#[test]
fn a_pivot_cache_is_refused_by_name() {
    let shape = PackageShape {
        workbook_relationships: vec![(
            "rIdPivot",
            format!("{REL}/pivotCacheDefinition"),
            "pivotCache/pivotCacheDefinition1.xml",
            false,
        )],
        members: vec![(
            "xl/pivotCache/pivotCacheDefinition1.xml",
            b"<pivotCacheDefinition/>".to_vec(),
        )],
        ..PackageShape::default()
    };
    let error = edit_a1(&worksheet("", ""), &shape).expect_err("a pivot cache is refused");
    assert_eq!(
        message(&error),
        format!(
            "invalid XLSX structure: value-only edits refuse workbook relationship '{REL}/pivotCacheDefinition'"
        )
    );
}

#[test]
fn a_table_and_a_query_table_are_refused_by_name() {
    for (kind, target) in [
        (format!("{REL}/table"), "../tables/table1.xml"),
        (
            format!("{REL}/queryTable"),
            "../queryTables/queryTable1.xml",
        ),
    ] {
        let shape = PackageShape {
            worksheet_relationships: vec![("rIdTable", kind.clone(), target, false)],
            members: vec![("xl/tables/table1.xml", b"<table/>".to_vec())],
            ..PackageShape::default()
        };
        let error = edit_a1(
            &worksheet(
                "",
                "<tableParts count=\"1\"><tablePart r:id=\"rIdTable\"/></tableParts>",
            ),
            &shape,
        )
        .expect_err("a table is refused");
        assert_eq!(
            message(&error),
            format!(
                "invalid XLSX structure: value-only edits refuse worksheet relationship '{kind}'"
            )
        );
    }
}

#[test]
fn external_relationships_stay_refused_at_every_level() {
    let package_external = PackageShape {
        package_relationships: vec![(
            "rIdExt",
            format!("{REL}/extended-properties"),
            "http://example.invalid/app.xml",
            true,
        )],
        ..PackageShape::default()
    };
    let error = edit_a1(&worksheet("", ""), &package_external)
        .expect_err("an external package relationship is refused");
    assert_eq!(
        message(&error),
        format!(
            "invalid XLSX structure: value-only edits refuse package relationship '{REL}/extended-properties'"
        )
    );

    let workbook_external = PackageShape {
        workbook_relationships: vec![(
            "rIdExt",
            format!("{REL}/externalLink"),
            "http://example.invalid/link.xml",
            true,
        )],
        ..PackageShape::default()
    };
    let error = edit_a1(&worksheet("", ""), &workbook_external)
        .expect_err("an external workbook relationship is refused");
    assert_eq!(
        message(&error),
        "invalid XLSX structure: value-only edits refuse external workbook relationships"
    );

    let worksheet_external = PackageShape {
        worksheet_relationships: vec![(
            "rIdLink",
            format!("{REL}/hyperlink"),
            "http://example.invalid/",
            true,
        )],
        ..PackageShape::default()
    };
    let error = edit_a1(&worksheet("", ""), &worksheet_external)
        .expect_err("an external worksheet relationship is refused");
    assert_eq!(
        message(&error),
        format!(
            "invalid XLSX structure: value-only edits refuse worksheet relationship '{REL}/hyperlink'"
        )
    );
}

#[test]
fn cell_metadata_is_admitted_by_the_vocabulary_and_refused_by_the_scalar_gate() {
    for attribute in ["cm=\"1\"", "vm=\"1\""] {
        let source = format!(
            "<worksheet xmlns=\"{SML}\"><dimension ref=\"A1:B1\"/><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\" {attribute}><v>2</v></c></row></sheetData></worksheet>"
        );
        let error = edit_a1(&source, &PackageShape::default()).expect_err("cell metadata refuses");
        assert_eq!(
            message(&error),
            "invalid XLSX structure: cell edits refuse unknown cells and cell metadata",
            "the refusal for {attribute} moved to the scalar-cell gate"
        );
    }
}

#[test]
fn a_relationship_reference_inside_sheet_data_is_refused() {
    let source = format!(
        "<worksheet xmlns=\"{SML}\" xmlns:r=\"{REL}\"><dimension ref=\"A1:B1\"/><sheetData><row r=\"1\"><c r=\"A1\" r:id=\"rIdSheet\"><v>1</v></c></row></sheetData></worksheet>"
    );
    let error = edit_a1(&source, &PackageShape::default())
        .expect_err("a planted relationship reference is refused");
    assert_eq!(
        message(&error),
        "invalid XLSX structure: value-only edits refuse relationship reference 'r:id' on 'c'"
    );
}

#[test]
fn an_unknown_element_inside_a_cell_record_is_refused() {
    let source = format!(
        "<worksheet xmlns=\"{SML}\"><dimension ref=\"A1:B1\"/><sheetData><row r=\"1\"><c r=\"A1\"><future/><v>1</v></c></row></sheetData></worksheet>"
    );
    let error =
        edit_a1(&source, &PackageShape::default()).expect_err("an unmodelled cell child refuses");
    assert_eq!(
        message(&error),
        "invalid XLSX structure: value-only edits refuse dependency-bearing or unknown element 'future'"
    );
}

#[test]
fn misplaced_merge_markup_is_now_refused_by_the_raw_parser() {
    // Before change 0657 the value-only vocabulary refused `<mergeCells>`
    // outright. It is now copied through, so the module that models it owns
    // the refusal, with its own typed message.
    let source = format!(
        "<worksheet xmlns=\"{SML}\"><dimension ref=\"A1:B1\"/><mergeCells count=\"1\"><mergeCell ref=\"A1:B1\"/></mergeCells><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData></worksheet>"
    );
    let error = edit_a1(&source, &PackageShape::default())
        .expect_err("merge markup before sheetData is refused");
    assert_eq!(
        message(&error),
        "invalid XLSX structure: worksheet mergeCells appears before sheetData"
    );
}

#[test]
fn two_package_owners_are_still_refused() {
    let shape = PackageShape {
        package_relationships: vec![(
            "rIdSecond",
            format!("{REL}/officeDocument"),
            "xl/workbook.xml",
            false,
        )],
        ..PackageShape::default()
    };
    let error = edit_a1(&worksheet("", ""), &shape).expect_err("a second owner is refused");
    assert_eq!(
        message(&error),
        "OPC package error: Invalid relationship: package has multiple main-document relationships",
        "the package reader owns this refusal; the value-only owner count is defence in depth"
    );
}

#[test]
fn a_protected_sheet_and_a_covering_data_validation_refuse_the_edit() {
    let protected = worksheet("", "<sheetProtection sheet=\"1\" objects=\"1\"/>");
    let error = edit_a1(&protected, &PackageShape::default())
        .expect_err("a protected sheet refuses an edit");
    match error {
        Error::EditBlocked { reason, .. } => assert_eq!(reason, EditBlock::ProtectedSheet),
        other => panic!("expected a protected-sheet refusal, got {other:?}"),
    }

    let validated = worksheet(
        "",
        "<dataValidations count=\"1\"><dataValidation type=\"whole\" sqref=\"A1:A9\"><formula1>0</formula1></dataValidation></dataValidations>",
    );
    let error = edit_a1(&validated, &PackageShape::default())
        .expect_err("a covering validation refuses an edit");
    match error {
        Error::EditBlocked { reason, .. } => assert_eq!(reason, EditBlock::DataValidation),
        other => panic!("expected a data-validation refusal, got {other:?}"),
    }
}

#[test]
fn an_exact_no_op_reproduces_the_producer_worksheet_byte_for_byte() {
    let source = worksheet_with(
        "<sheetPr codeName=\"Sheet1\"/>",
        "<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><cols><col min=\"1\" max=\"2\" width=\"9\"/></cols>",
        "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><extLst><ext xmlns:x14=\"urn:fixture:x14\" uri=\"{FIXTURE}\"><x14:payload/></ext></extLst>",
    );
    let editor = editor(&source, &PackageShape::default()).expect("producer-shaped worksheet");
    let commit = editor
        .edit_many([SheetCellValueEdit::set(
            "Sheet1",
            Address::from_a1("A1").expect("A1"),
            Value::Number(Number::new("1").expect("the value already stored")),
        )])
        .expect("plan the no-op")
        .commit()
        .expect("commit the no-op");
    assert!(!commit.changed(), "an exact no-op changes nothing");
    assert_eq!(
        commit
            .snapshot()
            .sheets()
            .first()
            .expect("one worksheet")
            .source_xml(),
        source.as_bytes(),
        "an exact no-op must reproduce the source bytes exactly"
    );
}

#[test]
fn publishing_a_producer_shaped_package_changes_only_the_expected_members() {
    let shape = PackageShape {
        package_relationships: vec![(
            "rIdApp",
            format!("{REL}/extended-properties"),
            "docProps/app.xml",
            false,
        )],
        worksheet_relationships: vec![(
            "rIdPrn",
            format!("{REL}/printerSettings"),
            "../printerSettings/printerSettings1.bin",
            false,
        )],
        members: vec![
            ("docProps/app.xml", b"<Properties/>".to_vec()),
            ("xl/printerSettings/printerSettings1.bin", vec![7u8; 64]),
        ],
        ..PackageShape::default()
    };
    let source = worksheet("", "<pageSetup r:id=\"rIdPrn\"/>");
    let bytes = package(&source, &shape);
    let editor = SourceBackedEditor::from_read_at(Arc::new(OwnedSource::new(bytes.clone())))
        .expect("producer-shaped package");
    let commit = editor
        .edit_many([SheetCellValueEdit::set(
            "Sheet1",
            Address::from_a1("A1").expect("A1"),
            Value::Number(Number::new("42").expect("numeral")),
        )])
        .expect("plan")
        .commit()
        .expect("commit");
    let mut published = Vec::new();
    editor
        .publish_multi_commit_to_stream(&mut published, &commit)
        .expect("publish");

    // Every member the plan does not name is transferred verbatim, so its
    // bytes still occur in the published archive (these members are stored,
    // not deflated, in this fixture).
    for (name, member) in &shape.members {
        assert!(
            published
                .windows(member.len())
                .any(|window| window == member.as_slice()),
            "member '{name}' must be transferred verbatim"
        );
    }
    let published_text = String::from_utf8_lossy(&published).into_owned();
    assert!(
        published_text.contains("r:id=\"rIdPrn\""),
        "the worksheet's relationship reference must survive publication"
    );
    assert!(
        published_text.contains("<v>42</v>"),
        "the edited value must be published"
    );
    assert!(
        published_text.contains("<c r=\"B1\"><v>2</v></c>"),
        "an unedited neighbour must be published verbatim"
    );
}

#[test]
fn a_copied_subtree_is_still_bounded_and_still_namespace_well_formed() {
    // The old vocabulary bounded nesting implicitly, because it admitted at
    // most five levels. A copied subtree may nest as deeply as the input
    // says, so the validator carries the raw parser's own 256-level bound.
    let mut deep = String::new();
    for _ in 0..300 {
        deep.push_str("<nest>");
    }
    for _ in 0..300 {
        deep.push_str("</nest>");
    }
    let source = worksheet("", &format!("<extLst>{deep}</extLst>"));
    let error =
        edit_a1(&source, &PackageShape::default()).expect_err("a deep copied subtree is bounded");
    assert_eq!(
        message(&error),
        "invalid XLSX structure: value-only XML nesting is too deep"
    );

    // An unbound prefix is malformed, and is refused even where the element
    // would otherwise be copied unread.
    let source = worksheet("", "<extLst><ghost:ext/></extLst>");
    let error = edit_a1(&source, &PackageShape::default())
        .expect_err("an unbound prefix is refused in a copied subtree");
    assert_eq!(
        message(&error),
        "invalid XLSX structure: value-only XML has an unbound element namespace"
    );
}
