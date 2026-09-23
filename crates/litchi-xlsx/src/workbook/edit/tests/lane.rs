//! End-to-end differential proof for the benign `<sheetData>` lane and the
//! reduced post-write readback it admits.
//!
//! Every edit is committed and saved twice from the same source snapshot,
//! once with every lane disabled (which also disables the reduced readback),
//! and the published bytes or the refusal must be identical.

use litchi_opc::PackURI;

use super::super::super::Workbook;
use super::support::styled_workbook;
use crate::raw::worksheet::edit::writer_fault;
use crate::raw::worksheet::lane::corpus::Lcg;
use crate::raw::worksheet::lane::route::{self, Pass};
use crate::{Address, Formula};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn workbook_with_sheet(xml: &str) -> Workbook {
    let baseline = styled_workbook();
    let mut package = baseline.inner.package.clone();
    package
        .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
        .expect("worksheet part")
        .set_blob(xml.as_bytes().to_vec());
    Workbook::from_package(package).expect("generated workbook")
}

fn column_name(column: u32) -> String {
    crate::raw::worksheet::lane::corpus::column_name(column)
}

/// A dense generated worksheet above the validated-store handoff bound.
fn dense_sheet(random: &mut Lcg, rows: u32, columns: u32, suffix: &str) -> String {
    dense_sheet_with(random, rows, columns, "", suffix, false)
}

/// A dense worksheet with `prefix` and `suffix` markup around its
/// `<sheetData>`; `mixed` adds formula and inline-string cells, which the
/// lane never admits, in columns congruent to 3 and 5 modulo 7.
fn dense_sheet_with(
    random: &mut Lcg,
    rows: u32,
    columns: u32,
    prefix: &str,
    suffix: &str,
    mixed: bool,
) -> String {
    let mut body = String::new();
    for row in 1..=rows {
        body.push_str(&format!("<row r=\"{row}\">"));
        for column in 0..columns {
            let address = format!("{}{row}", column_name(column));
            if mixed && column % 7 == 3 {
                body.push_str(&format!(
                    "<c r=\"{address}\"><f>{row}+1</f><v>{}</v></c>",
                    row + 1
                ));
                continue;
            }
            if mixed && column % 7 == 5 {
                body.push_str(&format!(
                    "<c r=\"{address}\" t=\"inlineStr\"><is><t>inline {row}</t></is></c>"
                ));
                continue;
            }
            match random.next() % 10 {
                0 => body.push_str(&format!("<c r=\"{address}\" s=\"1\"><v>{row}.5</v></c>")),
                1 => body.push_str(&format!(
                    "<c r=\"{address}\" t=\"str\"><v>t{row}x{column}</v></c>"
                )),
                2 => body.push_str(&format!("<c r=\"{address}\" t=\"b\"><v>1</v></c>")),
                3 => body.push_str(&format!("<c r=\"{address}\" s=\"1\"/>")),
                _ => body.push_str(&format!(
                    "<c r=\"{address}\"><v>{}</v></c>",
                    row * 1000 + column
                )),
            }
        }
        body.push_str("</row>");
    }
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
         <worksheet xmlns=\"{MAIN}\"><dimension ref=\"A1:{}{rows}\"/>{prefix}<sheetData>{body}</sheetData>{suffix}</worksheet>",
        column_name(columns - 1)
    )
}

#[derive(Debug, Clone)]
enum Operation {
    Number(Address, i32),
    Text(Address, String),
    Bool(Address, bool),
    Formula(Address),
    Clear(Address),
    Remove(Address),
    Style(Address),
    ResetStyle(Address),
}

fn random_operations(random: &mut Lcg, rows: u32, columns: u32, count: usize) -> Vec<Operation> {
    let mut operations = Vec::new();
    for _ in 0..count {
        // Mostly existing cells, sometimes a new cell to the right or below.
        let row = u32::try_from(random.next()).expect("u32") % (rows + 2);
        let column = u32::try_from(random.next()).expect("u32") % (columns + 2);
        let address = Address::at(row, column).expect("address");
        operations.push(match random.next() % 12 {
            0 => Operation::Text(address, format!("edited {row}/{column}")),
            1 => Operation::Bool(address, random.next() % 2 == 0),
            2 => Operation::Clear(address),
            3 => Operation::Remove(address),
            4 => Operation::Style(address),
            5 => Operation::ResetStyle(address),
            6 => Operation::Formula(address),
            _ => Operation::Number(
                address,
                i32::try_from(random.next() % 100_000).expect("i32"),
            ),
        });
    }
    operations
}

fn commit_and_save(source: &Workbook, operations: &[Operation]) -> Result<Vec<u8>, String> {
    let describe = |error: crate::Error| format!("{error:?} / {error}");
    let style = source.styles().map_err(describe)?.get(1).ok_or("style 1")?;
    let mut edit = source.edit().map_err(describe)?;
    {
        let mut sheet = edit.sheet("Sheet1").map_err(describe)?.ok_or("sheet")?;
        for operation in operations {
            match operation {
                Operation::Number(address, value) => sheet.set(*address, *value),
                Operation::Text(address, value) => sheet.set(*address, value.as_str()),
                Operation::Bool(address, value) => sheet.set(*address, *value),
                Operation::Formula(address) => {
                    sheet.set(*address, Formula::new("1+1").expect("formula"))
                },
                Operation::Clear(address) => sheet.clear(*address),
                Operation::Remove(address) => sheet.remove(*address),
                Operation::Style(address) => sheet.style(*address, &style),
                Operation::ResetStyle(address) => sheet.reset_style(*address),
            }
            .map_err(describe)?;
        }
    }
    let commit = edit.commit().map_err(describe)?;
    let bytes = commit.workbook().to_bytes().map_err(describe)?;
    // The published snapshot must read back the same cells on both routes.
    let sheet = commit
        .workbook()
        .sheet("Sheet1")
        .map_err(describe)?
        .ok_or("published sheet")?;
    let cells = sheet
        .cells("A1:XFD1048576")
        .map_err(describe)?
        .map(|(address, cell)| format!("{address:?}={cell:?}"))
        .collect::<Vec<_>>()
        .join(";");
    let mut outcome = bytes;
    outcome.extend_from_slice(cells.as_bytes());
    Ok(outcome)
}

/// Commit and save on both routes; return whether the reduced readback ran.
fn assert_parity(source_xml: &str, operations: &[Operation]) -> bool {
    let source = workbook_with_sheet(source_xml);
    route::reset();
    let fast = commit_and_save(&source, operations);
    let reduced = route::admitted(Pass::Readback) > 0;
    let slow_source = workbook_with_sheet(source_xml);
    let slow = route::without_lane(|| commit_and_save(&slow_source, operations));
    assert_eq!(route::admitted(Pass::Readback), usize::from(reduced));
    match (&fast, &slow) {
        (Ok(fast), Ok(slow)) => assert!(fast == slow, "published bytes differ for {operations:?}"),
        (Err(fast), Err(slow)) => assert_eq!(fast, slow, "{operations:?}"),
        _ => panic!("routes disagree for {operations:?}: {fast:?} vs {slow:?}"),
    }
    reduced
}

#[test]
fn random_edits_on_dense_sheets_publish_identical_bytes() {
    let mut random = Lcg(0x0744_0001);
    let mut reduced = 0usize;
    let mut refused = 0usize;
    for round in 0..40 {
        let rows = 70;
        let columns = 64;
        let suffix = if round % 4 == 0 {
            "<mergeCells count=\"1\"><mergeCell ref=\"B2:C3\"/></mergeCells>"
        } else {
            ""
        };
        let xml = dense_sheet(&mut random, rows, columns, suffix);
        let count = 1 + usize::try_from(random.next() % 6).expect("usize");
        let mut operations = random_operations(&mut random, rows, columns, count);
        if !suffix.is_empty() && round % 8 == 0 {
            // A covered merge follower must keep its typed refusal.
            operations.push(Operation::Number(Address::at(2, 2).expect("C3"), 1));
        }
        reduced += usize::from(assert_parity(&xml, &operations));
        let source = workbook_with_sheet(&xml);
        refused += usize::from(commit_and_save(&source, &operations).is_err());
    }
    assert!(reduced >= 5, "reduced readback ran {reduced} times");
    assert!(refused >= 1, "refusals were not exercised");
}

#[test]
fn value_updates_above_the_handoff_bound_take_the_reduced_readback() {
    let mut random = Lcg(7);
    let xml = dense_sheet(&mut random, 80, 60, "");
    let operations = [
        Operation::Number(Address::at(0, 0).expect("A1"), 42),
        Operation::Clear(Address::at(40, 30).expect("AE41")),
        Operation::Style(Address::at(41, 30).expect("AE42")),
        Operation::Bool(Address::at(79, 59).expect("last"), true),
    ];
    assert!(assert_parity(&xml, &operations));
    // The eager writer stores edited text as an inline string, which the
    // lane does not admit, so the complete readback verifies it.
    let text = [Operation::Text(
        Address::at(40, 30).expect("AE41"),
        "middle".to_owned(),
    )];
    assert!(!assert_parity(&xml, &text));
}

#[test]
fn small_sheets_keep_the_complete_readback_and_its_handoff() {
    let mut random = Lcg(11);
    let xml = dense_sheet(&mut random, 10, 10, "");
    let operations = [Operation::Number(Address::at(2, 2).expect("C3"), 5)];
    assert!(!assert_parity(&xml, &operations));
}

#[test]
fn markup_compatibility_and_noncompact_sources_keep_the_complete_readback() {
    let mut random = Lcg(13);
    let xml = dense_sheet(&mut random, 80, 60, "");
    let operations = [Operation::Number(Address::at(0, 0).expect("A1"), 42)];
    let with_mce = xml.replacen(
        &format!("<worksheet xmlns=\"{MAIN}\">"),
        &format!(
            "<worksheet xmlns=\"{MAIN}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">"
        ),
        1,
    );
    assert!(!assert_parity(&with_mce, &operations));
    let noncompact = xml.replacen("?>", "?>\r\n", 1);
    assert!(!assert_parity(&noncompact, &operations));
    let formatted = xml.replace("<row ", "\n<row ");
    assert!(!assert_parity(&formatted, &operations));
}

#[test]
fn refused_edits_keep_their_refusal_on_the_reduced_route() {
    let mut random = Lcg(17);
    let xml = dense_sheet(
        &mut random,
        80,
        60,
        "<mergeCells count=\"1\"><mergeCell ref=\"B2:C3\"/></mergeCells>",
    );
    // Covered merge follower.
    let operations = [Operation::Number(Address::at(1, 2).expect("C2"), 1)];
    assert_parity(&xml, &operations);
    let source = workbook_with_sheet(&xml);
    assert!(commit_and_save(&source, &operations).is_err());
}

#[test]
fn a_foreign_sheet_data_before_the_body_keeps_the_complete_readback() {
    let operations = [Operation::Number(Address::at(40, 0).expect("A41"), 42)];
    for foreign in [
        "<x:sheetData xmlns:x=\"urn:foreign\"><row r=\"1\"/></x:sheetData>",
        "<sheetData xmlns=\"urn:foreign\"><row r=\"1\"/></sheetData>",
    ] {
        // The worksheet's own body holds formulas and inline strings, so only
        // the foreign element's body is lane-benign.
        let mut random = Lcg(31);
        let mixed = dense_sheet_with(&mut random, 80, 60, foreign, "", true);
        assert!(!assert_parity(&mixed, &operations), "{foreign}");
        // A benign worksheet body behind the same foreign element still
        // verifies through the reduced readback.
        let mut random = Lcg(37);
        let benign = dense_sheet_with(&mut random, 80, 60, foreign, "", false);
        assert!(assert_parity(&benign, &operations), "{foreign}");
    }
}

#[test]
fn worksheets_above_the_web_reader_limit_never_take_the_reduced_readback() {
    let mut random = Lcg(41);
    // A comment survives compaction unchanged and pushes the part past the
    // limit, so the edit is refused at the web-extension step on both routes.
    let padding = format!("<!--{}-->", "x".repeat(crate::raw::web::MAX_XML_BYTES));
    let xml = dense_sheet(&mut random, 80, 60, &padding);
    let operations = [Operation::Number(Address::at(40, 30).expect("AE41"), 42)];
    assert!(!assert_parity(&xml, &operations));
    let source = workbook_with_sheet(&xml);
    let error = commit_and_save(&source, &operations).expect_err("above the web limit");
    assert!(error.contains("web-extension parser limit"), "{error}");
}

#[test]
fn a_changed_cell_reemitted_into_an_omitted_run_keeps_the_complete_readback() {
    let mut random = Lcg(43);
    let xml = dense_sheet(&mut random, 80, 60, "");
    let operations = [Operation::Number(Address::at(40, 30).expect("AE41"), 42)];
    // Without the fault this edit verifies through the reduced readback.
    assert!(assert_parity(&xml, &operations));
    // The faulty writer keeps the old AE41 inside the preceding omitted run
    // and writes the new one after it. Only the complete parse sees both.
    let (fast, reduced) = writer_fault::with_reemitted_changed_cells(|| {
        route::reset();
        let result = commit_and_save(&workbook_with_sheet(&xml), &operations);
        (result, route::admitted(Pass::Readback))
    });
    let slow = writer_fault::with_reemitted_changed_cells(|| {
        route::without_lane(|| commit_and_save(&workbook_with_sheet(&xml), &operations))
    });
    assert_eq!(
        reduced, 0,
        "the collision check must refuse the reduced store"
    );
    assert_eq!(fast, slow);
    let error = fast.expect_err("the published row holds two AE41 records");
    assert!(error.contains("duplicate worksheet cell"), "{error}");
}

#[test]
fn byte_order_marked_worksheets_publish_identical_outcomes() {
    let mut random = Lcg(47);
    let xml = format!("\u{feff}{}", dense_sheet(&mut random, 80, 60, ""));
    for operations in [
        vec![Operation::Number(Address::at(0, 0).expect("A1"), 42)],
        vec![Operation::Text(
            Address::at(3, 3).expect("D4"),
            "bom".to_owned(),
        )],
    ] {
        assert!(!assert_parity(&xml, &operations));
    }
}
