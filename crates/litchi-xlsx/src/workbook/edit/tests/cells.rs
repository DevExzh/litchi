//! Cell, merge, style, and package-edit regression tests.

use super::super::codec::calculation_chain_removal;
use super::super::*;
use super::support::{merged_workbook, part_text, styled_workbook, two_sheet_workbook};

use crate::cell::{Number, Value};
use crate::color::Rgb;
use crate::error::MergeEditBlock;
use crate::formula::Formula;
use crate::style::AlignmentChange;
use crate::style::BorderChange;
use crate::style::FillChange;
use crate::style::ProtectionChange;
use crate::style::Underline;
use crate::style::cell_format::{
    Color, Dir, Fill, GradientFill, GradientKind, GradientStop, Line, Pattern, PatternFill, Side,
    Tint,
};
use crate::style::stylesheet::alignment::{Horizontal, Indent, Reading, Rotation, Vertical};
use crate::style::stylesheet::{FontColor, FontColorKind, Scheme, Script};
use crate::{FontChange, StyleState};
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{BlobPart, PackURI, SourceLimits, TargetMode};
use std::mem::size_of;

fn table_workbook(with_existing_merge: bool) -> crate::Package {
    let baseline = crate::Package::new().expect("baseline");
    let mut package = baseline.inner.package.clone();
    let merge_cells = if with_existing_merge {
        r#"<mergeCells count="1"><mergeCell ref="B2:C3"/></mergeCells>"#
    } else {
        ""
    };
    package
        .get_part_mut(&baseline.inner.sheets[0].part_uri)
        .expect("worksheet part")
        .set_blob(
            format!(
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/>{merge_cells}</worksheet>"#
            )
            .into_bytes(),
        );
    package
        .get_part_mut(&baseline.inner.sheets[0].part_uri)
        .expect("worksheet part")
        .rels_mut()
        .try_add_relationship(
            rt::TABLE.to_owned(),
            "../tables/table1.xml".to_owned(),
            "rIdTable".to_owned(),
            TargetMode::Internal,
        )
        .expect("table relationship");
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/tables/table1.xml").expect("table URI"),
            litchi_opc::constants::content_type::SML_TABLE.to_owned(),
            br#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="B2:C3"><tableColumns count="2"><tableColumn id="1" name="One"/><tableColumn id="2" name="Two"/></tableColumns></table>"#.to_vec(),
        )))
        .expect("table part");
    crate::Package::from_opc(package).expect("table workbook")
}

fn number_format_workbook() -> crate::Package {
    let baseline = crate::Package::new().expect("baseline");
    let mut package = baseline.inner.package.clone();
    package
        .get_part_mut(&PackURI::new("/xl/styles.xml").expect("styles URI"))
        .expect("styles part")
        .set_blob(
            br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:f="urn:future"><numFmts count="1"><numFmt f:keep="yes" numFmtId="200" formatCode="0.000 &quot;kg&quot;"><f:format-future/></numFmt></numFmts><fonts count="1"><font/></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" f:xf="yes"><f:xf-future/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#
                .to_vec(),
        );
    crate::Package::from_opc(package).expect("number-format workbook")
}

fn fill_workbook() -> crate::Package {
    let baseline = crate::Package::new().expect("baseline");
    let mut package = baseline.inner.package.clone();
    package
        .get_part_mut(&PackURI::new("/xl/styles.xml").expect("styles URI"))
        .expect("styles part")
        .set_blob(
            br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:f="urn:future"><fonts count="1"><font/></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill f:target-fill="yes"><patternFill f:target-pattern="yes" patternType="solid"><fgColor rgb="FFFFFF00"/><f:fill-future/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" f:target-xf="yes"><f:xf-future/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#
                .to_vec(),
        );
    package
        .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
        .expect("worksheet part")
        .set_blob(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row></sheetData></worksheet>"#.to_vec(),
        );
    crate::Package::from_opc(package).expect("fill workbook")
}

fn border_workbook() -> crate::Package {
    let baseline = crate::Package::new().expect("baseline workbook");
    let mut package = baseline.inner.package.clone();
    package
        .get_part_mut(&PackURI::new("/xl/styles.xml").expect("styles URI"))
        .expect("styles part")
        .set_blob(
            br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:f="urn:future"><fonts count="1"><font/></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="2"><border/><border f:keep="yes" diagonalUp="1" diagonalDown="1" outline="1"><left style="thin"><color theme="4" tint=".5"/></left><top style="double"/><diagonal style="dashed"><color indexed="10"/></diagonal><f:future/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1" f:xf="yes"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0"/></cellStyles></styleSheet>"#
                .to_vec(),
        );
    package
        .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
        .expect("worksheet part")
        .set_blob(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row></sheetData></worksheet>"#
                .to_vec(),
        );
    crate::Package::from_opc(package).expect("border workbook")
}

fn low_part_workbook() -> crate::Package {
    let baseline = crate::Package::new().expect("baseline workbook");
    let baseline_bytes = baseline.to_bytes().expect("baseline bytes");
    let maximum_part = baseline
        .inner
        .package
        .try_iter_parts()
        .map(|part| part.expect("baseline part").blob().len())
        .max()
        .expect("baseline parts")
        .saturating_add(4096);
    let limits = SourceLimits::builder()
        .max_part_bytes(maximum_part as u64)
        .expect("part limit")
        .build()
        .expect("source limits");
    crate::Package::from_bytes_with_options(
        &baseline_bytes,
        crate::OpenOptions::new().source_limits(limits),
    )
    .expect("limited workbook")
}

fn lexical_gradient_workbook() -> crate::Package {
    let baseline = crate::Package::new().expect("baseline");
    let mut package = baseline.inner.package.clone();
    package
        .get_part_mut(&PackURI::new("/xl/styles.xml").expect("styles URI"))
        .expect("styles part")
        .set_blob(
            br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:f="urn:future"><fonts count="1"><font/></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill f:keep="yes"><gradientFill degree='045.0' type='linear' f:gradient='yes'><stop position='0.0'><color rgb='FF112233'/></stop><stop position='1.0'><color theme='5' tint='.5'/></stop></gradientFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#
                .to_vec(),
        );
    package
        .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
        .expect("worksheet part")
        .set_blob(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row></sheetData></worksheet>"#.to_vec(),
        );
    crate::Package::from_opc(package).expect("lexical gradient workbook")
}

fn authored_gradient_stops() -> Box<[GradientStop]> {
    vec![
        GradientStop::new(
            0.0,
            Color::rgb(0x11, 0x22, 0x33).with_tint(Tint::new(-0.25).expect("RGB tint")),
        )
        .expect("RGB stop"),
        GradientStop::new(
            0.5,
            Color::theme(5).with_tint(Tint::new(0.5).expect("theme tint")),
        )
        .expect("theme stop"),
        GradientStop::new(
            1.0,
            Color::indexed(12).with_tint(Tint::new(0.125).expect("indexed tint")),
        )
        .expect("indexed stop"),
    ]
    .into_boxed_slice()
}

fn authored_linear_gradient() -> GradientFill {
    GradientFill::linear(135.0, authored_gradient_stops()).expect("linear gradient")
}

fn authored_path_gradient() -> GradientFill {
    GradientFill::path(0.1, 0.2, 0.3, 0.4, authored_gradient_stops()).expect("path gradient")
}

#[test]
fn cell_crud_is_atomic_reversible_and_source_preserving() {
    let source = crate::Package::new().expect("source workbook");
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    {
        let mut sheet = edit.sheet("Sheet1").expect("sheet lookup").expect("sheet");
        sheet
            .set("A1", "hello")
            .and_then(|sheet| sheet.set("B2", 42_i32))
            .and_then(|sheet| sheet.set("C3", Formula::new("B2*2").expect("formula")))
            .expect("cell changes");
    }
    let committed = edit.commit().expect("commit");
    assert!(
        committed.package().inner.sheets[0].cells.get().is_some(),
        "the validated changed worksheet should be ready for its first public read"
    );
    assert_eq!(
        source.to_bytes().expect("source remains valid"),
        source_bytes
    );
    assert_eq!(committed.patch().operation_count(), 3);

    let book = committed.package();
    let sheet = book.sheet("Sheet1").expect("lookup").expect("sheet");
    assert!(matches!(
        sheet.cell("A1").expect("A1").stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "hello"
    ));
    assert!(matches!(
        sheet.cell("B2").expect("B2").stored(),
        Some(Cell::Value(Value::Number(number))) if number == &Number::new("42").expect("number")
    ));
    assert!(matches!(
        sheet.cell("C3").expect("C3").stored(),
        Some(Cell::Formula(_))
    ));
    let extents = sheet.extents().expect("committed extents");
    assert_eq!(extents.declared().map(Rect::a1).as_deref(), Some("A1:C3"));
    assert_eq!(extents.content().map(Rect::a1).as_deref(), Some("A1:C3"));

    let restored = book
        .apply(&committed.patch().inverse())
        .expect("inverse patch");
    assert_eq!(
        restored.package().to_bytes().expect("restored"),
        source_bytes
    );
    assert!(matches!(
        source.apply(committed.patch()),
        Ok(applied) if applied.package().sheet("Sheet1").expect("lookup").expect("sheet").cell("A1").expect("cell").stored().is_some()
    ));
    assert!(matches!(
        book.apply(committed.patch()),
        Err(Error::PatchConflict { .. })
    ));
}

#[test]
fn plain_cell_rewrite_preserves_neighbor_bytes_and_reopens_typed_formulas() {
    let baseline = crate::Package::new().expect("baseline");
    let mut package = baseline.inner.package.clone();
    let source_sheet = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>SUM(A1,1)</f><v>2</v></c><c r="C1" t="inlineStr"><is><t>typed</t></is></c><x:c r="D1"><x:f>SUM(A1,3)</x:f><x:v>4</x:v></x:c></row></sheetData></worksheet>"#;
    package
        .get_part_mut(&baseline.inner.sheets[0].part_uri)
        .expect("worksheet part")
        .set_blob(source_sheet.to_vec());
    let source = crate::Package::from_opc(package).expect("source workbook");
    let source_sheet_xml = part_text(&source, "/xl/worksheets/sheet1.xml").to_owned();
    let untouched_formula = r#"<c r="B1"><f>SUM(A1,1)</f><v>2</v></c>"#;
    assert_eq!(
        source_sheet_xml.matches(untouched_formula).count(),
        1,
        "fixture must contain one untouched plain formula cell"
    );

    let mut edit = source.edit().expect("edit");
    edit.sheet(0usize)
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", "edited")
        .expect("plain cell edit");
    let committed = edit.commit().expect("commit");
    let changed_sheet_xml = part_text(committed.package(), "/xl/worksheets/sheet1.xml");
    assert_eq!(
        changed_sheet_xml.matches(untouched_formula).count(),
        1,
        "untouched plain formula cell should retain its exact source bytes"
    );
    assert!(changed_sheet_xml.contains(r#"<c r="A1" t="inlineStr">"#));
    assert!(changed_sheet_xml.contains(r#"<c r="C1" t="inlineStr"><is><t>typed</t></is></c>"#));
    assert!(changed_sheet_xml.contains(r#"<x:c r="D1"><x:f>SUM(A1,3)</x:f><x:v>4</x:v></x:c>"#));

    let reopened = crate::Package::from_bytes(&committed.package().to_bytes().expect("serialized"))
        .expect("reopened workbook");
    let sheet = reopened
        .sheet(0usize)
        .expect("sheet lookup")
        .expect("worksheet");
    assert!(matches!(
        sheet.cell("A1").expect("A1").stored(),
        Some(Cell::Value(Value::Text(value))) if value.as_str() == "edited"
    ));
    assert!(matches!(
        sheet.cell("B1").expect("B1").stored(),
        Some(Cell::Formula(formula)) if formula.text() == "SUM(A1,1)"
    ));
    assert!(matches!(
        sheet.cell("C1").expect("C1").stored(),
        Some(Cell::Value(Value::Text(value))) if value.as_str() == "typed"
    ));
    assert!(matches!(
        sheet.cell("D1").expect("D1").stored(),
        Some(Cell::Formula(formula)) if formula.text() == "SUM(A1,3)"
    ));
    let reopened_sheet_xml = part_text(&reopened, "/xl/worksheets/sheet1.xml");
    assert!(reopened_sheet_xml.contains(r#"<x:c r="D1"><x:f>SUM(A1,3)</x:f>"#));
}

#[test]
fn commit_adopts_only_the_changed_worksheet_store() {
    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    assert!(
        source
            .inner
            .sheets
            .iter()
            .all(|sheet| sheet.cells.get().is_none())
    );

    let mut edit = source.edit().expect("edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("sheet")
        .set("A1", 7_i32)
        .expect("cell edit");
    let committed = edit.commit().expect("commit");

    assert!(committed.package().inner.sheets[0].cells.get().is_some());
    assert!(committed.package().inner.sheets[1].cells.get().is_none());
}

#[test]
fn commit_does_not_retain_an_oversized_validated_store() {
    let source = crate::Package::new().expect("source workbook");
    let mut edit = source.edit().expect("edit");
    {
        let mut sheet = edit.sheet("Sheet1").expect("sheet lookup").expect("sheet");
        for index in 0..4_097_u32 {
            sheet
                .set(
                    (index / 64, index % 64),
                    i32::try_from(index).expect("bounded value"),
                )
                .expect("cell edit");
        }
    }

    let committed = edit.commit().expect("commit");
    assert!(committed.package().inner.sheets[0].cells.get().is_none());
}

#[test]
fn large_plain_worksheet_sparse_edit_reopens_and_preserves_store_policy() {
    let source = crate::Package::new().expect("source workbook");
    let mut create = source.edit().expect("create edit");
    {
        let mut sheet = create
            .sheet("Sheet1")
            .expect("sheet lookup")
            .expect("sheet");
        for index in 0..4_097_u32 {
            sheet
                .set(
                    (index / 64, index % 64),
                    i32::try_from(index).expect("bounded value"),
                )
                .expect("cell edit");
        }
    }
    let created = create.commit().expect("large worksheet commit");
    let reopened = crate::Package::from_bytes(&created.package().to_bytes().expect("large bytes"))
        .expect("large worksheet reopen");

    let mut edit = reopened.edit().expect("sparse edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("sheet")
        .set("A1", 42_i32)
        .expect("changed value");
    let committed = edit.commit().expect("sparse commit");
    assert!(
        committed.package().inner.sheets[0].cells.get().is_none(),
        "large sparse verification must not retain a second full worksheet store"
    );
    let sheet = committed
        .package()
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("sheet");
    assert!(matches!(
        sheet.cell("A1").expect("updated cell").stored(),
        Some(Cell::Value(Value::Number(value))) if value.as_str() == "42"
    ));
    let final_bytes = committed.package().to_bytes().expect("final bytes");
    let final_reopen = crate::Package::from_bytes(&final_bytes).expect("final reopen");
    assert!(matches!(
        final_reopen
            .sheet("Sheet1")
            .expect("sheet lookup")
            .expect("sheet")
            .cell("A1")
            .expect("updated cell")
            .stored(),
        Some(Cell::Value(Value::Number(value))) if value.as_str() == "42"
    ));
}
#[test]
fn merged_range_crud_is_sparse_safe_reversible_and_composable() {
    let source = merged_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let source_sheet = source.sheet("Sheet1").expect("lookup").expect("sheet");
    assert_eq!(
        source_sheet
            .merges()
            .expect("merged ranges")
            .map(Rect::a1)
            .collect::<Vec<_>>(),
        ["A1:B2"]
    );
    assert!(matches!(
        source_sheet.cell("A1").expect("anchor"),
        crate::cell::View::Stored(Cell::Value(_))
    ));
    assert!(matches!(
        source_sheet.cell("B2").expect("covered"),
        crate::cell::View::Covered(range) if range == Rect::from_a1("A1:B2").expect("range")
    ));

    let mut edit = source.edit().expect("edit");
    edit.sheet("Sheet1")
        .expect("lookup")
        .expect("sheet")
        .unmerge("B2")
        .and_then(|sheet| sheet.set("B2", "revealed"))
        .and_then(|sheet| sheet.merge("C3:D4"))
        .expect("merged-range changes");
    let committed = edit.commit().expect("commit");
    assert_eq!(committed.patch().operation_count(), 3);
    let sheet = committed
        .package()
        .sheet("Sheet1")
        .expect("lookup")
        .expect("sheet");
    assert_eq!(
        sheet
            .merges()
            .expect("merged ranges")
            .map(Rect::a1)
            .collect::<Vec<_>>(),
        ["C3:D4"]
    );
    assert!(matches!(
        sheet.cell("B2").expect("uncovered cell").stored(),
        Some(Cell::Value(Value::Text(value))) if value.as_str() == "revealed"
    ));
    assert!(matches!(
        sheet.cell("D4").expect("covered cell"),
        crate::cell::View::Covered(range) if range == Rect::from_a1("C3:D4").expect("range")
    ));
    let restored = committed
        .package()
        .apply(&committed.patch().inverse())
        .expect("inverse");
    assert_eq!(
        restored.package().to_bytes().expect("restored bytes"),
        source_bytes
    );

    let mut blocked = source.edit().expect("blocked edit");
    blocked
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .merge("D1:E2")
        .expect("plan merge");
    assert!(matches!(
        blocked.commit(),
        Err(Error::MergeEditBlocked {
            reason: MergeEditBlock::FollowerContent { address },
            ..
        }) if address == Address::from_a1("E2").expect("address")
    ));

    let mut cleared = source.edit().expect("clear edit");
    cleared
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .clear("E2")
        .and_then(|sheet| sheet.merge("D1:E2"))
        .expect("clear then merge");
    let cleared = cleared.commit().expect("safe merge");
    assert!(matches!(
        cleared
            .package()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .cell("E2")
            .expect("covered"),
        crate::cell::View::Covered(range) if range == Rect::from_a1("D1:E2").expect("range")
    ));

    let mut merge = source.edit().expect("independent merge");
    merge
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .merge("D1:E2")
        .expect("merge");
    let mut clear = source.edit().expect("independent clear");
    clear
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .clear("E2")
        .expect("clear");
    merge.join(clear).expect("clear makes the merge safe");
    assert!(matches!(
        merge
            .commit()
            .expect("joined clear and merge")
            .package()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .cell("E2")
            .expect("covered"),
        crate::cell::View::Covered(_)
    ));

    let mut merge = source.edit().expect("independent merge");
    merge
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .merge("C3:D4")
        .expect("merge");
    let mut write = source.edit().expect("independent follower write");
    write
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .set("D4", "hidden")
        .expect("write");
    assert!(matches!(
        merge.join(write).expect_err("follower content must conflict").failure(),
        JoinFailure::Overlap(conflicts)
            if conflicts.conflicts().iter().any(|conflict| conflict.merges().is_some())
    ));

    let mut overlap = source.edit().expect("overlap edit");
    overlap
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .merge("B2:C3")
        .expect("plan overlap");
    assert!(matches!(
        overlap.commit(),
        Err(Error::MergeEditBlocked {
            reason: MergeEditBlock::Overlap { existing },
            ..
        }) if existing == Rect::from_a1("A1:B2").expect("existing")
    ));

    let mut created = source.edit().expect("create edit");
    created
        .add("Merged")
        .expect("new sheet")
        .set("A1", "anchor")
        .and_then(|sheet| sheet.merge("A1:C2"))
        .expect("new merged range");
    let created = created.commit().expect("create merged sheet");
    assert!(matches!(
        created
            .package()
            .sheet("Merged")
            .expect("lookup")
            .expect("sheet")
            .cell("C2")
            .expect("covered"),
        crate::cell::View::Covered(range) if range == Rect::from_a1("A1:C2").expect("range")
    ));

    let mut left = source.edit().expect("left edit");
    left.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .merge("C3:D4")
        .expect("left merge");
    let mut right = source.edit().expect("right edit");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .merge("F1:G2")
        .expect("right merge");
    left.join(right).expect("disjoint merge join");
    assert_eq!(
        left.commit()
            .expect("joined commit")
            .package()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .merges()
            .expect("merges")
            .count(),
        3
    );

    let mut left = source.edit().expect("left overlap");
    left.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .merge("C3:D4")
        .expect("left merge");
    let mut right = source.edit().expect("right overlap");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .merge("D4:E5")
        .expect("right merge");
    let error = left.join(right).expect_err("overlapping joins must fail");
    assert!(matches!(
        error.failure(),
        JoinFailure::Overlap(conflicts)
            if conflicts.conflicts().iter().any(|conflict| conflict.merges().is_some())
    ));

    let mut left = source.edit().expect("left unmerge");
    left.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .unmerge("A1")
        .expect("left unmerge");
    let mut right = source.edit().expect("right unmerge");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .unmerge("B2")
        .expect("right unmerge");
    let error = left
        .join(right)
        .expect_err("two selectors for one merge must conflict");
    assert!(matches!(
        error.failure(),
        JoinFailure::Overlap(conflicts)
            if conflicts.conflicts().iter().any(|conflict| {
                conflict.merges().is_some_and(|ranges| {
                    ranges == [Rect::from_a1("A1:B2").expect("range")]
                })
            })
    ));

    let mut missing = source.edit().expect("missing unmerge");
    missing
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .unmerge("H8")
        .expect("missing range is a no-op");
    assert!(missing.commit().expect("no-op commit").patch().is_noop());

    let maximum = source.retained_bytes().saturating_mul(2);
    let mut bounded_missing = source
        .edit_with_options(EditOptions::default().with_max_retained_bytes(maximum))
        .expect("bounded edit");
    bounded_missing
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .unmerge("H8")
        .expect("bounded missing range is a no-op");
    assert!(
        bounded_missing
            .commit()
            .expect("bounded no-op commit")
            .patch()
            .is_noop()
    );
}

#[test]
fn merge_addition_refuses_typed_table_overlap() {
    let source = table_workbook(false);
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    edit.sheet(0usize)
        .expect("sheet lookup")
        .expect("sheet")
        .merge("A1:B2")
        .expect("stage merge");
    assert!(matches!(
        edit.commit(),
        Err(Error::MergeEditBlocked {
            range,
            reason: MergeEditBlock::TableOverlap { table },
            ..
        }) if range == Rect::from_a1("A1:B2").expect("merge range")
            && table == Rect::from_a1("B2:C3").expect("table range")
    ));
    assert_eq!(
        source.to_bytes().expect("source remains unchanged"),
        source_bytes
    );
}

#[test]
fn exact_existing_merge_remains_a_noop_inside_a_table() {
    let source = table_workbook(true);
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    edit.sheet(0usize)
        .expect("sheet lookup")
        .expect("sheet")
        .merge("B2:C3")
        .expect("stage exact merge");
    let committed = edit.commit().expect("exact merge no-op");
    assert_eq!(committed.patch().operation_count(), 0);
    assert_eq!(
        committed.package().to_bytes().expect("no-op bytes"),
        source_bytes
    );
}

#[test]
fn merged_range_reads_share_one_snapshot_without_public_locks() {
    let workbook = Arc::new(merged_workbook());
    std::thread::scope(|scope| {
        for _ in 0..8 {
            let workbook = Arc::clone(&workbook);
            scope.spawn(move || {
                let sheet = workbook.sheet(0usize).expect("lookup").expect("sheet");
                assert!(matches!(
                    sheet.cell("B2").expect("covered"),
                    crate::cell::View::Covered(range)
                        if range == Rect::from_a1("A1:B2").expect("range")
                ));
                assert_eq!(sheet.merges().expect("merges").count(), 1);
            });
        }
    });
}

#[test]
fn clear_and_remove_have_distinct_missing_and_empty_semantics() {
    let source = crate::Package::new().expect("source workbook");
    let mut edit = source.edit().expect("edit");
    edit.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .set("A1", "value")
        .expect("set");
    let first = edit.commit().expect("first commit").into_package();

    let mut edit = first.edit().expect("edit");
    let mut sheet = edit.sheet(0usize).expect("lookup").expect("sheet");
    sheet.clear("A1").expect("clear");
    sheet.clear("B1").expect("clear missing");
    let cleared = edit.commit().expect("clear commit");
    assert_eq!(cleared.patch().operation_count(), 1);
    let sheet = cleared
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet");
    assert!(matches!(
        sheet.cell("A1").expect("cell").stored(),
        Some(Cell::Empty)
    ));
    assert!(sheet.cell("B1").expect("cell").is_missing());

    let mut edit = cleared.package().edit().expect("edit");
    edit.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .remove("A1")
        .expect("remove");
    let removed = edit.commit().expect("remove commit");
    assert!(
        removed
            .package()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .cell("A1")
            .expect("cell")
            .is_missing()
    );
}

#[test]
fn clearing_a_cell_created_in_the_same_transaction_keeps_an_empty_record() {
    let source = crate::Package::new().expect("source workbook");
    let mut edit = source.edit().expect("edit");
    let mut sheet = edit.sheet(0usize).expect("lookup").expect("sheet");
    sheet.set("A1", "temporary").expect("set");
    sheet.clear("A1").expect("clear");
    let committed = edit.commit().expect("commit");
    assert!(matches!(
        committed
            .package()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .cell("A1")
            .expect("cell")
            .stored(),
        Some(Cell::Empty)
    ));
}

#[test]
fn copying_an_explicit_empty_cell_creates_the_target_record() {
    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut create = source.edit().expect("create edit");
    create
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", "temporary")
        .and_then(|sheet| sheet.clear("A1"))
        .expect("explicit empty");
    let source = create.commit().expect("empty commit").into_package();

    let mut copy = source.edit().expect("copy edit");
    copy.copy_cells("Sheet1", "A1", "Sheet2", "B2")
        .expect("copy explicit empty");
    let committed = copy.commit().expect("copy commit");
    assert!(matches!(
        committed
            .package()
            .sheet("Sheet2")
            .expect("sheet lookup")
            .expect("worksheet")
            .cell("B2")
            .expect("target cell")
            .stored(),
        Some(Cell::Empty)
    ));
}

#[test]
fn calculation_chain_removal_is_atomic_and_reversible() {
    let baseline = crate::Package::new().expect("baseline");
    let mut package = baseline.inner.package.clone();
    let chain_uri = PackURI::new("/xl/calcChain.xml").expect("chain URI");
    package
            .try_add_part(Box::new(BlobPart::new(
                chain_uri.clone(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"
                    .to_owned(),
                br#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="1"/></calcChain>"#.to_vec(),
            )))
            .expect("chain part");
    package
        .get_part_mut(&baseline.inner.workbook_uri)
        .expect("workbook part")
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::CALC_CHAIN.to_owned(),
            "calcChain.xml".to_owned(),
            "rId3".to_owned(),
            TargetMode::Internal,
        )
        .expect("chain relationship");
    let source = crate::Package::from_opc(package).expect("workbook with chain");
    let source_bytes = source.to_bytes().expect("source bytes");
    let workbook_before = source
        .inner
        .package
        .get_part(&source.inner.workbook_uri)
        .expect("workbook part")
        .blob_arc();

    let mut visibility = source.edit().expect("visibility edit");
    visibility
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .row(1)
        .expect("row 2")
        .hide();
    let visibility = visibility.commit().expect("visibility commit");
    assert!(visibility.patch().graph.is_empty());
    assert!(
        visibility
            .package()
            .inner
            .package
            .get_part(&chain_uri)
            .is_ok()
    );
    assert_eq!(
        visibility
            .package()
            .inner
            .package
            .get_part(&source.inner.workbook_uri)
            .expect("unchanged workbook part")
            .blob(),
        workbook_before.as_slice()
    );

    let mut edit = source.edit().expect("edit");
    edit.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .set("A1", 7_i32)
        .expect("set");
    let committed = edit.commit().expect("commit");
    assert!(
        committed
            .package()
            .inner
            .package
            .get_part(&chain_uri)
            .is_err()
    );
    assert!(
        calculation_chain_removal(committed.package())
            .expect("chain query")
            .is_empty()
    );

    let restored = committed
        .package()
        .apply(&committed.patch().inverse())
        .expect("inverse");
    assert_eq!(
        restored.package().to_bytes().expect("restored"),
        source_bytes
    );
    assert!(
        !calculation_chain_removal(restored.package())
            .expect("restored chain")
            .is_empty()
    );

    let mut shared_target = source.inner.package.clone();
    let mut referrer = BlobPart::new(
        PackURI::new("/xl/custom.xml").expect("custom URI"),
        "application/xml".to_owned(),
        b"<custom/>".to_vec(),
    );
    referrer
        .rels_mut()
        .try_add_relationship(
            "urn:litchi:test-reference".to_owned(),
            "calcChain.xml".to_owned(),
            "rId1".to_owned(),
            TargetMode::Internal,
        )
        .expect("extra incoming relationship");
    shared_target
        .try_add_part(Box::new(referrer))
        .expect("referrer part");
    let shared_target = crate::Package::from_opc(shared_target).expect("shared target workbook");
    assert!(matches!(
        calculation_chain_removal(&shared_target),
        Err(Error::Invalid(message)) if message.contains("another incoming relationship")
    ));

    let mut outbound = source.inner.package.clone();
    outbound
        .get_part_mut(&chain_uri)
        .expect("chain part")
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::WORKSHEET.to_owned(),
            "worksheets/sheet1.xml".to_owned(),
            "rId1".to_owned(),
            TargetMode::Internal,
        )
        .expect("chain outbound relationship");
    let outbound = crate::Package::from_opc(outbound).expect("outbound chain workbook");
    let mut edit = outbound.edit().expect("outbound chain edit");
    edit.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .set("A1", 7_i32)
        .expect("set");
    assert!(matches!(
        edit.commit(),
        Err(Error::Invalid(message)) if message.contains("calculation-chain part cannot have relationships")
    ));
    assert!(outbound.inner.package.get_part(&chain_uri).is_ok());
}

#[test]
fn shared_style_crud_is_lineage_checked_reversible_and_exact() {
    let source = styled_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let styles = source.styles().expect("styles");
    assert_eq!(styles.len(), 2);
    let base = styles.base().expect("base style");
    let accent = styles.get(1).expect("accent style");
    assert_eq!(base.fan_out().expect("base fan-out"), 1);
    assert_eq!(accent.fan_out().expect("accent fan-out"), 1);

    let sheet = source.sheet("Sheet1").expect("lookup").expect("sheet");
    assert!(matches!(
        sheet.local_style("B1").expect("local style"),
        Some(crate::LocalStyle::Default)
    ));
    let Some(crate::LocalStyle::Shared(local)) = sheet.local_style("A1").expect("local style")
    else {
        panic!("A1 must have an explicit shared style")
    };
    assert!(local.same(&accent));
    assert!(
        sheet
            .style("B1")
            .expect("resolved style")
            .is_some_and(|style| style.same(&base))
    );

    let mut edit = source.edit().expect("edit");
    let mut sheet = edit.sheet("Sheet1").expect("lookup").expect("sheet");
    sheet
        .set("C1", 42_i32)
        .and_then(|sheet| sheet.style("C1", &accent))
        .and_then(|sheet| sheet.style("D1", &accent))
        .expect("style changes");
    let committed = edit.commit().expect("commit");
    assert_eq!(committed.patch().operation_count(), 2);
    assert!(matches!(
        committed.patch().changes()[0].cell(),
        Some((_, State::Missing, _))
    ));
    assert!(matches!(
        committed.patch().changes()[0].cell(),
        Some((_, _, State::Cell {
            content: Cell::Value(Value::Number(number)),
            style: StyleState::Shared(_),
            ..
        })) if number.as_str() == "42"
    ));

    let book = committed.package();
    let sheet = book.sheet(0usize).expect("lookup").expect("sheet");
    assert!(matches!(
        sheet.cell("D1").expect("D1").stored(),
        Some(Cell::Empty)
    ));
    let styles = book.styles().expect("styles");
    let inherited = styles.find(&accent.key()).expect("inherited style key");
    assert!(inherited.same(&accent));
    assert!(!inherited.same_workbook(&accent));
    assert_eq!(accent.fan_out().expect("source fan-out"), 1);
    assert_eq!(inherited.fan_out().expect("descendant fan-out"), 3);
    assert!(
        sheet
            .style("C1")
            .expect("style")
            .is_some_and(|style| style.same(&inherited))
    );

    let mut descendant = book.edit().expect("descendant edit");
    descendant
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .style("E1", &accent)
        .expect("reuse inherited style lineage");
    let descendant = descendant.commit().expect("descendant commit");
    assert!(matches!(
        descendant
            .package()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .local_style("E1")
            .expect("local style"),
        Some(crate::LocalStyle::Shared(_))
    ));

    let reopened = crate::Package::from_bytes(&source_bytes).expect("reopened source");
    assert!(
        reopened
            .styles()
            .expect("reopened styles")
            .find(&accent.key())
            .is_none()
    );
    let replayed = reopened
        .apply(committed.patch())
        .expect("replay onto byte-identical source");
    let (
        _,
        _,
        State::Cell {
            style: StyleState::Shared(replayed_key),
            ..
        },
    ) = replayed.patch().changes()[0].cell().expect("cell change")
    else {
        panic!("replayed change must retain its shared style")
    };
    assert!(
        replayed
            .package()
            .styles()
            .expect("replayed styles")
            .find(replayed_key)
            .is_some()
    );
    assert!(
        book.styles()
            .expect("original lineage")
            .find(replayed_key)
            .is_none()
    );

    let restored = book
        .apply(&committed.patch().inverse())
        .expect("inverse patch");
    assert_eq!(
        restored.package().to_bytes().expect("restored bytes"),
        source_bytes
    );

    let foreign = crate::Package::new()
        .expect("other workbook")
        .styles()
        .expect("other styles")
        .base()
        .expect("other base style");
    let mut edit = source.edit().expect("edit");
    assert!(matches!(
        edit.sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .style("A1", &foreign),
        Err(Error::ForeignStyle)
    ));

    let mut changed_package = source.inner.package.clone();
    let styles_uri = PackURI::new("/xl/styles.xml").expect("styles URI");
    let changed_xml = {
        let styles = changed_package.get_part(&styles_uri).expect("styles part");
        std::str::from_utf8(styles.blob())
            .expect("UTF-8 styles")
            .replace("FFFFFF00", "FFFF0000")
            .into_bytes()
    };
    changed_package
        .get_part_mut(&styles_uri)
        .expect("styles part")
        .set_blob(changed_xml);
    let changed = crate::Package::from_workbook(
        Workbook::from_package_with_styles(changed_package, Some(&source))
            .expect("changed style table"),
    );
    assert!(
        changed
            .styles()
            .expect("changed styles")
            .find(&accent.key())
            .is_none()
    );
    let mut edit = changed.edit().expect("changed edit");
    assert!(matches!(
        edit.sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .style("A1", &accent),
        Err(Error::ForeignStyle)
    ));
    assert!(matches!(
        changed.apply(committed.patch()),
        Err(Error::PatchConflict { part }) if part == "/xl/styles.xml"
    ));
}

#[test]
fn payload_and_style_effects_on_one_cell_join_without_locks() {
    let source = styled_workbook();
    let accent = source.styles().expect("styles").get(1).expect("accent");
    let (mut payload, style) = std::thread::scope(|scope| {
        let payload_source = source.clone();
        let style_source = source.clone();
        let accent = accent.clone();
        let payload = scope.spawn(move || {
            let mut edit = payload_source.edit().expect("payload edit");
            edit.sheet(0usize)
                .expect("lookup")
                .expect("sheet")
                .set("B1", 9_i32)
                .expect("payload");
            edit
        });
        let style = scope.spawn(move || {
            let mut edit = style_source.edit().expect("style edit");
            edit.sheet("Sheet1")
                .expect("lookup")
                .expect("sheet")
                .style("B1", &accent)
                .expect("style");
            edit
        });
        (
            payload.join().expect("payload worker"),
            style.join().expect("style worker"),
        )
    });
    payload.join(style).expect("disjoint cell facets");
    assert_eq!(payload.len(), 1);
    let committed = payload.commit().expect("commit");
    assert_eq!(committed.patch().operation_count(), 1);
    assert!(matches!(
        committed.patch().changes()[0].cell(),
        Some((_, _, State::Cell {
            content: Cell::Value(Value::Number(number)),
            style: StyleState::Shared(_),
            ..
        })) if number.as_str() == "9"
    ));

    let sheet = committed
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet");
    assert!(matches!(
        sheet.cell("B1").expect("B1").stored(),
        Some(Cell::Value(Value::Number(number))) if number.as_str() == "9"
    ));
    assert!(matches!(
        sheet.local_style("B1").expect("style"),
        Some(crate::LocalStyle::Shared(_))
    ));
}

#[test]
fn rich_shared_string_transfer_keeps_exact_dependency_through_join_replay_and_inverse() {
    let baseline = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut package = baseline.inner.package.clone();
    let shared_uri = PackURI::new("/xl/sharedStrings.xml").expect("shared strings URI");
    package
        .try_add_part(Box::new(BlobPart::new(
            shared_uri.clone(),
            litchi_opc::constants::content_type::SML_SHARED_STRINGS.to_owned(),
            br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><r><rPr><b/><color rgb="FFFF0000"/></rPr><t>Rich</t></r><r><rPr><i/></rPr><t> text</t></r></si></sst>"#.to_vec(),
        )))
        .expect("shared strings part");
    package
        .get_part_mut(&baseline.inner.workbook_uri)
        .expect("workbook part")
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::SHARED_STRINGS.to_owned(),
            "sharedStrings.xml".to_owned(),
            "rIdSharedStrings".to_owned(),
            TargetMode::Internal,
        )
        .expect("shared strings relationship");
    package
        .get_part_mut(&baseline.inner.sheets[0].part_uri)
        .expect("source worksheet")
        .set_blob(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>"#.to_vec(),
        );
    let source = crate::Package::from_opc(package).expect("rich-string workbook");
    let source_bytes = source.to_bytes().expect("source bytes");
    let shared_before = part_text(&source, "/xl/sharedStrings.xml").to_owned();

    let mut transfer = source.edit().expect("transfer edit");
    transfer
        .copy_cells("Sheet1", "A1", "Sheet2", "B2")
        .expect("copy rich string");
    let mut disjoint = source.edit().expect("disjoint edit");
    disjoint
        .sheet("Sheet2")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("C3", 7_i32)
        .expect("disjoint value");
    transfer.join(disjoint).expect("join disjoint edit");
    let committed = transfer.commit().expect("commit rich transfer");

    let target_xml = part_text(committed.package(), "/xl/worksheets/sheet2.xml");
    assert!(target_xml.contains(r#"<c r="B2" t="s"><v>0</v></c>"#));
    assert_eq!(
        part_text(committed.package(), "/xl/sharedStrings.xml"),
        shared_before
    );
    assert!(committed.patch().changes().iter().any(|change| matches!(
        change.cell(),
        Some((
            _,
            _,
            State::Cell {
                shared_string: Some(_),
                ..
            }
        ))
    )));

    let replayed = source
        .apply(committed.patch())
        .expect("replay rich transfer");
    assert!(
        part_text(replayed.package(), "/xl/worksheets/sheet2.xml")
            .contains(r#"<c r="B2" t="s"><v>0</v></c>"#)
    );
    let restored = committed
        .package()
        .apply(&committed.patch().inverse())
        .expect("inverse rich transfer");
    assert_eq!(
        restored.package().to_bytes().expect("restored bytes"),
        source_bytes
    );

    let shared_rich = source
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .rich_text("A1")
        .expect("shared rich read")
        .expect("shared rich value");
    let mut rewrite = source.edit().expect("shared-rich rewrite edit");
    rewrite
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set_rich_text("A1", shared_rich.clone())
        .expect("rewrite shared rich inline");
    let rewritten = rewrite.commit().expect("shared-rich rewrite commit");
    assert!(!rewritten.patch().is_noop());
    assert!(
        part_text(rewritten.package(), "/xl/worksheets/sheet1.xml").contains(r#"t="inlineStr""#)
    );
    assert_eq!(
        rewritten
            .package()
            .sheet("Sheet1")
            .expect("sheet lookup")
            .expect("worksheet")
            .rich_text("A1")
            .expect("rewritten rich read"),
        Some(shared_rich)
    );

    let mut plain = source
        .edit_with_options(EditOptions::default().with_max_retained_bytes(usize::MAX))
        .expect("bounded shared-rich plain edit");
    plain
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", "Rich text")
        .expect("erase shared runs");
    let plain = plain.commit().expect("bounded shared-rich plain commit");
    assert!(!plain.patch().is_noop());
    assert!(
        plain
            .package()
            .sheet("Sheet1")
            .expect("sheet lookup")
            .expect("worksheet")
            .rich_text("A1")
            .expect("plain rich read")
            .is_none()
    );

    // The inverse direction is also observable: adding a run property to
    // ordinary text must not be collapsed merely because the visible text is
    // unchanged.
    let plain = crate::Package::new().expect("plain workbook");
    let mut rich = crate::RichText::plain("same");
    rich.run_mut(0)
        .expect("plain run")
        .apply_font(&FontChange::new().italic(Some(true)))
        .expect("plain run formatting");
    let mut edit = plain.edit().expect("plain edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set_rich_text("A1", rich)
        .expect("plain-to-rich edit");
    let plain_to_rich = edit.commit().expect("plain-to-rich commit");
    assert!(!plain_to_rich.patch().is_noop());
}

#[test]
fn same_text_rich_run_property_change_is_not_collapsed_as_a_noop() {
    let baseline = crate::Package::new().expect("baseline workbook");
    let mut package = baseline.inner.package.clone();
    package
        .get_part_mut(&baseline.inner.sheets[0].part_uri)
        .expect("worksheet part")
        .set_blob(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><rPr><b/></rPr><t>same</t></r></is></c></row></sheetData></worksheet>"#.to_vec(),
        );
    let source = crate::Package::from_opc(package).expect("rich inline workbook");
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut rich = source
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .rich_text("A1")
        .expect("rich read")
        .expect("rich value");
    rich.run_mut(0)
        .expect("first run")
        .apply_font(&FontChange::new().italic(Some(true)))
        .expect("run formatting");

    let mut edit = source.edit().expect("edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set_rich_text("A1", rich)
        .expect("same-text rich edit");
    let committed = edit.commit().expect("commit rich edit");
    assert!(!committed.patch().is_noop());
    let target = committed
        .package()
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .rich_text("A1")
        .expect("rich target read")
        .expect("rich target");
    assert_eq!(target.text(), "same");
    assert!(target.runs()[0].font().expect("target font").italic);

    let replayed = source.apply(committed.patch()).expect("replay rich edit");
    assert_eq!(
        replayed.package().to_bytes().expect("replayed bytes"),
        committed.package().to_bytes().expect("target bytes")
    );
    let restored = committed
        .package()
        .apply(&committed.patch().inverse())
        .expect("inverse rich edit");
    assert_eq!(
        restored.package().to_bytes().expect("restored bytes"),
        source_bytes
    );
}

#[test]
fn plain_set_over_same_visible_rich_text_erases_runs() {
    let baseline = crate::Package::new().expect("baseline workbook");
    let mut package = baseline.inner.package.clone();
    package
        .get_part_mut(&baseline.inner.sheets[0].part_uri)
        .expect("worksheet part")
        .set_blob(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><rPr><b/></rPr><t>same</t></r></is></c></row></sheetData></worksheet>"#.to_vec(),
        );
    let source = crate::Package::from_opc(package).expect("rich inline workbook");
    let mut edit = source
        .edit_with_options(EditOptions::default().with_max_retained_bytes(usize::MAX))
        .expect("bounded edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", "same")
        .expect("plain replacement");
    let committed = edit.commit().expect("commit plain replacement");
    assert!(!committed.patch().is_noop());
    assert!(
        committed
            .package()
            .sheet("Sheet1")
            .expect("sheet lookup")
            .expect("worksheet")
            .rich_text("A1")
            .expect("rich read")
            .is_none()
    );
}

#[test]
fn unbounded_plain_set_preserves_shared_string_identity() {
    let baseline = crate::Package::new().expect("baseline workbook");
    let mut install = baseline.edit().expect("install edit");
    install
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", "plain shared text")
        .expect("plain shared value");
    let source = install.commit().expect("install commit").into_package();
    let source_bytes = source.to_bytes().expect("source bytes");

    let mut edit = source.edit().expect("unbounded edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", "plain shared text")
        .expect("same plain value");
    let committed = edit.commit().expect("source-equivalent commit");
    assert!(committed.patch().is_noop());
    assert_eq!(
        committed.package().to_bytes().expect("target bytes"),
        source_bytes
    );

    let mut bounded = source
        .edit_with_options(EditOptions::default().with_max_retained_bytes(usize::MAX))
        .expect("bounded edit");
    bounded
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", "plain shared text")
        .expect("same bounded plain value");
    let bounded = bounded.commit().expect("bounded no-op commit");
    assert!(bounded.patch().is_noop());
    assert_eq!(
        bounded.package().to_bytes().expect("bounded target bytes"),
        source_bytes
    );
}

#[test]
fn remove_then_differing_typed_style_recreates_the_cell() {
    let source = styled_workbook();
    let mut edit = source.edit().expect("edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .remove("A1")
        .and_then(|sheet| sheet.font("A1", &FontChange::new().bold(Some(true))))
        .expect("remove then typed style");
    let committed = edit.commit().expect("commit");
    let sheet = committed
        .package()
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet");
    assert!(sheet.cell("A1").expect("recreated cell").stored().is_some());
    assert_eq!(
        sheet
            .style("A1")
            .expect("style lookup")
            .expect("recreated style")
            .format()
            .expect("format")
            .font()
            .expect("font")
            .bold(),
        Some(true)
    );
}

#[test]
fn rich_transfer_keeps_the_source_shared_style() {
    let source = styled_workbook();
    let mut package = source.inner.package.clone();
    package
        .get_part_mut(&source.inner.sheets[0].part_uri)
        .expect("worksheet part")
        .set_blob(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1" t="inlineStr"><is><r><rPr><b/></rPr><t>styled rich</t></r></is></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>"#.to_vec(),
        );
    let source = crate::Package::from_opc(package).expect("styled rich workbook");
    let mut edit = source.edit().expect("edit");
    edit.copy_cells("Sheet1", "A1", "Sheet1", "B1")
        .expect("rich transfer")
        .expect("selectors");
    let committed = edit.commit().expect("commit rich transfer");
    let target = committed
        .package()
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet");
    assert!(matches!(
        target.local_style("B1").expect("target style"),
        Some(crate::LocalStyle::Shared(_))
    ));
    assert_eq!(
        target
            .style("B1")
            .expect("target style")
            .expect("shared style")
            .format()
            .expect("target format")
            .number_format()
            .code(),
        source
            .sheet("Sheet1")
            .expect("source lookup")
            .expect("source worksheet")
            .style("A1")
            .expect("source style")
            .expect("source shared style")
            .format()
            .expect("source format")
            .number_format()
            .code()
    );
    assert_eq!(
        target
            .rich_text("B1")
            .expect("rich read")
            .expect("rich value")
            .text(),
        "styled rich"
    );
}

#[test]
fn inline_rich_range_transfer_preserves_runs_and_noop_identity() {
    let baseline = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut package = baseline.inner.package.clone();
    package
        .get_part_mut(&baseline.inner.sheets[0].part_uri)
        .expect("source worksheet")
        .set_blob(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><rPr><b/></rPr><t>one</t></r></is></c><c r="B1"><v>1</v></c></row><row r="2"><c r="A2" t="inlineStr"><is><r><rPr><i/></rPr><t>two</t></r></is></c><c r="B2"><v>2</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><r><rPr><u val="single"/></rPr><t>three</t></r></is></c><c r="B3"><v>3</v></c></row><row r="4"><c r="A4" t="inlineStr"><is><r><rPr><color rgb="FFFF0000"/></rPr><t>four</t></r></is></c><c r="B4"><v>4</v></c></row></sheetData></worksheet>"#.to_vec(),
        );
    let source = crate::Package::from_opc(package).expect("rich range workbook");
    let expected = ["A1", "A2", "A3", "A4"].map(|address| {
        source
            .workbook()
            .sheet("Sheet1")
            .expect("sheet lookup")
            .expect("worksheet")
            .rich_text(address)
            .expect("source rich text")
            .expect("source runs")
    });

    let mut edit = source.edit().expect("edit");
    edit.copy_cells("Sheet1", "A1:A4", "Sheet1", "B1")
        .expect("bulk rich transfer")
        .expect("selectors");
    let committed = edit.commit().expect("commit bulk rich transfer");
    let target = committed
        .package()
        .workbook()
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet");
    for (index, address) in ["B1", "B2", "B3", "B4"].into_iter().enumerate() {
        assert_eq!(
            target.rich_text(address).expect("target rich text"),
            Some(expected[index].clone())
        );
    }

    let mut noop = committed.package().edit().expect("noop edit");
    noop.copy_cells("Sheet1", "A1:A4", "Sheet1", "B1")
        .expect("repeat bulk rich transfer")
        .expect("selectors");
    let noop = noop.commit().expect("noop commit");
    assert!(noop.patch().is_noop());
}

#[test]
fn moved_rich_source_does_not_reappear_in_later_copy() {
    let baseline = crate::Package::new().expect("baseline workbook");
    let mut package = baseline.inner.package.clone();
    package
        .get_part_mut(&baseline.inner.sheets[0].part_uri)
        .expect("source worksheet")
        .set_blob(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><rPr><b/></rPr><t>moved</t></r></is></c></row></sheetData></worksheet>"#.to_vec(),
        );
    let source = crate::Package::from_opc(package).expect("rich source workbook");
    let source_bytes = source.to_bytes().expect("source bytes");

    let mut edit = source.edit().expect("edit");
    edit.move_cells("Sheet1", "A1", "Sheet1", "B1")
        .expect("move rich source")
        .expect("move selectors");
    edit.copy_cells("Sheet1", "A1", "Sheet1", "C1")
        .expect("copy removed source")
        .expect("copy selectors");
    let committed = edit.commit().expect("commit move and copy");
    let sheet = committed
        .package()
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet");
    assert!(
        sheet
            .rich_text("A1")
            .expect("removed source rich text")
            .is_none()
    );
    assert_eq!(
        sheet
            .rich_text("B1")
            .expect("moved rich text")
            .expect("moved runs")
            .text(),
        "moved"
    );
    assert!(
        sheet
            .rich_text("C1")
            .expect("copied removed source rich text")
            .is_none()
    );

    let replayed = source
        .apply(committed.patch())
        .expect("replay move and copy");
    assert_eq!(
        replayed.package().to_bytes().expect("replayed bytes"),
        committed.package().to_bytes().expect("committed bytes")
    );
    let restored = committed
        .package()
        .apply(&committed.patch().inverse())
        .expect("inverse move and copy");
    assert_eq!(
        restored.package().to_bytes().expect("restored bytes"),
        source_bytes
    );
}

#[test]
fn typed_facet_equal_to_source_overrides_pending_complete_style() {
    let source = crate::Package::new().expect("source workbook");
    let mut configure = source.edit().expect("configure edit");
    configure
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", "styled")
        .and_then(|sheet| {
            sheet.font(
                "A1",
                &FontChange::new()
                    .bold(Some(true))
                    .name(Some("Accent"))
                    .expect("font name"),
            )
        })
        .expect("complete source style");
    let configured = configure.commit().expect("configure commit").into_package();
    let accent = configured
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .style("A1")
        .expect("style lookup")
        .expect("accent style");

    let mut edit = configured
        .edit_with_options(EditOptions::default().with_max_retained_bytes(usize::MAX))
        .expect("bounded edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .style("B1", &accent)
        .and_then(|sheet| {
            sheet.font(
                "B1",
                &FontChange::new().name(None).expect("clear font name"),
            )
        })
        .expect("typed override");
    let committed = edit.commit().expect("override commit");
    let format = committed
        .package()
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .style("B1")
        .expect("style lookup")
        .expect("styled B1")
        .format()
        .expect("cell format");
    let font = format.font().expect("font");
    assert_eq!(font.bold(), Some(true));
    assert_eq!(font.name(), None);
}

#[test]
fn typed_facet_equal_to_pending_complete_style_releases_capture() {
    let source = crate::Package::new().expect("source workbook");
    let mut configure = source.edit().expect("configure edit");
    configure
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", "styled")
        .and_then(|sheet| {
            sheet.font(
                "A1",
                &FontChange::new()
                    .bold(Some(true))
                    .name(Some("Accent"))
                    .expect("font name"),
            )
        })
        .expect("complete source style");
    let configured = configure.commit().expect("configure commit").into_package();
    let accent = configured
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .style("A1")
        .expect("style lookup")
        .expect("accent style");
    let sheet = &configured.inner.sheets[0];
    let worksheet_bytes = configured
        .inner
        .package
        .get_part(&sheet.part_uri)
        .expect("worksheet part")
        .blob()
        .len();
    let maximum = configured
        .retained_bytes()
        .saturating_mul(2)
        .saturating_add(worksheet_bytes)
        .saturating_add(size_of::<Change>())
        .saturating_add(size_of::<PartChange>())
        .saturating_add(size_of::<StyleEffect>());
    let mut edit = configured
        .edit_with_options(EditOptions::default().with_max_retained_bytes(maximum))
        .expect("bounded edit");
    {
        let mut sheet = edit
            .sheet("Sheet1")
            .expect("sheet lookup")
            .expect("worksheet");
        sheet.style("B1", &accent).expect("pending complete style");
        sheet
            .font("B1", &FontChange::new().bold(Some(true)))
            .expect("typed request equals pending style");
    }
    assert!(
        edit.sheets
            .get(&0)
            .is_none_or(|actions| actions.fonts.is_empty())
    );
}

#[test]
fn source_style_and_reset_release_superseded_typed_capture() {
    let source = styled_workbook();
    let source_style = source
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .style("A1")
        .expect("style lookup")
        .expect("source style");
    let worksheet_bytes = source
        .inner
        .package
        .get_part(&source.inner.sheets[0].part_uri)
        .expect("worksheet part")
        .blob()
        .len();
    let font = FontChange::new().bold(Some(true));
    let payload = Content::from("after style supersession");
    let maximum = source
        .retained_bytes()
        .saturating_mul(2)
        .saturating_add(worksheet_bytes)
        .saturating_add(size_of::<Change>().saturating_mul(2))
        .saturating_add(size_of::<PartChange>())
        .saturating_add(font.retained_bytes())
        .saturating_add(payload.retained_bytes());
    let mut edit = source
        .edit_with_options(EditOptions::default().with_max_retained_bytes(maximum))
        .expect("bounded edit");
    {
        let mut sheet = edit
            .sheet("Sheet1")
            .expect("sheet lookup")
            .expect("worksheet");
        sheet.font("A1", &font).expect("typed style");
        sheet
            .style("A1", &source_style)
            .expect("source style supersession");
        sheet
            .set("B1", payload)
            .expect("released typed capture leaves room for payload");
    }
    assert!(
        edit.sheets
            .get(&0)
            .is_none_or(|actions| actions.fonts.is_empty())
    );

    let source = crate::Package::new().expect("source workbook");
    let worksheet_bytes = source
        .inner
        .package
        .get_part(&source.inner.sheets[0].part_uri)
        .expect("worksheet part")
        .blob()
        .len();
    let payload = Content::from("after reset supersession");
    let maximum = source
        .retained_bytes()
        .saturating_mul(2)
        .saturating_add(worksheet_bytes)
        .saturating_add(size_of::<Change>().saturating_mul(2))
        .saturating_add(size_of::<PartChange>())
        .saturating_add(font.retained_bytes())
        .saturating_add(payload.retained_bytes());
    let mut edit = source
        .edit_with_options(EditOptions::default().with_max_retained_bytes(maximum))
        .expect("bounded edit");
    {
        let mut sheet = edit
            .sheet("Sheet1")
            .expect("sheet lookup")
            .expect("worksheet");
        sheet.font("B1", &font).expect("typed style");
        sheet.reset_style("B1").expect("missing reset supersession");
        sheet
            .set("C1", payload)
            .expect("released reset capture leaves room for payload");
    }
    assert!(
        edit.sheets
            .get(&0)
            .is_none_or(|actions| actions.fonts.is_empty())
    );
}

#[test]
fn complete_style_source_equivalent_releases_capture_before_next_payload() {
    let source = styled_workbook();
    let styles = source.styles().expect("styles");
    let base = styles.base().expect("base style");
    let source_style = styles.get(1).expect("source style");
    let worksheet_bytes = source
        .inner
        .package
        .get_part(&source.inner.sheets[0].part_uri)
        .expect("worksheet part")
        .blob()
        .len();
    let payload = Content::from("payload after complete-style cancellation");
    let maximum = source
        .retained_bytes()
        .saturating_mul(2)
        .saturating_add(worksheet_bytes)
        .saturating_add(size_of::<Change>())
        .saturating_add(size_of::<PartChange>())
        .saturating_add(payload.retained_bytes());

    let mut style_edit = source
        .edit_with_options(EditOptions::default().with_max_retained_bytes(maximum))
        .expect("bounded style edit");
    style_edit
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .style("A1", &base)
        .and_then(|sheet| sheet.style("A1", &source_style))
        .and_then(|sheet| sheet.set("C1", payload.clone()))
        .expect("source-equivalent style releases stale capture");

    let mut reset_edit = source
        .edit_with_options(EditOptions::default().with_max_retained_bytes(maximum))
        .expect("bounded reset edit");
    reset_edit
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .style("B1", &source_style)
        .and_then(|sheet| sheet.reset_style("B1"))
        .and_then(|sheet| sheet.set("C1", payload))
        .expect("source-equivalent reset releases stale capture");
}

#[test]
fn unbounded_source_equivalent_complete_style_and_reset_are_noops() {
    let source = styled_workbook();
    let source_style = source
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .style("A1")
        .expect("style lookup")
        .expect("source style");
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut style_edit = source.edit().expect("style edit");
    style_edit
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .style("A1", &source_style)
        .expect("source-equivalent style");
    let styled = style_edit.commit().expect("style no-op commit");
    assert!(styled.patch().is_noop());
    assert_eq!(
        styled.package().to_bytes().expect("style bytes"),
        source_bytes
    );

    let mut reset_edit = source.edit().expect("reset edit");
    reset_edit
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .reset_style("B1")
        .expect("missing reset");
    let reset = reset_edit.commit().expect("reset no-op commit");
    assert!(reset.patch().is_noop());
    assert_eq!(
        reset.package().to_bytes().expect("reset bytes"),
        source_bytes
    );
}

#[test]
fn hyperlink_ranges_with_one_start_keep_independent_captures() {
    let source = crate::Package::new().expect("source workbook");
    let mut edit = source
        .edit_with_options(EditOptions::default().with_max_retained_bytes(usize::MAX))
        .expect("bounded edit");
    let first = crate::hyperlinks::Hyperlink::external(
        crate::hyperlinks::HyperlinkReference::new("A1:B1").expect("first reference"),
        "https://example.test/first",
    )
    .expect("first hyperlink");
    let second = crate::hyperlinks::Hyperlink::external(
        crate::hyperlinks::HyperlinkReference::new("A1:C1").expect("second reference"),
        "https://example.test/second",
    )
    .expect("second hyperlink");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .put_hyperlink(first)
        .and_then(|sheet| sheet.put_hyperlink(second))
        .expect("independent hyperlinks");
    let committed = edit.commit().expect("commit hyperlinks");
    assert_eq!(
        committed
            .package()
            .sheet("Sheet1")
            .expect("sheet lookup")
            .expect("worksheet")
            .hyperlinks()
            .expect("hyperlinks")
            .len(),
        2
    );
}

#[test]
fn cell_transfer_clones_selected_picture_graph_and_is_reversible() {
    let baseline = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut package = baseline.inner.package.clone();
    package
        .get_part_mut(&baseline.inner.sheets[0].part_uri)
        .expect("source worksheet")
        .set_blob(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1"/><sheetData/><drawing r:id="rIdDrawing"/></worksheet>"#.to_vec(),
        );
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/drawings/drawing1.xml").expect("drawing URI"),
            litchi_opc::constants::content_type::OFC_DRAWING.to_owned(),
            br#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Picture 1"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#.to_vec(),
        )))
        .expect("drawing part");
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/media/image1.png").expect("image URI"),
            litchi_opc::constants::content_type::PNG.to_owned(),
            b"picture-bytes".to_vec(),
        )))
        .expect("image part");
    package
        .get_part_mut(&PackURI::new("/xl/drawings/drawing1.xml").expect("drawing URI"))
        .expect("drawing part")
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::IMAGE.to_owned(),
            "../media/image1.png".to_owned(),
            "rIdImage".to_owned(),
            TargetMode::Internal,
        )
        .expect("image relationship");
    package
        .get_part_mut(&baseline.inner.sheets[0].part_uri)
        .expect("source worksheet")
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::DRAWING.to_owned(),
            "../drawings/drawing1.xml".to_owned(),
            "rIdDrawing".to_owned(),
            TargetMode::Internal,
        )
        .expect("drawing relationship");
    let source = crate::Package::from_opc(package).expect("drawing workbook");
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    edit.copy_cells("Sheet1", "A1:C3", "Sheet2", "E5")
        .expect("picture transfer")
        .expect("selectors");
    let committed = edit.commit().expect("commit picture transfer");
    assert!(
        committed
            .patch()
            .package_changes()
            .iter()
            .any(|change| matches!(
                change,
                PackageChange::DrawingTransfer {
                    source,
                    target,
                    anchors: 1,
                    added: true,
                } if source.as_ref() == "Sheet1" && target.as_ref() == "Sheet2"
            ))
    );
    let target_sheet = part_text(committed.package(), "/xl/worksheets/sheet2.xml");
    assert!(target_sheet.contains(r#"r:id="rIdDrawingCopy1""#));
    let target_drawing = part_text(committed.package(), "/xl/drawings/drawing1_copy1.xml");
    assert!(target_drawing.contains("<xdr:col>4</xdr:col>"));
    assert!(target_drawing.contains("<xdr:row>4</xdr:row>"));
    assert!(target_drawing.contains("<xdr:col>6</xdr:col>"));
    assert!(target_drawing.contains("<xdr:row>6</xdr:row>"));
    assert!(target_drawing.contains(r#"r:embed="rIdImage""#));
    assert_eq!(
        committed
            .package()
            .inner
            .package
            .get_part(&PackURI::new("/xl/media/image1_copy1.png").expect("cloned image URI"))
            .expect("cloned image")
            .blob(),
        b"picture-bytes"
    );

    let replayed = source
        .apply(committed.patch())
        .expect("replay picture graph");
    assert_eq!(
        replayed.package().to_bytes().expect("replayed bytes"),
        committed.package().to_bytes().expect("committed bytes")
    );
    let restored = committed
        .package()
        .apply(&committed.patch().inverse())
        .expect("inverse picture graph");
    assert_eq!(
        restored.package().to_bytes().expect("restored"),
        source_bytes
    );
}

#[test]
fn cell_transfer_clones_classic_chart_graph_through_merge_and_durable_inverse() {
    let baseline = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut package = baseline.inner.package.clone();
    package
        .get_part_mut(&baseline.inner.sheets[0].part_uri)
        .expect("source worksheet")
        .set_blob(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1"/><sheetData/><drawing r:id="rIdDrawing"/></worksheet>"#.to_vec(),
        );
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/drawings/drawing1.xml").expect("drawing URI"),
            litchi_opc::constants::content_type::OFC_DRAWING.to_owned(),
            br#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#.to_vec(),
        )))
        .expect("drawing part");
    let chart_xml = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:plotArea/></c:chart><c:externalData r:id="rIdWorkbook"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>"#.to_vec();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/charts/chart1.xml").expect("chart URI"),
            litchi_opc::constants::content_type::DML_CHART.to_owned(),
            chart_xml.clone(),
        )))
        .expect("chart part");
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/embeddings/embedded1.xlsx").expect("embedded workbook URI"),
            litchi_opc::constants::content_type::OFC_PACKAGE.to_owned(),
            b"embedded-workbook-bytes".to_vec(),
        )))
        .expect("embedded workbook part");
    package
        .get_part_mut(&PackURI::new("/xl/charts/chart1.xml").expect("chart URI"))
        .expect("chart part")
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::PACKAGE.to_owned(),
            "../embeddings/embedded1.xlsx".to_owned(),
            "rIdWorkbook".to_owned(),
            TargetMode::Internal,
        )
        .expect("embedded workbook relationship");
    package
        .get_part_mut(&PackURI::new("/xl/drawings/drawing1.xml").expect("drawing URI"))
        .expect("drawing part")
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::CHART.to_owned(),
            "../charts/chart1.xml".to_owned(),
            "rIdChart".to_owned(),
            TargetMode::Internal,
        )
        .expect("chart relationship");
    package
        .get_part_mut(&baseline.inner.sheets[0].part_uri)
        .expect("source worksheet")
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::DRAWING.to_owned(),
            "../drawings/drawing1.xml".to_owned(),
            "rIdDrawing".to_owned(),
            TargetMode::Internal,
        )
        .expect("drawing relationship");
    let source = crate::Package::from_opc(package).expect("chart workbook");
    let source_bytes = source.to_bytes().expect("source bytes");

    let mut transfer = source.edit().expect("transfer edit");
    transfer
        .copy_cells("Sheet1", "A1:C3", "Sheet2", "E5")
        .expect("chart transfer")
        .expect("selectors");
    let mut disjoint = source.edit().expect("disjoint edit");
    disjoint
        .sheet("Sheet2")
        .expect("lookup")
        .expect("target sheet")
        .set("H8", "joined")
        .expect("disjoint cell");
    let merged = transfer
        .plan_three_way(disjoint, MergeLimits::new(2, 64, 128, 64))
        .expect("plan chart transfer merge");
    assert!(merged.conflicts().is_empty());
    let committed = merged
        .finish()
        .expect("finish chart transfer merge")
        .commit()
        .expect("commit chart transfer");

    let target_drawing = part_text(committed.package(), "/xl/drawings/drawing1_copy1.xml");
    assert!(target_drawing.contains(r#"r:id="rIdChart""#));
    let target_drawing_uri =
        PackURI::new("/xl/drawings/drawing1_copy1.xml").expect("cloned drawing URI");
    let target_chart_uri = PackURI::new("/xl/charts/chart1_copy1.xml").expect("cloned chart URI");
    let target_chart_relationship = committed
        .package()
        .inner
        .package
        .get_part(&target_drawing_uri)
        .expect("cloned drawing")
        .rels()
        .get("rIdChart")
        .expect("cloned chart relationship");
    assert_eq!(
        target_chart_relationship
            .target_partname()
            .expect("chart target"),
        target_chart_uri
    );
    assert_eq!(
        committed
            .package()
            .inner
            .package
            .get_part(&target_chart_uri)
            .expect("cloned chart")
            .blob(),
        chart_xml.as_slice()
    );
    let target_chart = committed
        .package()
        .inner
        .package
        .get_part(&target_chart_uri)
        .expect("cloned chart");
    let target_embedding_uri =
        PackURI::new("/xl/embeddings/embedded1_copy1.xlsx").expect("cloned embedding URI");
    assert_eq!(
        target_chart
            .rels()
            .get("rIdWorkbook")
            .expect("cloned embedded workbook relationship")
            .target_partname()
            .expect("embedded workbook target"),
        target_embedding_uri
    );
    assert_eq!(
        committed
            .package()
            .inner
            .package
            .get_part(&target_embedding_uri)
            .expect("cloned embedded workbook")
            .blob(),
        b"embedded-workbook-bytes"
    );
    assert!(matches!(
        committed
            .package()
            .sheet("Sheet2")
            .expect("lookup")
            .expect("target")
            .cell("H8")
            .expect("joined cell")
            .stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "joined"
    ));

    let durable = committed.patch().durable().expect("durable chart patch");
    let json = durable
        .to_deterministic_json()
        .expect("deterministic chart patch");
    let durable = DurablePatch::from_deterministic_json(&json).expect("parse chart patch");
    let replayed = durable.apply(&source).expect("replay chart graph");
    assert_eq!(
        replayed.to_bytes().expect("replayed bytes"),
        committed.package().to_bytes().expect("committed bytes")
    );
    assert_eq!(
        durable
            .inverse()
            .apply(&replayed)
            .expect("inverse chart graph")
            .to_bytes()
            .expect("restored bytes"),
        source_bytes
    );
}

#[test]
fn resetting_style_is_distinct_from_removal_and_missing_is_a_no_op() {
    let source = styled_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    let mut sheet = edit.sheet(0usize).expect("lookup").expect("sheet");
    sheet
        .reset_style("A1")
        .and_then(|sheet| sheet.reset_style("Z99"))
        .expect("style resets");
    let committed = edit.commit().expect("commit");
    assert_eq!(committed.patch().operation_count(), 1);
    assert!(matches!(
        committed.patch().changes()[0].cell(),
        Some((
            _,
            _,
            State::Cell {
                style: StyleState::Default,
                ..
            }
        ))
    ));
    let sheet = committed
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet");
    assert!(matches!(
        sheet.local_style("A1").expect("local style"),
        Some(crate::LocalStyle::Default)
    ));
    assert!(sheet.cell("Z99").expect("missing").is_missing());

    let restored = committed
        .package()
        .apply(&committed.patch().inverse())
        .expect("inverse");
    assert_eq!(restored.package().to_bytes().expect("bytes"), source_bytes);
}

#[test]
fn signed_packages_refuse_changed_edits_and_keep_exact_noops() {
    let baseline = crate::Package::new().expect("baseline");
    let mut package = baseline.inner.package.clone();
    package
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            "_xmlsignatures/origin.sigs".to_owned(),
            "rIdSignature".to_owned(),
            TargetMode::Internal,
        )
        .expect("signature relationship");
    let signed = crate::Package::from_opc(package).expect("signed workbook snapshot");

    let source_bytes = signed.to_bytes().expect("signed bytes");
    // Exact no-ops return the signed source (ADR 0006).
    let noop = signed.edit().expect("edit").commit().expect("no-op");
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().to_bytes().expect("bytes"), source_bytes);
    let replayed = signed.apply(&Patch::default()).expect("empty patch");
    assert_eq!(replayed.package().to_bytes().expect("bytes"), source_bytes);
    // A change is refused before anything is published.
    let mut edit = signed.edit().expect("edit");
    edit.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .set("A1", 1)
        .expect("stage");
    assert!(matches!(edit.commit(), Err(Error::Signed)));
    let mut unsigned_edit = baseline.edit().expect("edit");
    unsigned_edit
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .set("A1", 1)
        .expect("stage");
    let patch = unsigned_edit.commit().expect("commit").patch().clone();
    assert!(matches!(signed.apply(&patch), Err(Error::Signed)));
}

#[test]
fn signed_unbounded_source_equivalent_set_is_an_exact_noop() {
    let baseline = crate::Package::new().expect("baseline");
    let mut install = baseline.edit().expect("install edit");
    install
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", "signed plain value")
        .expect("install value");
    let unsigned = install.commit().expect("install commit").into_package();
    let mut package = unsigned.inner.package.clone();
    package
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            "_xmlsignatures/origin.sigs".to_owned(),
            "rIdSignature".to_owned(),
            TargetMode::Internal,
        )
        .expect("signature relationship");
    let signed = crate::Package::from_opc(package).expect("signed package");
    let source_bytes = signed.to_bytes().expect("source bytes");

    let mut edit = signed.edit().expect("unbounded signed edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", "signed plain value")
        .expect("source-equivalent set");
    let committed = edit.commit().expect("signed exact no-op");
    assert!(committed.patch().is_noop());
    assert_eq!(
        committed.package().to_bytes().expect("target bytes"),
        source_bytes
    );
}

#[test]
fn font_change_strike_and_underline_round_trip_by_coordinate() {
    let source = crate::Package::new().expect("source workbook");
    let source_bytes = source.to_bytes().expect("source bytes");
    let address = Address::from_a1("A1").expect("address");

    let mut edit = source.edit().expect("edit");
    let mut sheet = edit.sheet(0usize).expect("lookup").expect("worksheet");
    sheet.set(address, "styled").expect("cell value");
    sheet
        .font(address, &FontChange::new().strike(Some(true)))
        .expect("strike");
    sheet
        .font(
            address,
            &FontChange::new().underline(Some(Underline::Double)),
        )
        .expect("underline");
    let committed = edit.commit().expect("font commit");
    assert_eq!(committed.patch().operation_count(), 1);

    let committed_bytes = committed.package().to_bytes().expect("committed bytes");
    let styles = part_text(committed.package(), "/xl/styles.xml");
    assert!(styles.contains("<strike/>"));
    assert!(styles.contains(r#"<u val="double"/>"#));
    let sheet = committed
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet");
    let style = sheet
        .style(address)
        .expect("style lookup")
        .expect("styled cell");
    assert_eq!(
        style
            .format()
            .expect("cell format")
            .font()
            .and_then(|font| font.underline()),
        Some(Underline::Double)
    );

    let reopened = crate::Package::from_bytes(&committed_bytes).expect("reopened workbook");
    assert_eq!(part_text(&reopened, "/xl/styles.xml"), styles);
    assert_eq!(
        reopened
            .sheet(0usize)
            .expect("lookup")
            .expect("worksheet")
            .style(address)
            .expect("style lookup")
            .expect("styled cell")
            .format()
            .expect("cell format")
            .font()
            .and_then(|font| font.underline()),
        Some(Underline::Double)
    );

    let replayed = source.apply(committed.patch()).expect("replay patch");
    assert_eq!(
        replayed.package().to_bytes().expect("replayed bytes"),
        committed_bytes
    );
    let restored = committed
        .package()
        .apply(&committed.patch().inverse())
        .expect("inverse patch");
    assert_eq!(
        restored.package().to_bytes().expect("restored bytes"),
        source_bytes
    );
}

#[test]
fn font_change_same_property_conflicts_but_disjoint_properties_join() {
    let source = crate::Package::new().expect("source workbook");
    let address = Address::from_a1("A1").expect("address");

    let mut strike = source.edit().expect("strike edit");
    strike
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(address, &FontChange::new().strike(Some(true)))
        .expect("strike");
    let mut underline = source.edit().expect("underline edit");
    underline
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(
            address,
            &FontChange::new().underline(Some(Underline::Single)),
        )
        .expect("underline");
    strike.join(underline).expect("disjoint font properties");
    let joined = strike.commit().expect("joined font commit");
    let joined_styles = part_text(joined.package(), "/xl/styles.xml");
    assert!(joined_styles.contains("<strike/>"));
    assert!(joined_styles.contains(r#"<u/>"#));

    let mut left = source.edit().expect("left edit");
    left.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(address, &FontChange::new().strike(Some(true)))
        .expect("left strike");
    let mut right = source.edit().expect("right edit");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(address, &FontChange::new().strike(Some(false)))
        .expect("right strike");
    let error = left
        .join(right)
        .expect_err("same font property must conflict");
    let conflicts = error.conflicts().expect("cell conflict details");
    assert_eq!(conflicts.conflicts().len(), 1);
    assert_eq!(
        conflicts.conflicts()[0].cells().expect("cell conflicts"),
        &[address]
    );
}

#[test]
fn mixed_fill_and_border_staging_budget_is_atomic_in_both_orders_and_join() {
    let source = low_part_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let fill = FillChange::replace(Fill::Gradient(
        GradientFill::linear(
            90.0,
            vec![
                GradientStop::new(0.0, Color::rgb(0x10, 0x20, 0x30)).expect("first gradient stop"),
                GradientStop::new(1.0, Color::rgb(0xA0, 0xB0, 0xC0)).expect("last gradient stop"),
            ]
            .into_boxed_slice(),
        )
        .expect("gradient fill"),
    ));
    let border = BorderChange::new().left_line(Some(Line::Thin));

    // Probe each category independently. The low source-part limit makes
    // each half-sized request fit alone while their aggregate exceeds the
    // shared fill-plus-border budget.
    let fill_capacity = {
        let mut probe = source.edit().expect("fill probe edit");
        let mut capacity = 0;
        {
            let mut sheet = probe
                .sheet(0usize)
                .expect("fill probe lookup")
                .expect("fill probe worksheet");
            for row in 1..=4096 {
                let address = Address::from_a1(&format!("A{row}")).expect("fill probe address");
                if sheet.fill(address, &fill).is_err() {
                    break;
                }
                capacity += 1;
            }
        }
        capacity
    };
    let border_capacity = {
        let mut probe = source.edit().expect("border probe edit");
        let mut capacity = 0;
        {
            let mut sheet = probe
                .sheet(0usize)
                .expect("border probe lookup")
                .expect("border probe worksheet");
            for row in 1..=4096 {
                let address = Address::from_a1(&format!("A{row}")).expect("border probe address");
                if sheet.border(address, &border).is_err() {
                    break;
                }
                capacity += 1;
            }
        }
        capacity
    };
    assert!(fill_capacity >= 2, "fill budget must admit two entries");
    assert!(border_capacity >= 2, "border budget must admit two entries");
    assert!(fill_capacity < 4096, "fill probe must reach its budget");
    assert!(border_capacity < 4096, "border probe must reach its budget");
    let fill_copies = fill_capacity / 2 + 1;
    let border_copies = border_capacity / 2 + 1;

    // Fill first: a rejected border must not be inserted, and discarding the
    // staged entries must make the same transaction reusable.
    let mut fill_first = source.edit().expect("fill-first edit");
    {
        let mut sheet = fill_first
            .sheet(0usize)
            .expect("fill-first lookup")
            .expect("fill-first worksheet");
        for row in 1..=fill_copies {
            let address = Address::from_a1(&format!("A{row}")).expect("fill-first address");
            sheet.fill(address, &fill).expect("fill-first staging");
        }
    }
    let fill_first_before_failure = fill_first.len();
    let (failed_border, accepted_borders) = {
        let mut sheet = fill_first
            .sheet(0usize)
            .expect("fill-first border lookup")
            .expect("fill-first border worksheet");
        let mut accepted = Vec::new();
        let mut failed = None;
        for row in 1..=border_copies {
            let address = Address::from_a1(&format!("B{row}")).expect("fill-first border address");
            match sheet.border(address, &border) {
                Ok(_) => accepted.push(address),
                Err(error) => {
                    assert!(matches!(error, Error::ResourceLimit(_)));
                    failed = Some(address);
                    break;
                },
            }
        }
        (
            failed.expect("fill-first border must hit aggregate budget"),
            accepted,
        )
    };
    assert_eq!(
        fill_first.len(),
        fill_first_before_failure + accepted_borders.len()
    );
    {
        let mut sheet = fill_first
            .sheet(0usize)
            .expect("fill-first recovery lookup")
            .expect("fill-first recovery worksheet");
        for row in 1..=fill_copies {
            let address = Address::from_a1(&format!("A{row}")).expect("fill-first discard address");
            sheet.reset_style(address).expect("discard fill");
        }
        for address in accepted_borders {
            sheet.reset_style(address).expect("discard accepted border");
        }
        sheet
            .border(failed_border, &border)
            .expect("fill-first recovery border");
    }
    fill_first.commit().expect("fill-first recovery commit");

    // Border first exercises the symmetric failed-fill and discard path.
    let mut border_first = source.edit().expect("border-first edit");
    {
        let mut sheet = border_first
            .sheet(0usize)
            .expect("border-first lookup")
            .expect("border-first worksheet");
        for row in 1..=border_copies {
            let address = Address::from_a1(&format!("A{row}")).expect("border-first address");
            sheet
                .border(address, &border)
                .expect("border-first staging");
        }
    }
    let border_first_before_failure = border_first.len();
    let (failed_fill, accepted_fills) = {
        let mut sheet = border_first
            .sheet(0usize)
            .expect("border-first fill lookup")
            .expect("border-first fill worksheet");
        let mut accepted = Vec::new();
        let mut failed = None;
        for row in 1..=fill_copies {
            let address = Address::from_a1(&format!("B{row}")).expect("border-first fill address");
            match sheet.fill(address, &fill) {
                Ok(_) => accepted.push(address),
                Err(error) => {
                    assert!(matches!(error, Error::ResourceLimit(_)));
                    failed = Some(address);
                    break;
                },
            }
        }
        (
            failed.expect("border-first fill must hit aggregate budget"),
            accepted,
        )
    };
    assert_eq!(
        border_first.len(),
        border_first_before_failure + accepted_fills.len()
    );
    {
        let mut sheet = border_first
            .sheet(0usize)
            .expect("border-first recovery lookup")
            .expect("border-first recovery worksheet");
        for address in accepted_fills {
            sheet.reset_style(address).expect("discard accepted fill");
        }
        for row in 1..=border_copies {
            let address =
                Address::from_a1(&format!("A{row}")).expect("border-first discard address");
            sheet.reset_style(address).expect("discard border");
        }
        sheet
            .fill(failed_fill, &fill)
            .expect("border-first recovery fill");
    }
    border_first.commit().expect("border-first recovery commit");

    // Joining independently valid edits must reject the aggregate in either
    // operand order while preserving both the left edit and the rejected edit.
    for fill_on_left in [true, false] {
        let mut left = source.edit().expect("join left edit");
        let mut right = source.edit().expect("join right edit");
        {
            let mut sheet = left
                .sheet(0usize)
                .expect("join left lookup")
                .expect("join left worksheet");
            if fill_on_left {
                for row in 1..=fill_copies {
                    let address =
                        Address::from_a1(&format!("A{row}")).expect("join left fill address");
                    sheet.fill(address, &fill).expect("join left fill");
                }
            } else {
                for row in 1..=border_copies {
                    let address =
                        Address::from_a1(&format!("A{row}")).expect("join left border address");
                    sheet.border(address, &border).expect("join left border");
                }
            }
        }
        {
            let mut sheet = right
                .sheet(0usize)
                .expect("join right lookup")
                .expect("join right worksheet");
            if fill_on_left {
                for row in 1..=border_copies {
                    let address =
                        Address::from_a1(&format!("B{row}")).expect("join right border address");
                    sheet.border(address, &border).expect("join right border");
                }
            } else {
                for row in 1..=fill_copies {
                    let address =
                        Address::from_a1(&format!("B{row}")).expect("join right fill address");
                    sheet.fill(address, &fill).expect("join right fill");
                }
            }
        }
        let left_len = left.len();
        let right_len = right.len();
        let error = left
            .join(right)
            .expect_err("mixed fill and border join must exceed budget");
        assert!(matches!(
            error.failure(),
            JoinFailure::ResourceLimit(limit)
                if limit.scope.as_ref() == "worksheet style staged bytes"
        ));
        assert_eq!(left.len(), left_len);
        assert_eq!(error.rejected().len(), right_len);
        let (failure, rejected) = error.into_parts();
        assert!(matches!(
            failure,
            JoinFailure::ResourceLimit(limit)
                if limit.scope.as_ref() == "worksheet style staged bytes"
        ));
        left.commit().expect("left independent commit");
        rejected.commit().expect("rejected independent commit");
    }

    assert_eq!(
        source.to_bytes().expect("source remains unchanged"),
        source_bytes
    );
}

#[test]
fn number_format_round_trip_composes_with_native_style_facets_and_payload() {
    let source = crate::Package::new().expect("source workbook");
    let source_bytes = source.to_bytes().expect("source bytes");
    let address = Address::from_a1("A1").expect("address");
    let custom = "0.000 \"kg\"";
    let alignment = AlignmentChange::new()
        .horizontal(Some(Horizontal::Center))
        .wrap_text(Some(true));
    let protection = ProtectionChange::new()
        .locked(Some(false))
        .hidden(Some(true));

    let mut edit = source.edit().expect("edit");
    let mut sheet = edit.sheet(0usize).expect("lookup").expect("worksheet");
    sheet
        .set(address, "payload")
        .and_then(|sheet| sheet.font(address, &FontChange::new().bold(Some(true))))
        .and_then(|sheet| sheet.alignment(address, &alignment))
        .and_then(|sheet| sheet.protection(address, &protection))
        .and_then(|sheet| sheet.quote_prefix(address, Some(true)))
        .and_then(|sheet| sheet.number_format(address, Some(custom)))
        .expect("all style facets");
    let committed = edit.commit().expect("commit");
    assert_eq!(committed.patch().operation_count(), 1);

    let committed_bytes = committed.package().to_bytes().expect("committed bytes");
    let styles = part_text(committed.package(), "/xl/styles.xml");
    assert!(styles.contains(r#"formatCode="0.000 &quot;kg&quot;"#));
    let format = committed
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let number = format.number_format();
    let number_id = crate::style::cell_format::raw::number_format_id(&number);
    assert!(number_id >= crate::style::stylesheet::FIRST_CUSTOM_ID);
    let number_marker = format!(r#"numFmtId="{number_id}""#);
    assert!(styles.contains(number_marker.as_str()));
    assert_eq!(number.code(), Some(custom));
    assert!(number.is_declared());
    assert_eq!(format.font().and_then(|font| font.bold()), Some(true));
    let alignment_view = format.alignment().expect("alignment");
    assert_eq!(alignment_view.horizontal, Some(Horizontal::Center));
    assert!(alignment_view.wrap_text);
    let protection_view = format.protection().expect("protection");
    assert_eq!(protection_view.locked(), Some(false));
    assert_eq!(protection_view.hidden(), Some(true));
    assert!(format.quote_prefix());

    let reopened = crate::Package::from_bytes(&committed_bytes).expect("reopened workbook");
    let reopened_format = reopened
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    assert_eq!(reopened_format.number_format().code(), Some(custom));
    assert_eq!(
        crate::style::cell_format::raw::number_format_id(&reopened_format.number_format()),
        number_id
    );
    assert_eq!(
        reopened_format.font().and_then(|font| font.bold()),
        Some(true)
    );
    assert_eq!(
        reopened_format
            .alignment()
            .and_then(|alignment| alignment.horizontal),
        Some(Horizontal::Center)
    );
    assert_eq!(
        reopened_format
            .protection()
            .and_then(|value| value.locked()),
        Some(false)
    );
    assert!(reopened_format.quote_prefix());

    let replayed = source.apply(committed.patch()).expect("replay patch");
    assert_eq!(
        replayed.package().to_bytes().expect("replayed bytes"),
        committed_bytes
    );
    let restored = committed
        .package()
        .apply(&committed.patch().inverse())
        .expect("inverse patch");
    assert_eq!(
        restored.package().to_bytes().expect("restored bytes"),
        source_bytes
    );
}

#[test]
fn number_format_reuses_builtin_and_declared_codes_and_clears_to_general() {
    let source = number_format_workbook();
    let addresses = [
        Address::from_a1("A1").expect("A1"),
        Address::from_a1("B1").expect("B1"),
        Address::from_a1("C1").expect("C1"),
        Address::from_a1("D1").expect("D1"),
    ];
    let built_in = "0.00";
    let declared = "0.000 \"kg\"";
    let fresh = "0.0000 \"new\"";

    let mut edit = source.edit().expect("edit");
    edit.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .number_format(addresses[0], Some(built_in))
        .and_then(|sheet| sheet.number_format(addresses[1], Some(declared)))
        .and_then(|sheet| sheet.number_format(addresses[2], Some(declared)))
        .and_then(|sheet| sheet.number_format(addresses[3], Some(fresh)))
        .expect("number formats");
    let configured = edit.commit().expect("configured commit");
    let configured_bytes = configured.package().to_bytes().expect("configured bytes");
    let configured_sheet = configured
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet");
    let built_a_format = configured_sheet
        .style(addresses[0])
        .expect("A1 style")
        .expect("A1")
        .format()
        .expect("A1 format");
    let built_a = built_a_format.number_format();
    let declared_b_format = configured_sheet
        .style(addresses[1])
        .expect("B1 style")
        .expect("B1")
        .format()
        .expect("B1 format");
    let declared_b = declared_b_format.number_format();
    let declared_c_format = configured_sheet
        .style(addresses[2])
        .expect("C1 style")
        .expect("C1")
        .format()
        .expect("C1 format");
    let declared_c = declared_c_format.number_format();
    let fresh_d_format = configured_sheet
        .style(addresses[3])
        .expect("D1 style")
        .expect("D1")
        .format()
        .expect("D1 format");
    let fresh_d = fresh_d_format.number_format();
    assert_eq!(built_a.code(), Some(built_in));
    assert_eq!(
        crate::style::cell_format::raw::number_format_id(&built_a),
        2
    );
    assert!(!built_a.is_declared());
    assert_eq!(declared_b.code(), Some(declared));
    assert_eq!(declared_c.code(), Some(declared));
    assert!(declared_b.is_declared());
    assert_eq!(
        crate::style::cell_format::raw::number_format_id(&declared_b),
        200
    );
    assert_eq!(
        crate::style::cell_format::raw::number_format_id(&declared_b),
        crate::style::cell_format::raw::number_format_id(&declared_c)
    );
    assert_eq!(fresh_d.code(), Some(fresh));
    let fresh_id = crate::style::cell_format::raw::number_format_id(&fresh_d);
    assert!(fresh_id >= crate::style::stylesheet::FIRST_CUSTOM_ID);
    assert_ne!(fresh_id, 200);

    let styles = part_text(configured.package(), "/xl/styles.xml");
    assert_eq!(styles.matches("<numFmt ").count(), 2);
    assert_eq!(
        styles
            .matches(r#"formatCode="0.000 &quot;kg&quot;""#)
            .count(),
        1
    );
    assert_eq!(
        styles
            .matches(r#"formatCode="0.0000 &quot;new&quot;""#)
            .count(),
        1
    );
    let fresh_marker = format!(r#"numFmtId="{fresh_id}""#);
    assert!(styles.contains(fresh_marker.as_str()));

    let mut repeat = configured.package().edit().expect("repeat edit");
    repeat
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .number_format(addresses[0], Some(built_in))
        .and_then(|sheet| sheet.number_format(addresses[1], Some(declared)))
        .and_then(|sheet| sheet.number_format(addresses[2], Some(declared)))
        .and_then(|sheet| sheet.number_format(addresses[3], Some(fresh)))
        .expect("repeat number formats");
    let repeat = repeat.commit().expect("repeat commit");
    assert!(repeat.patch().is_noop());
    assert_eq!(
        repeat.package().to_bytes().expect("repeat bytes"),
        configured_bytes
    );

    let mut clear = configured.package().edit().expect("clear edit");
    clear
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .number_format(addresses[1], None)
        .expect("clear number format");
    let cleared = clear.commit().expect("clear commit");
    let cleared_format = cleared
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(addresses[1])
        .expect("style lookup")
        .expect("cleared cell")
        .format()
        .expect("cleared format");
    assert!(cleared_format.number_format().is_general());
    assert_eq!(
        crate::style::cell_format::raw::number_format_id(&cleared_format.number_format()),
        0
    );
    assert_eq!(cleared_format.number_format().code(), Some("General"));
    let cleared_styles = part_text(cleared.package(), "/xl/styles.xml");
    assert_eq!(
        cleared_styles
            .matches(r#"formatCode="0.000 &quot;kg&quot;""#)
            .count(),
        1
    );

    let mut clear_repeat = cleared.package().edit().expect("clear repeat edit");
    clear_repeat
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .number_format(addresses[1], None)
        .expect("repeat clear number format");
    let clear_repeat = clear_repeat.commit().expect("clear repeat commit");
    assert!(clear_repeat.patch().is_noop());
    assert_eq!(
        clear_repeat
            .package()
            .to_bytes()
            .expect("clear repeat bytes"),
        cleared.package().to_bytes().expect("cleared bytes")
    );
}

#[test]
fn number_format_joins_disjoint_style_facets_and_conflicts_on_same_field() {
    let source = crate::Package::new().expect("source workbook");
    let address = Address::from_a1("A1").expect("address");

    let mut number = source.edit().expect("number edit");
    number
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .number_format(address, Some("0.00"))
        .expect("number format");
    let mut facets = source.edit().expect("facets edit");
    facets
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(address, &FontChange::new().bold(Some(true)))
        .and_then(|sheet| {
            sheet.alignment(
                address,
                &AlignmentChange::new().horizontal(Some(Horizontal::Center)),
            )
        })
        .and_then(|sheet| sheet.protection(address, &ProtectionChange::new().locked(Some(false))))
        .and_then(|sheet| sheet.quote_prefix(address, Some(true)))
        .expect("other style facets");
    number.join(facets).expect("disjoint style facets");
    let joined = number.commit().expect("joined commit");
    let format = joined
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    assert_eq!(format.number_format().code(), Some("0.00"));
    assert_eq!(format.font().and_then(|font| font.bold()), Some(true));
    assert_eq!(
        format
            .alignment()
            .and_then(|alignment| alignment.horizontal),
        Some(Horizontal::Center)
    );
    assert_eq!(
        format.protection().and_then(|value| value.locked()),
        Some(false)
    );
    assert!(format.quote_prefix());

    let mut left = source.edit().expect("left edit");
    left.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .number_format(address, Some("0.00"))
        .expect("left number format");
    let mut right = source.edit().expect("right edit");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .number_format(address, Some("0.00%"))
        .expect("right number format");
    let error = left
        .join(right)
        .expect_err("same number-format field must conflict");
    let conflicts = error.conflicts().expect("cell conflict details");
    assert_eq!(conflicts.conflicts().len(), 1);
    assert_eq!(
        conflicts.conflicts()[0].cells().expect("cell conflicts"),
        &[address]
    );

    let base = source.styles().expect("styles").get(0).expect("base style");
    let mut number = source.edit().expect("number edit");
    number
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .number_format(address, Some("0.00"))
        .expect("number format");
    let mut complete = source.edit().expect("complete style edit");
    complete
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address, &base)
        .expect("complete style");
    let error = number
        .join(complete)
        .expect_err("number format and complete style must conflict");
    let conflicts = error.conflicts().expect("cell conflict details");
    assert_eq!(conflicts.conflicts().len(), 1);
    assert_eq!(
        conflicts.conflicts()[0].cells().expect("cell conflicts"),
        &[address]
    );
}

#[test]
fn number_format_staging_is_discarded_by_style_reset_replacement_and_removal() {
    let source = styled_workbook();
    let source_styles = part_text(&source, "/xl/styles.xml").to_owned();
    let accent = source
        .styles()
        .expect("styles")
        .get(1)
        .expect("accent style");
    let replacement_address = Address::from_a1("B1").expect("replacement address");
    let styled_address = Address::from_a1("A1").expect("styled address");
    let custom = "0.000 \"kg\"";

    let mut replacement = source.edit().expect("replacement edit");
    replacement
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .number_format(replacement_address, Some(custom))
        .and_then(|sheet| sheet.style(replacement_address, &accent))
        .expect("number format then style replacement");
    let replacement = replacement.commit().expect("replacement commit");
    let replacement_format = replacement
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(replacement_address)
        .expect("style lookup")
        .expect("resolved style")
        .format()
        .expect("cell format");
    assert_eq!(replacement_format.number_format().code(), Some("0.00"));
    assert_ne!(replacement_format.number_format().code(), Some(custom));

    let mut reset = source.edit().expect("reset edit");
    reset
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .number_format(styled_address, Some(custom))
        .and_then(|sheet| sheet.reset_style(styled_address))
        .expect("number format then reset");
    let reset = reset.commit().expect("reset commit");
    let reset_sheet = reset
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet");
    assert!(matches!(
        reset_sheet
            .local_style(styled_address)
            .expect("local style"),
        Some(crate::LocalStyle::Default)
    ));
    let reset_format = reset_sheet
        .style(styled_address)
        .expect("style lookup")
        .expect("resolved style")
        .format()
        .expect("cell format");
    assert!(reset_format.number_format().is_general());

    let mut remove = source.edit().expect("remove edit");
    remove
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .number_format(styled_address, Some(custom))
        .and_then(|sheet| sheet.remove(styled_address))
        .expect("number format then removal");
    let remove = remove.commit().expect("remove commit");
    assert!(
        remove
            .package()
            .sheet(0usize)
            .expect("lookup")
            .expect("worksheet")
            .cell(styled_address)
            .expect("removed cell")
            .is_missing()
    );
    assert_eq!(part_text(remove.package(), "/xl/styles.xml"), source_styles);
}

#[test]
fn number_format_rewrite_preserves_opaque_numfmt_and_xf_markup() {
    let source = number_format_workbook();
    let source_styles = part_text(&source, "/xl/styles.xml").to_owned();
    let address = Address::from_a1("A1").expect("address");
    let fresh = "0.0000 \"new\"";

    let mut edit = source.edit().expect("edit");
    edit.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .number_format(address, Some(fresh))
        .expect("number format");
    let committed = edit.commit().expect("commit");
    let styles = part_text(committed.package(), "/xl/styles.xml");
    assert_eq!(
        styles.matches(r#"f:keep="yes""#).count(),
        source_styles.matches(r#"f:keep="yes""#).count()
    );
    assert_eq!(
        styles.matches("<f:format-future/>").count(),
        source_styles.matches("<f:format-future/>").count()
    );
    assert_eq!(
        styles.matches(r#"f:xf="yes""#).count(),
        source_styles.matches(r#"f:xf="yes""#).count() * 2
    );
    assert_eq!(
        styles.matches("<f:xf-future/>").count(),
        source_styles.matches("<f:xf-future/>").count() * 2
    );
    let reopened = crate::Package::from_bytes(&committed.package().to_bytes().expect("bytes"))
        .expect("reopened workbook");
    let format = reopened
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    assert_eq!(format.number_format().code(), Some(fresh));
}

#[test]
fn fill_change_round_trip_composes_with_native_style_facets_and_payload() {
    let source = crate::Package::new().expect("source workbook");
    let source_bytes = source.to_bytes().expect("source bytes");
    let address = Address::from_a1("A1").expect("address");
    let tint = Tint::new(0.25).expect("tint");
    let foreground = Color::rgb(0x11, 0x22, 0x33).with_tint(tint);
    let background = Color::theme(5).with_tint(tint);
    let fill = FillChange::new()
        .pattern(Some(Pattern::Solid))
        .foreground(Some(foreground))
        .background(Some(background));
    let alignment = AlignmentChange::new()
        .horizontal(Some(Horizontal::Center))
        .wrap_text(Some(true));
    let protection = ProtectionChange::new()
        .locked(Some(false))
        .hidden(Some(true));

    let mut edit = source.edit().expect("edit");
    let mut sheet = edit.sheet(0usize).expect("lookup").expect("worksheet");
    sheet
        .set(address, "payload")
        .and_then(|sheet| sheet.font(address, &FontChange::new().bold(Some(true))))
        .and_then(|sheet| sheet.fill(address, &fill))
        .and_then(|sheet| sheet.alignment(address, &alignment))
        .and_then(|sheet| sheet.protection(address, &protection))
        .and_then(|sheet| sheet.quote_prefix(address, Some(true)))
        .expect("all style facets");
    let committed = edit.commit().expect("commit");
    assert_eq!(committed.patch().operation_count(), 1);

    let committed_bytes = committed.package().to_bytes().expect("committed bytes");
    let styles = part_text(committed.package(), "/xl/styles.xml");
    assert!(styles.contains(r#"patternType="solid""#));
    assert!(styles.contains(r#"rgb="FF112233""#));
    assert!(styles.contains(r#"theme="5""#));
    let format = committed
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let pattern = format.fill().and_then(Fill::pattern).expect("pattern fill");
    assert_eq!(pattern.pattern(), Some(Pattern::Solid));
    assert_eq!(pattern.foreground(), Some(foreground));
    assert_eq!(pattern.background(), Some(background));
    assert_eq!(format.font().and_then(|font| font.bold()), Some(true));
    assert_eq!(
        format
            .alignment()
            .and_then(|alignment| alignment.horizontal),
        Some(Horizontal::Center)
    );
    assert_eq!(
        format.protection().and_then(|value| value.locked()),
        Some(false)
    );
    assert!(format.quote_prefix());

    let reopened = crate::Package::from_bytes(&committed_bytes).expect("reopened workbook");
    let reopened_format = reopened
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let reopened_pattern = reopened_format
        .fill()
        .and_then(Fill::pattern)
        .expect("reopened pattern fill");
    assert_eq!(reopened_pattern.pattern(), Some(Pattern::Solid));
    assert_eq!(reopened_pattern.foreground(), Some(foreground));
    assert_eq!(reopened_pattern.background(), Some(background));
    assert_eq!(
        reopened_format.font().and_then(|font| font.bold()),
        Some(true)
    );
    assert!(reopened_format.quote_prefix());

    let replayed = source.apply(committed.patch()).expect("replay patch");
    assert_eq!(
        replayed.package().to_bytes().expect("replayed bytes"),
        committed_bytes
    );
    let restored = committed
        .package()
        .apply(&committed.patch().inverse())
        .expect("inverse patch");
    assert_eq!(
        restored.package().to_bytes().expect("restored bytes"),
        source_bytes
    );
}

#[test]
fn fill_change_colors_are_typed_and_selected_fields_do_not_imply_solid() {
    let source = crate::Package::new().expect("source workbook");
    let addresses = [
        Address::from_a1("A1").expect("A1"),
        Address::from_a1("B1").expect("B1"),
        Address::from_a1("C1").expect("C1"),
        Address::from_a1("D1").expect("D1"),
    ];
    let tint = Tint::new(-0.25).expect("tint");
    let colors = [
        Color::rgb(0x11, 0x22, 0x33).with_tint(tint),
        Color::theme(7).with_tint(tint),
        Color::indexed(12),
        Color::auto(),
    ];

    let mut edit = source.edit().expect("edit");
    edit.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill(
            addresses[0],
            &FillChange::new()
                .pattern(Some(Pattern::Solid))
                .foreground(Some(colors[0])),
        )
        .and_then(|sheet| {
            sheet.fill(
                addresses[1],
                &FillChange::new()
                    .pattern(Some(Pattern::LightGray))
                    .foreground(Some(colors[1])),
            )
        })
        .and_then(|sheet| {
            sheet.fill(
                addresses[2],
                &FillChange::new()
                    .pattern(Some(Pattern::DarkGray))
                    .background(Some(colors[2])),
            )
        })
        .and_then(|sheet| {
            sheet.fill(
                addresses[3],
                &FillChange::new()
                    .pattern(Some(Pattern::Solid))
                    .foreground(Some(colors[3])),
            )
        })
        .expect("typed colors");
    let configured = edit.commit().expect("configured commit");
    let sheet = configured
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet");
    for (address, expected_pattern, expected_foreground, expected_background) in [
        (addresses[0], Pattern::Solid, Some(colors[0]), None),
        (addresses[1], Pattern::LightGray, Some(colors[1]), None),
        (addresses[2], Pattern::DarkGray, None, Some(colors[2])),
        (addresses[3], Pattern::Solid, Some(colors[3]), None),
    ] {
        let format = sheet
            .style(address)
            .expect("style lookup")
            .expect("styled cell")
            .format()
            .expect("cell format");
        let pattern = format.fill().and_then(Fill::pattern).expect("pattern fill");
        assert_eq!(pattern.pattern(), Some(expected_pattern));
        assert_eq!(pattern.foreground(), expected_foreground);
        assert_eq!(pattern.background(), expected_background);
    }

    let mut selected = source.edit().expect("selected edit");
    selected
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill(
            addresses[0],
            &FillChange::new().foreground(Some(Color::rgb(0xAA, 0xBB, 0xCC))),
        )
        .expect("foreground only");
    let selected = selected.commit().expect("selected commit");
    let selected_format = selected
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(addresses[0])
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let selected_pattern = selected_format
        .fill()
        .and_then(Fill::pattern)
        .expect("selected pattern fill");
    assert_eq!(selected_pattern.pattern(), Some(Pattern::None));
    assert_eq!(
        selected_pattern.foreground(),
        Some(Color::rgb(0xAA, 0xBB, 0xCC))
    );
    assert_eq!(selected_pattern.background(), None);
    let selected_bytes = selected.package().to_bytes().expect("selected bytes");

    let mut repeat = selected.package().edit().expect("repeat edit");
    repeat
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill(
            addresses[0],
            &FillChange::new().foreground(Some(Color::rgb(0xAA, 0xBB, 0xCC))),
        )
        .expect("repeat foreground only");
    let repeat = repeat.commit().expect("repeat commit");
    assert!(repeat.patch().is_noop());
    assert_eq!(
        repeat.package().to_bytes().expect("repeat bytes"),
        selected_bytes
    );
}

#[test]
fn fill_change_defaults_clear_and_replay_are_distinct_and_exact() {
    let source = styled_workbook();
    let address = Address::from_a1("A1").expect("address");
    let solid_none = FillChange::new().pattern(Some(Pattern::None));

    let mut explicit = source.edit().expect("explicit edit");
    explicit
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill(address, &solid_none)
        .expect("explicit none pattern");
    let explicit = explicit.commit().expect("explicit commit");
    let explicit_format = explicit
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let explicit_pattern = explicit_format
        .fill()
        .and_then(Fill::pattern)
        .expect("explicit pattern");
    assert_eq!(explicit_pattern.pattern(), Some(Pattern::None));

    let explicit_bytes = explicit.package().to_bytes().expect("explicit bytes");
    let mut repeat = explicit.package().edit().expect("repeat edit");
    repeat
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill(address, &solid_none)
        .expect("repeat explicit none pattern");
    let repeat = repeat.commit().expect("repeat commit");
    assert!(repeat.patch().is_noop());
    assert_eq!(
        repeat.package().to_bytes().expect("repeat bytes"),
        explicit_bytes
    );

    let mut clear = explicit.package().edit().expect("clear edit");
    clear
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill(address, &FillChange::new().pattern(None).foreground(None))
        .expect("clear pattern and foreground");
    let clear = clear.commit().expect("clear commit");
    let clear_format = clear
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let clear_pattern = clear_format
        .fill()
        .and_then(Fill::pattern)
        .expect("cleared pattern");
    assert_eq!(clear_pattern.pattern(), None);
    assert_eq!(clear_pattern.foreground(), None);
}

#[test]
fn fill_change_joins_disjoint_fields_and_conflicts_on_overlap() {
    let source = crate::Package::new().expect("source workbook");
    let address = Address::from_a1("A1").expect("address");

    let mut pattern = source.edit().expect("pattern edit");
    pattern
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill(address, &FillChange::new().pattern(Some(Pattern::Solid)))
        .expect("pattern");
    let mut foreground = source.edit().expect("foreground edit");
    foreground
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill(
            address,
            &FillChange::new().foreground(Some(Color::rgb(0x11, 0x22, 0x33))),
        )
        .expect("foreground");
    pattern.join(foreground).expect("disjoint fill fields");
    let joined = pattern.commit().expect("joined commit");
    let joined_format = joined
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let joined_pattern = joined_format
        .fill()
        .and_then(Fill::pattern)
        .expect("joined pattern");
    assert_eq!(joined_pattern.pattern(), Some(Pattern::Solid));
    assert_eq!(
        joined_pattern.foreground(),
        Some(Color::rgb(0x11, 0x22, 0x33))
    );

    let mut left = source.edit().expect("left edit");
    left.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill(address, &FillChange::new().pattern(Some(Pattern::Solid)))
        .expect("left pattern");
    let mut right = source.edit().expect("right edit");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill(address, &FillChange::new().pattern(Some(Pattern::None)))
        .expect("right pattern");
    let error = left
        .join(right)
        .expect_err("same fill pattern field must conflict");
    let conflicts = error.conflicts().expect("cell conflict details");
    assert_eq!(conflicts.conflicts().len(), 1);
    assert_eq!(
        conflicts.conflicts()[0].cells().expect("cell conflicts"),
        &[address]
    );

    let mut left = source.edit().expect("left color edit");
    left.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill(
            address,
            &FillChange::new().foreground(Some(Color::rgb(0x11, 0x22, 0x33))),
        )
        .expect("left color");
    let mut right = source.edit().expect("right color edit");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill(
            address,
            &FillChange::new().foreground(Some(Color::rgb(0x44, 0x55, 0x66))),
        )
        .expect("right color");
    let error = left
        .join(right)
        .expect_err("same fill color field must conflict");
    let conflicts = error.conflicts().expect("cell conflict details");
    assert_eq!(conflicts.conflicts().len(), 1);
    assert_eq!(
        conflicts.conflicts()[0].cells().expect("cell conflicts"),
        &[address]
    );
}

#[test]
fn fill_change_staging_is_discarded_by_style_reset_replacement_and_removal() {
    let source = styled_workbook();
    let source_styles = part_text(&source, "/xl/styles.xml").to_owned();
    let accent = source
        .styles()
        .expect("styles")
        .get(1)
        .expect("accent style");
    let replacement_address = Address::from_a1("B1").expect("replacement address");
    let styled_address = Address::from_a1("A1").expect("styled address");
    let requested = FillChange::new()
        .pattern(Some(Pattern::Solid))
        .foreground(Some(Color::rgb(0x11, 0x22, 0x33)));

    let mut replacement = source.edit().expect("replacement edit");
    replacement
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill(replacement_address, &requested)
        .and_then(|sheet| sheet.style(replacement_address, &accent))
        .expect("fill then style replacement");
    let replacement = replacement.commit().expect("replacement commit");
    let replacement_format = replacement
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(replacement_address)
        .expect("style lookup")
        .expect("resolved style")
        .format()
        .expect("cell format");
    let replacement_pattern = replacement_format
        .fill()
        .and_then(Fill::pattern)
        .expect("replacement pattern");
    assert_eq!(
        replacement_pattern.foreground(),
        Some(Color::rgb(255, 255, 0))
    );
    assert_ne!(
        replacement_pattern.foreground(),
        Some(Color::rgb(0x11, 0x22, 0x33))
    );

    let mut reset = source.edit().expect("reset edit");
    reset
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill(styled_address, &requested)
        .and_then(|sheet| sheet.reset_style(styled_address))
        .expect("fill then reset");
    let reset = reset.commit().expect("reset commit");
    let reset_sheet = reset
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet");
    assert!(matches!(
        reset_sheet
            .local_style(styled_address)
            .expect("local style"),
        Some(crate::LocalStyle::Default)
    ));
    let reset_format = reset_sheet
        .style(styled_address)
        .expect("style lookup")
        .expect("resolved style")
        .format()
        .expect("cell format");
    assert!(reset_format.has_default_fill());

    let mut remove = source.edit().expect("remove edit");
    remove
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill(styled_address, &requested)
        .and_then(|sheet| sheet.remove(styled_address))
        .expect("fill then removal");
    let remove = remove.commit().expect("remove commit");
    assert!(
        remove
            .package()
            .sheet(0usize)
            .expect("lookup")
            .expect("worksheet")
            .cell(styled_address)
            .expect("removed cell")
            .is_missing()
    );
    assert_eq!(part_text(remove.package(), "/xl/styles.xml"), source_styles);
}

#[test]
fn fill_change_rewrite_preserves_opaque_fill_and_xf_markup() {
    let source = fill_workbook();
    let source_styles = part_text(&source, "/xl/styles.xml").to_owned();
    let address = Address::from_a1("A1").expect("address");
    let requested = FillChange::new().foreground(Some(Color::theme(4)));

    let mut edit = source.edit().expect("edit");
    edit.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill(address, &requested)
        .expect("fill");
    let committed = edit.commit().expect("commit");
    let styles = part_text(committed.package(), "/xl/styles.xml");
    assert_eq!(
        styles.matches(r#"f:target-fill="yes""#).count(),
        source_styles.matches(r#"f:target-fill="yes""#).count() * 2
    );
    assert_eq!(
        styles.matches(r#"f:target-pattern="yes""#).count(),
        source_styles.matches(r#"f:target-pattern="yes""#).count() * 2
    );
    assert_eq!(
        styles.matches("<f:fill-future/>").count(),
        source_styles.matches("<f:fill-future/>").count() * 2
    );
    assert_eq!(
        styles.matches(r#"f:target-xf="yes""#).count(),
        source_styles.matches(r#"f:target-xf="yes""#).count() * 2
    );
    assert_eq!(
        styles.matches("<f:xf-future/>").count(),
        source_styles.matches("<f:xf-future/>").count() * 2
    );
    let reopened = crate::Package::from_bytes(&committed.package().to_bytes().expect("bytes"))
        .expect("reopened workbook");
    let format = reopened
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    assert_eq!(
        format
            .fill()
            .and_then(Fill::pattern)
            .and_then(|pattern| pattern.foreground()),
        Some(Color::theme(4))
    );
}

#[test]
fn gradient_fill_replacements_round_trip_typed_stops_and_companion_styles() {
    let source = crate::Package::new().expect("source workbook");
    let source_bytes = source.to_bytes().expect("source bytes");
    let linear = authored_linear_gradient();
    let path = authored_path_gradient();
    let first = Address::from_a1("A1").expect("A1");
    let second = Address::from_a1("B1").expect("B1");
    let alignment = AlignmentChange::new()
        .horizontal(Some(Horizontal::Center))
        .wrap_text(Some(true));
    let protection = ProtectionChange::new().locked(Some(false));

    let mut edit = source.edit().expect("edit");
    let mut sheet = edit.sheet(0usize).expect("lookup").expect("worksheet");
    sheet
        .set(first, "gradient payload")
        .and_then(|sheet| sheet.font(first, &FontChange::new().bold(Some(true))))
        .and_then(|sheet| sheet.alignment(first, &alignment))
        .and_then(|sheet| sheet.protection(first, &protection))
        .and_then(|sheet| sheet.quote_prefix(first, Some(true)))
        .and_then(|sheet| sheet.fill(first, &FillChange::replace(Fill::Gradient(linear.clone()))))
        .and_then(|sheet| sheet.fill(second, &FillChange::replace(Fill::Gradient(path.clone()))))
        .expect("gradient and companion edits");
    let committed = edit.commit().expect("gradient commit");
    let committed_bytes = committed.package().to_bytes().expect("committed bytes");
    let styles = part_text(committed.package(), "/xl/styles.xml");
    assert!(styles.matches("gradientFill").count() >= 2);

    let sheet = committed
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet");
    let first_format = sheet
        .style(first)
        .expect("first style lookup")
        .expect("first style")
        .format()
        .expect("first format");
    let first_gradient = first_format
        .fill()
        .and_then(Fill::gradient)
        .expect("linear gradient");
    assert_eq!(first_gradient.kind(), GradientKind::Linear);
    assert_eq!(first_gradient.degree(), linear.degree());
    assert_eq!(first_gradient.stops(), linear.stops());
    assert_eq!(first_format.font().and_then(|font| font.bold()), Some(true));
    assert_eq!(
        first_format.alignment().and_then(|value| value.horizontal),
        Some(Horizontal::Center)
    );
    assert_eq!(
        first_format.protection().and_then(|value| value.locked()),
        Some(false)
    );
    assert!(first_format.quote_prefix());
    assert!(matches!(
        sheet.cell(first).expect("payload cell").stored(),
        Some(Cell::Value(Value::Text(value))) if value.as_str() == "gradient payload"
    ));

    let second_format = sheet
        .style(second)
        .expect("second style lookup")
        .expect("second style")
        .format()
        .expect("second format");
    let second_gradient = second_format
        .fill()
        .and_then(Fill::gradient)
        .expect("path gradient");
    assert_eq!(second_gradient.kind(), GradientKind::Path);
    assert_eq!(
        (
            second_gradient.left(),
            second_gradient.right(),
            second_gradient.top(),
            second_gradient.bottom(),
        ),
        (path.left(), path.right(), path.top(), path.bottom())
    );
    assert_eq!(second_gradient.stops(), path.stops());

    let reopened = crate::Package::from_bytes(&committed_bytes).expect("reopened workbook");
    let reopened_sheet = reopened
        .sheet(0usize)
        .expect("reopened lookup")
        .expect("reopened worksheet");
    for (address, expected, kind) in [
        (first, &linear, GradientKind::Linear),
        (second, &path, GradientKind::Path),
    ] {
        let reopened_format = reopened_sheet
            .style(address)
            .expect("reopened style lookup")
            .expect("reopened style")
            .format()
            .expect("reopened format");
        let gradient = reopened_format
            .fill()
            .and_then(Fill::gradient)
            .expect("reopened gradient");
        assert_eq!(gradient.kind(), kind);
        assert_eq!(gradient.stops(), expected.stops());
    }

    let replayed = source
        .apply(committed.patch())
        .expect("replay gradient patch");
    assert_eq!(
        replayed.package().to_bytes().expect("replayed bytes"),
        committed_bytes
    );
    let restored = committed
        .package()
        .apply(&committed.patch().inverse())
        .expect("inverse gradient patch");
    assert_eq!(
        restored.package().to_bytes().expect("restored bytes"),
        source_bytes
    );
}

#[test]
fn fill_replacement_swaps_pattern_and_gradient() {
    let source = crate::Package::new().expect("source workbook");
    let address = Address::from_a1("A1").expect("address");
    let pattern = PatternFill::new(
        Some(Pattern::Solid),
        Some(Color::rgb(0x11, 0x22, 0x33)),
        Some(Color::theme(4)),
    );
    let gradient = authored_linear_gradient();

    let mut pattern_edit = source.edit().expect("pattern edit");
    pattern_edit
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill(address, &FillChange::replace(Fill::Pattern(pattern)))
        .expect("pattern replacement");
    let patterned = pattern_edit.commit().expect("pattern commit");
    let patterned_format = patterned
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("style")
        .format()
        .expect("format");
    assert_eq!(
        patterned_format
            .fill()
            .and_then(Fill::pattern)
            .and_then(PatternFill::pattern),
        Some(Pattern::Solid)
    );

    let mut gradient_edit = patterned.package().edit().expect("gradient edit");
    gradient_edit
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill(
            address,
            &FillChange::replace(Fill::Gradient(gradient.clone())),
        )
        .expect("gradient replacement");
    let gradient_commit = gradient_edit.commit().expect("gradient commit");
    let gradient_format = gradient_commit
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("style")
        .format()
        .expect("format");
    let actual_gradient = gradient_format
        .fill()
        .and_then(Fill::gradient)
        .expect("gradient fill");
    assert_eq!(actual_gradient.kind(), GradientKind::Linear);
    assert_eq!(actual_gradient.stops(), gradient.stops());

    let mut pattern_again = gradient_commit.package().edit().expect("pattern edit");
    pattern_again
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill(
            address,
            &FillChange::replace(Fill::Pattern(PatternFill::new(
                Some(Pattern::DarkGray),
                None,
                None,
            ))),
        )
        .expect("pattern replacement");
    let pattern_again = pattern_again.commit().expect("pattern commit");
    assert_eq!(
        pattern_again
            .package()
            .sheet(0usize)
            .expect("lookup")
            .expect("worksheet")
            .style(address)
            .expect("style lookup")
            .expect("style")
            .format()
            .expect("format")
            .fill()
            .and_then(Fill::pattern)
            .and_then(PatternFill::pattern),
        Some(Pattern::DarkGray)
    );
}

#[test]
fn matching_gradient_replacement_is_an_exact_lexical_no_op() {
    let source = lexical_gradient_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let expected = GradientFill::linear(
        45.0,
        vec![
            GradientStop::new(0.0, Color::rgb(0x11, 0x22, 0x33)).expect("RGB stop"),
            GradientStop::new(
                1.0,
                Color::theme(5).with_tint(Tint::new(0.5).expect("theme tint")),
            )
            .expect("theme stop"),
        ]
        .into_boxed_slice(),
    )
    .expect("expected gradient");

    let mut edit = source.edit().expect("edit");
    edit.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill("A1", &FillChange::replace(Fill::Gradient(expected.clone())))
        .expect("matching replacement");
    let committed = edit.commit().expect("no-op commit");
    assert!(committed.patch().is_noop());
    assert_eq!(
        committed.package().to_bytes().expect("no-op bytes"),
        source_bytes
    );
    let actual_format = committed
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style("A1")
        .expect("style lookup")
        .expect("style")
        .format()
        .expect("format");
    let actual = actual_format
        .fill()
        .and_then(Fill::gradient)
        .expect("gradient");
    assert_eq!(actual, &expected);
}

#[test]
fn fill_replacement_joins_conflict_and_later_replacement_wins() {
    let source = crate::Package::new().expect("source workbook");
    let address = Address::from_a1("A1").expect("address");
    let pattern = FillChange::replace(Fill::Pattern(PatternFill::new(
        Some(Pattern::Solid),
        Some(Color::rgb(0x10, 0x20, 0x30)),
        None,
    )));
    let gradient = FillChange::replace(Fill::Gradient(authored_linear_gradient()));

    let mut left = source.edit().expect("left edit");
    left.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill(address, &pattern)
        .expect("left replacement");
    let mut right = source.edit().expect("right edit");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill(address, &gradient)
        .expect("right replacement");
    let error = left
        .join(right)
        .expect_err("whole-fill replacements must conflict");
    let conflicts = error.conflicts().expect("cell conflict details");
    assert_eq!(conflicts.conflicts().len(), 1);
    assert_eq!(
        conflicts.conflicts()[0].cells().expect("cell conflicts"),
        &[address]
    );

    let mut staged = source.edit().expect("staged edit");
    staged
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill(address, &pattern)
        .and_then(|sheet| sheet.fill(address, &gradient))
        .expect("staged replacements");
    let staged = staged.commit().expect("staged commit");
    assert!(
        staged
            .package()
            .sheet(0usize)
            .expect("lookup")
            .expect("worksheet")
            .style(address)
            .expect("style lookup")
            .expect("style")
            .format()
            .expect("format")
            .fill()
            .and_then(Fill::gradient)
            .is_some()
    );
}

#[test]
fn granular_fill_after_pattern_replacement_composes_but_gradient_is_refused() {
    let source = crate::Package::new().expect("source workbook");
    let address = Address::from_a1("A1").expect("address");
    let replacement = FillChange::replace(Fill::Pattern(PatternFill::new(
        Some(Pattern::Solid),
        None,
        None,
    )));
    let foreground = FillChange::new().foreground(Some(Color::theme(3)));

    let mut pattern_edit = source.edit().expect("pattern edit");
    pattern_edit
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill(address, &replacement)
        .and_then(|sheet| sheet.fill(address, &foreground))
        .expect("pattern replacement and granular fill");
    let patterned = pattern_edit.commit().expect("pattern commit");
    let patterned_format = patterned
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("style")
        .format()
        .expect("format");
    let pattern = patterned_format
        .fill()
        .and_then(Fill::pattern)
        .expect("pattern fill");
    assert_eq!(pattern.pattern(), Some(Pattern::Solid));
    assert_eq!(pattern.foreground(), Some(Color::theme(3)));

    let gradient_replacement = FillChange::replace(Fill::Gradient(authored_linear_gradient()));
    let mut gradient_edit = source.edit().expect("gradient edit");
    gradient_edit
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .fill(address, &gradient_replacement)
        .and_then(|sheet| sheet.fill(address, &foreground))
        .expect("gradient replacement and granular fill");
    let error = gradient_edit
        .commit()
        .expect_err("pattern-only granular fields cannot edit gradients");
    assert!(matches!(error, Error::Unsupported { .. }), "{error:?}");
}

#[test]
fn failed_gradient_stage_does_not_poison_the_fill_budget() {
    let baseline = crate::Package::new().expect("baseline workbook");
    let baseline_bytes = baseline.to_bytes().expect("baseline bytes");
    let maximum_part = baseline
        .inner
        .package
        .try_iter_parts()
        .map(|part| part.expect("baseline part").blob().len())
        .max()
        .expect("baseline parts")
        .saturating_add(4096);
    let limits = SourceLimits::builder()
        .max_part_bytes(maximum_part as u64)
        .expect("part limit")
        .build()
        .expect("source limits");
    let source = crate::Package::from_bytes_with_options(
        &baseline_bytes,
        crate::OpenOptions::new().source_limits(limits),
    )
    .expect("limited workbook");
    let address = Address::from_a1("A1").expect("address");
    let stops = (0..4096)
        .map(|index| {
            GradientStop::new(index as f64 / 4095.0, Color::rgb(0x10, 0x20, 0x30))
                .expect("gradient stop")
        })
        .collect::<Vec<_>>()
        .into_boxed_slice();
    let oversized = FillChange::replace(Fill::Gradient(
        GradientFill::linear(90.0, stops).expect("oversized gradient"),
    ));
    let fallback = FillChange::replace(Fill::Pattern(PatternFill::new(
        Some(Pattern::Solid),
        Some(Color::rgb(0xAA, 0xBB, 0xCC)),
        None,
    )));

    let mut edit = source.edit().expect("edit");
    let mut sheet = edit.sheet(0usize).expect("lookup").expect("worksheet");
    let failed = sheet.fill(address, &oversized);
    assert!(matches!(
        failed,
        Err(Error::ResourceLimit(limit))
            if limit.scope.as_ref() == "worksheet fill staged bytes"
    ));
    sheet.fill(address, &fallback).expect("fallback fill");
    let committed = edit.commit().expect("fallback commit");
    let committed_format = committed
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("style")
        .format()
        .expect("format");
    let pattern = committed_format
        .fill()
        .and_then(Fill::pattern)
        .expect("fallback pattern");
    assert_eq!(pattern.pattern(), Some(Pattern::Solid));
    assert_eq!(pattern.foreground(), Some(Color::rgb(0xAA, 0xBB, 0xCC)));
}

#[test]
fn gradient_join_resource_limit_returns_rejected_edit_and_allows_individual_commits() {
    let baseline = crate::Package::new().expect("baseline workbook");
    let baseline_bytes = baseline.to_bytes().expect("baseline bytes");
    let maximum_part = baseline
        .inner
        .package
        .try_iter_parts()
        .map(|part| part.expect("baseline part").blob().len())
        .max()
        .expect("baseline parts")
        .saturating_add(4096);
    let limits = SourceLimits::builder()
        .max_part_bytes(maximum_part as u64)
        .expect("part limit")
        .build()
        .expect("source limits");
    let source = crate::Package::from_bytes_with_options(
        &baseline_bytes,
        crate::OpenOptions::new().source_limits(limits),
    )
    .expect("limited workbook");

    // One small gradient is cheap to serialize. Reusing the same Arc-backed
    // request across many disjoint cells keeps the generated stylesheet
    // deduplicated while the semantic transaction still charges every staged
    // cell copy against its retained-payload budget.
    let stops = vec![
        GradientStop::new(0.0, Color::rgb(0x10, 0x20, 0x30)).expect("first stop"),
        GradientStop::new(1.0, Color::rgb(0xA0, 0xB0, 0xC0)).expect("last stop"),
    ]
    .into_boxed_slice();
    let change = FillChange::replace(Fill::Gradient(
        GradientFill::linear(90.0, stops).expect("two-stop gradient"),
    ));
    let per_cell = change.retained_bytes();
    let copies = maximum_part
        .checked_div(per_cell)
        .expect("nonzero fill retention")
        / 2
        + 1;
    let per_edit = per_cell
        .checked_mul(copies)
        .expect("per-edit fill retention");
    assert!(per_edit < maximum_part);
    assert!(per_edit.saturating_mul(2) > maximum_part);

    let mut left = source.edit().expect("left edit");
    {
        let mut sheet = left
            .sheet(0usize)
            .expect("left lookup")
            .expect("left worksheet");
        for row in 1..=copies {
            let address = Address::from_a1(&format!("A{row}")).expect("left address");
            sheet.fill(address, &change).expect("left fill");
        }
    }
    let left_len = left.len();

    let mut right = source.edit().expect("right edit");
    {
        let mut sheet = right
            .sheet(0usize)
            .expect("right lookup")
            .expect("right worksheet");
        for row in 1..=copies {
            let address = Address::from_a1(&format!("B{row}")).expect("right address");
            sheet.fill(address, &change).expect("right fill");
        }
    }
    let right_len = right.len();

    let error = left
        .join(right)
        .expect_err("fill payloads exceed join budget");
    assert!(matches!(
        error.failure(),
        JoinFailure::ResourceLimit(limit)
            if limit.scope.as_ref() == "worksheet fill staged bytes"
    ));
    assert!(error.conflicts().is_none());
    assert_eq!(left.len(), left_len);
    assert_eq!(error.rejected().len(), right_len);

    let (failure, rejected) = error.into_parts();
    assert!(matches!(
        failure,
        JoinFailure::ResourceLimit(limit)
            if limit.scope.as_ref() == "worksheet fill staged bytes"
    ));
    assert_eq!(rejected.len(), right_len);

    let left_commit = left.commit().expect("left commit after rejected join");
    let right_commit = rejected
        .commit()
        .expect("rejected edit commits independently");
    assert_eq!(
        part_text(left_commit.package(), "/xl/styles.xml")
            .matches("<gradientFill")
            .count(),
        1
    );
    assert_eq!(
        part_text(right_commit.package(), "/xl/styles.xml")
            .matches("<gradientFill")
            .count(),
        1
    );
}

#[test]
fn border_change_round_trips_edges_colors_diagonals_and_replays() {
    let source = crate::Package::new().expect("source workbook");
    let source_bytes = source.to_bytes().expect("source bytes");
    let address = Address::from_a1("A1").expect("address");
    let left = Side::new(Line::Thick).with_color(Color::rgb(0x10, 0x20, 0x30));
    let right_color = Color::indexed(17);
    let top_color = Color::theme(4).with_tint(Tint::new(-0.25).expect("theme tint"));
    let bottom = Side::new(Line::Hair).with_color(Color::auto_value(false));
    let diagonal = Side::new(Line::Dashed).with_color(Color::rgb(0xA0, 0xB0, 0xC0));
    let change = BorderChange::new()
        .left(Some(left.clone()))
        .right_line(Some(Line::Medium))
        .right_color(Some(right_color))
        .top_line(Some(Line::Double))
        .top_color(Some(top_color))
        .bottom(Some(bottom.clone()))
        .diagonal(Some(diagonal.clone()))
        .diagonal_up(Some(true))
        .diagonal_down(Some(true))
        .outline(Some(false));

    let mut edit = source.edit().expect("edit");
    edit.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .border(address, &change)
        .expect("border change");
    let committed = edit.commit().expect("border commit");
    let format = committed
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let border = format.border().expect("border");
    assert_eq!(border.left.as_ref(), Some(&left));
    assert_eq!(
        border.right.as_ref().map(|side| side.line),
        Some(Line::Medium)
    );
    assert_eq!(
        border.right.as_ref().and_then(|side| side.color),
        Some(right_color)
    );
    assert_eq!(
        border.top.as_ref().map(|side| side.line),
        Some(Line::Double)
    );
    assert_eq!(
        border.top.as_ref().and_then(|side| side.color),
        Some(top_color)
    );
    assert_eq!(border.bottom.as_ref(), Some(&bottom));
    let actual_diagonal = border.diagonal.as_ref().expect("diagonal");
    assert_eq!(actual_diagonal.side(), Some(&diagonal));
    assert_eq!(actual_diagonal.dir(), Some(Dir::Both));
    assert_eq!(border.outline, Some(false));

    let replayed = source
        .apply(committed.patch())
        .expect("replay border patch");
    assert_eq!(
        replayed.package().to_bytes().expect("replayed bytes"),
        committed.package().to_bytes().expect("committed bytes")
    );
    let restored = committed
        .package()
        .apply(&committed.patch().inverse())
        .expect("inverse border patch");
    assert_eq!(
        restored.package().to_bytes().expect("restored bytes"),
        source_bytes
    );
}

#[test]
fn border_change_noop_is_exact_and_known_edits_preserve_opaque_markup() {
    let source = border_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let tint = Tint::new(0.5).expect("source tint");
    let change = BorderChange::new()
        .left(Some(
            Side::new(Line::Thin).with_color(Color::theme(4).with_tint(tint)),
        ))
        .top_line(Some(Line::Double))
        .diagonal(Some(Side::new(Line::Dashed).with_color(Color::indexed(10))))
        .diagonal_up(Some(true))
        .diagonal_down(Some(true))
        .outline(Some(true));

    let mut noop = source.edit().expect("no-op edit");
    noop.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .border("A1", &change)
        .expect("matching border");
    let noop = noop.commit().expect("no-op commit");
    assert!(noop.patch().is_noop());
    assert_eq!(
        noop.package().to_bytes().expect("no-op bytes"),
        source_bytes
    );

    let mut edit = source.edit().expect("opaque edit");
    edit.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .border("A1", &BorderChange::new().left_line(Some(Line::Thick)))
        .expect("known side change");
    let committed = edit.commit().expect("opaque border commit");
    let styles = part_text(committed.package(), "/xl/styles.xml");
    assert_eq!(styles.matches("<f:future/>").count(), 2);
    assert!(styles.contains(r#"<left style="thick">"#));
}

#[test]
fn border_change_joins_disjoint_fields_conflicts_on_overlap_and_later_whole_wins() {
    let source = crate::Package::new().expect("source workbook");
    let address = Address::from_a1("A1").expect("address");

    let mut left = source.edit().expect("left edit");
    left.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .border(address, &BorderChange::new().left_line(Some(Line::Thin)))
        .expect("left border");
    let mut right = source.edit().expect("right edit");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .border(
            address,
            &BorderChange::new()
                .top_line(Some(Line::Double))
                .outline(Some(true)),
        )
        .expect("right border");
    left.join(right).expect("disjoint border fields");
    let joined = left.commit().expect("joined border commit");
    let joined_format = joined
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let joined_border = joined_format.border().expect("joined border");
    assert_eq!(
        joined_border.left.as_ref().map(|side| side.line),
        Some(Line::Thin)
    );
    assert_eq!(
        joined_border.top.as_ref().map(|side| side.line),
        Some(Line::Double)
    );
    assert_eq!(joined_border.outline, Some(true));

    let mut left = source.edit().expect("left overlap");
    left.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .border(address, &BorderChange::new().left_line(Some(Line::Thin)))
        .expect("left line");
    let mut right = source.edit().expect("right overlap");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .border(address, &BorderChange::new().left_line(Some(Line::Thick)))
        .expect("right line");
    let error = left
        .join(right)
        .expect_err("same border line must conflict");
    let conflicts = error.conflicts().expect("border conflict details");
    assert_eq!(
        conflicts.conflicts()[0].cells().expect("cell conflicts"),
        &[address]
    );

    let mut whole = source.edit().expect("whole edit");
    whole
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .border(
            address,
            &BorderChange::new().left(Some(Side::new(Line::Thick))),
        )
        .expect("whole side");
    let mut granular = source.edit().expect("granular edit");
    granular
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .border(
            address,
            &BorderChange::new().left_color(Some(Color::rgb(0x11, 0x22, 0x33))),
        )
        .expect("granular color");
    let error = whole
        .join(granular)
        .expect_err("whole and granular side changes must conflict");
    assert!(error.conflicts().is_some());

    let expected = Side::new(Line::Double).with_color(Color::theme(6));
    let mut later = source.edit().expect("later whole edit");
    later
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .border(
            address,
            &BorderChange::new()
                .left_line(Some(Line::Thin))
                .left_color(Some(Color::rgb(0xAA, 0xBB, 0xCC))),
        )
        .and_then(|sheet| sheet.border(address, &BorderChange::new().left(Some(expected.clone()))))
        .expect("later whole side");
    let later = later.commit().expect("later whole commit");
    let later_format = later
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    assert_eq!(
        later_format.border().expect("border").left.as_ref(),
        Some(&expected)
    );
}

#[test]
fn border_change_is_discarded_by_style_replacement_and_reset() {
    let source = styled_workbook();
    let accent = source.styles().expect("styles").get(1).expect("accent");
    let address = Address::from_a1("A1").expect("address");
    let requested = BorderChange::new().left_line(Some(Line::Thick));

    let mut replacement = source.edit().expect("replacement edit");
    replacement
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .border(address, &requested)
        .and_then(|sheet| sheet.style(address, &accent))
        .expect("border then style replacement");
    let replacement = replacement.commit().expect("replacement commit");
    let replacement_format = replacement
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    assert!(replacement_format.has_default_border());

    let mut reset = source.edit().expect("reset edit");
    reset
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .border(address, &requested)
        .and_then(|sheet| sheet.reset_style(address))
        .expect("border then reset");
    let reset = reset.commit().expect("reset commit");
    let reset_sheet = reset
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet");
    assert!(matches!(
        reset_sheet.local_style(address).expect("local style"),
        Some(crate::LocalStyle::Default)
    ));
    assert!(
        reset_sheet
            .style(address)
            .expect("style lookup")
            .expect("resolved style")
            .format()
            .expect("cell format")
            .has_default_border()
    );
}

#[test]
fn failed_border_stage_is_atomic_and_recovers_after_discard() {
    let source = low_part_workbook();
    let change = BorderChange::new().left_line(Some(Line::Thin));
    let mut edit = source.edit().expect("edit");
    let mut sheet = edit.sheet(0usize).expect("lookup").expect("worksheet");
    let mut accepted = Vec::new();
    let mut failed = None;
    for row in 1..=128 {
        let address = Address::from_a1(&format!("A{row}")).expect("border address");
        match sheet.border(address, &change) {
            Ok(_) => accepted.push(address),
            Err(error) => {
                failed = Some((address, error));
                break;
            },
        }
    }
    assert!(
        accepted.len() >= 2,
        "the low budget must admit prior borders"
    );
    let (failed_address, error) = failed.expect("a later border must hit the budget");
    assert!(matches!(
        error,
        Error::ResourceLimit(limit)
            if limit.scope.as_ref() == "worksheet style staged bytes"
    ));

    // Discarding one admitted entry frees exactly one map/node charge. The
    // failed address was never inserted, so it can then be admitted and the
    // earlier, unrelated border remains staged.
    sheet
        .reset_style(accepted[0])
        .expect("discard admitted border");
    sheet
        .border(failed_address, &change)
        .expect("recover after discard");
    let committed = edit.commit().expect("recovered border commit");

    let remaining_format = committed
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(accepted[1])
        .expect("remaining style lookup")
        .expect("remaining style")
        .format()
        .expect("remaining format");
    assert_eq!(
        remaining_format
            .border()
            .and_then(|border| border.left.as_ref().map(|side| side.line)),
        Some(Line::Thin)
    );
    let recovered_format = committed
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(failed_address)
        .expect("recovered style lookup")
        .expect("recovered style")
        .format()
        .expect("recovered format");
    assert_eq!(
        recovered_format
            .border()
            .and_then(|border| border.left.as_ref().map(|side| side.line)),
        Some(Line::Thin)
    );
}

#[test]
fn border_join_resource_limit_returns_rejected_edit_and_allows_individual_commits() {
    let source = low_part_workbook();
    let change = BorderChange::new().left_line(Some(Line::Thin));

    // Probe the fixed per-cell admission charge, then use just over half of
    // the admitted capacity in each branch. Each edit fits; their disjoint
    // aggregate exceeds the same source-derived budget.
    let capacity = {
        let mut probe = source.edit().expect("probe edit");
        let mut sheet = probe.sheet(0usize).expect("lookup").expect("worksheet");
        let mut capacity = 0;
        for row in 1..=128 {
            let address = Address::from_a1(&format!("A{row}")).expect("probe address");
            if sheet.border(address, &change).is_err() {
                break;
            }
            capacity += 1;
        }
        capacity
    };
    assert!(capacity >= 2, "the low budget must admit multiple borders");
    let copies = capacity / 2 + 1;

    let mut left = source.edit().expect("left edit");
    {
        let mut sheet = left
            .sheet(0usize)
            .expect("left lookup")
            .expect("left worksheet");
        for row in 1..=copies {
            let address = Address::from_a1(&format!("A{row}")).expect("left address");
            sheet.border(address, &change).expect("left border");
        }
    }
    let left_len = left.len();

    let mut right = source.edit().expect("right edit");
    {
        let mut sheet = right
            .sheet(0usize)
            .expect("right lookup")
            .expect("right worksheet");
        for row in 1..=copies {
            let address = Address::from_a1(&format!("B{row}")).expect("right address");
            sheet.border(address, &change).expect("right border");
        }
    }
    let right_len = right.len();

    let error = left
        .join(right)
        .expect_err("aggregate borders exceed budget");
    assert!(matches!(
        error.failure(),
        JoinFailure::ResourceLimit(limit)
            if limit.scope.as_ref() == "worksheet border staged bytes"
    ));
    assert!(error.conflicts().is_none());
    assert_eq!(left.len(), left_len);
    assert_eq!(error.rejected().len(), right_len);

    let (failure, rejected) = error.into_parts();
    assert!(matches!(
        failure,
        JoinFailure::ResourceLimit(limit)
            if limit.scope.as_ref() == "worksheet border staged bytes"
    ));
    assert_eq!(rejected.len(), right_len);

    let left_commit = left.commit().expect("left independent commit");
    let right_commit = rejected.commit().expect("right independent commit");
    for (workbook, address) in [
        (
            left_commit.package(),
            Address::from_a1("A1").expect("left A1"),
        ),
        (
            right_commit.package(),
            Address::from_a1("B1").expect("right B1"),
        ),
    ] {
        let format = workbook
            .sheet(0usize)
            .expect("lookup")
            .expect("worksheet")
            .style(address)
            .expect("style lookup")
            .expect("styled cell")
            .format()
            .expect("cell format");
        assert_eq!(
            format
                .border()
                .and_then(|border| border.left.as_ref().map(|side| side.line)),
            Some(Line::Thin)
        );
    }
}

#[test]
fn font_change_repeated_values_and_unset_values_are_exact_no_ops() {
    let source = crate::Package::new().expect("source workbook");
    let address = Address::from_a1("A1").expect("address");
    let mut create = source.edit().expect("create edit");
    let mut sheet = create.sheet(0usize).expect("lookup").expect("worksheet");
    sheet.set(address, "styled").expect("cell value");
    sheet
        .font(
            address,
            &FontChange::new()
                .strike(Some(true))
                .underline(Some(Underline::Double)),
        )
        .expect("font properties");
    let configured = create.commit().expect("configured commit").into_package();
    let configured_bytes = configured.to_bytes().expect("configured bytes");

    let mut repeat = configured.edit().expect("repeat edit");
    repeat
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(
            address,
            &FontChange::new()
                .strike(Some(true))
                .underline(Some(Underline::Double)),
        )
        .expect("repeat properties");
    let repeat = repeat.commit().expect("repeat commit");
    assert!(repeat.patch().is_noop());
    assert_eq!(
        repeat.package().to_bytes().expect("repeat bytes"),
        configured_bytes
    );

    let mut unset = configured.edit().expect("unset edit");
    unset
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(address, &FontChange::new().strike(None).underline(None))
        .expect("unset properties");
    let unset = unset.commit().expect("unset commit");
    assert_eq!(unset.patch().operation_count(), 1);
    assert_eq!(
        unset
            .package()
            .sheet(0usize)
            .expect("lookup")
            .expect("worksheet")
            .style(address)
            .expect("style lookup")
            .expect("styled cell")
            .format()
            .expect("cell format")
            .font()
            .and_then(|font| font.underline()),
        None
    );
}

#[test]
fn font_change_name_size_script_round_trip_replays_and_owns_input() {
    let source = crate::Package::new().expect("source workbook");
    let source_bytes = source.to_bytes().expect("source bytes");
    let address = Address::from_a1("A1").expect("address");
    let mut authored_name = String::from("A&B <quoted> \"family\"\r\n\t字体");
    let expected_name = authored_name.clone();
    let change = FontChange::new()
        .name(Some(authored_name.as_str()))
        .expect("font name")
        .size(Some(14.5))
        .expect("font size")
        .script(Some(Script::Superscript));
    authored_name.clear();
    drop(authored_name);

    let mut edit = source.edit().expect("edit");
    let mut sheet = edit.sheet(0usize).expect("lookup").expect("worksheet");
    sheet.set(address, "styled").expect("cell value");
    sheet.font(address, &change).expect("font properties");
    let committed = edit.commit().expect("font commit");
    assert_eq!(committed.patch().operation_count(), 1);

    let committed_bytes = committed.package().to_bytes().expect("committed bytes");
    let styles = part_text(committed.package(), "/xl/styles.xml");
    assert!(styles.contains("A&amp;B &lt;quoted"));
    assert!(styles.contains("&quot;family&quot;"));
    assert!(styles.contains(r#"<sz val="14.5"/>"#));
    assert!(styles.contains(r#"<vertAlign val="superscript"/>"#));
    let style = committed
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell");
    let format = style.format().expect("cell format");
    let font = format.font().expect("font");
    assert_eq!(font.name(), Some(expected_name.as_str()));
    assert_eq!(font.size(), Some(14.5));
    assert_eq!(font.script(), Some(Script::Superscript));

    let reopened = crate::Package::from_bytes(&committed_bytes).expect("reopened workbook");
    assert_eq!(part_text(&reopened, "/xl/styles.xml"), styles);
    let reopened_format = reopened
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let reopened_font = reopened_format.font().expect("font");
    assert_eq!(reopened_font.name(), Some(expected_name.as_str()));
    assert_eq!(reopened_font.size(), Some(14.5));
    assert_eq!(reopened_font.script(), Some(Script::Superscript));

    let replayed = source.apply(committed.patch()).expect("replay patch");
    assert_eq!(
        replayed.package().to_bytes().expect("replayed bytes"),
        committed_bytes
    );
    let restored = committed
        .package()
        .apply(&committed.patch().inverse())
        .expect("inverse patch");
    assert_eq!(
        restored.package().to_bytes().expect("restored bytes"),
        source_bytes
    );
}

#[test]
fn font_change_name_size_script_joins_disjoint_fields_and_conflicts_on_script() {
    let source = crate::Package::new().expect("source workbook");
    let address = Address::from_a1("A1").expect("address");

    let mut name = source.edit().expect("name edit");
    name.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(
            address,
            &FontChange::new().name(Some("Joined")).expect("font name"),
        )
        .expect("name");
    let mut size = source.edit().expect("size edit");
    size.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(
            address,
            &FontChange::new().size(Some(9.5)).expect("font size"),
        )
        .expect("size");
    name.join(size).expect("disjoint name and size");
    let mut script = source.edit().expect("script edit");
    script
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(address, &FontChange::new().script(Some(Script::Subscript)))
        .expect("script");
    name.join(script).expect("disjoint script");
    let joined = name.commit().expect("joined font commit");
    let format = joined
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let font = format.font().expect("font");
    assert_eq!(font.name(), Some("Joined"));
    assert_eq!(font.size(), Some(9.5));
    assert_eq!(font.script(), Some(Script::Subscript));

    let mut left = source.edit().expect("left edit");
    left.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(
            address,
            &FontChange::new().script(Some(Script::Superscript)),
        )
        .expect("left script");
    let mut right = source.edit().expect("right edit");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(address, &FontChange::new().script(Some(Script::Subscript)))
        .expect("right script");
    let error = left
        .join(right)
        .expect_err("same script property must conflict");
    let conflicts = error.conflicts().expect("cell conflict details");
    assert_eq!(conflicts.conflicts().len(), 1);
    assert_eq!(
        conflicts.conflicts()[0].cells().expect("cell conflicts"),
        &[address]
    );
}

#[test]
fn font_change_name_size_script_repeats_and_unsets_without_spurious_changes() {
    let source = crate::Package::new().expect("source workbook");
    let address = Address::from_a1("A1").expect("address");
    let mut create = source.edit().expect("create edit");
    let mut sheet = create.sheet(0usize).expect("lookup").expect("worksheet");
    sheet.set(address, "styled").expect("cell value");
    sheet
        .font(
            address,
            &FontChange::new()
                .name(Some("Repeat"))
                .expect("font name")
                .size(Some(12.25))
                .expect("font size")
                .script(Some(Script::Superscript)),
        )
        .expect("font properties");
    let configured = create.commit().expect("configured commit").into_package();
    let configured_bytes = configured.to_bytes().expect("configured bytes");

    let mut repeat = configured.edit().expect("repeat edit");
    repeat
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(
            address,
            &FontChange::new()
                .name(Some("Repeat"))
                .expect("font name")
                .size(Some(12.25))
                .expect("font size")
                .script(Some(Script::Superscript)),
        )
        .expect("repeat properties");
    let repeat = repeat.commit().expect("repeat commit");
    assert!(repeat.patch().is_noop());
    assert_eq!(
        repeat.package().to_bytes().expect("repeat bytes"),
        configured_bytes
    );

    let mut unset = configured.edit().expect("unset edit");
    unset
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(
            address,
            &FontChange::new()
                .name(None)
                .expect("unset font name")
                .size(None)
                .expect("unset font size")
                .script(None),
        )
        .expect("unset properties");
    let unset = unset.commit().expect("unset commit");
    assert_eq!(unset.patch().operation_count(), 1);
    let format = unset
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let font = format.font().expect("font");
    assert_eq!(font.name(), None);
    assert_eq!(font.size(), None);
    assert_eq!(font.script(), None);
}

#[test]
fn font_change_color_replaces_rgb_with_theme_tint_and_preserves_companions() {
    let source = crate::Package::new().expect("source workbook");
    let source_bytes = source.to_bytes().expect("source bytes");
    let address = Address::from_a1("A1").expect("address");
    let rgb = Rgb::argb(0x80, 0x11, 0x22, 0x33);
    let rgb_color = FontColor::rgb(rgb);
    assert_eq!(rgb_color.kind(), FontColorKind::Rgb(rgb));

    let mut edit = source.edit().expect("edit");
    let mut sheet = edit.sheet(0usize).expect("lookup").expect("worksheet");
    sheet.set(address, "styled").expect("cell value");
    sheet
        .font(
            address,
            &FontChange::new()
                .color(Some(rgb_color))
                .name(Some("Companion"))
                .expect("font name")
                .script(Some(Script::Superscript)),
        )
        .expect("font properties");
    let rgb_commit = edit.commit().expect("RGB commit");
    let rgb_bytes = rgb_commit.package().to_bytes().expect("RGB bytes");
    let rgb_styles = part_text(rgb_commit.package(), "/xl/styles.xml");
    assert!(rgb_styles.contains(r#"<color rgb="80112233"/>"#));
    let rgb_format = rgb_commit
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let rgb_font = rgb_format.font().expect("font");
    assert_eq!(rgb_font.color(), Some(Color::from_rgb(rgb)));
    assert_eq!(rgb_font.name(), Some("Companion"));
    assert_eq!(rgb_font.script(), Some(Script::Superscript));

    let replayed = source.apply(rgb_commit.patch()).expect("replay RGB patch");
    assert_eq!(
        replayed.package().to_bytes().expect("replayed RGB bytes"),
        rgb_bytes
    );
    let restored = rgb_commit
        .package()
        .apply(&rgb_commit.patch().inverse())
        .expect("inverse RGB patch");
    assert_eq!(
        restored
            .package()
            .to_bytes()
            .expect("restored source bytes"),
        source_bytes
    );

    let tint = Tint::new(0.25).expect("tint");
    let theme_color = FontColor::theme(5).with_tint(tint);
    assert_eq!(theme_color.kind(), FontColorKind::Theme(5));
    assert_eq!(theme_color.tint(), Some(tint));
    let mut replace = rgb_commit.package().edit().expect("replace edit");
    replace
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(address, &FontChange::new().color(Some(theme_color)))
        .expect("theme color");
    let theme_commit = replace.commit().expect("theme commit");
    let theme_bytes = theme_commit.package().to_bytes().expect("theme bytes");
    let theme_styles = part_text(theme_commit.package(), "/xl/styles.xml");
    assert!(theme_styles.contains(r#"<color theme="5" tint="0.25"/>"#));
    let theme_format = theme_commit
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let theme_font = theme_format.font().expect("font");
    assert_eq!(theme_font.color(), Some(Color::theme(5).with_tint(tint)));
    assert_eq!(theme_font.name(), Some("Companion"));
    assert_eq!(theme_font.script(), Some(Script::Superscript));

    let reopened = crate::Package::from_bytes(&theme_bytes).expect("reopened theme workbook");
    let reopened_format = reopened
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let reopened_font = reopened_format.font().expect("font");
    assert_eq!(reopened_font.color(), Some(Color::theme(5).with_tint(tint)));
    assert_eq!(reopened_font.name(), Some("Companion"));
    assert_eq!(reopened_font.script(), Some(Script::Superscript));

    let replayed_theme = rgb_commit
        .package()
        .apply(theme_commit.patch())
        .expect("replay theme patch");
    assert_eq!(
        replayed_theme
            .package()
            .to_bytes()
            .expect("replayed theme bytes"),
        theme_bytes
    );
    let restored_rgb = theme_commit
        .package()
        .apply(&theme_commit.patch().inverse())
        .expect("inverse theme patch");
    assert_eq!(
        restored_rgb
            .package()
            .to_bytes()
            .expect("restored RGB bytes"),
        rgb_bytes
    );
}

#[test]
fn font_change_color_semantics_ignore_source_lexical_spelling_and_deduplicate() {
    let baseline = styled_workbook();
    let styles_uri = PackURI::new("/xl/styles.xml").expect("styles URI");
    let mut package = baseline.inner.package.clone();
    let styles = std::str::from_utf8(package.get_part(&styles_uri).expect("styles part").blob())
        .expect("styles XML")
        .replace(r#"<font/>"#, r#"<font><color rgb="ff112233"/></font>"#);
    package
        .get_part_mut(&styles_uri)
        .expect("styles part")
        .set_blob(styles.into_bytes());
    let lexical_source = crate::Package::from_opc(package).expect("lexical source workbook");
    let before_bytes = lexical_source.to_bytes().expect("source bytes");
    let before_styles = part_text(&lexical_source, "/xl/styles.xml").to_owned();
    let semantic_color = FontColor::rgb(Rgb::argb(0xFF, 0x11, 0x22, 0x33));

    let mut repeat = lexical_source.edit().expect("repeat edit");
    repeat
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(
            Address::from_a1("A1").expect("address"),
            &FontChange::new().color(Some(semantic_color.clone())),
        )
        .expect("repeat color");
    let repeated = repeat.commit().expect("repeat commit");
    assert!(repeated.patch().is_noop());
    assert_eq!(
        repeated.package().to_bytes().expect("repeat bytes"),
        before_bytes
    );
    assert_eq!(
        part_text(repeated.package(), "/xl/styles.xml"),
        before_styles
    );

    let source = crate::Package::new().expect("dedup source");
    let before = part_text(&source, "/xl/styles.xml").to_owned();
    let mut edit = source.edit().expect("dedup edit");
    let mut sheet = edit.sheet(0usize).expect("lookup").expect("worksheet");
    for (address, value) in [("A1", "one"), ("B1", "two"), ("C1", "three")] {
        sheet.set(address, value).expect("cell value");
        sheet
            .font(
                address,
                &FontChange::new().color(Some(semantic_color.clone())),
            )
            .expect("shared color");
    }
    let deduplicated = edit.commit().expect("dedup commit").into_package();
    let after = part_text(&deduplicated, "/xl/styles.xml");
    assert_eq!(
        after.matches("<color").count(),
        before.matches("<color").count() + 1
    );
    assert_eq!(
        after.matches("<xf").count(),
        before.matches("<xf").count() + 1
    );
    for address in ["A1", "B1", "C1"] {
        let format = deduplicated
            .sheet(0usize)
            .expect("lookup")
            .expect("worksheet")
            .style(address)
            .expect("style lookup")
            .expect("styled cell")
            .format()
            .expect("cell format");
        assert_eq!(
            format.font().expect("font").color(),
            Some(Color::from_rgb(Rgb::argb(0xFF, 0x11, 0x22, 0x33)))
        );
    }

    let mut null = source.edit().expect("null edit");
    null.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(
            Address::from_a1("A1").expect("address"),
            &FontChange::new().color(None),
        )
        .expect("null color");
    let null = null.commit().expect("null commit");
    assert!(null.patch().is_noop());
    assert_eq!(
        null.package().to_bytes().expect("null bytes"),
        source.to_bytes().expect("source bytes")
    );
}

#[test]
fn font_change_color_joins_disjoint_fields_and_conflicts_on_color() {
    let source = crate::Package::new().expect("source workbook");
    let address = Address::from_a1("A1").expect("address");
    let rgb = Rgb::new(0x20, 0x40, 0x60);

    let mut color = source.edit().expect("color edit");
    color
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(address, &FontChange::new().color(Some(FontColor::rgb(rgb))))
        .expect("color");
    let mut name = source.edit().expect("name edit");
    name.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(
            address,
            &FontChange::new().name(Some("Joined")).expect("font name"),
        )
        .expect("name");
    color.join(name).expect("disjoint color and name");
    let joined = color.commit().expect("joined commit");
    let format = joined
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let font = format.font().expect("font");
    assert_eq!(font.color(), Some(Color::from_rgb(rgb)));
    assert_eq!(font.name(), Some("Joined"));

    let mut left = source.edit().expect("left edit");
    left.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(
            address,
            &FontChange::new().color(Some(FontColor::rgb(Rgb::new(1, 2, 3)))),
        )
        .expect("left color");
    let mut right = source.edit().expect("right edit");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(
            address,
            &FontChange::new().color(Some(FontColor::rgb(Rgb::new(4, 5, 6)))),
        )
        .expect("right color");
    let error = left
        .join(right)
        .expect_err("same color property must conflict");
    let conflicts = error.conflicts().expect("cell conflict details");
    assert_eq!(conflicts.conflicts().len(), 1);
    assert_eq!(
        conflicts.conflicts()[0].cells().expect("cell conflicts"),
        &[address]
    );
}

#[test]
fn font_change_native_bool_integer_and_scheme_fields_round_trip_with_companions() {
    let source = crate::Package::new().expect("source workbook");
    let source_bytes = source.to_bytes().expect("source bytes");
    let address = Address::from_a1("A1").expect("address");
    let change = FontChange::new()
        .name(Some("Companion"))
        .expect("font name")
        .size(Some(12.0))
        .expect("font size")
        .script(Some(Script::Superscript))
        .outline(Some(false))
        .shadow(Some(true))
        .condense(Some(false))
        .extend(Some(true))
        .charset(Some(0))
        .expect("font charset")
        .family(Some(0))
        .expect("font family")
        .scheme(Some(Scheme::None));

    let mut edit = source.edit().expect("edit");
    let mut sheet = edit.sheet(0usize).expect("lookup").expect("worksheet");
    sheet.set(address, "styled").expect("cell value");
    sheet.font(address, &change).expect("font properties");
    let committed = edit.commit().expect("font commit");
    assert_eq!(committed.patch().operation_count(), 1);

    let committed_bytes = committed.package().to_bytes().expect("committed bytes");
    let styles = part_text(committed.package(), "/xl/styles.xml");
    assert!(styles.contains(r#"<outline val="0"/>"#));
    assert!(styles.contains(r#"<shadow/>"#));
    assert!(styles.contains(r#"<condense val="0"/>"#));
    assert!(styles.contains(r#"<extend/>"#));
    assert!(styles.contains(r#"<charset val="0"/>"#));
    assert!(styles.contains(r#"<family val="0"/>"#));
    assert!(styles.contains(r#"<scheme val="none"/>"#));
    let format = committed
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let font = format.font().expect("font");
    assert_eq!(font.outline(), Some(false));
    assert_eq!(font.shadow(), Some(true));
    assert_eq!(font.condense(), Some(false));
    assert_eq!(font.extend(), Some(true));
    assert_eq!(font.charset(), Some(0));
    assert_eq!(font.family(), Some(0));
    assert_eq!(font.scheme(), Some(Scheme::None));
    assert_eq!(font.name(), Some("Companion"));
    assert_eq!(font.size(), Some(12.0));
    assert_eq!(font.script(), Some(Script::Superscript));

    let reopened = crate::Package::from_bytes(&committed_bytes).expect("reopened workbook");
    let reopened_format = reopened
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let reopened_font = reopened_format.font().expect("font");
    assert_eq!(reopened_font.outline(), Some(false));
    assert_eq!(reopened_font.shadow(), Some(true));
    assert_eq!(reopened_font.condense(), Some(false));
    assert_eq!(reopened_font.extend(), Some(true));
    assert_eq!(reopened_font.charset(), Some(0));
    assert_eq!(reopened_font.family(), Some(0));
    assert_eq!(reopened_font.scheme(), Some(Scheme::None));
    assert_eq!(reopened_font.name(), Some("Companion"));
    assert_eq!(reopened_font.size(), Some(12.0));
    assert_eq!(reopened_font.script(), Some(Script::Superscript));

    let replayed = source.apply(committed.patch()).expect("replay patch");
    assert_eq!(
        replayed.package().to_bytes().expect("replayed bytes"),
        committed_bytes
    );
    let restored = committed
        .package()
        .apply(&committed.patch().inverse())
        .expect("inverse patch");
    assert_eq!(
        restored.package().to_bytes().expect("restored bytes"),
        source_bytes
    );
}

#[test]
fn font_change_native_fields_join_and_distinguish_false_zero_none_and_unset() {
    let source = crate::Package::new().expect("source workbook");
    let address = Address::from_a1("A1").expect("address");

    let mut outline = source.edit().expect("outline edit");
    outline
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(address, &FontChange::new().outline(Some(false)))
        .expect("outline");
    let mut family = source.edit().expect("family edit");
    family
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(
            address,
            &FontChange::new().family(Some(0)).expect("font family"),
        )
        .expect("family");
    outline.join(family).expect("disjoint native fields");
    let joined = outline.commit().expect("joined commit");
    let joined_format = joined
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let joined_font = joined_format.font().expect("font");
    assert_eq!(joined_font.outline(), Some(false));
    assert_eq!(joined_font.family(), Some(0));

    let mut left = source.edit().expect("left edit");
    left.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(address, &FontChange::new().outline(Some(false)))
        .expect("left outline");
    let mut right = source.edit().expect("right edit");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(address, &FontChange::new().outline(Some(true)))
        .expect("right outline");
    let error = left
        .join(right)
        .expect_err("same native field must conflict");
    let conflicts = error.conflicts().expect("cell conflict details");
    assert_eq!(conflicts.conflicts().len(), 1);
    assert_eq!(
        conflicts.conflicts()[0].cells().expect("cell conflicts"),
        &[address]
    );

    let mut create = source.edit().expect("create edit");
    create
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .set(address, "styled")
        .expect("cell value")
        .font(
            address,
            &FontChange::new()
                .outline(Some(false))
                .shadow(Some(true))
                .condense(Some(false))
                .extend(Some(true))
                .charset(Some(0))
                .expect("font charset")
                .family(Some(0))
                .expect("font family")
                .scheme(Some(Scheme::None)),
        )
        .expect("native fields");
    let configured = create.commit().expect("configured commit").into_package();
    let configured_bytes = configured.to_bytes().expect("configured bytes");

    let mut repeat = configured.edit().expect("repeat edit");
    repeat
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(
            address,
            &FontChange::new()
                .outline(Some(false))
                .shadow(Some(true))
                .condense(Some(false))
                .extend(Some(true))
                .charset(Some(0))
                .expect("font charset")
                .family(Some(0))
                .expect("font family")
                .scheme(Some(Scheme::None)),
        )
        .expect("repeat native fields");
    let repeat = repeat.commit().expect("repeat commit");
    assert!(repeat.patch().is_noop());
    assert_eq!(
        repeat.package().to_bytes().expect("repeat bytes"),
        configured_bytes
    );

    let mut unset = configured.edit().expect("unset edit");
    unset
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(
            address,
            &FontChange::new()
                .outline(None)
                .shadow(None)
                .condense(None)
                .extend(None)
                .charset(None)
                .expect("unset font charset")
                .family(None)
                .expect("unset font family")
                .scheme(None),
        )
        .expect("unset native fields");
    let unset = unset.commit().expect("unset commit");
    assert_eq!(unset.patch().operation_count(), 1);
    let unset_format = unset
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let unset_font = unset_format.font().expect("font");
    assert_eq!(unset_font.outline(), None);
    assert_eq!(unset_font.shadow(), None);
    assert_eq!(unset_font.condense(), None);
    assert_eq!(unset_font.extend(), None);
    assert_eq!(unset_font.charset(), None);
    assert_eq!(unset_font.family(), None);
    assert_eq!(unset_font.scheme(), None);
}

#[test]
fn alignment_change_all_fields_round_trip_with_font_and_payload() {
    let source = crate::Package::new().expect("source workbook");
    let source_bytes = source.to_bytes().expect("source bytes");
    let address = Address::from_a1("A1").expect("address");
    let rotation = Rotation::degrees(45).expect("rotation");
    let change = AlignmentChange::new()
        .horizontal(Some(Horizontal::Center))
        .vertical(Some(Vertical::Bottom))
        .text_rotation(Some(rotation))
        .wrap_text(Some(true))
        .indent(Some(Indent::new(2)))
        .relative_indent(Some(-2))
        .justify_last_line(Some(false))
        .shrink_to_fit(Some(true))
        .reading_order(Some(Reading::RightToLeft));
    let protection = ProtectionChange::new()
        .locked(Some(false))
        .hidden(Some(true));

    let mut edit = source.edit().expect("edit");
    let mut sheet = edit.sheet(0usize).expect("lookup").expect("worksheet");
    sheet.set(address, "aligned").expect("cell value");
    sheet
        .font(address, &FontChange::new().bold(Some(true)))
        .expect("font companion");
    sheet.alignment(address, &change).expect("alignment");
    sheet
        .protection(address, &protection)
        .expect("protection companion");
    sheet
        .quote_prefix(address, Some(true))
        .expect("quote-prefix companion");
    let committed = edit.commit().expect("alignment commit");
    assert_eq!(committed.patch().operation_count(), 1);

    let committed_bytes = committed.package().to_bytes().expect("committed bytes");
    let styles = part_text(committed.package(), "/xl/styles.xml");
    assert!(styles.contains(r#"horizontal="center""#));
    assert!(styles.contains(r#"vertical="bottom""#));
    assert!(styles.contains(r#"textRotation="45""#));
    assert!(styles.contains(r#"wrapText="1""#));
    assert!(styles.contains(r#"indent="2""#));
    assert!(styles.contains(r#"relativeIndent="-2""#));
    assert!(styles.contains(r#"justifyLastLine="0""#));
    assert!(styles.contains(r#"shrinkToFit="1""#));
    assert!(styles.contains(r#"readingOrder="2""#));
    assert!(styles.contains(r#"locked="0""#));
    assert!(styles.contains(r#"hidden="1""#));
    assert!(styles.contains(r#"quotePrefix="1""#));

    let format = committed
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let alignment = format.alignment().expect("alignment");
    assert_eq!(alignment.horizontal, Some(Horizontal::Center));
    assert_eq!(alignment.vertical, Some(Vertical::Bottom));
    assert_eq!(alignment.text_rotation, Some(rotation));
    assert!(alignment.wrap_text);
    assert_eq!(alignment.indent, Some(Indent::new(2)));
    assert_eq!(alignment.relative_indent, Some(-2));
    assert_eq!(alignment.justify_last_line, Some(false));
    assert!(alignment.shrink_to_fit);
    assert_eq!(alignment.reading_order, Some(Reading::RightToLeft));
    assert_eq!(format.font().and_then(|font| font.bold()), Some(true));
    let protection_view = format.protection().expect("protection");
    assert_eq!(protection_view.locked(), Some(false));
    assert_eq!(protection_view.hidden(), Some(true));
    assert!(format.quote_prefix());

    let reopened = crate::Package::from_bytes(&committed_bytes).expect("reopened workbook");
    let reopened_format = reopened
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    assert_eq!(reopened_format.alignment(), Some(alignment));
    assert_eq!(
        reopened_format.font().and_then(|font| font.bold()),
        Some(true)
    );
    let reopened_protection = reopened_format.protection().expect("protection");
    assert_eq!(reopened_protection.locked(), Some(false));
    assert_eq!(reopened_protection.hidden(), Some(true));
    assert!(reopened_format.quote_prefix());

    let replayed = source.apply(committed.patch()).expect("replay patch");
    assert_eq!(
        replayed.package().to_bytes().expect("replayed bytes"),
        committed_bytes
    );
    let restored = committed
        .package()
        .apply(&committed.patch().inverse())
        .expect("inverse patch");
    assert_eq!(
        restored.package().to_bytes().expect("restored bytes"),
        source_bytes
    );
}

#[test]
fn alignment_change_preserves_source_shape_and_reopens() {
    let baseline = styled_workbook();
    let mut package = baseline.inner.package.clone();
    package
        .get_part_mut(&PackURI::new("/xl/styles.xml").expect("styles URI"))
        .expect("styles part")
        .set_blob(
            br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:f="urn:future"><fonts count="1"><font/></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment f:keep="yes" horizontal='center' wrapText='0'><!-- preserved alignment comment --><f:future/></alignment><protection f:keep="yes" locked='1'><!-- preserved protection comment --><f:future/></protection></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#
                .to_vec(),
        );
    package
        .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
        .expect("worksheet part")
        .set_blob(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>"#
                .to_vec(),
        );
    let source = crate::Package::from_opc(package).expect("source workbook");
    let source_styles = part_text(&source, "/xl/styles.xml").to_owned();
    assert!(source_styles.contains("<!-- preserved alignment comment -->"));
    assert!(source_styles.contains("<!-- preserved protection comment -->"));
    assert!(source_styles.contains("horizontal='center'"));
    assert_eq!(source_styles.matches(r#"f:keep="yes""#).count(), 2);
    assert_eq!(source_styles.matches("<f:future/>").count(), 2);

    let address = Address::from_a1("A1").expect("address");
    let mut edit = source.edit().expect("edit");
    edit.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .alignment(
            address,
            &AlignmentChange::new().vertical(Some(Vertical::Bottom)),
        )
        .and_then(|sheet| {
            sheet.protection(
                address,
                &ProtectionChange::new()
                    .locked(Some(false))
                    .hidden(Some(true)),
            )
        })
        .and_then(|sheet| sheet.quote_prefix(address, Some(true)))
        .expect("alignment");
    let committed = edit.commit().expect("alignment commit");
    let styles = part_text(committed.package(), "/xl/styles.xml");
    assert!(styles.contains("<!-- preserved alignment comment -->"));
    assert!(styles.contains("<!-- preserved protection comment -->"));
    assert!(styles.contains("horizontal='center'"));
    assert_eq!(
        styles.matches(r#"f:keep="yes""#).count(),
        source_styles.matches(r#"f:keep="yes""#).count() * 2
    );
    assert_eq!(
        styles.matches("<f:future/>").count(),
        source_styles.matches("<f:future/>").count() * 2
    );
    assert!(styles.contains(r#"vertical="bottom""#));
    assert!(styles.contains("locked='0'"));
    assert!(styles.contains("hidden=\"1\""));
    assert!(styles.contains(r#"quotePrefix="1""#));

    let reopened = crate::Package::from_bytes(&committed.package().to_bytes().expect("bytes"))
        .expect("reopened workbook");
    let format = reopened
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    assert_eq!(
        format
            .alignment()
            .and_then(|alignment| alignment.horizontal),
        Some(Horizontal::Center)
    );
    assert_eq!(
        format.alignment().and_then(|alignment| alignment.vertical),
        Some(Vertical::Bottom)
    );
    let protection = format.protection().expect("protection");
    assert_eq!(protection.locked(), Some(false));
    assert_eq!(protection.hidden(), Some(true));
    assert!(format.quote_prefix());
}

#[test]
fn alignment_change_repeats_and_distinguishes_explicit_defaults_from_unset_and_clear() {
    let source = crate::Package::new().expect("source workbook");
    let address = Address::from_a1("A1").expect("address");

    let mut unset = source.edit().expect("unset edit");
    unset
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .alignment(address, &AlignmentChange::new())
        .expect("empty alignment");
    let unset = unset.commit().expect("unset commit");
    assert!(unset.patch().is_noop());

    let mut explicit = source.edit().expect("explicit edit");
    explicit
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .alignment(address, &AlignmentChange::new().wrap_text(Some(false)))
        .expect("explicit default");
    let explicit = explicit.commit().expect("explicit commit");
    assert_eq!(explicit.patch().operation_count(), 1);
    let explicit_styles = part_text(explicit.package(), "/xl/styles.xml");
    assert!(explicit_styles.contains(r#"wrapText="0""#));
    let explicit_format = explicit
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let explicit_alignment = explicit_format.alignment().expect("explicit alignment");
    assert!(!explicit_alignment.wrap_text);

    let explicit_bytes = explicit.package().to_bytes().expect("explicit bytes");
    let mut repeat = explicit.package().edit().expect("repeat edit");
    repeat
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .alignment(address, &AlignmentChange::new().wrap_text(Some(false)))
        .expect("repeat explicit default");
    let repeat = repeat.commit().expect("repeat commit");
    assert!(repeat.patch().is_noop());
    assert_eq!(
        repeat.package().to_bytes().expect("repeat bytes"),
        explicit_bytes
    );

    let mut clear = explicit.package().edit().expect("clear edit");
    clear
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .alignment(address, &AlignmentChange::new().wrap_text(None))
        .expect("clear explicit default");
    let clear = clear.commit().expect("clear commit");
    assert_eq!(clear.patch().operation_count(), 1);
    let clear_styles = part_text(clear.package(), "/xl/styles.xml");
    assert!(clear_styles.contains("<alignment/>") || clear_styles.contains("<alignment />"));
    let clear_format = clear
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    assert_eq!(
        clear_format.alignment(),
        Some(&crate::style::stylesheet::Alignment::new())
    );
}

#[test]
fn alignment_change_joins_disjoint_fields_and_font_but_conflicts_on_overlap() {
    let source = crate::Package::new().expect("source workbook");
    let address = Address::from_a1("A1").expect("address");

    let mut horizontal = source.edit().expect("horizontal edit");
    horizontal
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .alignment(
            address,
            &AlignmentChange::new().horizontal(Some(Horizontal::Center)),
        )
        .expect("horizontal");
    let mut vertical = source.edit().expect("vertical edit");
    vertical
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .alignment(
            address,
            &AlignmentChange::new().vertical(Some(Vertical::Bottom)),
        )
        .expect("vertical");
    horizontal
        .join(vertical)
        .expect("disjoint alignment fields");
    let joined = horizontal.commit().expect("joined alignment commit");
    let joined_format = joined
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let joined_alignment = joined_format.alignment().expect("joined alignment");
    assert_eq!(joined_alignment.horizontal, Some(Horizontal::Center));
    assert_eq!(joined_alignment.vertical, Some(Vertical::Bottom));

    let mut font = source.edit().expect("font edit");
    font.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .font(address, &FontChange::new().bold(Some(true)))
        .expect("font");
    let mut alignment = source.edit().expect("alignment edit");
    alignment
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .alignment(address, &AlignmentChange::new().wrap_text(Some(true)))
        .expect("alignment");
    font.join(alignment).expect("font and alignment join");
    let joined = font.commit().expect("font/alignment commit");
    let joined_format = joined
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    assert_eq!(
        joined_format.font().and_then(|font| font.bold()),
        Some(true)
    );
    assert!(joined_format.alignment().expect("alignment").wrap_text);

    let mut left = source.edit().expect("left edit");
    left.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .alignment(
            address,
            &AlignmentChange::new().horizontal(Some(Horizontal::Center)),
        )
        .expect("left horizontal");
    let mut right = source.edit().expect("right edit");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .alignment(
            address,
            &AlignmentChange::new().horizontal(Some(Horizontal::Right)),
        )
        .expect("right horizontal");
    let error = left
        .join(right)
        .expect_err("same alignment field must conflict");
    let conflicts = error.conflicts().expect("cell conflict details");
    assert_eq!(conflicts.conflicts().len(), 1);
    assert_eq!(
        conflicts.conflicts()[0].cells().expect("cell conflicts"),
        &[address]
    );
}

#[test]
fn alignment_staging_is_discarded_by_style_replacement_reset_and_removal() {
    let source = styled_workbook();
    let source_styles = part_text(&source, "/xl/styles.xml").to_owned();
    let accent = source
        .styles()
        .expect("styles")
        .get(1)
        .expect("accent style");
    let replacement_address = Address::from_a1("B1").expect("replacement address");
    let styled_address = Address::from_a1("A1").expect("styled address");
    let change = AlignmentChange::new().horizontal(Some(Horizontal::Center));
    let protection = ProtectionChange::new()
        .locked(Some(false))
        .hidden(Some(true));

    let mut replacement = source.edit().expect("replacement edit");
    replacement
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .alignment(replacement_address, &change)
        .and_then(|sheet| sheet.protection(replacement_address, &protection))
        .and_then(|sheet| sheet.quote_prefix(replacement_address, Some(true)))
        .and_then(|sheet| sheet.style(replacement_address, &accent))
        .expect("alignment then style replacement");
    let replacement = replacement.commit().expect("replacement commit");
    let replacement_sheet = replacement
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet");
    assert!(
        replacement_sheet
            .style(replacement_address)
            .expect("style lookup")
            .is_some_and(|style| style.same(&accent))
    );
    assert_eq!(
        replacement_sheet
            .style(replacement_address)
            .expect("style lookup")
            .expect("resolved style")
            .format()
            .expect("cell format")
            .alignment(),
        None
    );
    let replacement_format = replacement_sheet
        .style(replacement_address)
        .expect("style lookup")
        .expect("resolved style")
        .format()
        .expect("cell format");
    assert_eq!(replacement_format.protection(), None);
    assert!(!replacement_format.quote_prefix());

    let mut reset = source.edit().expect("reset edit");
    reset
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .alignment(styled_address, &change)
        .and_then(|sheet| sheet.protection(styled_address, &protection))
        .and_then(|sheet| sheet.quote_prefix(styled_address, Some(true)))
        .and_then(|sheet| sheet.reset_style(styled_address))
        .expect("alignment then style reset");
    let reset = reset.commit().expect("reset commit");
    let reset_sheet = reset
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet");
    assert!(matches!(
        reset_sheet
            .local_style(styled_address)
            .expect("local style"),
        Some(crate::LocalStyle::Default)
    ));
    assert_eq!(
        reset_sheet
            .style(styled_address)
            .expect("style lookup")
            .expect("resolved style")
            .format()
            .expect("cell format")
            .alignment(),
        None
    );
    let reset_format = reset_sheet
        .style(styled_address)
        .expect("style lookup")
        .expect("resolved style")
        .format()
        .expect("cell format");
    assert_eq!(reset_format.protection(), None);
    assert!(!reset_format.quote_prefix());

    let mut remove = source.edit().expect("remove edit");
    remove
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .alignment(styled_address, &change)
        .and_then(|sheet| sheet.protection(styled_address, &protection))
        .and_then(|sheet| sheet.quote_prefix(styled_address, Some(true)))
        .and_then(|sheet| sheet.remove(styled_address))
        .expect("alignment then removal");
    let remove = remove.commit().expect("remove commit");
    assert!(
        remove
            .package()
            .sheet(0usize)
            .expect("lookup")
            .expect("worksheet")
            .cell(styled_address)
            .expect("removed cell")
            .is_missing()
    );
    assert_eq!(part_text(remove.package(), "/xl/styles.xml"), source_styles);
}

#[test]
fn alignment_after_style_reset_uses_reset_base_and_inverse_restores_source() {
    let source = styled_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let address = Address::from_a1("A1").expect("address");
    let change = AlignmentChange::new()
        .horizontal(Some(Horizontal::Center))
        .wrap_text(Some(true));

    let mut edit = source.edit().expect("edit");
    edit.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .reset_style(address)
        .and_then(|sheet| sheet.alignment(address, &change))
        .expect("style reset then alignment");
    let committed = edit.commit().expect("commit");
    let sheet = committed
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet");
    let format = sheet
        .style(address)
        .expect("style lookup")
        .expect("resolved style")
        .format()
        .expect("cell format");
    let alignment = format.alignment().expect("alignment after reset");
    assert_eq!(alignment.horizontal, Some(Horizontal::Center));
    assert!(alignment.wrap_text);
    assert!(matches!(
        sheet.local_style(address).expect("local style"),
        Some(crate::LocalStyle::Shared(_))
    ));

    let restored = committed
        .package()
        .apply(&committed.patch().inverse())
        .expect("inverse");
    assert_eq!(
        restored.package().to_bytes().expect("restored bytes"),
        source_bytes
    );
}

#[test]
fn alignment_and_complete_style_actions_conflict_when_joined() {
    let source = styled_workbook();
    let replacement_style = source.styles().expect("styles").get(0).expect("base style");
    let address = Address::from_a1("A1").expect("address");
    let change = AlignmentChange::new().horizontal(Some(Horizontal::Center));

    let mut alignment = source.edit().expect("alignment edit");
    alignment
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .alignment(address, &change)
        .expect("alignment");
    let mut replacement = source.edit().expect("replacement edit");
    replacement
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address, &replacement_style)
        .expect("style replacement");
    let error = alignment
        .join(replacement)
        .expect_err("alignment and style replacement must conflict");
    let conflicts = error.conflicts().expect("cell conflict details");
    assert_eq!(conflicts.conflicts().len(), 1);
    assert_eq!(
        conflicts.conflicts()[0].cells().expect("cell conflicts"),
        &[address]
    );

    let mut alignment = source.edit().expect("alignment edit");
    alignment
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .alignment(address, &change)
        .expect("alignment");
    let mut reset = source.edit().expect("reset edit");
    reset
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .reset_style(address)
        .expect("style reset");
    let error = alignment
        .join(reset)
        .expect_err("alignment and style reset must conflict");
    let conflicts = error.conflicts().expect("cell conflict details");
    assert_eq!(conflicts.conflicts().len(), 1);
    assert_eq!(
        conflicts.conflicts()[0].cells().expect("cell conflicts"),
        &[address]
    );

    let mut alignment = source.edit().expect("alignment edit");
    alignment
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .alignment(address, &change)
        .expect("alignment");
    let mut remove = source.edit().expect("remove edit");
    remove
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .remove(address)
        .expect("cell removal");
    let error = alignment
        .join(remove)
        .expect_err("alignment and removal must conflict");
    let conflicts = error.conflicts().expect("cell conflict details");
    assert_eq!(conflicts.conflicts().len(), 1);
    assert_eq!(
        conflicts.conflicts()[0].cells().expect("cell conflicts"),
        &[address]
    );
}

#[test]
fn protection_quote_and_extended_alignment_repeat_clear_without_spurious_changes() {
    let source = crate::Package::new().expect("source workbook");
    let address = Address::from_a1("A1").expect("address");
    let alignment = AlignmentChange::new()
        .relative_indent(Some(-3))
        .justify_last_line(Some(false));
    let protection = ProtectionChange::new()
        .locked(Some(false))
        .hidden(Some(true));

    let mut create = source.edit().expect("create edit");
    create
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .alignment(address, &alignment)
        .and_then(|sheet| sheet.protection(address, &protection))
        .and_then(|sheet| sheet.quote_prefix(address, Some(false)))
        .expect("style facets");
    let configured = create.commit().expect("configured commit");
    let configured_bytes = configured.package().to_bytes().expect("configured bytes");
    let configured_styles = part_text(configured.package(), "/xl/styles.xml").to_owned();
    let configured_format = configured
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let configured_alignment = configured_format.alignment().expect("alignment");
    assert_eq!(configured_alignment.relative_indent, Some(-3));
    assert_eq!(configured_alignment.justify_last_line, Some(false));
    let configured_protection = configured_format.protection().expect("protection");
    assert_eq!(configured_protection.locked(), Some(false));
    assert_eq!(configured_protection.hidden(), Some(true));
    assert!(!configured_format.quote_prefix());

    let mut repeat = configured.package().edit().expect("repeat edit");
    repeat
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .alignment(address, &alignment)
        .and_then(|sheet| sheet.protection(address, &protection))
        .and_then(|sheet| sheet.quote_prefix(address, Some(false)))
        .expect("repeat style facets");
    let repeat = repeat.commit().expect("repeat commit");
    assert!(repeat.patch().is_noop());
    assert_eq!(
        repeat.package().to_bytes().expect("repeat bytes"),
        configured_bytes
    );

    let mut clear = configured.package().edit().expect("clear edit");
    clear
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .alignment(
            address,
            &AlignmentChange::new()
                .relative_indent(None)
                .justify_last_line(None),
        )
        .and_then(|sheet| {
            sheet.protection(address, &ProtectionChange::new().locked(None).hidden(None))
        })
        .and_then(|sheet| sheet.quote_prefix(address, None))
        .expect("clear style facets");
    let clear = clear.commit().expect("clear commit");
    assert_eq!(clear.patch().operation_count(), 1);
    let styles = part_text(clear.package(), "/xl/styles.xml");
    assert!(styles.contains("<alignment/>") || styles.contains("<alignment />"));
    assert!(styles.contains("<protection/>") || styles.contains("<protection />"));
    assert_eq!(
        styles.matches(r#"relativeIndent="-3""#).count(),
        configured_styles.matches(r#"relativeIndent="-3""#).count()
    );
    assert_eq!(
        styles.matches(r#"justifyLastLine="0""#).count(),
        configured_styles.matches(r#"justifyLastLine="0""#).count()
    );
    assert_eq!(
        styles.matches(r#"locked="0""#).count(),
        configured_styles.matches(r#"locked="0""#).count()
    );
    assert_eq!(
        styles.matches(r#"hidden="1""#).count(),
        configured_styles.matches(r#"hidden="1""#).count()
    );
    assert_eq!(
        styles.matches("quotePrefix=").count(),
        configured_styles.matches("quotePrefix=").count()
    );
    let clear_format = clear
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    assert_eq!(
        clear_format.alignment(),
        Some(&crate::style::stylesheet::Alignment::new())
    );
    assert_eq!(
        clear_format.protection(),
        Some(crate::style::cell_format::Protection::default())
    );
    assert!(!clear_format.quote_prefix());
}

#[test]
fn protection_alignment_and_quote_facets_join_when_disjoint_and_conflict_when_overlapping() {
    let source = crate::Package::new().expect("source workbook");
    let address = Address::from_a1("A1").expect("address");

    let mut locked = source.edit().expect("locked edit");
    locked
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .protection(address, &ProtectionChange::new().locked(Some(true)))
        .expect("locked");
    let mut hidden = source.edit().expect("hidden edit");
    hidden
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .protection(address, &ProtectionChange::new().hidden(Some(false)))
        .expect("hidden");
    locked.join(hidden).expect("disjoint protection fields");
    let joined = locked.commit().expect("protection join commit");
    let joined_format = joined
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    let joined_protection = joined_format.protection().expect("protection");
    assert_eq!(joined_protection.locked(), Some(true));
    assert_eq!(joined_protection.hidden(), Some(false));

    let mut protection = source.edit().expect("protection edit");
    protection
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .protection(address, &ProtectionChange::new().locked(Some(true)))
        .expect("protection");
    let mut alignment = source.edit().expect("alignment edit");
    alignment
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .alignment(
            address,
            &AlignmentChange::new().horizontal(Some(Horizontal::Center)),
        )
        .expect("alignment");
    protection
        .join(alignment)
        .expect("protection/alignment join");
    let joined = protection.commit().expect("protection/alignment commit");
    let joined_format = joined
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    assert_eq!(
        joined_format.protection().expect("protection").locked(),
        Some(true)
    );
    assert_eq!(
        joined_format.alignment().expect("alignment").horizontal,
        Some(Horizontal::Center)
    );

    let mut quote = source.edit().expect("quote edit");
    quote
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .quote_prefix(address, Some(true))
        .expect("quote");
    let mut quote_protection = source.edit().expect("quote/protection edit");
    quote_protection
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .protection(address, &ProtectionChange::new().hidden(Some(true)))
        .expect("protection");
    quote.join(quote_protection).expect("quote/protection join");
    let joined = quote.commit().expect("quote/protection commit");
    let joined_format = joined
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .style(address)
        .expect("style lookup")
        .expect("styled cell")
        .format()
        .expect("cell format");
    assert!(joined_format.quote_prefix());
    assert_eq!(
        joined_format.protection().expect("protection").hidden(),
        Some(true)
    );

    let mut left = source.edit().expect("left quote edit");
    left.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .quote_prefix(address, Some(true))
        .expect("left quote");
    let mut right = source.edit().expect("right quote edit");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .quote_prefix(address, Some(false))
        .expect("right quote");
    let error = left
        .join(right)
        .expect_err("same quote-prefix facet must conflict");
    let conflicts = error.conflicts().expect("cell conflict details");
    assert_eq!(conflicts.conflicts().len(), 1);
    assert_eq!(
        conflicts.conflicts()[0].cells().expect("cell conflicts"),
        &[address]
    );

    let mut left = source.edit().expect("left protection edit");
    left.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .protection(address, &ProtectionChange::new().locked(Some(true)))
        .expect("left protection");
    let mut right = source.edit().expect("right protection edit");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .protection(address, &ProtectionChange::new().locked(Some(false)))
        .expect("right protection");
    let error = left
        .join(right)
        .expect_err("same protection facet must conflict");
    let conflicts = error.conflicts().expect("cell conflict details");
    assert_eq!(conflicts.conflicts().len(), 1);
    assert_eq!(
        conflicts.conflicts()[0].cells().expect("cell conflicts"),
        &[address]
    );
}

/// Record 0279: a formula cache write keeps the formula markup (shared
/// members included), sets the cell type from the value, and inverts exactly.
#[test]
fn formula_cache_writes_keep_formula_markup() {
    let baseline = crate::Package::new().expect("baseline");
    let mut package = baseline.inner.package.clone();
    let uri = baseline.inner.sheets[0].part_uri.clone();
    package
        .get_part_mut(&uri)
        .expect("worksheet part")
        .set_blob(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>2</v></c><c r="B1"><f t="shared" ref="B1:B2" si="0">A1*2</f><v>0</v></c><c r="C1" t="str"><f>"x"&amp;A1</f><v>old</v></c><c r="D1"><v>5</v></c></row><row r="2"><c r="B2"><f t="shared" si="0"/></c></row></sheetData></worksheet>"#
                .to_vec(),
        );
    let source = crate::Package::from_opc(package).expect("source");
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    {
        let mut sheet = edit.sheet("Sheet1").expect("lookup").expect("sheet");
        sheet
            .set_cached("B1", Value::Number(Number::new("4").expect("number")))
            .expect("B1");
        sheet.set_cached("B2", Value::Bool(true)).expect("B2");
        sheet.set_cached("C1", Value::text("x2")).expect("C1");
        assert!(sheet.set_cached("D1", Value::Bool(false)).is_err());
    }
    let committed = edit.commit().expect("commit");
    let xml = part_text(committed.package(), uri.as_str());
    assert!(
        xml.contains(r#"<c r="B1"><f t="shared" ref="B1:B2" si="0">A1*2</f><v>4</v></c>"#),
        "{xml}"
    );
    assert!(
        xml.contains(r#"<c r="B2" t="b"><f t="shared" si="0"/><v>1</v></c>"#),
        "{xml}"
    );
    assert!(
        xml.contains(r#"<c r="C1" t="str"><f>"x"&amp;A1</f><v>x2</v></c>"#),
        "{xml}"
    );
    let restored = committed
        .package()
        .apply(&committed.patch().inverse())
        .expect("inverse");
    assert_eq!(
        restored.package().to_bytes().expect("restored"),
        source_bytes
    );
}
