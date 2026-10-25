//! Unified XLSX cell writes through the public `File` and transaction APIs.
//!
//! The fixture starts as a native owning workbook, then each test opens the
//! same bytes through both the retained source and eager owner routes.  The
//! assertions exercise the facade contract rather than a format-specific
//! editor: old snapshots stay unchanged, cell properties round-trip through
//! the node view, and retained patches replay onto independently opened equal
//! content after their source proof.
#![cfg(feature = "xlsx")]

use std::fs;

use litchi_core::FileFormat;
use litchi_opc::PackURI;
use litchi_ops::capabilities::Capabilities;
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::value::{Color, ColorContext};
use litchi_ops::{
    BatchItem, BatchMode, Code, File, ItemResult, OpenOptions, PropValue, Props, SourceMode, props,
};

const CELL_A1: &str = "/sheet[1]/A1";
const CELL_B1: &str = "/sheet[1]/B1";
const CELL_C1: &str = "/sheet[1]/C1";
const CELL_D1: &str = "/sheet[1]/D1";
const STYLES_URI: &str = "/xl/styles.xml";
const SHEET1_URI: &str = "/xl/worksheets/sheet1.xml";
const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const MCE_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const X14AC_NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";

/// Build a small workbook with every scalar kind used by the first XLSX
/// write surface, plus a cacheless formula and an explicit boolean cell.
fn fixture() -> Vec<u8> {
    let workbook = litchi_xlsx::Package::new().expect("create native workbook");
    let mut edit = workbook.edit().expect("start native workbook edit");
    {
        let mut sheet = edit
            .sheet("Sheet1")
            .expect("look up native worksheet")
            .expect("minimal workbook has Sheet1");
        sheet.set("A1", "seed").expect("seed text cell");
        sheet.set("B1", 2_i32).expect("seed number cell");
        sheet
            .set(
                "C1",
                litchi_xlsx::Formula::new("B1*2").expect("seed formula"),
            )
            .expect("seed formula cell");
        sheet.set("D1", true).expect("seed boolean cell");
    }
    edit.commit()
        .expect("commit native fixture")
        .into_package()
        .to_bytes()
        .expect("serialize native fixture")
}

fn custom_styles_fixture() -> Vec<u8> {
    let workbook = litchi_xlsx::Package::from_bytes(&fixture()).expect("open native XLSX fixture");
    let package = litchi_xlsx::Package::from_bytes(
        &workbook.to_bytes().expect("serialize native XLSX fixture"),
    )
    .expect("open fixture package")
    .into_plain_opc();
    let styles_uri = PackURI::new(STYLES_URI).expect("styles URI");
    let styles = std::str::from_utf8(package.get_part(&styles_uri).expect("styles").blob())
        .expect("styles XML")
        .replace(
            &format!(r#"<styleSheet xmlns="{MAIN_NS}">"#),
            &format!(
                r#"<styleSheet xmlns="{MAIN_NS}" xmlns:mc="{MCE_NS}" xmlns:x14ac="{X14AC_NS}" mc:Ignorable="x14ac">"#
            ),
        )
        .replace(
            r#"<fonts count="1">"#,
            r#"<fonts count="1" x14ac:knownFonts="1">"#,
        )
        .replace(
            r#"<font/>"#,
            r#"<font x14ac:fontMarker="keep"><x14ac:fontExtension marker="keep"/></font>"#,
        )
        .replace(
            r#"<cellXfs count="1"><xf "#,
            r#"<cellXfs count="1"><xf x14ac:xfMarker="keep" "#,
        );
    let mut package = package;
    package
        .get_part_mut(&styles_uri)
        .expect("styles part")
        .set_blob(styles.into_bytes());
    litchi_xlsx::Package::from_opc(package)
        .expect("adopt custom styles package")
        .to_plain_bytes()
        .expect("serialize custom styles package")
}

/// Add a producer-authored second font and cell format rather than creating
/// the long-tail children through the unified writer.  A1 is the only cell
/// that references the explicit style, so query results can distinguish
/// authored false/zero/none values from absent font metadata.
fn source_long_tail_font_fixture() -> Vec<u8> {
    let workbook = litchi_xlsx::Package::from_bytes(&fixture()).expect("open native XLSX fixture");
    let package = litchi_xlsx::Package::from_bytes(
        &workbook.to_bytes().expect("serialize native XLSX fixture"),
    )
    .expect("open fixture package")
    .into_plain_opc();
    let styles_uri = PackURI::new(STYLES_URI).expect("styles URI");
    let styles = std::str::from_utf8(package.get_part(&styles_uri).expect("styles").blob())
        .expect("styles XML")
        .replace(
            r#"<fonts count="1"><font/></fonts>"#,
            r#"<fonts count="2"><font/><font><condense val="0"/><extend val="0"/><outline val="0"/><shadow val="0"/><family val="0"/><charset val="0"/><scheme val="none"/></font></fonts>"#,
        )
        .replace(r#"<cellXfs count="1">"#, r#"<cellXfs count="2">"#)
        .replace(
            r#"</cellXfs>"#,
            r#"<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>"#,
        );
    assert!(styles.contains(r#"<font><condense val="0"/>"#));
    assert!(styles.contains(r#"fontId="1""#));

    let sheet_uri = PackURI::new(SHEET1_URI).expect("sheet URI");
    let sheet = std::str::from_utf8(package.get_part(&sheet_uri).expect("sheet").blob())
        .expect("sheet XML")
        .replace(r#"<c r="A1""#, r#"<c r="A1" s="1""#);
    assert!(sheet.contains(r#"<c r="A1" s="1""#));

    let mut package = package;
    package
        .get_part_mut(&styles_uri)
        .expect("styles part")
        .set_blob(styles.into_bytes());
    package
        .get_part_mut(&sheet_uri)
        .expect("sheet part")
        .set_blob(sheet.into_bytes());
    litchi_xlsx::Package::from_opc(package)
        .expect("adopt source-authored font package")
        .to_plain_bytes()
        .expect("serialize source-authored font package")
}

/// Add a producer-authored alignment and an opaque alignment attribute.  The
/// latter is deliberately outside the shared facade vocabulary: a cell edit
/// must preserve it while changing a known alignment child.
fn source_alignment_fixture() -> Vec<u8> {
    let workbook = litchi_xlsx::Package::from_bytes(&fixture()).expect("open native XLSX fixture");
    let package = litchi_xlsx::Package::from_bytes(
        &workbook.to_bytes().expect("serialize native XLSX fixture"),
    )
    .expect("open fixture package")
    .into_plain_opc();
    let styles_uri = PackURI::new(STYLES_URI).expect("styles URI");
    let mut styles = std::str::from_utf8(package.get_part(&styles_uri).expect("styles").blob())
        .expect("styles XML")
        .replace(
            &format!(r#"<styleSheet xmlns="{MAIN_NS}">"#),
            &format!(r#"<styleSheet xmlns="{MAIN_NS}" xmlns:future="urn:fixture:future">"#),
        );
    let cell_xfs_start = styles
        .find(r#"<cellXfs count="1">"#)
        .expect("cell-XF collection");
    let cell_xfs_end = cell_xfs_start
        + styles[cell_xfs_start..]
            .find(r#"</cellXfs>"#)
            .expect("cell-XF collection end")
        + r#"</cellXfs>"#.len();
    let cell_xfs = &styles[cell_xfs_start..cell_xfs_end];
    let xf_end = cell_xfs.find(r#"/>"#).expect("cell-XF end");
    let xf_start = cell_xfs.find("<xf ").expect("cell-XF start");
    let replacement = format!(
        r#"<cellXfs count="2">{}{} applyAlignment="1" future:xfMarker="keep"><alignment horizontal="center" vertical="bottom" wrapText="1" textRotation="90" indent="3" relativeIndent="0" justifyLastLine="0" shrinkToFit="1" readingOrder="2" future:alignmentMarker="keep"/></xf>{}"#,
        &cell_xfs[xf_start..xf_end + 2],
        &cell_xfs[xf_start..xf_end],
        &cell_xfs[xf_end + 2..],
    );
    styles.replace_range(cell_xfs_start..cell_xfs_end, &replacement);

    let sheet_uri = PackURI::new(SHEET1_URI).expect("sheet URI");
    let sheet = std::str::from_utf8(package.get_part(&sheet_uri).expect("sheet").blob())
        .expect("sheet XML")
        .replace(r#"<c r="A1""#, r#"<c r="A1" s="1""#);
    assert!(styles.contains(r#"future:alignmentMarker="keep""#));
    assert!(sheet.contains(r#"<c r="A1" s="1""#));

    let mut package = package;
    package
        .get_part_mut(&styles_uri)
        .expect("styles part")
        .set_blob(styles.into_bytes());
    package
        .get_part_mut(&sheet_uri)
        .expect("sheet part")
        .set_blob(sheet.into_bytes());
    litchi_xlsx::Package::from_opc(package)
        .expect("adopt source-authored alignment package")
        .to_plain_bytes()
        .expect("serialize source-authored alignment package")
}

/// Add two source-authored pattern fills, a gradient fill, and
/// opaque producer markup inside the first pattern.  A1 and B1 deliberately
/// start with different pattern/background pairs so range edits exercise
/// per-cell source-style composition rather than one shared default style.
fn source_fill_fixture() -> Vec<u8> {
    let workbook = litchi_xlsx::Package::from_bytes(&fixture()).expect("open native XLSX fixture");
    let mut edit = workbook.edit().expect("start source fill edit");
    {
        let mut sheet = edit
            .sheet("Sheet1")
            .expect("look up native worksheet")
            .expect("minimal workbook has Sheet1");
        let first = litchi_xlsx::style::FillChange::new()
            .pattern(Some(litchi_xlsx::style::FillPattern::DarkGray))
            .foreground(Some(litchi_xlsx::style::FillColor::rgb(0x10, 0x20, 0x30)))
            .background(Some(
                litchi_xlsx::style::FillColor::theme(4).with_tint(
                    litchi_xlsx::style::stylesheet::border::Tint::new(0.4)
                        .expect("valid source fill tint"),
                ),
            ));
        sheet
            .fill("A1", &first)
            .expect("author first source pattern");
        let second = litchi_xlsx::style::FillChange::new()
            .pattern(Some(litchi_xlsx::style::FillPattern::LightGray))
            .foreground(Some(litchi_xlsx::style::FillColor::rgb(0x40, 0x50, 0x60)))
            .background(Some(litchi_xlsx::style::FillColor::rgb(0xF0, 0xE0, 0xD0)));
        sheet
            .fill("B1", &second)
            .expect("author second source pattern");
    }
    let workbook = edit
        .commit()
        .expect("commit source fill styles")
        .into_package();
    let package = litchi_xlsx::Package::from_bytes(
        &workbook.to_bytes().expect("serialize source fill workbook"),
    )
    .expect("open source fill package")
    .into_plain_opc();
    let styles_uri = PackURI::new(STYLES_URI).expect("styles URI");
    let mut styles = std::str::from_utf8(package.get_part(&styles_uri).expect("styles").blob())
        .expect("styles XML")
        .replace(
            &format!(r#"<styleSheet xmlns="{MAIN_NS}">"#),
            &format!(r#"<styleSheet xmlns="{MAIN_NS}" xmlns:future="urn:fixture:future">"#),
        );

    let dark_pattern = r#"<patternFill patternType="darkGray">"#;
    let dark_start = styles.find(dark_pattern).expect("source dark pattern");
    styles.replace_range(
        dark_start..dark_start + dark_pattern.len(),
        r#"<patternFill patternType="darkGray" future:patternMarker="keep">"#,
    );
    let dark_end = styles[dark_start..]
        .find("</patternFill>")
        .map(|offset| dark_start + offset)
        .expect("source dark pattern end");
    styles.insert_str(dark_end, r#"<future:fillExtension marker="keep"/>"#);

    let fills_start = styles.find(r#"<fills count="#).expect("fills collection");
    let fills_count_start = fills_start + "<fills count=\"".len();
    let fills_count_end = styles[fills_count_start..]
        .find('"')
        .map(|offset| fills_count_start + offset)
        .expect("fills count end");
    let fills_count = styles[fills_count_start..fills_count_end]
        .parse::<usize>()
        .expect("numeric fills count");
    styles.replace_range(
        fills_count_start..fills_count_end,
        &(fills_count + 1).to_string(),
    );
    let fills_end = styles.find("</fills>").expect("fills collection end");
    styles.insert_str(
        fills_end,
        r#"<fill><gradientFill type="linear" degree="45"><stop position="0"><color rgb="FFFF0000"/></stop><stop position="1"><color rgb="FF0000FF"/></stop></gradientFill></fill>"#,
    );

    let xfs_start = styles
        .find(r#"<cellXfs count="#)
        .expect("cell-XF collection");
    let xfs_count_start = xfs_start + "<cellXfs count=\"".len();
    let xfs_count_end = styles[xfs_count_start..]
        .find('"')
        .map(|offset| xfs_count_start + offset)
        .expect("cell-XF count end");
    let xfs_count = styles[xfs_count_start..xfs_count_end]
        .parse::<usize>()
        .expect("numeric cell-XF count");
    styles.replace_range(xfs_count_start..xfs_count_end, &(xfs_count + 1).to_string());
    let xfs_end = styles.find("</cellXfs>").expect("cell-XF collection end");
    styles.insert_str(
        xfs_end,
        r#"<xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>"#,
    );

    let sheet_uri = PackURI::new(SHEET1_URI).expect("sheet URI");
    let mut sheet = std::str::from_utf8(package.get_part(&sheet_uri).expect("sheet").blob())
        .expect("sheet XML")
        .to_owned();
    sheet = sheet.replace(r#"<c r="C1""#, &format!(r#"<c r="C1" s="{}""#, xfs_count));
    assert!(styles.contains(r#"future:patternMarker="keep""#));
    assert!(styles.contains(r#"future:fillExtension marker="keep""#));
    assert!(sheet.contains(&format!(r#"<c r="C1" s="{}""#, xfs_count)));

    let mut package = package;
    package
        .get_part_mut(&styles_uri)
        .expect("styles part")
        .set_blob(styles.into_bytes());
    package
        .get_part_mut(&sheet_uri)
        .expect("sheet part")
        .set_blob(sheet.into_bytes());
    litchi_xlsx::Package::from_opc(package)
        .expect("adopt source-authored fill package")
        .to_plain_bytes()
        .expect("serialize source-authored fill package")
}

/// Replace the simple source gradient with a path gradient carrying three
/// stops.  The ordinary node projection intentionally reports only the
/// OfficeCLI string; the native workbook read below checks that geometry,
/// positions, and every authored colour remain available to the leaf.
fn source_path_gradient_fixture() -> Vec<u8> {
    let workbook =
        litchi_xlsx::Package::from_bytes(&source_fill_fixture()).expect("open source fill fixture");
    let package = litchi_xlsx::Package::from_bytes(
        &workbook
            .to_bytes()
            .expect("serialize source path-gradient fixture"),
    )
    .expect("open source path-gradient package")
    .into_plain_opc();
    let styles_uri = PackURI::new(STYLES_URI).expect("styles URI");
    let mut styles = std::str::from_utf8(package.get_part(&styles_uri).expect("styles").blob())
        .expect("styles XML")
        .to_owned();
    let linear = r#"<gradientFill type="linear" degree="45"><stop position="0"><color rgb="FFFF0000"/></stop><stop position="1"><color rgb="FF0000FF"/></stop></gradientFill>"#;
    let path = r#"<gradientFill type="path" left="0.1" right="0.8" top="0.2" bottom="0.9"><stop position="0.125"><color rgb="FFFF0000"/></stop><stop position="0.5"><color rgb="FF00FF00"/></stop><stop position="0.875"><color rgb="FF0000FF"/></stop></gradientFill>"#;
    assert!(styles.contains(linear), "source gradient fixture changed");
    styles = styles.replace(linear, path);

    let mut package = package;
    package
        .get_part_mut(&styles_uri)
        .expect("styles part")
        .set_blob(styles.into_bytes());
    litchi_xlsx::Package::from_opc(package)
        .expect("adopt source path-gradient package")
        .to_plain_bytes()
        .expect("serialize source path-gradient package")
}

/// Add a producer-authored border with every typed side/color variant and
/// opaque extension markup.  A1 alone points at the second cell-XF, so a
/// border edit can be checked against both the source projection and the
/// native format reader without changing the shared default border.
fn source_border_fixture() -> Vec<u8> {
    let workbook = litchi_xlsx::Package::from_bytes(&fixture()).expect("open native XLSX fixture");
    let package = litchi_xlsx::Package::from_bytes(
        &workbook.to_bytes().expect("serialize native XLSX fixture"),
    )
    .expect("open fixture package")
    .into_plain_opc();
    let styles_uri = PackURI::new(STYLES_URI).expect("styles URI");
    let mut styles = std::str::from_utf8(package.get_part(&styles_uri).expect("styles").blob())
        .expect("styles XML")
        .replace(
            &format!(r#"<styleSheet xmlns="{MAIN_NS}">"#),
            &format!(r#"<styleSheet xmlns="{MAIN_NS}" xmlns:future="urn:fixture:future">"#),
        );

    let borders_start = styles
        .find(r#"<borders count="#)
        .expect("borders collection");
    let count_start = borders_start + r#"<borders count="#.len() + 1;
    let count_end = count_start + styles[count_start..].find('"').expect("borders count end");
    let count = styles[count_start..count_end]
        .parse::<usize>()
        .expect("numeric borders count");
    styles.replace_range(count_start..count_end, &(count + 1).to_string());
    let borders_end = styles.find("</borders>").expect("borders collection end");
    styles.insert_str(
        borders_end,
        r#"<border diagonalUp="1" diagonalDown="1" outline="0" future:borderMarker="keep"><left style="thin" future:leftMarker="keep"><color rgb="FF102030"/></left><right style="double"><color theme="4" tint="0.25"/></right><top style="medium"/><bottom style="dotted"><color indexed="12"/></bottom><diagonal style="hair"><color auto="1"/></diagonal><future:borderExtension marker="keep"/></border>"#,
    );

    let xfs_start = styles
        .find(r#"<cellXfs count="#)
        .expect("cell-XF collection");
    let xfs_count_start = xfs_start + r#"<cellXfs count="#.len() + 1;
    let xfs_count_end = xfs_count_start
        + styles[xfs_count_start..]
            .find('"')
            .expect("cell-XF count end");
    let xfs_count = styles[xfs_count_start..xfs_count_end]
        .parse::<usize>()
        .expect("numeric cell-XF count");
    styles.replace_range(xfs_count_start..xfs_count_end, &(xfs_count + 1).to_string());
    let xfs_end = styles.find("</cellXfs>").expect("cell-XF collection end");
    styles.insert_str(
        xfs_end,
        r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" future:xfBorderMarker="keep"/>"#,
    );

    let sheet_uri = PackURI::new(SHEET1_URI).expect("sheet URI");
    let sheet = std::str::from_utf8(package.get_part(&sheet_uri).expect("sheet").blob())
        .expect("sheet XML")
        .replace(r#"<c r="A1""#, r#"<c r="A1" s="1""#);
    assert!(styles.contains(r#"future:borderMarker="keep""#));
    assert!(styles.contains(r#"future:borderExtension marker="keep""#));
    assert!(sheet.contains(r#"<c r="A1" s="1""#));

    let mut package = package;
    package
        .get_part_mut(&styles_uri)
        .expect("styles part")
        .set_blob(styles.into_bytes());
    package
        .get_part_mut(&sheet_uri)
        .expect("sheet part")
        .set_blob(sheet.into_bytes());
    litchi_xlsx::Package::from_opc(package)
        .expect("adopt source-authored border package")
        .to_plain_bytes()
        .expect("serialize source-authored border package")
}

/// Make the workbook's style-zero cell format point at the authored dark
/// pattern.  A missing cell inherits this format even though the ordinary
/// node projection intentionally omits inherited/default format details.
fn source_base_pattern_fixture() -> Vec<u8> {
    let workbook =
        litchi_xlsx::Package::from_bytes(&source_fill_fixture()).expect("open source fill fixture");
    let package = litchi_xlsx::Package::from_bytes(
        &workbook
            .to_bytes()
            .expect("serialize source base-pattern fixture"),
    )
    .expect("open source base-pattern package")
    .into_plain_opc();
    let styles_uri = PackURI::new(STYLES_URI).expect("styles URI");
    let mut styles = std::str::from_utf8(package.get_part(&styles_uri).expect("styles").blob())
        .expect("styles XML")
        .to_owned();
    let cell_xfs_start = styles
        .find(r#"<cellXfs count="#)
        .expect("cell-XF collection");
    let first_xf_start = cell_xfs_start
        + styles[cell_xfs_start..]
            .find("<xf ")
            .expect("first cell-XF");
    let first_xf_end = first_xf_start
        + styles[first_xf_start..]
            .find("/>")
            .expect("first cell-XF end")
        + 2;
    let first_xf = styles[first_xf_start..first_xf_end].replace(r#"fillId="0""#, r#"fillId="2""#);
    assert!(first_xf.contains(r#"fillId="2""#));
    styles.replace_range(first_xf_start..first_xf_end, &first_xf);

    let mut package = package;
    package
        .get_part_mut(&styles_uri)
        .expect("styles part")
        .set_blob(styles.into_bytes());
    litchi_xlsx::Package::from_opc(package)
        .expect("adopt source base-pattern package")
        .to_plain_bytes()
        .expect("serialize source base-pattern package")
}

fn open_owner(bytes: &[u8], eager: bool) -> File {
    if eager {
        litchi_ops::open_bytes_with(
            bytes.to_vec(),
            &OpenOptions::new().source_mode(SourceMode::Eager),
        )
        .expect("open eager XLSX owner")
    } else {
        litchi_ops::open_bytes(bytes.to_vec()).expect("open source XLSX owner")
    }
}

fn cell_text(file: &File, path: &str) -> Option<String> {
    file.get(path)
        .expect("resolve cell")
        .view()
        .text()
        .expect("read cell text")
        .map(std::borrow::Cow::into_owned)
}

fn cell_property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .expect("resolve cell")
        .view()
        .format()
        .expect("read cell properties")
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn styles_bytes(file: &File) -> Vec<u8> {
    let package = litchi_xlsx::Package::from_bytes(&file.to_bytes().expect("publish XLSX"))
        .expect("open published XLSX package")
        .into_plain_opc();
    package
        .get_part(&PackURI::new(STYLES_URI).expect("styles URI"))
        .expect("styles part")
        .blob()
        .to_vec()
}

fn assert_value_refusal(error: litchi_ops::Error) {
    assert!(
        matches!(
            error.code(),
            Code::InvalidArgument | Code::InvalidValue | Code::InvalidInput
        ),
        "expected a typed property refusal, got {error:?}"
    );
}

fn batch_set(path: &'static str, properties: Props<'static>) -> BatchItem<'static> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(properties);
    item
}

#[test]
fn xlsx_cell_capabilities_match_the_unified_write_registry() {
    let cell = Capabilities::of(FileFormat::Xlsx)
        .element("cell")
        .expect("XLSX cell capability");
    for (key, get, set) in [
        ("value", false, true),
        ("formula", true, true),
        ("clear", false, true),
        ("font.bold", true, true),
        ("font.italic", true, true),
        ("strike", true, true),
        ("underline", true, true),
        ("font.name", true, true),
        ("font.size", true, true),
        ("superscript", true, true),
        ("subscript", true, true),
        ("font.color", true, true),
        ("border.left", true, true),
        ("border.left.color", true, true),
        ("border.right", true, true),
        ("border.right.color", true, true),
        ("border.top", true, true),
        ("border.top.color", true, true),
        ("border.bottom", true, true),
        ("border.bottom.color", true, true),
        ("border.diagonal", true, true),
        ("border.diagonal.color", true, true),
        ("border.diagonalUp", true, true),
        ("border.diagonalDown", true, true),
        ("alignment.horizontal", true, true),
        ("alignment.vertical", true, true),
        ("alignment.wrapText", true, true),
        ("alignment.textRotation", true, true),
        ("alignment.indent", true, true),
        ("alignment.relativeIndent", true, true),
        ("alignment.justifyLastLine", true, true),
        ("alignment.shrinkToFit", true, true),
        ("alignment.readingOrder", true, true),
        ("protection.locked", true, true),
        ("protection.hidden", true, true),
        ("quotePrefix", true, true),
    ] {
        let property = cell
            .property(key)
            .unwrap_or_else(|| panic!("missing XLSX cell property {key}"));
        assert_eq!(property.get(), get, "XLSX cell {key} get flag");
        assert_eq!(property.set(), set, "XLSX cell {key} set flag");
    }
}

#[test]
fn font_assignments_preserve_payloads_and_shared_neighbours() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("original bytes");
        let updated = original
            .set(
                CELL_A1,
                props! { "font.bold" => true, "font.italic" => false, "value" => "styled" },
            )
            .expect("combined payload and font assignment");
        assert_eq!(cell_text(&updated, CELL_A1).as_deref(), Some("styled"));
        assert_eq!(
            cell_property(&updated, CELL_A1, "font.bold").as_deref(),
            Some("true")
        );
        assert_eq!(
            cell_property(&updated, CELL_A1, "font.italic").as_deref(),
            Some("false")
        );
        assert_eq!(cell_property(&updated, CELL_B1, "font.bold"), None);
        assert_eq!(original.to_bytes().expect("unchanged original"), before);
        let reopened = open_owner(&updated.to_bytes().expect("publish font edit"), eager);
        assert_eq!(
            cell_property(&reopened, CELL_A1, "font.bold").as_deref(),
            Some("true")
        );
        assert_eq!(
            cell_property(&reopened, CELL_A1, "font.italic").as_deref(),
            Some("false")
        );

        let cleared = updated
            .set(CELL_A1, Props::new().with("font.bold", PropValue::Null))
            .expect("clear only bold");
        assert_eq!(cell_property(&cleared, CELL_A1, "font.bold"), None);
        assert_eq!(
            cell_property(&cleared, CELL_A1, "font.italic").as_deref(),
            Some("false")
        );
        assert_eq!(cell_text(&cleared, CELL_A1).as_deref(), Some("styled"));
    }
}

#[test]
fn font_name_size_and_script_assignments_round_trip_source_and_eager() {
    let bytes = fixture();
    let authored_name = "A&B <quoted> \"family\"\r\n\t字体";

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("original bytes");
        let updated = original
            .set(
                CELL_A1,
                props! {
                    "value" => "styled",
                    "font.name" => authored_name,
                    "font.size" => "14.5pt",
                    "superscript" => true,
                },
            )
            .expect("font name, size and superscript assignment");

        assert_eq!(cell_text(&updated, CELL_A1).as_deref(), Some("styled"));
        assert_eq!(
            cell_property(&updated, CELL_A1, "font.name").as_deref(),
            Some(authored_name)
        );
        assert_eq!(
            cell_property(&updated, CELL_A1, "font.size").as_deref(),
            Some("14.5pt")
        );
        assert_eq!(
            cell_property(&updated, CELL_A1, "superscript").as_deref(),
            Some("true")
        );
        assert_eq!(cell_property(&updated, CELL_A1, "subscript"), None);
        assert_eq!(cell_property(&updated, CELL_B1, "font.name"), None);

        let styles = styles_bytes(&updated);
        let styles = std::str::from_utf8(&styles).expect("styles XML");
        assert!(
            styles.contains("A&amp;B &lt;quoted") && styles.contains("&quot;family&quot;"),
            "font names must be XML escaped: {styles}"
        );

        let reopened = open_owner(&updated.to_bytes().expect("publish font edit"), eager);
        assert_eq!(
            cell_property(&reopened, CELL_A1, "font.name").as_deref(),
            Some(authored_name)
        );
        assert_eq!(
            cell_property(&reopened, CELL_A1, "font.size").as_deref(),
            Some("14.5pt")
        );
        assert_eq!(
            cell_property(&reopened, CELL_A1, "superscript").as_deref(),
            Some("true")
        );

        let baseline = original
            .set(CELL_B1, Props::new().with("superscript", false))
            .expect("false superscript authors baseline");
        assert_eq!(cell_property(&baseline, CELL_B1, "superscript"), None);
        assert_eq!(cell_property(&baseline, CELL_B1, "subscript"), None);
        let baseline_styles = styles_bytes(&baseline);
        let baseline_styles = std::str::from_utf8(&baseline_styles).expect("styles XML");
        assert!(
            baseline_styles.contains(r#"<vertAlign val="baseline"/>"#),
            "false script assignment must author SpreadsheetML baseline: {baseline_styles}"
        );

        let subscript = original
            .set(CELL_B1, Props::new().with("subscript", true))
            .expect("subscript assignment");
        assert_eq!(
            cell_property(&subscript, CELL_B1, "subscript").as_deref(),
            Some("true")
        );
        assert_eq!(cell_property(&subscript, CELL_B1, "superscript"), None);

        let cleared = updated
            .set(
                CELL_A1,
                Props::new()
                    .with("font.name", PropValue::Null)
                    .with("font.size", PropValue::Null)
                    .with("superscript", PropValue::Null),
            )
            .expect("null removes authored name, size and script");
        assert_eq!(cell_property(&cleared, CELL_A1, "font.name"), None);
        assert_eq!(cell_property(&cleared, CELL_A1, "font.size"), None);
        assert_eq!(cell_property(&cleared, CELL_A1, "superscript"), None);
        assert_eq!(cell_text(&cleared, CELL_A1).as_deref(), Some("styled"));

        let absent_clear = original
            .set(
                CELL_A1,
                Props::new()
                    .with("font.name", PropValue::Null)
                    .with("font.size", PropValue::Null)
                    .with("superscript", PropValue::Null),
            )
            .expect("clearing absent font properties is a no-op");
        assert_eq!(absent_clear.to_bytes().expect("absent clear bytes"), before);
    }
}

#[test]
fn font_name_and_size_aliases_share_canonical_assignments() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        for (key, value) in [("font", "Alias & Face"), ("fontname", "Alias Name")] {
            let updated = original
                .set(CELL_A1, Props::new().with(key, value))
                .expect("font name alias");
            assert_eq!(
                cell_property(&updated, CELL_A1, "font.name").as_deref(),
                Some(value),
                "font name alias {key}"
            );
        }
        for (key, value) in [("size", "12pt"), ("fontsize", "18pt")] {
            let updated = original
                .set(CELL_A1, Props::new().with(key, value))
                .expect("font size alias");
            assert_eq!(
                cell_property(&updated, CELL_A1, "font.size").as_deref(),
                Some(value),
                "font size alias {key}"
            );
        }
        for (key, canonical) in [
            ("font.superscript", "superscript"),
            ("font.subscript", "subscript"),
        ] {
            let updated = original
                .set(CELL_A1, Props::new().with(key, true))
                .expect("script alias");
            assert_eq!(
                cell_property(&updated, CELL_A1, canonical).as_deref(),
                Some("true"),
                "script alias {key}"
            );
        }

        let before = original.to_bytes().expect("original bytes");
        for properties in [
            Props::new()
                .with("font", "first")
                .with("font.name", "second"),
            Props::new().with("size", "12pt").with("font.size", "13pt"),
        ] {
            let error = original
                .set(CELL_A1, properties)
                .expect_err("canonical and alias duplicate must refuse");
            assert_eq!(error.code(), Code::InvalidArgument);
            assert_eq!(original.to_bytes().expect("unchanged bytes"), before);
        }
    }
}

#[test]
fn font_name_size_and_script_ranges_replay_inverse_and_repeat_without_changes() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("original bytes");
        let commit = original
            .edit(|edit| {
                edit.set(
                    CELL_A1,
                    props! { "value" => "mixed", "font.name" => "Mixed Face" },
                )?;
                edit.set(
                    "/sheet[1]/A1:D2",
                    props! { "font.size" => "18pt", "superscript" => true },
                )?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("mixed payload and font range edits");

        for path in [CELL_A1, CELL_B1, CELL_C1, CELL_D1, "/sheet[1]/D2"] {
            assert_eq!(
                cell_property(commit.file(), path, "font.size").as_deref(),
                Some("18pt"),
                "{path} font size"
            );
            assert_eq!(
                cell_property(commit.file(), path, "superscript").as_deref(),
                Some("true"),
                "{path} superscript"
            );
        }
        assert_eq!(cell_text(commit.file(), CELL_A1).as_deref(), Some("mixed"));
        assert_eq!(
            cell_property(commit.file(), CELL_A1, "font.name").as_deref(),
            Some("Mixed Face")
        );
        assert_eq!(
            cell_property(commit.file(), CELL_C1, "formula").as_deref(),
            Some("B1*2")
        );

        let replayed = commit
            .patch()
            .apply(&original)
            .expect("replay font range patch");
        assert_eq!(
            replayed.to_bytes().expect("replayed bytes"),
            commit.file().to_bytes().expect("committed bytes")
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(&replayed)
            .expect("inverse font range patch");
        assert_eq!(restored.to_bytes().expect("restored bytes"), before);

        let repeated = commit
            .file()
            .set(
                "/sheet[1]/A1:D2",
                props! { "font.size" => "18pt", "superscript" => true },
            )
            .expect("repeat same range formatting");
        assert_eq!(
            repeated.to_bytes().expect("repeat bytes"),
            commit.file().to_bytes().expect("first bytes")
        );
    }
}

#[test]
fn invalid_font_name_size_and_script_assignments_are_atomic() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("original bytes");

        for value in ["", "bad\u{1}name"] {
            let error = original
                .set(CELL_A1, Props::new().with("font.name", value))
                .expect_err("empty or XML-invalid font name");
            assert_eq!(error.code(), Code::InvalidValue);
            assert_eq!(error.property(), Some("font.name"));
            assert_eq!(original.to_bytes().expect("unchanged name owner"), before);
        }

        for value in ["0pt", "0.25pt", "409.01pt", "410pt", "1cm"] {
            let error = original
                .set(CELL_A1, Props::new().with("font.size", value))
                .expect_err("font size outside the XLSX facade domain");
            assert_eq!(error.code(), Code::InvalidValue, "size {value}");
            assert_eq!(error.property(), Some("font.size"), "size {value}");
            assert_eq!(original.to_bytes().expect("unchanged size owner"), before);
        }

        let max_size = original
            .set(CELL_A1, Props::new().with("font.size", "409pt"))
            .expect("409pt is the inclusive maximum");
        assert_eq!(
            cell_property(&max_size, CELL_A1, "font.size").as_deref(),
            Some("409pt")
        );

        let invalid_script = original
            .set(CELL_A1, Props::new().with("superscript", "sometimes"))
            .expect_err("script values require booleans");
        assert_eq!(invalid_script.code(), Code::InvalidValue);
        assert_eq!(invalid_script.property(), Some("superscript"));
        assert_eq!(original.to_bytes().expect("unchanged script owner"), before);

        let conflict = original
            .set(
                CELL_A1,
                Props::new()
                    .with("superscript", true)
                    .with("subscript", true),
            )
            .expect_err("superscript and subscript share one native field");
        assert_eq!(conflict.code(), Code::InvalidArgument);
        assert_eq!(
            original.to_bytes().expect("unchanged conflict owner"),
            before
        );
    }
}

#[test]
fn staged_font_name_bytes_are_checked_before_xlsx_publication() {
    let bytes = fixture();
    let long_name = "N".repeat(256);
    let limits = SemanticLimits::builder()
        .max_staged_bytes(64)
        .expect("staged-byte limit is representable")
        .build()
        .expect("build staged-byte limit");

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("original bytes");
        let error = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |edit| {
                    edit.set(CELL_A1, Props::new().with("font.name", long_name.as_str()))?;
                    Ok::<_, litchi_ops::Error>(())
                },
            )
            .expect_err("font name bytes must count against staged budget");
        assert_eq!(error.code(), Code::LimitExceeded);
        assert_eq!(
            original.to_bytes().expect("unchanged limited owner"),
            before
        );
    }
}

#[test]
fn font_colors_accept_named_hex_css_and_typed_values_with_exact_argb_order() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        for (input, expected, native) in [
            ("red", "#FF0000", r#"rgb="FFFF0000""#),
            ("#10203080", "#10203080", r#"rgb="80102030""#),
            ("rgba(10,20,30,0.5)", "#0A141E80", r#"rgb="800A141E""#),
        ] {
            let updated = original
                .set(CELL_A1, Props::new().with("font.color", input))
                .expect("spreadsheet color grammar");
            assert_eq!(
                cell_property(&updated, CELL_A1, "font.color").as_deref(),
                Some(expected),
                "color input {input}"
            );
            let styles = styles_bytes(&updated);
            let styles = std::str::from_utf8(&styles).expect("styles XML");
            assert!(styles.contains(native), "{input}: {styles}");
        }

        let typed = original
            .set(
                CELL_A1,
                Props::new().with("value", "typed color").with(
                    "font.color",
                    PropValue::Color(Color::rgba(0x11, 0x22, 0x33, 0x44)),
                ),
            )
            .expect("typed spreadsheet color");
        assert_eq!(cell_text(&typed, CELL_A1).as_deref(), Some("typed color"));
        assert_eq!(
            cell_property(&typed, CELL_A1, "font.color").as_deref(),
            Some("#11223344")
        );
        let styles = styles_bytes(&typed);
        let styles = std::str::from_utf8(&styles).expect("styles XML");
        assert!(styles.contains(r#"rgb="44112233""#), "{styles}");

        let reopened = open_owner(&typed.to_bytes().expect("publish typed color"), eager);
        assert_eq!(
            cell_property(&reopened, CELL_A1, "font.color").as_deref(),
            Some("#11223344")
        );
    }
}

#[test]
fn font_color_theme_tint_is_preserved_while_rgb_tint_readback_is_dropped() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let theme = original
            .set(CELL_A1, Props::new().with("font.color", "accent1+tint40"))
            .expect("theme tint color");
        assert_eq!(
            cell_property(&theme, CELL_A1, "font.color").as_deref(),
            Some("accent1+tint40")
        );
        let theme_styles = styles_bytes(&theme);
        let theme_styles = std::str::from_utf8(&theme_styles).expect("styles XML");
        assert!(
            theme_styles.contains(r#"<color theme="4" tint="0.4"/>"#),
            "theme tint must remain native: {theme_styles}"
        );

        let rgb_tinted = original
            .set(CELL_A1, Props::new().with("font.color", "#102030+tint40"))
            .expect("RGB tint color");
        assert_eq!(
            cell_property(&rgb_tinted, CELL_A1, "font.color").as_deref(),
            Some("#102030"),
            "the ordinary projection follows OfficeCLI and drops RGB tint"
        );
        let rgb_styles = styles_bytes(&rgb_tinted);
        let rgb_styles = std::str::from_utf8(&rgb_styles).expect("styles XML");
        assert!(
            rgb_styles.contains(r#"<color rgb="FF102030" tint="0.4"/>"#),
            "RGB tint must remain in the native stylesheet: {rgb_styles}"
        );
    }
}

#[test]
fn font_color_ranges_companions_replay_inverse_and_repeat_without_changes() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("original bytes");
        let commit = original
            .edit(|edit| {
                edit.set(
                    CELL_A1,
                    Props::new().with("value", "colored").with(
                        "font.color",
                        PropValue::Color(Color::rgba(0x11, 0x22, 0x33, 0x44)),
                    ),
                )?;
                edit.set("/sheet[1]/A1:D2", Props::new().with("font.color", "red"))?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("payload and range color edits");

        assert_eq!(
            cell_text(commit.file(), CELL_A1).as_deref(),
            Some("colored")
        );
        for path in [CELL_A1, CELL_B1, CELL_C1, CELL_D1, "/sheet[1]/D2"] {
            assert_eq!(
                cell_property(commit.file(), path, "font.color").as_deref(),
                Some("#FF0000"),
                "{path} font color"
            );
        }
        assert_eq!(
            cell_property(commit.file(), CELL_C1, "formula").as_deref(),
            Some("B1*2")
        );

        let replayed = commit.patch().apply(&original).expect("replay color patch");
        assert_eq!(
            replayed.to_bytes().expect("replayed bytes"),
            commit.file().to_bytes().expect("committed bytes")
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(&replayed)
            .expect("inverse color patch");
        assert_eq!(restored.to_bytes().expect("restored bytes"), before);

        let repeated = commit
            .file()
            .set(
                "/sheet[1]/A1:D2",
                Props::new().with("font.color", "#FF0000"),
            )
            .expect("repeat color range");
        assert_eq!(
            repeated.to_bytes().expect("repeat bytes"),
            commit.file().to_bytes().expect("first bytes")
        );

        let cleared = commit
            .file()
            .set(CELL_A1, Props::new().with("font.color", PropValue::Null))
            .expect("clear font color");
        assert_eq!(cell_property(&cleared, CELL_A1, "font.color"), None);
        assert_eq!(cell_text(&cleared, CELL_A1).as_deref(), Some("colored"));
    }
}

#[test]
fn font_color_aliases_and_null_noops_are_atomic() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("original bytes");
        let alias = original
            .set(CELL_A1, Props::new().with("color", "red"))
            .expect("font color alias");
        assert_eq!(
            cell_property(&alias, CELL_A1, "font.color").as_deref(),
            Some("#FF0000")
        );

        let absent_clear = original
            .set(CELL_A1, Props::new().with("font.color", PropValue::Null))
            .expect("clear absent font color");
        assert_eq!(absent_clear.to_bytes().expect("absent clear bytes"), before);

        let duplicate = original
            .set(
                CELL_A1,
                Props::new().with("font.color", "red").with("color", "blue"),
            )
            .expect_err("font color aliases must share duplicate admission");
        assert_eq!(duplicate.code(), Code::InvalidArgument);
        assert_eq!(
            original.to_bytes().expect("unchanged duplicate owner"),
            before
        );
    }
}

#[test]
fn invalid_font_colors_and_context_transforms_are_rejected_atomically() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("original bytes");
        for value in [
            "not-a-color",
            "auto",
            "accent1+lumMod75",
            "accent7",
            "accent1+tint101",
            "accent1+tint=-101",
            "rgba(0,0,0,1.5)",
            "#12345",
        ] {
            let error = original
                .set(CELL_A1, Props::new().with("font.color", value))
                .expect_err("invalid SpreadsheetML color input");
            assert_eq!(error.code(), Code::InvalidValue, "color {value}");
            assert_eq!(error.property(), Some("font.color"), "color {value}");
            assert_eq!(original.to_bytes().expect("unchanged color owner"), before);
        }

        let typed_wrong_context = Color::parse("accent1+lumMod75", ColorContext::Drawing)
            .expect("Drawing color parses before spreadsheet admission");
        let error = original
            .set(
                CELL_A1,
                Props::new().with("font.color", typed_wrong_context),
            )
            .expect_err("Drawing transforms must be rejected by spreadsheet color admission");
        assert_eq!(error.code(), Code::InvalidValue);
        assert_eq!(error.property(), Some("font.color"));
        assert_eq!(
            original.to_bytes().expect("unchanged typed color owner"),
            before
        );
    }
}

#[test]
fn strike_and_underline_assignments_preserve_payloads_and_null_clears() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("original bytes");
        let updated = original
            .set(
                CELL_A1,
                props! {
                    "value" => "styled",
                    "strike" => true,
                    "underline" => "double",
                },
            )
            .expect("payload and strike/underline assignment");
        assert_eq!(cell_text(&updated, CELL_A1).as_deref(), Some("styled"));
        assert_eq!(
            cell_property(&updated, CELL_A1, "strike").as_deref(),
            Some("true")
        );
        assert_eq!(
            cell_property(&updated, CELL_A1, "underline").as_deref(),
            Some("double")
        );
        assert_eq!(cell_property(&updated, CELL_B1, "strike"), None);
        assert_eq!(original.to_bytes().expect("unchanged original"), before);

        let cleared = updated
            .set(
                CELL_A1,
                Props::new()
                    .with("strike", PropValue::Null)
                    .with("underline", PropValue::Null),
            )
            .expect("clear strike and underline");
        assert_eq!(cell_property(&cleared, CELL_A1, "strike"), None);
        assert_eq!(cell_property(&cleared, CELL_A1, "underline"), None);
        assert_eq!(cell_text(&cleared, CELL_A1).as_deref(), Some("styled"));
    }
}

#[test]
fn underline_tokens_are_exact_and_explicit_none_is_preserved() {
    let bytes = fixture();
    let tokens = [
        "single",
        "double",
        "singleAccounting",
        "doubleAccounting",
        "none",
    ];
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        for token in tokens {
            let updated = original
                .set(CELL_A1, Props::new().with("underline", token))
                .expect("exact native underline token");
            if token == "none" {
                // The ordinary cell projection deliberately omits an
                // explicitly authored no-underline marker.  The native
                // stylesheet still retains the exact token, distinguishing
                // it from a null clear.
                assert_eq!(cell_property(&updated, CELL_A1, "underline"), None);
                let styles = styles_bytes(&updated);
                let styles = std::str::from_utf8(&styles).expect("styles XML");
                assert!(styles.contains(r#"<u val="none"/>"#), "{styles}");
            } else {
                assert_eq!(
                    cell_property(&updated, CELL_A1, "underline").as_deref(),
                    Some(token),
                    "underline token {token}"
                );
            }
        }

        let changed = original
            .set(
                CELL_A1,
                props! { "strike" => true, "underline" => "double" },
            )
            .expect("set strike and underline");
        let repeated = changed
            .set(
                CELL_A1,
                props! { "strike" => true, "underline" => "double" },
            )
            .expect("repeat strike and underline")
            .to_bytes()
            .expect("repeated bytes");
        assert_eq!(repeated, changed.to_bytes().expect("changed bytes"));
        assert_eq!(
            original
                .set(
                    CELL_A1,
                    Props::new()
                        .with("strike", PropValue::Null)
                        .with("underline", PropValue::Null),
                )
                .expect("clear absent font properties")
                .to_bytes()
                .expect("clear absent bytes"),
            original.to_bytes().expect("original bytes")
        );
    }
}

#[test]
fn strike_and_underline_broadcast_over_ranges_and_replay_inverse() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("original bytes");
        let commit = original
            .edit(|edit| {
                edit.set(
                    "/sheet[1]/A1:D2",
                    props! { "strike" => true, "underline" => "singleAccounting" },
                )?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("broadcast strike and underline");
        assert!(commit.diagnostics().changed());
        for path in [CELL_A1, CELL_B1, CELL_C1, CELL_D1, "/sheet[1]/D2"] {
            assert_eq!(
                cell_property(commit.file(), path, "strike").as_deref(),
                Some("true"),
                "{path} strike"
            );
            assert_eq!(
                cell_property(commit.file(), path, "underline").as_deref(),
                Some("singleAccounting"),
                "{path} underline"
            );
        }
        let replayed = commit.patch().apply(&original).expect("replay font patch");
        assert_eq!(
            replayed.to_bytes().expect("replayed bytes"),
            commit.file().to_bytes().expect("committed bytes")
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(&replayed)
            .expect("inverse font patch");
        assert_eq!(restored.to_bytes().expect("restored bytes"), before);
    }
}

#[test]
fn strike_and_underline_aliases_use_the_same_native_assignments() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        for (input, expected) in [
            ("Single", Some("single")),
            ("dbl", Some("double")),
            ("singleacct", Some("singleAccounting")),
            ("DOUBLEACCT", Some("doubleAccounting")),
            ("true", Some("single")),
            ("yes", Some("single")),
            ("on", Some("single")),
            ("1", Some("single")),
            ("false", None),
            ("no", None),
            ("off", None),
            ("0", None),
        ] {
            let updated = original
                .set(CELL_A1, Props::new().with("font.underline", input))
                .expect("underline alias");
            assert_eq!(
                cell_property(&updated, CELL_A1, "underline").as_deref(),
                expected
            );
        }
        for (value, expected) in [(true, Some("single")), (false, None)] {
            let updated = original
                .set(CELL_A1, props! { "underline" => value })
                .expect("typed boolean underline alias");
            assert_eq!(
                cell_property(&updated, CELL_A1, "underline").as_deref(),
                expected
            );
        }
        for key in ["strikethrough", "font.strike", "font.strikethrough"] {
            let updated = original
                .set(CELL_A1, Props::new().with(key, true))
                .expect("strike alias");
            assert_eq!(
                cell_property(&updated, CELL_A1, "strike").as_deref(),
                Some("true")
            );
            let duplicate = original
                .set(CELL_A1, Props::new().with(key, true).with("strike", false))
                .expect_err("aliases share duplicate admission");
            assert_eq!(duplicate.code(), Code::InvalidArgument);
        }
    }
}

#[test]
fn invalid_and_duplicate_strike_underline_assignments_are_atomic() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("original bytes");

        let invalid_strike = original
            .set(
                CELL_A1,
                props! { "value" => "must not publish", "strike" => "sometimes" },
            )
            .expect_err("strike requires a boolean");
        assert_eq!(invalid_strike.code(), Code::InvalidValue);
        assert_eq!(invalid_strike.property(), Some("strike"));
        assert_eq!(original.to_bytes().expect("after strike refusal"), before);

        let invalid_underline = original
            .set(
                CELL_A1,
                props! { "value" => "must not publish", "underline" => "wavy" },
            )
            .expect_err("unknown underline tokens refuse");
        assert_eq!(invalid_underline.code(), Code::InvalidValue);
        assert_eq!(invalid_underline.property(), Some("underline"));
        assert_eq!(
            original.to_bytes().expect("after underline refusal"),
            before
        );

        let duplicate_strike = original
            .set(
                CELL_A1,
                Props::new()
                    .with("value", "must not publish")
                    .with("strike", true)
                    .with("STRIKE", false),
            )
            .expect_err("duplicate strike keys must refuse atomically");
        assert_eq!(duplicate_strike.code(), Code::InvalidArgument);
        assert_eq!(
            original.to_bytes().expect("after duplicate refusal"),
            before
        );

        let duplicate_underline = original
            .set(
                CELL_A1,
                Props::new()
                    .with("underline", "single")
                    .with("UNDERLINE", "double"),
            )
            .expect_err("duplicate underline keys must refuse atomically");
        assert_eq!(duplicate_underline.code(), Code::InvalidArgument);
        assert_eq!(
            original.to_bytes().expect("after duplicate underline"),
            before
        );
    }
}

#[test]
fn custom_native_styles_survive_source_eager_font_payload_range_and_inverse() {
    let bytes = custom_styles_fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("original bytes");
        let original_styles = styles_bytes(&original);
        let original_styles_text = std::str::from_utf8(&original_styles).expect("styles XML");
        assert!(original_styles_text.contains(r#"mc:Ignorable="x14ac""#));
        assert!(original_styles_text.contains(r#"x14ac:knownFonts="1""#));
        assert!(original_styles_text.contains(r#"x14ac:fontMarker="keep""#));

        let commit = original
            .edit(|transaction| {
                transaction.set(
                    CELL_A1,
                    props! { "value" => "custom styled", "font.bold" => true },
                )?;
                transaction.set("/sheet[1]/A1:D2", props! { "font.italic" => true })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("commit custom style edits");
        assert_eq!(
            cell_text(commit.file(), CELL_A1).as_deref(),
            Some("custom styled")
        );
        for path in [CELL_A1, CELL_B1, CELL_C1, CELL_D1, "/sheet[1]/D2"] {
            assert_eq!(
                cell_property(commit.file(), path, "font.italic").as_deref(),
                Some("true"),
                "{path} italic"
            );
        }
        assert_eq!(
            cell_property(commit.file(), CELL_A1, "font.bold").as_deref(),
            Some("true")
        );

        let edited_styles = styles_bytes(commit.file());
        let edited_styles_text = std::str::from_utf8(&edited_styles).expect("edited styles XML");
        assert!(edited_styles_text.contains(r#"mc:Ignorable="x14ac""#));
        assert!(edited_styles_text.contains(r#"x14ac:knownFonts="1""#));
        assert!(edited_styles_text.contains(r#"x14ac:fontMarker="keep""#));
        assert!(edited_styles_text.contains(r#"<x14ac:fontExtension marker="keep"/>"#));
        assert!(edited_styles_text.contains(r#"x14ac:xfMarker="keep""#));

        let reopened = open_owner(&commit.file().to_bytes().expect("edited bytes"), eager);
        assert_eq!(
            cell_property(&reopened, CELL_A1, "font.bold").as_deref(),
            Some("true")
        );
        assert_eq!(
            cell_property(&reopened, CELL_C1, "font.italic").as_deref(),
            Some("true")
        );

        let replayed = commit.patch().apply(&original).expect("forward replay");
        assert_eq!(
            replayed.to_bytes().expect("replayed bytes"),
            commit.file().to_bytes().expect("committed bytes")
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(&replayed)
            .expect("inverse replay");
        assert_eq!(restored.to_bytes().expect("restored bytes"), original_bytes);
        assert_eq!(
            original.to_bytes().expect("unchanged source"),
            original_bytes
        );
    }
}

#[test]
fn font_ranges_replay_inverse_and_repeat_without_changes() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("original bytes");
        let commit = original
            .edit(|edit| {
                edit.set(
                    "/sheet[1]/A1:D2",
                    props! { "font.bold" => true, "font.italic" => true },
                )?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("font range commit");
        for path in [CELL_A1, CELL_B1, CELL_C1, CELL_D1, "/sheet[1]/D2"] {
            assert_eq!(
                cell_property(commit.file(), path, "font.bold").as_deref(),
                Some("true")
            );
            assert_eq!(
                cell_property(commit.file(), path, "font.italic").as_deref(),
                Some("true")
            );
        }
        assert_eq!(
            cell_property(commit.file(), CELL_C1, "formula").as_deref(),
            Some("B1*2")
        );
        let replayed = commit.patch().apply(&original).expect("replay font range");
        let restored = commit
            .patch()
            .inverse()
            .apply(&replayed)
            .expect("undo font range");
        assert_eq!(restored.to_bytes().expect("restored bytes"), before);
        let repeat = commit
            .file()
            .edit(|edit| {
                edit.set(
                    "/sheet[1]/A1:D2",
                    props! { "font.bold" => true, "font.italic" => true },
                )?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("repeat same formatting");
        assert!(!repeat.diagnostics().changed());
        assert!(repeat.patch().is_noop());
        assert_eq!(
            repeat.file().to_bytes().expect("repeat bytes"),
            commit.file().to_bytes().expect("first bytes")
        );
    }
}

#[test]
fn invalid_font_assignment_rejects_combined_payload_atomically() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("original bytes");
        let error = original
            .set(
                CELL_A1,
                props! { "value" => "must not publish", "font.bold" => "invalid" },
            )
            .expect_err("invalid font flag");
        assert_eq!(error.code(), Code::InvalidValue);
        assert_eq!(error.property(), Some("font.bold"));
        assert_eq!(original.to_bytes().expect("original after refusal"), before);
        let empty_clear = original
            .edit(|edit| {
                edit.set(
                    "/sheet[1]/Z9",
                    Props::new().with("font.bold", PropValue::Null),
                )?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("clear absent font flag");
        assert!(!empty_clear.diagnostics().changed());
    }
}

#[test]
fn border_assignments_round_trip_source_and_eager() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original border owner");
        let updated = original
            .set(
                CELL_A1,
                props! {
                    "value" => "border payload",
                    "font.bold" => true,
                    "border.left" => "thin",
                    "border.left.color" => "#102030",
                    "border.right" => "double",
                    "border.right.color" => "#405060",
                    "border.top" => "medium",
                    "border.top.color" => "#708090",
                    "border.bottom" => "dotted",
                    "border.bottom.color" => "#A0B0C0",
                    "border.diagonal" => "hair",
                    "border.diagonal.color" => "#D0E0F0",
                    "border.diagonalUp" => true,
                    "border.diagonalDown" => true,
                },
            )
            .expect("assign every facade border property with a payload");

        assert_eq!(
            cell_text(&updated, CELL_A1).as_deref(),
            Some("border payload")
        );
        assert_eq!(
            cell_property(&updated, CELL_A1, "font.bold").as_deref(),
            Some("true")
        );
        for (key, expected) in [
            ("border.left", "thin"),
            ("border.left.color", "#102030"),
            ("border.right", "double"),
            ("border.right.color", "#405060"),
            ("border.top", "medium"),
            ("border.top.color", "#708090"),
            ("border.bottom", "dotted"),
            ("border.bottom.color", "#A0B0C0"),
            ("border.diagonal", "hair"),
            ("border.diagonal.color", "#D0E0F0"),
            ("border.diagonalUp", "true"),
            ("border.diagonalDown", "true"),
        ] {
            assert_eq!(
                cell_property(&updated, CELL_A1, key).as_deref(),
                Some(expected),
                "{key} readback"
            );
        }

        let styles = styles_bytes(&updated);
        let styles = std::str::from_utf8(&styles).expect("border styles XML");
        for fragment in [
            r#"<left style="thin"><color rgb="FF102030"/></left>"#,
            r#"<right style="double"><color rgb="FF405060"/></right>"#,
            r#"<top style="medium"><color rgb="FF708090"/></top>"#,
            r#"<bottom style="dotted"><color rgb="FFA0B0C0"/></bottom>"#,
            r#"<diagonal style="hair"><color rgb="FFD0E0F0"/></diagonal>"#,
            r#"diagonalUp="1" diagonalDown="1""#,
        ] {
            assert!(styles.contains(fragment), "missing {fragment} in {styles}");
        }

        let reopened = open_owner(&updated.to_bytes().expect("publish border edit"), eager);
        for (key, expected) in [
            ("border.left", "thin"),
            ("border.left.color", "#102030"),
            ("border.right", "double"),
            ("border.right.color", "#405060"),
            ("border.top", "medium"),
            ("border.top.color", "#708090"),
            ("border.bottom", "dotted"),
            ("border.bottom.color", "#A0B0C0"),
            ("border.diagonal", "hair"),
            ("border.diagonal.color", "#D0E0F0"),
            ("border.diagonalUp", "true"),
            ("border.diagonalDown", "true"),
        ] {
            assert_eq!(
                cell_property(&reopened, CELL_A1, key).as_deref(),
                Some(expected),
                "{key} after reopen"
            );
        }
        assert_eq!(
            original.to_bytes().expect("old owner remains unchanged"),
            before
        );
    }
}

#[test]
fn source_borders_keep_native_typed_values_and_opaque_markup() {
    let bytes = source_border_fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        for (key, expected) in [
            ("border.left", "thin"),
            ("border.left.color", "#102030"),
            ("border.right", "double"),
            ("border.top", "medium"),
            ("border.bottom", "dotted"),
            ("border.diagonal", "hair"),
            ("border.diagonalUp", "true"),
            ("border.diagonalDown", "true"),
        ] {
            assert_eq!(
                cell_property(&original, CELL_A1, key).as_deref(),
                Some(expected),
                "source {key} projection"
            );
        }
        for key in [
            "border.right.color",
            "border.bottom.color",
            "border.diagonal.color",
        ] {
            assert_eq!(
                cell_property(&original, CELL_A1, key),
                None,
                "OfficeCLI omits non-RGB source color for {key}"
            );
        }

        let native =
            litchi_xlsx::Package::from_bytes(&bytes).expect("open native source border workbook");
        let native_format = native
            .workbook_ref()
            .sheet(0usize)
            .expect("native sheet lookup")
            .expect("native worksheet")
            .style("A1")
            .expect("native style lookup")
            .expect("native style")
            .format()
            .expect("native format");
        let border = native_format.border().expect("native source border");
        let native_color = litchi_xlsx::style::cell_format::Color::rgb;
        let tint =
            litchi_xlsx::style::stylesheet::border::Tint::new(0.25).expect("valid source tint");
        assert_eq!(
            border.left,
            Some(
                litchi_xlsx::style::stylesheet::border::Side::new(
                    litchi_xlsx::style::stylesheet::border::Line::Thin,
                )
                .with_color(native_color(0x10, 0x20, 0x30)),
            )
        );
        assert_eq!(
            border.right,
            Some(
                litchi_xlsx::style::stylesheet::border::Side::new(
                    litchi_xlsx::style::stylesheet::border::Line::Double,
                )
                .with_color(litchi_xlsx::style::cell_format::Color::theme(4).with_tint(tint),),
            )
        );
        assert_eq!(
            border.top,
            Some(litchi_xlsx::style::stylesheet::border::Side::new(
                litchi_xlsx::style::stylesheet::border::Line::Medium,
            ))
        );
        assert_eq!(
            border.bottom,
            Some(
                litchi_xlsx::style::stylesheet::border::Side::new(
                    litchi_xlsx::style::stylesheet::border::Line::Dotted,
                )
                .with_color(litchi_xlsx::style::cell_format::Color::indexed(12)),
            )
        );
        let diagonal = border.diagonal.as_ref().expect("native diagonal");
        assert_eq!(
            diagonal.side(),
            Some(
                &litchi_xlsx::style::stylesheet::border::Side::new(
                    litchi_xlsx::style::stylesheet::border::Line::Hair,
                )
                .with_color(litchi_xlsx::style::cell_format::Color::auto()),
            )
        );
        assert_eq!(
            diagonal.dir(),
            Some(litchi_xlsx::style::stylesheet::border::Dir::Both)
        );
        assert_eq!(border.outline, Some(false));

        let edited = original
            .set(CELL_A1, Props::new().with("border.left", "thick"))
            .expect("edit known source border side");
        let styles = styles_bytes(&edited);
        let styles = std::str::from_utf8(&styles).expect("edited source border XML");
        assert!(styles.contains(r#"future:borderMarker="keep""#));
        assert!(styles.contains(r#"future:leftMarker="keep""#));
        assert!(styles.contains(r#"future:borderExtension marker="keep""#));
        assert!(styles.contains(r#"future:xfBorderMarker="keep""#));
        assert_eq!(
            cell_property(&edited, CELL_A1, "border.left").as_deref(),
            Some("thick")
        );
        assert_eq!(
            cell_property(&edited, CELL_A1, "border.right").as_deref(),
            Some("double")
        );
        let reopened = open_owner(
            &edited.to_bytes().expect("publish source border edit"),
            eager,
        );
        assert_eq!(
            cell_property(&reopened, CELL_A1, "border.left").as_deref(),
            Some("thick")
        );
        assert_eq!(
            cell_property(&reopened, CELL_A1, "border.right").as_deref(),
            Some("double")
        );
    }
}

#[test]
fn border_shorthands_and_style_aliases_have_stable_side_overrides() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let all_first = original
            .set(
                CELL_A1,
                Props::new()
                    .with("border.all", "thin")
                    .with("border.all.color", "#102030")
                    .with("border.left", "double")
                    .with("border.left.color", "#405060")
                    .with("border.top", "dotted")
                    .with("border.top.color", "#708090"),
            )
            .expect("border shorthand with side overrides");
        let side_first = original
            .set(
                CELL_A1,
                Props::new()
                    .with("border.left.style", "double")
                    .with("border.left.color", "#405060")
                    .with("border.top.style", "dotted")
                    .with("border.top.color", "#708090")
                    .with("border", "thin")
                    .with("border.color", "#102030"),
            )
            .expect("border all aliases with side overrides");

        for (name, file) in [("all-first", &all_first), ("side-first", &side_first)] {
            for (key, expected) in [
                ("border.left", "double"),
                ("border.left.color", "#405060"),
                ("border.right", "thin"),
                ("border.right.color", "#102030"),
                ("border.top", "dotted"),
                ("border.top.color", "#708090"),
                ("border.bottom", "thin"),
                ("border.bottom.color", "#102030"),
            ] {
                assert_eq!(
                    cell_property(file, CELL_A1, key).as_deref(),
                    Some(expected),
                    "{name} {key} readback"
                );
            }
        }

        let aliases = original
            .set(
                CELL_A1,
                Props::new()
                    .with("border.all.style", "medium")
                    .with("border.all.color", "#A0B0C0")
                    .with("border.left.style", "double")
                    .with("border.left.color", "#102030")
                    .with("border.right.style", "dashed")
                    .with("border.right.color", "#405060")
                    .with("border.top.style", "hair")
                    .with("border.top.color", "#708090")
                    .with("border.bottom.style", "dashDot")
                    .with("border.bottom.color", "#D0E0F0")
                    .with("border.diagonal.style", "thin")
                    .with("border.diagonal.color", "#010203"),
            )
            .expect("all border style and color aliases");
        for (key, expected) in [
            ("border.left", "double"),
            ("border.left.color", "#102030"),
            ("border.right", "dashed"),
            ("border.right.color", "#405060"),
            ("border.top", "hair"),
            ("border.top.color", "#708090"),
            ("border.bottom", "dashDot"),
            ("border.bottom.color", "#D0E0F0"),
            ("border.diagonal", "thin"),
            ("border.diagonal.color", "#010203"),
        ] {
            assert_eq!(
                cell_property(&aliases, CELL_A1, key).as_deref(),
                Some(expected),
                "style alias {key} readback"
            );
        }
    }
}

#[test]
fn border_color_policy_is_atomic_and_clear_line_preserves_source_color() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original
            .to_bytes()
            .expect("original absent-border policy bytes");
        for key in [
            "border.color",
            "border.all.color",
            "border.left.color",
            "border.right.color",
            "border.top.color",
            "border.bottom.color",
            "border.diagonal.color",
        ] {
            let error = original
                .set(
                    CELL_A1,
                    Props::new()
                        .with("value", "must not publish")
                        .with(key, "#102030"),
                )
                .expect_err("non-null color without a resulting line must refuse atomically");
            assert_eq!(error.code(), Code::InvalidValue, "{key}");
            assert_eq!(
                original
                    .to_bytes()
                    .expect("unchanged absent-border policy owner"),
                before,
                "invalid {key} must not publish the payload"
            );
        }

        for key in [
            "border.color",
            "border.all.color",
            "border.left.color",
            "border.right.color",
            "border.top.color",
            "border.bottom.color",
            "border.diagonal.color",
        ] {
            let null_clear = original
                .set(CELL_A1, Props::new().with(key, PropValue::Null))
                .expect("null color without a line is an accepted no-op");
            assert_eq!(
                null_clear
                    .to_bytes()
                    .expect("null absent-border color bytes"),
                before,
                "null {key} must not create a border"
            );
        }

        let source = open_owner(&source_border_fixture(), eager);
        let source_before = source
            .to_bytes()
            .expect("source border policy bytes before refusal");
        let invalid_companion = source
            .set(
                CELL_A1,
                Props::new()
                    .with("value", "must not publish")
                    .with("border.left", "none")
                    .with("border.left.color", "#405060"),
            )
            .expect_err("a color cannot accompany a cleared border line");
        assert_eq!(invalid_companion.code(), Code::InvalidValue);
        assert_eq!(
            source
                .to_bytes()
                .expect("unchanged source border policy owner"),
            source_before
        );

        let cleared_line = source
            .set(CELL_A1, Props::new().with("border.left", "none"))
            .expect("clear source border line while retaining its color markup");
        assert_eq!(cell_property(&cleared_line, CELL_A1, "border.left"), None);
        let cleared_styles_bytes = styles_bytes(&cleared_line);
        let cleared_styles =
            std::str::from_utf8(&cleared_styles_bytes).expect("cleared border XML");
        assert_eq!(
            cleared_styles.matches(r#"rgb="FF102030""#).count(),
            2,
            "clearing the line must retain the source color in the cloned border"
        );
        assert_eq!(
            cleared_styles
                .matches(r#"future:leftMarker="keep""#)
                .count(),
            2,
            "clearing the line must retain opaque side markup"
        );
        assert_eq!(
            cleared_styles.matches(r#"<left style="thin""#).count(),
            1,
            "the cloned side must no longer carry the source line"
        );
        let reopened = open_owner(
            &cleared_line
                .to_bytes()
                .expect("publish cleared source border line"),
            eager,
        );
        assert_eq!(cell_property(&reopened, CELL_A1, "border.left"), None);

        let cleared_color = source
            .set(
                CELL_A1,
                Props::new().with("border.left.color", PropValue::Null),
            )
            .expect("clear source border color without clearing its line");
        assert_eq!(
            cell_property(&cleared_color, CELL_A1, "border.left").as_deref(),
            Some("thin")
        );
        assert_eq!(
            cell_property(&cleared_color, CELL_A1, "border.left.color"),
            None
        );
        let cleared_color_styles_bytes = styles_bytes(&cleared_color);
        let cleared_color_styles =
            std::str::from_utf8(&cleared_color_styles_bytes).expect("cleared color XML");
        assert_eq!(
            cleared_color_styles.matches(r#"rgb="FF102030""#).count(),
            1,
            "null color removes only the authored source color"
        );
    }
}

#[test]
fn border_ranges_keep_payloads_and_ordered_batches_replay_inverse() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("original border bytes");
        let commit = original
            .edit(|edit| {
                edit.set(
                    CELL_A1,
                    props! {
                        "value" => "border companion",
                        "font.bold" => true,
                        "border.left" => "thin",
                        "border.left.color" => "#102030",
                    },
                )?;
                edit.set(
                    "/sheet[1]/A1:D2",
                    props! {
                        "border.top" => "medium",
                        "border.bottom" => "double",
                        "border.diagonal" => "hair",
                        "border.diagonalUp" => true,
                    },
                )?;
                edit.set(
                    CELL_B1,
                    props! {
                        "border.right" => "dotted",
                        "border.right.color" => "#405060",
                    },
                )?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("ordered border batch with companions");

        for path in [CELL_A1, CELL_B1, CELL_C1, CELL_D1, "/sheet[1]/D2"] {
            for (key, expected) in [
                ("border.top", "medium"),
                ("border.bottom", "double"),
                ("border.diagonal", "hair"),
                ("border.diagonalUp", "true"),
            ] {
                assert_eq!(
                    cell_property(commit.file(), path, key).as_deref(),
                    Some(expected),
                    "{path} {key}"
                );
            }
        }
        assert_eq!(
            cell_property(commit.file(), CELL_A1, "border.left").as_deref(),
            Some("thin")
        );
        assert_eq!(
            cell_property(commit.file(), CELL_A1, "border.left.color").as_deref(),
            Some("#102030")
        );
        assert_eq!(
            cell_property(commit.file(), CELL_B1, "border.right").as_deref(),
            Some("dotted")
        );
        assert_eq!(
            cell_property(commit.file(), CELL_B1, "border.right.color").as_deref(),
            Some("#405060")
        );
        assert_eq!(
            cell_property(commit.file(), CELL_A1, "font.bold").as_deref(),
            Some("true")
        );
        assert_eq!(
            cell_text(commit.file(), CELL_A1).as_deref(),
            Some("border companion")
        );
        assert_eq!(
            cell_property(commit.file(), CELL_C1, "formula").as_deref(),
            Some("B1*2")
        );

        let replayed = commit
            .patch()
            .apply(&original)
            .expect("replay border patch");
        assert_eq!(
            replayed.to_bytes().expect("replayed border bytes"),
            commit.file().to_bytes().expect("committed border bytes")
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(&replayed)
            .expect("inverse border patch");
        assert_eq!(restored.to_bytes().expect("restored border bytes"), before);

        let repeated = commit
            .file()
            .edit(|edit| {
                edit.set(
                    CELL_A1,
                    props! {
                        "value" => "border companion",
                        "font.bold" => true,
                        "border.left" => "thin",
                        "border.left.color" => "#102030",
                    },
                )?;
                edit.set(
                    "/sheet[1]/A1:D2",
                    props! {
                        "border.top" => "medium",
                        "border.bottom" => "double",
                        "border.diagonal" => "hair",
                        "border.diagonalUp" => true,
                    },
                )?;
                edit.set(
                    CELL_B1,
                    props! {
                        "border.right" => "dotted",
                        "border.right.color" => "#405060",
                    },
                )?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("repeat ordered border batch");
        assert!(repeated.patch().is_noop());
        assert_eq!(
            repeated.file().to_bytes().expect("repeated border bytes"),
            commit.file().to_bytes().expect("first border bytes")
        );

        let later_wins = original
            .edit(|edit| {
                edit.set(CELL_A1, Props::new().with("border.left", "thin"))?;
                edit.set(CELL_A1, Props::new().with("border.left", "double"))?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("ordered same-side border batch");
        assert_eq!(
            cell_property(later_wins.file(), CELL_A1, "border.left").as_deref(),
            Some("double")
        );
    }
}

#[test]
fn border_null_noops_and_clears_keep_distinct_semantics() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("original absent-border bytes");
        let absent_clear = original
            .set(
                CELL_A1,
                Props::new()
                    .with("border.left", PropValue::Null)
                    .with("border.left.color", PropValue::Null)
                    .with("border.right", PropValue::Null)
                    .with("border.right.color", PropValue::Null)
                    .with("border.top", PropValue::Null)
                    .with("border.top.color", PropValue::Null)
                    .with("border.bottom", PropValue::Null)
                    .with("border.bottom.color", PropValue::Null)
                    .with("border.diagonal", PropValue::Null)
                    .with("border.diagonal.color", PropValue::Null)
                    .with("border.diagonalUp", PropValue::Null)
                    .with("border.diagonalDown", PropValue::Null),
            )
            .expect("clearing absent borders is a no-op");
        assert_eq!(
            absent_clear.to_bytes().expect("absent border clear bytes"),
            before
        );

        let assigned = original
            .set(
                CELL_A1,
                props! {
                    "border.left" => "thin",
                    "border.left.color" => "#102030",
                    "border.right" => "double",
                    "border.right.color" => "#405060",
                    "border.top" => "medium",
                    "border.bottom" => "dotted",
                    "border.diagonal" => "hair",
                    "border.diagonalUp" => true,
                    "border.diagonalDown" => true,
                },
            )
            .expect("assign borders before repeat");
        let repeated = assigned
            .set(
                CELL_A1,
                props! {
                    "border.left" => "thin",
                    "border.left.color" => "#102030",
                    "border.right" => "double",
                    "border.right.color" => "#405060",
                    "border.top" => "medium",
                    "border.bottom" => "dotted",
                    "border.diagonal" => "hair",
                    "border.diagonalUp" => true,
                    "border.diagonalDown" => true,
                },
            )
            .expect("repeat exact borders");
        assert_eq!(
            repeated.to_bytes().expect("repeated border bytes"),
            assigned.to_bytes().expect("assigned border bytes")
        );

        let cleared = assigned
            .set(
                CELL_A1,
                Props::new()
                    .with("border.left", PropValue::Null)
                    .with("border.left.color", PropValue::Null)
                    .with("border.right", PropValue::Null)
                    .with("border.right.color", PropValue::Null)
                    .with("border.top", PropValue::Null)
                    .with("border.top.color", PropValue::Null)
                    .with("border.bottom", PropValue::Null)
                    .with("border.bottom.color", PropValue::Null)
                    .with("border.diagonal", PropValue::Null)
                    .with("border.diagonal.color", PropValue::Null)
                    .with("border.diagonalUp", PropValue::Null)
                    .with("border.diagonalDown", PropValue::Null),
            )
            .expect("clear assigned borders");
        for key in [
            "border.left",
            "border.left.color",
            "border.right",
            "border.right.color",
            "border.top",
            "border.top.color",
            "border.bottom",
            "border.bottom.color",
            "border.diagonal",
            "border.diagonal.color",
            "border.diagonalUp",
            "border.diagonalDown",
        ] {
            assert_eq!(cell_property(&cleared, CELL_A1, key), None, "{key} clear");
        }
        let repeated_clear = cleared
            .set(
                CELL_A1,
                Props::new()
                    .with("border.left", PropValue::Null)
                    .with("border.left.color", PropValue::Null)
                    .with("border.right", PropValue::Null)
                    .with("border.right.color", PropValue::Null)
                    .with("border.top", PropValue::Null)
                    .with("border.top.color", PropValue::Null)
                    .with("border.bottom", PropValue::Null)
                    .with("border.bottom.color", PropValue::Null)
                    .with("border.diagonal", PropValue::Null)
                    .with("border.diagonal.color", PropValue::Null)
                    .with("border.diagonalUp", PropValue::Null)
                    .with("border.diagonalDown", PropValue::Null),
            )
            .expect("repeat border clear");
        assert_eq!(
            repeated_clear.to_bytes().expect("repeated clear bytes"),
            cleared.to_bytes().expect("clear bytes")
        );
    }
}

#[test]
fn invalid_border_values_and_companions_are_atomic() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("original invalid-border bytes");
        for (key, value) in [
            ("border.left", "invalid"),
            ("border.right", "wavy"),
            ("border.top", "not-a-line"),
            ("border.bottom", ""),
            ("border.diagonal", "unknown"),
            ("border.left.color", "not-a-color"),
            ("border.right.color", "accent7"),
            ("border.top.color", "#12345"),
            ("border.bottom.color", "accent1+tint101"),
            ("border.diagonal.color", "#GG0000"),
            ("border.diagonalUp", "sometimes"),
            ("border.diagonalDown", "sometimes"),
        ] {
            let error = original
                .set(
                    CELL_A1,
                    Props::new()
                        .with("value", "must not publish")
                        .with(key, value),
                )
                .expect_err("invalid border value must refuse atomically");
            assert_eq!(error.code(), Code::InvalidValue, "{key}");
            assert_eq!(error.property(), Some(key), "{key}");
            assert_eq!(
                original.to_bytes().expect("unchanged invalid-border owner"),
                before,
                "invalid {key} must not publish payload"
            );
        }

        for properties in [
            Props::new()
                .with("border.left", "thin")
                .with("border.left", "double"),
            Props::new()
                .with("border.left.color", "#102030")
                .with("border.left.color", "#405060"),
            Props::new()
                .with("border.diagonalUp", true)
                .with("border.diagonalUp", false),
            Props::new()
                .with("border.right", "thin")
                .with("border.right.color", "not-a-color"),
        ] {
            let error = original
                .set(CELL_A1, properties)
                .expect_err("duplicate or invalid border companion must refuse");
            assert!(
                matches!(error.code(), Code::InvalidArgument | Code::InvalidValue),
                "unexpected border companion error: {error:?}"
            );
            assert_eq!(
                original
                    .to_bytes()
                    .expect("unchanged border companion owner"),
                before
            );
        }
    }
}

#[test]
fn scalar_value_formula_and_clear_work_for_source_and_eager_owners() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let updated = original
            .set(CELL_A1, props! { "value" => "changed" })
            .expect("set text value")
            .set(CELL_B1, props! { "value" => 7 })
            .expect("set numeric value")
            .set(CELL_D1, props! { "value" => false })
            .expect("set boolean value")
            .set(CELL_C1, props! { "formula" => "B1*2+1" })
            .expect("set formula");

        assert_eq!(cell_text(&updated, CELL_A1).as_deref(), Some("changed"));
        assert_eq!(cell_text(&updated, CELL_B1).as_deref(), Some("7"));
        assert_eq!(cell_text(&updated, CELL_D1).as_deref(), Some("FALSE"));
        assert_eq!(
            cell_property(&updated, CELL_C1, "formula").as_deref(),
            Some("B1*2+1")
        );
        assert_eq!(cell_property(&updated, CELL_C1, "cachedValue"), None);
        assert_eq!(
            cell_property(&updated, CELL_C1, "evaluated").as_deref(),
            Some("false")
        );

        let cleared = updated
            .set(CELL_C1, props! { "clear" => true })
            .expect("clear formula");
        assert_eq!(cell_text(&cleared, CELL_C1).as_deref(), Some(""));
        assert_eq!(cell_property(&cleared, CELL_C1, "formula"), None);
        assert_eq!(
            cell_property(&cleared, CELL_C1, "empty").as_deref(),
            Some("true")
        );

        assert_eq!(
            original.to_bytes().expect("republish original owner"),
            original_bytes,
            "the source/eager input was mutated by a detached cell edit"
        );
    }
}

#[test]
fn exact_cell_noops_preserve_bytes_and_old_owner() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");

        let same_value = original
            .set(CELL_A1, props! { "value" => "seed" })
            .expect("repeat text value");
        assert_eq!(
            same_value.to_bytes().expect("publish value no-op"),
            original_bytes,
            "same text value changed the {eager:?} owner"
        );

        let same_formula = original
            .set(CELL_C1, props! { "formula" => "B1*2" })
            .expect("repeat formula");
        assert_eq!(
            same_formula.to_bytes().expect("publish formula no-op"),
            original_bytes,
            "same formula changed the {eager:?} owner"
        );
        assert_eq!(
            original.to_bytes().expect("republish old owner"),
            original_bytes,
            "the old owner changed after a no-op"
        );
    }
}

#[test]
fn a_round_trip_transaction_reuses_the_exact_input_snapshot() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let commit = original
            .edit(|tx| {
                tx.set(CELL_A1, props! { "value" => "temporary" })?;
                tx.set(CELL_A1, props! { "value" => "seed" })?;
                Ok(())
            })
            .expect("round-trip edit");
        assert!(!commit.diagnostics().changed());
        assert!(commit.patch().is_noop());
        assert_eq!(commit.file().to_bytes().unwrap(), bytes);
        assert_eq!(
            commit.patch().apply(&original).unwrap().to_bytes().unwrap(),
            bytes
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .unwrap()
                .to_bytes()
                .unwrap(),
            bytes
        );
    }
}

#[test]
fn null_value_and_formula_are_explicit_clears() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let cleared = original
            .set(CELL_A1, Props::new().with("value", PropValue::Null))
            .expect("null value clears the cell")
            .set(CELL_C1, Props::new().with("formula", PropValue::Null))
            .expect("null formula clears the cell");

        assert_eq!(cell_text(&cleared, CELL_A1).as_deref(), Some(""));
        assert_eq!(
            cell_property(&cleared, CELL_A1, "empty").as_deref(),
            Some("true")
        );
        assert_eq!(cell_property(&cleared, CELL_C1, "formula"), None);
        assert_eq!(
            cell_property(&cleared, CELL_C1, "empty").as_deref(),
            Some("true")
        );
        assert_eq!(
            original.to_bytes().expect("republish original owner"),
            original_bytes,
            "null clears mutated the detached source/eager owner"
        );
    }
}

#[test]
fn value_strings_remain_text_when_they_resemble_other_scalar_kinds() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let updated = original
            .set(CELL_A1, props! { "value" => "123" })
            .expect("numeric-looking text")
            .set(CELL_B1, props! { "value" => "true" })
            .expect("boolean-looking text")
            .set(CELL_D1, props! { "value" => "=B1*2" })
            .expect("formula-looking text");

        assert_eq!(cell_text(&updated, CELL_A1).as_deref(), Some("123"));
        assert_eq!(cell_text(&updated, CELL_B1).as_deref(), Some("true"));
        assert_eq!(cell_text(&updated, CELL_D1).as_deref(), Some("=B1*2"));
        assert_eq!(cell_property(&updated, CELL_A1, "formula"), None);
        assert_eq!(cell_property(&updated, CELL_B1, "formula"), None);
        assert_eq!(cell_property(&updated, CELL_D1, "formula"), None);
        assert_eq!(
            cell_property(&updated, CELL_A1, "type").as_deref(),
            Some("SharedString")
        );
        assert_eq!(
            cell_property(&updated, CELL_B1, "type").as_deref(),
            Some("SharedString")
        );
        assert_eq!(
            cell_property(&updated, CELL_D1, "type").as_deref(),
            Some("SharedString")
        );
    }
}

#[test]
fn invalid_xml_controls_are_rejected_before_xlsx_cell_mutation() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        for (path, property) in [(CELL_A1, "value"), (CELL_C1, "formula")] {
            let error = original
                .set(path, Props::new().with(property, "bad\u{1}text"))
                .expect_err("XML control characters must be refused");
            assert_eq!(error.code(), Code::InvalidValue);
            assert_eq!(error.property(), Some(property));
        }
        assert_eq!(
            original.to_bytes().expect("publish unchanged owner"),
            original_bytes,
            "invalid XML text escaped into a partially edited owner"
        );
    }
}

#[test]
fn values_outside_the_i64_scalar_range_are_rejected() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let error = original
            .set(
                CELL_B1,
                Props::new().with("value", PropValue::Unsigned(i64::MAX as u64 + 1)),
            )
            .expect_err("values outside the scalar i64 range must be refused");
        assert_eq!(error.code(), Code::InvalidValue);
        assert_eq!(error.property(), Some("value"));
        assert_eq!(
            original.to_bytes().expect("publish unchanged owner"),
            original_bytes
        );
    }
}

#[test]
fn xlsx_batches_are_atomic_or_best_effort_by_mode() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let atomic_error = original
            .batch(
                [
                    batch_set(CELL_A1, Props::new().with("value", "atomic")),
                    batch_set(CELL_B1, Props::new().with("unknownCellProperty", true)),
                    batch_set(CELL_D1, Props::new().with("value", false)),
                ],
                BatchMode::Atomic,
            )
            .expect_err("atomic batches must reject the whole candidate");
        assert_eq!(atomic_error.item(), Some(1));
        assert_eq!(
            original.to_bytes().expect("publish unchanged atomic owner"),
            original_bytes,
            "atomic failure published a prefix or suffix"
        );

        let best_effort = original
            .batch(
                [
                    batch_set(CELL_A1, Props::new().with("value", "best")),
                    batch_set(CELL_B1, Props::new().with("unknownCellProperty", true)),
                    batch_set(CELL_D1, Props::new().with("value", false)),
                ],
                BatchMode::BestEffort,
            )
            .expect("best-effort batches retain successful items");
        assert_eq!(best_effort.value().len(), 3);
        assert!(matches!(
            &best_effort.value()[0],
            ItemResult::Applied(result) if result.changed
        ));
        assert!(matches!(
            &best_effort.value()[1],
            ItemResult::Failed(error) if error.item() == Some(1)
        ));
        assert!(matches!(
            &best_effort.value()[2],
            ItemResult::Applied(result) if result.changed
        ));
        assert_eq!(
            cell_text(best_effort.file(), CELL_A1).as_deref(),
            Some("best")
        );
        assert_eq!(
            cell_text(best_effort.file(), CELL_D1).as_deref(),
            Some("FALSE")
        );

        let restored = best_effort
            .patch()
            .inverse()
            .apply(best_effort.file())
            .expect("best-effort inverse restores the batch input");
        assert_eq!(
            restored.to_bytes().expect("publish restored batch owner"),
            original_bytes
        );
    }
}

#[test]
fn xlsx_snapshot_limits_roll_back_a_failed_transaction() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let probe = original
            .edit(|transaction| {
                transaction.set(CELL_A1, props! { "value" => "snapshot-charge" })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("measure a valid XLSX snapshot");
        let charge = probe.diagnostics().snapshot_bytes();
        assert!(charge > 0, "a changed XLSX patch must retain native bytes");
        let limits = SemanticLimits::builder()
            .max_snapshot_bytes(charge - 1)
            .expect("snapshot limit is representable")
            .build()
            .expect("build snapshot limit");

        let error = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |transaction| {
                    transaction.set(CELL_A1, props! { "value" => "snapshot-charge" })?;
                    Ok::<_, litchi_ops::Error>(())
                },
            )
            .expect_err("snapshot limit must reject before publication");
        assert_eq!(error.code(), Code::LimitExceeded);
        assert_eq!(
            original
                .to_bytes()
                .expect("publish unchanged limited owner"),
            original_bytes,
            "snapshot-limit failure escaped a staged XLSX candidate"
        );
    }
}

#[test]
fn xlsx_transaction_retains_native_patch_and_inverse_endpoints() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let commit = original
            .edit(|transaction| {
                transaction.set(CELL_A1, props! { "value" => "staged" })?;
                transaction.set(CELL_B1, props! { "value" => 9 })?;
                transaction.set(CELL_C1, props! { "clear" => true })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("commit XLSX transaction");

        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), 3);
        assert!(commit.diagnostics().commits() > 0);

        let forward = commit
            .patch()
            .apply(&original)
            .expect("apply forward XLSX patch");
        assert_eq!(
            forward.to_bytes().expect("publish forward endpoint"),
            commit.file().to_bytes().expect("publish committed file")
        );

        let restored = commit
            .patch()
            .inverse()
            .apply(&forward)
            .expect("apply inverse XLSX patch");
        assert_eq!(
            restored.to_bytes().expect("publish restored endpoint"),
            original_bytes
        );

        let reopened = open_owner(&original_bytes, eager);
        assert_eq!(
            commit
                .patch()
                .apply(&reopened)
                .expect("a byte-identical copy takes the patch")
                .to_bytes()
                .expect("publish reopened replay"),
            commit.file().to_bytes().expect("publish committed file")
        );
    }
}

#[test]
fn failed_later_xlsx_set_poisons_transaction_and_discards_candidate() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");
        let error = original
            .edit::<()>(|transaction| {
                transaction.set(CELL_A1, props! { "value" => "staged" })?;
                let first = transaction
                    .set(CELL_B1, props! { "unknownCellProperty" => true })
                    .expect_err("unknown property must fail");
                assert_eq!(first.item(), Some(1));
                let second = transaction
                    .set(CELL_D1, props! { "value" => false })
                    .expect_err("a poisoned transaction must reject later work");
                assert_eq!(second.item(), Some(1));
                assert_eq!(second.code(), first.code());
                Ok(())
            })
            .expect_err("poisoned closure must not publish");
        assert_eq!(error.item(), Some(1));
        assert_eq!(
            original.to_bytes().expect("publish unchanged owner"),
            original_bytes,
            "failed XLSX transaction escaped a staged candidate"
        );
    }
}

#[test]
fn malformed_xlsx_cell_props_are_rejected_without_partial_writes() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("publish original owner");

        assert_value_refusal(
            original
                .set(CELL_C1, props! { "formula" => "=B1" })
                .expect_err("leading equals must be refused"),
        );
        assert_value_refusal(
            original
                .set(CELL_C1, props! { "formula" => "" })
                .expect_err("empty formula must be refused"),
        );
        assert_value_refusal(
            original
                .set(CELL_A1, props! { "value" => "one", "formula" => "B1" })
                .expect_err("value/formula conflict must be refused"),
        );
        assert_value_refusal(
            original
                .set(CELL_A1, props! { "clear" => true, "value" => "one" })
                .expect_err("clear/value conflict must be refused"),
        );
        assert_value_refusal(
            original
                .set(CELL_A1, props! { "value" => "one", "VALUE" => "two" })
                .expect_err("duplicate property names must be refused"),
        );
        assert_eq!(
            original.to_bytes().expect("publish unchanged owner"),
            original_bytes
        );
    }
}

#[test]
fn long_tail_font_children_round_trip_source_and_eager() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original owner");
        let updated = original
            .set(
                CELL_A1,
                props! {
                    "value" => "long-tail",
                    "font.outline" => true,
                    "font.shadow" => true,
                    "font.condense" => true,
                    "font.extend" => true,
                    "font.charset" => 128,
                    "font.family" => 2,
                    "font.scheme" => "minor",
                },
            )
            .expect("long-tail font assignment");

        assert_eq!(cell_text(&updated, CELL_A1).as_deref(), Some("long-tail"));
        for (key, expected) in [
            ("font.outline", "true"),
            ("font.shadow", "true"),
            ("font.condense", "true"),
            ("font.extend", "true"),
            ("font.charset", "128"),
            ("font.family", "2"),
            ("font.scheme", "minor"),
        ] {
            assert_eq!(
                cell_property(&updated, CELL_A1, key).as_deref(),
                Some(expected),
                "{key} readback"
            );
        }

        let styles = styles_bytes(&updated);
        let styles = std::str::from_utf8(&styles).expect("styles XML");
        for child in ["<outline", "<shadow", "<condense", "<extend"] {
            assert!(styles.contains(child), "missing {child} in {styles}");
        }
        for child in [
            r#"<charset val="128"/>"#,
            r#"<family val="2"/>"#,
            r#"<scheme val="minor"/>"#,
        ] {
            assert!(styles.contains(child), "missing {child} in {styles}");
        }

        let reopened = open_owner(&updated.to_bytes().expect("publish long-tail edit"), eager);
        for (key, expected) in [
            ("font.outline", "true"),
            ("font.shadow", "true"),
            ("font.condense", "true"),
            ("font.extend", "true"),
            ("font.charset", "128"),
            ("font.family", "2"),
            ("font.scheme", "minor"),
        ] {
            assert_eq!(
                cell_property(&reopened, CELL_A1, key).as_deref(),
                Some(expected),
                "{key} after reopen"
            );
        }
        assert_eq!(
            original.to_bytes().expect("old owner remains unchanged"),
            before
        );
    }
}

#[test]
fn source_authored_long_tail_font_values_are_read_and_queryable() {
    let bytes = source_long_tail_font_fixture();

    for eager in [false, true] {
        let file = open_owner(&bytes, eager);
        for (key, expected) in [
            ("font.outline", "false"),
            ("font.shadow", "false"),
            ("font.condense", "false"),
            ("font.extend", "false"),
            ("font.charset", "0"),
            ("font.family", "0"),
            ("font.scheme", "none"),
        ] {
            assert_eq!(
                cell_property(&file, CELL_A1, key).as_deref(),
                Some(expected),
                "{key} source readback"
            );
        }

        for selector in [
            "cell[font.outline=false]",
            "cell[font.charset=0]",
            "cell[font.scheme=none]",
        ] {
            let matches = file
                .query(selector)
                .expect("compile long-tail font query")
                .collect::<litchi_ops::Result<Vec<_>>>()
                .expect("evaluate long-tail font query");
            assert_eq!(matches.len(), 1, "{selector}");
            assert_eq!(matches[0].path().to_string(), CELL_A1, "{selector}");
        }

        let unstyled = file.get(CELL_B1).expect("resolve unstyled cell");
        assert_eq!(
            unstyled
                .view()
                .format()
                .expect("unstyled format")
                .find(|property| property.key() == "font.outline"),
            None
        );
    }
}

#[test]
fn long_tail_font_ranges_keep_payloads_and_replay_inverse() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original owner");
        let commit = original
            .edit(|edit| {
                edit.set(
                    CELL_A1,
                    props! {
                        "value" => "range companion",
                        "font.scheme" => "major",
                    },
                )?;
                edit.set(
                    "/sheet[1]/A1:D2",
                    props! {
                        "font.outline" => true,
                        "font.shadow" => true,
                        "font.condense" => true,
                        "font.extend" => true,
                        "font.charset" => 1,
                        "font.family" => 3,
                        "font.scheme" => "minor",
                    },
                )?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("long-tail range assignment");

        for path in [CELL_A1, CELL_B1, CELL_C1, CELL_D1, "/sheet[1]/D2"] {
            for (key, expected) in [
                ("font.outline", "true"),
                ("font.shadow", "true"),
                ("font.condense", "true"),
                ("font.extend", "true"),
                ("font.charset", "1"),
                ("font.family", "3"),
                ("font.scheme", "minor"),
            ] {
                assert_eq!(
                    cell_property(commit.file(), path, key).as_deref(),
                    Some(expected),
                    "{path} {key}"
                );
            }
        }
        assert_eq!(
            cell_text(commit.file(), CELL_A1).as_deref(),
            Some("range companion")
        );
        assert_eq!(
            cell_property(commit.file(), CELL_C1, "formula").as_deref(),
            Some("B1*2")
        );

        let replayed = commit
            .patch()
            .apply(&original)
            .expect("replay long-tail patch");
        assert_eq!(
            replayed.to_bytes().expect("replayed bytes"),
            commit.file().to_bytes().expect("committed bytes")
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(&replayed)
            .expect("inverse long-tail patch");
        assert_eq!(restored.to_bytes().expect("restored bytes"), before);

        let repeated = commit
            .file()
            .set(
                "/sheet[1]/A1:D2",
                props! {
                    "font.outline" => true,
                    "font.shadow" => true,
                    "font.condense" => true,
                    "font.extend" => true,
                    "font.charset" => 1,
                    "font.family" => 3,
                    "font.scheme" => "minor",
                },
            )
            .expect("repeat long-tail range");
        assert_eq!(
            repeated.to_bytes().expect("repeat bytes"),
            commit.file().to_bytes().expect("first bytes")
        );
    }
}

#[test]
fn long_tail_font_false_none_and_null_clears_keep_distinct_semantics() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original owner");

        let absent_clear = original
            .set(
                CELL_A1,
                Props::new()
                    .with("font.outline", PropValue::Null)
                    .with("font.shadow", PropValue::Null)
                    .with("font.condense", PropValue::Null)
                    .with("font.extend", PropValue::Null)
                    .with("font.charset", PropValue::Null)
                    .with("font.family", PropValue::Null)
                    .with("font.scheme", PropValue::Null),
            )
            .expect("clear absent long-tail font children");
        assert_eq!(absent_clear.to_bytes().expect("absent clear bytes"), before);

        let assigned = original
            .set(
                CELL_A1,
                props! {
                    "font.outline" => true,
                    "font.shadow" => true,
                    "font.condense" => true,
                    "font.extend" => true,
                    "font.charset" => 128,
                    "font.family" => 2,
                    "font.scheme" => "minor",
                },
            )
            .expect("assign long-tail font children");
        let reset = assigned
            .set(
                CELL_A1,
                props! {
                    "font.outline" => false,
                    "font.shadow" => false,
                    "font.condense" => false,
                    "font.extend" => false,
                    "font.charset" => 0,
                    "font.family" => 0,
                    "font.scheme" => "none",
                },
            )
            .expect("author explicit false, zero and none values");
        for (key, expected) in [
            ("font.outline", "false"),
            ("font.shadow", "false"),
            ("font.condense", "false"),
            ("font.extend", "false"),
            ("font.charset", "0"),
            ("font.family", "0"),
            ("font.scheme", "none"),
        ] {
            assert_eq!(
                cell_property(&reset, CELL_A1, key).as_deref(),
                Some(expected),
                "{key} explicit reset"
            );
        }
        let repeated = reset
            .set(
                CELL_A1,
                props! {
                    "font.outline" => false,
                    "font.shadow" => false,
                    "font.condense" => false,
                    "font.extend" => false,
                    "font.charset" => 0,
                    "font.family" => 0,
                    "font.scheme" => "none",
                },
            )
            .expect("repeat explicit reset");
        assert_eq!(
            repeated.to_bytes().expect("repeated reset bytes"),
            reset.to_bytes().expect("reset bytes")
        );

        let cleared = reset
            .set(
                CELL_A1,
                Props::new()
                    .with("font.outline", PropValue::Null)
                    .with("font.shadow", PropValue::Null)
                    .with("font.condense", PropValue::Null)
                    .with("font.extend", PropValue::Null)
                    .with("font.charset", PropValue::Null)
                    .with("font.family", PropValue::Null)
                    .with("font.scheme", PropValue::Null),
            )
            .expect("null clears long-tail font children");
        for key in [
            "font.outline",
            "font.shadow",
            "font.condense",
            "font.extend",
            "font.charset",
            "font.family",
            "font.scheme",
        ] {
            assert_eq!(cell_property(&cleared, CELL_A1, key), None, "{key} clear");
        }
    }
}

#[test]
fn invalid_long_tail_font_values_and_duplicate_keys_are_atomic() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original owner");

        for key in [
            "font.outline",
            "font.shadow",
            "font.condense",
            "font.extend",
        ] {
            let error = original
                .set(CELL_A1, Props::new().with(key, "sometimes"))
                .expect_err("font flag must be boolean");
            assert_eq!(error.code(), Code::InvalidValue, "{key}");
            assert_eq!(error.property(), Some(key), "{key}");
            assert_eq!(original.to_bytes().expect("unchanged flag owner"), before);
        }

        for (key, value) in [
            ("font.charset", -1_i32),
            ("font.charset", 256_i32),
            ("font.family", -1_i32),
            ("font.family", 6_i32),
        ] {
            let error = original
                .set(CELL_A1, Props::new().with(key, value))
                .expect_err("font integer must stay in its XML domain");
            assert_eq!(error.code(), Code::InvalidValue, "{key}={value}");
            assert_eq!(error.property(), Some(key), "{key}={value}");
            assert_eq!(
                original.to_bytes().expect("unchanged integer owner"),
                before
            );
        }

        let error = original
            .set(CELL_A1, Props::new().with("font.scheme", "body"))
            .expect_err("font scheme has a closed token domain");
        assert_eq!(error.code(), Code::InvalidValue);
        assert_eq!(error.property(), Some("font.scheme"));
        assert_eq!(original.to_bytes().expect("unchanged scheme owner"), before);

        for key in [
            "font.outline",
            "font.shadow",
            "font.condense",
            "font.extend",
            "font.charset",
            "font.family",
            "font.scheme",
        ] {
            let properties = match key {
                "font.charset" | "font.family" => Props::new().with(key, 1_i32).with(key, 2_i32),
                "font.scheme" => Props::new().with(key, "major").with(key, "minor"),
                _ => Props::new().with(key, true).with(key, false),
            };
            let error = original
                .set(CELL_A1, properties)
                .expect_err("duplicate long-tail keys must refuse");
            assert_eq!(error.code(), Code::InvalidArgument, "{key}");
            assert_eq!(
                original.to_bytes().expect("unchanged duplicate owner"),
                before
            );
        }

        for key in [
            "outline", "shadow", "condense", "extend", "charset", "family", "scheme",
        ] {
            let error = original
                .set(CELL_A1, Props::new().with(key, true))
                .expect_err("long-tail fields have no bare OfficeCLI aliases");
            assert_eq!(error.code(), Code::UnsupportedProperty, "{key}");
            assert_eq!(
                original.to_bytes().expect("unchanged bare alias owner"),
                before
            );
        }
    }
}

#[test]
fn alignment_assignments_round_trip_source_and_eager() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original owner");
        let updated = original
            .set(
                CELL_A1,
                props! {
                    "value" => "aligned payload",
                    "font.bold" => true,
                    "alignment.horizontal" => "center",
                    "alignment.vertical" => "bottom",
                    "alignment.wrapText" => true,
                    "alignment.textRotation" => 90,
                    "alignment.indent" => 3,
                    "alignment.shrinkToFit" => true,
                    "alignment.readingOrder" => "rtl",
                },
            )
            .expect("assign all alignment children with a payload");

        assert_eq!(
            cell_text(&updated, CELL_A1).as_deref(),
            Some("aligned payload")
        );
        assert_eq!(
            cell_property(&updated, CELL_A1, "font.bold").as_deref(),
            Some("true")
        );
        for (key, expected) in [
            ("alignment.horizontal", "center"),
            ("alignment.vertical", "bottom"),
            ("alignment.wrapText", "true"),
            ("alignment.textRotation", "90"),
            ("alignment.indent", "3"),
            ("alignment.shrinkToFit", "true"),
            ("alignment.readingOrder", "rtl"),
        ] {
            assert_eq!(
                cell_property(&updated, CELL_A1, key).as_deref(),
                Some(expected),
                "{key} readback"
            );
        }

        let styles = styles_bytes(&updated);
        let styles = std::str::from_utf8(&styles).expect("styles XML");
        for fragment in [
            r#"horizontal="center""#,
            r#"vertical="bottom""#,
            r#"wrapText="1""#,
            r#"textRotation="90""#,
            r#"indent="3""#,
            r#"shrinkToFit="1""#,
            r#"readingOrder="2""#,
        ] {
            assert!(styles.contains(fragment), "missing {fragment} in {styles}");
        }

        let reopened = open_owner(&updated.to_bytes().expect("publish alignment edit"), eager);
        for (key, expected) in [
            ("alignment.horizontal", "center"),
            ("alignment.vertical", "bottom"),
            ("alignment.wrapText", "true"),
            ("alignment.textRotation", "90"),
            ("alignment.indent", "3"),
            ("alignment.shrinkToFit", "true"),
            ("alignment.readingOrder", "rtl"),
        ] {
            assert_eq!(
                cell_property(&reopened, CELL_A1, key).as_deref(),
                Some(expected),
                "{key} after reopen"
            );
        }
        assert_eq!(
            original.to_bytes().expect("old owner remains unchanged"),
            before
        );
    }
}

#[test]
fn alignment_ranges_keep_payloads_and_replay_inverse() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original owner");
        let commit = original
            .edit(|edit| {
                edit.set(
                    CELL_A1,
                    props! {
                        "value" => "alignment companion",
                        "font.bold" => true,
                        "alignment.horizontal" => "center",
                    },
                )?;
                edit.set(
                    "/sheet[1]/A1:D2",
                    props! {
                        "alignment.vertical" => "center",
                        "alignment.wrapText" => true,
                        "alignment.textRotation" => 45,
                        "alignment.indent" => 2,
                        "alignment.shrinkToFit" => true,
                        "alignment.readingOrder" => "ltr",
                    },
                )?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("commit alignment range with companions");

        for path in [CELL_A1, CELL_B1, CELL_C1, CELL_D1, "/sheet[1]/D2"] {
            for (key, expected) in [
                ("alignment.vertical", "center"),
                ("alignment.wrapText", "true"),
                ("alignment.textRotation", "45"),
                ("alignment.indent", "2"),
                ("alignment.shrinkToFit", "true"),
                ("alignment.readingOrder", "ltr"),
            ] {
                assert_eq!(
                    cell_property(commit.file(), path, key).as_deref(),
                    Some(expected),
                    "{path} {key}"
                );
            }
        }
        assert_eq!(
            cell_property(commit.file(), CELL_A1, "alignment.horizontal").as_deref(),
            Some("center")
        );
        assert_eq!(
            cell_property(commit.file(), CELL_A1, "font.bold").as_deref(),
            Some("true")
        );
        assert_eq!(
            cell_text(commit.file(), CELL_A1).as_deref(),
            Some("alignment companion")
        );
        assert_eq!(
            cell_property(commit.file(), CELL_C1, "formula").as_deref(),
            Some("B1*2")
        );

        let replayed = commit
            .patch()
            .apply(&original)
            .expect("replay alignment range patch");
        assert_eq!(
            replayed.to_bytes().expect("replayed alignment bytes"),
            commit.file().to_bytes().expect("committed alignment bytes")
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(&replayed)
            .expect("inverse alignment range patch");
        assert_eq!(
            restored.to_bytes().expect("restored alignment bytes"),
            before
        );

        let repeated = commit
            .file()
            .set(
                "/sheet[1]/A1:D2",
                props! {
                    "alignment.vertical" => "center",
                    "alignment.wrapText" => true,
                    "alignment.textRotation" => 45,
                    "alignment.indent" => 2,
                    "alignment.shrinkToFit" => true,
                    "alignment.readingOrder" => "ltr",
                },
            )
            .expect("repeat alignment range");
        assert_eq!(
            repeated.to_bytes().expect("repeat alignment bytes"),
            commit.file().to_bytes().expect("first alignment bytes")
        );
    }
}

#[test]
fn alignment_aliases_round_trip_and_alias_duplicates_are_atomic() {
    let bytes = fixture();
    let aliases = [
        (
            "align",
            "alignment.horizontal",
            PropValue::from("centerAcrossSelection"),
            "centerContinuous",
        ),
        (
            "readingorder",
            "alignment.readingOrder",
            PropValue::from(2_i32),
            "rtl",
        ),
        (
            "dir",
            "alignment.readingOrder",
            PropValue::from("leftToRight"),
            "ltr",
        ),
        (
            "halign",
            "alignment.horizontal",
            PropValue::from("right"),
            "right",
        ),
        (
            "align",
            "alignment.horizontal",
            PropValue::from("left"),
            "left",
        ),
        (
            "valign",
            "alignment.vertical",
            PropValue::from("bottom"),
            "bottom",
        ),
        ("wrap", "alignment.wrapText", PropValue::from(true), "true"),
        (
            "wraptext",
            "alignment.wrapText",
            PropValue::from(true),
            "true",
        ),
        (
            "rotation",
            "alignment.textRotation",
            PropValue::from(30_i32),
            "30",
        ),
        (
            "textrotation",
            "alignment.textRotation",
            PropValue::from(31_i32),
            "31",
        ),
        (
            "alignment.rotation",
            "alignment.textRotation",
            PropValue::from(32_i32),
            "32",
        ),
        ("indent", "alignment.indent", PropValue::from(4_i32), "4"),
        (
            "shrinktofit",
            "alignment.shrinkToFit",
            PropValue::from(true),
            "true",
        ),
        (
            "shrink",
            "alignment.shrinkToFit",
            PropValue::from(true),
            "true",
        ),
        (
            "alignment.shrink",
            "alignment.shrinkToFit",
            PropValue::from(true),
            "true",
        ),
        (
            "readingorder",
            "alignment.readingOrder",
            PropValue::from("ltr"),
            "ltr",
        ),
        (
            "direction",
            "alignment.readingOrder",
            PropValue::from("rtl"),
            "rtl",
        ),
        (
            "dir",
            "alignment.readingOrder",
            PropValue::from("ltr"),
            "ltr",
        ),
    ];

    for (alias, canonical, value, expected) in aliases {
        for eager in [false, true] {
            let original = open_owner(&bytes, eager);
            let updated = original
                .set(CELL_A1, Props::new().with(alias, value.clone()))
                .unwrap_or_else(|error| panic!("alias {alias} should be accepted: {error:?}"));
            assert_eq!(
                cell_property(&updated, CELL_A1, canonical).as_deref(),
                Some(expected),
                "{alias} -> {canonical}"
            );
        }
    }

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original owner");
        for properties in [
            Props::new()
                .with("align", "left")
                .with("alignment.horizontal", "right"),
            Props::new()
                .with("wrap", true)
                .with("alignment.wrapText", false),
            Props::new()
                .with("rotation", 30_i32)
                .with("textrotation", 31_i32),
            Props::new()
                .with("direction", "ltr")
                .with("alignment.readingOrder", "rtl"),
        ] {
            let error = original
                .set(CELL_A1, properties)
                .expect_err("canonical and alias duplicates must refuse");
            assert_eq!(error.code(), Code::InvalidArgument);
            assert_eq!(original.to_bytes().expect("unchanged alias owner"), before);
        }
    }
}

#[test]
fn alignment_explicit_defaults_and_null_clears_are_stable() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original owner");
        let cleared_absent = original
            .set(
                CELL_A1,
                Props::new()
                    .with("alignment.horizontal", PropValue::Null)
                    .with("alignment.vertical", PropValue::Null)
                    .with("alignment.wrapText", PropValue::Null)
                    .with("alignment.textRotation", PropValue::Null)
                    .with("alignment.indent", PropValue::Null)
                    .with("alignment.shrinkToFit", PropValue::Null)
                    .with("alignment.readingOrder", PropValue::Null),
            )
            .expect("clearing absent alignment is a no-op");
        assert_eq!(
            cleared_absent.to_bytes().expect("absent clear bytes"),
            before
        );

        let defaults = original
            .set(
                CELL_A1,
                props! {
                    "alignment.horizontal" => "general",
                    "alignment.vertical" => "top",
                    "alignment.wrapText" => false,
                    "alignment.textRotation" => 0,
                    "alignment.indent" => 0,
                    "alignment.shrinkToFit" => false,
                    "alignment.readingOrder" => "context",
                },
            )
            .expect("assign explicit alignment defaults");
        assert_eq!(
            cell_property(&defaults, CELL_A1, "alignment.horizontal").as_deref(),
            Some("general")
        );
        assert_eq!(
            cell_property(&defaults, CELL_A1, "alignment.vertical").as_deref(),
            Some("top")
        );
        for key in [
            "alignment.wrapText",
            "alignment.textRotation",
            "alignment.indent",
            "alignment.shrinkToFit",
            "alignment.readingOrder",
        ] {
            assert_eq!(
                cell_property(&defaults, CELL_A1, key),
                None,
                "ordinary read omits the explicit {key} default"
            );
        }
        let repeated = defaults
            .set(
                CELL_A1,
                props! {
                    "alignment.horizontal" => "general",
                    "alignment.vertical" => "top",
                    "alignment.wrapText" => false,
                    "alignment.textRotation" => 0,
                    "alignment.indent" => 0,
                    "alignment.shrinkToFit" => false,
                    "alignment.readingOrder" => "context",
                },
            )
            .expect("repeat explicit defaults");
        assert_eq!(
            repeated.to_bytes().expect("repeated defaults bytes"),
            defaults.to_bytes().expect("defaults bytes")
        );

        let cleared = defaults
            .set(
                CELL_A1,
                Props::new()
                    .with("alignment.horizontal", PropValue::Null)
                    .with("alignment.vertical", PropValue::Null)
                    .with("alignment.wrapText", PropValue::Null)
                    .with("alignment.textRotation", PropValue::Null)
                    .with("alignment.indent", PropValue::Null)
                    .with("alignment.shrinkToFit", PropValue::Null)
                    .with("alignment.readingOrder", PropValue::Null),
            )
            .expect("null clear explicit alignment defaults");
        for key in [
            "alignment.horizontal",
            "alignment.vertical",
            "alignment.wrapText",
            "alignment.textRotation",
            "alignment.indent",
            "alignment.shrinkToFit",
            "alignment.readingOrder",
        ] {
            assert_eq!(cell_property(&cleared, CELL_A1, key), None, "{key} clear");
        }
        let repeated_clear = cleared
            .set(
                CELL_A1,
                Props::new()
                    .with("alignment.horizontal", PropValue::Null)
                    .with("alignment.vertical", PropValue::Null)
                    .with("alignment.wrapText", PropValue::Null)
                    .with("alignment.textRotation", PropValue::Null)
                    .with("alignment.indent", PropValue::Null)
                    .with("alignment.shrinkToFit", PropValue::Null)
                    .with("alignment.readingOrder", PropValue::Null),
            )
            .expect("repeat null clear");
        assert_eq!(
            repeated_clear.to_bytes().expect("repeated clear bytes"),
            cleared.to_bytes().expect("clear bytes")
        );
    }
}

#[test]
fn invalid_alignment_values_and_duplicate_keys_are_atomic() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original owner");
        for (key, value) in [
            ("alignment.horizontal", "middle"),
            ("alignment.vertical", "middle"),
            ("alignment.readingOrder", "sideways"),
            ("alignment.wrapText", "sometimes"),
            ("alignment.shrinkToFit", "sometimes"),
        ] {
            let error = original
                .set(CELL_A1, Props::new().with(key, value))
                .expect_err("unknown alignment token or flag must refuse");
            assert_eq!(error.code(), Code::InvalidValue, "{key}");
            assert_eq!(error.property(), Some(key), "{key}");
            assert_eq!(
                original.to_bytes().expect("unchanged alignment owner"),
                before
            );
        }
        for (key, value) in [
            ("alignment.textRotation", -1_i32),
            ("alignment.textRotation", 181_i32),
            ("alignment.textRotation", 254_i32),
            ("alignment.textRotation", 256_i32),
            ("alignment.indent", -1_i32),
            ("alignment.indent", 256_i32),
        ] {
            let error = original
                .set(CELL_A1, Props::new().with(key, value))
                .expect_err("alignment integer must stay in its OfficeCLI domain");
            assert_eq!(error.code(), Code::InvalidValue, "{key}={value}");
            assert_eq!(error.property(), Some(key), "{key}={value}");
            assert_eq!(
                original.to_bytes().expect("unchanged alignment integer"),
                before
            );
        }

        let combined = original
            .set(
                CELL_A1,
                props! {
                    "value" => "must not publish",
                    "alignment.horizontal" => "middle",
                },
            )
            .expect_err("invalid alignment must reject its payload atomically");
        assert_eq!(combined.code(), Code::InvalidValue);
        assert_eq!(combined.property(), Some("alignment.horizontal"));
        assert_eq!(
            original.to_bytes().expect("unchanged combined owner"),
            before
        );

        for properties in [
            Props::new()
                .with("alignment.horizontal", "left")
                .with("alignment.horizontal", "right"),
            Props::new()
                .with("alignment.vertical", "top")
                .with("alignment.vertical", "bottom"),
            Props::new()
                .with("alignment.wrapText", true)
                .with("alignment.wrapText", false),
            Props::new()
                .with("alignment.textRotation", 1_i32)
                .with("alignment.textRotation", 2_i32),
            Props::new()
                .with("alignment.indent", 1_i32)
                .with("alignment.indent", 2_i32),
            Props::new()
                .with("alignment.shrinkToFit", true)
                .with("alignment.shrinkToFit", false),
            Props::new()
                .with("alignment.readingOrder", "ltr")
                .with("alignment.readingOrder", "rtl"),
        ] {
            let error = original
                .set(CELL_A1, properties)
                .expect_err("duplicate alignment keys must refuse");
            assert_eq!(error.code(), Code::InvalidArgument);
            assert_eq!(
                original.to_bytes().expect("unchanged duplicate owner"),
                before
            );
        }
    }
}

#[test]
fn source_alignment_values_and_unknown_attributes_survive_known_edit() {
    let bytes = source_alignment_fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        for (key, expected) in [
            ("alignment.horizontal", "center"),
            ("alignment.vertical", "bottom"),
            ("alignment.wrapText", "true"),
            ("alignment.textRotation", "90"),
            ("alignment.indent", "3"),
            ("alignment.relativeIndent", "0"),
            ("alignment.justifyLastLine", "false"),
            ("alignment.shrinkToFit", "true"),
            ("alignment.readingOrder", "rtl"),
        ] {
            assert_eq!(
                cell_property(&original, CELL_A1, key).as_deref(),
                Some(expected),
                "{key} source value"
            );
        }

        let changed = original
            .set(CELL_A1, props! { "alignment.horizontal" => "right" })
            .expect("change known alignment while preserving unknown attribute");
        assert_eq!(
            cell_property(&changed, CELL_A1, "alignment.horizontal").as_deref(),
            Some("right")
        );
        let styles = styles_bytes(&changed);
        let styles = std::str::from_utf8(&styles).expect("edited styles XML");
        assert!(
            styles.contains(r#"future:alignmentMarker="keep""#),
            "known alignment edit dropped the opaque alignment attribute: {styles}"
        );
        assert!(
            styles.contains(r#"future:xfMarker="keep""#),
            "known alignment edit dropped the opaque XF attribute: {styles}"
        );
    }
}

#[test]
fn alignment_protection_and_quote_round_trip_source_and_eager() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original owner");
        let updated = original
            .set(
                CELL_A1,
                props! {
                    "value" => "protected payload",
                    "font.bold" => true,
                    "alignment.horizontal" => "center",
                    "alignment.relativeIndent" => -2,
                    "alignment.justifyLastLine" => true,
                    "protection.locked" => false,
                    "protection.hidden" => true,
                    "quotePrefix" => true,
                },
            )
            .expect("assign extended alignment and cell-format properties");

        assert_eq!(
            cell_text(&updated, CELL_A1).as_deref(),
            Some("protected payload")
        );
        assert_eq!(
            cell_property(&updated, CELL_A1, "font.bold").as_deref(),
            Some("true")
        );
        for (key, expected) in [
            ("alignment.horizontal", "center"),
            ("alignment.relativeIndent", "-2"),
            ("alignment.justifyLastLine", "true"),
            ("protection.locked", "false"),
            ("protection.hidden", "true"),
            ("quotePrefix", "true"),
        ] {
            assert_eq!(
                cell_property(&updated, CELL_A1, key).as_deref(),
                Some(expected),
                "{key} readback"
            );
        }

        let styles = styles_bytes(&updated);
        let styles = std::str::from_utf8(&styles).expect("styles XML");
        for fragment in [
            r#"relativeIndent="-2""#,
            r#"justifyLastLine="1""#,
            r#"locked="0""#,
            r#"hidden="1""#,
            r#"quotePrefix="1""#,
        ] {
            assert!(styles.contains(fragment), "missing {fragment} in {styles}");
        }

        let reopened = open_owner(
            &updated
                .to_bytes()
                .expect("publish extended cell-format edit"),
            eager,
        );
        for (key, expected) in [
            ("alignment.relativeIndent", "-2"),
            ("alignment.justifyLastLine", "true"),
            ("protection.locked", "false"),
            ("protection.hidden", "true"),
            ("quotePrefix", "true"),
        ] {
            assert_eq!(
                cell_property(&reopened, CELL_A1, key).as_deref(),
                Some(expected),
                "{key} after reopen"
            );
        }
        assert_eq!(
            original.to_bytes().expect("old owner remains unchanged"),
            before
        );
    }
}

#[test]
fn extended_cell_format_ranges_keep_companions_and_replay_inverse() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original owner");
        let commit = original
            .edit(|edit| {
                edit.set(
                    CELL_A1,
                    props! {
                        "value" => "extended companion",
                        "font.bold" => true,
                        "alignment.horizontal" => "center",
                    },
                )?;
                edit.set(
                    "/sheet[1]/A1:D2",
                    props! {
                        "alignment.relativeIndent" => 4,
                        "alignment.justifyLastLine" => true,
                        "protection.locked" => true,
                        "protection.hidden" => false,
                        "quotePrefix" => true,
                    },
                )?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("commit extended cell-format range");

        for path in [CELL_A1, CELL_B1, CELL_C1, CELL_D1, "/sheet[1]/D2"] {
            for (key, expected) in [
                ("alignment.relativeIndent", "4"),
                ("alignment.justifyLastLine", "true"),
                ("protection.locked", "true"),
                ("protection.hidden", "false"),
                ("quotePrefix", "true"),
            ] {
                assert_eq!(
                    cell_property(commit.file(), path, key).as_deref(),
                    Some(expected),
                    "{path} {key}"
                );
            }
        }
        assert_eq!(
            cell_property(commit.file(), CELL_A1, "alignment.horizontal").as_deref(),
            Some("center")
        );
        assert_eq!(
            cell_property(commit.file(), CELL_A1, "font.bold").as_deref(),
            Some("true")
        );
        assert_eq!(
            cell_text(commit.file(), CELL_A1).as_deref(),
            Some("extended companion")
        );
        assert_eq!(
            cell_property(commit.file(), CELL_C1, "formula").as_deref(),
            Some("B1*2")
        );

        let replayed = commit
            .patch()
            .apply(&original)
            .expect("replay extended cell-format range patch");
        assert_eq!(
            replayed.to_bytes().expect("replayed extended bytes"),
            commit.file().to_bytes().expect("committed extended bytes")
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(&replayed)
            .expect("inverse extended cell-format range patch");
        assert_eq!(
            restored.to_bytes().expect("restored extended bytes"),
            before
        );

        let repeated = commit
            .file()
            .set(
                "/sheet[1]/A1:D2",
                props! {
                    "alignment.relativeIndent" => 4,
                    "alignment.justifyLastLine" => true,
                    "protection.locked" => true,
                    "protection.hidden" => false,
                    "quotePrefix" => true,
                },
            )
            .expect("repeat extended cell-format range");
        assert_eq!(
            repeated.to_bytes().expect("repeat extended bytes"),
            commit.file().to_bytes().expect("first extended bytes")
        );
    }
}

#[test]
fn extended_cell_format_aliases_round_trip_and_duplicate_atomically() {
    let bytes = fixture();
    let aliases = [
        (
            "locked",
            "protection.locked",
            PropValue::from(false),
            "false",
        ),
        (
            "formulahidden",
            "protection.hidden",
            PropValue::from(true),
            "true",
        ),
    ];

    for (alias, canonical, value, expected) in aliases {
        for eager in [false, true] {
            let original = open_owner(&bytes, eager);
            let updated = original
                .set(CELL_A1, Props::new().with(alias, value.clone()))
                .unwrap_or_else(|error| panic!("alias {alias} should be accepted: {error:?}"));
            assert_eq!(
                cell_property(&updated, CELL_A1, canonical).as_deref(),
                Some(expected),
                "{alias} -> {canonical}"
            );
        }
    }

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original owner");
        for properties in [
            Props::new()
                .with("locked", true)
                .with("protection.locked", false),
            Props::new()
                .with("formulahidden", true)
                .with("protection.hidden", false),
        ] {
            let error = original
                .set(CELL_A1, properties)
                .expect_err("canonical and protection alias duplicates must refuse");
            assert_eq!(error.code(), Code::InvalidArgument);
            assert_eq!(
                original.to_bytes().expect("unchanged protection owner"),
                before
            );
        }
    }
}

#[test]
fn extended_cell_format_defaults_and_null_clears_are_stable() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original owner");
        let absent_clear = original
            .set(
                CELL_A1,
                Props::new()
                    .with("alignment.relativeIndent", PropValue::Null)
                    .with("alignment.justifyLastLine", PropValue::Null)
                    .with("protection.locked", PropValue::Null)
                    .with("protection.hidden", PropValue::Null)
                    .with("quotePrefix", PropValue::Null),
            )
            .expect("clearing absent extended cell-format properties");
        assert_eq!(absent_clear.to_bytes().expect("absent clear bytes"), before);

        let defaults = original
            .set(
                CELL_A1,
                props! {
                    "alignment.relativeIndent" => 0,
                    "alignment.justifyLastLine" => false,
                    "protection.locked" => false,
                    "protection.hidden" => false,
                    "quotePrefix" => false,
                },
            )
            .expect("assign explicit extended defaults");
        assert_eq!(
            cell_property(&defaults, CELL_A1, "alignment.relativeIndent").as_deref(),
            Some("0")
        );
        assert_eq!(
            cell_property(&defaults, CELL_A1, "protection.locked").as_deref(),
            Some("false")
        );
        assert_eq!(
            cell_property(&defaults, CELL_A1, "protection.hidden").as_deref(),
            Some("false")
        );
        assert_eq!(
            cell_property(&defaults, CELL_A1, "alignment.justifyLastLine").as_deref(),
            Some("false")
        );
        assert_eq!(
            cell_property(&defaults, CELL_A1, "quotePrefix"),
            None,
            "ordinary read omits the explicit quotePrefix false default"
        );
        let repeated = defaults
            .set(
                CELL_A1,
                props! {
                    "alignment.relativeIndent" => 0,
                    "alignment.justifyLastLine" => false,
                    "protection.locked" => false,
                    "protection.hidden" => false,
                    "quotePrefix" => false,
                },
            )
            .expect("repeat explicit extended defaults");
        assert_eq!(
            repeated.to_bytes().expect("repeated default bytes"),
            defaults.to_bytes().expect("default bytes")
        );

        let cleared = defaults
            .set(
                CELL_A1,
                Props::new()
                    .with("alignment.relativeIndent", PropValue::Null)
                    .with("alignment.justifyLastLine", PropValue::Null)
                    .with("protection.locked", PropValue::Null)
                    .with("protection.hidden", PropValue::Null)
                    .with("quotePrefix", PropValue::Null),
            )
            .expect("null clear extended defaults");
        for key in [
            "alignment.relativeIndent",
            "alignment.justifyLastLine",
            "protection.locked",
            "protection.hidden",
            "quotePrefix",
        ] {
            assert_eq!(cell_property(&cleared, CELL_A1, key), None, "{key} clear");
        }
        let repeated_clear = cleared
            .set(
                CELL_A1,
                Props::new()
                    .with("alignment.relativeIndent", PropValue::Null)
                    .with("alignment.justifyLastLine", PropValue::Null)
                    .with("protection.locked", PropValue::Null)
                    .with("protection.hidden", PropValue::Null)
                    .with("quotePrefix", PropValue::Null),
            )
            .expect("repeat null clear");
        assert_eq!(
            repeated_clear.to_bytes().expect("repeated clear bytes"),
            cleared.to_bytes().expect("clear bytes")
        );
    }
}

#[test]
fn invalid_extended_cell_format_values_are_atomic() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original owner");
        for (key, value) in [
            ("alignment.justifyLastLine", "sometimes"),
            ("protection.locked", "sometimes"),
            ("protection.hidden", "sometimes"),
            ("quotePrefix", "sometimes"),
        ] {
            let error = original
                .set(CELL_A1, Props::new().with(key, value))
                .expect_err("extended cell-format flag must be boolean");
            assert_eq!(error.code(), Code::InvalidValue, "{key}");
            assert_eq!(error.property(), Some(key), "{key}");
            assert_eq!(original.to_bytes().expect("unchanged flag owner"), before);
        }
        for value in [i64::from(i32::MIN) - 1, i64::from(i32::MAX) + 1] {
            let error = original
                .set(
                    CELL_A1,
                    Props::new().with("alignment.relativeIndent", value),
                )
                .expect_err("relative alignment indent must fit i32");
            assert_eq!(error.code(), Code::InvalidValue, "relative indent {value}");
            assert_eq!(error.property(), Some("alignment.relativeIndent"));
            assert_eq!(
                original
                    .to_bytes()
                    .expect("unchanged relative indent owner"),
                before
            );
        }

        let combined = original
            .set(
                CELL_A1,
                props! {
                    "value" => "must not publish",
                    "protection.locked" => "sometimes",
                },
            )
            .expect_err("invalid protection must reject its payload atomically");
        assert_eq!(combined.code(), Code::InvalidValue);
        assert_eq!(combined.property(), Some("protection.locked"));
        assert_eq!(
            original.to_bytes().expect("unchanged combined owner"),
            before
        );

        for properties in [
            Props::new()
                .with("alignment.relativeIndent", 1_i64)
                .with("alignment.relativeIndent", 2_i64),
            Props::new()
                .with("alignment.justifyLastLine", true)
                .with("alignment.justifyLastLine", false),
            Props::new()
                .with("protection.locked", true)
                .with("protection.locked", false),
            Props::new()
                .with("protection.hidden", true)
                .with("protection.hidden", false),
            Props::new()
                .with("quotePrefix", true)
                .with("quotePrefix", false),
        ] {
            let error = original
                .set(CELL_A1, properties)
                .expect_err("duplicate extended cell-format keys must refuse");
            assert_eq!(error.code(), Code::InvalidArgument);
            assert_eq!(
                original.to_bytes().expect("unchanged duplicate owner"),
                before
            );
        }
    }
}

#[test]
fn number_format_aliases_and_friendly_keywords_round_trip_source_and_eager() {
    let bytes = fixture();
    let aliases = [
        ("numberformat", "0.000"),
        ("numfmt", "0.000"),
        ("format", "0.000"),
    ];

    for eager in [false, true] {
        for (key, code) in aliases {
            let original = open_owner(&bytes, eager);
            let updated = original
                .set(CELL_A1, Props::new().with(key, code))
                .unwrap_or_else(|error| panic!("{key} number format should write: {error:?}"));
            assert_eq!(
                cell_property(&updated, CELL_A1, "numberformat").as_deref(),
                Some(code),
                "{key} ordinary read"
            );
            let reopened = open_owner(&updated.to_bytes().expect("publish number format"), eager);
            assert_eq!(
                cell_property(&reopened, CELL_A1, "numberformat").as_deref(),
                Some(code),
                "{key} read after reopen"
            );
            let repeated = updated
                .set(CELL_A1, Props::new().with(key, code))
                .unwrap_or_else(|error| panic!("{key} repeat should write: {error:?}"));
            assert_eq!(
                repeated.to_bytes().expect("repeat number-format bytes"),
                updated.to_bytes().expect("number-format bytes")
            );
        }

        // These names are the OfficeCLI-friendly spellings accepted by the
        // unified adapter.  The expected strings are the native SpreadsheetML
        // format codes, so this also exercises built-in and custom resolution.
        let friendly = [
            ("general", None),
            ("number", Some("#,##0.00")),
            ("comma", Some("#,##0.00")),
            ("currency", Some("\"$\"#,##0.00")),
            (
                "accounting",
                Some(r#"_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)"#),
            ),
            ("percent", Some("0.00%")),
            ("percentage", Some("0.00%")),
            ("scientific", Some("0.00E+00")),
            ("text", Some("@")),
            ("date", Some("yyyy-mm-dd")),
            ("time", Some("h:mm:ss")),
            ("datetime", Some("yyyy-mm-dd h:mm:ss")),
        ];
        for (keyword, expected) in friendly {
            let original = open_owner(&bytes, eager);
            let updated = original
                .set(CELL_A1, Props::new().with("numberformat", keyword))
                .unwrap_or_else(|error| {
                    panic!("friendly format {keyword} should write: {error:?}")
                });
            assert_eq!(
                cell_property(&updated, CELL_A1, "numberformat").as_deref(),
                expected,
                "friendly format {keyword}"
            );
            let reopened = open_owner(&updated.to_bytes().expect("publish friendly format"), eager);
            assert_eq!(
                cell_property(&reopened, CELL_A1, "numberformat").as_deref(),
                expected,
                "friendly format {keyword} after reopen"
            );
        }
    }
}

#[test]
fn number_format_builtins_custom_ranges_and_replay_inverse() {
    let bytes = fixture();
    let builtins = [
        ("0", "0"),
        ("0.00", "0.00"),
        ("#,##0", "#,##0"),
        ("0%", "0%"),
        ("mm-dd-yy", "mm-dd-yy"),
        ("@", "@"),
    ];

    for eager in [false, true] {
        for (code, expected) in builtins {
            let original = open_owner(&bytes, eager);
            let updated = original
                .set(CELL_B1, Props::new().with("numberformat", code))
                .unwrap_or_else(|error| panic!("built-in format {code} should write: {error:?}"));
            assert_eq!(
                cell_property(&updated, CELL_B1, "numberformat").as_deref(),
                Some(expected),
                "built-in format {code}"
            );
            let reopened = open_owner(&updated.to_bytes().expect("publish built-in format"), eager);
            assert_eq!(
                cell_property(&reopened, CELL_B1, "numberformat").as_deref(),
                Some(expected),
                "built-in format {code} after reopen"
            );
        }

        let original = open_owner(&bytes, eager);
        let before = original
            .to_bytes()
            .expect("publish original number-format owner");
        let custom = r#"0.000 "kg""#;
        let commit = original
            .edit(|edit| {
                edit.set(
                    CELL_A1,
                    props! {
                        "value" => "number-format payload",
                        "numberformat" => custom,
                    },
                )?;
                edit.set(
                    "/sheet[1]/B1:D2",
                    Props::new().with("numberformat", "0.00%"),
                )?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("commit custom and range number formats");

        assert_eq!(
            cell_property(commit.file(), CELL_A1, "numberformat").as_deref(),
            Some(custom)
        );
        for path in [CELL_B1, CELL_C1, CELL_D1, "/sheet[1]/D2"] {
            assert_eq!(
                cell_property(commit.file(), path, "numberformat").as_deref(),
                Some("0.00%"),
                "range number format at {path}"
            );
        }
        assert_eq!(
            cell_text(commit.file(), CELL_A1).as_deref(),
            Some("number-format payload")
        );
        assert_eq!(
            cell_property(commit.file(), CELL_C1, "formula").as_deref(),
            Some("B1*2")
        );
        let styles = styles_bytes(commit.file());
        let styles = std::str::from_utf8(&styles).expect("number-format styles XML");
        assert!(
            styles.contains(r#"formatCode="0.000 &quot;kg&quot;""#)
                && styles.contains(r#"numFmtId="10""#),
            "custom/range number formats must be retained in styles XML: {styles}"
        );

        let replayed = commit
            .patch()
            .apply(&original)
            .expect("replay number-format patch");
        assert_eq!(
            replayed.to_bytes().expect("replayed number-format bytes"),
            commit.file().to_bytes().expect("number-format bytes")
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(&replayed)
            .expect("inverse number-format patch");
        assert_eq!(
            restored.to_bytes().expect("restored number-format bytes"),
            before
        );

        let repeated = commit
            .file()
            .set(
                "/sheet[1]/B1:D2",
                Props::new().with("numberformat", "0.00%"),
            )
            .expect("repeat range number format");
        assert_eq!(
            repeated
                .to_bytes()
                .expect("repeat range number-format bytes"),
            commit
                .file()
                .to_bytes()
                .expect("first range number-format bytes")
        );
    }
}

#[test]
fn number_format_general_and_null_clears_are_exact_noops() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        // The fixture authors numFmtId=0. The first clear removes that
        // reference; only the subsequent clear is an absent-attribute no-op.
        let first_clear = original
            .set(CELL_A1, Props::new().with("numberformat", PropValue::Null))
            .expect("clear authored default number format");
        assert_eq!(cell_property(&first_clear, CELL_A1, "numberformat"), None);
        let absent_clear = first_clear
            .set(CELL_A1, Props::new().with("numberformat", PropValue::Null))
            .expect("clear absent number format");
        assert_eq!(
            absent_clear
                .to_bytes()
                .expect("absent number-format clear bytes"),
            first_clear
                .to_bytes()
                .expect("first number-format clear bytes")
        );

        let custom = original
            .set(CELL_A1, Props::new().with("numberformat", r#"0.000 "kg""#))
            .expect("assign custom number format");
        assert_eq!(
            cell_property(&custom, CELL_A1, "numberformat").as_deref(),
            Some(r#"0.000 "kg""#)
        );
        let cleared = custom
            .set(CELL_A1, Props::new().with("numberformat", PropValue::Null))
            .expect("clear custom number format");
        assert_eq!(cell_property(&cleared, CELL_A1, "numberformat"), None);
        let repeated_clear = cleared
            .set(CELL_A1, Props::new().with("numberformat", PropValue::Null))
            .expect("repeat null number-format clear");
        assert_eq!(
            repeated_clear
                .to_bytes()
                .expect("repeated null number-format bytes"),
            cleared.to_bytes().expect("null number-format bytes")
        );

        let general = custom
            .set(CELL_A1, Props::new().with("numberformat", "general"))
            .expect("clear custom number format with General");
        assert_eq!(cell_property(&general, CELL_A1, "numberformat"), None);
        let repeated_general = general
            .set(CELL_A1, Props::new().with("numberformat", "General"))
            .expect("repeat General number format");
        assert_eq!(
            repeated_general.to_bytes().expect("repeated General bytes"),
            general.to_bytes().expect("General bytes")
        );
    }
}

#[test]
fn invalid_and_duplicate_number_formats_are_atomic() {
    let bytes = fixture();
    let invalid = [
        ("numberformat", "".to_owned()),
        ("numberformat", "0\"".to_owned()),
        ("numberformat", "[Red".to_owned()),
        ("numberformat", "0;0;0;0;0".to_owned()),
        ("numberformat", "0\\".to_owned()),
        ("numberformat", "0_".to_owned()),
        ("numberformat", "0*".to_owned()),
        ("numberformat", "0]".to_owned()),
        ("numberformat", "a\u{1}b".to_owned()),
        ("numfmt", "[Red".to_owned()),
        ("format", "0\\".to_owned()),
    ];

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original
            .to_bytes()
            .expect("publish original number-format owner");
        for (key, value) in &invalid {
            let error = original
                .set(CELL_A1, Props::new().with(*key, value.as_str()))
                .unwrap_err();
            assert_value_refusal(error);
            assert_eq!(
                original
                    .to_bytes()
                    .expect("unchanged invalid number-format owner"),
                before,
                "invalid {key}={value:?} must be atomic"
            );
        }

        for value in [
            format!("\"{}\"", "€".repeat(253)),
            format!("\"{}\"", "€".repeat(254)),
        ] {
            let result = original.set(CELL_A1, Props::new().with("numberformat", value.as_str()));
            if value.chars().count() == 255 {
                let updated = result.expect("255 Unicode scalar number-format code is valid");
                assert_eq!(
                    cell_property(&updated, CELL_A1, "numberformat").as_deref(),
                    Some(value.as_str())
                );
            } else {
                assert_value_refusal(result.expect_err("256 Unicode scalar number-format code"));
                assert_eq!(
                    original
                        .to_bytes()
                        .expect("unchanged oversized number-format owner"),
                    before
                );
            }
        }

        for properties in [
            Props::new()
                .with("numberformat", "0.00")
                .with("numfmt", "0.000"),
            Props::new()
                .with("numberformat", "0.00")
                .with("format", "0.000"),
            Props::new().with("numfmt", "0.00").with("format", "0.000"),
        ] {
            let error = original
                .set(CELL_A1, properties)
                .expect_err("duplicate number-format aliases must refuse");
            assert_eq!(error.code(), Code::InvalidArgument);
            assert_eq!(
                original
                    .to_bytes()
                    .expect("unchanged duplicate number-format owner"),
                before
            );
        }

        let combined = original
            .set(
                CELL_A1,
                props! {
                    "value" => "must not publish",
                    "numberformat" => "[Red",
                },
            )
            .expect_err("invalid number format must reject its payload atomically");
        assert_value_refusal(combined);
        assert_eq!(
            original
                .to_bytes()
                .expect("unchanged combined number-format owner"),
            before
        );
    }
}

#[test]
fn fill_colors_aliases_tints_and_null_round_trip_source_and_eager() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        for (key, value) in [
            ("fill", "red"),
            ("bgcolor", "#10203080"),
            ("bg", "rgba(10,20,30,0.5)"),
        ] {
            let updated = original
                .set(CELL_A1, Props::new().with(key, value))
                .unwrap_or_else(|error| panic!("{key} fill alias should write: {error:?}"));
            let expected = match value {
                "red" => "#FF0000",
                "#10203080" => "#10203080",
                _ => "#0A141E80",
            };
            assert_eq!(
                cell_property(&updated, CELL_A1, "fill").as_deref(),
                Some(expected),
                "{key} ordinary fill"
            );
            let reopened = open_owner(&updated.to_bytes().expect("publish fill alias"), eager);
            assert_eq!(
                cell_property(&reopened, CELL_A1, "fill").as_deref(),
                Some(expected),
                "{key} fill after reopen"
            );
        }

        let theme = original
            .set(CELL_A1, Props::new().with("fill", "accent1+tint40"))
            .expect("theme fill tint");
        assert_eq!(
            cell_property(&theme, CELL_A1, "fill").as_deref(),
            Some("accent1+tint40")
        );
        let theme_styles = styles_bytes(&theme);
        let theme_styles = std::str::from_utf8(&theme_styles).expect("theme fill styles XML");
        assert!(
            theme_styles.contains(r#"<fgColor theme="4" tint="0.4"/>"#),
            "theme fill tint must remain native: {theme_styles}"
        );

        let rgb_tinted = original
            .set(CELL_A1, Props::new().with("fill", "#102030+tint40"))
            .expect("RGB fill tint");
        assert_eq!(
            cell_property(&rgb_tinted, CELL_A1, "fill").as_deref(),
            Some("#102030"),
            "ordinary fill read follows OfficeCLI and drops RGB tint"
        );
        let rgb_styles = styles_bytes(&rgb_tinted);
        let rgb_styles = std::str::from_utf8(&rgb_styles).expect("RGB fill styles XML");
        assert!(
            rgb_styles.contains(r#"<fgColor rgb="FF102030" tint="0.4"/>"#),
            "RGB fill tint must remain native: {rgb_styles}"
        );

        let typed = original
            .set(
                CELL_A1,
                Props::new().with(
                    "fill",
                    PropValue::Color(Color::rgba(0x11, 0x22, 0x33, 0x44)),
                ),
            )
            .expect("typed fill color");
        assert_eq!(
            cell_property(&typed, CELL_A1, "fill").as_deref(),
            Some("#11223344")
        );
        let typed_styles = styles_bytes(&typed);
        let typed_styles = std::str::from_utf8(&typed_styles).expect("typed fill styles XML");
        assert!(typed_styles.contains(r#"<fgColor rgb="44112233"/>"#));

        let cleared = typed
            .set(CELL_A1, Props::new().with("fill", PropValue::Null))
            .expect("clear authored fill");
        assert_eq!(cell_property(&cleared, CELL_A1, "fill"), None);
        assert_eq!(cell_text(&cleared, CELL_A1).as_deref(), Some("seed"));

        let before = original.to_bytes().expect("original fill bytes");
        let absent_clear = original
            .set(CELL_A1, Props::new().with("fill", PropValue::Null))
            .expect("clear absent fill");
        assert_eq!(
            absent_clear.to_bytes().expect("absent fill clear bytes"),
            before,
            "clearing an absent fill must be an exact no-op"
        );
    }
}

#[test]
fn fill_patterns_and_friendly_gray_aliases_round_trip_with_backgrounds() {
    let bytes = fixture();
    let patterns = [
        ("mediumGray", "mediumGray"),
        ("darkGray", "darkGray"),
        ("lightGray", "lightGray"),
        ("darkHorizontal", "darkHorizontal"),
        ("darkVertical", "darkVertical"),
        ("darkDown", "darkDown"),
        ("darkUp", "darkUp"),
        ("darkGrid", "darkGrid"),
        ("darkTrellis", "darkTrellis"),
        ("lightHorizontal", "lightHorizontal"),
        ("lightVertical", "lightVertical"),
        ("lightDown", "lightDown"),
        ("lightUp", "lightUp"),
        ("lightGrid", "lightGrid"),
        ("lightTrellis", "lightTrellis"),
        ("gray125", "gray125"),
        ("gray0625", "gray0625"),
    ];
    let friendly = [
        ("gray50", "mediumGray"),
        ("gray-50", "mediumGray"),
        ("gray 50", "mediumGray"),
        ("gray75", "darkGray"),
        ("gray-75", "darkGray"),
        ("gray 75", "darkGray"),
        ("gray25", "lightGray"),
        ("gray-25", "lightGray"),
        ("gray 25", "lightGray"),
    ];

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        for (input, expected) in patterns.into_iter().chain(friendly) {
            let updated = original
                .set(CELL_A1, Props::new().with("fillPattern", input))
                .unwrap_or_else(|error| panic!("fill pattern {input} should write: {error:?}"));
            assert_eq!(
                cell_property(&updated, CELL_A1, "fillPattern").as_deref(),
                Some(expected),
                "pattern {input}"
            );
            assert_eq!(
                cell_property(&updated, CELL_A1, "fill"),
                None,
                "a pattern replacement clears its omitted foreground: {input}"
            );
            assert_eq!(
                cell_property(&updated, CELL_A1, "fillBg"),
                None,
                "a pattern replacement clears its omitted background: {input}"
            );
            let styles = styles_bytes(&updated);
            let styles = std::str::from_utf8(&styles).expect("pattern styles XML");
            assert!(
                styles.contains(&format!(r#"patternType="{expected}""#)),
                "pattern {input} must be native {expected}: {styles}"
            );
        }

        let solid = original
            .set(CELL_A1, Props::new().with("fillPattern", "solid"))
            .expect("solid pattern token");
        assert_eq!(cell_property(&solid, CELL_A1, "fillPattern"), None);
        let solid_styles = styles_bytes(&solid);
        let solid_styles = std::str::from_utf8(&solid_styles).expect("solid styles XML");
        assert!(solid_styles.contains(r#"patternType="solid""#));

        let none = original
            .set(CELL_A1, Props::new().with("fillPattern", "none"))
            .expect("none pattern token");
        assert_eq!(cell_property(&none, CELL_A1, "fillPattern"), None);
        let none_styles = styles_bytes(&none);
        let none_styles = std::str::from_utf8(&none_styles).expect("none styles XML");
        assert!(none_styles.contains(r#"patternType="none""#));

        let updated = original
            .set(
                CELL_A1,
                Props::new()
                    .with("fillPattern", "darkGray")
                    .with("fillBg", "accent1+tint40"),
            )
            .expect("background with explicit pattern should write");
        assert_eq!(
            cell_property(&updated, CELL_A1, "fillPattern").as_deref(),
            Some("darkGray")
        );
        assert_eq!(
            cell_property(&updated, CELL_A1, "fillBg").as_deref(),
            Some("accent1+tint40")
        );

        let combined = original
            .set(
                CELL_A1,
                props! {
                    "fill" => "#102030",
                    "fillPattern" => "mediumGray",
                    "fillBg" => "#FFEEDD",
                },
            )
            .expect("all fill facets in one operation");
        assert_eq!(
            cell_property(&combined, CELL_A1, "fill").as_deref(),
            Some("#102030")
        );
        assert_eq!(
            cell_property(&combined, CELL_A1, "fillPattern").as_deref(),
            Some("mediumGray")
        );
        assert_eq!(
            cell_property(&combined, CELL_A1, "fillBg").as_deref(),
            Some("#FFEEDD")
        );
    }
}

#[test]
fn source_fill_patterns_retain_incremental_colors_and_unknown_markup() {
    let bytes = source_fill_fixture();
    let source = open_owner(&bytes, false);
    assert_eq!(
        cell_property(&source, CELL_A1, "fillPattern").as_deref(),
        Some("darkGray")
    );
    assert_eq!(
        cell_property(&source, CELL_A1, "fill").as_deref(),
        Some("#102030")
    );
    assert_eq!(
        cell_property(&source, CELL_A1, "fillBg").as_deref(),
        Some("accent1+tint40")
    );
    assert_eq!(
        cell_property(&source, CELL_B1, "fillPattern").as_deref(),
        Some("lightGray")
    );
    assert_eq!(
        cell_property(&source, CELL_B1, "fill").as_deref(),
        Some("#405060")
    );
    assert_eq!(
        cell_property(&source, CELL_B1, "fillBg").as_deref(),
        Some("#F0E0D0")
    );

    let foreground = source
        .set(CELL_A1, Props::new().with("fill", "#AABBCC"))
        .expect("incremental foreground fill");
    assert_eq!(
        cell_property(&foreground, CELL_A1, "fillPattern").as_deref(),
        Some("darkGray")
    );
    assert_eq!(
        cell_property(&foreground, CELL_A1, "fillBg").as_deref(),
        Some("accent1+tint40")
    );
    let foreground_styles = styles_bytes(&foreground);
    let foreground_styles = std::str::from_utf8(&foreground_styles).expect("source styles XML");
    assert!(foreground_styles.contains(r#"future:patternMarker="keep""#));
    assert!(foreground_styles.contains(r#"<future:fillExtension marker="keep"/>"#));

    let replacement = source
        .set(CELL_A1, Props::new().with("fillPattern", "lightGray"))
        .expect("replace source pattern");
    assert_eq!(
        cell_property(&replacement, CELL_A1, "fillPattern").as_deref(),
        Some("lightGray")
    );
    assert_eq!(cell_property(&replacement, CELL_A1, "fill"), None);
    assert_eq!(cell_property(&replacement, CELL_A1, "fillBg"), None);
    let replacement_styles = styles_bytes(&replacement);
    let replacement_styles =
        std::str::from_utf8(&replacement_styles).expect("replacement styles XML");
    assert!(replacement_styles.contains(r#"future:patternMarker="keep""#));
    assert!(replacement_styles.contains(r#"<future:fillExtension marker="keep"/>"#));

    let pattern_cleared = source
        .set(CELL_A1, Props::new().with("fillPattern", PropValue::Null))
        .expect("clear authored pattern attribute");
    assert_eq!(
        cell_property(&pattern_cleared, CELL_A1, "fillPattern"),
        None
    );
    assert_eq!(
        cell_property(&pattern_cleared, CELL_A1, "fill").as_deref(),
        Some("#102030"),
        "clearing only the pattern must retain the foreground color"
    );
    let pattern_cleared_styles = styles_bytes(&pattern_cleared);
    let pattern_cleared_styles =
        std::str::from_utf8(&pattern_cleared_styles).expect("cleared pattern styles XML");
    assert!(
        pattern_cleared_styles.contains(r#"<bgColor theme="4" tint="0.4"/>"#),
        "clearing only the pattern must retain the background color"
    );

    let background_cleared = source
        .set(CELL_A1, Props::new().with("fillBg", PropValue::Null))
        .expect("clear authored fill background");
    assert_eq!(
        cell_property(&background_cleared, CELL_A1, "fillPattern").as_deref(),
        Some("darkGray")
    );
    assert_eq!(
        cell_property(&background_cleared, CELL_A1, "fill").as_deref(),
        Some("#102030")
    );
    assert_eq!(cell_property(&background_cleared, CELL_A1, "fillBg"), None);
    let background_cleared_again = background_cleared
        .set(CELL_A1, Props::new().with("fillBg", PropValue::Null))
        .expect("repeat fill background clear");
    assert_eq!(
        background_cleared_again
            .to_bytes()
            .expect("repeated fill background clear bytes"),
        background_cleared
            .to_bytes()
            .expect("fill background clear bytes")
    );

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let background_without_pattern = original
            .set(CELL_D1, Props::new().with("fillBg", "#112233"))
            .expect_err("background-only fill requires a non-solid pattern");
        assert_eq!(
            background_without_pattern.code(),
            Code::UnsupportedCapability
        );
        assert_eq!(background_without_pattern.property(), Some("fillBg"));

        let gradient_to_solid = original
            .set(CELL_C1, Props::new().with("fill", "#AABBCC"))
            .expect("gradient-to-solid replacement");
        assert_eq!(
            cell_property(&gradient_to_solid, CELL_C1, "fill").as_deref(),
            Some("#AABBCC")
        );
        assert_eq!(
            cell_property(&gradient_to_solid, CELL_C1, "fillPattern"),
            None,
            "solid conversion replaces the gradient as one fill"
        );

        let gradient_to_pattern = original
            .set(CELL_C1, Props::new().with("fillPattern", "darkGray"))
            .expect("gradient-to-pattern replacement");
        assert_eq!(
            cell_property(&gradient_to_pattern, CELL_C1, "fillPattern").as_deref(),
            Some("darkGray")
        );

        let background_on_gradient = original
            .set(CELL_C1, Props::new().with("fillBg", "#112233"))
            .expect_err("background-only gradient edit remains unsupported");
        assert_eq!(background_on_gradient.code(), Code::UnsupportedCapability);
        assert_eq!(background_on_gradient.property(), Some("fillBg"));

        let gradient_before_nulls = original
            .to_bytes()
            .expect("gradient bytes before null attribute clears");
        let pattern_null = original
            .set(CELL_C1, Props::new().with("fillPattern", PropValue::Null))
            .expect("clearing absent gradient pattern is a no-op");
        assert_eq!(
            pattern_null
                .to_bytes()
                .expect("gradient bytes after null pattern"),
            gradient_before_nulls,
            "null fillPattern must not remove or rewrite a gradient"
        );
        let background_null = original
            .set(CELL_C1, Props::new().with("fillBg", PropValue::Null))
            .expect("clearing absent gradient background is a no-op");
        assert_eq!(
            background_null
                .to_bytes()
                .expect("gradient bytes after null background"),
            gradient_before_nulls,
            "null fillBg must not remove or rewrite a gradient"
        );
    }
}

#[test]
fn gradient_fill_readback_retains_all_stops_and_path_geometry() {
    let bytes = source_path_gradient_fixture();

    for eager in [false, true] {
        let file = open_owner(&bytes, eager);
        assert_eq!(
            cell_property(&file, CELL_C1, "fill").as_deref(),
            Some("gradient;#FF0000;#00FF00;#0000FF;0"),
            "the ordinary projection reports all RGB stops"
        );

        let styles = styles_bytes(&file);
        let styles = std::str::from_utf8(&styles).expect("path-gradient styles XML");
        for fragment in [
            r#"<gradientFill type="path" left="0.1" right="0.8" top="0.2" bottom="0.9">"#,
            r#"<stop position="0.125"><color rgb="FFFF0000"/></stop>"#,
            r#"<stop position="0.5"><color rgb="FF00FF00"/></stop>"#,
            r#"<stop position="0.875"><color rgb="FF0000FF"/></stop>"#,
        ] {
            assert!(
                styles.contains(fragment),
                "missing source gradient fragment: {fragment}"
            );
        }

        let workbook = litchi_xlsx::Package::from_bytes(&bytes).expect("open native path gradient");
        let native_format = workbook
            .workbook_ref()
            .sheet(0usize)
            .expect("native sheet lookup")
            .expect("native worksheet")
            .style("C1")
            .expect("native style lookup")
            .expect("native style")
            .format()
            .expect("native format");
        let gradient = native_format
            .fill()
            .and_then(litchi_xlsx::style::cell_format::Fill::gradient)
            .expect("native path gradient");
        assert_eq!(
            gradient.kind(),
            litchi_xlsx::style::cell_format::GradientKind::Path
        );
        assert_eq!(
            (
                gradient.left(),
                gradient.right(),
                gradient.top(),
                gradient.bottom()
            ),
            (0.1, 0.8, 0.2, 0.9)
        );
        assert_eq!(
            gradient
                .stops()
                .iter()
                .map(|stop| stop.position())
                .collect::<Vec<_>>(),
            vec![0.125, 0.5, 0.875]
        );
        assert_eq!(
            gradient
                .stops()
                .iter()
                .map(|stop| stop.color())
                .collect::<Vec<_>>(),
            vec![
                litchi_xlsx::style::cell_format::Color::rgb(0xFF, 0x00, 0x00),
                litchi_xlsx::style::cell_format::Color::rgb(0x00, 0xFF, 0x00),
                litchi_xlsx::style::cell_format::Color::rgb(0x00, 0x00, 0xFF),
            ]
        );

        let reopened = open_owner(&file.to_bytes().expect("publish path gradient"), eager);
        assert_eq!(
            cell_property(&reopened, CELL_C1, "fill").as_deref(),
            Some("gradient;#FF0000;#00FF00;#0000FF;0")
        );
    }
}

#[test]
fn gradient_fill_writes_support_linear_path_aliases_angles_and_tinted_stops() {
    let bytes = fixture();
    let inputs = [
        ("gradient;#FF0000;#00FF00;#0000FF;135", "135"),
        ("#FF0000-#00FF00-#0000FF-135", "135"),
        ("radial:#FF0000-#00FF00-#0000FF", "0"),
        ("#FF0000+tint-40-#00FF00+tint25-#0000FF-45", "45"),
    ];

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        for (input, degree) in inputs {
            let updated = original
                .set(CELL_A1, Props::new().with("fill", input))
                .unwrap_or_else(|error| panic!("gradient {input} should write: {error:?}"));
            let observed =
                cell_property(&updated, CELL_A1, "fill").expect("gradient ordinary readback");
            assert!(
                observed.starts_with("gradient;"),
                "gradient projection: {observed}"
            );
            assert!(
                observed.ends_with(&format!(";{degree}")),
                "gradient angle: {observed}"
            );
            let styles = styles_bytes(&updated);
            let styles = std::str::from_utf8(&styles).expect("gradient styles XML");
            assert!(styles.contains("gradientFill"), "gradient XML: {styles}");
            if input.contains("tint-40") {
                assert!(
                    styles.contains(r#"tint="-0.4""#),
                    "negative stop tint must remain native: {styles}"
                );
            }

            let repeated = updated
                .set(CELL_A1, Props::new().with("fill", input))
                .expect("repeat exact gradient assignment");
            assert_eq!(
                repeated.to_bytes().expect("repeated exact gradient bytes"),
                updated.to_bytes().expect("gradient bytes"),
                "repeating {input} is an exact no-op",
            );

            // OfficeCLI text readback omits path geometry and stop tints.
            // Only losslessly represented gradients can replay that projection.
            if !input.starts_with("radial:") && !input.contains("tint") {
                let repeated = updated
                    .set(CELL_A1, Props::new().with("fill", observed.as_str()))
                    .expect("repeat canonical gradient");
                assert_eq!(
                    repeated.to_bytes().expect("repeated gradient bytes"),
                    updated.to_bytes().expect("gradient bytes")
                );
            }
        }

        let path = original
            .set(
                CELL_A1,
                Props::new().with("fill", "radial:#FF0000-#00FF00-#0000FF"),
            )
            .expect("radial path gradient");
        let path_styles = styles_bytes(&path);
        let path_styles = std::str::from_utf8(&path_styles).expect("path gradient XML");
        assert!(path_styles.contains(r#"gradientFill type="path""#));
        assert!(
            path_styles.matches("<stop position=").count() >= 3,
            "radial alias retains all authored stops: {path_styles}"
        );
    }
}

#[test]
fn gradient_fill_converts_patterns_source_and_eager_with_ordered_batches() {
    let bytes = source_fill_fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original
            .to_bytes()
            .expect("original gradient conversion bytes");
        let gradient =
            cell_property(&original, CELL_C1, "fill").expect("source gradient projection");

        let opaque_error = original
            .set(CELL_A1, Props::new().with("fill", gradient.as_str()))
            .expect_err("replacing opaque source pattern must refuse loss");
        assert_eq!(opaque_error.code(), Code::UnsupportedCapability);
        assert_eq!(
            original
                .to_bytes()
                .expect("opaque source remains unchanged"),
            before
        );

        let commit = original
            .edit(|edit| {
                edit.set(CELL_B1, Props::new().with("fill", gradient.as_str()))?;
                edit.set(CELL_C1, Props::new().with("fillPattern", "darkGray"))?;
                edit.set(CELL_B1, props! { "value" => "gradient payload" })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("ordered gradient conversion");

        assert!(
            cell_property(commit.file(), CELL_B1, "fill")
                .is_some_and(|value| value.starts_with("gradient;")),
            "gradient conversion at B1"
        );
        assert_eq!(
            cell_property(commit.file(), CELL_C1, "fillPattern").as_deref(),
            Some("darkGray")
        );
        assert_eq!(
            cell_text(commit.file(), CELL_B1).as_deref(),
            Some("gradient payload")
        );

        let replayed = commit
            .patch()
            .apply(&original)
            .expect("replay gradient conversion");
        assert_eq!(
            replayed.to_bytes().expect("replayed gradient bytes"),
            commit.file().to_bytes().expect("gradient commit bytes")
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(&replayed)
            .expect("inverse gradient conversion");
        assert_eq!(
            restored.to_bytes().expect("restored gradient bytes"),
            before
        );

        let repeat = commit
            .file()
            .set(CELL_B1, Props::new().with("fill", gradient.as_str()))
            .expect("repeat gradient conversion");
        assert_eq!(
            repeat.to_bytes().expect("repeat gradient bytes"),
            commit.file().to_bytes().expect("first gradient bytes")
        );

        let reverse = commit
            .file()
            .edit(|edit| {
                edit.set(CELL_B1, Props::new().with("fillPattern", "solid"))?;
                edit.set(CELL_B1, Props::new().with("fill", gradient.as_str()))?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("later gradient replacement wins");
        assert!(
            cell_property(reverse.file(), CELL_B1, "fill")
                .is_some_and(|value| value.starts_with("gradient;"))
        );
    }
}

#[test]
fn invalid_gradient_values_and_companions_are_atomic() {
    let bytes = fixture();
    let invalid = [
        "gradient",
        "gradient;#FF0000",
        "gradient;#FF0000;not-a-color;45",
        "gradient;#FF0000;#00FF00;-1",
        "gradient;#FF0000;#00FF00;361",
        "gradient;#FF0000;#00FF00;NaN",
        "gradient;#FF0000;#00FF00;45;not-a-color",
    ];

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original
            .to_bytes()
            .expect("original invalid-gradient bytes");
        for input in invalid {
            let error = original
                .set(CELL_A1, Props::new().with("fill", input))
                .expect_err("invalid gradient must refuse");
            assert_value_refusal(error);
            assert_eq!(
                original
                    .to_bytes()
                    .expect("unchanged invalid-gradient owner"),
                before,
                "invalid gradient {input} must be atomic"
            );
        }

        for properties in [
            Props::new()
                .with("fill", "gradient;#FF0000;#00FF00;45")
                .with("fillPattern", "solid"),
            Props::new()
                .with("fill", "gradient;#FF0000;#00FF00;45")
                .with("fillBg", "#102030"),
            Props::new()
                .with("fill", "gradient;#FF0000;#00FF00;45")
                .with("fillPattern", PropValue::Null),
            Props::new()
                .with("fill", "gradient;#FF0000;#00FF00;45")
                .with("fillBg", PropValue::Null),
        ] {
            let error = original
                .set(CELL_A1, properties)
                .expect_err("gradient companions must refuse atomically");
            assert_eq!(error.code(), Code::InvalidValue);
            assert_eq!(
                original
                    .to_bytes()
                    .expect("unchanged gradient companion owner"),
                before
            );
        }

        let gradient = open_owner(&source_fill_fixture(), eager);
        let gradient_before = gradient.to_bytes().expect("gradient owner bytes");
        let cleared = gradient
            .set(CELL_C1, Props::new().with("fill", PropValue::Null))
            .expect("null fill clears a whole gradient");
        assert_eq!(cell_property(&cleared, CELL_C1, "fill"), None);
        let repeated = cleared
            .set(CELL_C1, Props::new().with("fill", PropValue::Null))
            .expect("repeat gradient clear");
        assert_eq!(
            repeated.to_bytes().expect("repeated gradient clear bytes"),
            cleared.to_bytes().expect("gradient clear bytes")
        );
        assert_ne!(
            cleared.to_bytes().expect("cleared gradient bytes"),
            gradient_before,
            "the first null clear removes the authored whole gradient"
        );

        let cleared_with_null_attributes = gradient
            .set(
                CELL_C1,
                Props::new()
                    .with("fill", PropValue::Null)
                    .with("fillPattern", PropValue::Null)
                    .with("fillBg", PropValue::Null),
            )
            .expect("whole-fill null may clear null gradient attributes together");
        assert_eq!(
            cell_property(&cleared_with_null_attributes, CELL_C1, "fill"),
            None,
            "fill=null clears the whole gradient even with null companions"
        );
        assert_ne!(
            cleared_with_null_attributes
                .to_bytes()
                .expect("whole-fill null with attributes bytes"),
            gradient_before,
            "whole-fill null with null companions must remove the gradient"
        );
    }
}

#[test]
fn fill_changes_preserve_non_solid_style_zero_for_new_cells_source_and_eager() {
    let bytes = source_base_pattern_fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let updated = original
            .set(
                "/sheet[1]/D2",
                props! {
                    "value" => "inherited pattern",
                    "fill" => "#AABBCC",
                },
            )
            .expect("new cell fill should preserve style-zero pattern");
        assert_eq!(
            cell_text(&updated, "/sheet[1]/D2").as_deref(),
            Some("inherited pattern")
        );
        assert_eq!(
            cell_property(&updated, "/sheet[1]/D2", "fill").as_deref(),
            Some("#AABBCC")
        );
        assert_eq!(
            cell_property(&updated, "/sheet[1]/D2", "fillPattern").as_deref(),
            Some("darkGray"),
            "a fill-only edit must retain the inherited non-solid pattern"
        );

        let reopened = open_owner(&updated.to_bytes().expect("publish inherited fill"), eager);
        assert_eq!(
            cell_property(&reopened, "/sheet[1]/D2", "fillPattern").as_deref(),
            Some("darkGray")
        );
    }
}

#[test]
fn fill_ranges_keep_mixed_source_patterns_payloads_and_replay_inverse() {
    let bytes = source_fill_fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("original mixed fill bytes");
        let commit = original
            .edit(|edit| {
                edit.set(
                    CELL_A1,
                    props! {
                        "value" => "range payload",
                    },
                )?;
                edit.set("/sheet[1]/A1:B1", Props::new().with("fill", "#AABBCC"))?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("mixed source pattern range edit");

        for path in [CELL_A1, CELL_B1] {
            assert_eq!(
                cell_property(commit.file(), path, "fill").as_deref(),
                Some("#AABBCC"),
                "range foreground at {path}"
            );
        }
        assert_eq!(
            cell_text(commit.file(), CELL_A1).as_deref(),
            Some("range payload")
        );
        assert_eq!(
            cell_property(commit.file(), CELL_A1, "fillPattern").as_deref(),
            Some("darkGray")
        );
        assert_eq!(
            cell_property(commit.file(), CELL_B1, "fillPattern").as_deref(),
            Some("lightGray")
        );
        assert_eq!(
            cell_property(commit.file(), CELL_A1, "fillBg").as_deref(),
            Some("accent1+tint40")
        );
        assert_eq!(
            cell_property(commit.file(), CELL_B1, "fillBg").as_deref(),
            Some("#F0E0D0")
        );

        let replayed = commit
            .patch()
            .apply(&original)
            .expect("replay mixed fill patch");
        assert_eq!(
            replayed.to_bytes().expect("replayed mixed fill bytes"),
            commit.file().to_bytes().expect("mixed fill bytes")
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(&replayed)
            .expect("inverse mixed fill patch");
        assert_eq!(
            restored.to_bytes().expect("restored mixed fill bytes"),
            before
        );

        let repeated = commit
            .file()
            .set("/sheet[1]/A1:B1", Props::new().with("fill", "#AABBCC"))
            .expect("repeat mixed fill range");
        assert_eq!(
            repeated.to_bytes().expect("repeated mixed fill bytes"),
            commit.file().to_bytes().expect("first mixed fill bytes")
        );
    }
}

#[test]
fn invalid_and_duplicate_fill_assignments_are_atomic() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("original fill owner");
        for value in [
            "not-a-color",
            "auto",
            "accent1+lumMod75",
            "accent7",
            "accent1+tint101",
            "accent1+tint=-101",
            "rgba(0,0,0,1.5)",
            "#12345",
        ] {
            let error = original
                .set(CELL_A1, Props::new().with("fill", value))
                .expect_err("invalid fill color");
            assert_eq!(error.code(), Code::InvalidValue, "fill {value}");
            assert_eq!(error.property(), Some("fill"), "fill {value}");
            assert_eq!(original.to_bytes().expect("unchanged fill owner"), before);
        }

        for value in ["diagonal", "gray60", "gray-60", "unknown", "gradient"] {
            let error = original
                .set(CELL_A1, Props::new().with("fillPattern", value))
                .expect_err("invalid fill pattern");
            assert_eq!(error.code(), Code::InvalidValue, "pattern {value}");
            assert_eq!(error.property(), Some("fillPattern"), "pattern {value}");
            assert_eq!(
                original.to_bytes().expect("unchanged pattern owner"),
                before
            );
        }

        for properties in [
            Props::new().with("fill", "red").with("fill", "blue"),
            Props::new()
                .with("fillPattern", "darkGray")
                .with("fillPattern", "lightGray"),
            Props::new().with("fill", "red").with("bgcolor", "blue"),
            Props::new().with("fill", "red").with("bg", "blue"),
            Props::new().with("bgcolor", "red").with("bg", "blue"),
            Props::new().with("fillBg", "red").with("fillBg", "blue"),
        ] {
            let error = original
                .set(CELL_A1, properties)
                .expect_err("duplicate fill assignments");
            assert_eq!(error.code(), Code::InvalidArgument);
            assert_eq!(
                original.to_bytes().expect("unchanged duplicate fill owner"),
                before
            );
        }

        let combined = original
            .set(
                CELL_A1,
                props! {
                    "value" => "must not publish",
                    "fillPattern" => "gray60",
                },
            )
            .expect_err("invalid pattern must reject payload atomically");
        assert_eq!(combined.code(), Code::InvalidValue);
        assert_eq!(
            original.to_bytes().expect("unchanged combined fill owner"),
            before
        );

        let cleared_default = original
            .set(CELL_A1, Props::new().with("fillPattern", PropValue::Null))
            .expect("clear authored default pattern");
        assert_eq!(
            cell_property(&cleared_default, CELL_A1, "fillPattern"),
            None,
            "ordinary pattern read remains absent after clearing the default token"
        );
        let repeated_clear = cleared_default
            .set(CELL_A1, Props::new().with("fillPattern", PropValue::Null))
            .expect("repeat default pattern clear");
        assert_eq!(
            repeated_clear.to_bytes().expect("repeated pattern clear"),
            cleared_default.to_bytes().expect("default pattern clear")
        );
    }
}

#[test]
fn stale_source_and_released_source_are_refused_before_cell_edit() {
    let directory = tempfile::tempdir().expect("temporary source directory");
    let path = directory.path().join("cells.xlsx");
    let bytes = fixture();
    fs::write(&path, &bytes).expect("write source fixture");

    let source = litchi_ops::open_with(&path, &OpenOptions::new().source_mode(SourceMode::Auto))
        .expect("open path-backed source owner");
    fs::write(&path, b"changed physical source").expect("change source after open");
    assert_eq!(
        source
            .set(CELL_A1, props! { "value" => "stale" })
            .expect_err("stale source must be fenced")
            .code(),
        Code::IoError
    );

    let released =
        litchi_ops::open_bytes_with(bytes.clone(), &OpenOptions::new().retain_source(false))
            .expect("open owner without retained source");
    assert_eq!(
        released
            .set(CELL_A1, props! { "value" => "released" })
            .expect_err("released source must refuse mutation")
            .code(),
        Code::UnsupportedCapability
    );

    let origin = directory.path().join("retained-origin.xlsx");
    fs::write(&origin, &bytes).expect("write retained source fixture");
    let retained =
        litchi_ops::open_with(&origin, &OpenOptions::new().source_mode(SourceMode::Auto))
            .expect("open retained source owner");
    let changed = retained
        .set(CELL_A1, props! { "value" => "saved" })
        .expect("edit retained source owner");
    let saved = changed.save().expect("save changed source owner at origin");
    assert_eq!(saved.path, origin);
    let reopened = litchi_ops::open(&origin).expect("reopen changed source origin");
    assert_eq!(cell_text(&reopened, CELL_A1).as_deref(), Some("saved"));
}

#[cfg(feature = "encryption")]
#[test]
fn unlocked_encrypted_xlsx_refuses_cell_edit_and_transaction_admission() {
    let password = "xlsx-cell-write-password";
    let encrypted = litchi_xlsx::Package::from_bytes(&fixture())
        .expect("open clear XLSX through native package")
        .to_encrypted(password, litchi_xlsx::encryption::Mode::Agile)
        .expect("encrypt XLSX fixture");
    let locked = litchi_ops::open_bytes(encrypted).expect("open encrypted XLSX");
    let File::Locked(locked) = locked else {
        panic!("encrypted XLSX must open locked");
    };
    let unlocked = locked
        .unlock(litchi_ops::Password::new(password.to_owned()))
        .expect("unlock encrypted XLSX");
    for verb in [
        litchi_ops::capabilities::Verb::Set,
        litchi_ops::capabilities::Verb::Edit,
        litchi_ops::capabilities::Verb::Batch,
        litchi_ops::capabilities::Verb::Save,
    ] {
        assert!(matches!(
            unlocked.capabilities().verb(verb),
            litchi_ops::capabilities::Support::Unsupported {
                code: Code::DocumentProtected,
                ..
            }
        ));
    }
    assert_eq!(
        unlocked.to_bytes().unwrap_err().code(),
        Code::DocumentProtected
    );
    assert_eq!(
        unlocked
            .set(CELL_A1, props! { "value" => "refused" })
            .expect_err("encrypted owner must refuse cell mutation")
            .code(),
        Code::DocumentProtected
    );
    assert_eq!(
        unlocked
            .edit::<()>(|_| panic!("protected XLSX transaction closure ran"))
            .expect_err("encrypted owner must refuse transaction admission")
            .code(),
        Code::DocumentProtected
    );
}
