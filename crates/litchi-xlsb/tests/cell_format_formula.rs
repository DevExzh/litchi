#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::float_cmp,
    reason = "integration tests use panic-on-failure extraction and exact fixture values"
)]

//! Partial cell formats and formula text through package cell edits
//! (record 0261).

use litchi_core::sheet::{Cell as _, CellValue};
use litchi_xlsb::builder::{self, Builder};
use litchi_xlsb::cell_values::{BorderLine, Format, Reference};
use litchi_xlsb::package::PackageError;
use litchi_xlsb::package::styles as read;
use litchi_xlsb::styles::{BorderStyle, HorizontalAlignment, VerticalAlignment};
use litchi_xlsb::{Package, Workbook};

fn package() -> Package {
    let mut sheet = builder::Worksheet::new("Data");
    sheet.set_cell(0, 0, 1.5_f64);
    sheet.set_cell(0, 1, "label");
    sheet.set_cell(1, 0, 2.0_f64);
    let mut builder = Builder::new();
    builder.add_worksheet(sheet);
    builder.add_worksheet(builder::Worksheet::new("Other"));
    builder.build().unwrap()
}

fn at(row: u32, column: u32) -> Reference {
    Reference::new(row, column).unwrap()
}

fn style(workbook: &Workbook, row: u32, column: u32) -> u32 {
    workbook
        .worksheet(0)
        .unwrap()
        .get_cell(row, column)
        .map_or(0, litchi_xlsb::Cell::style_id)
}

fn format_all() -> Format {
    let mut format = Format::default();
    format.bold = Some(true);
    format.italic = Some(true);
    format.underline = Some(true);
    format.font_size = Some(14.0);
    format.font_name = Some("Arial".to_owned());
    format.font_color = Some(Some(0xFF1F_4E79));
    format.fill = Some(Some(0xFFFF_EE00));
    format.border_bottom = Some(Some(BorderLine {
        style: BorderStyle::Thick,
        color: Some(0xFF00_0000),
    }));
    format.number_format = Some("0.000".to_owned());
    format.horizontal = Some(HorizontalAlignment::Center);
    format.vertical = Some(VerticalAlignment::Top);
    format.wrap_text = Some(true);
    format.rotation = Some(45);
    format.indent = Some(2);
    format
}

#[test]
fn a_format_change_is_interned_and_reads_back() {
    let source = package();
    let mut edit = source
        .edit_cell_values(0_usize)
        .unwrap()
        .resolve()
        .unwrap()
        .unwrap();
    assert!(edit.set_format(at(0, 0), &format_all()).unwrap());
    // A blank coordinate becomes a formatted blank cell with the same XF.
    assert!(edit.set_format(at(4, 4), &format_all()).unwrap());
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let (after, patch) = commit.into_parts();

    let workbook = after.workbook().unwrap();
    let xf_index = style(workbook, 0, 0);
    assert_ne!(xf_index, 0);
    assert_eq!(style(workbook, 4, 4), xf_index, "one interned XF");
    assert_eq!(
        style(workbook, 1, 0),
        style(source.workbook().unwrap(), 1, 0)
    );
    let styles = workbook.styles();
    let xf = styles.get_cell_format(xf_index as usize).unwrap();
    let font = styles.get_font(xf.font_id as usize).unwrap();
    assert!(font.bold && font.italic && font.underline && !font.strike);
    assert_eq!(font.size, 14.0);
    assert_eq!(font.name, "Arial");
    assert_eq!(font.color, Some(0xFF1F_4E79));
    let fill = styles.get_fill(xf.fill_id as usize).unwrap();
    assert_eq!((fill.pattern_type, fill.fg_color), (1, Some(0xFFFF_EE00)));
    let border = styles.get_border(xf.border_id as usize).unwrap();
    let bottom = border.bottom.as_ref().unwrap();
    assert_eq!(
        (bottom.style, bottom.color),
        (read::BorderStyle::Thick, Some(0xFF00_0000))
    );
    assert!(border.top.is_none());
    assert_eq!(styles.get_num_fmt(xf.num_fmt_id), Some("0.000"));
    let alignment = xf.alignment.as_ref().unwrap();
    assert_eq!(alignment.horizontal, read::HorizontalAlignment::Center);
    assert_eq!(alignment.vertical, read::VerticalAlignment::Top);
    assert!(alignment.wrap_text);
    assert_eq!((alignment.rotation, alignment.indent), (45, 2));

    // The values are untouched and the patch replays and inverts exactly.
    assert_eq!(
        workbook
            .worksheet(0)
            .unwrap()
            .get_cell(0, 0)
            .unwrap()
            .value(),
        &CellValue::Float(1.5)
    );
    assert_eq!(
        source
            .apply_cell_values(&patch)
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        after.to_bytes().unwrap()
    );
    assert_eq!(
        after
            .apply_cell_values(&patch.inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        source.to_bytes().unwrap()
    );
}

#[test]
fn a_partial_change_keeps_every_other_attribute() {
    let source = package();
    let mut first = source
        .edit_cell_values(0_usize)
        .unwrap()
        .resolve()
        .unwrap()
        .unwrap();
    first.set_format(at(0, 0), &format_all()).unwrap();
    let styled = first.commit().unwrap().into_package();
    let mut second = styled
        .edit_cell_values(0_usize)
        .unwrap()
        .resolve()
        .unwrap()
        .unwrap();
    let mut format = Format::default();
    format.bold = Some(false);
    format.border_bottom = Some(None);
    assert!(second.set_format(at(0, 0), &format).unwrap());
    let after = second.commit().unwrap().into_package();
    let workbook = after.workbook().unwrap();
    let styles = workbook.styles();
    let xf = styles
        .get_cell_format(style(workbook, 0, 0) as usize)
        .unwrap();
    let font = styles.get_font(xf.font_id as usize).unwrap();
    assert!(!font.bold && font.italic && font.underline);
    assert_eq!(font.name, "Arial");
    assert!(
        styles
            .get_border(xf.border_id as usize)
            .unwrap()
            .bottom
            .is_none()
    );
    assert_eq!(
        styles.get_fill(xf.fill_id as usize).unwrap().fg_color,
        Some(0xFFFF_EE00)
    );
    assert_eq!(styles.get_num_fmt(xf.num_fmt_id), Some("0.000"));
}

#[test]
fn an_equal_format_is_an_exact_no_op() {
    let source = package();
    let mut first = source
        .edit_cell_values(0_usize)
        .unwrap()
        .resolve()
        .unwrap()
        .unwrap();
    first.set_format(at(0, 0), &format_all()).unwrap();
    let styled = first.commit().unwrap().into_package();
    let mut again = styled
        .edit_cell_values(0_usize)
        .unwrap()
        .resolve()
        .unwrap()
        .unwrap();
    assert!(!again.set_format(at(0, 0), &format_all()).unwrap());
    assert!(!again.set_format(at(0, 0), &Format::default()).unwrap());
    let commit = again.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(
        commit.into_package().to_bytes().unwrap(),
        styled.to_bytes().unwrap()
    );
}

#[test]
fn built_in_number_formats_keep_their_identifier_and_invalid_values_are_refused() {
    let source = package();
    let mut edit = source
        .edit_cell_values(0_usize)
        .unwrap()
        .resolve()
        .unwrap()
        .unwrap();
    let mut format = Format::default();
    format.number_format = Some("0.00%".to_owned());
    edit.set_format(at(0, 0), &format).unwrap();
    let after = edit.commit().unwrap().into_package();
    let workbook = after.workbook().unwrap();
    let xf = workbook
        .styles()
        .get_cell_format(style(workbook, 0, 0) as usize)
        .unwrap()
        .clone();
    assert_eq!(xf.num_fmt_id, 10);

    let mut edit = source
        .edit_cell_values(0_usize)
        .unwrap()
        .resolve()
        .unwrap()
        .unwrap();
    for invalid in [
        {
            let mut f = Format::default();
            f.font_size = Some(0.5);
            f
        },
        {
            let mut f = Format::default();
            f.font_name = Some(String::new());
            f
        },
        {
            let mut f = Format::default();
            f.rotation = Some(200);
            f
        },
    ] {
        assert!(matches!(
            edit.set_format(at(0, 0), &invalid),
            Err(PackageError::InvalidFormat(_))
        ));
    }
    assert!(edit.commit().unwrap().patch().is_noop());
}

fn formula(workbook: &Workbook, row: u32, column: u32) -> (String, Option<CellValue>) {
    let sheet = workbook.worksheet(0).unwrap();
    let cell = sheet.get_cell(row, column).unwrap();
    match cell.value() {
        CellValue::Formula {
            formula,
            cached_value,
            ..
        } => (formula.clone(), cached_value.as_deref().cloned()),
        other => panic!("not a formula: {other:?}"),
    }
}

#[test]
fn formula_text_compiles_into_new_value_and_formula_cells() {
    let source = package();
    let mut edit = source
        .edit_cell_values(0_usize)
        .unwrap()
        .resolve()
        .unwrap()
        .unwrap();
    assert!(edit.set_formula_text(at(2, 0), "=SUM(A1:A2)*2").unwrap());
    assert!(edit.set_formula_text(at(1, 0), "A1+1").unwrap());
    let after = edit.commit().unwrap().into_package();
    let workbook = after.workbook().unwrap();
    let (text, cache) = formula(workbook, 2, 0);
    // Reads render binary operators fully parenthesized.
    assert_eq!(text, "(SUM(A1:A2)*2)");
    assert_eq!(cache, Some(CellValue::String(String::new())));
    let (text, cache) = formula(workbook, 1, 0);
    assert_eq!(text, "(A1+1)");
    assert_eq!(cache, Some(CellValue::Float(2.0)), "the value is the cache");
    let sheet = workbook.worksheet(0).unwrap();
    assert_eq!(sheet.get_cell(2, 0).unwrap().formula_flags(), Some(0x0002));

    // Replacing the formula text keeps the cache; the same text is a no-op.
    let mut edit = after
        .edit_cell_values(0_usize)
        .unwrap()
        .resolve()
        .unwrap()
        .unwrap();
    assert!(!edit.set_formula_text(at(1, 0), "A1+1").unwrap());
    assert!(edit.set_formula_text(at(1, 0), "A1+2").unwrap());
    let replaced = edit.commit().unwrap().into_package();
    let (text, cache) = formula(replaced.workbook().unwrap(), 1, 0);
    assert_eq!(text, "(A1+2)");
    assert_eq!(cache, Some(CellValue::Float(2.0)));
}

#[test]
fn formulas_needing_workbook_context_are_refused_before_staging() {
    let source = package();
    let mut edit = source
        .edit_cell_values(0_usize)
        .unwrap()
        .resolve()
        .unwrap()
        .unwrap();
    for text in ["Other!A1", "SUM(Missing)", "=("] {
        let error = edit.set_formula_text(at(3, 3), text).unwrap_err();
        assert!(
            matches!(
                error,
                PackageError::Refused(_) | PackageError::InvalidFormula(_)
            ),
            "{text}: {error:?}"
        );
    }
    assert!(edit.commit().unwrap().patch().is_noop());
}
