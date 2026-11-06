//! Saved native Excel view and print metadata, without rendering or printing.

use litchi_sheet::view::{Mode, Position, State};
use litchi_xlsx::{Package, SourceBackedWorkbook};
use std::{collections::BTreeMap, path::PathBuf};

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/excel")
        .join(name)
}

fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let reader = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    reader
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let bytes = reader.read_member(&name).unwrap();
            (name, bytes)
        })
        .collect()
}

fn assert_only_sheet_changed(before: &[u8], after: &[u8]) {
    let a = members(before);
    let b = members(after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    let changed = a
        .iter()
        .filter(|(name, bytes)| **bytes != b[*name])
        .map(|(name, _)| name.as_str())
        .collect::<Vec<_>>();
    assert_eq!(changed, ["xl/worksheets/sheet1.xml"]);
}

#[test]
fn native_print_headers_footers_and_default_gridline_flag_are_typed() {
    let package = Package::open(fixture("excel-interaction-print.xlsx")).unwrap();
    let workbook = package.workbook();
    let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
    let settings = sheet.header_footer().unwrap().unwrap();
    assert!(!settings.different_odd_even() && !settings.different_first());
    assert!(settings.scale_with_document() && settings.align_with_margins());
    assert!(settings.even_header().is_none() && settings.even_footer().is_none());
    assert!(settings.first_header().is_none() && settings.first_footer().is_none());
    let header = settings.odd_header().unwrap();
    let footer = settings.odd_footer().unwrap();
    let font = "&\"Aptos Narrow,Regular\"&K000000";
    assert_eq!(
        header.left(),
        Some(format!("{font}HEADER_LEFT_002").as_str())
    );
    assert_eq!(
        header.center(),
        Some(format!("{font}Native Excel print fixture").as_str())
    );
    assert_eq!(header.right(), Some(format!("{font}&A").as_str()));
    assert_eq!(footer.left(), Some(format!("{font}FOOTER_002").as_str()));
    assert_eq!(
        footer.center(),
        Some(format!("{font}Page &P of &N").as_str())
    );
    assert_eq!(footer.right(), Some(format!("{font}&F").as_str()));
    let options = sheet.print_options().unwrap().unwrap();
    assert!(options.horizontal_centered() && options.print_headings());
    assert!(!options.vertical_centered());
    assert!(options.grid_lines());
    // CT_PrintOptions defaults gridLinesSet to true, even though Word/Excel
    // Boolean attributes often default to false. This native part omits it.
    assert!(options.grid_lines_set());
    assert!(options.prints_grid_lines());
    assert!(
        package
            .edit_printer_settings(0_usize)
            .unwrap()
            .printer_settings()
            .unwrap()
            .is_none()
    );
    let reopened = Package::from_bytes(&package.to_bytes().unwrap()).unwrap();
    let reopened_workbook = reopened.workbook();
    let reopened_sheet = reopened_workbook.sheet("Sheet1").unwrap().unwrap();
    assert_eq!(reopened_sheet.header_footer().unwrap().unwrap(), settings);
    assert_eq!(reopened_sheet.print_options().unwrap().unwrap(), options);
}

#[test]
fn native_print_option_edit_changes_only_its_element_and_restores_exact_source() {
    let source = SourceBackedWorkbook::open(fixture("excel-interaction-print.xlsx")).unwrap();
    let before = source.to_bytes().unwrap();
    let mut noop = source.edit_print_options(0_usize).unwrap();
    let original = *noop.print_options().unwrap();
    assert!(original.prints_grid_lines());
    assert!(!noop.set(original));
    assert_eq!(
        noop.commit().unwrap().into_package().to_bytes().unwrap(),
        before
    );
    let mut edit = source.edit_print_options(0_usize).unwrap();
    let mut settings = original;
    settings
        .set_print_grid_lines(false)
        .set_vertical_centered(true);
    assert!(edit.set(settings));
    let (updated, patch) = edit.commit().unwrap().into_parts();
    let after = updated.to_bytes().unwrap();
    assert_only_sheet_changed(&before, &after);
    let source_xml =
        String::from_utf8(members(&before)["xl/worksheets/sheet1.xml"].clone()).unwrap();
    let expected = source_xml.replace(
        "<printOptions horizontalCentered=\"1\" headings=\"1\" gridLines=\"1\"/>",
        "<printOptions horizontalCentered=\"1\" verticalCentered=\"1\" headings=\"1\" gridLines=\"0\" gridLinesSet=\"0\"/>",
    );
    assert_ne!(expected, source_xml);
    assert!(
        members(&after)["xl/worksheets/sheet1.xml"] == expected.as_bytes(),
        "print-option edit changed bytes outside the selected element"
    );
    let output = tempfile::NamedTempFile::with_suffix(".xlsx").unwrap();
    updated.save(output.path()).unwrap();
    let reopened = Package::open(output.path()).unwrap();
    let workbook = reopened.workbook();
    let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
    assert_eq!(sheet.print_options().unwrap().unwrap(), settings);
    assert!(!sheet.print_options().unwrap().unwrap().prints_grid_lines());
    let restored = updated
        .apply_print_options(&patch.inverse())
        .unwrap()
        .into_package();
    assert!(restored.to_bytes().unwrap() == before);
    assert!(source.to_bytes().unwrap() == before);
}

#[test]
fn native_frozen_panes_and_saved_selections_match_eager_and_source_readers() {
    for (name, active, range, selected) in [
        ("excel-values-formulas-chart.xlsx", "B2", "B2", false),
        ("excel-interaction-print.xlsx", "B2", "B2:I5", true),
        ("excel-date1904-breaks.xlsx", "B6", "B6", true),
    ] {
        let package = Package::open(fixture(name)).unwrap();
        let workbook = package.workbook();
        let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
        let views = sheet.views().unwrap().unwrap();
        assert_eq!(views.entries().len(), 1);
        let view = views.entries()[0].view();
        assert_eq!(view.window.get(), 0);
        assert_eq!(view.mode, Mode::Normal);
        assert_eq!(view.zoom.current.get(), 100);
        assert_eq!(view.tab_selected, selected);
        assert!(view.display.grid_lines && view.display.row_column_headers);
        if name == "excel-values-formulas-chart.xlsx" {
            let pane = view.pane.as_ref().unwrap();
            assert_eq!(pane.position, Position::BottomRight);
            assert_eq!(pane.state, State::Frozen);
            assert_eq!(pane.horizontal.unwrap().get(), 1.0);
            assert_eq!(pane.vertical.unwrap().get(), 1.0);
            assert_eq!(pane.top_left.to_string(), "B2");
            assert_eq!(view.selections.len(), 3);
            for (selection, position, address) in [
                (&view.selections[0], Position::TopRight, "B1"),
                (&view.selections[1], Position::BottomLeft, "A2"),
                (&view.selections[2], Position::BottomRight, "B2"),
            ] {
                assert_eq!(selection.position(), position);
                assert_eq!(selection.active_cell().to_string(), address);
                assert_eq!(selection.active_range(), 0);
                assert_eq!(selection.ranges().len(), 1);
                assert_eq!(selection.ranges()[0].to_string(), address);
            }
        } else {
            assert!(view.pane.is_none());
            assert_eq!(view.selections.len(), 1);
        }
        let selection = view.selections.last().unwrap();
        assert_eq!(selection.active_cell().to_string(), active);
        assert_eq!(selection.ranges()[0].to_string(), range);
        assert!(sheet.named_sheet_views().unwrap().is_none());
        let source = SourceBackedWorkbook::open(fixture(name)).unwrap();
        assert_eq!(
            source
                .sheet("Sheet1")
                .unwrap()
                .unwrap()
                .views()
                .unwrap()
                .unwrap(),
            views
        );
    }
}
