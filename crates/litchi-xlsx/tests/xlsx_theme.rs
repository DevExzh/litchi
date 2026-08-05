//! Tests for the XLSX theme part reader against real workbooks.

use litchi_ooxml::xlsx::theme::{Theme, ThemeColorSlot, ThemeColorValue};
use litchi_ooxml::xlsx::{Workbook, template};
use std::path::PathBuf;

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ooxml/xlsx")
        .join(name)
}

#[test]
fn parses_the_default_writer_theme_template() {
    let theme = Theme::parse(template::default_theme_xml()).unwrap();
    assert_eq!(theme.color_scheme_name(), "Office");
    assert!(theme.major_font().is_some());
    assert!(theme.minor_font().is_some());
    assert!(!theme.format_scheme_xml().is_empty());
    for slot in ThemeColorSlot::ALL {
        let _ = theme.rgb(slot);
    }
}

#[test]
fn reads_theme_parts_from_real_workbooks() {
    for name in ["cell-borders.xlsx"] {
        let workbook = Workbook::new(litchi_opc::OpcPackage::open(fixture(name)).unwrap()).unwrap();
        let theme = workbook
            .theme()
            .unwrap_or_else(|error| panic!("{name}: {error}"))
            .unwrap_or_else(|| panic!("{name} has no theme part"));
        // Every theme slot resolves to a concrete RGB triple.
        for slot in ThemeColorSlot::ALL {
            let _ = theme.rgb(slot);
        }
        assert!(matches!(
            theme.color(ThemeColorSlot::Dk1),
            ThemeColorValue::Srgb(_) | ThemeColorValue::System { .. }
        ));
    }
}

#[test]
fn workbooks_without_theme_parts_report_none() {
    for name in ["autofilter.xlsx", "column_style.xlsx"] {
        let workbook = Workbook::new(litchi_opc::OpcPackage::open(fixture(name)).unwrap()).unwrap();
        assert!(workbook.theme().unwrap().is_none(), "{name}");
    }
}
