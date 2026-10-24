//! Complete tab identity and formula provenance for canonical read adapters.
#![allow(clippy::unwrap_used, reason = "acceptance assertions fail fast")]
use litchi_xls::{Builder, Package, SourceBackedLimits};
use std::sync::Arc;

fn source(bytes: Vec<u8>) -> Package {
    let cfb = Arc::new(
        litchi_cfb::SharedOleFile::open(
            Arc::new(litchi_core::OwnedSource::new(bytes)),
            litchi_cfb::SharedOleFileLimits::default(),
        )
        .unwrap(),
    );
    Package::from_source(
        litchi_xls::raw::source_backed_workbook_from_shared_ole_file(
            cfb,
            SourceBackedLimits::default(),
        )
        .unwrap(),
    )
    .unwrap()
}

#[test]
fn complete_catalog_matches_eager_visibility_and_chart_tab_identity() {
    for name in ["45761.xls", "44010-TwoCharts.xls"] {
        let bytes = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/poi/test-data/spreadsheet")
                .join(name),
        )
        .unwrap();
        let eager = Package::from_bytes(&bytes).unwrap();
        let source = source(bytes);
        assert_eq!(
            source.workbook().sheets().unwrap(),
            eager.workbook().sheets().unwrap()
        );
        for package in [&eager, &source] {
            let workbook = package.workbook();
            let catalog = workbook.sheets().unwrap();
            assert_eq!(workbook.sheet_count().unwrap(), 3);
            for (position, metadata) in catalog.iter().enumerate() {
                assert_eq!(workbook.sheet(position).unwrap().as_ref(), Some(metadata));
                assert_eq!(
                    workbook
                        .sheet(metadata.name().to_ascii_lowercase().as_str())
                        .unwrap()
                        .as_ref(),
                    Some(metadata)
                );
            }
            assert!(workbook.sheet(3).unwrap().is_none());
            assert!(workbook.sheet("no such tab").unwrap().is_none());
        }
        if name == "44010-TwoCharts.xls" {
            let sheets = source.workbook().sheets().unwrap();
            assert_eq!(sheets[1].kind(), litchi_xls::SheetKind::ChartSheet);
            assert_eq!(sheets[1].parsed_worksheet_index(), None);
            assert_eq!(source.workbook().worksheet_count().unwrap(), 1);
        }
    }
}

#[test]
fn formula_provenance_survives_scan_visit_and_index_replay() {
    let mut builder = Builder::new();
    let sheet = builder.add_worksheet("Data").unwrap();
    builder.write_number(sheet, 0, 0, 5.0).unwrap();
    builder.write_formula(sheet, 1, 0, "A1*2").unwrap();
    let source = source(builder.build().unwrap().to_bytes().unwrap());
    let worksheet = source
        .source_workbook()
        .unwrap()
        .worksheet_by_index(0)
        .unwrap()
        .unwrap();
    for _ in 0..3 {
        assert!(!worksheet.cell(0, 0).unwrap().unwrap().is_formula());
        assert!(worksheet.cell(1, 0).unwrap().unwrap().is_formula());
    }
    let mut flags = Vec::new();
    worksheet
        .visit_cells(|cell| {
            flags.push((cell.row(), cell.is_formula()));
            Ok(())
        })
        .unwrap();
    assert_eq!(flags, [(0, false), (1, true)]);
}
