//! Source freshness and identity policy at the unified XLSX boundary.
use super::{REL, SML, SheetPart, both, data_sheet, get, kitchen, small_sheet, styles, workbook};
use crate::{Code, File, OpenOptions};
use litchi_core::{ReadAt, SourceVersion};
use std::io;
use std::sync::{
    Arc,
    atomic::{AtomicU64, Ordering},
};

struct Source {
    bytes: Vec<u8>,
    revision: AtomicU64,
}
impl ReadAt for Source {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }
    fn read_at(&self, offset: u64, out: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset).unwrap_or(usize::MAX);
        let data = self.bytes.get(start..).unwrap_or_default();
        let count = out.len().min(data.len());
        out[..count].copy_from_slice(&data[..count]);
        Ok(count)
    }
    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(56, self.revision.load(Ordering::SeqCst)))
    }
}

#[test]
fn retained_nodes_and_lazy_reads_refuse_changed_sources() {
    let source = Arc::new(Source {
        bytes: kitchen(),
        revision: AtomicU64::new(0),
    });
    let package = litchi_opc::SourceBackedPackage::from_read_at(
        source.clone(),
        litchi_opc::SourceLimits::default(),
        litchi_opc::SourceCacheLimits::default(),
        litchi_opc::SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let handle = super::Xlsx::from_source_package(package, &OpenOptions::default()).unwrap();
    let file = File::Xlsx(handle);
    let sheet = file.get("/Data").unwrap();
    let cell = file.get("/Data/B2").unwrap();
    let _warm = cell.view().format().unwrap();
    assert!(cell.as_xlsx().unwrap().cell_format().unwrap().is_some());
    let mut children = sheet.children();
    assert!(children.next().unwrap().is_ok());
    let mut matches = file.query("cell").unwrap();
    assert!(matches.next().unwrap().is_ok());
    source.revision.store(1, Ordering::SeqCst);
    assert_eq!(file.root().unwrap_err().code(), Code::IoError);
    assert_eq!(cell.view().text().unwrap_err().code(), Code::IoError);
    assert_eq!(cell.view().format().unwrap_err().code(), Code::IoError);
    assert_eq!(
        cell.as_xlsx().unwrap().cell_format().unwrap_err().code(),
        Code::IoError
    );
    assert_eq!(
        sheet.as_xlsx().unwrap().sheet_name().unwrap_err().code(),
        Code::IoError
    );
    assert_eq!(children.next().unwrap().unwrap_err().code(), Code::IoError);
    assert!(children.next().is_none());
    assert_eq!(matches.next().unwrap().unwrap_err().code(), Code::IoError);
    assert!(matches.next().is_none());
    assert_eq!(
        file.view(crate::view::Mode::Text).unwrap_err().code(),
        Code::IoError
    );
    assert_eq!(
        file.view_json(crate::view::Mode::Stats).unwrap_err().code(),
        Code::IoError
    );
}

#[test]
fn typed_cell_format_preserves_gradient_geometry_and_all_color_kinds() {
    use litchi_xlsx::style::cell_format::{Color, GradientKind, Tint};

    let styles = styles().replace(
        r#"<patternFill patternType="lightGray"><fgColor theme="4" tint="0.4"/><bgColor rgb="FF0000FF"/></patternFill>"#,
        r#"<gradientFill type="path" degree="45.5" left="0.1" right="0.2" top="0.3" bottom="0.4"><stop position="0"><color rgb="80112233" tint="-0.25"/></stop><stop position="0.2"><color theme="4" tint="0.5"/></stop><stop position="0.65"><color indexed="12"/></stop><stop position="0.9"><color auto="0"/></stop><stop position="1"><color tint="0.1"/></stop></gradientFill>"#,
    );
    let bytes = workbook(
        &[small_sheet(
            "One",
            r#"<row r="1"><c r="A1" s="4"><v>1</v></c><c r="B1"><v>2</v></c><c r="D1" s="0"><v>3</v></c></row>"#,
        )],
        "",
        &[],
        &styles,
    );
    for handle in both(&bytes) {
        let file = File::Xlsx(handle);
        let cell = file.get("/One/A1").unwrap();
        let node = cell.as_xlsx().unwrap();
        let format = node.cell_format().unwrap().unwrap();
        assert!(std::ptr::eq(format, node.cell_format().unwrap().unwrap()));
        let gradient = format.fill().unwrap().gradient().unwrap();
        assert_eq!(gradient.kind(), GradientKind::Path);
        assert_eq!(gradient.degree(), 45.5);
        assert_eq!(gradient.left(), 0.1);
        assert_eq!(gradient.right(), 0.2);
        assert_eq!(gradient.top(), 0.3);
        assert_eq!(gradient.bottom(), 0.4);
        let stops = gradient.stops();
        assert_eq!(stops.len(), 5);
        for (stop, position) in stops.iter().zip([0.0, 0.2, 0.65, 0.9, 1.0]) {
            assert_eq!(stop.position(), position);
        }
        assert_eq!(
            stops[0].color(),
            Color::argb(0x80, 0x11, 0x22, 0x33).with_tint(Tint::new(-0.25).unwrap())
        );
        assert_eq!(
            stops[1].color(),
            Color::theme(4).with_tint(Tint::new(0.5).unwrap())
        );
        assert_eq!(stops[2].color(), Color::indexed(12));
        assert_eq!(stops[3].color(), Color::auto_value(false));
        assert_eq!(
            stops[4].color(),
            Color::default_base().with_tint(Tint::new(0.1).unwrap())
        );
        assert!(
            file.get("/One/D1")
                .unwrap()
                .as_xlsx()
                .unwrap()
                .cell_format()
                .unwrap()
                .unwrap()
                .is_base()
        );
        for path in ["/One", "/One/B1", "/One/C1"] {
            assert!(
                file.get(path)
                    .unwrap()
                    .as_xlsx()
                    .unwrap()
                    .cell_format()
                    .unwrap()
                    .is_none()
            );
        }
    }
}

#[test]
fn scoped_defined_names_require_unambiguous_selection() {
    let bytes = workbook(
        &[small_sheet("One", ""), small_sheet("Two", "")],
        r#"<definedNames><definedName name="Shared" localSheetId="0">One!$A$1</definedName><definedName name="SHARED" localSheetId="1">Two!$A$1</definedName></definedNames>"#,
        &[],
        &styles(),
    );
    for file in both(&bytes) {
        assert_eq!(
            get(&file, "/namedrange[shared]").unwrap_err().code(),
            Code::InvalidPath
        );
        assert_eq!(
            get(&file, "/namedrange[1]").unwrap().path().to_string(),
            "/namedrange[1]"
        );
        assert_eq!(
            get(&file, "/namedrange[2]").unwrap().path().to_string(),
            "/namedrange[2]"
        );
    }
}

#[test]
fn duplicate_table_names_do_not_select_the_first_match() {
    let mut first = data_sheet();
    let mut second = data_sheet();
    first.name = "One";
    second.name = "Two";
    let bytes = workbook(&[first, second], "", super::STRINGS, &styles());
    for file in both(&bytes) {
        let error = get(&file, "/table[@name=Sales]").unwrap_err();
        assert_eq!(error.code(), Code::InvalidPath);
        assert!(!format!("{error:?}").contains("Sales"));
        assert!(!error.to_string().contains("Sales"));
    }
}

#[test]
fn native_format_ids_require_explicit_raw_query() {
    for handle in both(&kitchen()) {
        let file = File::Xlsx(handle);
        let error = file.query("cell[numFmtId=164]").unwrap_err();
        assert_eq!(error.code(), Code::InvalidSelector);
        let raw = crate::query::raw::query("cell[numFmtId=164]").unwrap();
        let paths: Vec<_> = file
            .query(raw)
            .unwrap()
            .map(|node| node.unwrap().path().to_string())
            .collect();
        assert_eq!(paths, ["/sheet[1]/B2", "/sheet[1]/D2"]);
        let semantic: Vec<_> = file
            .query("cell[numberformat='0.00%']")
            .unwrap()
            .map(|node| node.unwrap().path().to_string())
            .collect();
        assert_eq!(semantic, paths);
        let node = file.get("/Data/B2").unwrap();
        assert!(
            node.view()
                .format()
                .unwrap()
                .all(|prop| prop.key() != "numFmtId")
        );
        assert!(
            !serde_json::to_string(&file.capabilities())
                .unwrap()
                .contains("numFmtId")
        );
    }
}

#[test]
fn malformed_table_bands_are_not_silently_empty_header_candidates() {
    let mut sheet = data_sheet();
    sheet.table = sheet.table.map(|xml| {
        xml.replace(
            "totalsRowShown=\"0\"",
            "totalsRowShown=\"1\" totalsRowCount=\"100\"",
        )
    });
    let bytes = workbook(&[sheet], "", super::STRINGS, &styles());
    for handle in both(&bytes) {
        let file = File::Xlsx(handle);
        let result = file
            .query("row[Sales>0]")
            .unwrap()
            .collect::<crate::Result<Vec<_>>>();
        assert_eq!(result.unwrap_err().code(), Code::CorruptFile);
    }
}

#[test]
fn unmodeled_cells_are_not_misclassified_as_numbers() {
    let bytes = workbook(
        &[small_sheet(
            "Data",
            r#"<row r="1"><c r="A1" t="vendor"><v>7</v></c></row>"#,
        )],
        "",
        &[],
        &styles(),
    );
    for handle in both(&bytes) {
        let file = File::Xlsx(handle);
        let node = file.get("/sheet[1]/A1").unwrap();
        let props: Vec<_> = node.view().format().unwrap().collect();
        assert_eq!(
            props
                .iter()
                .find(|prop| prop.key() == "type")
                .unwrap()
                .value()
                .as_str(),
            Some("Unknown")
        );
        assert_eq!(file.query("cell[type=Unknown]").unwrap().count(), 1);
    }
}

/// A source that counts its version observations and can be changed.
struct Counted {
    bytes: Vec<u8>,
    revision: AtomicU64,
    versions: AtomicU64,
}
impl ReadAt for Counted {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }
    fn read_at(&self, offset: u64, out: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset).unwrap_or(usize::MAX);
        let data = self.bytes.get(start..).unwrap_or_default();
        let count = out.len().min(data.len());
        out[..count].copy_from_slice(&data[..count]);
        Ok(count)
    }
    fn version(&self) -> io::Result<SourceVersion> {
        self.versions.fetch_add(1, Ordering::SeqCst);
        Ok(SourceVersion::new(57, self.revision.load(Ordering::SeqCst)))
    }
}

/// Observations of a header-row query's second full run over a `Data`
/// sheet with a `Sales` header and `rows` rows of 10, and its match count.
fn observed_header_query(rows: u32) -> (usize, u64) {
    let last = rows + 1;
    let mut xml = format!(
        r#"<worksheet xmlns="{SML}" xmlns:r="{REL}"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>"#
    );
    for row in 2..=last {
        xml.push_str(&format!(
            r#"<row r="{row}"><c r="A{row}" t="s"><v>2</v></c><c r="B{row}"><v>10</v></c></row>"#
        ));
    }
    xml.push_str(
        r#"</sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>"#,
    );
    let table = format!(
        r#"<table xmlns="{SML}" id="1" name="Sales" displayName="Sales" ref="A1:B{last}" totalsRowShown="0"><autoFilter ref="A1:B{last}"/><tableColumns count="2"><tableColumn id="1" name="Region"/><tableColumn id="2" name="Sales"/></tableColumns></table>"#
    );
    let sheet = SheetPart {
        name: "Data",
        state: None,
        xml,
        table: Some(table),
    };
    let source = Arc::new(Counted {
        bytes: workbook(&[sheet], "", &["Region", "Sales", "East"], &styles()),
        revision: AtomicU64::new(0),
        versions: AtomicU64::new(0),
    });
    let package = litchi_opc::SourceBackedPackage::from_read_at(
        source.clone(),
        litchi_opc::SourceLimits::default(),
        litchi_opc::SourceCacheLimits::default(),
        litchi_opc::SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let handle = super::Xlsx::from_source_package(package, &OpenOptions::default()).unwrap();
    let file = File::Xlsx(handle);
    let matches = file.query("row[Sales>5]").unwrap().count();
    let before = source.versions.load(Ordering::SeqCst);
    assert_eq!(file.query("row[Sales>5]").unwrap().count(), matches);
    (matches, source.versions.load(Ordering::SeqCst) - before)
}

/// Record 0310: a header-row query probes each row of the bound block
/// through the sheet's cached grid. Inside the query step those probes
/// share the step's closing check instead of re-entering the leaf's fenced
/// grid read, so each further matching row costs its one per-match check
/// (owner decision 2026-10-14) and nothing more. A change between two
/// pulls still fails the next pull.
#[test]
fn header_row_probes_share_the_step_check() {
    let (small_matches, small) = observed_header_query(10);
    let (large_matches, large) = observed_header_query(20);
    assert_eq!((small_matches, large_matches), (10, 20));
    assert_eq!(large - small, 10, "{small} -> {large}");

    let source = Arc::new(Counted {
        bytes: kitchen(),
        revision: AtomicU64::new(0),
        versions: AtomicU64::new(0),
    });
    let package = litchi_opc::SourceBackedPackage::from_read_at(
        source.clone(),
        litchi_opc::SourceLimits::default(),
        litchi_opc::SourceCacheLimits::default(),
        litchi_opc::SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let handle = super::Xlsx::from_source_package(package, &OpenOptions::default()).unwrap();
    let file = File::Xlsx(handle);
    let paths = file
        .query("row[Sales>1000]")
        .unwrap()
        .map(|node| node.unwrap().path().to_string())
        .collect::<Vec<_>>();
    assert_eq!(paths, ["/sheet[1]/row[2]", "/sheet[1]/row[5]"]);
    let mut matches = file.query("row[Sales>1000]").unwrap();
    assert!(matches.next().unwrap().is_ok());
    source.revision.store(1, Ordering::SeqCst);
    assert_eq!(matches.next().unwrap().unwrap_err().code(), Code::IoError);
    assert!(matches.next().is_none());
}
