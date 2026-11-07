//! Native worksheet metadata readers and detached codecs; no evaluation.
use litchi_xlsx::cell_watches::{
    CellWatchConformance as Conformance, CellWatchReference, CellWatches, parse_cell_watches,
    write_cell_watches,
};
use litchi_xlsx::data_consolidation::{self, DataConsolidation, Function, ReferenceSource};
use litchi_xlsx::ignored_errors::{IgnoredErrorType as ErrorKind, parse_worksheet_ignored_errors};
use litchi_xlsx::{Address, Cell, ErrorValue, Package, SourceBackedWorkbook, Value};
use std::{collections::BTreeMap, io::Cursor};
fn native() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel/excel-metadata.xlsx"),
    )
    .unwrap()
}
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let r = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    r.member_names()
        .unwrap()
        .into_iter()
        .map(|n| {
            let b = r.read_member(&n).unwrap();
            (n, b)
        })
        .collect()
}
fn fragment(xml: &[u8], name: &[u8]) -> Vec<u8> {
    use quick_xml::events::Event;
    let mut r = quick_xml::Reader::from_reader(xml);
    let mut start = None;
    let mut depth = 0usize;
    loop {
        let at = r.buffer_position() as usize;
        let e = r.read_event().unwrap();
        let end = r.buffer_position() as usize;
        match e {
            Event::Start(e) => {
                if start.is_none() && e.name().as_ref() == name {
                    start = Some((at, depth));
                }
                depth += 1;
            },
            Event::Empty(e) if start.is_none() && e.name().as_ref() == name => {
                return xml[at..end].to_vec();
            },
            Event::End(_) => {
                depth -= 1;
                if start.is_some_and(|(_, d)| d == depth) {
                    return xml[start.unwrap().0..end].to_vec();
                }
            },
            Event::Eof => panic!("missing XML fragment"),
            _ => {},
        }
    }
}
fn caches(p: &Package) {
    let b = p.workbook();
    let s = b.sheet("Sheet1").unwrap().unwrap();
    for (a, formula, value) in [
        ("A4", "A3+1", Value::Error(ErrorValue::Value)),
        ("A5", "1/0", Value::Error(ErrorValue::DivZero)),
        (
            "A7",
            "SUM(10,20)",
            Value::Number(litchi_xlsx::Number::new("30").unwrap()),
        ),
    ] {
        let c = s.cell(a).unwrap();
        let Some(Cell::Formula(f)) = c.stored() else {
            panic!("cached formula")
        };
        assert_eq!(f.text(), formula);
        assert_eq!(f.cached().unwrap().value(), &value);
    }
}

#[test]
fn native_ignored_errors_and_a7_watch_keep_error_kinds_and_saved_cache_values() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let b = p.workbook();
    let s = b.sheet("Sheet1").unwrap().unwrap();
    let errors = s.ignored_errors().unwrap().unwrap();
    assert_eq!(errors.entries().len(), 1);
    assert!(errors.extensions().is_empty());
    let e = &errors.entries()[0];
    assert_eq!(
        e.ranges().iter().map(|r| r.as_str()).collect::<Vec<_>>(),
        ["A4:A5"]
    );
    for kind in [
        ErrorKind::CalculatedColumn,
        ErrorKind::EmptyCellReference,
        ErrorKind::EvaluationError,
        ErrorKind::Formula,
        ErrorKind::FormulaRange,
        ErrorKind::ListDataValidation,
        ErrorKind::NumberStoredAsText,
        ErrorKind::TwoDigitTextYear,
        ErrorKind::UnlockedFormula,
    ] {
        assert_eq!(e.ignores(kind), kind == ErrorKind::EvaluationError);
    }
    let parts = members(&bytes);
    assert_eq!(
        parse_worksheet_ignored_errors(&parts["xl/worksheets/sheet1.xml"]).unwrap(),
        Some(errors.clone())
    );
    let watches = parse_cell_watches(&parts["xl/worksheets/sheet1.xml"])
        .unwrap()
        .unwrap();
    assert_eq!(
        watches
            .references()
            .iter()
            .map(|r| r.as_str())
            .collect::<Vec<_>>(),
        ["A7"]
    );
    caches(&p);
    assert!(members(&p.to_bytes().unwrap()) == parts);
    for (mode, ns) in [
        (
            Conformance::Transitional,
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        ),
        (
            Conformance::Strict,
            "http://purl.oclc.org/ooxml/spreadsheetml/main",
        ),
    ] {
        let xml = write_cell_watches(&watches, mode).unwrap();
        let wrapped = format!("<worksheet xmlns=\"{ns}\">{xml}</worksheet>");
        assert_eq!(
            parse_cell_watches(wrapped.as_bytes()).unwrap(),
            Some(watches.clone())
        );
    }
    let detached = CellWatches::new(vec![
        CellWatchReference::new("A4").unwrap(),
        CellWatchReference::new("A7").unwrap(),
    ])
    .unwrap();
    assert!(CellWatchReference::new("A1:B2").is_err());
    let xml = write_cell_watches(&detached, Conformance::Transitional).unwrap();
    let wrapped = format!(
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">{xml}</worksheet>"
    );
    assert_eq!(
        parse_cell_watches(wrapped.as_bytes()).unwrap(),
        Some(detached)
    );
    assert!(members(&p.to_bytes().unwrap()) == parts);
}

#[test]
fn native_consolidation_sources_and_detached_writing_preserve_saved_formula_results() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let b = p.workbook();
    let sheet = b.sheet("Consolidated_015").unwrap().unwrap();
    let v = sheet.data_consolidation().unwrap().unwrap();
    assert_eq!(v.function(), Function::Sum);
    assert!(v.left_labels() && v.top_labels() && v.link());
    assert!(!v.start_labels());
    let refs = v.data_references().unwrap();
    assert_eq!(refs.declared_count(), Some(2));
    assert_eq!(refs.references().len(), 2);
    for (r, name) in refs
        .references()
        .iter()
        .zip(["Source_A_015", "Source_B_015"])
    {
        assert_eq!(r.relationship_id(), None);
        let ReferenceSource::Range { sheet, reference } = r.source() else {
            panic!("native range")
        };
        assert_eq!(sheet, name);
        assert_eq!(reference.as_str(), "A1:B3");
    }
    for (a, formula, cache) in [
        ("C4", "Source_A_015!$B$2", "10"),
        ("C5", "Source_B_015!$B$2", "5"),
        ("C6", "SUM(C4:C5)", "15"),
        ("C7", "Source_A_015!$B$3", "20"),
        ("C8", "Source_B_015!$B$3", "7"),
        ("C9", "SUM(C7:C8)", "27"),
    ] {
        let c = sheet.cell(a).unwrap();
        let Some(Cell::Formula(f)) = c.stored() else {
            panic!("linked formula")
        };
        assert_eq!(f.text(), formula);
        let Value::Number(n) = f.cached().unwrap().value() else {
            panic!("stored number")
        };
        assert_eq!(n.as_str(), cache);
    }
    for (mode, ns) in [
        (
            data_consolidation::Conformance::Transitional,
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        ),
        (
            data_consolidation::Conformance::Strict,
            "http://purl.oclc.org/ooxml/spreadsheetml/main",
        ),
    ] {
        let xml = data_consolidation::write_worksheet_data_consolidation(&v, mode).unwrap();
        let wrapped = format!("<worksheet xmlns=\"{ns}\">{xml}</worksheet>");
        assert_eq!(
            data_consolidation::parse_worksheet_data_consolidation(wrapped.as_bytes()).unwrap(),
            Some(v.clone())
        );
    }
    let detached = DataConsolidation::new(Function::Average, Some(refs.clone()))
        .with_top_labels(true)
        .with_link(false);
    let xml = data_consolidation::write_worksheet_data_consolidation(
        &detached,
        data_consolidation::Conformance::Transitional,
    )
    .unwrap();
    let wrapped = format!(
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">{xml}</worksheet>"
    );
    assert_eq!(
        data_consolidation::parse_worksheet_data_consolidation(wrapped.as_bytes()).unwrap(),
        Some(detached)
    );
    assert!(members(&p.to_bytes().unwrap()) == members(&bytes));
}

#[test]
fn native_scalar_publication_preserves_watches_ignored_errors_consolidation_and_pictures() {
    let bytes = native();
    let before = members(&bytes);
    let initial = Package::from_bytes(&bytes).unwrap();
    let workbook = initial.workbook();
    let sheet = workbook.sheet("Print_015").unwrap().unwrap();
    assert!(matches!(
        sheet.cell("B2").unwrap().stored(),
        Some(Cell::Value(Value::Number(n))) if n.as_str() == "42"
    ));
    let source = SourceBackedWorkbook::from_reader(Cursor::new(&bytes)).unwrap();
    let mut e = source.edit_cell_values(["Print_015".into()]).unwrap();
    e.set("Print_015", Address::from_a1("B2").unwrap(), 43u32)
        .unwrap();
    let c = e.commit().unwrap();
    let f = tempfile::NamedTempFile::with_suffix(".xlsx").unwrap();
    c.package().save(f.path()).unwrap();
    let output = std::fs::read(f.path()).unwrap();
    let after = members(&output);
    assert_eq!(
        after.keys().map(String::as_str).collect::<Vec<_>>(),
        before
            .keys()
            .map(String::as_str)
            .filter(|n| *n != "xl/calcChain.xml")
            .collect::<Vec<_>>()
    );
    let p = Package::open(f.path()).unwrap();
    caches(&p);
    let b = p.workbook();
    let s = b.sheet("Print_015").unwrap().unwrap();
    let cell = s.cell("B2").unwrap();
    let Some(Cell::Value(Value::Number(n))) = cell.stored() else {
        panic!("B2")
    };
    assert_eq!(n.as_str(), "43");
    for (name, data) in &before {
        if ![
            "[Content_Types].xml",
            "xl/_rels/workbook.xml.rels",
            "xl/workbook.xml",
            "xl/calcChain.xml",
            "xl/worksheets/sheet5.xml",
        ]
        .contains(&name.as_str())
        {
            assert!(after[name] == *data, "changed {name}");
        }
    }
    assert_eq!(
        fragment(&before["xl/worksheets/sheet5.xml"], b"headerFooter"),
        fragment(&after["xl/worksheets/sheet5.xml"], b"headerFooter")
    );
    let restored = SourceBackedWorkbook::from_reader(Cursor::new(&output))
        .unwrap()
        .apply_cell_values(&c.patch().inverse())
        .unwrap();
    let mut restored_bytes = Vec::new();
    restored.package().write_to(&mut restored_bytes).unwrap();
    assert!(members(&restored_bytes) == before);
    assert_eq!(native(), bytes);
}
