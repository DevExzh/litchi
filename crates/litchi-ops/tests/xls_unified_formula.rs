//! Unified XLS formula writes and formula/cached-value reads through the
//! public [`File`] API, on both the retained positional owner and the eager
//! byte owner.
#![cfg(feature = "xls")]
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::panic,
    reason = "integration fixtures fail fast"
)]

use std::ops::Deref;

use litchi_core::Position;
use litchi_ops::{Code, File, OpenOptions, SourceMode, props};
use litchi_xls::cell_values::{FormulaCache, Reference, Selector, Storage, Value};

fn fixture() -> Vec<u8> {
    let mut builder = litchi_xls::Builder::new();
    let data = builder.add_worksheet("Data").unwrap();
    builder.write_number(data, 0, 0, 3.0).unwrap();
    builder.write_number(data, 0, 1, 4.0).unwrap();
    builder.write_string(data, 1, 0, "label").unwrap();
    builder.write_formula(data, 2, 0, "A1*2").unwrap();
    let package = builder.build().unwrap();
    // Give the authored formula a cached result, as a calculating
    // application would.
    let mut edit = package.edit().unwrap();
    edit.set_value(
        "Data",
        Reference::new(2, 0).unwrap(),
        Value::FormulaCache(FormulaCache::Number(6.0)),
    )
    .unwrap();
    edit.commit().unwrap().into_package().to_bytes().unwrap()
}

struct Owner {
    file: File,
    _guard: Option<tempfile::TempDir>,
}

impl Deref for Owner {
    type Target = File;

    fn deref(&self) -> &File {
        &self.file
    }
}

fn open_owner(bytes: &[u8], eager: bool) -> Owner {
    if eager {
        return Owner {
            file: litchi_ops::open_bytes_with(
                bytes.to_vec(),
                &OpenOptions::new().source_mode(SourceMode::Eager),
            )
            .unwrap(),
            _guard: None,
        };
    }
    let guard = tempfile::tempdir().unwrap();
    let path = guard.path().join("source.xls");
    std::fs::write(&path, bytes).unwrap();
    Owner {
        file: litchi_ops::open_with(&path, &OpenOptions::new().source_mode(SourceMode::Auto))
            .unwrap(),
        _guard: Some(guard),
    }
}

fn property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .format()
        .unwrap()
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .map(std::borrow::Cow::into_owned)
        .unwrap_or_default()
}

fn native_cell(bytes: &[u8], row: u32, column: u32) -> Option<(Storage, u16)> {
    let snapshot = litchi_xls::Package::from_bytes(bytes)
        .and_then(|package| package.cell_values())
        .unwrap();
    snapshot
        .worksheet(Selector::Position(Position::new(0)))
        .unwrap()
        .unwrap()
        .cell(Reference::new(row, column).unwrap())
        .unwrap()
        .map(|cell| (cell.storage(), cell.style().get()))
}

#[test]
fn formula_and_cached_value_read_on_both_routes() {
    let bytes = fixture();
    for eager in [false, true] {
        let owner = open_owner(&bytes, eager);
        assert_eq!(
            property(&owner, "/Data/A3", "formula").as_deref(),
            Some("(A1*2)")
        );
        assert_eq!(
            property(&owner, "/Data/A3", "cachedValue").as_deref(),
            Some("6")
        );
        assert_eq!(text(&owner, "/Data/A3"), "6");
        assert_eq!(property(&owner, "/Data/A1", "formula"), None);
        assert_eq!(property(&owner, "/Data/A1", "cachedValue"), None);
        // Query projections carry the same properties.
        let query = litchi_core::selector::Query::parse("cell[formula]")
            .unwrap()
            .with_scope(litchi_core::selector::Scope::Path(Box::new(
                litchi_core::selector::Path::parse("/Data").unwrap(),
            )));
        let matches: Vec<_> = owner
            .query(query)
            .unwrap()
            .map(|node| node.unwrap().path().to_string())
            .collect();
        assert_eq!(matches, ["/sheet[1]/A3"], "eager={eager}");
    }
}

#[test]
fn formulas_needing_workbook_tables_render_on_the_positional_route() {
    let bytes = std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/xls/FormulaSheetRange.xls"),
    )
    .unwrap();
    let eager = open_owner(&bytes, true);
    let source = open_owner(&bytes, false);
    let expected = property(&eager, "/test/D11", "formula").expect("eager 3-D formula");
    assert!(expected.contains("Sheet2:Sheet5"), "{expected}");
    assert_eq!(property(&source, "/test/D11", "formula"), Some(expected));
    assert_eq!(
        property(&source, "/test/D11", "cachedValue"),
        property(&eager, "/test/D11", "cachedValue")
    );
}

#[test]
fn formulas_are_written_into_absent_and_scalar_cells() {
    let bytes = fixture();
    for eager in [false, true] {
        let owner = open_owner(&bytes, eager);
        let style = native_cell(&bytes, 0, 1).unwrap().1;
        let written = owner
            .set("/Data/B2", props! { "formula" => "A1+B1" })
            .unwrap()
            .set("/Data/B1", props! { "formula" => "SUM(A1:A3)" })
            .unwrap();
        assert_eq!(
            property(&written, "/Data/B2", "formula").as_deref(),
            Some("(A1+B1)")
        );
        assert_eq!(
            property(&written, "/Data/B1", "formula").as_deref(),
            Some("SUM((A1:A3))")
        );
        // An authored formula has no cached result until recalculation.
        assert_eq!(property(&written, "/Data/B1", "cachedValue"), None);
        let published = written.to_bytes().unwrap();
        assert_eq!(
            native_cell(&published, 0, 1),
            Some((Storage::Formula, style))
        );
        let reopened = litchi_ops::open_bytes(published).unwrap();
        assert_eq!(
            property(&reopened, "/Data/B2", "formula").as_deref(),
            Some("(A1+B1)")
        );
        assert_eq!(text(&reopened, "/Data/A1"), "3");
        assert_eq!(owner.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn ranges_broadcast_one_formula_in_one_native_commit() {
    let owner = open_owner(&fixture(), true);
    let commit = owner
        .edit(|tx| tx.set("/Data/C1:D2", props! { "formula" => "1+1" }))
        .unwrap();
    assert_eq!(commit.diagnostics().operations(), 1);
    for cell in ["C1", "C2", "D1", "D2"] {
        assert_eq!(
            property(commit.file(), &format!("/Data/{cell}"), "formula").as_deref(),
            Some("(1+1)"),
            "{cell}"
        );
    }
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.file())
            .unwrap()
            .to_bytes()
            .unwrap(),
        fixture()
    );
}

#[test]
fn invalid_and_unsupported_formula_writes_are_typed_and_atomic() {
    let bytes = fixture();
    let owner = open_owner(&bytes, true);
    let cases: [(&str, Code); 4] = [
        ("=A1", Code::InvalidValue),
        ("", Code::InvalidValue),
        ("Other!A1", Code::InvalidValue),
        ("SUM(", Code::InvalidValue),
    ];
    for (expression, code) in cases {
        let error = owner
            .set("/Data/B3", props! { "formula" => expression })
            .expect_err(expression);
        assert_eq!(error.code(), code, "{expression:?}: {error}");
    }
    let error = owner
        .set(
            "/Data/B3",
            litchi_ops::Props::from_props(props! { "formula" => "A1", "value" => 1 }).unwrap(),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidArgument);
    // An existing formula record is not replaced.
    let error = owner
        .set("/Data/A3", props! { "formula" => "A1*3" })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert_eq!(error.property(), Some("formula"));
    assert_eq!(owner.to_bytes().unwrap(), bytes);
}

#[test]
fn formula_writes_are_refused_on_signed_workbooks() {
    let mut ole = litchi_cfb::OleFile::open(
        std::io::Cursor::new(fixture()),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    let mut writer = litchi_cfb::OleWriter::new();
    for path in ole.list_streams() {
        let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
        let data = ole.open_stream(&refs).unwrap();
        writer.create_stream(&refs, &data).unwrap();
    }
    writer
        .create_stream(&["_signatures"], b"synthetic")
        .unwrap();
    let mut signed = std::io::Cursor::new(Vec::new());
    writer.write_to(&mut signed).unwrap();
    let signed = litchi_ops::open_bytes(signed.into_inner()).unwrap();
    let error = signed
        .set("/Data/B3", props! { "formula" => "A1" })
        .unwrap_err();
    assert_eq!(error.code(), Code::DocumentProtected);
}
