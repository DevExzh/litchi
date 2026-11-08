//! External contract tests for the unified legacy XLS read adapter.
//!
//! The fixture is authored through the native XLS builder so these tests do
//! not duplicate BIFF records.  Each read is then exercised through the
//! public `File` surface in both of the owner states users can obtain: an
//! eager byte input and a source-backed path input.
#![cfg(feature = "xls")]

use std::io::Cursor;

use litchi_ops::node::{Node, SerializeError, ViewOptions};
use litchi_ops::view::Mode;
use litchi_ops::{Code, File, OpenOptions, SourceMode};
use litchi_xls::writer::Builder;

fn workbook_bytes() -> Vec<u8> {
    let mut builder = Builder::new();
    let first = builder
        .add_worksheet("Cafe\u{301}")
        .expect("first worksheet");
    let second = builder.add_worksheet("Tail").expect("second worksheet");

    builder
        .write_string(first, 0, 0, "header")
        .expect("string cell");
    builder
        .write_number(first, 2, 2, 42.5)
        .expect("numeric cell");
    builder
        .write_formula(first, 4, 4, "C3*2")
        .expect("formula cell");
    builder
        .write_string(second, 0, 1, "tail")
        .expect("second-sheet cell");

    let package = builder.build().expect("build XLS workbook");
    let mut output = Cursor::new(Vec::new());
    package.write_to(&mut output).expect("write XLS workbook");
    output.into_inner()
}

fn signed_workbook_bytes() -> Vec<u8> {
    let mut ole = litchi_cfb::OleFile::open(
        Cursor::new(workbook_bytes()),
        litchi_cfb::OleFileLimits::default(),
    )
    .expect("open native XLS package");
    let workbook = ole
        .open_stream(&["Workbook"])
        .expect("open workbook stream");
    let mut writer = litchi_cfb::OleWriter::new();
    writer
        .create_stream(&["Workbook"], &workbook)
        .expect("copy workbook stream");
    writer
        .create_stream(&["_signatures"], b"opaque-signature")
        .expect("add signature stream");
    let mut output = Cursor::new(Vec::new());
    writer
        .write_to(&mut output)
        .expect("write signed XLS package");
    output.into_inner()
}

fn eager(bytes: &[u8]) -> File {
    litchi_ops::open_bytes_with(
        bytes.to_vec(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )
    .expect("open eager XLS")
}

fn source_file(bytes: &[u8]) -> (tempfile::NamedTempFile, File) {
    let temporary = tempfile::NamedTempFile::new().expect("temporary XLS");
    std::fs::write(temporary.path(), bytes).expect("write temporary XLS");
    let file = litchi_ops::open(temporary.path()).expect("open source-backed XLS");
    (temporary, file)
}

fn child_paths(node: &Node<'_>) -> Vec<String> {
    node.children()
        .map(|child| child.expect("child node").path().to_string())
        .collect()
}

fn query_paths(file: &File, selector: &str) -> Vec<String> {
    file.query(selector)
        .unwrap_or_else(|error| panic!("{selector}: {error}"))
        .map(|node| node.expect("query node").path().to_string())
        .collect()
}

fn property_text(node: &Node<'_>, key: &str) -> Option<String> {
    node.view()
        .format()
        .expect("node properties")
        .find(|property| property.key() == key)
        .map(|property| property.to_string())
}

fn assert_code(error: &litchi_ops::Error, code: Code) {
    assert_eq!(error.code(), code, "{error}");
}

#[test]
fn navigation_query_and_views_are_public_and_owner_parity_is_stable() {
    let bytes = workbook_bytes();
    let (temporary, source) = source_file(&bytes);
    let eager = eager(&bytes);
    let _keep_source_alive = temporary;

    for file in [&source, &eager] {
        let root = file.get("/").expect("workbook root");
        assert_eq!(root.kind(), "workbook");
        assert_eq!(root.path().to_string(), "/");
        assert_eq!(
            child_paths(&root),
            ["/sheet[1]", "/sheet[2]"],
            "physical worksheet order is stable"
        );

        let first = file.get("/sheet[1]/A1").expect("first string cell");
        assert_eq!(first.kind(), "cell");
        assert_eq!(first.path().to_string(), "/sheet[1]/A1");
        assert_eq!(
            first.view().text().expect("cell text").as_deref(),
            Some("header")
        );

        let numeric = file.get("/sheet[1]/C3").expect("sparse numeric cell");
        assert_eq!(numeric.path().to_string(), "/sheet[1]/C3");
        assert_eq!(
            numeric.view().text().expect("numeric text").as_deref(),
            Some("42.5")
        );
        assert_eq!(property_text(&numeric, "type").as_deref(), Some("Number"));

        // The native fixture contains a formula cell. It is a real,
        // queryable cell with the same canonical coordinate as scalar cells,
        // and both owner routes render its formula text (record 0259).
        let formula = file.get("/sheet[1]/E5").expect("formula cell");
        assert_eq!(formula.kind(), "cell");
        assert_eq!(formula.path().to_string(), "/sheet[1]/E5");
        // The builder leaves an empty string cache and requests recalculation;
        // writing a formula does not invent a numeric result.
        assert_eq!(
            property_text(&formula, "type").as_deref(),
            Some("SharedString")
        );
        assert_eq!(
            property_text(&formula, "formula").as_deref(),
            Some("(C3*2)")
        );

        assert!(file.find("/missing").expect("find missing sheet").is_none());
        assert_eq!(
            query_paths(file, "cell"),
            [
                "/sheet[1]/A1",
                "/sheet[1]/C3",
                "/sheet[1]/E5",
                "/sheet[2]/B1",
            ]
        );

        let text = file.view(Mode::Text).expect("text view");
        assert!(text.contains("header"));
        assert!(text.contains("42.5"));
        assert!(text.contains("tail"));
        let outline = serde_json::to_value(file.view_json(Mode::Outline).expect("outline view"))
            .expect("serialize outline");
        assert!(outline["sheets"].is_array());
        assert!(outline.to_string().contains("Cafe"));
        let stats = serde_json::to_value(file.view_json(Mode::Stats).expect("stats view"))
            .expect("serialize stats");
        assert_eq!(stats["formulaCells"], 1);
    }

    assert_eq!(
        source.view(Mode::Text).expect("source text"),
        eager.view(Mode::Text).expect("eager text")
    );
    assert_eq!(
        serde_json::to_value(source.view_json(Mode::Stats).expect("source stats"))
            .expect("source stats JSON"),
        serde_json::to_value(eager.view_json(Mode::Stats).expect("eager stats"))
            .expect("eager stats JSON")
    );
}

#[test]
fn signed_xls_remains_readable_through_the_unified_file_surface() {
    let file = litchi_ops::open_bytes(signed_workbook_bytes()).expect("open signed XLS");
    assert!(file.is_signed());
    assert_eq!(
        file.get("/sheet[1]/A1")
            .expect("signed cell")
            .view()
            .text()
            .expect("signed text")
            .as_deref(),
        Some("header")
    );
    assert!(file.find("/sheet[2]/B1").expect("signed find").is_some());
    assert_eq!(
        query_paths(&file, "cell"),
        [
            "/sheet[1]/A1",
            "/sheet[1]/C3",
            "/sheet[1]/E5",
            "/sheet[2]/B1",
        ]
    );
}

#[test]
fn sheet_names_are_unicode_normalized_and_cell_coordinates_are_bounded() {
    let bytes = workbook_bytes();
    let (temporary, source) = source_file(&bytes);
    let eager = eager(&bytes);
    let _keep_source_alive = temporary;

    for file in [&source, &eager] {
        let normalized = file.get("/CAFÉ/A1").expect("NFD sheet name normalizes");
        assert_eq!(normalized.path().to_string(), "/sheet[1]/A1");
        assert_eq!(
            file.get("Cafe\u{301}!A1")
                .expect("A1 sheet notation")
                .path()
                .to_string(),
            "/sheet[1]/A1"
        );

        for path in [
            "/sheet[1]/A0",
            "/sheet[1]/XFE1",
            "/sheet[1]/IW1",
            "/sheet[1]/A65537",
            "/sheet[1]/A1:IW2",
            "/sheet[1]/row[65537]",
            "/sheet[1]/col[257]",
        ] {
            let error = file.get(path).expect_err(path);
            assert_code(&error, Code::InvalidPath);
        }
        assert!(
            file.find("/sheet[1]/Z99")
                .expect("missing sparse cell")
                .is_some()
        );
    }
}

#[test]
fn unsupported_legacy_axes_return_typed_capability_gaps() {
    let file = eager(&workbook_bytes());
    for path in ["/sheet[1]/table[1]", "/sheet[1]/chart[1]"] {
        let error = file.get(path).expect_err(path);
        assert_code(&error, Code::UnsupportedCapability);
    }
    // Notes are readable (record 0327); this fixture has none.
    assert_code(
        &file.get("/sheet[1]/comment[1]").expect_err("no note"),
        Code::PathNotFound,
    );
    // Defined names are readable (record 0259); this fixture has none.
    assert_code(
        &file.get("/namedrange[1]").expect_err("no defined name"),
        Code::PathNotFound,
    );
}

#[test]
fn range_views_are_bounded_before_dense_serialization() {
    let file = eager(&workbook_bytes());
    let range = file.get("/sheet[1]/A1:E5").expect("cell range");
    assert_eq!(range.kind(), "range");
    assert_eq!(range.path().to_string(), "/sheet[1]/A1:E5");
    let children = child_paths(&range);
    assert_eq!(
        children,
        ["/sheet[1]/A1", "/sheet[1]/C3", "/sheet[1]/E5"],
        "the default range walk stays sparse"
    );

    let mut sink = Vec::new();
    let error = range
        .view_with(ViewOptions::default().dense(4))
        .serialize_into(&mut serde_json::Serializer::new(&mut sink))
        .expect_err("dense output exceeds its node budget");
    assert!(matches!(
        error,
        SerializeError::View(error) if error.code() == Code::OutputTooLarge
    ));
}

#[cfg(any(unix, windows))]
#[test]
fn source_nodes_queries_and_views_refuse_a_changed_path() {
    let bytes = workbook_bytes();
    let (temporary, file) = source_file(&bytes);
    let node = file.get("/sheet[1]/A1").expect("source cell");
    let node_view = node.view();
    let mut query = file.query("cell").expect("source query");
    let root = file.get("/").expect("source root");
    let mut children = root.children();

    std::fs::write(temporary.path(), b"changed XLS source").expect("change source");

    assert_code(
        &node_view.text().expect_err("stale node view"),
        Code::IoError,
    );
    let mut sink = Vec::new();
    let error = node_view
        .serialize_into(&mut serde_json::Serializer::new(&mut sink))
        .expect_err("stale node serialization");
    assert!(matches!(
        error,
        SerializeError::View(error) if error.code() == Code::IoError
    ));
    assert_code(
        &query
            .next()
            .expect("stale query reports one item")
            .expect_err("stale query item"),
        Code::IoError,
    );
    assert_code(
        &children
            .next()
            .expect("stale child iterator reports one item")
            .expect_err("stale child item"),
        Code::IoError,
    );
    assert_code(
        &file.view(Mode::Text).expect_err("stale workbook view"),
        Code::IoError,
    );
}

#[cfg(feature = "encryption")]
#[test]
fn encrypted_xls_unlocks_into_the_same_read_projection() {
    let mut builder = Builder::new();
    let sheet = builder.add_worksheet("Secret").expect("worksheet");
    builder
        .write_string(sheet, 0, 0, "protected")
        .expect("protected cell");
    builder
        .set_password("secret", litchi_xls::EncryptionProfile::OfficeBinaryRc4)
        .expect("encryption profile");
    builder.set_encryption_entropy(litchi_xls::EncryptionEntropy::generate().expect("entropy"));
    let encrypted = builder
        .build()
        .expect("encrypted package")
        .to_encrypted_bytes()
        .expect("encrypted bytes");

    let locked = litchi_ops::open_bytes(encrypted).expect("open locked XLS");
    let File::Locked(locked) = locked else {
        panic!("encrypted XLS must stay locked until password consumption");
    };
    let unlocked = locked
        .unlock(litchi_ops::Password::new("secret".to_owned()))
        .expect("unlock XLS");
    assert_eq!(
        unlocked
            .get("/Secret/A1")
            .expect("unlocked cell")
            .view()
            .text()
            .expect("unlocked text")
            .as_deref(),
        Some("protected")
    );
}

#[test]
fn complete_tab_catalog_preserves_chart_positions_and_hidden_states() {
    let fixture = |name: &str| {
        std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/poi/test-data/spreadsheet")
                .join(name),
        )
        .unwrap()
    };
    let bytes = fixture("44010-TwoCharts.xls");
    let (temporary, source) = source_file(&bytes);
    for file in [&source, &eager(&bytes)] {
        assert_eq!(
            child_paths(&file.root().unwrap()),
            ["/sheet[1]", "/sheet[2]", "/sheet[3]"]
        );
        let chart = file.get("/Graph1").unwrap();
        assert_eq!(chart.path().to_string(), "/sheet[2]");
        assert_code(
            &file.get("/sheet[2]/A1").unwrap_err(),
            Code::UnsupportedCapability,
        );
    }
    drop(temporary);
    let bytes = fixture("45761.xls");
    let (_temporary, source) = source_file(&bytes);
    for file in [&source, &eager(&bytes)] {
        assert_eq!(
            property_text(&file.get("/sheet[2]").unwrap(), "visibility").as_deref(),
            Some("hidden")
        );
        assert_eq!(
            property_text(&file.get("/sheet[3]").unwrap(), "visibility").as_deref(),
            Some("veryHidden")
        );
    }
}

#[test]
fn workbook_date_system_uses_retained_native_metadata() {
    let mut builder = Builder::new();
    builder.set_1904_dates(true);
    let bytes = builder.build().unwrap().to_bytes().unwrap();
    let (_temporary, source) = source_file(&bytes);
    for file in [&source, &eager(&bytes)] {
        assert_eq!(
            property_text(&file.root().unwrap(), "workbook.date1904").as_deref(),
            Some("true")
        );
    }
}

#[test]
fn formula_cached_errors_are_counted_and_annotated_on_both_owners() {
    use litchi_xls::cell_values::{CellError, FormulaCache, Reference, Value};
    let mut builder = Builder::new();
    let sheet = builder.add_worksheet("Data").unwrap();
    builder.write_formula(sheet, 0, 0, "1/0").unwrap();
    let package = builder.build().unwrap();
    let mut edit = package.edit().unwrap();
    edit.set_value(
        "Data",
        Reference::new(0, 0).unwrap(),
        Value::FormulaCache(FormulaCache::Error(CellError::new(0x07).unwrap())),
    )
    .unwrap();
    let (package, _) = edit.commit().unwrap().into_parts();
    let bytes = package.to_bytes().unwrap();
    let (_temporary, source) = source_file(&bytes);
    for file in [&source, &eager(&bytes)] {
        let stats = serde_json::to_value(file.view_json(Mode::Stats).unwrap()).unwrap();
        assert_eq!(stats["formulaCells"], 1);
        assert_eq!(stats["errorCells"], 1);
        let outline = serde_json::to_value(file.view_json(Mode::Outline).unwrap()).unwrap();
        assert_eq!(outline["sheets"][0]["errorCells"], 1);
        assert!(file.view(Mode::Annotated).unwrap().contains("⚠ error"));
    }
}

#[test]
fn defined_names_are_readable_on_both_owner_routes() {
    let bytes = std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/xls/SimpleWithPrintArea.xls"),
    )
    .expect("print-area fixture");
    let (temporary, source) = source_file(&bytes);
    let eager = eager(&bytes);
    let _keep_source_alive = temporary;
    for file in [&source, &eager] {
        let name = file.get("/namedrange[1]").expect("first defined name");
        assert_eq!(name.kind(), "namedrange");
        assert_eq!(name.path().to_string(), "/namedrange[1]");
        assert_eq!(property_text(&name, "name").as_deref(), Some("Print_Area"));
        assert_eq!(
            property_text(&name, "ref").as_deref(),
            Some("'Sheet1'!$A$1:$C$5")
        );
        assert_eq!(property_text(&name, "scope").as_deref(), Some("Sheet1"));
        assert_eq!(
            name.view().text().expect("name text").as_deref(),
            Some("'Sheet1'!$A$1:$C$5")
        );
        assert!(file.get("/namedrange[2]").is_err());
        assert_eq!(
            file.get("/namedrange[Print_Area]")
                .expect("named lookup")
                .path()
                .to_string(),
            "/namedrange[1]"
        );
        assert_eq!(query_paths(file, "namedrange"), ["/namedrange[1]"]);
    }
}
