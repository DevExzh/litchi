//! Unified XLSB read coverage (M1 R4d).
//!
//! These tests deliberately enter through [`crate::File`].  The leaf tests in
//! `tests.rs` already cover the `Xlsb` handle and its worksheet trait; this
//! module pins the contract at the family boundary: canonical positional
//! paths, lazy source reads, owner parity, bounded node views and the typed
//! refusals for tabs the leaf exposes but does not model as worksheets.

use std::io::Cursor;
use std::num::{NonZeroU64, NonZeroUsize};
use std::path::{Path, PathBuf};
use std::sync::Arc;

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as BudgetLimits,
    OwnedSource, ReadAt,
};
use litchi_opc::constants::relationship_type;
use litchi_opc::{OpcPackage, PackURI, SourceBackedPackage, SourceCacheLimits, SourceLimits};
use litchi_xlsb::builder::chart::{Anchor, Chart};
use litchi_xlsb::builder::{self, Builder};

use crate::File;
use crate::error::{Code, Result};
use crate::node::{Node, SerializeError, ViewOptions};
use crate::options::{OpenOptions, SourceMode};
use crate::view::Mode;
use litchi_core::sheet::Worksheet as _;

fn fixture_path(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ooxml")
        .join(name)
}

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(fixture_path(name)).expect("read XLSB fixture")
}

fn open(bytes: Vec<u8>, mode: SourceMode) -> File {
    let options = OpenOptions::new().source_mode(mode);
    crate::open_bytes_with(bytes, &options).expect("open XLSB fixture")
}

fn open_with_execution_context(bytes: Vec<u8>, context: ExecutionContext) -> File {
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::from_arc(Arc::new(bytes)));
    let package = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        litchi_opc::SourceReadPolicy::exact(),
        Some(context),
    )
    .expect("indexed OPC package with execution context");
    let xlsb = super::Xlsb::from_source_package(package, &OpenOptions::new())
        .expect("source-backed XLSB with execution context");
    File::Xlsb(xlsb)
}

fn both(bytes: &[u8]) -> [File; 2] {
    [
        open(bytes.to_vec(), SourceMode::Auto),
        open(bytes.to_vec(), SourceMode::Eager),
    ]
}

fn query_paths(file: &File, selector: &str) -> Result<Vec<String>> {
    file.query(selector)?
        .map(|node| node.map(|node| node.path().to_string()))
        .collect()
}

fn child_paths(node: &Node<'_>) -> Result<Vec<String>> {
    node.children()
        .map(|child| child.map(|child| child.path().to_string()))
        .collect()
}

fn node_json(node: &Node<'_>, options: ViewOptions) -> String {
    let mut bytes = Vec::new();
    node.view_with(options)
        .serialize_into(&mut serde_json::Serializer::new(&mut bytes))
        .expect("node view serializes");
    String::from_utf8(bytes).expect("node view is UTF-8")
}

fn save(mut writer: Builder) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    writer
        .build()
        .expect("build XLSB fixture")
        .write_to(&mut output)
        .expect("write XLSB fixture");
    output.into_inner()
}

fn sparse_workbook() -> Vec<u8> {
    let mut sheet = builder::Worksheet::new("Sparse");
    sheet.set_cell(0, 0, "A1");
    sheet.set_cell(2, 2, "C3");
    sheet.set_cell(4, 4, 5.0);

    let mut tail = builder::Worksheet::new("Tail");
    tail.set_cell(0, 0, "tail");

    let mut writer = Builder::new();
    writer.add_worksheet(sheet);
    writer.add_worksheet(tail);
    save(writer)
}

fn unicode_workbook() -> Vec<u8> {
    let mut decomposed = builder::Worksheet::new("Cafe\u{301}");
    decomposed.set_cell(0, 0, "accent");
    let mut other = builder::Worksheet::new("Other");
    other.set_cell(0, 0, "other");

    let mut writer = Builder::new();
    writer.add_worksheet(decomposed);
    writer.add_worksheet(other);
    save(writer)
}

fn unicode_ambiguous_workbook() -> Vec<u8> {
    // Excel's leaf name check lower-cases these two differently, while the
    // unified canonical key applies full Unicode case folding (`ß` -> `ss`).
    let mut sharp_s = builder::Worksheet::new("Straße");
    sharp_s.set_cell(0, 0, "sharp s");
    let mut ascii = builder::Worksheet::new("STRASSE");
    ascii.set_cell(0, 0, "ascii");

    let mut writer = Builder::new();
    writer.add_worksheet(sharp_s);
    writer.add_worksheet(ascii);
    save(writer)
}

fn chart_workbook() -> Vec<u8> {
    let chart = Chart::bar_chart_with_cache(
        "Sales",
        "Data!$A$1:$A$2",
        &["North", "South"],
        "Data!$B$1:$B$2",
        &[42.0, 55.0],
        Anchor::new(0, 0, 10, 20),
    )
    .expect("chart fixture");

    let mut data = builder::Worksheet::new("Data");
    data.set_cell(0, 0, "DATA");
    let mut tail = builder::Worksheet::new("Tail");
    tail.set_cell(0, 0, "TAIL");

    let mut writer = Builder::new();
    writer.add_worksheet(data);
    writer
        .add_chart_sheet(builder::ChartSheet::new("Sales Chart", chart))
        .expect("chart tab");
    writer.add_worksheet(tail);
    save(writer)
}

fn malformed_second_sheet() -> Vec<u8> {
    let mut package = OpcPackage::from_reader(
        Cursor::new(fixture("xlsb/Simple.xlsb")),
        SourceLimits::default(),
    )
    .expect("read OPC package");
    package
        .get_part_mut(&PackURI::new("/xl/worksheets/sheet2.bin").expect("sheet URI"))
        .expect("second worksheet")
        .set_blob(vec![0xff; 8]);
    let mut output = Vec::new();
    package
        .write_to(&mut output)
        .expect("write malformed package");
    output
}

/// Add the same deferred pivot edge used by the leaf source-backed tests.
/// The indexed fixture remains otherwise unchanged, so the worksheet's
/// binary-index projection can be selected after full materialization refuses
/// the deferred pivot formula context.
fn indexed_pivot_workbook() -> Vec<u8> {
    let mut package = OpcPackage::from_reader(
        Cursor::new(fixture("xlsb/sample.xlsb")),
        SourceLimits::default(),
    )
    .expect("read indexed OPC package");
    package
        .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.bin").expect("sheet URI"))
        .expect("indexed first worksheet")
        .rels_mut()
        .add_relationship(
            relationship_type::PIVOT_TABLE.to_owned(),
            "../pivotTables/pivotTable1.bin".to_owned(),
            "rIdPivot".to_owned(),
            true,
        );
    let mut output = Vec::new();
    package
        .write_to(&mut output)
        .expect("write indexed pivot package");
    output
}

fn corrupt_index_workbook() -> Vec<u8> {
    let mut package = OpcPackage::from_reader(
        Cursor::new(indexed_pivot_workbook()),
        SourceLimits::default(),
    )
    .expect("read indexed pivot OPC package");
    package
        .get_part_mut(&PackURI::new("/xl/worksheets/binaryIndex1.bin").expect("index URI"))
        .expect("first worksheet index")
        .set_blob(vec![0xff; 8]);
    let mut output = Vec::new();
    package
        .write_to(&mut output)
        .expect("write corrupt indexed package");
    output
}

fn managed_context() -> (CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "xlsb-unified-read-test",
        BudgetLimits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (source, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("nonzero operation limit"),
        NonZeroUsize::new(1).expect("nonzero depth limit"),
        NonZeroU64::new(u64::MAX).expect("nonzero byte limit"),
        0,
    )
    .expect("execution limits");
    (source, ExecutionContext::new(budget, token, limits))
}

fn assert_code(error: &crate::Error, code: Code) {
    assert_eq!(error.code(), code, "{error}");
}

#[test]
fn file_navigation_query_and_views_match_source_and_eager() {
    let files = both(&fixture("xlsb/Simple.xlsb"));
    let mut owner_outputs = Vec::new();

    for file in &files {
        let root = file.get("/").expect("workbook root");
        assert_eq!(root.kind(), "workbook");
        assert_eq!(root.path().to_string(), "/");

        let sheets = child_paths(&root).expect("sheet children");
        assert_eq!(sheets, ["/sheet[1]", "/sheet[2]", "/sheet[3]"]);

        for (input, canonical, kind) in [
            ("/Sheet1/A1", "/sheet[1]/A1", "cell"),
            ("/sheet[2]/B2", "/sheet[2]/B2", "cell"),
            ("Sheet3!C3", "/sheet[3]/C3", "cell"),
        ] {
            let node = file
                .get(input)
                .unwrap_or_else(|error| panic!("{input}: {error}"));
            assert_eq!(node.path().to_string(), canonical, "{input}");
            assert_eq!(node.kind(), kind, "{input}");
        }

        assert!(file.find("/missing").expect("find missing").is_none());
        assert_eq!(
            query_paths(file, "cell").expect("cell query"),
            query_paths(&files[0], "cell").unwrap_or_default()
        );

        let text = file.view(Mode::Text).expect("text view");
        assert!(!text.is_empty(), "text view must render worksheet content");
        let leaf_text = file.text().expect("leaf text");
        let rendered = [Mode::Annotated, Mode::Outline, Mode::Stats]
            .into_iter()
            .map(|mode| {
                file.view(mode)
                    .unwrap_or_else(|error| panic!("{mode:?}: {error}"))
            })
            .collect::<Vec<_>>();
        owner_outputs.push((leaf_text, text, rendered));
    }

    assert_eq!(owner_outputs[0], owner_outputs[1]);
}

#[test]
fn sheet_names_use_unicode_casefold_and_report_ambiguity() {
    let files = both(&unicode_workbook());
    for file in &files {
        let node = file.get("/CAFÉ/A1").expect("NFD sheet name folds to NFC");
        assert_eq!(node.path().to_string(), "/sheet[1]/A1");
    }

    let files = both(&unicode_ambiguous_workbook());
    for file in &files {
        assert!(file.get("/sheet[1]/A1").is_ok());
        assert!(file.get("/sheet[2]/A1").is_ok());
        let error = file
            .get("/strasse")
            .expect_err("casefold collision is ambiguous");
        assert_code(&error, Code::InvalidPath);
        let error = file
            .get("/STRASSE/A1")
            .expect_err("casefold collision is ambiguous for cells too");
        assert_code(&error, Code::InvalidPath);
    }
}

#[test]
fn physical_tab_ordinals_include_chart_tabs_and_grid_access_is_refused() {
    let files = both(&chart_workbook());
    for (owner, file) in files.iter().enumerate() {
        let root = file.get("/").expect("workbook root");
        assert_eq!(
            child_paths(&root).expect("physical tabs"),
            ["/sheet[1]", "/sheet[2]", "/sheet[3]"]
        );
        let first = file
            .get("/sheet[1]/A1")
            .unwrap_or_else(|error| panic!("owner {owner} first worksheet: {error}"));
        assert_eq!(first.kind(), "cell");
        let tail = file
            .get("/sheet[3]/A1")
            .unwrap_or_else(|error| panic!("owner {owner} tail worksheet: {error}"));
        assert_eq!(tail.kind(), "cell");

        let chart = file.get("/sheet[2]").expect("chart tab node");
        assert!(child_paths(&chart).expect("chart children").is_empty());
        for path in ["/sheet[2]/A1", "/Sales Chart/A1"] {
            let error = file.get(path).expect_err(path);
            assert_code(&error, Code::UnsupportedType);
        }
    }
}

#[test]
fn typed_node_coordinates_are_zero_based_against_one_based_paths() {
    use litchi_core::sheet::WorkbookTrait;

    let files = both(&fixture("xlsb/Simple.xlsb"));
    for (owner, file) in files.iter().enumerate() {
        let workbook = file.workbook().expect("workbook view");
        let worksheet = WorkbookTrait::worksheet_by_index(&workbook, 0).expect("first sheet");

        // The generic node contract keeps zero-based coordinates in Rust,
        // while the path grammar writes A1 references with one-based rows and
        // columns. The compatibility worksheet trait remains useful here to
        // assert that a zero coordinate is rejected before leaf materialization.
        let a1 = file.get("/sheet[1]/A1").expect("A1");
        let b2 = file.get("/sheet[1]/B2").expect("B2");
        let typed_b2 = match &b2 {
            Node::Xlsb(node) => node.cell().expect("typed B2 coordinate"),
            _ => panic!("XLSB path returned a different node variant"),
        };
        assert_eq!(typed_b2, (1, 1));
        // The compatibility trait still exposes the legacy source/eager
        // coordinate split; the unified node above has one stable basis.
        assert_eq!(
            worksheet.cell_value(0, 0).is_err(),
            owner == 0,
            "legacy worksheet coordinate policy differs by owner"
        );
        assert_eq!(a1.path().to_string(), "/sheet[1]/A1");
        assert_eq!(b2.path().to_string(), "/sheet[1]/B2");
    }
}

#[test]
fn missing_stored_sparse_ranges_and_dense_views_obey_the_budget() {
    let bytes = sparse_workbook();
    let files = both(&bytes);
    for file in &files {
        let missing = file.get("/Sparse/B2").expect("missing cells have paths");
        assert_eq!(missing.kind(), "cell");
        assert_eq!(missing.path().to_string(), "/sheet[1]/B2");

        let range = file.get("/Sparse/A1:E5").expect("sparse range");
        assert_eq!(range.kind(), "range");
        let stored = child_paths(&range).expect("stored range children");
        assert_eq!(stored.len(), 3);
        assert!(stored.iter().all(|path| !path.ends_with("B2")));

        let sparse_json = node_json(&range, ViewOptions::default());
        let dense_json = node_json(&range, ViewOptions::default().dense(25));
        let sparse_json: serde_json::Value =
            serde_json::from_str(&sparse_json).expect("sparse JSON");
        let dense_json: serde_json::Value = serde_json::from_str(&dense_json).expect("dense JSON");
        assert_eq!(sparse_json["children"].as_array().unwrap().len(), 3);
        assert_eq!(dense_json["children"].as_array().unwrap().len(), 25);

        let mut sink = Vec::new();
        let error = range
            .view_with(ViewOptions::default().dense(4))
            .serialize_into(&mut serde_json::Serializer::new(&mut sink))
            .expect_err("a 5x5 dense range exceeds four nodes");
        assert!(matches!(
            error,
            SerializeError::View(error) if error.code() == Code::OutputTooLarge
        ));
    }
}

#[test]
fn malformed_deferred_sheet_is_not_read_by_root_or_limited_walk() {
    let file = open(malformed_second_sheet(), SourceMode::Auto);
    let root = file.get("/").expect("catalog-only root");
    assert_eq!(root.kind(), "workbook");

    let mut sink = Vec::new();
    root.view_with(ViewOptions::default().depth(0))
        .serialize_into(&mut serde_json::Serializer::new(&mut sink))
        .expect("depth-zero root does not read worksheets");
    assert_eq!(
        file.get("/sheet[1]/A1")
            .expect("first worksheet is still readable")
            .kind(),
        "cell"
    );

    let first = file
        .query("sheet")
        .expect("compile sheet query")
        .limit(1)
        .next()
        .expect("one limited match")
        .expect("first worksheet");
    assert_eq!(first.path().to_string(), "/sheet[1]");

    assert!(
        query_paths(&file, "table")
            .expect("table query does not read deferred worksheet grids")
            .is_empty(),
        "an unselected malformed worksheet must not be parsed by a table query"
    );

    let error = file
        .get("/sheet[2]/A1")
        .expect_err("selected malformed worksheet is refused");
    assert_code(&error, Code::CorruptFile);
}

#[cfg(any(unix, windows))]
#[test]
fn source_change_fences_acquired_nodes_queries_iterators_and_views() {
    let temporary = tempfile::NamedTempFile::new().expect("temporary workbook");
    std::fs::write(temporary.path(), fixture("xlsb/Simple.xlsb")).expect("write source");
    let file = crate::open(temporary.path()).expect("source-backed open");

    let node = file.get("/sheet[1]/A1").expect("acquire node");
    let view = node.view_with(ViewOptions::default());
    let mut query = file.query("cell").expect("acquire query");
    let root = file.get("/").expect("acquire root");
    let mut children = root.children();

    std::fs::write(temporary.path(), sparse_workbook()).expect("change source");

    assert_code(&view.text().expect_err("stale node view"), Code::IoError);
    let mut sink = Vec::new();
    let error = view
        .serialize_into(&mut serde_json::Serializer::new(&mut sink))
        .expect_err("stale node serialization");
    assert!(matches!(
        error,
        SerializeError::View(error) if error.code() == Code::IoError
    ));

    let error = query
        .next()
        .expect("query reports one stale item")
        .expect_err("stale query");
    assert_code(&error, Code::IoError);
    let error = children
        .next()
        .expect("iterator reports one stale item")
        .expect_err("stale child iterator");
    assert_code(&error, Code::IoError);
    assert_code(
        &file.view(Mode::Text).expect_err("stale workbook view"),
        Code::IoError,
    );
}

#[test]
fn indexed_cached_fallback_projects_scalars_without_formula_or_style_metadata() {
    let bytes = indexed_pivot_workbook();
    let source = open(bytes.clone(), SourceMode::Auto);

    // B10 is a formula cell in the indexed fixture.  The cached-only route
    // exposes its stored scalar, while retaining no formula or XF metadata.
    let cached_formula = source
        .get("/sheet[1]/B10")
        .expect("indexed cached formula scalar");
    assert_eq!(cached_formula.kind(), "cell");
    let text = cached_formula
        .view()
        .text()
        .expect("cached formula text")
        .expect("cached formula has a scalar")
        .into_owned();
    assert!(
        !text.is_empty(),
        "cached formula scalar is not fabricated empty text"
    );

    let properties = cached_formula
        .view()
        .format()
        .expect("cached formula properties")
        .map(|property| property.key().to_owned())
        .collect::<Vec<_>>();
    assert!(properties.iter().any(|key| key == "type"));
    assert!(!properties.iter().any(|key| key == "formula"));
    assert!(!properties.iter().any(|key| key == "cachedValue"));
    assert!(!properties.iter().any(|key| {
        key.starts_with("font.")
            || key.starts_with("border.")
            || key.starts_with("alignment.")
            || key == "fill"
            || key == "fillPattern"
            || key == "fillBg"
            || key == "numberformat"
    }));

    let cells = query_paths(&source, "cell").expect("cached scalar query");
    assert!(!cells.is_empty(), "indexed cached cells are queryable");
    assert!(cells.iter().any(|path| path == "/sheet[1]/B10"));

    let text_view = source.view(Mode::Text).expect("cached scalar text view");
    assert!(text_view.contains("[/sheet[1]/row[10]]"));
    let stats = serde_json::to_value(
        source
            .view_json(Mode::Stats)
            .expect("cached scalar stats view"),
    )
    .expect("serialize cached scalar stats");
    assert!(
        stats.get("formulaCells").is_none(),
        "formulaCells is unavailable in a cached-only projection, not zero"
    );

    // The synthetic pivot edge is deliberately unresolved.  Eager ownership
    // validates that closure while opening and refuses the package before a
    // node can be acquired; the source owner remains lazy and can use its
    // indexed scalar projection.
    let error = crate::open_bytes_with(bytes, &OpenOptions::new().source_mode(SourceMode::Eager))
        .expect_err("eager full grid has no indexed fallback");
    assert_code(&error, Code::CorruptFile);
}

#[test]
fn indexed_cached_fallback_lists_rows_without_row_metadata() {
    let file = open(indexed_pivot_workbook(), SourceMode::Auto);
    let sheet = file.get("/sheet[1]").expect("indexed worksheet");
    let rows = child_paths(&sheet).expect("cached-only row children");
    let expected = (1..=10)
        .map(|row| format!("/sheet[1]/row[{row}]"))
        .collect::<Vec<_>>();
    assert_eq!(rows, expected);
}

#[test]
fn indexed_cached_fallback_preserves_corruption_and_cancellation_errors() {
    let corrupt = open(corrupt_index_workbook(), SourceMode::Auto);
    let error = corrupt
        .get("/sheet[1]/B10")
        .expect_err("corrupt binary index is not an empty cached projection");
    assert_code(&error, Code::CorruptFile);

    let (cancellation, context) = managed_context();
    let file = open_with_execution_context(indexed_pivot_workbook(), context);
    cancellation.cancel();
    let error = file
        .get("/sheet[1]/B10")
        .expect_err("cancellation is not swallowed by cached fallback");
    assert_code(&error, Code::Cancelled);
}

#[cfg(any(unix, windows))]
#[test]
fn indexed_cached_fallback_preserves_stale_source_errors() {
    let temporary = tempfile::NamedTempFile::new().expect("temporary workbook");
    std::fs::write(temporary.path(), indexed_pivot_workbook()).expect("write pivot source");
    let file = crate::open(temporary.path()).expect("source-backed indexed workbook");
    assert_eq!(
        file.get("/sheet[1]/B10")
            .expect("initial cached fallback")
            .kind(),
        "cell"
    );

    std::fs::write(temporary.path(), fixture("xlsb/Simple.xlsb")).expect("change source");
    let error = file
        .get("/sheet[1]/B10")
        .expect_err("stale source is not hidden by cached fallback");
    assert_code(&error, Code::IoError);
}

#[test]
fn cached_pivot_fixture_keeps_leaf_only_data_as_a_typed_refusal() {
    let bytes = fixture("xlsb/pivot-cache-ac-prefixed.xlsb");
    for file in &both(&bytes) {
        let error = file
            .get("/sheet[1]/pivottable[1]")
            .expect_err("pivot cache is not a readable pivot node");
        assert_code(&error, Code::UnsupportedCapability);
    }
}
