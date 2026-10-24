#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration regressions use immediate fixture diagnostics"
)]

//! Public integration coverage for the opt-in drawing-skipped Workbook view.
//!
//! The fixture has a valid standard drawing.  The tests make that part
//! malformed after opening through the skipped projection and then exercise
//! publication paths which must keep it opaque.  A second malformed drawing
//! is deliberately left unrelated to any worksheet relationship so package
//! ownership and save/reopen behavior are checked as well.

use litchi_core::sheet::traits::WorkbookTrait;
use litchi_opc::{BlobPart, PackURI};
use litchi_xlsb::Workbook;
use litchi_xlsb::cell_values::structure::Patch;
use litchi_xlsb::cell_values::{Reference, TransferLimits, Value};
use std::collections::BTreeMap;
use std::io::Cursor;
use std::path::PathBuf;

#[path = "support/projection.rs"]
mod projection;
use projection::{InPlace as _, OwnedWorkbook as _};

const DRAWING_CONTENT_TYPE: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
const FIXTURE: &str = "test-data/poi/test-data/spreadsheet/testVarious.xlsb";

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

fn fixture_bytes() -> Vec<u8> {
    std::fs::read(fixture(FIXTURE)).expect("drawing fixture")
}

fn skipped(bytes: &[u8]) -> Workbook {
    Workbook::new_with_options(
        Cursor::new(bytes.to_vec()),
        litchi_xlsb::OpenOptions::new().parse_drawings(false),
    )
    .expect("drawing-skipped workbook")
}

fn saved(workbook: &Workbook) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    workbook
        .to_package()
        .write_to(&mut output)
        .expect("workbook save");
    output.into_inner()
}

fn first_drawing_uri(workbook: &Workbook) -> PackURI {
    workbook
        .opc_package()
        .iter_parts()
        .find(|part| part.content_type() == DRAWING_CONTENT_TYPE)
        .map(|part| part.partname().clone())
        .expect("fixture standard drawing part")
}

fn part_bytes(workbook: &Workbook, uri: &PackURI) -> Vec<u8> {
    workbook
        .opc_package()
        .get_part(uri)
        .expect("package part")
        .blob()
        .to_vec()
}

fn package_blobs(workbook: &Workbook) -> BTreeMap<String, Vec<u8>> {
    workbook
        .opc_package()
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
        .map(|part| (part.partname().as_str().to_owned(), part.blob().to_vec()))
        .collect()
}

fn add_unrelated_malformed_drawing(workbook: &mut Workbook) -> PackURI {
    let uri = PackURI::new("/xl/drawings/unrelated-malformed.xml").expect("orphan URI");
    workbook
        .edit_opc_in_place(|package| {
            package.add_part(Box::new(BlobPart::new(
                uri.clone(),
                DRAWING_CONTENT_TYPE.to_owned(),
                b"<unrelated-broken-drawing/>".to_vec(),
            )));
            Ok(())
        })
        .expect("add unrelated malformed drawing");
    uri
}

fn corrupt_referenced_drawing(workbook: &mut Workbook, uri: &PackURI) {
    workbook
        .edit_opc_in_place(|package| {
            package
                .get_part_mut(uri)?
                .set_blob(b"<broken-drawing/>".to_vec());
            Ok(())
        })
        .expect("corrupt referenced drawing through skipped projection");
}

fn malformed_skipped_fixture() -> (Workbook, PackURI, PackURI) {
    let mut workbook = skipped(&fixture_bytes());
    let referenced_uri = first_drawing_uri(&workbook);
    let unrelated_uri = add_unrelated_malformed_drawing(&mut workbook);
    corrupt_referenced_drawing(&mut workbook, &referenced_uri);
    assert!(workbook.sheet_drawings().is_empty());
    (workbook, referenced_uri, unrelated_uri)
}

fn first_finite_number(workbook: &Workbook) -> (usize, Reference, f64) {
    (0..workbook.worksheet_count())
        .find_map(|sheet| {
            workbook
                .cell_values(sheet)
                .expect("worksheet cell values")
                .numbers()
                .find(|number| number.value().is_finite())
                .map(|number| (sheet, number.reference(), number.value()))
        })
        .expect("fixture finite numeric cell")
}

#[test]
fn skipped_projection_keeps_opaque_graph_across_noop_reparse_and_save() {
    let bytes = fixture_bytes();
    let eager = Workbook::new(Cursor::new(bytes.clone())).expect("eager fixture");
    assert!(
        !eager.sheet_drawings().is_empty(),
        "the control parse must expose the fixture drawing"
    );

    let (mut workbook, referenced_uri, unrelated_uri) = malformed_skipped_fixture();
    let malformed = b"<broken-drawing/>".to_vec();
    let unrelated = b"<unrelated-broken-drawing/>".to_vec();
    assert_eq!(part_bytes(&workbook, &referenced_uri), malformed);
    assert_eq!(part_bytes(&workbook, &unrelated_uri), unrelated);

    let before_noop = package_blobs(&workbook);
    workbook
        .edit_opc_in_place(|_| Ok(()))
        .expect("exact package no-op publication");
    assert!(workbook.sheet_drawings().is_empty());
    assert_eq!(package_blobs(&workbook), before_noop);

    let saved = saved(&workbook);
    let reopened = skipped(&saved);
    assert!(reopened.sheet_drawings().is_empty());
    assert_eq!(package_blobs(&reopened), before_noop);
    assert!(
        Workbook::new(Cursor::new(saved)).is_err(),
        "the eager projection must reject the malformed referenced drawing"
    );
}

#[test]
fn skipped_projection_cell_noop_and_scalar_change_keep_malformed_drawings_opaque() {
    let (mut workbook, referenced_uri, unrelated_uri) = malformed_skipped_fixture();
    let before = package_blobs(&workbook);

    let package = workbook.to_package();
    let name = workbook.worksheet_names()[0].clone();
    let no_op = package
        .edit_cell_values(name)
        .expect("cell no-op selector")
        .resolve()
        .expect("worksheet resolution")
        .expect("cell no-op edit")
        .commit()
        .expect("cell no-op commit");
    assert!(no_op.patch().is_noop());
    workbook = no_op
        .into_package()
        .owned_workbook()
        .expect("cell no-op publication");
    assert!(workbook.sheet_drawings().is_empty());
    assert_eq!(package_blobs(&workbook), before);

    let (sheet, reference, old_value) = first_finite_number(&workbook);
    let replacement = if old_value.to_bits() == 12_345.25_f64.to_bits() {
        54_321.5
    } else {
        12_345.25
    };
    let package = workbook.to_package();
    let name = workbook.worksheet_names()[sheet].clone();
    let mut edit = package
        .edit_cell_values(name)
        .expect("scalar selector")
        .resolve()
        .expect("worksheet resolution")
        .expect("scalar edit");
    edit.set_number(reference, replacement)
        .expect("replace finite number");
    let commit = edit.commit().expect("scalar commit");
    workbook = commit
        .into_package()
        .owned_workbook()
        .expect("scalar publication");

    assert!(workbook.sheet_drawings().is_empty());
    assert_eq!(part_bytes(&workbook, &referenced_uri), b"<broken-drawing/>");
    assert_eq!(
        part_bytes(&workbook, &unrelated_uri),
        b"<unrelated-broken-drawing/>"
    );
    assert_eq!(
        workbook
            .cell_values(sheet)
            .expect("published cell values")
            .number(reference)
            .expect("number lookup")
            .expect("published number")
            .value()
            .to_bits(),
        replacement.to_bits()
    );
}

#[test]
fn skipped_malformed_donor_transfers_to_eager_target_without_drawing_parse() {
    let original = fixture_bytes();
    let (donor, donor_drawing_uri, _) = malformed_skipped_fixture();
    let donor_before = package_blobs(&donor);
    let (source_sheet, source_reference, source_value) = first_finite_number(&donor);
    let destination = Reference::new(10_000, 100).expect("transfer destination");

    let target = Workbook::new(Cursor::new(original.clone())).expect("eager target");
    assert!(
        !target.sheet_drawings().is_empty(),
        "the target starts with its eager drawing projection"
    );
    assert!(target.drawings_parsed());
    let target_before = package_blobs(&target);
    let target_drawing_uri = first_drawing_uri(&target);
    let target_drawing_before = part_bytes(&target, &target_drawing_uri);

    let mut edit = target.to_package().edit().expect("target structure edit");
    edit.transfer_cell(&donor, source_sheet, source_reference, 0, destination)
        .expect("transfer scalar from skipped malformed donor");
    let commit = edit.commit().expect("transfer commit");
    assert_eq!(commit.patch().operation_count(), 1);
    assert_eq!(package_blobs(&donor), donor_before);
    assert_eq!(part_bytes(&donor, &donor_drawing_uri), b"<broken-drawing/>");

    let durable = commit
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .expect("durable transfer patch");
    let decoded = Patch::from_bytes(
        &durable,
        TransferLimits::DEFAULT,
        litchi_xlsb::raw::Limits::DEFAULT,
    )
    .expect("decode durable transfer patch");

    let mut published = Workbook::new(Cursor::new(original.clone())).expect("published target");
    published
        .apply_structure(&decoded)
        .expect("publish durable transfer");
    assert!(
        !published.sheet_drawings().is_empty(),
        "an eager target remains eagerly projected after transfer"
    );
    assert!(published.drawings_parsed());
    assert_eq!(
        part_bytes(&published, &target_drawing_uri),
        target_drawing_before
    );
    let published_cells = published.cell_values(0).expect("published cells");
    let transferred = published_cells
        .cell(destination)
        .expect("destination lookup")
        .expect("transferred cell");
    match transferred.value() {
        Value::Number(value) => assert_eq!(value.to_bits(), source_value.to_bits()),
        value => panic!("expected transferred number, found {value:?}"),
    }

    let inverse = decoded.inverse();
    published
        .apply_structure(&inverse)
        .expect("inverse durable transfer");
    assert!(
        published
            .cell_values(0)
            .expect("inverse cells")
            .cell(destination)
            .expect("inverse destination lookup")
            .is_none(),
        "inverse removes the transferred cell from the exact target image"
    );
    assert!(
        !published.sheet_drawings().is_empty(),
        "inverse keeps the eager drawing projection"
    );
    assert!(published.drawings_parsed());
    assert_eq!(package_blobs(&published), target_before);
}

#[test]
fn workbook_patch_respects_skipped_target_and_durable_decode_boundary() {
    let original = fixture_bytes();
    let source = skipped(&original);
    let drawing_uri = first_drawing_uri(&source);
    let drawing_before = part_bytes(&source, &drawing_uri);
    let old_name = source.worksheet_names()[0].clone();
    let mut edit = source.to_package().edit().expect("structure edit");
    edit.set_sheet_name(0, format!("{old_name}_patched"))
        .expect("sheet rename")
        .expect("resolved worksheet selector");
    let commit = edit.commit().expect("structure commit");
    let durable = commit
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .expect("durable workbook patch");
    let decoded = Patch::from_bytes(
        &durable,
        TransferLimits::DEFAULT,
        litchi_xlsb::raw::Limits::DEFAULT,
    )
    .expect("decode valid durable workbook patch");

    let mut target = skipped(&original);
    target
        .apply_structure(&decoded)
        .expect("apply decoded patch");
    assert!(target.sheet_drawings().is_empty());
    assert_eq!(target.worksheet_names()[0], format!("{old_name}_patched"));
    assert_eq!(part_bytes(&target, &drawing_uri), drawing_before);

    let (malformed, malformed_drawing, _) = malformed_skipped_fixture();
    let malformed_before = saved(&malformed);
    let malformed_name = malformed.worksheet_names()[0].clone();
    let mut malformed_edit = malformed
        .to_package()
        .edit()
        .expect("malformed structure edit");
    malformed_edit
        .set_sheet_name(0, format!("{malformed_name}_patched"))
        .expect("malformed sheet rename")
        .expect("resolved worksheet selector");
    let malformed_commit = malformed_edit.commit().expect("malformed structure commit");
    let malformed_durable = malformed_commit
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .expect("serialize malformed-image patch");
    assert!(
        Patch::from_bytes(
            &malformed_durable,
            TransferLimits::DEFAULT,
            litchi_xlsb::raw::Limits::DEFAULT
        )
        .is_err(),
        "durable decoding remains eager and therefore rejects malformed drawing images"
    );

    let decoded_skipped = Patch::from_bytes_without_drawing_parse(
        &malformed_durable,
        TransferLimits::DEFAULT,
        litchi_xlsb::raw::Limits::DEFAULT,
    )
    .expect("decode malformed-image patch through explicit skipped projection");
    let mut durable_skipped_target = skipped(&malformed_before);
    durable_skipped_target
        .apply_structure(&decoded_skipped)
        .expect("apply decoded skipped patch");
    assert!(durable_skipped_target.sheet_drawings().is_empty());
    assert_eq!(
        durable_skipped_target.worksheet_names()[0],
        format!("{malformed_name}_patched")
    );
    assert_eq!(
        part_bytes(&durable_skipped_target, &malformed_drawing),
        b"<broken-drawing/>"
    );

    let mut malformed_target = skipped(&malformed_before);
    malformed_target
        .apply_structure(malformed_commit.patch())
        .expect("apply in-memory patch to skipped target");
    assert!(malformed_target.sheet_drawings().is_empty());
    assert_eq!(
        malformed_target.worksheet_names()[0],
        format!("{malformed_name}_patched")
    );
    assert_eq!(
        part_bytes(&malformed_target, &malformed_drawing),
        b"<broken-drawing/>"
    );
}

#[test]
fn package_conversion_and_publication_retain_skipped_drawings() {
    let (workbook, referenced, unrelated) = malformed_skipped_fixture();
    let original = package_blobs(&workbook);
    let package = workbook.into_package();
    assert!(!package.drawings_parsed());
    let projected = package.workbook().expect("skipped package projection");
    assert!(!projected.drawings_parsed());
    assert_eq!(package_blobs(projected), original);

    // Even an unrelated package publication must not enter the drawing parser.
    let published = package.unsign();
    assert!(!published.drawings_parsed());
    let reopened = published.owned_workbook().expect("consuming projection");
    assert!(!reopened.drawings_parsed());
    assert_eq!(package_blobs(&reopened), original);
    assert_eq!(part_bytes(&reopened, &referenced), b"<broken-drawing/>");
    assert_eq!(
        part_bytes(&reopened, &unrelated),
        b"<unrelated-broken-drawing/>"
    );
    assert!(Workbook::new(Cursor::new(saved(&reopened))).is_err());
}
