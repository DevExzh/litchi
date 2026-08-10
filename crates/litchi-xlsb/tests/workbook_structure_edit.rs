#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration tests use panic-on-failure assertions"
)]

use litchi_core::sheet::traits::WorkbookTrait;
use litchi_xlsb::Workbook;
use litchi_xlsb::cell_values::{
    AuthoredStyle, CellFormula, Reference, TransferLimits, Value, WorkbookHistory, WorkbookPatch,
};
use litchi_xlsb::package::{SharedString, SharedStringRun};
use litchi_xlsb::styles::{Alignment, Fill, Font, HorizontalAlignment, VerticalAlignment};
use litchi_xlsb::writer::{ChartAnchor, Image, ImageFormat, MutableWorksheet, WorkbookWriter};
use std::fs::File;
use std::io::Cursor;
use std::path::PathBuf;

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

fn producer_workbook(name: &str, initial_value: Option<&str>) -> Workbook {
    let mut sheet = MutableWorksheet::new(name);
    if let Some(value) = initial_value {
        sheet.set_cell(0, 0, value);
    }
    let mut producer = WorkbookWriter::new();
    producer.add_worksheet(sheet);
    let mut bytes = Cursor::new(Vec::new());
    producer.save(&mut bytes).expect("producer save");
    Workbook::new(Cursor::new(bytes.into_inner())).expect("producer reopen")
}

#[test]
fn producer_reopen_transfers_sst_and_renames_on_a_durable_root() {
    let source = producer_workbook("Source", Some("shared transfer value"));
    let mut target = producer_workbook("Target", None);
    let destination = Reference::new(4, 3).expect("destination");

    let mut edit = target
        .edit_workbook_structure()
        .expect("detached workbook edit");
    edit.rename_sheet(0, "Imported".to_string())
        .expect("rename");
    edit.transfer_cell(
        &source,
        0,
        Reference::new(0, 0).expect("source reference"),
        0,
        destination,
    )
    .expect("dependency-managed transfer");
    let commit = edit.commit().expect("workbook commit");
    let encoded = commit
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .expect("durable patch");
    let decoded = WorkbookPatch::from_bytes(&encoded, TransferLimits::DEFAULT)
        .expect("validated durable patch");
    decoded.apply(&mut target).expect("atomic publication");

    assert_eq!(target.worksheet_names(), &["Imported".to_string()]);
    let snapshot = target.cell_values(0).expect("target cells");
    let cell = snapshot
        .cell(destination)
        .expect("unique destination")
        .expect("transferred destination");
    let Value::SharedStringIndex(index) = cell.value() else {
        panic!("writer-produced string should remain SST-backed")
    };
    assert_eq!(
        target.shared_strings()[usize::try_from(*index).expect("SST index")].text,
        "shared transfer value"
    );

    let mut bytes = Cursor::new(Vec::new());
    target.save(&mut bytes).expect("save transferred workbook");
    let reopened = Workbook::new(Cursor::new(bytes.into_inner())).expect("consumer reopen");
    assert_eq!(reopened.worksheet_names(), &["Imported".to_string()]);
    assert!(
        reopened
            .cell_values(0)
            .expect("reopened cells")
            .cell(destination)
            .expect("unique destination")
            .is_some()
    );
}

#[test]
fn ordinary_root_authors_sst_rich_style_and_formula_resources() {
    const GIF_1X1: &[u8] = &[
        0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x01, 0x00, 0x01, 0x00, 0x80, 0x00, 0x00, 0x00, 0x00,
        0x00, 0xff, 0xff, 0xff, 0x21, 0xf9, 0x04, 0x01, 0x00, 0x00, 0x00, 0x00, 0x2c, 0x00, 0x00,
        0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00, 0x02, 0x02, 0x44, 0x01, 0x00, 0x3b,
    ];
    let mut workbook = producer_workbook("Target", None);
    let style = AuthoredStyle {
        font: Font {
            name: "Arial".to_string(),
            size: 12.0,
            color: Some(0xff_20_40_80),
            bold: true,
            italic: false,
            underline: false,
            strike: false,
        },
        fill: Fill {
            pattern_type: 1,
            fg_color: Some(0xff_f0_e0_20),
            bg_color: None,
        },
        number_format: Some("0.0000".to_string()),
        alignment: Some(Alignment {
            horizontal: HorizontalAlignment::Center,
            vertical: VerticalAlignment::Center,
            rotation: 0,
            indent: 0,
            text_direction: 0,
            wrap_text: true,
            shrink_to_fit: false,
        }),
        ..AuthoredStyle::default()
    };
    let shared_ref = Reference::new(1, 1).expect("shared reference");
    let rich_ref = Reference::new(1, 2).expect("rich reference");
    let formula_ref = Reference::new(1, 3).expect("formula reference");
    let formula_string_ref = Reference::new(1, 4).expect("string-formula reference");
    let string = SharedString {
        text: "authored rich value".to_string(),
        runs: vec![SharedStringRun {
            character_index: 0,
            font_id: u16::MAX,
        }],
        phonetic: None,
    };
    let mut edit = workbook.edit_workbook_structure().expect("root edit");
    edit.insert_shared_string(0, shared_ref, string.clone(), &style)
        .expect("SST authoring");
    edit.insert_rich_string(0, rich_ref, string, &style)
        .expect("rich-string authoring");
    edit.insert_formula_number(
        0,
        formula_ref,
        2.0,
        CellFormula::new(0, vec![0x1e, 2, 0], vec![]).expect("constant formula"),
        &style,
    )
    .expect("formula authoring");
    edit.insert_formula(
        0,
        formula_string_ref,
        Value::FormulaStringCache("two".to_string()),
        CellFormula::new(0, vec![0x1e, 2, 0], vec![]).expect("constant string-cache formula"),
        &style,
    )
    .expect("string-cache formula authoring");
    let image = Image::new(
        GIF_1X1.to_vec(),
        ImageFormat::Gif,
        ChartAnchor::new(0, 3, 2, 7),
    )
    .expect("valid image")
    .with_description("authored one-pixel image")
    .expect("valid image description");
    edit.insert_image(0, image).expect("image authoring");
    let commit = edit.commit().expect("root commit");
    let bytes = commit
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .expect("durable resources");
    let patch = WorkbookPatch::from_bytes(&bytes, TransferLimits::DEFAULT)
        .expect("durable resource replay");
    patch.apply(&mut workbook).expect("publish resources");

    let snapshot = workbook.cell_values(0).expect("authored snapshot");
    assert!(matches!(
        snapshot
            .cell(shared_ref)
            .expect("lookup")
            .expect("shared")
            .value(),
        Value::SharedStringIndex(_)
    ));
    assert!(matches!(
        snapshot
            .cell(rich_ref)
            .expect("lookup")
            .expect("rich")
            .value(),
        Value::RichString(_)
    ));
    assert_eq!(
        snapshot
            .cell(formula_ref)
            .expect("lookup")
            .expect("formula")
            .value(),
        &Value::FormulaNumberCache(2.0)
    );
    assert_eq!(
        snapshot
            .cell(formula_string_ref)
            .expect("lookup")
            .expect("string formula")
            .value(),
        &Value::FormulaStringCache("two".to_string())
    );
    assert!(workbook.styles().get_font(1).is_some());
    let drawing = workbook.sheet_drawing(0).expect("authored drawing");
    assert_eq!(drawing.images.len(), 1);
    assert_eq!(drawing.images[0].data.as_ref(), GIF_1X1);
    let mut reopened_bytes = Cursor::new(Vec::new());
    workbook.save(&mut reopened_bytes).expect("save authored");
    let reopened =
        Workbook::new(Cursor::new(reopened_bytes.into_inner())).expect("reopen authored");
    assert_eq!(
        reopened.cell_values(0).expect("reopen cells").cells().len(),
        4
    );
    assert_eq!(
        reopened
            .sheet_drawing(0)
            .expect("reopened drawing")
            .images
            .len(),
        1
    );

    let mut transfer_target = producer_workbook("Image target", None);
    let mut transfer = transfer_target
        .edit_workbook_structure()
        .expect("image transfer edit");
    transfer
        .transfer_image(&reopened, 0, 0, 0, ChartAnchor::new(3, 4, 6, 9))
        .expect("dependency-closed image transfer");
    let transfer_commit = transfer.commit().expect("image transfer commit");
    transfer_target
        .apply_workbook_structure(&transfer_commit)
        .expect("publish transferred image");
    let transferred = transfer_target
        .sheet_drawing(0)
        .expect("transferred drawing");
    assert_eq!(transferred.images[0].data.as_ref(), GIF_1X1);
    assert_eq!(
        transferred.images[0].description.as_deref(),
        Some("authored one-pixel image")
    );
}

#[test]
fn third_party_fixture_cell_transfers_with_its_style_dependency_closure() {
    let fixtures = [
        "test-data/ooxml/xlsb/sample.xlsb",
        "test-data/ooxml/xlsb/Simple.xlsb",
        "test-data/ooxml/xlsb/date.xlsb",
        "test-data/ooxml/xlsb/51519.xlsb",
    ];
    let (source, sheet, source_reference) = fixtures
        .iter()
        .find_map(|relative| {
            let workbook = Workbook::new(File::open(fixture(relative)).ok()?).ok()?;
            let selected = (0..workbook.worksheet_count()).find_map(|sheet| {
                workbook.cell_values(sheet).ok().and_then(|snapshot| {
                    snapshot
                        .cells()
                        .find(|cell| cell.formula().is_none())
                        .map(|cell| (sheet, cell.reference()))
                })
            });
            selected.map(|(sheet, reference)| (workbook, sheet, reference))
        })
        .expect("a third-party fixture contains a direct stored cell");
    let destination = Reference::new(10, 7).expect("destination");
    let mut target = producer_workbook("Target", None);
    let mut edit = target.edit_workbook_structure().expect("root edit");
    edit.transfer_cell(&source, sheet, source_reference, 0, destination)
        .expect("third-party transfer plan");
    let commit = edit.commit().expect("third-party transfer commit");
    target
        .apply_workbook_structure(&commit)
        .expect("publish third-party transfer");

    let transferred_snapshot = target.cell_values(0).expect("transferred snapshot");
    let transferred = transferred_snapshot
        .cell(destination)
        .expect("unique destination")
        .expect("transferred cell");
    assert!(
        target
            .styles()
            .get_cell_format(usize::try_from(transferred.style().get()).expect("style index"))
            .is_some()
    );
    let mut bytes = Cursor::new(Vec::new());
    target.save(&mut bytes).expect("save transferred fixture");
    let reopened = Workbook::new(Cursor::new(bytes.into_inner())).expect("reopen transfer");
    assert!(
        reopened
            .cell_values(0)
            .expect("reopened snapshot")
            .cell(destination)
            .expect("unique destination")
            .is_some()
    );
}

#[test]
fn three_way_history_and_adversarial_envelope_remain_atomic() {
    let source = producer_workbook("Source", Some("merge value"));
    let mut target = producer_workbook("Target", None);
    let destination = Reference::new(8, 2).expect("destination");

    let mut left_edit = target.edit_workbook_structure().expect("left plan");
    left_edit
        .rename_sheet(0, "Merged".to_string())
        .expect("rename");
    let left_commit = left_edit.commit().expect("left commit");

    let mut right_edit = target.edit_workbook_structure().expect("right plan");
    right_edit
        .transfer_cell(
            &source,
            0,
            Reference::new(0, 0).expect("source reference"),
            0,
            destination,
        )
        .expect("transfer");
    let right_commit = right_edit.commit().expect("right commit");
    let outcome = left_commit
        .patch()
        .merge_three_way(right_commit.patch())
        .expect("three-way merge");
    assert!(outcome.conflicts().is_empty());
    let merged = outcome.patch().expect("merged patch").clone();

    let mut corrupt = merged
        .to_bytes(TransferLimits::DEFAULT)
        .expect("encoded merge");
    let last = corrupt.len().checked_sub(1).expect("nonempty patch");
    corrupt[last] ^= 0x5a;
    assert!(WorkbookPatch::from_bytes(&corrupt, TransferLimits::DEFAULT).is_err());
    assert_eq!(target.worksheet_names(), &["Target".to_string()]);

    merged.apply(&mut target).expect("apply merge");
    assert_eq!(target.worksheet_names(), &["Merged".to_string()]);
    assert!(
        merged.apply(&mut target).is_err(),
        "stale reapply must fail"
    );

    let mut history = WorkbookHistory::new(TransferLimits::DEFAULT).expect("history");
    history.push(merged).expect("retain patch");
    history.undo(&mut target).expect("undo");
    assert_eq!(target.worksheet_names(), &["Target".to_string()]);
    assert!(
        target
            .cell_values(0)
            .expect("undone cells")
            .cell(destination)
            .expect("unique destination")
            .is_none()
    );
    history.redo(&mut target).expect("redo");
    assert_eq!(target.worksheet_names(), &["Merged".to_string()]);
}
