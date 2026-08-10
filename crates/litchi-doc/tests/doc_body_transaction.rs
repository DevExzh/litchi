#![expect(
    clippy::expect_used,
    reason = "integration fixtures intentionally fail fast with contextual assertions"
)]

use litchi_cfb::OleWriter;
use litchi_core::Position;
use litchi_doc::Package;
use litchi_doc::body_text::{CharacterProperty, Projection, Snapshot, Story, TextTarget};
use litchi_doc::writer::{FloatingPosition, Kind, Picture, Shape, Writer};
use std::io::Cursor;
use std::path::PathBuf;

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

#[test]
fn real_multi_generation_docs_resize_and_fully_reopen() {
    let fixtures = [
        ("test-data/ole/doc/NoHeadFoot.doc", 0x00C1),
        ("test-data/ole/doc/documentProperties.doc", 0x0101),
    ];
    for (relative, expected_nfib) in fixtures {
        let bytes = std::fs::read(fixture(relative)).expect("real DOC fixture");
        let source = Snapshot::parse(&bytes).expect("fixture has a safe edit basis");
        let paragraphs = source
            .paragraphs(Projection::All)
            .expect("fixture paragraphs");
        let target = paragraphs
            .iter()
            .find(|paragraph| !paragraph.text().is_empty())
            .expect("fixture has an ordinary paragraph");
        let replacement = format!("{} [resized]", target.text());
        let mut transaction = source.edit().expect("fixture transaction");
        transaction
            .replace_paragraph(target.position(), &replacement)
            .expect("fixture length change");
        let commit = transaction.commit().expect("fixture commit");
        let mut package = Package::from_reader(Cursor::new(commit.snapshot().finish()))
            .expect("edited CFB reopens");
        let document = package.document().expect("edited DOC reopens");
        assert_eq!(document.fib().version(), expected_nfib);
        assert_eq!(
            commit
                .snapshot()
                .paragraphs(Projection::All)
                .expect("semantic readback")[target.position().get()]
            .text(),
            replacement
        );
    }
}

#[test]
fn durable_stale_source_is_rejected_without_mutation() {
    let first = Snapshot::parse(
        &std::fs::read(fixture("test-data/ole/doc/NoHeadFoot.doc")).expect("first fixture"),
    )
    .expect("first snapshot");
    let second = Snapshot::parse(
        &std::fs::read(fixture("test-data/ole/doc/ThreeColHeadFoot.doc")).expect("second fixture"),
    )
    .expect("second snapshot");
    let paragraph = first
        .paragraphs(Projection::All)
        .expect("paragraphs")
        .into_iter()
        .find(|paragraph| !paragraph.text().is_empty())
        .expect("ordinary paragraph");
    let mut edit = first.edit().expect("edit");
    edit.replace_paragraph(
        Position::new(paragraph.position().get()),
        "durable replacement",
    )
    .expect("replacement");
    let commit = edit.commit().expect("commit");
    let limits = litchi_core::PatchLimits::new(
        litchi_core::BlobLimits::new(0, 0, 0),
        128 * 1024,
        16,
        8,
        16 * 1024,
        64 * 1024,
    );
    let patch = commit.patch().to_durable(limits).expect("durable patch");
    assert!(second.apply_durable(&patch).is_err());
}

#[test]
fn real_story_field_and_table_targets_fully_reopen() {
    let header_bytes =
        std::fs::read(fixture("test-data/ole/doc/ThreeColHeadFoot.doc")).expect("header fixture");
    let header_source = Snapshot::parse(&header_bytes).expect("header snapshot");
    let header = header_source
        .story_paragraphs(Story::Header)
        .expect("header story")
        .into_iter()
        .find(|item| {
            !item.text().is_empty()
                && !item
                    .text()
                    .chars()
                    .any(|character| character.is_control() && character != '\t')
        })
        .expect("header paragraph");
    let mut header_edit = header_source.edit().expect("header edit");
    header_edit
        .set_character_property(header.target(), CharacterProperty::Italic, true)
        .expect("header direct format");
    let header_commit = header_edit.commit().expect("header commit");
    let mut header_package = Package::from_reader(Cursor::new(header_commit.snapshot().finish()))
        .expect("formatted header CFB reopens");
    assert_eq!(
        header_package
            .document()
            .expect("formatted header DOC reopens")
            .fib()
            .version(),
        0x00C1
    );

    let field_bytes =
        std::fs::read(fixture("test-data/ole/doc/hyperlink.doc")).expect("field fixture");
    let field_source = Snapshot::parse(&field_bytes).expect("field snapshot");
    let field = field_source
        .field_results()
        .expect("field results")
        .into_iter()
        .find(|item| !item.text().is_empty())
        .expect("simple field result");
    let field_replacement = format!("{} updated", field.text());
    let mut field_edit = field_source.edit().expect("field edit");
    field_edit
        .replace_text(field.target(), &field_replacement)
        .expect("field result resize");
    let field_commit = field_edit.commit().expect("field commit");
    let mut field_package = Package::from_reader(Cursor::new(field_commit.snapshot().finish()))
        .expect("field CFB reopens");
    field_package.document().expect("field DOC reopens");
    assert_eq!(
        field_commit
            .snapshot()
            .field_results()
            .expect("field semantic readback")
            .into_iter()
            .find(|item| item.target() == field.target())
            .expect("same field selector")
            .text(),
        field_replacement
    );

    let table_bytes =
        std::fs::read(fixture("test-data/ole/doc/commented-table.doc")).expect("table fixture");
    let table_source = Snapshot::parse(&table_bytes).expect("table snapshot");
    let cell = table_source
        .table_cells()
        .expect("simple real table cells")
        .into_iter()
        .find(|item| !item.text().is_empty())
        .expect("non-empty real table cell");
    let cell_replacement = format!("{}!", cell.text());
    let mut table_edit = table_source.edit().expect("table edit");
    table_edit
        .replace_text(cell.target(), &cell_replacement)
        .expect("table cell resize");
    let table_commit = table_edit.commit().expect("table commit");
    let mut table_package = Package::from_reader(Cursor::new(table_commit.snapshot().finish()))
        .expect("table CFB reopens");
    table_package.document().expect("table DOC reopens");
}

#[test]
fn genuine_embedded_object_transfer_closes_cfb_field_preview_and_storage() {
    let donor_base =
        std::fs::read(fixture("test-data/ole/doc/NoHeadFoot.doc")).expect("real donor fixture");
    let mut object = OleWriter::new();
    object
        .create_stream(&["CONTENTS"], b"inert transfer payload")
        .expect("standalone object stream");
    let mut object_output = Cursor::new(Vec::new());
    object
        .write_to(&mut object_output)
        .expect("standalone object CFB");
    let mut preview = 12u32.to_le_bytes().to_vec();
    preview.extend_from_slice(&7_701u32.to_le_bytes());
    preview.extend_from_slice(&[0; 4]);
    let mut donor_owner =
        litchi_doc::Editor::open(donor_base, litchi_doc::embedded_object::Limits::default())
            .expect("real donor embedded owner");
    donor_owner
        .add(litchi_doc::WriteOptions::new(
            7_701,
            object_output.into_inner(),
            preview,
        ))
        .expect("seed inert resource on genuine donor");
    let donor = Snapshot::parse(&donor_owner.finish().expect("genuine donor publication"))
        .expect("genuine donor root snapshot");

    let receiver = Snapshot::parse(
        &std::fs::read(fixture("test-data/ole/doc/documentProperties.doc"))
            .expect("real receiver fixture"),
    )
    .expect("real receiver snapshot");
    let plan = receiver
        .plan_embedded_transfer_from(&donor, 7_701, 9_001)
        .expect("genuine dependency-closed transfer plan");

    let mut edit = receiver.edit().expect("real resource transfer edit");
    edit.apply_embedded_transfer(&plan)
        .expect("real resource transfer");
    let commit = edit.commit().expect("real resource transfer commit");
    let mut package = Package::from_reader(Cursor::new(commit.snapshot().finish()))
        .expect("transferred resource CFB reopens");
    assert_eq!(
        package
            .document()
            .expect("transferred resource DOC fully reopens")
            .fib()
            .version(),
        0x0101
    );
    assert!(
        commit
            .snapshot()
            .embedded_objects()
            .expect("transferred real inventory")
            .get(9_001)
            .is_some()
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .expect("real resource exact inverse"),
        receiver
    );

    let limits = litchi_core::PatchLimits::new(
        litchi_core::BlobLimits::new(8, 32 * 1024 * 1024, 64 * 1024 * 1024),
        96 * 1024 * 1024,
        16,
        8,
        16 * 1024,
        96 * 1024 * 1024,
    );
    let durable = commit
        .patch()
        .to_durable(limits)
        .expect("real durable closure");
    let replay = receiver
        .apply_durable(&durable)
        .expect("real durable resource replay");
    assert!(
        replay
            .embedded_objects()
            .expect("durable real inventory")
            .get(9_001)
            .is_some()
    );
}

#[test]
fn genuine_receivers_accept_pictures_beside_shapes_and_textboxes() {
    let image = std::fs::read(fixture("test-data/images/png/lena.png")).expect("PNG fixture");
    for floating in [false, true] {
        let mut writer = Writer::new();
        writer
            .insert_floating_shape(
                Shape::new(Kind::Rectangle, 640, 320).expect("neighbor shape"),
                FloatingPosition::new(120, 160),
            )
            .expect("neighbor shape run");
        writer
            .insert_floating_text_box(
                Shape::new(Kind::Rectangle, 680, 360).expect("neighbor textbox"),
                FloatingPosition::new(180, 220),
                "unrelated textbox content",
            )
            .expect("neighbor textbox run");
        let picture = Picture::new(image.clone()).expect("native PNG picture");
        if floating {
            writer
                .insert_floating_picture(picture, FloatingPosition::new(720, 360))
                .expect("floating donor picture");
        } else {
            writer
                .insert_picture(picture)
                .expect("inline donor picture");
        }
        let mut donor_bytes = Cursor::new(Vec::new());
        writer
            .write_to(&mut donor_bytes)
            .expect("picture donor DOC");
        let donor = Snapshot::parse(&donor_bytes.into_inner()).expect("picture donor snapshot");

        let receiver = Snapshot::parse(
            &std::fs::read(fixture("test-data/ole/doc/documentProperties.doc"))
                .expect("genuine receiver fixture"),
        )
        .expect("genuine receiver snapshot");
        let destination = receiver
            .paragraphs(Projection::All)
            .expect("genuine receiver paragraphs")
            .into_iter()
            .find(|paragraph| !paragraph.text().is_empty())
            .expect("genuine receiver placeholder")
            .position();
        let plan = receiver
            .plan_picture_transfer_from(
                &donor,
                TextTarget::body_paragraph(Position::new(2)),
                TextTarget::body_paragraph(destination),
            )
            .expect("bounded genuine picture transfer");
        let mut edit = receiver.edit().expect("genuine picture edit");
        edit.apply_picture_transfer(&plan)
            .expect("install genuine picture graph");
        let commit = edit.commit().expect("genuine picture full reopen");

        let mut package = Package::from_reader(Cursor::new(commit.snapshot().finish()))
            .expect("picture receiver CFB reopens");
        let document = package.document().expect("picture receiver DOC reopens");
        assert_eq!(document.fib().version(), 0x0101);
        let picture_count = document
            .paragraphs()
            .expect("reopened paragraphs")
            .into_iter()
            .flat_map(|paragraph| paragraph.runs().expect("reopened runs"))
            .filter(|run| run.image().is_some())
            .count();
        assert_eq!(picture_count, 1);
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.snapshot())
                .expect("genuine picture exact inverse"),
            receiver
        );
        let durable_limits = litchi_core::PatchLimits::new(
            litchi_core::BlobLimits::new(8, 16 * 1024 * 1024, 32 * 1024 * 1024),
            40 * 1024 * 1024,
            16,
            8,
            16 * 1024,
            40 * 1024 * 1024,
        );
        let durable = commit
            .patch()
            .to_durable(durable_limits)
            .expect("genuine picture durable patch");
        let replay = receiver
            .apply_durable(&durable)
            .expect("genuine picture durable replay");
        let mut replay_package = Package::from_reader(Cursor::new(replay.finish()))
            .expect("durable picture CFB reopens");
        assert_eq!(
            replay_package
                .document()
                .expect("durable picture DOC reopens")
                .fib()
                .version(),
            0x0101
        );
        let restored = replay
            .apply_durable(&durable.inverse())
            .expect("genuine picture durable inverse");
        assert_eq!(
            restored
                .paragraphs(Projection::All)
                .expect("durable inverse paragraphs")[destination.get()]
            .text(),
            receiver
                .paragraphs(Projection::All)
                .expect("receiver paragraphs")[destination.get()]
            .text()
        );
    }
}
