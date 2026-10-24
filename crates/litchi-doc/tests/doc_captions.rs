#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_wrap,
    clippy::let_underscore_must_use,
    clippy::manual_midpoint,
    clippy::map_unwrap_or,
    clippy::needless_pass_by_value,
    clippy::shadow_reuse,
    clippy::wildcard_enum_match_arm,
    clippy::bool_assert_comparison,
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    clippy::decimal_bitwise_operands,
    clippy::default_trait_access,
    clippy::doc_markdown,
    clippy::expect_used,
    clippy::field_reassign_with_default,
    clippy::float_cmp,
    clippy::implicit_clone,
    clippy::items_after_statements,
    clippy::manual_let_else,
    clippy::manual_repeat_n,
    clippy::manual_string_new,
    clippy::match_wildcard_for_single_variants,
    clippy::needless_raw_string_hashes,
    clippy::redundant_closure_for_method_calls,
    clippy::shadow_unrelated,
    clippy::similar_names,
    clippy::uninlined_format_args,
    clippy::unreadable_literal,
    clippy::unwrap_used,
    reason = "integration-test fixtures favor explicit wire values and concise panic-driven assertions over production-style ergonomics"
)]

use litchi_cfb::OleFile;
use litchi_doc::Package;
use litchi_doc::captions::{
    AutoEntry, AutoTable, Definition, Format, Heading, Info, LabelTable, Location, Numbering,
    Separator, Tables,
};
use std::io::Cursor;

mod common;

#[test]
fn facade_exposes_contextual_caption_types_without_repeated_prefixes() {
    let info = Info::new(
        Location::Above,
        Some(Numbering::new(Heading::Level1, Separator::Period)),
        false,
        Format::Arabic,
    );
    let labels =
        LabelTable::try_new(vec![Definition::try_new("Figure".into(), info).unwrap()]).unwrap();
    let auto = AutoTable::try_new(vec![
        AutoEntry::try_new("Word.Picture.8".into(), 0).unwrap(),
    ])
    .unwrap();
    let tables = Tables::try_new(Some(labels.clone()), Some(auto.clone())).unwrap();

    assert_eq!(tables.labels(), Some(&labels));
    assert_eq!(tables.auto(), Some(&auto));
    assert_eq!(
        labels.to_bytes().unwrap(),
        LabelTable::parse_bytes(&labels.to_bytes().unwrap())
            .unwrap()
            .to_bytes()
            .unwrap()
    );
    assert_eq!(
        auto.to_bytes().unwrap(),
        AutoTable::parse_bytes(&auto.to_bytes().unwrap())
            .unwrap()
            .to_bytes()
            .unwrap()
    );
}

#[test]
fn package_edit_publishes_caption_crud_without_rewriting_unrelated_table_bytes() {
    let original = common::authored_template("caption source");
    let original_table = table_stream(&original);
    let package = Package::from_bytes(&original).expect("template package opens");
    let mut edit = package.edit_captions().expect("template edit opens");
    assert!(!edit.captions().expect("caption snapshot").is_present());

    let tables = sample_tables();
    edit.set(tables.clone()).expect("caption tables stage");
    let committed = edit.commit().expect("caption tables publish");
    assert!(committed.diagnostics().changed());
    let bytes = committed.patch().after_bytes().to_vec();
    let mut reopened = committed
        .package()
        .edit_captions()
        .expect("published package reopens");
    assert_eq!(reopened.tables().expect("caption tables"), &tables);

    let table = table_stream(&bytes);
    assert!(table.starts_with(&original_table));
    let restored = committed
        .package()
        .apply_captions(&committed.patch().inverse())
        .expect("inverse patch applies");
    assert_eq!(restored.package().to_bytes().unwrap(), original);

    let mut cleared = committed
        .package()
        .edit_captions()
        .expect("published package opens for clear");
    cleared.clear().expect("caption ranges clear");
    let cleared_bytes = cleared
        .commit()
        .expect("clear renders")
        .patch()
        .after_bytes()
        .to_vec();
    let mut ole = OleFile::open(
        Cursor::new(cleared_bytes),
        litchi_cfb::OleFileLimits::default(),
    )
    .expect("cleared CFB opens");
    let word = ole
        .open_stream(&["WordDocument"])
        .expect("WordDocument stream exists");
    let caption_pointer = 154 + 52 * 8;
    let auto_pointer = 154 + 53 * 8;
    assert_eq!(&word[caption_pointer..caption_pointer + 8], &[0; 8]);
    assert_eq!(&word[auto_pointer..auto_pointer + 8], &[0; 8]);
}

#[test]
fn package_edit_rejects_stale_record_edits_and_keeps_noop_bytes_exact() {
    let package =
        Package::from_bytes(&common::authored_template("untouched")).expect("template opens");
    let mut edit = package.edit_captions().expect("template edit opens");
    let labels = sample_tables().labels().cloned().expect("labels");
    edit.replace_labels(labels).expect("labels stage");
    let published = edit.commit().expect("labels publish");
    let exact = published.patch().after_bytes().to_vec();

    let mut noop = published
        .package()
        .edit_captions()
        .expect("published package reopens");
    let tables = noop.tables().expect("caption tables").clone();
    noop.set(tables).expect("semantic no-op stages");
    let noop_commit = noop.commit().expect("semantic no-op publishes");
    assert!(!noop_commit.diagnostics().changed());
    assert_eq!(noop_commit.package().to_bytes().unwrap(), exact);

    let mut editor = published
        .package()
        .edit_captions()
        .expect("published package reopens");
    let stale = editor.captions().expect("caption snapshot").edit();
    editor
        .replace_auto(
            AutoTable::try_new(vec![
                AutoEntry::try_new("Word.Picture.8".into(), 0).unwrap(),
            ])
            .unwrap(),
        )
        .expect("auto-caption rules stage");
    assert!(editor.apply(stale).is_err());
}

#[test]
fn package_edit_defers_caption_range_rejection_on_ordinary_documents() {
    let package = Package::from_bytes(&common::authored_document("table")).expect("document opens");
    let mut edit = package
        .edit_captions()
        .expect("caption edit construction is deferred");
    assert!(edit.captions().is_err());
}

fn table_stream(bytes: &[u8]) -> Vec<u8> {
    let mut ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
        .expect("fixture CFB opens");
    let name = if ole.exists(&["1Table"]) {
        "1Table"
    } else {
        "0Table"
    };
    ole.open_stream(&[name]).expect("table stream exists")
}

fn sample_tables() -> Tables {
    let info = Info::new(Location::Above, None, false, Format::Arabic);
    let labels =
        LabelTable::try_new(vec![Definition::try_new("Figure".into(), info).unwrap()]).unwrap();
    let auto = AutoTable::try_new(vec![
        AutoEntry::try_new("Word.Picture.8".into(), 0).unwrap(),
    ])
    .unwrap();
    Tables::try_new(Some(labels), Some(auto)).unwrap()
}
