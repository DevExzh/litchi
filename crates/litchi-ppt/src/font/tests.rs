#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::{
    Facet, FontCollection, FontCollections, FontEmbeddingFlags, PackageLimits, Scope, Snapshot,
};
use crate::{Record, RecordType};
use std::io::Cursor;
use std::sync::Arc;

fn record_bytes(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    data.extend_from_slice(&kind.to_le_bytes());
    data.extend_from_slice(&u32::try_from(payload.len()).unwrap().to_le_bytes());
    data.extend_from_slice(payload);
    data
}

fn minimal_eot() -> Vec<u8> {
    let mut bytes = vec![0u8; 96];
    bytes[0..4].copy_from_slice(&96u32.to_le_bytes());
    bytes[8..12].copy_from_slice(&0x0001_0000u32.to_le_bytes());
    bytes[34..36].copy_from_slice(&0x504cu16.to_le_bytes());
    bytes
}

fn collection(kind: RecordType, payload: Vec<u8>) -> Record {
    Record {
        record_type: kind,
        record_type_raw: kind.as_u16(),
        version: 0x0f,
        instance: 0,
        data_length: u32::try_from(payload.len()).unwrap(),
        data: payload,
        children: Vec::new(),
    }
}

fn prog_tags_record(blob_payload: &[u8]) -> Record {
    let name: Vec<u8> = "___PPT10"
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect();
    let mut tag_payload = record_bytes(0, 0, 4026, &name);
    tag_payload.extend_from_slice(&record_bytes(0, 0, 0x138b, blob_payload));
    let tag = record_bytes(0x0f, 0, 0x138a, &tag_payload);
    Record {
        record_type: RecordType::ProgTags,
        record_type_raw: RecordType::ProgTags.as_u16(),
        version: 0x0f,
        instance: 0,
        data_length: u32::try_from(tag.len()).unwrap(),
        data: tag,
        children: Vec::new(),
    }
}

fn doc_info_list(prog_tags: Record) -> Record {
    Record {
        record_type: RecordType::DocInfoList,
        record_type_raw: RecordType::DocInfoList.as_u16(),
        version: 0x0f,
        instance: 0,
        data_length: prog_tags.data_length + 8,
        data: Vec::new(),
        children: vec![prog_tags],
    }
}

fn entity(instance: u16, name: &str, ignored_type_bits: u8) -> Vec<u8> {
    let mut payload = vec![0u8; 68];
    for (index, unit) in name.encode_utf16().chain(std::iter::once(0)).enumerate() {
        payload[index * 2..index * 2 + 2].copy_from_slice(&unit.to_le_bytes());
    }
    payload[65] = 0x81;
    payload[66] = ignored_type_bits | 0x0c;
    record_bytes(0, instance, RecordType::FontEntityAtom.as_u16(), &payload)
}

fn document(base: Record, ppt10: &[u8]) -> Record {
    let environment = Record {
        record_type: RecordType::Environment,
        record_type_raw: RecordType::Environment.as_u16(),
        version: 0x0f,
        instance: 0,
        data_length: base.data_length + 8,
        data: Vec::new(),
        children: vec![base],
    };
    Record {
        record_type: RecordType::Document,
        record_type_raw: RecordType::Document.as_u16(),
        version: 0x0f,
        instance: 0,
        data_length: 0,
        data: Vec::new(),
        children: vec![environment, doc_info_list(prog_tags_record(ppt10))],
    }
}

#[test]
fn ordinals_are_distinct_from_raw_instances_and_ignored_bits_round_trip() {
    let mut payload = entity(7, "Noto Sans", 0xa0);
    payload.extend_from_slice(&record_bytes(
        0,
        0,
        RecordType::FontEmbeddedData.as_u16(),
        b"opaque-eot",
    ));
    let parsed = FontCollection::parse(&collection(RecordType::FontCollection, payload)).unwrap();
    let font = parsed.get(0).unwrap();
    assert_eq!(font.index, 0);
    assert_eq!(font.raw_instance, 7);
    assert_eq!(font.font_flags, 0x81);
    assert_eq!(font.font_type_flags, 0xac);
    assert_eq!(font.facet(Facet::Plain).unwrap().bytes(), b"opaque-eot");

    let reparsed = FontCollection::parse(&parsed.to_record().unwrap()).unwrap();
    assert_eq!(reparsed, parsed);
    assert_eq!(
        reparsed.to_record().unwrap().data,
        parsed.to_record().unwrap().data
    );
}

#[test]
fn owned_parser_moves_large_facet_allocation_and_clones_share_it() {
    let large = vec![0x5a; 4 * 1024 * 1024];
    let mut payload = entity(0, "Large", 0);
    payload.extend_from_slice(&record_bytes(
        0,
        0,
        RecordType::FontEmbeddedData.as_u16(),
        &large,
    ));
    let children = Record::parse_sequence_strict(&payload, "large font fixture").unwrap();
    let original = children[1].data.as_ptr();
    let mut record = collection(RecordType::FontCollection, payload);
    record.children = children;

    let parsed = FontCollection::take_with_limits(&mut record, super::Limits::default()).unwrap();
    let data = &parsed.get(0).unwrap().facet(Facet::Plain).unwrap().data;
    assert_eq!(data.as_ptr(), original);
    let clone = data.clone();
    assert!(data.ptr_eq(&clone));
    assert!(record.children[1].data.is_empty());
}

#[test]
fn accepts_exactly_32_non_nul_utf16_units_by_ordinal() {
    let name = "12345678901234567890123456789012";
    let mut payload = vec![0u8; 68];
    for (index, unit) in name.encode_utf16().enumerate() {
        payload[index * 2..index * 2 + 2].copy_from_slice(&unit.to_le_bytes());
    }
    payload[66] = 4;
    let parsed = FontCollection::parse(&collection(
        RecordType::FontCollection,
        record_bytes(0, 11, RecordType::FontEntityAtom.as_u16(), &payload),
    ))
    .unwrap();
    assert_eq!(parsed.get(0).unwrap().name, name);
    assert_eq!(parsed.get(0).unwrap().index, 0);
    assert_eq!(parsed.get(0).unwrap().raw_instance, 11);
    assert_eq!(
        FontCollection::parse(&parsed.to_record().unwrap()).unwrap(),
        parsed
    );
}

#[test]
fn rejects_malformed_utf16_and_authored_names_over_32_units() {
    let mut malformed = vec![0u8; 68];
    malformed[..2].copy_from_slice(&0xd800u16.to_le_bytes());
    malformed[66] = 4;
    assert!(
        FontCollection::parse(&collection(
            RecordType::FontCollection,
            record_bytes(0, 0, RecordType::FontEntityAtom.as_u16(), &malformed),
        ))
        .is_err()
    );

    let mut authored = FontCollection::new(Scope::Base);
    authored.try_push(super::Font::new("x".repeat(33))).unwrap();
    assert!(authored.to_record_canonical().is_err());
}

#[test]
fn duplicate_record_instances_still_resolve_by_collection_ordinal() {
    let mut payload = entity(7, "First", 0);
    payload.extend_from_slice(&entity(7, "Second", 0));
    let parsed = FontCollection::parse(&collection(RecordType::FontCollection, payload)).unwrap();
    assert_eq!(parsed.get(0).unwrap().name, "First");
    assert_eq!(parsed.get(1).unwrap().name, "Second");
    assert_eq!(parsed.get(0).unwrap().raw_instance, 7);
    assert_eq!(parsed.get(1).unwrap().raw_instance, 7);
}

#[test]
fn authoring_accepts_129_fonts_and_refuses_the_130th() {
    let mut collection = FontCollection::new(Scope::Base);
    for ordinal in 0..129 {
        assert_eq!(
            collection
                .try_push(super::Font::new(format!("Font {ordinal}")))
                .unwrap(),
            ordinal
        );
    }
    assert!(collection.try_push(super::Font::new("Font 129")).is_err());
}

#[test]
fn contextual_owner_headers_and_zero_aggregate_limits_are_enforced() {
    let base = collection(RecordType::FontCollection, entity(0, "Base", 0));
    let valid = document(base.clone(), &[]);
    let mut malformed = valid.clone();
    malformed.children[1].version = 0;
    assert!(FontCollections::parse(&malformed).is_err());

    let mut direct = valid.clone();
    direct.children[1] = prog_tags_record(&[]);
    assert!(
        FontCollections::parse(&direct)
            .unwrap()
            .international
            .is_none()
    );

    let limits = super::Limits {
        max_fonts_per_collection: 0,
        ..super::Limits::default()
    };
    assert!(FontCollections::parse_with_limits(&valid, limits).is_err());

    let mut copied = super::Limits::default();
    copied.records.max_copied_payload_bytes = 67;
    assert!(FontCollections::parse_with_limits(&valid, copied).is_err());
}

#[test]
fn nested_programmable_tags_share_one_record_budget() {
    let mut tags = Vec::new();
    for _ in 0..8 {
        let name: Vec<u8> = "___PPT9"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        let mut pair = record_bytes(0, 0, RecordType::CString.as_u16(), &name);
        pair.extend_from_slice(&record_bytes(0, 0, RecordType::BinaryTagData.as_u16(), &[]));
        tags.extend_from_slice(&record_bytes(
            0x0f,
            0,
            RecordType::ProgBinaryTag.as_u16(),
            &pair,
        ));
    }
    let prog_tags = Record {
        record_type: RecordType::ProgTags,
        record_type_raw: RecordType::ProgTags.as_u16(),
        version: 0x0f,
        instance: 0,
        data_length: u32::try_from(tags.len()).unwrap(),
        data: tags,
        children: Vec::new(),
    };
    let mut root = document(
        collection(RecordType::FontCollection, entity(0, "Base", 0)),
        &[],
    );
    root.children[1] = doc_info_list(prog_tags);
    let mut limits = super::Limits::default();
    limits.records.max_records = 12;
    assert!(matches!(
        FontCollections::parse_with_limits(&root, limits),
        Err(crate::Error::ResourceLimit(_))
    ));
}

#[test]
fn eot_metadata_rejects_forged_prefixes_and_trailing_bytes() {
    let mut bytes = minimal_eot();
    let facet = super::EmbeddedFont::new(Facet::Plain, bytes.clone()).unwrap();
    assert!(facet.eot_metadata().is_some());

    bytes.extend_from_slice(&[0]);
    assert!(super::EmbeddedFont::new(Facet::Plain, bytes.clone()).is_err());
    assert!(
        super::EmbeddedFont::from_preserved(Facet::Plain, bytes)
            .eot_metadata()
            .is_none()
    );
    assert!(super::EmbeddedFont::new(Facet::Plain, vec![0u8; 36]).is_err());
}

#[test]
fn discovers_only_live_environment_and_document_ppt10_owners() {
    let base = collection(RecordType::FontCollection, entity(12, "Base", 0));
    let mut ppt10 = record_bytes(
        0x0f,
        0,
        RecordType::FontCollection10.as_u16(),
        &entity(44, "Intl", 0),
    );
    ppt10.extend_from_slice(&record_bytes(
        0,
        0,
        RecordType::FontEmbedFlags10Atom.as_u16(),
        &0xffff_ffffu32.to_le_bytes(),
    ));
    let parsed = FontCollections::parse(&document(base, &ppt10)).unwrap();
    assert_eq!(parsed.get_base(0).unwrap().raw_instance, 12);
    assert_eq!(parsed.get_international(0).unwrap().raw_instance, 44);
    assert_eq!(parsed.embedding_flags.unwrap().raw, u32::MAX);
}

#[test]
fn collection_mutations_keep_ordinals_stable() {
    let mut collection = FontCollection::parse(&collection(
        RecordType::FontCollection,
        entity(7, "First", 0),
    ))
    .unwrap();
    let mut appended = collection.get(0).unwrap().clone();
    appended.name = "Second".into();
    assert_eq!(collection.try_push(appended).unwrap(), 1);
    collection.set_facet(1, Facet::Bold, minimal_eot()).unwrap();
    assert_eq!(collection.get(0).unwrap().raw_instance, 7);
    assert_eq!(collection.get(1).unwrap().index, 1);
    assert!(collection.remove_facet(1, Facet::Bold).unwrap().is_some());
}

#[test]
fn whole_package_transaction_is_exact_reversible_and_refuses_reindexing() {
    let mut writer = crate::Writer::new();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let source = Snapshot::from_bytes(output.into_inner()).unwrap();
    let mut transaction = source.edit().unwrap();
    let mut replacement = source.fonts().get_base(0).unwrap().clone();
    replacement.name = "Transaction Font".into();
    transaction
        .replace_font(Scope::Base, 0, replacement)
        .unwrap();
    assert!(transaction.remove_font(Scope::Base, 0).is_err());
    assert!(transaction.reorder_fonts(Scope::Base, &[0]).is_err());
    let commit = transaction.commit().unwrap();
    assert_eq!(commit.fonts().get_base(0).unwrap().name, "Transaction Font");
    let replay = commit.patch().apply(&source).unwrap();
    assert_eq!(replay.bytes(), commit.snapshot().bytes());
    assert_eq!(
        commit.patch().apply(&replay).unwrap().bytes(),
        replay.bytes()
    );
    assert_eq!(commit.undo(&replay).unwrap().bytes(), source.bytes());
}

fn minimal_live_snapshot() -> Snapshot {
    let mut writer = crate::Writer::new();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    Snapshot::from_bytes(output.into_inner()).unwrap()
}

fn document_saves_fonts(bytes: &[u8]) -> bool {
    let mut package = crate::Package::from_reader(Cursor::new(bytes)).unwrap();
    package
        .presentation()
        .unwrap()
        .document_atom()
        .unwrap()
        .unwrap()
        .save_with_fonts
}

#[test]
fn live_cfb_facet_lifecycle_synchronizes_document_atom_and_stable_index() {
    let source = minimal_live_snapshot();
    assert!(!document_saves_fonts(source.bytes()));
    let raw_instance = source.fonts().get_base(0).unwrap().raw_instance;

    let mut add = source.edit().unwrap();
    let eot = minimal_eot();
    add.set_facet(Scope::Base, 0, Facet::Plain, eot.clone())
        .unwrap();
    let embedded = add.commit().unwrap();
    assert!(document_saves_fonts(embedded.snapshot().bytes()));
    let embedded_payload = embedded
        .fonts()
        .get_base(0)
        .unwrap()
        .facet(Facet::Plain)
        .unwrap()
        .bytes();
    assert_eq!(embedded_payload, eot.as_slice());

    let mut replace = embedded.snapshot().edit().unwrap();
    assert_eq!(
        replace
            .fonts()
            .get_base(0)
            .unwrap()
            .facet(Facet::Plain)
            .unwrap()
            .bytes()
            .as_ptr(),
        embedded_payload.as_ptr(),
        "snapshot edits must share unchanged EOT payload storage"
    );
    let mut replacement = replace.fonts().get_base(0).unwrap().clone();
    replacement.name = "Stable Ordinal".into();
    replace.replace_font(Scope::Base, 0, replacement).unwrap();
    let replaced = replace.commit().unwrap();
    assert_eq!(replaced.fonts().get_base(0).unwrap().index, 0);
    assert_eq!(
        replaced.fonts().get_base(0).unwrap().raw_instance,
        raw_instance
    );
    assert!(document_saves_fonts(replaced.snapshot().bytes()));

    let mut remove = replaced.snapshot().edit().unwrap();
    assert!(remove.remove_facet(Scope::Base, 0, Facet::Plain).unwrap());
    let cleared = remove.commit().unwrap();
    assert!(!document_saves_fonts(cleared.snapshot().bytes()));
}

#[test]
fn failed_mutation_rolls_back_without_staging_changes() {
    let source = minimal_live_snapshot();
    let mut transaction = source.edit().unwrap();
    let before = transaction.fonts().clone();
    assert!(
        transaction
            .set_facet(Scope::Base, u16::MAX, Facet::Plain, vec![1])
            .is_err()
    );
    assert_eq!(transaction.fonts(), &before);
    assert!(transaction.changes().is_empty());
    assert_eq!(transaction.rollback().bytes(), source.bytes());
}

#[test]
fn stale_patch_is_rejected_and_borrowed_source_limit_precedes_copy() {
    let source = minimal_live_snapshot();
    let mut left = source.edit().unwrap();
    let mut left_font = left.fonts().get_base(0).unwrap().clone();
    left_font.name = "Left".into();
    left.replace_font(Scope::Base, 0, left_font).unwrap();
    let left_commit = left.commit().unwrap();

    let mut right = source.edit().unwrap();
    let mut right_font = right.fonts().get_base(0).unwrap().clone();
    right_font.name = "Right".into();
    right.replace_font(Scope::Base, 0, right_font).unwrap();
    let right_commit = right.commit().unwrap();
    assert!(left_commit.patch().apply(right_commit.snapshot()).is_err());
    assert!(left_commit.patch().undo(right_commit.snapshot()).is_err());

    let limits = PackageLimits {
        max_source_bytes: source.bytes().len() - 1,
        ..PackageLimits::default()
    };
    assert!(Snapshot::parse_with_limits(source.bytes(), limits).is_err());
}

#[test]
fn embedding_flags_preserve_undefined_bits() {
    let flags = FontEmbeddingFlags {
        raw: 0xffff_fffd,
        subset: true,
        subset_option_confirmed: false,
    };
    assert_eq!(
        FontEmbeddingFlags::parse(&flags.to_record().unwrap()).unwrap(),
        flags
    );
}

#[test]
fn transaction_preserves_vec_and_arc_facet_allocations() {
    let source = minimal_live_snapshot();
    let mut transaction = source.edit().unwrap();

    let plain = minimal_eot();
    let plain_ptr = plain.as_ptr();
    transaction
        .set_facet(Scope::Base, 0, Facet::Plain, plain)
        .unwrap();
    assert_eq!(
        transaction
            .fonts()
            .get_base(0)
            .unwrap()
            .facet(Facet::Plain)
            .unwrap()
            .bytes()
            .as_ptr(),
        plain_ptr
    );

    let bold: Arc<[u8]> = Arc::from(minimal_eot());
    let bold_owner = super::SharedFontData::from(bold.clone());
    transaction
        .set_facet(Scope::Base, 0, Facet::Bold, bold)
        .unwrap();
    assert!(
        transaction
            .fonts()
            .get_base(0)
            .unwrap()
            .facet(Facet::Bold)
            .unwrap()
            .data
            .ptr_eq(&bold_owner)
    );
}

#[test]
fn repeated_staging_failure_retains_the_last_committable_candidate() {
    let source = minimal_live_snapshot();
    let mut limits = PackageLimits::default();
    limits.fonts.max_facets = 1;
    limits.fonts.max_embedded_bytes = minimal_eot().len();
    let bounded = Snapshot::from_bytes_with_limits(source.bytes().to_vec(), limits).unwrap();
    let mut transaction = bounded.edit().unwrap();

    transaction
        .set_facet(Scope::Base, 0, Facet::Plain, minimal_eot())
        .unwrap();
    assert!(
        transaction
            .set_facet(Scope::Base, 0, Facet::Bold, minimal_eot())
            .is_err()
    );
    assert_eq!(transaction.changes().len(), 1);
    let font = transaction.fonts().get_base(0).unwrap();
    assert!(font.facet(Facet::Plain).is_some());
    assert!(font.facet(Facet::Bold).is_none());
}

#[test]
fn changed_publication_is_refused_before_exceeding_the_source_budget() {
    let source = minimal_live_snapshot();
    let limits = PackageLimits {
        max_source_bytes: source.bytes().len(),
        ..PackageLimits::default()
    };
    let bounded = Snapshot::from_bytes_with_limits(source.bytes().to_vec(), limits).unwrap();
    let mut transaction = bounded.edit().unwrap();
    let mut font = transaction.fonts().get_base(0).unwrap().clone();
    font.name = "Budgeted".into();
    transaction.replace_font(Scope::Base, 0, font).unwrap();
    assert!(transaction.commit().is_err());
}

#[test]
fn last_facet_removal_clears_subset_state_and_missing_owner_is_unchanged() {
    let source = minimal_live_snapshot();
    let mut transaction = source.edit().unwrap();
    transaction
        .set_facet(Scope::Base, 0, Facet::Plain, minimal_eot())
        .unwrap();
    let mut font = transaction.fonts().get_base(0).unwrap().clone();
    font.embedded_subset = true;
    font.font_flags |= 1;
    transaction.replace_font(Scope::Base, 0, font).unwrap();
    assert!(
        transaction
            .remove_facet(Scope::Base, 0, Facet::Plain)
            .unwrap()
    );
    let cleared_font = transaction.fonts().get_base(0).unwrap();
    assert!(!cleared_font.embedded_subset);
    assert_eq!(cleared_font.font_flags & 1, 0);

    let before = transaction.fonts().clone();
    let changes = transaction.changes().len();
    assert!(
        transaction
            .append_font(Scope::International, super::Font::new("Unavailable"))
            .is_err()
    );
    assert_eq!(transaction.fonts(), &before);
    assert_eq!(transaction.changes().len(), changes);
}

#[test]
fn live_transaction_appends_base_and_international_fonts_and_updates_flags() {
    let mut writer = crate::Writer::new();
    writer
        .add_font_model(Scope::International, super::Font::new("Intl Zero"))
        .unwrap();
    writer
        .set_font_embedding_flags(Some(FontEmbeddingFlags::new(false, false)))
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let source = Snapshot::from_bytes(output.into_inner()).unwrap();

    let mut transaction = source.edit().unwrap();
    assert_eq!(
        transaction
            .append_font(Scope::Base, super::Font::new("Base One"))
            .unwrap(),
        1
    );
    assert_eq!(
        transaction
            .append_font(Scope::International, super::Font::new("Intl One"))
            .unwrap(),
        1
    );
    let flags = FontEmbeddingFlags::new(true, true);
    transaction.set_embedding_flags(Some(flags)).unwrap();
    let commit = transaction.commit().unwrap();
    assert_eq!(commit.fonts().get_base(1).unwrap().name, "Base One");
    assert_eq!(
        commit.fonts().get_international(1).unwrap().name,
        "Intl One"
    );
    assert_eq!(commit.fonts().embedding_flags, Some(flags));
}

#[test]
fn live_transaction_accepts_font_129_and_atomically_refuses_font_130() {
    let mut writer = crate::Writer::new();
    for ordinal in 1..128 {
        writer
            .add_font_model(Scope::Base, super::Font::new(format!("Boundary {ordinal}")))
            .unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let source = Snapshot::from_bytes(output.into_inner()).unwrap();
    assert_eq!(source.fonts().base.as_ref().unwrap().fonts.len(), 128);

    let mut transaction = source.edit().unwrap();
    assert_eq!(
        transaction
            .append_font(Scope::Base, super::Font::new("Boundary 128"))
            .unwrap(),
        128
    );
    let commit = transaction.commit().unwrap();
    assert_eq!(commit.fonts().base.as_ref().unwrap().fonts.len(), 129);

    let mut refused = commit.snapshot().edit().unwrap();
    let before = refused.fonts().clone();
    assert!(
        refused
            .append_font(Scope::Base, super::Font::new("Boundary 129"))
            .is_err()
    );
    assert_eq!(refused.fonts(), &before);
    assert!(refused.changes().is_empty());
}

#[test]
fn inserts_pp10_font_flags_after_comments_and_before_later_or_opaque_tail() {
    let base = collection(RecordType::FontCollection, entity(0, "Base", 0));
    let mut ppt10 = record_bytes(
        0x0f,
        0,
        RecordType::FontCollection10.as_u16(),
        &entity(0, "Intl", 0),
    );
    ppt10.extend_from_slice(&record_bytes(
        0,
        0,
        RecordType::GridSpacing10Atom.as_u16(),
        &[0; 8],
    ));
    ppt10.extend_from_slice(&record_bytes(
        0x0f,
        0,
        RecordType::CommentIndex10.as_u16(),
        &[],
    ));
    ppt10.extend_from_slice(&record_bytes(
        0x0f,
        0,
        RecordType::OutlineTextProps10.as_u16(),
        &[],
    ));
    ppt10.extend_from_slice(&record_bytes(0, 0, 0x7abc, &[9, 8, 7]));
    let mut root = document(base, &ppt10);
    let before = FontCollections::parse(&root).unwrap();
    let mut after = before.clone();
    after.embedding_flags = Some(FontEmbeddingFlags::new(true, false));
    after
        .apply_to_document_from(&before, &mut root, super::Limits::default())
        .unwrap();

    let records = root.versioned_binary_tag_records(10).unwrap();
    let position = |kind| {
        records
            .iter()
            .position(|record| record.record_type == kind)
            .unwrap()
    };
    assert!(position(RecordType::CommentIndex10) < position(RecordType::FontEmbedFlags10Atom));
    assert!(position(RecordType::FontEmbedFlags10Atom) < position(RecordType::OutlineTextProps10));
    let opaque = records
        .iter()
        .position(|record| record.record_type_raw == 0x7abc)
        .unwrap();
    assert!(position(RecordType::FontEmbedFlags10Atom) < opaque);
}
