#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::codec::{
    ATOM_VERSION, BROADCAST_DOC_INFO_9, COMMENT_10, CONTAINER_VERSION, COPYRIGHT_INSTANCE,
    ENVELOPE_DATA_9_ATOM, ENVELOPE_FLAGS_9_ATOM, HTML_DOC_INFO_9_ATOM, HTML_PUBLISH_INFO_9,
    KEYWORDS_INSTANCE, MODIFY_PASSWORD_INSTANCE, PRES_ADVISOR_FLAGS_9_ATOM,
};
use super::*;
use crate::consts::RecordType;
use crate::prog_tags::{ProgBinaryTagVersion, ProgTagLimits, ProgTagScope, ProgTags};
use crate::records::Record;

fn record_bytes(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(8 + payload.len());
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&kind.to_le_bytes());
    bytes.extend_from_slice(&u32::try_from(payload.len()).unwrap().to_le_bytes());
    bytes.extend_from_slice(payload);
    bytes
}

fn parse_payload(bytes: &[u8]) -> Vec<Record> {
    Record::parse_sequence_strict(bytes, "test payload").unwrap()
}

fn atom(instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
    record_bytes(ATOM_VERSION, instance, kind, payload)
}

fn container(kind: u16, payload: &[u8]) -> Vec<u8> {
    record_bytes(CONTAINER_VERSION, 0, kind, payload)
}

fn cstring(instance: u16, text: &str) -> Vec<u8> {
    let data: Vec<u8> = text.encode_utf16().flat_map(u16::to_le_bytes).collect();
    atom(instance, RecordType::CString.as_u16(), &data)
}

/// Encode the full PP9 document grammar, in spec order.
fn pp9_doc_payload() -> Vec<u8> {
    [
        atom(1, RecordType::TextMasterStyle9Atom.as_u16(), &[0; 20]),
        atom(2, RecordType::TextMasterStyle9Atom.as_u16(), &[0; 20]),
        container(RecordType::BlipCollection9.as_u16(), &[]),
        atom(0, RecordType::TextDefaults9Atom.as_u16(), &[0; 8]),
        container(RecordType::Kinsoku.as_u16(), &[]),
        container(RecordType::ExternalHyperlink9.as_u16(), &[]),
        atom(0, PRES_ADVISOR_FLAGS_9_ATOM, &[0; 4]),
        atom(0, ENVELOPE_DATA_9_ATOM, &[1, 2, 3]),
        atom(0, ENVELOPE_FLAGS_9_ATOM, &[0; 4]),
        atom(0, HTML_DOC_INFO_9_ATOM, &[0; 16]),
        container(HTML_PUBLISH_INFO_9, &[]),
        container(BROADCAST_DOC_INFO_9, &[]),
        container(BROADCAST_DOC_INFO_9, &[]),
        container(RecordType::OutlineTextProps9.as_u16(), &[]),
    ]
    .concat()
}

/// Encode the full PP10 document grammar, in spec order.
fn pp10_doc_payload() -> Vec<u8> {
    [
        container(RecordType::FontCollection10.as_u16(), &[]),
        atom(0, RecordType::TextMasterStyle10Atom.as_u16(), &[0; 12]),
        atom(0, RecordType::TextDefaults10Atom.as_u16(), &[0; 8]),
        atom(0, RecordType::GridSpacing10Atom.as_u16(), &[0; 8]),
        container(RecordType::CommentIndex10.as_u16(), &[]),
        atom(0, RecordType::FontEmbedFlags10Atom.as_u16(), &[0; 4]),
        cstring(COPYRIGHT_INSTANCE, "(c) Ada"),
        cstring(KEYWORDS_INSTANCE, "ppt,test"),
        atom(0, RecordType::FilterPrivacyFlags10Atom.as_u16(), &[0; 4]),
        container(RecordType::OutlineTextProps10.as_u16(), &[]),
        atom(0, RecordType::DocToolbarStates10Atom.as_u16(), &[0]),
        container(RecordType::SlideListTable10.as_u16(), &[]),
        container(RecordType::DiffTree10.as_u16(), &[]),
        cstring(MODIFY_PASSWORD_INSTANCE, "secret"),
        atom(0, RecordType::PhotoAlbumInfo10Atom.as_u16(), &[0; 6]),
    ]
    .concat()
}

/// Encode the full PP10 slide grammar, in spec order.
fn pp10_slide_payload() -> Vec<u8> {
    let mut linked_slide = Vec::new();
    linked_slide.extend_from_slice(&42u32.to_le_bytes());
    linked_slide.extend_from_slice(&2i32.to_le_bytes());
    [
        atom(0, RecordType::TextMasterStyle10Atom.as_u16(), &[0; 12]),
        container(COMMENT_10, &[]),
        atom(0, RecordType::LinkedSlide10Atom.as_u16(), &linked_slide),
        atom(0, RecordType::LinkedShape10Atom.as_u16(), &[0; 8]),
        atom(0, RecordType::LinkedShape10Atom.as_u16(), &[0; 8]),
        atom(0, RecordType::SlideFlags10Atom.as_u16(), &[0; 4]),
        atom(0, RecordType::SlideTime10Atom.as_u16(), &[0; 8]),
        atom(0, RecordType::HashCode10Atom.as_u16(), &[0; 4]),
        container(RecordType::ExtTimeNode.as_u16(), &[]),
        container(RecordType::BuildList.as_u16(), &[]),
    ]
    .concat()
}

/// Wrap an extension payload in a `ProgBinaryTag`/`ProgTags` record pair.
fn prog_tags_record(tag_name: &str, extension_payload: &[u8]) -> (Vec<u8>, Record) {
    let name: Vec<u8> = tag_name.encode_utf16().flat_map(u16::to_le_bytes).collect();
    let binary_tag = record_bytes(
        CONTAINER_VERSION,
        0,
        RecordType::ProgBinaryTag.as_u16(),
        &[
            atom(0, RecordType::CString.as_u16(), &name),
            atom(0, RecordType::BinaryTagData.as_u16(), extension_payload),
        ]
        .concat(),
    );
    let bytes = record_bytes(
        CONTAINER_VERSION,
        0,
        RecordType::ProgTags.as_u16(),
        &binary_tag,
    );
    let (record, consumed) = Record::parse_strict(&bytes, 0).unwrap();
    assert_eq!(consumed, bytes.len());
    (bytes, record)
}

#[test]
fn pp9_doc_extension_assigns_every_slot_and_round_trips_exactly() {
    let payload = pp9_doc_payload();
    let extension = DocBinaryTagExtension9::parse_records(parse_payload(&payload)).unwrap();

    assert_eq!(extension.text_master_styles.len(), 2);
    assert!(extension.blip_collection.is_some());
    assert!(extension.text_defaults.is_some());
    assert!(extension.kinsoku.is_some());
    assert_eq!(extension.external_hyperlinks.len(), 1);
    assert!(extension.advisor_flags.is_some());
    assert!(extension.envelope_data.is_some());
    assert!(extension.envelope_flags.is_some());
    assert!(extension.html_doc_info.is_some());
    assert!(extension.html_publish_info.is_some());
    assert_eq!(extension.broadcasts.len(), 2);
    assert!(extension.outline_text_props.is_some());
    assert_eq!(extension.to_payload().unwrap(), payload);
}

#[test]
fn pp10_doc_extension_assigns_every_slot_and_round_trips_exactly() {
    let payload = pp10_doc_payload();
    let extension = DocBinaryTagExtension10::parse_records(parse_payload(&payload)).unwrap();

    assert!(extension.font_collection.is_some());
    assert_eq!(extension.text_master_styles.len(), 1);
    assert!(extension.text_defaults.is_some());
    assert!(extension.grid_spacing.is_some());
    assert_eq!(extension.comment_indices.len(), 1);
    assert!(extension.font_embed_flags.is_some());
    assert!(extension.copyright.is_some());
    assert!(extension.keywords.is_some());
    assert!(extension.filter_privacy_flags.is_some());
    assert!(extension.outline_text_props.is_some());
    assert!(extension.toolbar_states.is_some());
    assert!(extension.slide_list_table.is_some());
    assert_eq!(extension.diff_trees.len(), 1);
    assert!(extension.modify_password.is_some());
    assert!(extension.photo_album_info.is_some());
    assert_eq!(extension.to_payload().unwrap(), payload);
}

#[test]
fn pp10_doc_extension_allows_minimal_grammar() {
    // Only the optional GridSpacing10Atom.
    let payload = atom(0, RecordType::GridSpacing10Atom.as_u16(), &[0; 8]);
    let extension = DocBinaryTagExtension10::parse_records(parse_payload(&payload)).unwrap();
    assert!(extension.grid_spacing.is_some());
    assert!(extension.font_collection.is_none());
    assert!(extension.modify_password.is_none());
    assert_eq!(extension.to_payload().unwrap(), payload);
}

#[test]
fn pp11_and_pp12_doc_extensions_round_trip() {
    let pp11_payload = [
        container(RecordType::SmartTagStore11.as_u16(), &[]),
        container(RecordType::OutlineTextProps11.as_u16(), &[]),
    ]
    .concat();
    let pp11 = DocBinaryTagExtension11::parse_records(parse_payload(&pp11_payload)).unwrap();
    assert!(pp11.smart_tag_store.is_some());
    assert!(pp11.outline_text_props.is_some());
    assert_eq!(pp11.to_payload().unwrap(), pp11_payload);

    let pp12_payload = atom(0, RecordType::RoundTripDocFlags12Atom.as_u16(), &[0]);
    let pp12 = DocBinaryTagExtension12::parse_records(parse_payload(&pp12_payload)).unwrap();
    assert!(pp12.doc_flags.is_some());
    assert_eq!(pp12.to_payload().unwrap(), pp12_payload);

    // PP12 with all-optional grammar accepts an empty payload.
    let empty = DocBinaryTagExtension12::parse_records(Vec::new()).unwrap();
    assert!(empty.doc_flags.is_none());
    assert_eq!(empty.to_payload().unwrap(), Vec::<u8>::new());
}

#[test]
fn pp9_and_pp12_slide_extensions_round_trip() {
    let pp9_payload = [
        atom(0, RecordType::TextMasterStyle9Atom.as_u16(), &[0; 20]),
        atom(3, RecordType::TextMasterStyle9Atom.as_u16(), &[0; 20]),
    ]
    .concat();
    let pp9 = SlideBinaryTagExtension9::parse_records(parse_payload(&pp9_payload)).unwrap();
    assert_eq!(pp9.text_master_styles.len(), 2);
    assert_eq!(pp9.to_payload().unwrap(), pp9_payload);

    let pp12_payload = atom(
        0,
        RecordType::RoundTripHeaderFooterDefaults12Atom.as_u16(),
        &[0],
    );
    let pp12 = SlideBinaryTagExtension12::parse_records(parse_payload(&pp12_payload)).unwrap();
    assert!(pp12.header_footer_defaults.is_some());
    assert_eq!(pp12.to_payload().unwrap(), pp12_payload);
}

#[test]
fn pp10_slide_extension_assigns_slots_and_round_trips_exactly() {
    let payload = pp10_slide_payload();
    let extension = SlideBinaryTagExtension10::parse_records(parse_payload(&payload)).unwrap();

    assert_eq!(extension.text_master_styles.len(), 1);
    assert_eq!(extension.comments.len(), 1);
    assert!(extension.linked_slide.is_some());
    assert_eq!(extension.linked_shapes.len(), 2);
    assert!(extension.slide_flags.is_some());
    assert!(extension.slide_time.is_some());
    assert!(extension.hash_code.is_some());
    assert!(extension.timing.is_some());
    assert!(extension.build_list.is_some());
    assert_eq!(extension.to_payload().unwrap(), payload);
}

#[test]
fn pp10_slide_extension_validates_linked_shape_count() {
    // Count says 1, array holds 2.
    let mut linked_slide = Vec::new();
    linked_slide.extend_from_slice(&42u32.to_le_bytes());
    linked_slide.extend_from_slice(&1i32.to_le_bytes());
    let mismatched = [
        atom(0, RecordType::LinkedSlide10Atom.as_u16(), &linked_slide),
        atom(0, RecordType::LinkedShape10Atom.as_u16(), &[0; 8]),
        atom(0, RecordType::LinkedShape10Atom.as_u16(), &[0; 8]),
    ]
    .concat();
    assert!(SlideBinaryTagExtension10::parse_records(parse_payload(&mismatched)).is_err());

    // Shapes without the linked-slide atom.
    let orphan = atom(0, RecordType::LinkedShape10Atom.as_u16(), &[0; 8]);
    assert!(SlideBinaryTagExtension10::parse_records(parse_payload(&orphan)).is_err());

    // Truncated linked-slide atom.
    let truncated = atom(0, RecordType::LinkedSlide10Atom.as_u16(), &[0; 4]);
    assert!(SlideBinaryTagExtension10::parse_records(parse_payload(&truncated)).is_err());
}

#[test]
fn grammars_accept_optional_grid_and_reject_out_of_order_or_trailing_records() {
    // PP9 doc: OutlineTextProps9 before the broadcast array is out of order.
    let out_of_order = [
        container(RecordType::OutlineTextProps9.as_u16(), &[]),
        container(BROADCAST_DOC_INFO_9, &[]),
    ]
    .concat();
    assert!(DocBinaryTagExtension9::parse_records(parse_payload(&out_of_order)).is_err());

    // PP10 doc: GridSpacing10Atom is genuinely optional.
    let omitted_grid = atom(0, RecordType::TextMasterStyle10Atom.as_u16(), &[0; 12]);
    let without_grid =
        DocBinaryTagExtension10::parse_records(parse_payload(&omitted_grid)).unwrap();
    assert!(without_grid.grid_spacing.is_none());

    // PP10 doc: GridSpacing10Atom after the comment array is out of order.
    let pp10_out_of_order = [
        container(RecordType::CommentIndex10.as_u16(), &[]),
        atom(0, RecordType::GridSpacing10Atom.as_u16(), &[0; 8]),
    ]
    .concat();
    assert!(DocBinaryTagExtension10::parse_records(parse_payload(&pp10_out_of_order)).is_err());

    // PP10 doc: a ModifyPasswordAtom where the CopyrightAtom belongs.
    let wrong_instance = [
        atom(0, RecordType::GridSpacing10Atom.as_u16(), &[0; 8]),
        cstring(MODIFY_PASSWORD_INSTANCE, "secret"),
    ]
    .concat();
    assert!(DocBinaryTagExtension10::parse_records(parse_payload(&wrong_instance)).is_err());

    // PP9 slide: any non-TextMasterStyle9Atom record is outside the grammar.
    let foreign = atom(0, RecordType::TextDefaults9Atom.as_u16(), &[0; 8]);
    assert!(SlideBinaryTagExtension9::parse_records(parse_payload(&foreign)).is_err());

    // Wrong version nibble on an array element.
    let bad_version = record_bytes(
        CONTAINER_VERSION,
        0,
        RecordType::TextMasterStyle9Atom.as_u16(),
        &[0; 20],
    );
    assert!(SlideBinaryTagExtension9::parse_records(parse_payload(&bad_version)).is_err());
}

#[test]
fn tag_and_container_level_dispatch_decode_extensions() {
    let limits = ProgTagLimits::default();
    let (bytes, record) = prog_tags_record("___PPT9", &pp9_doc_payload());
    let tags = ProgTags::parse(&record, ProgTagScope::Document, limits).unwrap();

    let extensions = tags.document_extensions().unwrap();
    let pp9 = extensions.powerpoint9.as_ref().unwrap();
    assert_eq!(pp9.text_master_styles.len(), 2);
    assert_eq!(pp9.broadcasts.len(), 2);
    assert!(extensions.powerpoint10.is_none());

    let tag = tags.binary_tag(ProgBinaryTagVersion::PowerPoint9).unwrap();
    let extension = tag.doc_extension().unwrap().unwrap();
    assert_eq!(extension.to_payload().unwrap(), pp9_doc_payload());
    // Container-level bytes are unaffected by extension decoding.
    assert_eq!(tags.to_bytes(limits).unwrap(), bytes);

    // Document tags cannot decode slide extensions and vice versa.
    assert!(tags.slide_extensions().is_err());
    // The doc-scoped ___PPT9 payload is not a valid PP9 slide grammar.
    assert!(tag.slide_extension().is_err());
}

#[test]
fn slide_scope_dispatch_decodes_pp10_slide_extension() {
    let limits = ProgTagLimits::default();
    let (_, record) = prog_tags_record("___PPT10", &pp10_slide_payload());
    let tags = ProgTags::parse(&record, ProgTagScope::Slide, limits).unwrap();

    let extensions = tags.slide_extensions().unwrap();
    let pp10 = extensions.powerpoint10.as_ref().unwrap();
    assert_eq!(pp10.linked_shapes.len(), 2);
    assert!(pp10.build_list.is_some());
    assert!(extensions.powerpoint9.is_none());
    assert!(tags.document_extensions().is_err());

    let tag = tags.binary_tag(ProgBinaryTagVersion::PowerPoint10).unwrap();
    assert_eq!(
        tag.slide_extension()
            .unwrap()
            .unwrap()
            .to_payload()
            .unwrap(),
        pp10_slide_payload()
    );
}

#[test]
fn ppt11_versioned_tag_rejects_slide_extension_decode() {
    let limits = ProgTagLimits::default();
    let payload = container(RecordType::SmartTagStore11.as_u16(), &[]);
    let (_, record) = prog_tags_record("___PPT11", &payload);
    let tags = ProgTags::parse(&record, ProgTagScope::Document, limits).unwrap();
    let tag = tags.binary_tag(ProgBinaryTagVersion::PowerPoint11).unwrap();
    assert!(tag.slide_extension().is_err());
    assert!(tag.doc_extension().unwrap().is_some());
}
