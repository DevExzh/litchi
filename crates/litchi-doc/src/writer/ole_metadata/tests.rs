//! Focused tests for the DOC OLE metadata model and codecs.

use super::codec::{COMP_OBJ_LEN, OLE_LEN, write_comp_obj, write_ole};
use super::model::{ClassId, CompObj, Metadata, Ole};
use super::validation;
use super::{generate_compobj_stream, generate_ole_stream};

#[test]
fn default_metadata_is_the_word_document_profile() {
    let metadata = Metadata::default();
    assert_eq!(metadata, Metadata::word_document());
    assert_eq!(metadata.comp_obj().class_id(), ClassId::WORD_DOCUMENT);
    assert_eq!(metadata.comp_obj().user_type(), "Microsoft Word Document");
    assert_eq!(metadata.comp_obj().clipboard_format(), "MSWordDoc");
    assert_eq!(metadata.comp_obj().prog_id(), "Word.Document.8");
    assert_eq!(metadata.ole().version(), 0x0200_0001);
}

#[test]
fn class_id_round_trips_without_allocation() {
    let bytes = *ClassId::WORD_DOCUMENT.as_bytes();
    assert_eq!(ClassId::from_bytes(bytes), ClassId::WORD_DOCUMENT);
}

#[test]
fn comp_obj_codec_preserves_the_existing_wire_profile() {
    let data = write_comp_obj(CompObj::word_document());
    assert_eq!(data.len(), COMP_OBJ_LEN);
    assert_eq!(&data[..4], &[0x01, 0x00, 0xFE, 0xFF]);
    assert_eq!(&data[12..28], ClassId::WORD_DOCUMENT.as_bytes());
    assert_eq!(&data[28..32], &24u32.to_le_bytes());
    assert_eq!(&data[32..56], b"Microsoft Word Document\0");
    assert_eq!(&data[56..60], &10u32.to_le_bytes());
    assert_eq!(&data[60..70], b"MSWordDoc\0");
    assert_eq!(&data[70..74], &16u32.to_le_bytes());
    assert_eq!(&data[74..90], b"Word.Document.8\0");
    assert!(validation::comp_obj(&data, CompObj::word_document()).is_ok());
}

#[test]
fn ole_codec_preserves_the_existing_wire_profile() {
    let data = write_ole(Ole::word_document());
    assert_eq!(data.len(), OLE_LEN);
    assert_eq!(&data[..4], &[0x01, 0x00, 0x00, 0x02]);
    assert!(data[4..].iter().all(|&byte| byte == 0));
    assert!(validation::ole(&data, Ole::word_document()).is_ok());
}

#[test]
fn public_generators_match_the_typed_codecs() {
    assert_eq!(
        generate_compobj_stream(),
        write_comp_obj(CompObj::word_document())
    );
    assert_eq!(generate_ole_stream(), write_ole(Ole::word_document()));
}

#[test]
fn validation_rejects_truncation_and_reserved_byte_changes() {
    let mut comp_obj = generate_compobj_stream();
    comp_obj.pop();
    assert!(validation::comp_obj(&comp_obj, CompObj::word_document()).is_err());

    let mut ole = generate_ole_stream();
    ole[4] = 1;
    assert!(validation::ole(&ole, Ole::word_document()).is_err());
}
