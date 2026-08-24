//! Selector-first integration coverage for Pages body-table dimensions.

use std::error::Error as StdError;

use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_common::{decode_varint_from_bytes, wire::WireView};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldPath, RawMessage, SnappyStream};
use litchi_iwa_protos::{tp, tsa, tsd, tsp, tst, tswp};
use litchi_pages::table::dimension::{Dimension, Points, Size};
use litchi_pages::{BodyTableDimensionError as Error, BodyTableSelector, Limits, Package};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const ROW_BUCKET_MEMBER: &str = "Index/RowHeaderBuckets.iwa";
const ROOT_IDENTIFIER: u64 = 1;
const BODY_IDENTIFIER: u64 = 2;
const FIRST_ATTACHMENT_IDENTIFIER: u64 = 10;
const FIRST_DRAWABLE_IDENTIFIER: u64 = 11;
const FIRST_MODEL_IDENTIFIER: u64 = 12;
const ROW_BUCKET_IDENTIFIER: u64 = 100;
const COLUMN_BUCKET_IDENTIFIER: u64 = 101;
const DATA_OBJECT_START: u64 = 200;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
const ROOT_MESSAGE_TYPE: u32 = 10_000;
const BODY_MESSAGE_TYPE: u32 = 2_001;
const ATTACHMENT_MESSAGE_TYPE: u32 = 2_003;
const PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
const UNKNOWN_BUCKET_FIELD: u32 = 99;
const UNKNOWN_BUCKET_VALUE: u64 = 0xfeed_beef;
const UNKNOWN_HEADER_FIELD: u32 = 98;
const UNKNOWN_HEADER_VALUE: u64 = 0xcafe_babe;

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

trait ExactBytes {
    fn exact_bytes(&self) -> Vec<u8>;
}

impl ExactBytes for Package {
    fn exact_bytes(&self) -> Vec<u8> {
        let mut bytes = Vec::new();
        self.write_to(&mut bytes)
            .expect("an in-memory Vec accepts package bytes");
        bytes
    }
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn field_reference(path: impl Into<FieldPath>, identifier: u64) -> FieldInfo {
    let mut field = FieldInfo::new(path);
    field.object_references.push(identifier);
    field
}

fn object(
    identifier: u64,
    type_: u32,
    data: Vec<u8>,
    references: &[u64],
) -> TestResult<ArchiveObject> {
    let mut object = ArchiveObject::new(identifier, vec![RawMessage { type_, data }])?;
    object.archive_info.message_infos[0]
        .object_references
        .extend_from_slice(references);
    Ok(object)
}

fn storage_reference_ids() -> [u64; 7] {
    [
        DATA_OBJECT_START,
        DATA_OBJECT_START + 1,
        DATA_OBJECT_START + 2,
        DATA_OBJECT_START + 3,
        DATA_OBJECT_START + 4,
        DATA_OBJECT_START + 5,
        DATA_OBJECT_START + 6,
    ]
}

fn data_store() -> tst::DataStore {
    let ids = storage_reference_ids();
    tst::DataStore {
        row_headers: tst::HeaderStorage {
            bucket_hash_function: 1,
            buckets: vec![reference(ROW_BUCKET_IDENTIFIER)],
        },
        column_headers: reference(COLUMN_BUCKET_IDENTIFIER),
        tiles: tst::TileStorage::default(),
        string_table: reference(ids[0]),
        style_table: reference(ids[1]),
        formula_table: reference(ids[2]),
        format_table_pre_bnc: reference(ids[3]),
        next_row_strip_id: 1,
        next_column_strip_id: 1,
        row_tile_tree: tst::TableRbTree::default(),
        column_tile_tree: tst::TableRbTree::default(),
        ..tst::DataStore::default()
    }
}

fn model_payload(name: &str) -> Vec<u8> {
    let store = data_store();
    let mut model = tst::TableModelArchive {
        table_id: format!("table-{name}"),
        table_style: reference(DATA_OBJECT_START + 10),
        body_text_style: reference(DATA_OBJECT_START + 11),
        header_row_text_style: reference(DATA_OBJECT_START + 12),
        header_column_text_style: reference(DATA_OBJECT_START + 13),
        footer_row_text_style: reference(DATA_OBJECT_START + 14),
        body_cell_style: reference(DATA_OBJECT_START + 15),
        header_row_style: reference(DATA_OBJECT_START + 16),
        header_column_style: reference(DATA_OBJECT_START + 17),
        footer_row_style: reference(DATA_OBJECT_START + 18),
        base_data_store: store,
        number_of_rows: 3,
        number_of_columns: 3,
        table_name: name.to_owned(),
        default_row_height: 18.0,
        default_column_width: 64.0,
        ..tst::TableModelArchive::default()
    };
    let mut payload = model.encode_to_vec();
    litchi_iwa_common::wire::append_varint_field(&mut payload, UNKNOWN_BUCKET_FIELD, 7)
        .expect("unknown model field fits");
    model = tst::TableModelArchive::decode(payload.as_slice()).expect("model remains decodable");
    let mut payload = model.encode_to_vec();
    litchi_iwa_common::wire::append_varint_field(&mut payload, UNKNOWN_BUCKET_FIELD, 7)
        .expect("unknown model field fits");
    payload
}

fn table_info_payload() -> Vec<u8> {
    tst::TableInfoArchive {
        super_: tsd::DrawableArchive {
            parent: Some(reference(BODY_IDENTIFIER)),
            ..tsd::DrawableArchive::default()
        },
        table_model: reference(FIRST_MODEL_IDENTIFIER),
        ..tst::TableInfoArchive::default()
    }
    .encode_to_vec()
}

fn bucket_header(index: u32, size: f32) -> Vec<u8> {
    tst::header_storage_bucket::Header {
        index,
        size,
        hiding_state: 0,
        number_of_cells: 0,
        cell_style: None,
        text_style: None,
    }
    .encode_to_vec()
}

fn bucket_payload(headers: &[(u32, f32)], unknown_header: bool) -> TestResult<Vec<u8>> {
    let mut raw_headers = Vec::new();
    for &(index, size) in headers {
        let mut raw = bucket_header(index, size);
        if unknown_header {
            litchi_iwa_common::wire::append_varint_field(
                &mut raw,
                UNKNOWN_HEADER_FIELD,
                UNKNOWN_HEADER_VALUE,
            )?;
        }
        raw_headers.push(raw);
    }
    let mut payload = Vec::new();
    litchi_iwa_common::wire::append_varint_field(&mut payload, 1, 1)?;
    for raw in raw_headers {
        litchi_iwa_common::wire::append_length_delimited_field(&mut payload, 2, &raw)?;
    }
    litchi_iwa_common::wire::append_varint_field(
        &mut payload,
        UNKNOWN_BUCKET_FIELD,
        UNKNOWN_BUCKET_VALUE,
    )?;
    Ok(payload)
}

fn synthetic_package() -> TestResult<Vec<u8>> {
    let root = tp::DocumentArchive {
        super_: tsa::DocumentArchive::default(),
        body_storage: Some(reference(BODY_IDENTIFIER)),
        ..tp::DocumentArchive::default()
    };
    let body = tswp::StorageArchive {
        kind: Some(tswp::storage_archive::KindType::Body as i32),
        text: vec!["\u{fffc}".to_owned()],
        table_attachment: Some(tswp::ObjectAttributeTable {
            entries: vec![tswp::object_attribute_table::ObjectAttribute {
                character_index: 0,
                object: Some(reference(FIRST_ATTACHMENT_IDENTIFIER)),
            }],
        }),
        ..tswp::StorageArchive::default()
    };
    let attachment = tswp::DrawableAttachmentArchive {
        drawable: Some(reference(FIRST_DRAWABLE_IDENTIFIER)),
        ..tswp::DrawableAttachmentArchive::default()
    };
    let mut root_object = object(
        ROOT_IDENTIFIER,
        ROOT_MESSAGE_TYPE,
        root.encode_to_vec(),
        &[BODY_IDENTIFIER],
    )?;
    root_object.archive_info.message_infos[0]
        .field_infos
        .push(field_reference(vec![4], BODY_IDENTIFIER));
    let mut body_object = object(
        BODY_IDENTIFIER,
        BODY_MESSAGE_TYPE,
        body.encode_to_vec(),
        &[FIRST_ATTACHMENT_IDENTIFIER],
    )?;
    body_object.archive_info.message_infos[0]
        .field_infos
        .push(field_reference(vec![9], FIRST_ATTACHMENT_IDENTIFIER));
    let mut attachment_object = object(
        FIRST_ATTACHMENT_IDENTIFIER,
        ATTACHMENT_MESSAGE_TYPE,
        attachment.encode_to_vec(),
        &[FIRST_DRAWABLE_IDENTIFIER],
    )?;
    attachment_object.archive_info.message_infos[0]
        .field_infos
        .push(field_reference(vec![1], FIRST_DRAWABLE_IDENTIFIER));
    let mut info_object = object(
        FIRST_DRAWABLE_IDENTIFIER,
        TABLE_INFO_MESSAGE_TYPE,
        table_info_payload(),
        &[BODY_IDENTIFIER, FIRST_MODEL_IDENTIFIER],
    )?;
    info_object.archive_info.message_infos[0]
        .field_infos
        .extend([
            field_reference(vec![1, 2], BODY_IDENTIFIER),
            field_reference(vec![2], FIRST_MODEL_IDENTIFIER),
        ]);
    let model_refs = [
        ROW_BUCKET_IDENTIFIER,
        COLUMN_BUCKET_IDENTIFIER,
        FIRST_MODEL_IDENTIFIER + 188,
    ];
    let mut model_object = object(
        FIRST_MODEL_IDENTIFIER,
        TABLE_MODEL_MESSAGE_TYPE,
        model_payload("Revenue"),
        &model_refs,
    )?;
    model_object.archive_info.message_infos[0]
        .field_infos
        .extend([
            field_reference(vec![4, 1, 2], ROW_BUCKET_IDENTIFIER),
            field_reference(vec![4, 2], COLUMN_BUCKET_IDENTIFIER),
        ]);
    let row_bucket = object(
        ROW_BUCKET_IDENTIFIER,
        HEADER_BUCKET_MESSAGE_TYPE,
        bucket_payload(&[(1, 25.0)], true)?,
        &[],
    )?;
    let column_bucket = object(
        COLUMN_BUCKET_IDENTIFIER,
        HEADER_BUCKET_MESSAGE_TYPE,
        bucket_payload(&[(2, 40.0)], true)?,
        &[],
    )?;
    let mut objects = vec![
        root_object,
        body_object,
        attachment_object,
        info_object,
        model_object,
        row_bucket,
        column_bucket,
    ];
    for (offset, identifier) in storage_reference_ids().into_iter().enumerate() {
        objects.push(object(
            identifier,
            7_000 + u32::try_from(offset)?,
            vec![0x08, 0x01],
            &[],
        )?);
    }
    for identifier in (DATA_OBJECT_START + 10)..=(DATA_OBJECT_START + 18) {
        objects.push(object(identifier, 2_022, vec![0x08, 0x01], &[])?);
    }
    let archive = Archive { objects };
    let component = SnappyStream::compress(&archive.to_bytes()?)?;
    let mut members: Vec<(&str, &[u8])> = vec![
        ("Data/sentinel.bin", b"untouched-sentinel"),
        (DOCUMENT_MEMBER, component.as_slice()),
    ];
    members.extend(
        PREVIEWS
            .into_iter()
            .map(|name| (name, b"preview".as_slice())),
    );
    Ok(litchi_iwa_archive::package::to_bytes(
        members,
        Limits::default(),
    )?)
}

fn document_archive(package: &[u8]) -> TestResult<Archive> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or("missing document member")?;
    Ok(Archive::parse(
        SnappyStream::decompress(entry.data())?.as_bytes(),
    )?)
}

fn member_bytes(package: &[u8], name: &str) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    Ok(catalog
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or_else(|| format!("missing package member {name}"))?
        .data()
        .to_vec())
}

fn split_row_bucket_component(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let document_entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or("missing document member")?;
    let mut document = Archive::parse(SnappyStream::decompress(document_entry.data())?.as_bytes())?;
    let bucket_index = document
        .objects
        .iter()
        .position(|object| object.archive_info.identifier == Some(ROW_BUCKET_IDENTIFIER))
        .ok_or("missing row bucket")?;
    let bucket = document.objects.remove(bucket_index);
    let document_payload = SnappyStream::compress(&document.to_bytes()?)?;
    let bucket_payload = SnappyStream::compress(
        &Archive {
            objects: vec![bucket],
        }
        .to_bytes()?,
    )?;
    let mut entries = Vec::new();
    for entry in catalog.iter() {
        let data = if entry.name() == DOCUMENT_MEMBER {
            document_payload.as_slice()
        } else {
            entry.data()
        };
        entries.push((entry.name().to_owned(), data.to_vec()));
    }
    entries.push((ROW_BUCKET_MEMBER.to_owned(), bucket_payload));
    let borrowed = entries
        .iter()
        .map(|(name, data)| (name.as_str(), data.as_slice()))
        .collect::<Vec<_>>();
    Ok(litchi_iwa_archive::package::to_bytes(
        borrowed,
        Limits::default(),
    )?)
}

fn without_previews(package: &[u8]) -> TestResult<Vec<u8>> {
    Ok(
        Catalog::from_bytes(package)?.reassemble_with_deletions_to_bytes(
            &[],
            &PREVIEWS,
            Limits::default(),
        )?,
    )
}

fn bucket_bytes(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let archive = document_archive(package)?;
    Ok(archive
        .object(identifier)
        .ok_or("missing bucket")?
        .messages
        .iter()
        .find(|message| message.type_ == HEADER_BUCKET_MESSAGE_TYPE)
        .ok_or("missing bucket payload")?
        .data
        .clone())
}

fn rewrite_document_archive(
    package: &[u8],
    mutate: impl FnOnce(&mut Archive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or("missing document member")?;
    let mut archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
    mutate(&mut archive)?;
    let component = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(DOCUMENT_MEMBER, &component)],
        Limits::default(),
    )?)
}

fn append_duplicate_header(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let object = archive.object_mut(identifier).ok_or("missing bucket")?;
        let message = object
            .messages
            .iter_mut()
            .find(|message| message.type_ == HEADER_BUCKET_MESSAGE_TYPE)
            .ok_or("missing bucket message")?;
        let headers = WireView::parse(&message.data)?
            .fields()
            .filter(|field| field.number() == 2)
            .map(|field| field.payload().to_vec())
            .next()
            .ok_or("missing header")?;
        litchi_iwa_common::wire::append_length_delimited_field(&mut message.data, 2, &headers)?;
        Ok(())
    })
}

fn replace_bucket_message_type(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let object = archive.object_mut(identifier).ok_or("missing bucket")?;
        let message = object
            .messages
            .iter_mut()
            .find(|message| message.type_ == HEADER_BUCKET_MESSAGE_TYPE)
            .ok_or("missing bucket message")?;
        message.type_ = HEADER_BUCKET_MESSAGE_TYPE + 1;
        object.archive_info.message_infos[0].type_ = HEADER_BUCKET_MESSAGE_TYPE + 1;
        Ok(())
    })
}

fn remove_bucket(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        archive
            .objects
            .retain(|object| object.archive_info.identifier != Some(identifier));
        Ok(())
    })
}

fn sentinel(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    Ok(catalog
        .iter()
        .find(|entry| entry.name() == "Data/sentinel.bin")
        .ok_or("missing sentinel")?
        .data()
        .to_vec())
}

fn points(value: f32) -> TestResult<Size> {
    Ok(Size::points(value)?)
}

#[test]
fn selectors_read_rows_columns_and_default_presence() -> TestResult {
    let source = synthetic_package()?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.body_table_dimension_size(0usize, Dimension::Row(0))?,
        Size::Default
    );
    assert_eq!(
        package.body_table_dimension_size(BodyTableSelector::name("Revenue"), Dimension::Row(1))?,
        points(25.0)?
    );
    assert_eq!(
        package.body_table_dimension_size(0usize, Dimension::Column(2))?,
        points(40.0)?
    );
    assert!(
        package
            .body_table_dimension_size(BodyTableSelector::name("Missing"), Dimension::Row(0))
            .is_err()
    );
    assert!(
        package
            .body_table_dimension_size(0usize, Dimension::Row(3))
            .is_err()
    );
    Ok(())
}

#[test]
fn no_op_is_exact_and_reset_restores_default() -> TestResult {
    let source = synthetic_package()?;
    let package = Package::from_bytes(&source)?;
    let before = package.body_table_dimension_size(0usize, Dimension::Row(1))?;
    let commit = package
        .edit_body_table_dimension_size(BodyTableSelector::name("Revenue"), Dimension::Row(1))?
        .set(before)
        .commit()?;
    assert!(commit.patch().is_noop());
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert_eq!(commit.package().exact_bytes(), source);
    let changed = package
        .edit_body_table_dimension_size(0usize, Dimension::Row(0))?
        .set(points(32.0)?)
        .commit()?;
    assert_eq!(
        changed
            .package()
            .body_table_dimension_size(0usize, Dimension::Row(0))?,
        points(32.0)?
    );
    let reset = changed
        .package()
        .edit_body_table_dimension_size(0usize, Dimension::Row(0))?
        .reset()
        .commit()?;
    assert_eq!(
        reset
            .package()
            .body_table_dimension_size(0usize, Dimension::Row(0))?,
        Size::Default
    );
    let reset_bytes = reset.package().exact_bytes();
    assert_eq!(
        member_bytes(&reset_bytes, DOCUMENT_MEMBER)?,
        member_bytes(&source, DOCUMENT_MEMBER)?
    );
    assert_eq!(sentinel(&reset_bytes)?, b"untouched-sentinel");
    assert_eq!(
        Catalog::from_bytes(&reset_bytes)?.iter().count(),
        2,
        "a changed transaction invalidates the three preview members"
    );
    Ok(())
}

#[test]
fn row_and_column_changes_are_local_reversible_and_preview_invalidating() -> TestResult {
    let source = synthetic_package()?;
    for (dimension, after) in [
        (Dimension::Row(0), points(32.0)?),
        (Dimension::Column(2), points(80.0)?),
    ] {
        let package = Package::from_bytes(&source)?;
        let commit = package
            .edit_body_table_dimension_size(0usize, dimension)?
            .set(after)
            .commit()?;
        assert_eq!(commit.patch().after(), after);
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().touched_components(), 1);
        assert_eq!(commit.diagnostics().deleted_previews(), PREVIEWS.len());
        assert!(commit.diagnostics().full_reparse_performed());
        assert_eq!(
            sentinel(&commit.package().exact_bytes())?,
            b"untouched-sentinel"
        );
        assert_eq!(
            commit
                .package()
                .body_table_dimension_size(0usize, dimension)?,
            after
        );
        let target = commit.package().exact_bytes();
        assert_eq!(
            package
                .apply_body_table_dimension_size(commit.patch())?
                .package()
                .exact_bytes(),
            target
        );
        let restored = commit
            .package()
            .apply_body_table_dimension_size(&commit.patch().inverse())?;
        assert_eq!(restored.package().exact_bytes(), source);
    }
    Ok(())
}

#[test]
fn cross_component_header_bucket_rewrites_only_bucket_member() -> TestResult {
    let source = split_row_bucket_component(&synthetic_package()?)?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.body_table_dimension_size(0usize, Dimension::Row(1))?,
        points(25.0)?
    );
    let document_before = member_bytes(&source, DOCUMENT_MEMBER)?;
    let bucket_before = member_bytes(&source, ROW_BUCKET_MEMBER)?;
    let commit = package
        .edit_body_table_dimension_size(0usize, Dimension::Row(1))?
        .set(points(31.0)?)
        .commit()?;
    let target = commit.package().exact_bytes();
    assert_eq!(
        commit
            .package()
            .body_table_dimension_size(0usize, Dimension::Row(1))?,
        points(31.0)?
    );
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert_eq!(commit.diagnostics().deleted_previews(), PREVIEWS.len());
    assert_eq!(member_bytes(&target, DOCUMENT_MEMBER)?, document_before);
    assert_ne!(member_bytes(&target, ROW_BUCKET_MEMBER)?, bucket_before);
    assert_eq!(
        sentinel(&target)?,
        b"untouched-sentinel",
        "an out-of-component rewrite must not perturb unrelated members"
    );
    let restored = commit
        .package()
        .apply_body_table_dimension_size(&commit.patch().inverse())?;
    assert_eq!(restored.package().exact_bytes(), source);
    Ok(())
}

#[test]
fn unknown_bucket_and_header_fields_survive_size_rewrite() -> TestResult {
    let source = synthetic_package()?;
    let package = Package::from_bytes(&source)?;
    let before = bucket_bytes(&source, ROW_BUCKET_IDENTIFIER)?;
    let commit = package
        .edit_body_table_dimension_size(0usize, Dimension::Row(1))?
        .set(points(31.0)?)
        .commit()?;
    let after = bucket_bytes(&commit.package().exact_bytes(), ROW_BUCKET_IDENTIFIER)?;
    let bucket_view = WireView::parse(&after)?;
    let bucket_unknown = bucket_view
        .fields()
        .find(|field| field.number() == UNKNOWN_BUCKET_FIELD)
        .ok_or("bucket unknown field was dropped")?;
    assert_eq!(
        decode_varint_from_bytes(bucket_unknown.payload())?.0,
        UNKNOWN_BUCKET_VALUE
    );
    let header = bucket_view
        .fields()
        .find(|field| field.number() == 2)
        .ok_or("header was dropped")?;
    let header_view = WireView::parse(header.payload())?;
    let header_unknown = header_view
        .fields()
        .find(|field| field.number() == UNKNOWN_HEADER_FIELD)
        .ok_or("header unknown field was dropped")?;
    assert_eq!(
        decode_varint_from_bytes(header_unknown.payload())?.0,
        UNKNOWN_HEADER_VALUE
    );
    assert_ne!(before, after);
    Ok(())
}

#[test]
fn output_limit_rejects_growth_before_publication() -> TestResult {
    let source = without_previews(&synthetic_package()?)?;
    let package = Package::from_bytes(&source)?;
    let unrestricted = package
        .edit_body_table_dimension_size(0usize, Dimension::Row(0))?
        .set(points(32.0)?)
        .commit()?;
    let target = unrestricted.package().exact_bytes();
    assert!(
        target.len() > source.len(),
        "default-header insertion should grow the no-preview package"
    );
    let limits = Limits::new(
        u64::try_from(source.len())?,
        32,
        1024 * 1024,
        1024 * 1024,
        1024 * 1024,
    )?;
    let bounded = Package::from_bytes_with_limits(&source, limits)?;
    let before = bounded.exact_bytes();
    let error = bounded
        .edit_body_table_dimension_size(0usize, Dimension::Row(0))?
        .set(points(32.0)?)
        .commit()
        .expect_err("output growth must be rejected at the physical ceiling");
    assert!(matches!(error, Error::LimitExceeded { .. }));
    assert_eq!(bounded.exact_bytes(), before);
    Ok(())
}

#[test]
fn header_field_limit_has_a_prepublication_boundary() -> TestResult {
    let source = synthetic_package()?;
    let requested = points(31.0)?;
    let mut successful_fields = None;
    for fields in 1..=4_096 {
        let archive_limits = litchi_iwa_core::Limits::default().with_header_fields(fields)?;
        let limits = Limits::default().with_archive_limits(archive_limits)?;
        let Ok(package) = Package::from_bytes_with_limits(&source, limits) else {
            continue;
        };
        let result = package
            .edit_body_table_dimension_size(0usize, Dimension::Row(1))
            .and_then(|edit| edit.set(requested).commit());
        if result.is_ok() {
            successful_fields = Some(fields);
            break;
        }
    }
    let exact_fields = successful_fields.ok_or("no field-budget boundary found")?;
    assert!(exact_fields > 1);
    let archive_limits = litchi_iwa_core::Limits::default().with_header_fields(exact_fields - 1)?;
    let limits = Limits::default().with_archive_limits(archive_limits)?;
    if let Ok(package) = Package::from_bytes_with_limits(&source, limits) {
        let before = package.exact_bytes();
        let result = package
            .edit_body_table_dimension_size(0usize, Dimension::Row(1))
            .and_then(|edit| edit.set(requested).commit());
        assert!(matches!(result, Err(Error::LimitExceeded { .. })));
        assert_eq!(package.exact_bytes(), before);
    }
    Ok(())
}

#[test]
fn lock_alias_missing_and_malformed_buckets_fail_atomically() -> TestResult {
    let source = synthetic_package()?;
    let requested = points(31.0)?;
    for malformed in [
        append_duplicate_header(&source, ROW_BUCKET_IDENTIFIER)?,
        replace_bucket_message_type(&source, ROW_BUCKET_IDENTIFIER)?,
        remove_bucket(&source, ROW_BUCKET_IDENTIFIER)?,
    ] {
        let package = Package::from_bytes(&malformed)?;
        let before = package.exact_bytes();
        let result = package
            .edit_body_table_dimension_size(0usize, Dimension::Row(1))
            .and_then(|edit| edit.set(requested).commit());
        assert!(matches!(
            result,
            Err(Error::InvalidSource)
                | Err(Error::UnsupportedSource)
                | Err(Error::LimitExceeded { .. })
        ));
        assert_eq!(package.exact_bytes(), before);
    }
    let package = Package::from_bytes(&source)?;
    let mut lock = package.edit_body_table_lock(0usize)?;
    lock.lock();
    let locked = lock.commit()?.into_package();
    let error = locked
        .edit_body_table_dimension_size(0usize, Dimension::Row(0))?
        .set(requested)
        .commit()
        .expect_err("locked table dimension must refuse");
    assert!(matches!(error, Error::TableLocked));
    Ok(())
}

#[test]
fn limits_and_public_values_are_typed_and_redacted() -> TestResult {
    fn assert_send_sync_debug<T: Send + Sync + std::fmt::Debug>() {}
    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<Dimension>();
    assert_send_sync_debug::<Points>();
    assert_send_sync_debug::<Size>();
    assert_send_sync_debug::<litchi_pages::BodyTableDimensionEdit<'static>>();
    assert_send_sync_debug::<litchi_pages::BodyTableDimensionPatch>();
    assert_send_sync_debug::<litchi_pages::BodyTableDimensionCommit>();
    assert_send_sync_debug::<litchi_pages::BodyTableDimensionDiagnostics>();
    assert_send_sync_debug::<litchi_pages::BodyTableDimensionLimitKind>();
    assert_send_sync_debug::<Error>();
    let source = synthetic_package()?;
    let archive_limits = litchi_iwa_core::Limits::default().with_header_fields(1)?;
    let limits = Limits::default().with_archive_limits(archive_limits)?;
    if let Ok(package) = Package::from_bytes_with_limits(&source, limits) {
        let error = package
            .edit_body_table_dimension_size(0usize, Dimension::Row(0))?
            .set_points(Points::new(31.0)?)
            .commit()
            .expect_err("tight header limits must reject before publication");
        assert!(matches!(
            error,
            Error::LimitExceeded { .. } | Error::InvalidSource
        ));
    }
    let package = Package::from_bytes(&source)?;
    let edit = package.edit_body_table_dimension_size(0usize, Dimension::Row(0))?;
    let rendered = format!("{edit:?}");
    assert!(!rendered.contains(DOCUMENT_MEMBER));
    Ok(())
}
