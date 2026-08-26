//! Native integration coverage for selector-first Keynote slide-table titles.

use std::error::Error as StdError;
use std::io;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::{
    decode_varint_from_bytes,
    wire::{WireView, append_varint_field},
};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldPath, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsd, tsk, tsp, tss, tst, tswp};
use litchi_keynote::slide::table::title::Settings;
use litchi_keynote::{Package, SlideSelector, TableSelector};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const SLIDE_NODE: u64 = 3;
const SLIDE: u64 = 4;
const TABLE_INFOS: [u64; 2] = [100, 101];
const MODELS: [u64; 2] = [110, 111];
const TITLE_STYLE: u64 = 120;
const SHAPE_STYLE: u64 = 121;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const UNKNOWN_MODEL_FIELD: u32 = 99;
const UNKNOWN_MODEL_VALUE: u64 = 0xfeed_beef;
const PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn field_reference(path: impl Into<FieldPath>, references: &[u64]) -> FieldInfo {
    let mut field = FieldInfo::new(path);
    field.object_references.extend_from_slice(references);
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

fn table_model(name: &str, visible: Option<bool>, outlined: Option<bool>) -> Vec<u8> {
    let mut payload = tst::TableModelArchive {
        table_id: format!("table-{name}"),
        table_name: name.to_owned(),
        table_name_enabled: visible,
        table_name_height: Some(20.0),
        table_name_border_enabled: outlined,
        table_name_style: Some(reference(TITLE_STYLE)),
        table_name_shape_style: Some(reference(SHAPE_STYLE)),
        number_of_rows: 1,
        number_of_columns: 1,
        default_row_height: 20.0,
        default_column_width: 64.0,
        ..tst::TableModelArchive::default()
    }
    .encode_to_vec();
    append_varint_field(&mut payload, UNKNOWN_MODEL_FIELD, UNKNOWN_MODEL_VALUE)
        .expect("synthetic unknown title field fits");
    payload
}

fn paragraph_style_payload() -> Vec<u8> {
    tswp::ParagraphStyleArchive {
        super_: tss::StyleArchive {
            style_identifier: Some("slide-table-title".to_owned()),
            ..tss::StyleArchive::default()
        },
        ..tswp::ParagraphStyleArchive::default()
    }
    .encode_to_vec()
}

fn shape_style_payload() -> Vec<u8> {
    tswp::ShapeStyleArchive {
        super_: tsd::ShapeStyleArchive {
            super_: tss::StyleArchive {
                style_identifier: Some("slide-table-title-shape".to_owned()),
                ..tss::StyleArchive::default()
            },
            ..tsd::ShapeStyleArchive::default()
        },
        ..tswp::ShapeStyleArchive::default()
    }
    .encode_to_vec()
}

fn synthetic_package() -> TestResult<Vec<u8>> {
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..tsa::DocumentArchive::default()
        },
        show: reference(2),
        ..kn::DocumentArchive::default()
    };
    let show = kn::ShowArchive {
        theme: reference(80),
        slide_tree: kn::SlideTreeArchive {
            slides: vec![reference(SLIDE_NODE)],
            ..kn::SlideTreeArchive::default()
        },
        size: tsp::Size {
            width: 1_024.0,
            height: 768.0,
        },
        stylesheet: reference(81),
        ..kn::ShowArchive::default()
    };
    #[allow(deprecated, reason = "native schema retains cache fields")]
    let node = kn::SlideNodeArchive {
        slide: Some(reference(SLIDE)),
        is_skipped: false,
        has_builds: false,
        has_transition: false,
        ..kn::SlideNodeArchive::default()
    };
    let slide = kn::SlideArchive {
        style: reference(90),
        transition: kn::TransitionArchive::default(),
        owned_drawables: TABLE_INFOS.iter().copied().map(reference).collect(),
        drawables_z_order: TABLE_INFOS.iter().copied().map(reference).collect(),
        name: Some("Tables".to_owned()),
        in_document: true,
        ..kn::SlideArchive::default()
    };
    let mut slide_object = object(SLIDE, 5, slide.encode_to_vec(), &TABLE_INFOS)?;
    slide_object.archive_info.message_infos[0]
        .field_infos
        .extend([
            field_reference(vec![7], &TABLE_INFOS),
            field_reference(vec![42], &TABLE_INFOS),
        ]);
    let mut objects = vec![
        object(1, 1, document.encode_to_vec(), &[2])?,
        object(2, 2, show.encode_to_vec(), &[SLIDE_NODE, 80, 81])?,
        object(SLIDE_NODE, 4, node.encode_to_vec(), &[SLIDE])?,
        slide_object,
    ];
    for index in 0..TABLE_INFOS.len() {
        let info = tst::TableInfoArchive {
            super_: tsd::DrawableArchive {
                parent: Some(reference(SLIDE)),
                locked: Some(false),
                ..tsd::DrawableArchive::default()
            },
            table_model: reference(MODELS[index]),
            ..tst::TableInfoArchive::default()
        };
        let mut info_object = object(
            TABLE_INFOS[index],
            TABLE_INFO_MESSAGE_TYPE,
            info.encode_to_vec(),
            &[SLIDE, MODELS[index]],
        )?;
        info_object.archive_info.message_infos[0]
            .field_infos
            .push(field_reference(vec![2], &[MODELS[index]]));
        objects.push(info_object);

        let mut model_object = object(
            MODELS[index],
            TABLE_MODEL_MESSAGE_TYPE,
            table_model(
                if index == 0 { "Revenue" } else { "Costs" },
                Some(index == 0),
                if index == 0 { None } else { Some(false) },
            ),
            &[TITLE_STYLE, SHAPE_STYLE],
        )?;
        model_object.archive_info.message_infos[0]
            .field_infos
            .extend([
                field_reference(vec![30], &[TITLE_STYLE]),
                field_reference(vec![36], &[SHAPE_STYLE]),
            ]);
        objects.push(model_object);
    }
    objects.push(object(TITLE_STYLE, 2_022, paragraph_style_payload(), &[])?);
    objects.push(object(SHAPE_STYLE, 2_025, shape_style_payload(), &[])?);
    let compressed = SnappyStream::compress(&Archive { objects }.to_bytes()?)?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", b"unrelated ZIP sentinel".as_slice()),
            (DOCUMENT_MEMBER, compressed.as_slice()),
            (PREVIEWS[0], b"large preview".as_slice()),
            (PREVIEWS[1], b"micro preview".as_slice()),
            (PREVIEWS[2], b"web preview".as_slice()),
        ],
        Limits::default(),
    )?)
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn model_payload(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("missing synthetic document member"))?;
    let archive = Archive::parse(&SnappyStream::decompress(entry.data())?.into_bytes())?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing synthetic table model"))?;
    Ok(object
        .messages
        .iter()
        .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("missing synthetic table-model message"))?
        .data
        .clone())
}

fn replace_model_payload(package: &[u8], identifier: u64, payload: Vec<u8>) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let document = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("missing synthetic document member"))?;
    let mut archive = Archive::parse(&SnappyStream::decompress(document.data())?.into_bytes())?;
    let object = archive
        .object_mut(identifier)
        .ok_or_else(|| io::Error::other("missing synthetic table model"))?;
    let message = object
        .messages
        .iter_mut()
        .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("missing synthetic table-model message"))?;
    message.data = payload;
    object.archive_info.message_infos[0].length = message.data.len().try_into()?;
    let replacement = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            DOCUMENT_MEMBER,
            &replacement,
        )],
        Limits::default(),
    )?)
}

fn unknown_value(payload: &[u8]) -> TestResult<u64> {
    let fields = WireView::parse(payload)?;
    let field = fields
        .fields()
        .find(|field| field.number() == UNKNOWN_MODEL_FIELD)
        .ok_or_else(|| io::Error::other("missing unknown model field"))?;
    Ok(decode_varint_from_bytes(field.payload())?.0)
}

#[test]
fn selector_read_noop_change_apply_inverse_and_locality() -> TestResult {
    let source_bytes = synthetic_package()?;
    let package = Package::from_bytes(&source_bytes)?;
    assert_eq!(
        package.slide_table_title_settings("Tables", TableSelector::index(0))?,
        Settings::new(Some(true), None),
    );
    assert_eq!(
        package.slide_table_title_settings(SlideSelector::index(0), 1usize)?,
        Settings::new(Some(false), Some(false)),
    );

    let noop = package
        .edit_slide_table_title("Tables", TableSelector::index(0))?
        .commit()?;
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(exact_bytes(noop.package())?, source_bytes);

    let before_first = model_payload(&source_bytes, MODELS[0])?;
    let before_second = model_payload(&source_bytes, MODELS[1])?;
    let replacement = Settings::new(Some(false), Some(true));
    let commit = package
        .edit_slide_table_title("Tables", TableSelector::index(0))?
        .set(replacement)
        .commit()?;
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert_eq!(commit.diagnostics().deleted_previews(), 3);
    assert_eq!(
        commit
            .package()
            .slide_table_title_settings("Tables", TableSelector::index(0))?,
        replacement,
    );
    let candidate_bytes = exact_bytes(commit.package())?;
    assert_eq!(
        unknown_value(&model_payload(&candidate_bytes, MODELS[0])?)?,
        UNKNOWN_MODEL_VALUE
    );
    assert_eq!(model_payload(&candidate_bytes, MODELS[1])?, before_second);
    assert_ne!(model_payload(&candidate_bytes, MODELS[0])?, before_first);
    let candidate_catalog = Catalog::from_bytes(&candidate_bytes)?;
    assert!(
        PREVIEWS
            .iter()
            .all(|name| candidate_catalog.iter().all(|entry| entry.name() != *name))
    );
    assert_eq!(
        candidate_catalog
            .iter()
            .find(|entry| entry.name() == "Data/sentinel.bin")
            .map(|entry| entry.data()),
        Some(b"unrelated ZIP sentinel".as_slice()),
    );

    let applied = package.apply_slide_table_title(commit.patch())?;
    assert_eq!(exact_bytes(applied.package())?, candidate_bytes);
    assert!(matches!(
        commit.package().apply_slide_table_title(commit.patch()),
        Err(litchi_keynote::SlideTableTitleError::PatchConflict)
    ));
    let restored = commit
        .package()
        .apply_slide_table_title(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source_bytes);
    Ok(())
}

#[test]
fn positions_and_locks_fail_closed_without_mutating_source() -> TestResult {
    let source = synthetic_package()?;
    let package = Package::from_bytes(&source)?;
    assert!(matches!(
        package.slide_table_title_settings("Tables", TableSelector::index(2)),
        Err(litchi_keynote::SlideTableTitleError::TablePositionNotFound { .. })
    ));
    assert_eq!(exact_bytes(&package)?, source);

    let catalog = Catalog::from_bytes(&source)?;
    let document = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("missing document"))?;
    let mut archive = Archive::parse(&SnappyStream::decompress(document.data())?.into_bytes())?;
    let table = archive
        .object_mut(TABLE_INFOS[0])
        .ok_or_else(|| io::Error::other("missing table info"))?;
    let info = tst::TableInfoArchive::decode(table.messages[0].data.as_slice())?;
    table.messages[0].data = tst::TableInfoArchive {
        super_: tsd::DrawableArchive {
            locked: Some(true),
            ..info.super_
        },
        ..info
    }
    .encode_to_vec();
    table.archive_info.message_infos[0].length = table.messages[0].data.len().try_into()?;
    let locked_member = SnappyStream::compress(&archive.to_bytes()?)?;
    let locked_source = catalog.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            DOCUMENT_MEMBER,
            &locked_member,
        )],
        Limits::default(),
    )?;
    let locked = Package::from_bytes(&locked_source)?;
    assert_eq!(
        locked.slide_table_title_settings("Tables", 0usize)?,
        Settings::new(Some(true), None)
    );
    assert!(matches!(
        locked
            .edit_slide_table_title("Tables", 0usize)?
            .set(Settings::new(Some(false), None))
            .commit(),
        Err(litchi_keynote::SlideTableTitleError::Locked)
    ));
    assert_eq!(exact_bytes(&locked)?, locked_source);
    Ok(())
}

#[test]
fn every_optional_presence_combination_roundtrips_exactly() -> TestResult {
    let source = synthetic_package()?;
    for visible in [None, Some(false), Some(true)] {
        for outlined in [None, Some(false), Some(true)] {
            let desired = Settings::new(visible, outlined);
            let package = Package::from_bytes(&source)?;
            let commit = package
                .edit_slide_table_title(SlideSelector::index(0), TableSelector::index(0))?
                .set(desired)
                .commit()?;
            assert_eq!(
                commit
                    .package()
                    .slide_table_title_settings(SlideSelector::index(0), TableSelector::index(0))?,
                desired,
            );
            let restored = commit
                .package()
                .apply_slide_table_title(&commit.patch().inverse())?;
            assert_eq!(exact_bytes(restored.package())?, source);
        }
    }
    Ok(())
}

#[test]
fn duplicate_known_title_field_is_rejected_atomically() -> TestResult {
    let source = synthetic_package()?;
    let mut malformed_payload = model_payload(&source, MODELS[0])?;
    append_varint_field(&mut malformed_payload, 22, 0)?;
    let malformed_source = replace_model_payload(&source, MODELS[0], malformed_payload)?;
    let package = Package::from_bytes(&malformed_source)?;
    assert!(matches!(
        package.slide_table_title_settings(SlideSelector::index(0), TableSelector::index(0)),
        Err(litchi_keynote::SlideTableTitleError::InvalidSource)
    ));
    assert!(matches!(
        package.edit_slide_table_title(SlideSelector::index(0), TableSelector::index(0)),
        Err(litchi_keynote::SlideTableTitleError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, malformed_source);
    Ok(())
}
