//! Negative proof cases for iWork root-comment clearing.
//!
//! These cases deliberately keep successful root-only clears and then alter
//! only the selected type-3056 storage object or its graph.  An opaque sibling
//! message, an unknown payload field, or an unsafe Numbers header/table graph
//! must produce a typed terminal refusal; the package and selected semantic
//! comment must remain unchanged.

use litchi_iwa::keynote;
use litchi_iwa::pages;
use litchi_iwa_archive::iwa::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_common::{WireLimits, decode_varint_from_bytes, encode_varint_into};
use litchi_iwa_common::{table::coordinate::At, wire};
use litchi_keynote::slide::table::comments::Error as KeynoteCommentError;
use litchi_keynote::{Package as KeynotePackage, SlideSelector, TableSelector};
use litchi_numbers::Package as NumbersPackage;
use litchi_numbers::cell::comment::Error as NumbersCommentError;
use litchi_pages::table::comments::Error as PagesCommentError;
use litchi_pages::{BodyTableSelector, Package as PagesPackage};

use crate::harness::{TestResult, fixture};

const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const OPAQUE_SIBLING_MESSAGE_TYPE: u32 = 99_999;
const UNKNOWN_PAYLOAD_FIELD: u32 = 99;
const NUMBERS_ROOT_ID: u64 = 51;
const NUMBERS_DOCUMENT_MEMBER: &str = "Index/Document.iwa";

const NUMBERS_SOURCE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../litchi-numbers/tests/fixtures/comment-edit-root.numbers"
));

#[derive(Clone, Copy)]
enum StorageMutation {
    OpaqueSibling,
    UnknownPayload,
}

#[derive(Clone, Copy)]
enum NumbersMutation {
    UnknownSurvivingReference,
    SelectedStorageShouldMerge,
    MalformedTableInfo,
}

fn pages_root_source() -> TestResult<(Vec<u8>, At)> {
    let input = fixture("iwork/pages/body-table-cells-native.pages")?;
    let mut editor = pages::Draft::from_bytes(&input.bytes)?;
    let model = editor
        .tables()?
        .first()
        .ok_or("Pages source has no table")?
        .model_object_id;
    let (row, column) = (1, 1);
    if editor.table(model)?.get_cell(row, column).is_none() {
        return Err("Pages source has no stored root-only target cell".into());
    }
    editor.set_table_cell_comment(model, row, column, "clear-safety root")?;
    if !editor
        .table_cell_comment_replies(model, row, column)?
        .is_empty()
    {
        return Err("Pages host baseline unexpectedly has replies".into());
    }
    Ok((
        editor.to_bytes()?,
        At::new(u32::try_from(row)?, u32::try_from(column)?),
    ))
}

fn keynote_root_source() -> TestResult<(Vec<u8>, At)> {
    let input = fixture("iwork/keynote/slide-table-read-native.key")?;
    let mut editor = keynote::Draft::from_bytes(&input.bytes)?;
    let slide = 0;
    let model = editor
        .slide_tables(slide)?
        .first()
        .ok_or("Keynote source has no table")?
        .model_object_id;
    let position = editor
        .slide_table(slide, model)?
        .iter_cells()
        .map(|((row, column), _)| Ok(At::new(u32::try_from(row)?, u32::try_from(column)?)))
        .collect::<TestResult<Vec<_>>>()?
        .into_iter()
        .find(|position| {
            editor
                .slide_table_cell_comment_replies(
                    slide,
                    model,
                    usize::try_from(position.row()).unwrap_or(usize::MAX),
                    usize::try_from(position.column()).unwrap_or(usize::MAX),
                )
                .is_ok_and(|replies| replies.is_empty())
        })
        .ok_or("Keynote source has no root-only table cell")?;
    editor.set_slide_table_cell_comment(
        slide,
        model,
        usize::try_from(position.row())?,
        usize::try_from(position.column())?,
        "clear-safety root",
    )?;
    if !editor
        .slide_table_cell_comment_replies(
            slide,
            model,
            usize::try_from(position.row())?,
            usize::try_from(position.column())?,
        )?
        .is_empty()
    {
        return Err("Keynote host baseline unexpectedly has replies".into());
    }
    Ok((editor.to_bytes()?, position))
}

fn mutate_selected_storage(source: &[u8], mutation: StorageMutation) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source, Default::default())?;
    let mut edited = Vec::new();
    let mut selected = 0usize;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data(), Default::default())?;
        let mut archive = Archive::parse(stream.as_bytes(), Default::default())?;
        let mut changed = false;
        for object in &mut archive.objects {
            if !object
                .messages
                .iter()
                .any(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
            {
                continue;
            }
            selected = selected.saturating_add(1);
            if selected != 1 {
                return Err("root-only baseline contains multiple type-3056 objects".into());
            }
            let message = object
                .messages
                .iter_mut()
                .find(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
                .ok_or("selected type-3056 message disappeared")?;
            match mutation {
                StorageMutation::OpaqueSibling => {
                    object.push_message(
                        RawMessage {
                            type_: OPAQUE_SIBLING_MESSAGE_TYPE,
                            data: vec![0x08, 0x01],
                        },
                        Default::default(),
                    )?;
                },
                StorageMutation::UnknownPayload => {
                    litchi_iwa_common::wire::append_varint_field(
                        &mut message.data,
                        UNKNOWN_PAYLOAD_FIELD,
                        1,
                    )?;
                },
            }
            changed = true;
        }
        if changed {
            let compressed =
                SnappyStream::compress(&archive.to_bytes(Default::default())?, Default::default())?;
            edited.push((entry.name().to_owned(), compressed));
        }
    }
    if selected != 1 {
        return Err(format!("expected one selected type-3056 object, found {selected}").into());
    }
    let edits = edited
        .iter()
        .map(|(name, bytes)| EntryEdit::new(name, bytes))
        .collect::<Vec<_>>();
    Ok(catalog.reassemble_to_bytes(&edits, Default::default())?)
}

fn rewrite_numbers_member(
    source: &[u8],
    rewrite: impl FnOnce(Vec<u8>) -> TestResult<Vec<u8>>,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source, Default::default())?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == NUMBERS_DOCUMENT_MEMBER)
        .ok_or("Numbers comment fixture has no Document member")?;
    let stream = SnappyStream::decompress(entry.data(), Default::default())?;
    let rewritten = rewrite(stream.as_bytes().to_vec())?;
    let compressed = SnappyStream::compress(&rewritten, Default::default())?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(NUMBERS_DOCUMENT_MEMBER, &compressed)],
        Default::default(),
    )?)
}

fn rewrite_numbers_object_header(
    stream: &[u8],
    identifier: u64,
    mutate: impl FnOnce(&mut Vec<u8>) -> TestResult,
) -> TestResult<Vec<u8>> {
    let archive = Archive::parse(stream, Default::default())?;
    let object = archive
        .object(identifier)
        .ok_or("Numbers header mutation target is missing")?;
    let header_offset = usize::try_from(object.header_offset)?;
    let data_offset = usize::try_from(object.data_offset)?;
    let (header_length, prefix_length) = decode_varint_from_bytes(&stream[header_offset..])?;
    let header_start = header_offset
        .checked_add(prefix_length)
        .ok_or_else(|| std::io::Error::other("Numbers header mutation start overflow"))?;
    let header_end = header_start
        .checked_add(usize::try_from(header_length)?)
        .ok_or_else(|| std::io::Error::other("Numbers header mutation end overflow"))?;
    if header_end != data_offset {
        return Err("Numbers object header framing is inconsistent".into());
    }
    let mut header = stream[header_start..header_end].to_vec();
    mutate(&mut header)?;
    let mut rewritten = Vec::with_capacity(stream.len().saturating_add(header.len()));
    rewritten.extend_from_slice(&stream[..header_offset]);
    encode_varint_into(&mut rewritten, u64::try_from(header.len())?);
    rewritten.extend_from_slice(&header);
    rewritten.extend_from_slice(&stream[data_offset..]);
    Archive::parse(&rewritten, Default::default())?;
    Ok(rewritten)
}

fn mutate_numbers_source(mutation: NumbersMutation) -> TestResult<Vec<u8>> {
    match mutation {
        NumbersMutation::UnknownSurvivingReference => {
            rewrite_numbers_member(NUMBERS_SOURCE, |stream| {
                rewrite_numbers_object_header(&stream, 1, |header| {
                    let mut reference = Vec::new();
                    wire::append_varint_field(&mut reference, 1, NUMBERS_ROOT_ID)?;
                    wire::append_length_delimited_field(
                        header,
                        UNKNOWN_PAYLOAD_FIELD,
                        &reference,
                        WireLimits::DEFAULT,
                    )?;
                    Ok(())
                })
            })
        },
        NumbersMutation::SelectedStorageShouldMerge => {
            rewrite_numbers_member(NUMBERS_SOURCE, |stream| {
                rewrite_numbers_object_header(&stream, NUMBERS_ROOT_ID, |header| {
                    wire::append_varint_field(header, 3, 1)?;
                    Ok(())
                })
            })
        },
        NumbersMutation::MalformedTableInfo => rewrite_numbers_member(NUMBERS_SOURCE, |stream| {
            let mut archive = Archive::parse(&stream, Default::default())?;
            archive.objects.push(ArchiveObject::new(
                99_998,
                vec![RawMessage {
                    type_: TABLE_INFO_MESSAGE_TYPE,
                    data: vec![0xff],
                }],
                Default::default(),
            )?);
            Ok(archive.to_bytes(Default::default())?)
        }),
    }
}

fn numbers_target(package: &NumbersPackage) -> TestResult<(String, String)> {
    let sheet = package
        .sheets()
        .first()
        .ok_or("Numbers comment fixture has no sheet")?;
    let table = sheet
        .tables()
        .next()
        .ok_or("Numbers comment fixture has no table")?;
    Ok((sheet.name().to_owned(), table.name().to_owned()))
}

fn numbers_bytes(package: &NumbersPackage) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn assert_numbers_clear_refuses(source: &[u8], expected: fn(&NumbersCommentError)) -> TestResult {
    let package = NumbersPackage::from_bytes(source)?;
    let before = numbers_bytes(&package)?;
    assert_eq!(
        before, source,
        "Numbers mutation did not preserve source bytes"
    );
    let (sheet, table) = numbers_target(&package)?;
    let position = At::new(1, 1);
    let selected_before = package
        .table_cell_comment(&sheet, &table, position)?
        .ok_or("Numbers mutation lost the selected root before clear")?;
    let result = package.clear_table_cell_comment(&sheet, &table, position);
    match result {
        Ok(_) => panic!("Numbers clear unexpectedly published a lossy candidate"),
        Err(error) => expected(&error),
    }
    assert_eq!(
        numbers_bytes(&package)?,
        before,
        "Numbers refusal mutated source"
    );
    assert_eq!(
        NumbersPackage::from_bytes(source)?.table_cell_comment(&sheet, &table, position)?,
        Some(selected_before),
        "Numbers refusal changed the selected root"
    );
    Ok(())
}

fn expect_numbers_unsupported(error: &NumbersCommentError) {
    assert!(
        matches!(
            error,
            NumbersCommentError::Refused(
                litchi_numbers::cell::comment::Refusal::UnsupportedDependency { .. }
            )
        ),
        "Numbers clear returned the wrong typed error: {error:?}"
    );
}

fn expect_numbers_invalid(error: &NumbersCommentError) {
    assert!(
        matches!(error, NumbersCommentError::InvalidSource { .. }),
        "Numbers clear returned the wrong typed error: {error:?}"
    );
}

fn pages_bytes(package: &PagesPackage) -> TestResult<Vec<u8>> {
    Ok(package.to_bytes()?)
}

fn keynote_bytes(package: &KeynotePackage) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn assert_pages_clear_refuses(source: &[u8], position: At) -> TestResult {
    let package = PagesPackage::from_bytes(source)?;
    let before = pages_bytes(&package)?;
    assert_eq!(
        before, source,
        "Pages mutation did not preserve its source bytes"
    );
    let selected_before = package
        .body_table_cell_comment(BodyTableSelector::index(0), position)?
        .ok_or("Pages mutation lost the selected root before clear")?;
    let result = package
        .edit_body_table_cell_comment(BodyTableSelector::index(0), position)?
        .clear()?
        .commit();
    match result {
        Ok(_) => panic!("Pages clear unexpectedly published a lossy candidate"),
        Err(PagesCommentError::UnsupportedDependency) => {},
        Err(error) => panic!("Pages clear returned the wrong typed error: {error:?}"),
    }
    assert_eq!(
        pages_bytes(&package)?,
        before,
        "Pages refusal mutated source"
    );
    assert_eq!(
        PagesPackage::from_bytes(source)?
            .body_table_cell_comment(BodyTableSelector::index(0), position,)?,
        Some(selected_before),
        "Pages refusal changed the selected root"
    );
    Ok(())
}

fn assert_keynote_clear_refuses(source: &[u8], position: At) -> TestResult {
    let package = KeynotePackage::from_bytes(source)?;
    let before = keynote_bytes(&package)?;
    assert_eq!(
        before, source,
        "Keynote mutation did not preserve its source bytes"
    );
    let selected_before = package
        .slide_table_cell_comment(SlideSelector::index(0), TableSelector::index(0), position)?
        .ok_or("Keynote mutation lost the selected root before clear")?;
    let result = package
        .edit_slide_table_cell_comment(SlideSelector::index(0), TableSelector::index(0), position)?
        .clear()?
        .commit();
    match result {
        Ok(_) => panic!("Keynote clear unexpectedly published a lossy candidate"),
        Err(KeynoteCommentError::UnsupportedDependency) => {},
        Err(error) => panic!("Keynote clear returned the wrong typed error: {error:?}"),
    }
    assert_eq!(
        keynote_bytes(&package)?,
        before,
        "Keynote refusal mutated source"
    );
    assert_eq!(
        KeynotePackage::from_bytes(source)?.slide_table_cell_comment(
            SlideSelector::index(0),
            TableSelector::index(0),
            position,
        )?,
        Some(selected_before),
        "Keynote refusal changed the selected root"
    );
    Ok(())
}

#[test]
fn pages_root_clear_baseline_is_successful_and_reversible() -> TestResult {
    let (source, position) = pages_root_source()?;
    let package = PagesPackage::from_bytes(&source)?;
    assert!(
        package
            .body_table_cell_comment(BodyTableSelector::index(0), position)?
            .is_some()
    );
    let clear = package
        .edit_body_table_cell_comment(BodyTableSelector::index(0), position)?
        .clear()?
        .commit()?;
    assert!(
        clear
            .package()
            .body_table_cell_comment(BodyTableSelector::index(0), position)?
            .is_none()
    );
    let restored = clear
        .package()
        .apply_body_table_cell_comment(&clear.patch().inverse())?;
    assert_eq!(pages_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn pages_root_clear_refuses_opaque_storage_sibling() -> TestResult {
    let (source, position) = pages_root_source()?;
    let corrupted = mutate_selected_storage(&source, StorageMutation::OpaqueSibling)?;
    assert_pages_clear_refuses(&corrupted, position)
}

#[test]
fn pages_root_clear_refuses_unknown_storage_payload() -> TestResult {
    let (source, position) = pages_root_source()?;
    let corrupted = mutate_selected_storage(&source, StorageMutation::UnknownPayload)?;
    assert_pages_clear_refuses(&corrupted, position)
}

#[test]
fn keynote_root_clear_baseline_is_successful_and_reversible() -> TestResult {
    let (source, position) = keynote_root_source()?;
    let package = KeynotePackage::from_bytes(&source)?;
    assert!(
        package
            .slide_table_cell_comment(SlideSelector::index(0), TableSelector::index(0), position)?
            .is_some()
    );
    let clear = package
        .edit_slide_table_cell_comment(SlideSelector::index(0), TableSelector::index(0), position)?
        .clear()?
        .commit()?;
    assert!(
        clear
            .package()
            .slide_table_cell_comment(SlideSelector::index(0), TableSelector::index(0), position,)?
            .is_none()
    );
    let restored = clear
        .package()
        .apply_slide_table_cell_comment(&clear.patch().inverse())?;
    assert_eq!(keynote_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn keynote_root_clear_refuses_opaque_storage_sibling() -> TestResult {
    let (source, position) = keynote_root_source()?;
    let corrupted = mutate_selected_storage(&source, StorageMutation::OpaqueSibling)?;
    assert_keynote_clear_refuses(&corrupted, position)
}

#[test]
fn keynote_root_clear_refuses_unknown_storage_payload() -> TestResult {
    let (source, position) = keynote_root_source()?;
    let corrupted = mutate_selected_storage(&source, StorageMutation::UnknownPayload)?;
    assert_keynote_clear_refuses(&corrupted, position)
}

#[test]
fn numbers_root_clear_baseline_is_successful_and_reversible() -> TestResult {
    let package = NumbersPackage::from_bytes(NUMBERS_SOURCE)?;
    let (sheet, table) = numbers_target(&package)?;
    let position = At::new(1, 1);
    let before = package
        .table_cell_comment(&sheet, &table, position)?
        .ok_or("Numbers fixture root comment is missing")?;
    let clear = package.clear_table_cell_comment(&sheet, &table, position)?;
    assert!(
        clear
            .package()
            .table_cell_comment(&sheet, &table, position)?
            .is_none()
    );
    let restored = clear
        .package()
        .apply_table_cell_comment(&clear.patch().inverse())?;
    assert_eq!(
        restored
            .package()
            .table_cell_comment(&sheet, &table, position)?,
        Some(before)
    );
    assert_eq!(numbers_bytes(restored.package())?, NUMBERS_SOURCE);
    Ok(())
}

#[test]
fn numbers_root_clear_refuses_unknown_surviving_archive_reference() -> TestResult {
    let corrupted = mutate_numbers_source(NumbersMutation::UnknownSurvivingReference)?;
    assert_numbers_clear_refuses(&corrupted, expect_numbers_unsupported)
}

#[test]
fn numbers_root_clear_refuses_selected_storage_header_mismatch() -> TestResult {
    let corrupted = mutate_numbers_source(NumbersMutation::SelectedStorageShouldMerge)?;
    assert_numbers_clear_refuses(&corrupted, expect_numbers_invalid)
}

#[test]
fn numbers_root_clear_refuses_nonempty_malformed_table_info() -> TestResult {
    let corrupted = mutate_numbers_source(NumbersMutation::MalformedTableInfo)?;
    assert_numbers_clear_refuses(&corrupted, expect_numbers_invalid)
}
