//! Negative proof cases for Pages and Keynote root-comment clearing.
//!
//! These cases deliberately keep the successful baseline root-only clear and
//! then alter only the selected type-3056 storage object.  An opaque sibling
//! message or an unknown payload field must produce a typed terminal refusal;
//! the package and the selected semantic comment must remain unchanged.

use litchi_iwa::keynote;
use litchi_iwa::pages;
use litchi_iwa_archive::iwa::{Archive, RawMessage, SnappyStream};
use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_common::table::coordinate::At;
use litchi_keynote::{Package as KeynotePackage, SlideSelector, TableSelector};
use litchi_pages::{BodyTableSelector, Package as PagesPackage};

use crate::harness::{TestResult, fixture};

const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
const OPAQUE_SIBLING_MESSAGE_TYPE: u32 = 99_999;
const UNKNOWN_HEADER_FIELD: u32 = 99;

#[derive(Clone, Copy)]
enum StorageMutation {
    OpaqueSibling,
    UnknownHeader,
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
                StorageMutation::UnknownHeader => {
                    litchi_iwa_common::wire::append_varint_field(
                        &mut message.data,
                        UNKNOWN_HEADER_FIELD,
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

fn pages_bytes(package: &PagesPackage) -> TestResult<Vec<u8>> {
    Ok(package.to_bytes()?)
}

fn keynote_bytes(package: &KeynotePackage) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn assert_typed_refusal(debug: String, family: &str) {
    assert!(
        debug.contains("InvalidSource")
            || debug.contains("UnsupportedDependency")
            || debug.contains("UnsupportedSource"),
        "{family} clear returned an untyped/unexpected error: {debug}"
    );
}

fn assert_pages_clear_refuses(source: &[u8], position: At) -> TestResult {
    let package = PagesPackage::from_bytes(source)?;
    let before = pages_bytes(&package)?;
    let result = package
        .edit_body_table_cell_comment(BodyTableSelector::index(0), position)?
        .clear()?
        .commit();
    match result {
        Ok(_) => panic!("Pages clear unexpectedly published a lossy candidate"),
        Err(error) => assert_typed_refusal(format!("{error:?}"), "Pages"),
    }
    assert_eq!(
        pages_bytes(&package)?,
        before,
        "Pages refusal mutated source"
    );
    assert!(
        PagesPackage::from_bytes(source)?
            .body_table_cell_comment(BodyTableSelector::index(0), position)?
            .is_some(),
        "Pages refusal lost the selected root"
    );
    Ok(())
}

fn assert_keynote_clear_refuses(source: &[u8], position: At) -> TestResult {
    let package = KeynotePackage::from_bytes(source)?;
    let before = keynote_bytes(&package)?;
    let result = package
        .edit_slide_table_cell_comment(SlideSelector::index(0), TableSelector::index(0), position)?
        .clear()?
        .commit();
    match result {
        Ok(_) => panic!("Keynote clear unexpectedly published a lossy candidate"),
        Err(error) => assert_typed_refusal(format!("{error:?}"), "Keynote"),
    }
    assert_eq!(
        keynote_bytes(&package)?,
        before,
        "Keynote refusal mutated source"
    );
    assert!(
        KeynotePackage::from_bytes(source)?
            .slide_table_cell_comment(SlideSelector::index(0), TableSelector::index(0), position,)?
            .is_some(),
        "Keynote refusal lost the selected root"
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
fn pages_root_clear_refuses_unknown_storage_header() -> TestResult {
    let (source, position) = pages_root_source()?;
    let corrupted = mutate_selected_storage(&source, StorageMutation::UnknownHeader)?;
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
fn keynote_root_clear_refuses_unknown_storage_header() -> TestResult {
    let (source, position) = keynote_root_source()?;
    let corrupted = mutate_selected_storage(&source, StorageMutation::UnknownHeader)?;
    assert_keynote_clear_refuses(&corrupted, position)
}
