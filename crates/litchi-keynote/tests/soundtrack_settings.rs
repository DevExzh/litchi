//! Native absence and archive-free validation for soundtrack settings.

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::wire::WireView;
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_keynote::{
    Package,
    soundtrack::{Error, Mode, Settings},
};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn fixture() -> std::path::PathBuf {
    std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/keynote/basic.key")
}

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut result = Vec::new();
    package.write_to(&mut result)?;
    Ok(result)
}

fn without_soundtrack_reference(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let mut archive = Archive::parse(stream.as_bytes())?;
        for object in &mut archive.objects {
            let Some(index) = object
                .messages
                .iter()
                .position(|message| message.type_ == 2)
            else {
                continue;
            };
            let view = WireView::parse(&object.messages[index].data)?;
            if !view.fields().any(|field| field.number() == 17) {
                continue;
            }
            let mut rewritten = Vec::with_capacity(object.messages[index].data.len());
            for field in view.fields() {
                if field.number() != 17 {
                    rewritten.extend_from_slice(field.raw());
                }
            }
            object.replace_message_preserving_header(
                index,
                RawMessage {
                    type_: 2,
                    data: rewritten,
                },
            )?;
            let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
            return Ok(catalog.reassemble_to_bytes(
                &[EntryEdit::new(entry.name(), &compressed)],
                Limits::default(),
            )?);
        }
    }
    Err("native show soundtrack reference was not found".into())
}

fn soundtrack_media_fields(source: &[u8]) -> TestResult<Vec<Vec<u8>>> {
    let catalog = Catalog::from_bytes(source)?;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let archive = Archive::parse(stream.as_bytes())?;
        for object in &archive.objects {
            if let Some(message) = object.messages.iter().find(|message| message.type_ == 21) {
                let view = WireView::parse(&message.data)?;
                return Ok(view
                    .fields()
                    .filter(|field| field.number() == 3)
                    .map(|field| field.raw().to_vec())
                    .collect());
            }
        }
    }
    Err("native soundtrack payload was not found".into())
}

#[test]
fn native_read_noop_change_apply_inverse_and_conflict() -> TestResult<()> {
    let package = Package::open(fixture())?;
    let source = bytes(&package)?;
    let before = package.soundtrack_settings()?.ok_or("missing soundtrack")?;
    let noop = package.edit_soundtrack_settings()?.set(before).commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(bytes(noop.package())?, source);
    let changed = Settings::new(Some(0.5), Some(Mode::Loop))?;
    let commit = package.edit_soundtrack_settings()?.set(changed).commit()?;
    assert_eq!(commit.package().soundtrack_settings()?, Some(changed));
    let target = bytes(commit.package())?;
    let applied = package.apply_soundtrack_settings(commit.patch())?;
    assert_eq!(bytes(applied.package())?, target);
    let restored = commit
        .package()
        .apply_soundtrack_settings(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    assert!(matches!(
        commit.package().apply_soundtrack_settings(commit.patch()),
        Err(Error::PatchConflict)
    ));
    Ok(())
}

#[test]
fn archive_free_settings_validate_presence_future_modes_and_redaction() -> TestResult<()> {
    let settings = Settings::new(Some(1.0), Some(Mode::Unknown(19)))?;
    assert_eq!(settings.volume(), Some(1.0));
    assert_eq!(settings.mode(), Some(Mode::Unknown(19)));
    assert!(Settings::new(Some(f64::NAN), None).is_err());
    assert!(Settings::new(Some(-0.1), None).is_err());
    assert!(Settings::new(Some(1.1), None).is_err());
    assert!(Settings::new(None, Some(Mode::Unknown(0))).is_err());
    assert!(!format!("{settings:?}").contains("Index/"));
    Ok(())
}

#[test]
fn synthetic_absence_reads_none_and_edit_is_typed() -> TestResult<()> {
    let native = Package::open(fixture())?;
    let absent = without_soundtrack_reference(&bytes(&native)?)?;
    let package = Package::from_bytes(&absent)?;
    assert_eq!(package.soundtrack_settings()?, None);
    assert!(matches!(
        package.edit_soundtrack_settings(),
        Err(Error::SoundtrackNotFound)
    ));
    assert_eq!(bytes(&package)?, absent);
    Ok(())
}

#[test]
fn native_soundtrack_media_records_are_preserved_by_settings_change_and_inverse() -> TestResult<()>
{
    let package = Package::open(fixture())?;
    let source = bytes(&package)?;
    let media = soundtrack_media_fields(&source)?;
    let before = package.soundtrack_settings()?.ok_or("missing soundtrack")?;
    let changed = Settings::new(Some(0.25), Some(Mode::DoNotPlay))?;
    let commit = package.edit_soundtrack_settings()?.set(changed).commit()?;
    let target = bytes(commit.package())?;
    assert_eq!(soundtrack_media_fields(&target)?, media);
    let restored = commit
        .package()
        .apply_soundtrack_settings(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    assert_eq!(restored.package().soundtrack_settings()?, Some(before));
    Ok(())
}
