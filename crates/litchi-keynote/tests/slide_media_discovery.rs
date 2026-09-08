//! Lazy semantic discovery regressions for slide media source order.
//!
//! The role-alias cases deliberately rewrite only the native slide owner.  The
//! package remains immutable, so a failed lazy projection must preserve its
//! exact source bytes.

use std::io;

use litchi_iwa_archive::{Limits, package::Catalog, package::EntryEdit};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, SnappyStream};
use litchi_iwa_protos::{kn, tsp};
use litchi_keynote::{MovieKind, Package, ReadError};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const SLIDE_MESSAGE_TYPE: u32 = 5;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const TITLE_PLACEHOLDER_FIELD: u32 = 5;
const BODY_PLACEHOLDER_FIELD: u32 = 6;
const SLIDE_NUMBER_PLACEHOLDER_FIELD: u32 = 20;

const NATIVE_BASELINE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-comments-baseline-native.key");
const NATIVE_PLACEHOLDER: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-properties-placeholder-native.key");

#[derive(Clone, Copy)]
enum PlaceholderRole {
    Title,
    Body,
    SlideNumber,
}

impl PlaceholderRole {
    const fn field(self) -> u32 {
        match self {
            Self::Title => TITLE_PLACEHOLDER_FIELD,
            Self::Body => BODY_PLACEHOLDER_FIELD,
            Self::SlideNumber => SLIDE_NUMBER_PLACEHOLDER_FIELD,
        }
    }
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn slide_component(source: &[u8]) -> TestResult<(String, Vec<u8>)> {
    for entry in Catalog::from_bytes(source)?.iter() {
        // These fixed native fixtures contain one presentation slide; template
        // slides also use type 5 and must never become the test mutation target.
        if !entry.name().starts_with("Index/Slide-") || !entry.name().ends_with(".iwa") {
            continue;
        }
        let stream = SnappyStream::decompress(entry.data())?.into_bytes();
        let archive = Archive::parse(&stream)?;
        if archive.objects.iter().any(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == SLIDE_MESSAGE_TYPE)
        }) {
            return Ok((entry.name().to_owned(), stream));
        }
    }
    Err(io::Error::other("missing native slide component").into())
}

fn rewrite_document(
    source: &[u8],
    mutate: impl FnOnce(&mut Archive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let (member, stream) = slide_component(source)?;
    let mut archive = Archive::parse(&stream)?;
    mutate(&mut archive)?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    let catalog = Catalog::from_bytes(source)?;
    Ok(catalog.reassemble_to_bytes(&[EntryEdit::new(&member, &compressed)], Limits::default())?)
}

fn slide_object(archive: &Archive) -> TestResult<&ArchiveObject> {
    archive
        .objects
        .iter()
        .find(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == SLIDE_MESSAGE_TYPE)
        })
        .ok_or_else(|| io::Error::other("missing native slide object").into())
}

fn slide_object_mut(archive: &mut Archive) -> TestResult<&mut ArchiveObject> {
    archive
        .objects
        .iter_mut()
        .find(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == SLIDE_MESSAGE_TYPE)
        })
        .ok_or_else(|| io::Error::other("missing native slide object").into())
}

fn first_movie_identifier(archive: &Archive, slide: &kn::SlideArchive) -> TestResult<u64> {
    slide
        .owned_drawables
        .iter()
        .find_map(|reference| {
            archive.object(reference.identifier).and_then(|object| {
                object
                    .messages
                    .iter()
                    .any(|message| message.type_ == MOVIE_MESSAGE_TYPE)
                    .then_some(reference.identifier)
            })
        })
        .ok_or_else(|| io::Error::other("native slide has no movie drawable").into())
}

fn with_movie_role_alias(source: &[u8], role: PlaceholderRole) -> TestResult<Vec<u8>> {
    rewrite_document(source, |archive| {
        let slide = slide_object(archive)?;
        let slide_message = slide
            .messages
            .iter()
            .find(|message| message.type_ == SLIDE_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing native slide message"))?;
        let slide_archive = kn::SlideArchive::decode(slide_message.data.as_slice())?;
        let movie_identifier = first_movie_identifier(archive, &slide_archive)?;

        let slide = slide_object_mut(archive)?;
        let message_index = slide
            .messages
            .iter()
            .position(|message| message.type_ == SLIDE_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing native slide message"))?;
        let mut replacement =
            kn::SlideArchive::decode(slide.messages[message_index].data.as_slice())?;
        let previous_identifier = match role {
            PlaceholderRole::Title => replacement
                .title_placeholder
                .as_ref()
                .map(|reference| reference.identifier),
            PlaceholderRole::Body => replacement
                .body_placeholder
                .as_ref()
                .map(|reference| reference.identifier),
            PlaceholderRole::SlideNumber => replacement
                .slide_number_placeholder
                .as_ref()
                .map(|reference| reference.identifier),
        };
        let movie_reference = tsp::Reference {
            identifier: movie_identifier,
            ..Default::default()
        };
        match role {
            PlaceholderRole::Title => replacement.title_placeholder = Some(movie_reference),
            PlaceholderRole::Body => replacement.body_placeholder = Some(movie_reference),
            PlaceholderRole::SlideNumber => {
                replacement.slide_number_placeholder = Some(movie_reference)
            },
        }
        let data = replacement.encode_to_vec();
        slide.messages[message_index].data = data;
        slide.archive_info.message_infos[message_index].length =
            u32::try_from(slide.messages[message_index].data.len())?;

        let info = &mut slide.archive_info.message_infos[message_index];
        if let Some(previous_identifier) = previous_identifier {
            let reference = info
                .object_references
                .iter_mut()
                .find(|identifier| **identifier == previous_identifier)
                .ok_or_else(|| io::Error::other("native role reference metadata is missing"))?;
            *reference = movie_identifier;
        } else {
            info.object_references.push(movie_identifier);
        }
        if let Some(field) = info
            .field_infos
            .iter_mut()
            .find(|field| field.path.as_slice() == [role.field()])
        {
            if let Some(previous_identifier) = previous_identifier {
                let reference = field
                    .object_references
                    .iter_mut()
                    .find(|identifier| **identifier == previous_identifier)
                    .ok_or_else(|| io::Error::other("native role field metadata is missing"))?;
                *reference = movie_identifier;
            } else {
                field.object_references.push(movie_identifier);
            }
        } else {
            let mut field = FieldInfo::new(vec![role.field()]);
            field.object_references.push(movie_identifier);
            info.field_infos.push(field);
        }
        Ok(())
    })
}

#[test]
fn native_media_projection_retains_complete_source_order() -> TestResult {
    for (source, expected) in [
        (
            NATIVE_BASELINE,
            vec![
                MovieKind::Audio,
                MovieKind::Audio,
                MovieKind::File,
                MovieKind::File,
            ],
        ),
        (
            NATIVE_PLACEHOLDER,
            vec![
                MovieKind::Audio,
                MovieKind::Audio,
                MovieKind::File,
                MovieKind::File,
                MovieKind::Placeholder,
            ],
        ),
    ] {
        let package = Package::from_bytes(source)?;
        let before = exact_bytes(&package)?;
        let movies = package.show()?.slides()[0].movies();
        assert_eq!(
            movies.iter().map(|movie| movie.kind()).collect::<Vec<_>>(),
            expected
        );
        assert_eq!(exact_bytes(&package)?, before);
    }
    Ok(())
}

#[test]
fn native_movie_role_aliases_are_rejected_by_lazy_slide_projection() -> TestResult {
    for role in [
        PlaceholderRole::Title,
        PlaceholderRole::Body,
        PlaceholderRole::SlideNumber,
    ] {
        let hostile = with_movie_role_alias(NATIVE_BASELINE, role)?;
        let package = Package::from_bytes(&hostile)?;
        let before = exact_bytes(&package)?;
        let error = package
            .slides()
            .expect_err("movie role alias must be rejected");
        if matches!(role, PlaceholderRole::SlideNumber) {
            assert!(matches!(
                error,
                ReadError::InvalidFormat(message)
                    if message.contains("aliases the slide slide-number placeholder")
            ));
        } else {
            assert!(matches!(error, ReadError::InvalidFormat(_)), "{error:?}");
        }
        assert_eq!(exact_bytes(&package)?, before);
    }
    Ok(())
}
