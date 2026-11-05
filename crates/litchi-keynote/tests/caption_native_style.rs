//! The paragraph style of a new caption against captions Keynote authored
//! (record 0435).
//!
//! `chart-caption-native.key` and `slide-movie-geometry-source-native.key` hold
//! captions typed in Keynote 14.4, on a chart and on a movie. Both use the
//! template's caption style, which the theme lists after the title's. A new
//! caption takes it; a new title keeps the title's style.

use std::path::PathBuf;

use litchi_iwa_archive::Limits;
use litchi_iwa_archive::iwa::{Archive, SnappyStream};
use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_common::WireLimits;
use litchi_iwa_common::wire::WireView;
use litchi_keynote::chart::caption::{Error, Refusal};
use litchi_keynote::slide::image::text::Text;
use litchi_keynote::{ChartSelector, ImageSelector, Package, SlideSelector};

const CAPTION_IDENTIFIER: &str = "captions-0-paragraphstyle-Caption Text";
const TITLE_IDENTIFIER: &str = "captions-0-paragraphstyle-Caption Title";
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const STYLESHEET_MEMBER: &str = "Index/DocumentStylesheet.iwa";

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn fixture(name: &str) -> TestResult<Vec<u8>> {
    Ok(std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/keynote")
            .join(name),
    )?)
}

fn objects(bytes: &[u8]) -> TestResult<Vec<litchi_iwa_core::ArchiveObject>> {
    let catalog = Catalog::from_bytes(bytes, Limits::default())?;
    let mut objects = Vec::new();
    for entry in catalog.iter() {
        if entry.name().ends_with(".iwa") {
            let stream = SnappyStream::decompress(entry.data(), Default::default())?;
            objects.extend(Archive::parse(stream.as_bytes(), Default::default())?.objects);
        }
    }
    Ok(objects)
}

fn field<'a>(payload: &'a [u8], number: u32) -> TestResult<Option<&'a [u8]>> {
    let view = WireView::parse(payload, WireLimits::DEFAULT)?;
    Ok(view
        .fields()
        .find(|field| field.number() == number)
        .map(|field| field.payload()))
}

fn path<'a>(mut payload: &'a [u8], numbers: &[u32]) -> TestResult<&'a [u8]> {
    for number in numbers {
        payload = field(payload, *number)?.ok_or("a missing field")?;
    }
    Ok(payload)
}

/// `(object, style_identifier)` of the paragraph style at the start of the
/// text storage holding exactly `text`.
fn paragraph_style(bytes: &[u8], text: &str) -> TestResult<(u64, String)> {
    let objects = objects(bytes)?;
    let storage = objects
        .iter()
        .flat_map(|object| &object.messages)
        .filter(|message| message.type_ == STORAGE_MESSAGE_TYPE)
        .find(|message| {
            field(&message.data, 3)
                .ok()
                .flatten()
                .is_some_and(|stored| stored == text.as_bytes())
        })
        .ok_or("no storage holds the text")?;
    let reference = path(&storage.data, &[5, 1, 2, 1])?;
    let (identifier, _) = litchi_iwa_common::decode_varint_from_bytes(reference)?;
    let style = objects
        .iter()
        .find(|object| object.archive_info.identifier == Some(identifier))
        .ok_or("the paragraph style is missing")?;
    Ok((
        identifier,
        std::str::from_utf8(path(&style.messages[0].data, &[1, 2])?)?.to_owned(),
    ))
}

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    Ok(output)
}

fn slide() -> SlideSelector<'static> {
    SlideSelector::index(0)
}

fn without_caption_style(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source, Limits::default())?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == STYLESHEET_MEMBER)
        .ok_or("no stylesheet")?;
    let mut stream = SnappyStream::decompress(entry.data(), Default::default())?.into_bytes();
    let needle = CAPTION_IDENTIFIER.as_bytes();
    // The stylesheet's map and the style object both hold it.
    let places = stream
        .windows(needle.len())
        .enumerate()
        .filter_map(|(at, window)| (window == needle).then_some(at))
        .collect::<Vec<_>>();
    assert_eq!(places.len(), 2);
    for at in places {
        stream[at + needle.len() - 4] = b'N';
    }
    let compressed = SnappyStream::compress(&stream, Default::default())?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(STYLESHEET_MEMBER, &compressed)],
        Limits::default(),
    )?)
}

/// Keynote's own captions, on a chart and a movie, name the template's caption
/// style. A new image caption takes the same style.
#[test]
fn new_image_caption_equals_the_native_paragraph_style() -> TestResult {
    let chart = paragraph_style(
        &fixture("chart-caption-native.key")?,
        "Native chart caption marker",
    )?;
    let movie = paragraph_style(
        &fixture("slide-movie-geometry-source-native.key")?,
        "Shared native movie caption",
    )?;
    assert_eq!(chart, movie);
    assert_eq!(chart.1, CAPTION_IDENTIFIER);

    let package = Package::from_bytes(&fixture("image-adjustments-native.key")?)?;
    let captioned = package
        .edit_slide_image_text(slide(), ImageSelector::index(0), Text::Caption)?
        .set("A caption")?
        .commit()?;
    assert_eq!(
        paragraph_style(&bytes(captioned.package())?, "A caption")?,
        chart
    );
    Ok(())
}

/// A new caption on a native chart stays refused, as it was: the slide does
/// not list the document stylesheet, which the route requires (record 0435).
#[test]
fn a_new_caption_on_a_native_chart_stays_refused() -> TestResult {
    let package = Package::from_bytes(&fixture("chart-caption-native.key")?)?;
    let cleared = package
        .edit_slide_chart_caption(slide(), ChartSelector::index(0))?
        .clear()?
        .commit()?;
    let refused = cleared
        .package()
        .edit_slide_chart_caption(slide(), ChartSelector::index(0))?
        .set("New")?
        .commit();
    assert!(matches!(
        refused,
        Err(Error::Refused(Refusal::UnsupportedDependency))
    ));
    Ok(())
}

#[test]
fn new_image_labels_use_the_style_of_their_kind_and_invert() -> TestResult {
    let source = fixture("image-adjustments-native.key")?;
    let package = Package::from_bytes(&source)?;
    let captioned = package
        .edit_slide_image_text(slide(), ImageSelector::index(0), Text::Caption)?
        .set("A caption")?
        .commit()?;
    let captioned_bytes = bytes(captioned.package())?;
    assert_eq!(
        paragraph_style(&captioned_bytes, "A caption")?.1,
        CAPTION_IDENTIFIER
    );
    let reopened = Package::from_bytes(&captioned_bytes)?;
    assert_eq!(
        reopened
            .slide_image_text(slide(), ImageSelector::index(0), Text::Caption)?
            .as_deref(),
        Some("A caption")
    );

    let replaced = captioned
        .package()
        .edit_slide_image_text(slide(), ImageSelector::index(0), Text::Caption)?
        .set("Replaced")?
        .commit()?;
    assert_eq!(
        paragraph_style(&bytes(replaced.package())?, "Replaced")?.1,
        CAPTION_IDENTIFIER
    );
    let cleared = captioned
        .package()
        .edit_slide_image_text(slide(), ImageSelector::index(0), Text::Caption)?
        .clear()?
        .commit()?;
    assert_eq!(
        cleared
            .package()
            .slide_image_text(slide(), ImageSelector::index(0), Text::Caption)?,
        None
    );

    let restored = captioned
        .package()
        .apply_slide_image_text(&captioned.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);

    let titled = package
        .edit_slide_image_text(slide(), ImageSelector::index(0), Text::Title)?
        .set("A title")?
        .commit()?;
    assert_eq!(
        paragraph_style(&bytes(titled.package())?, "A title")?.1,
        TITLE_IDENTIFIER
    );
    Ok(())
}

/// A template that lists no such style keeps the earlier behaviour: the first
/// preset.
#[test]
fn a_template_without_the_caption_style_falls_back_to_the_first_preset() -> TestResult {
    let source = without_caption_style(&fixture("image-adjustments-native.key")?)?;
    let package = Package::from_bytes(&source)?;
    let captioned = package
        .edit_slide_image_text(slide(), ImageSelector::index(0), Text::Caption)?
        .set("A caption")?
        .commit()?;
    assert_eq!(
        paragraph_style(&bytes(captioned.package())?, "A caption")?.1,
        TITLE_IDENTIFIER
    );
    Ok(())
}
