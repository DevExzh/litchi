//! The paragraph style of a new caption against a caption Numbers authored
//! (record 0435).
//!
//! `chart-caption-native.numbers` is `basic.numbers` with a chart and a caption
//! typed in Numbers 14.4. Its caption uses the template's caption style, which
//! the theme lists after the title's. A new caption takes the same object; a
//! new title keeps the title's style.

use litchi_iwa_archive::Limits;
use litchi_iwa_archive::iwa::{Archive, SnappyStream};
use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_common::WireLimits;
use litchi_iwa_common::wire::WireView;
use litchi_numbers::Package;
use litchi_numbers::sheet::drawable::caption::Label;
use litchi_numbers::sheet::drawable::{DrawableSelector, Kind};

const NATIVE_CAPTION: &[u8] =
    include_bytes!("../../../test-data/iwork/numbers/chart-caption-native.numbers");
const NATIVE_CHART: &[u8] =
    include_bytes!("../../../test-data/iwork/numbers/chart-data-native.numbers");
const NATIVE_IMAGES: &[u8] =
    include_bytes!("../../../test-data/iwork/numbers/image-adjustments-native.numbers");
const LITCHI_BUILT: &[u8] =
    include_bytes!("../../../test-data/iwork/numbers/chart-grid-authored-native.numbers");
const CAPTION_IDENTIFIER: &str = "captions-0-paragraphstyle-Caption Text";
const TITLE_IDENTIFIER: &str = "captions-0-paragraphstyle-Caption Title";
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const STYLESHEET_MEMBER: &str = "Index/DocumentStylesheet.iwa";

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

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

fn field(payload: &[u8], number: u32) -> TestResult<Option<&[u8]>> {
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

/// `(object, style_identifier, name)` of the paragraph style at the start of
/// the text storage holding exactly `text`.
fn paragraph_style(bytes: &[u8], text: &str) -> TestResult<(u64, String, String)> {
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
    let base = path(&style.messages[0].data, &[1])?;
    Ok((
        identifier,
        std::str::from_utf8(path(base, &[2])?)?.to_owned(),
        std::str::from_utf8(path(base, &[1])?)?.to_owned(),
    ))
}

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    Ok(output)
}

fn chart() -> DrawableSelector {
    DrawableSelector::of(Kind::Chart, 0)
}

fn image() -> DrawableSelector {
    DrawableSelector::of(Kind::Image, 0)
}

/// `source` with the caption style's identifier changed, as in a template
/// that has no such style.
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

/// The native file records what a caption uses, and the leaf writes the same.
#[test]
fn new_chart_caption_equals_the_native_paragraph_style() -> TestResult {
    let native = paragraph_style(NATIVE_CAPTION, "Native caption")?;
    assert_eq!(
        native,
        (
            904_466,
            CAPTION_IDENTIFIER.to_owned(),
            "Object Caption".to_owned()
        )
    );
    let package = Package::from_bytes(NATIVE_CHART)?;
    let mut edit = package.edit_sheet_drawable_caption(0, chart(), Label::Caption)?;
    edit.set("Native caption")?;
    let written = bytes(edit.commit()?.package())?;
    assert_eq!(paragraph_style(&written, "Native caption")?, native);
    Ok(())
}

#[test]
fn new_inline_labels_use_the_style_of_their_kind() -> TestResult {
    let package = Package::from_bytes(NATIVE_IMAGES)?;
    let mut edit = package.edit_sheet_drawable_caption(0, image(), Label::Caption)?;
    edit.set("A caption")?;
    let captioned = edit.commit()?;
    let written = bytes(captioned.package())?;
    assert_eq!(
        paragraph_style(&written, "A caption")?.1,
        CAPTION_IDENTIFIER
    );

    let mut edit = captioned
        .package()
        .edit_sheet_drawable_caption(0, image(), Label::Title)?;
    edit.set("A title")?;
    let written = bytes(edit.commit()?.package())?;
    let title = paragraph_style(&written, "A title")?;
    assert_eq!(title.1, TITLE_IDENTIFIER);
    assert_eq!(title.2, "Object Title");
    assert_eq!(
        paragraph_style(&written, "A caption")?.1,
        CAPTION_IDENTIFIER
    );
    Ok(())
}

#[test]
fn replacing_and_removing_a_new_caption_keep_its_style_and_invert() -> TestResult {
    let package = Package::from_bytes(NATIVE_CHART)?;
    let mut edit = package.edit_sheet_drawable_caption(0, chart(), Label::Caption)?;
    edit.set("First")?;
    let created = edit.commit()?;
    let created_bytes = bytes(created.package())?;

    let mut edit = created
        .package()
        .edit_sheet_drawable_caption(0, chart(), Label::Caption)?;
    edit.set("Second")?;
    let replaced = edit.commit()?;
    assert_eq!(
        paragraph_style(&bytes(replaced.package())?, "Second")?.1,
        CAPTION_IDENTIFIER
    );

    let mut edit = replaced
        .package()
        .edit_sheet_drawable_caption(0, chart(), Label::Caption)?;
    edit.clear()?;
    let cleared = edit.commit()?;
    assert_eq!(
        cleared
            .package()
            .sheet_drawable_caption(0, chart(), Label::Caption)?,
        None
    );

    // A new label after a removal authors the same style again.
    let mut edit = cleared
        .package()
        .edit_sheet_drawable_caption(0, chart(), Label::Caption)?;
    edit.set("Third")?;
    assert_eq!(
        paragraph_style(&bytes(edit.commit()?.package())?, "Third")?.1,
        CAPTION_IDENTIFIER
    );

    // Every step inverts to the exact bytes before it.
    let restored = cleared
        .package()
        .apply_sheet_drawable_caption(&cleared.patch().inverse())?;
    assert_eq!(
        bytes(restored.package())?,
        bytes(replaced.package())?,
        "the removal's inverse"
    );
    let restored = replaced
        .package()
        .apply_sheet_drawable_caption(&replaced.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, created_bytes);
    let restored = created
        .package()
        .apply_sheet_drawable_caption(&created.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, NATIVE_CHART);
    Ok(())
}

/// Replacing Numbers' own caption changes its text, not its style.
#[test]
fn replacing_the_native_caption_keeps_the_native_style() -> TestResult {
    let package = Package::from_bytes(NATIVE_CAPTION)?;
    let mut edit = package.edit_sheet_drawable_caption(0, chart(), Label::Caption)?;
    edit.set("Replaced")?;
    let written = bytes(edit.commit()?.package())?;
    assert_eq!(paragraph_style(&written, "Replaced")?.0, 904_466);
    Ok(())
}

/// A template that lists no such style keeps the earlier behaviour: the first
/// preset. Documents Litchi built list one style for both labels.
#[test]
fn a_template_without_the_caption_style_falls_back_to_the_first_preset() -> TestResult {
    let package = Package::from_bytes(&without_caption_style(NATIVE_IMAGES)?)?;
    let mut edit = package.edit_sheet_drawable_caption(0, image(), Label::Caption)?;
    edit.set("A caption")?;
    assert_eq!(
        paragraph_style(&bytes(edit.commit()?.package())?, "A caption")?.1,
        TITLE_IDENTIFIER
    );

    let package = Package::from_bytes(LITCHI_BUILT)?;
    let mut edit = package.edit_sheet_drawable_caption(0, chart(), Label::Caption)?;
    edit.set("A caption")?;
    assert_eq!(
        paragraph_style(&bytes(edit.commit()?.package())?, "A caption")?.1,
        "litchi-paragraph-caption"
    );
    Ok(())
}
