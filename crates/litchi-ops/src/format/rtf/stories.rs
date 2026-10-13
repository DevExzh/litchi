//! Positional addressing of the native RTF stories behind unified nodes.
//!
//! The native owner keeps tables, header/footer stories, notes, comments,
//! fields, pictures, styles and lists in flat source-ordered collections.
//! These helpers map the unified ordinal axes onto them without building a
//! second document model.

use litchi_rtf::{HeaderFooter, HeaderFooterType, Note};

use super::Rtf;

/// Whether a native header/footer kind is a header (`true`) or footer.
pub(crate) const fn is_header(kind: HeaderFooterType) -> bool {
    matches!(
        kind,
        HeaderFooterType::Header
            | HeaderFooterType::HeaderFirst
            | HeaderFooterType::HeaderLeft
            | HeaderFooterType::HeaderRight
    )
}

/// The unified `type` token of a native header/footer kind.
pub(crate) const fn story_type(kind: HeaderFooterType) -> &'static str {
    match kind {
        HeaderFooterType::Header | HeaderFooterType::Footer => "default",
        HeaderFooterType::HeaderFirst | HeaderFooterType::FooterFirst => "first",
        HeaderFooterType::HeaderLeft | HeaderFooterType::FooterLeft => "even",
        HeaderFooterType::HeaderRight | HeaderFooterType::FooterRight => "odd",
    }
}

/// One header or footer story with its native coordinates.
#[derive(Clone, Copy)]
pub(crate) struct StoryRef<'f> {
    pub(crate) section: usize,
    pub(crate) story: usize,
    pub(crate) value: &'f HeaderFooter<'f>,
}

/// The header (`headers == true`) or footer stories in document order.
pub(crate) fn header_footers(doc: &Rtf, headers: bool) -> impl Iterator<Item = StoryRef<'_>> {
    doc.0
        .package
        .sections()
        .iter()
        .enumerate()
        .flat_map(move |(section, value)| {
            value
                .headers_footers
                .iter()
                .enumerate()
                .filter(move |(_, story)| is_header(story.header_type) == headers)
                .map(move |(story, value)| StoryRef {
                    section,
                    story,
                    value,
                })
        })
}

/// One footnote or endnote with its native index and 1-based number.
#[derive(Clone, Copy)]
pub(crate) struct NoteRef<'f> {
    pub(crate) index: usize,
    pub(crate) number: usize,
    pub(crate) value: &'f Note<'f>,
}

/// The footnotes (`footnotes == true`) or endnotes in source order.
pub(crate) fn notes(doc: &Rtf, footnotes: bool) -> impl Iterator<Item = NoteRef<'_>> {
    doc.0
        .package
        .notes()
        .iter()
        .enumerate()
        .filter(move |(_, note)| note.is_footnote == footnotes)
        .enumerate()
        .map(|(ordinal, (index, value))| NoteRef {
            index,
            number: ordinal.saturating_add(1),
            value,
        })
}

/// Indices of main-story `HYPERLINK` fields anchored in `range`.
pub(crate) fn hyperlinks(
    doc: &Rtf,
    range: std::ops::RangeInclusive<usize>,
) -> impl Iterator<Item = usize> + '_ {
    doc.0
        .package
        .fields()
        .iter()
        .enumerate()
        .filter(move |(_, field)| {
            field.owner == litchi_rtf::field::FieldOwner::Body
                && field.field_type == litchi_rtf::field::FieldType::Hyperlink
                && range.contains(&field.position)
        })
        .map(|(index, _)| index)
}

/// Indices of main-story pictures anchored in `range`.
pub(crate) fn pictures(
    doc: &Rtf,
    range: std::ops::RangeInclusive<usize>,
) -> impl Iterator<Item = usize> + '_ {
    let package = &doc.0.package;
    (0..package.pictures().len()).filter(move |index| {
        package
            .picture_body_offset(*index)
            .is_some_and(|offset| range.contains(&offset))
    })
}

/// The body-text range a paragraph's inline objects may be anchored in: its
/// text plus the position of its terminating mark.
pub(crate) fn anchor_range(
    paragraph: litchi_rtf::text::Paragraph<'_>,
) -> std::ops::RangeInclusive<usize> {
    let range = paragraph.text_range();
    range.start..=range.end
}

/// The name of the stylesheet entry of `kind` with native handle `id`.
pub(crate) fn style_name(doc: &Rtf, kind: litchi_rtf::style::Kind, id: u16) -> Option<&str> {
    doc.0
        .package
        .styles()
        .iter()
        .find(|style| style.id == id && style.style_type == kind)
        .map(|style| style.name.as_ref())
}
