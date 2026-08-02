//! Immutable, lifetime-free RTF document facade.

use crate::document::RtfDocument;
use crate::{ParseLimits, RtfResult};
use std::fmt;
use std::io::{self, Write};
use std::path::Path;
use std::sync::{Arc, OnceLock};

struct Inner {
    model: RtfDocument<'static>,
    text: OnceLock<Box<str>>,
    paragraph_count: OnceLock<usize>,
}

/// Immutable, cheap-to-share RTF document snapshot.
///
/// Parsing always detaches the retained model from the input. Cloning a
/// `Document` therefore shares one immutable snapshot rather than duplicating
/// its strings, pictures, objects, and other resources.
///
/// Attached document state is not publicly mutable:
///
/// ```compile_fail
/// use litchi_rtf::Document;
///
/// let mut document = Document::parse(r"{\rtf1 immutable}").unwrap();
/// document.clear_fields();
/// ```
///
/// Parser storage is not part of the ordinary facade:
///
/// ```compile_fail
/// use litchi_rtf::Document;
///
/// let document = Document::parse(r"{\rtf1 semantic}").unwrap();
/// let parser_blocks = document.blocks();
/// ```
#[derive(Clone)]
pub struct Document {
    inner: Arc<Inner>,
}

impl Document {
    /// Parse UTF-8 RTF source with the production-safe resource profile.
    pub fn parse(input: &str) -> RtfResult<Self> {
        Self::parse_with_limits(input, ParseLimits::default())
    }

    /// Parse UTF-8 RTF source with an explicit finite resource profile.
    pub fn parse_with_limits(input: &str, limits: ParseLimits) -> RtfResult<Self> {
        RtfDocument::parse_with_limits(input, limits).map(Self::from_model)
    }

    /// Parse original RTF transport bytes with the production-safe resource
    /// profile.
    ///
    /// This is the preferred entry point for legacy code-page data, `bin`
    /// destinations, and compressed RTF.
    pub fn from_bytes(input: &[u8]) -> RtfResult<Self> {
        Self::from_bytes_with_limits(input, ParseLimits::default())
    }

    /// Parse original RTF transport bytes with an explicit finite resource
    /// profile.
    pub fn from_bytes_with_limits(input: &[u8], limits: ParseLimits) -> RtfResult<Self> {
        RtfDocument::parse_bytes_with_limits(input, limits).map(Self::from_model)
    }

    /// Open an RTF file with the production-safe resource profile.
    pub fn open(path: impl AsRef<Path>) -> RtfResult<Self> {
        Self::open_with_limits(path, ParseLimits::default())
    }

    /// Open an RTF file with an explicit finite resource profile.
    pub fn open_with_limits(path: impl AsRef<Path>, limits: ParseLimits) -> RtfResult<Self> {
        RtfDocument::open_with_limits(path, limits).map(Self::from_model)
    }

    fn from_model(model: RtfDocument<'static>) -> Self {
        Self {
            inner: Arc::new(Inner {
                model,
                text: OnceLock::new(),
                paragraph_count: OnceLock::new(),
            }),
        }
    }

    /// Whether two handles refer to exactly the same immutable snapshot.
    #[inline]
    pub fn same_snapshot(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.inner, &other.inner)
    }

    /// Borrow the document's plain text.
    ///
    /// The flattened text is materialized at most once per snapshot and is
    /// then shared by every clone.
    pub fn text(&self) -> &str {
        self.inner
            .text
            .get_or_init(|| self.inner.model.text().into_boxed_str())
    }

    /// Whether the body contains no text.
    pub fn is_empty(&self) -> bool {
        self.body().is_empty()
    }

    /// Number of logical body paragraphs.
    ///
    /// The value is derived at most once per snapshot.
    pub fn paragraph_count(&self) -> usize {
        *self
            .inner
            .paragraph_count
            .get_or_init(|| self.body().paragraphs().count())
    }

    /// Borrow the main text story through lazy semantic views.
    pub fn body(&self) -> crate::text::Story<'_> {
        crate::text::Story::new(
            self.inner.model.retained_blocks(),
            self.inner.model.body_boundaries(),
            self.fonts(),
            self.colors(),
        )
    }

    /// Borrow the document font resources.
    pub fn fonts(&self) -> crate::font::Catalog<'_> {
        crate::font::Catalog::new(self.inner.model.font_table())
    }

    /// Borrow the document color resources.
    pub fn colors(&self) -> crate::color::Palette<'_> {
        crate::color::Palette::new(self.inner.model.color_table())
    }

    /// Borrow body tables in source order.
    pub fn tables(&self) -> &[crate::table::Table<'_>] {
        self.inner.model.tables()
    }

    /// Borrow pictures in source order.
    pub fn pictures(&self) -> &[crate::picture::Picture<'_>] {
        self.inner.model.pictures()
    }

    /// Borrow fields in source order.
    pub fn fields(&self) -> &[crate::field::Field<'_>] {
        self.inner.model.fields()
    }

    /// Borrow sections in source order.
    pub fn sections(&self) -> &[crate::section::Section<'_>] {
        self.inner.model.sections()
    }

    /// Borrow named styles in stylesheet order.
    pub fn styles(&self) -> &[crate::style::Style<'_>] {
        self.inner.model.stylesheet().styles()
    }

    /// Borrow inert document metadata.
    pub fn info(&self) -> &crate::metadata::Info<'_> {
        self.inner.model.info()
    }

    /// Borrow annotations in source order.
    pub fn annotations(&self) -> &[crate::review::Annotation<'_>] {
        self.inner.model.annotations()
    }

    /// Borrow footnotes and endnotes in source order.
    pub fn notes(&self) -> &[crate::review::Note<'_>] {
        self.inner.model.notes()
    }

    /// Lazily traverse footnotes in source order.
    pub fn footnotes(&self) -> impl DoubleEndedIterator<Item = &crate::review::Note<'_>> + '_ {
        self.notes().iter().filter(|note| note.is_footnote)
    }

    /// Lazily traverse endnotes in source order.
    pub fn endnotes(&self) -> impl DoubleEndedIterator<Item = &crate::review::Note<'_>> + '_ {
        self.notes().iter().filter(|note| !note.is_footnote)
    }

    /// Borrow tracked revisions in source order.
    pub fn revisions(&self) -> &[crate::review::Revision<'_>] {
        self.inner.model.revisions()
    }

    /// Serialize this snapshot to a sequential sink.
    ///
    /// A sink failure can leave caller-owned output incomplete. Filesystem
    /// replacement is intentionally left to an atomic save facade.
    pub fn write_to(&self, output: impl Write) -> io::Result<()> {
        let mut writer = crate::writer::RtfWriter::new(output);
        writer.write(self)?;
        writer.flush()
    }

    /// Serialize this snapshot to a newly allocated byte buffer.
    pub fn to_bytes(&self) -> io::Result<Vec<u8>> {
        let mut output = Vec::new();
        self.write_to(&mut output)?;
        Ok(output)
    }

    pub(crate) fn model(&self) -> &RtfDocument<'static> {
        &self.inner.model
    }
}

impl fmt::Debug for Document {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Document")
            .field("paragraphs", &self.paragraph_count())
            .field("tables", &self.tables().len())
            .field("pictures", &self.pictures().len())
            .field("fields", &self.fields().len())
            .field("sections", &self.sections().len())
            .finish_non_exhaustive()
    }
}

#[cfg(test)]
mod tests {
    use super::Document;

    #[test]
    fn semantic_story_traversal_does_not_flatten_the_snapshot() {
        let document = Document::parse(r"{\rtf1 one\line two\par three}").unwrap();
        assert!(document.inner.text.get().is_none());
        assert!(document.inner.paragraph_count.get().is_none());

        let run_bytes: usize = document
            .body()
            .paragraphs()
            .flat_map(|paragraph| paragraph.runs())
            .map(|run| run.text().len())
            .sum();

        assert_eq!(run_bytes, "onetwothree".len());
        assert!(document.inner.text.get().is_none());
        assert!(document.inner.paragraph_count.get().is_none());
    }
}
