//! Immutable, lifetime-free RTF document facade.

use crate::document::RtfDocument;
use crate::{ParseLimits, RtfResult};
use std::fmt;
use std::io::{self, Write};
use std::path::Path;
use std::sync::{Arc, OnceLock};

struct Inner {
    model: RtfDocument<'static>,
    limits: ParseLimits,
    text: OnceLock<Box<str>>,
}

fn visible_text_fragment<'a>(inline: crate::text::Inline<'a>) -> Option<&'a str> {
    match inline {
        crate::text::Inline::Text(run) => Some(run.text()),
        crate::text::Inline::Break(crate::text::Break::Line) => Some("\n"),
        crate::text::Inline::Break(crate::text::Break::Paragraph) => None,
    }
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
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn parse(input: &str) -> RtfResult<Self> {
        Self::parse_with_limits(input, ParseLimits::default())
    }

    /// Parse UTF-8 RTF source with an explicit finite resource profile.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn parse_with_limits(input: &str, limits: ParseLimits) -> RtfResult<Self> {
        RtfDocument::parse_with_limits(input, limits).map(|model| Self::from_model(model, limits))
    }

    /// Parse original RTF transport bytes with the production-safe resource
    /// profile.
    ///
    /// This is the preferred entry point for legacy code-page data, `bin`
    /// destinations, and compressed RTF.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn from_bytes(input: &[u8]) -> RtfResult<Self> {
        Self::from_bytes_with_limits(input, ParseLimits::default())
    }

    /// Parse original RTF transport bytes with an explicit finite resource
    /// profile.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn from_bytes_with_limits(input: &[u8], limits: ParseLimits) -> RtfResult<Self> {
        RtfDocument::parse_bytes_with_limits(input, limits)
            .map(|model| Self::from_model(model, limits))
    }

    /// Open an RTF file with the production-safe resource profile.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn open(path: impl AsRef<Path>) -> RtfResult<Self> {
        Self::open_with_limits(path, ParseLimits::default())
    }

    /// Open an RTF file with an explicit finite resource profile.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn open_with_limits(path: impl AsRef<Path>, limits: ParseLimits) -> RtfResult<Self> {
        RtfDocument::open_with_limits(path, limits).map(|model| Self::from_model(model, limits))
    }

    fn from_model(model: RtfDocument<'static>, limits: ParseLimits) -> Self {
        Self {
            inner: Arc::new(Inner {
                model,
                limits,
                text: OnceLock::new(),
            }),
        }
    }

    /// Whether two handles refer to exactly the same immutable snapshot.
    #[inline]
    #[must_use]
    pub fn same_snapshot(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.inner, &other.inner)
    }

    /// Borrow the document's plain text.
    ///
    /// The flattened text is materialized at most once per snapshot and is
    /// then shared by every clone.
    #[must_use]
    pub fn text(&self) -> &str {
        self.inner
            .text
            .get_or_init(|| self.inner.model.text().into_boxed_str())
    }

    /// Write body paragraphs as bounded UTF-8 text to a sequential, non-seek sink.
    ///
    /// Character formatting is intentionally omitted, inline line breaks remain
    /// `\n`, and paragraph separators come from `options`. Unsupported inert
    /// destinations are retained by the document but do not become visible text.
    /// The immutable document snapshot is never changed.
    ///
    /// # Errors
    /// Returns a typed resource-limit or partial sink failure.
    pub fn write_text_to<W: Write + ?Sized>(
        &self,
        output: &mut W,
        options: litchi_core::TextOutputOptions<'_>,
    ) -> Result<litchi_core::TextOutputReport, litchi_core::TextOutputError<crate::Error>> {
        let mut writer = litchi_core::SequentialTextWriter::new(output, options);
        for paragraph in self.body().paragraphs() {
            writer.write_joined_object(
                litchi_core::TextObjectKind::Paragraph,
                || paragraph.inlines().filter_map(visible_text_fragment),
                "",
            )?;
        }
        Ok(writer.finish())
    }

    /// Whether the body contains no text.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.body().is_empty()
    }

    /// Number of logical body paragraphs.
    ///
    /// The parser retains this exact value with the immutable snapshot.
    #[must_use]
    pub fn paragraph_count(&self) -> usize {
        self.inner.model.retained_body_paragraph_count()
    }

    /// Borrow the main text story through lazy semantic views.
    #[must_use]
    pub fn body(&self) -> crate::text::Story<'_> {
        crate::text::Story::new(
            self.inner.model.retained_blocks(),
            self.inner.model.body_boundaries(),
            self.inner.model.retained_text_len(),
            self.fonts(),
            self.colors(),
        )
    }

    /// Borrow the document font resources.
    #[must_use]
    pub fn fonts(&self) -> crate::font::Catalog<'_> {
        crate::font::Catalog::new(self.inner.model.font_table())
    }

    /// Borrow the document color resources.
    #[must_use]
    pub fn colors(&self) -> crate::color::Palette<'_> {
        crate::color::Palette::new(self.inner.model.color_table())
    }

    /// Borrow body tables in source order.
    #[must_use]
    pub fn tables(&self) -> &[crate::table::Table<'_>] {
        self.inner.model.tables()
    }

    /// Borrow pictures in source order.
    #[must_use]
    pub fn pictures(&self) -> &[crate::picture::Picture<'_>] {
        self.inner.model.pictures()
    }

    /// Borrow root body shapes in source order.
    #[must_use]
    pub fn shapes(&self) -> &[crate::Shape<'_>] {
        self.inner.model.shapes()
    }

    /// Borrow fields in source order.
    #[must_use]
    pub fn fields(&self) -> &[crate::field::Field<'_>] {
        self.inner.model.fields()
    }

    /// Borrow inert embedded and linked object records in body order.
    #[must_use]
    pub fn objects(&self) -> &[crate::EmbeddedObject<'_>] {
        self.inner.model.objects()
    }

    /// Produce a bounded, content-free semantic and security inventory.
    ///
    /// The report reuses this immutable snapshot. It never reparses the
    /// source, follows or resolves references, executes fields or objects,
    /// repairs syntax, or mutates the document.
    #[must_use]
    pub fn validation_report(&self) -> crate::ValidationReport {
        crate::ValidationReport::from_document(self)
    }

    /// Alias for [`Self::validation_report`].
    #[must_use]
    pub fn security_report(&self) -> crate::ValidationReport {
        self.validation_report()
    }

    /// Borrow list definitions in list-table order.
    #[must_use]
    pub fn lists(&self) -> &[crate::list::List<'_>] {
        self.inner.model.list_table().lists()
    }

    /// Borrow list-instance overrides in override-table order.
    #[must_use]
    pub fn list_overrides(&self) -> &[crate::list::ListOverride] {
        self.inner.model.list_override_table().overrides()
    }

    /// Borrow sections in source order.
    #[must_use]
    pub fn sections(&self) -> &[crate::section::Section<'_>] {
        self.inner.model.sections()
    }

    /// Borrow named styles in stylesheet order.
    #[must_use]
    pub fn styles(&self) -> &[crate::style::Style<'_>] {
        self.inner.model.stylesheet().styles()
    }

    /// Borrow inert document metadata.
    #[must_use]
    pub fn info(&self) -> &crate::metadata::Info<'_> {
        self.inner.model.info()
    }

    /// Borrow annotations in source order.
    #[must_use]
    pub fn annotations(&self) -> &[crate::review::Annotation<'_>] {
        self.inner.model.annotations()
    }

    /// Borrow footnotes and endnotes in source order.
    #[must_use]
    pub fn notes(&self) -> &[crate::review::Note<'_>] {
        self.inner.model.notes()
    }

    /// Lazily traverse footnotes in source order.
    #[must_use]
    pub fn footnotes(&self) -> impl DoubleEndedIterator<Item = &crate::review::Note<'_>> + '_ {
        self.notes().iter().filter(|note| note.is_footnote)
    }

    /// Lazily traverse endnotes in source order.
    #[must_use]
    pub fn endnotes(&self) -> impl DoubleEndedIterator<Item = &crate::review::Note<'_>> + '_ {
        self.notes().iter().filter(|note| !note.is_footnote)
    }

    /// Borrow tracked revisions in source order.
    #[must_use]
    pub fn revisions(&self) -> &[crate::review::Revision<'_>] {
        self.inner.model.revisions()
    }

    /// Borrow unsupported syntax retained as bounded inert data.
    #[must_use]
    pub fn opaque(&self) -> &[crate::opaque::Node] {
        self.inner.model.opaque_nodes()
    }

    /// Serialize this snapshot to a sequential sink.
    ///
    /// A sink failure can leave caller-owned output incomplete. Filesystem
    /// replacement is intentionally left to an atomic save facade.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_to(&self, output: impl Write) -> io::Result<()> {
        let mut writer = crate::writer::RtfWriter::new(output);
        writer.write(self)?;
        writer.flush()
    }

    /// Serialize this snapshot to a newly allocated byte buffer.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn to_bytes(&self) -> io::Result<Vec<u8>> {
        let mut output = Vec::new();
        self.write_to(&mut output)?;
        Ok(output)
    }

    pub(crate) fn model(&self) -> &RtfDocument<'static> {
        &self.inner.model
    }

    pub(crate) fn limits(&self) -> ParseLimits {
        self.inner.limits
    }

    pub(crate) fn source_bytes(&self) -> Option<&[u8]> {
        self.inner.model.preserved_source()
    }

    pub(crate) fn source_version(&self) -> litchi_core::SourceVersion {
        let identity = Arc::as_ptr(&self.inner) as usize as u64;
        let revision = self
            .inner
            .model
            .preserved_source()
            .map_or(0, |bytes| u64::try_from(bytes.len()).unwrap_or(u64::MAX));
        litchi_core::SourceVersion::new(identity, revision)
    }

    /// Starts a bounded, detached edit of this immutable snapshot.
    #[must_use]
    pub fn edit(&self) -> crate::edit::Edit {
        crate::edit::Edit::new(self.clone())
    }

    /// Starts a bounded detached edit with caller-selected operation limits.
    #[must_use]
    pub fn edit_with_limits(&self, limits: crate::edit::Limits) -> crate::edit::Edit {
        crate::edit::Edit::new_with_limits(self.clone(), limits)
    }

    /// Starts a bounded preservation-first append before the exact root close.
    ///
    /// This transaction is distinct from [`crate::streaming::StreamingRtfWriter`]:
    /// it is rooted at this immutable existing snapshot and publishes a new
    /// artifact only after a complete source-proof, byte splice, reopen, and
    /// semantic readback.
    #[must_use]
    pub fn tail_append(
        &self,
        selector: crate::tail_append::TailSelector,
    ) -> crate::tail_append::TailAppendEdit {
        crate::tail_append::TailAppendEdit::new(self, selector)
    }

    /// Starts a bounded preservation-first append with explicit limits.
    #[must_use]
    pub fn tail_append_with_limits(
        &self,
        selector: crate::tail_append::TailSelector,
        limits: crate::tail_append::TailAppendLimits,
    ) -> crate::tail_append::TailAppendEdit {
        crate::tail_append::TailAppendEdit::with_limits(self, selector, limits)
    }

    /// Applies a shared durable RTF semantic patch to this exact snapshot.
    ///
    /// # Errors
    /// Returns an error when the format vocabulary, source digest, semantic
    /// precondition, operation bounds, or candidate validation fails.
    pub fn apply_durable<Mode>(
        &self,
        patch: &litchi_core::patch::Patch<Mode>,
    ) -> Result<Self, crate::edit::Error> {
        crate::edit::apply_durable(self, patch)
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
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
    use super::Document;
    use crate::text::{Break, Inline, Paragraph};

    fn paragraph_signature(
        paragraph: Paragraph<'_>,
    ) -> (String, Option<String>, String, Vec<String>) {
        let inlines = paragraph
            .inlines()
            .map(|inline| match inline {
                Inline::Text(run) => format!(
                    "text:{:?}:{}:{}:{:?}",
                    run.text(),
                    run.format().bold(),
                    run.format().italic(),
                    run.format().underline()
                ),
                Inline::Break(Break::Line) => "break:line".to_string(),
                Inline::Break(Break::Paragraph) => "break:paragraph".to_string(),
            })
            .collect();
        (
            paragraph.to_text(),
            paragraph.as_str().map(str::to_string),
            format!("{:?}", paragraph.format()),
            inlines,
        )
    }

    fn assert_nth_matches_repeated_next(source: &str) {
        let document = Document::parse(source).unwrap();
        let paragraphs = document.body().paragraphs().count();
        assert_eq!(document.paragraph_count(), paragraphs);
        for index in 0..=paragraphs {
            let mut repeated = document.body().paragraphs();
            let mut expected = None;
            for _ in 0..=index {
                expected = repeated.next();
                if expected.is_none() {
                    break;
                }
            }
            let mut selected = document.body().paragraphs();
            let actual = selected.nth(index);
            assert_eq!(
                actual.map(paragraph_signature),
                expected.map(paragraph_signature),
                "paragraph {index} differed for {source:?}"
            );
            assert_eq!(
                selected.next().map(paragraph_signature),
                repeated.next().map(paragraph_signature),
                "paragraph {index} did not leave the same iterator state for {source:?}"
            );
        }
        let mut exhausted = document.body().paragraphs();
        assert!(exhausted.nth(usize::MAX).is_none());
        assert!(exhausted.next().is_none());
    }

    #[test]
    fn semantic_story_traversal_does_not_flatten_the_snapshot() {
        let document = Document::parse(r"{\rtf1 one\line two\par three}").unwrap();
        assert!(document.inner.text.get().is_none());
        assert_eq!(document.paragraph_count(), 2);

        let run_bytes: usize = document
            .body()
            .paragraphs()
            .flat_map(Paragraph::runs)
            .map(|run| run.text().len())
            .sum();

        assert_eq!(run_bytes, "onetwothree".len());
        assert!(document.inner.text.get().is_none());
        assert_eq!(document.paragraph_count(), 2);
    }

    #[test]
    fn valid_paragraph_edit_preserves_retained_count() {
        let document = Document::parse(r"{\rtf1 one\par two\par three}").unwrap();
        assert_eq!(document.paragraph_count(), 3);

        document
            .edit()
            .replace_paragraph_text(1, "changed")
            .unwrap();

        assert_eq!(document.paragraph_count(), 3);
    }

    #[test]
    fn sparse_paragraph_nth_matches_repeated_next_for_structural_variants() {
        for source in [
            r"{\rtf1 first\par second\par third}",
            r"{\rtf1 first\par\par third\par}",
            r"{\rtf1 first\line wrapped\par second}",
            r"{\rtf1 first\par\line second\par third}",
            r"{\rtf1 first\par\qc plain {\b bold} {\i italic}\par\ql third}",
            r"{\rtf1 first\par decoded\u10?linefeed\par third}",
            r"{\rtf1 trailing unterminated text}",
            r"{\rtf1}",
        ] {
            assert_nth_matches_repeated_next(source);
        }
    }
}
