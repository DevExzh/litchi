//! Immutable, lifetime-free RTF document facade.

use crate::document::Tree;
use crate::{OpenOptions, SourceLimits};
use std::fmt;
use std::io::{self, Write};
use std::path::Path;
use std::sync::{Arc, OnceLock};

struct Inner {
    model: Tree<'static>,
    options: OpenOptions,
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
/// `Package` therefore shares one immutable snapshot rather than duplicating
/// its strings, pictures, objects, and other resources.
///
/// Attached document state is not publicly mutable:
///
/// ```compile_fail
/// use litchi_rtf::Package;
///
/// let mut document = Package::parse(r"{\rtf1 immutable}").unwrap();
/// document.clear_fields();
/// ```
///
/// Parser storage is not part of the ordinary facade:
///
/// ```compile_fail
/// use litchi_rtf::Package;
///
/// let document = Package::parse(r"{\rtf1 semantic}").unwrap();
/// let parser_blocks = document.blocks();
/// ```
#[derive(Clone)]
pub struct Package {
    inner: Arc<Inner>,
}

impl Package {
    /// Construct the minimal valid empty RTF package.
    ///
    /// The package is produced through the canonical [`crate::Builder`] path,
    /// so it has the same strict source policy and native writer semantics as
    /// an explicitly authored package.
    ///
    /// # Errors
    /// Returns a typed authoring or strict-reopen error.
    pub fn new() -> crate::builder::Result<Self> {
        crate::Builder::default().build()
    }

    /// Obtain an owned read-only semantic projection of this snapshot.
    #[must_use]
    pub fn document(&self) -> Document {
        Document {
            owner: self.clone(),
        }
    }

    /// The checked source profile retained by this package.
    pub fn source_limits(&self) -> SourceLimits {
        self.inner.options.source_limits_value()
    }

    /// The open policy retained for edits and candidate admission.
    pub fn open_options(&self) -> OpenOptions {
        self.inner.options
    }

    /// Borrow the advanced retained model without mutable access.
    #[must_use]
    pub fn raw(&self) -> &Tree<'static> {
        &self.inner.model
    }

    /// Parse UTF-8 RTF source with the production-safe resource profile.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn parse(input: &str) -> crate::Result<Self> {
        Self::parse_with_options(input, OpenOptions::default())
    }

    /// Parse UTF-8 RTF source with explicit resource limits and repair
    /// posture.
    ///
    /// Word-compatible repairs are applied under every leniency; the opt-in
    /// repairs a [`Leniency`](crate::leniency::Leniency) variant names are
    /// applied only when selected. Every repair is listed by
    /// [`Self::tolerance_report`].
    ///
    /// # Errors
    /// Returns an error when the input is malformed beyond what the selected
    /// leniency repairs, or a configured limit is exceeded.
    pub fn parse_with_options(input: &str, options: OpenOptions) -> crate::Result<Self> {
        Tree::parse(input, options).map(|model| Self::from_model(model, options))
    }

    /// Parse original RTF transport bytes with the production-safe resource
    /// profile.
    ///
    /// This is the preferred entry point for legacy code-page data, `bin`
    /// destinations, and compressed RTF.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn from_bytes(input: &[u8]) -> crate::Result<Self> {
        Self::from_bytes_with_options(input, OpenOptions::default())
    }

    /// Parse original RTF transport bytes with explicit resource limits and
    /// repair posture.
    ///
    /// # Errors
    /// Returns an error when the input is malformed beyond what the selected
    /// leniency repairs, or a configured limit is exceeded.
    pub fn from_bytes_with_options(input: &[u8], options: OpenOptions) -> crate::Result<Self> {
        Tree::parse_bytes(input, options).map(|model| Self::from_model(model, options))
    }

    /// Parse an owned RTF transport allocation and retain it without copying.
    ///
    /// This is the consuming counterpart to [`Self::from_bytes`]. It is
    /// useful when the input already belongs to the caller and must remain
    /// available for exact-source serialization.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is
    /// exceeded.
    pub fn from_owned_bytes(input: Vec<u8>) -> crate::Result<Self> {
        Self::from_owned_bytes_with_options(input, OpenOptions::default())
    }

    /// Parse an owned RTF transport allocation with explicit resource limits
    /// and repair posture.
    ///
    /// # Errors
    /// Returns an error when the input is malformed beyond what the selected
    /// leniency repairs, or a configured limit is exceeded.
    pub fn from_owned_bytes_with_options(
        input: Vec<u8>,
        options: OpenOptions,
    ) -> crate::Result<Self> {
        Tree::from_owned_bytes(input, options).map(|model| Self::from_model(model, options))
    }

    /// Open an RTF file with the production-safe resource profile.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn open(path: impl AsRef<Path>) -> crate::Result<Self> {
        Self::open_with_options(path, OpenOptions::default())
    }

    /// Open an RTF file with explicit resource limits and repair posture.
    ///
    /// # Errors
    /// Returns an error when the input is malformed beyond what the selected
    /// leniency repairs, or a configured limit is exceeded.
    pub fn open_with_options(path: impl AsRef<Path>, options: OpenOptions) -> crate::Result<Self> {
        Tree::open(path, options).map(|model| Self::from_model(model, options))
    }

    fn from_model(model: Tree<'static>, options: OpenOptions) -> Self {
        Self {
            inner: Arc::new(Inner {
                model,
                options,
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

    /// Every repair the parse applied, in discovery order.
    ///
    /// An empty report means the source conformed and nothing was
    /// substituted. Word-compatible repairs are recorded under every
    /// [`Leniency`](crate::leniency::Leniency), including the default
    /// strict posture.
    #[inline]
    #[must_use]
    pub fn tolerance_report(&self) -> &crate::leniency::ToleranceReport {
        self.inner.model.tolerance_report()
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

    /// UTF-8 offset in [`Self::text`] where the picture at `index` is anchored.
    ///
    /// Only pictures opened directly in the main story have an anchor;
    /// pictures inside table cells, other stories, embedded-object results,
    /// shapes or compatibility wrappers return `None`.
    #[must_use]
    pub fn picture_body_offset(&self, index: usize) -> Option<usize> {
        self.inner.model.text_provenance().picture(index)
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

    /// Borrow inert SmartTag/factoid ranges in body source order.
    #[must_use]
    pub fn smart_tags(&self) -> &[crate::SmartTag<'_>] {
        self.inner.model.smart_tags()
    }

    /// Borrow complete inert main-body tracked-move start/end ranges in source order.
    ///
    /// Equal tags with opposite kinds identify the two move locations, but the
    /// bounded model does not expose a cross-location pair index or execute
    /// move/deleted/inserted semantics.
    #[must_use]
    pub fn move_bookmarks(&self) -> &[crate::MoveBookmark<'_>] {
        self.inner.model.move_bookmarks()
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
        crate::ValidationReport::from_package(self)
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

    /// Borrow inert top-level `nextfile` and `template` metadata without
    /// resolving either name.
    #[must_use]
    pub fn external_references(&self) -> &crate::DocumentExternalReferences<'_> {
        self.inner.model.external_references()
    }

    /// Capture the bounded source inventory used by the narrow external-
    /// reference redaction API.  The inventory is non-mutating and refuses no
    /// content merely because it is unsupported; strict planning reports those
    /// surfaces explicitly.
    pub fn external_reference_redaction_snapshot(
        &self,
    ) -> Result<crate::redact::Snapshot, crate::edit::Error> {
        crate::redact::Snapshot::from_package(self)
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

    /// Atomically save this package with full durability.
    ///
    /// Returns the number of serialized bytes. A failure before replacement
    /// leaves the destination untouched; a committed error means replacement
    /// succeeded but subsequent durability work failed.
    ///
    /// # Errors
    /// Returns typed publication or serialization errors.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<u64, crate::save::Error> {
        self.save_with_options(path, crate::save::Options::new())
    }

    /// Atomically save with explicit per-call durability.
    ///
    /// # Errors
    /// Returns the same typed failures as [`Self::save`].
    pub fn save_with_durability(
        &self,
        path: impl AsRef<Path>,
        durability: litchi_core::Durability,
    ) -> Result<u64, crate::save::Error> {
        self.save_with_options(path, crate::save::Options::new().durability(durability))
    }

    /// Atomically save with explicit durability and overwrite policies.
    ///
    /// Symbolic-link and non-file destinations are refused. The complete
    /// serialized artifact is staged in a sibling temporary file before
    /// publication. An immutable package's stored opening policy is unchanged.
    ///
    /// # Errors
    /// Returns typed publication or serialization errors. A committed failure
    /// must not be retried as though the destination were unchanged.
    pub fn save_with_options(
        &self,
        path: impl AsRef<Path>,
        options: crate::save::Options,
    ) -> Result<u64, crate::save::Error> {
        litchi_core::atomic::save(
            path.as_ref(),
            options.durability_value(),
            options.overwrite_value(),
            |sink| {
                self.write_to(&mut *sink)?;
                use std::io::Seek;
                Ok(sink.stream_position()?)
            },
            || Ok(()),
        )
    }

    /// Serialize this snapshot to a sequential sink.
    ///
    /// A sink failure can leave caller-owned output incomplete. Filesystem
    /// replacement is provided by [`Self::save`].
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_to(&self, output: impl Write) -> io::Result<()> {
        let mut writer = crate::write::Writer::new(output);
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

    pub(crate) fn model(&self) -> &Tree<'static> {
        &self.inner.model
    }

    pub(crate) fn limits(&self) -> SourceLimits {
        self.inner.options.source_limits_value()
    }

    /// The resource limits and repair posture this snapshot was opened with.
    ///
    /// Edits re-read their spliced output under the same options, so a
    /// document opened through an opt-in repair stays editable.
    pub(crate) fn options(&self) -> OpenOptions {
        self.inner.options
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

    /// Starts a bounded detached edit with caller-selected semantic policies.
    #[must_use]
    pub fn edit_with_options(&self, options: crate::edit::EditOptions) -> crate::edit::Edit {
        crate::edit::Edit::new_with_options(self.clone(), options)
    }

    /// Starts a bounded preservation-first append before the exact root close.
    ///
    /// This transaction is distinct from [`crate::streaming::Writer`]:
    /// it is rooted at this immutable existing snapshot and publishes a new
    /// artifact only after a complete source-proof, byte splice, reopen, and
    /// semantic readback.
    #[must_use]
    pub fn edit_tail_append(
        &self,
        selector: crate::tail_append::TailSelector,
    ) -> crate::tail_append::Edit {
        crate::tail_append::Edit::new(self, selector)
    }

    /// Starts a bounded preservation-first append with explicit limits.
    #[must_use]
    pub fn edit_tail_append_with_options(
        &self,
        selector: crate::tail_append::TailSelector,
        options: crate::tail_append::EditOptions,
    ) -> crate::tail_append::Edit {
        crate::tail_append::Edit::new_with_options(self, selector, options)
    }

    /// Apply an exact-source primary patch under this package's open policy.
    ///
    /// # Errors
    /// Returns an error for a source conflict or candidate admission failure.
    pub fn apply(
        &self,
        patch: &crate::edit::Patch,
    ) -> Result<crate::edit::Commit, crate::edit::Error> {
        patch.apply_commit(self)
    }

    /// Apply an exact-source logical-tail patch under this package's open policy.
    ///
    /// # Errors
    /// Returns an error for a source conflict or candidate admission failure.
    pub fn apply_tail_append(
        &self,
        patch: &crate::tail_append::Patch,
    ) -> Result<crate::tail_append::Commit, crate::tail_append::Error> {
        patch.apply_commit(self)
    }

    /// Applies a shared durable RTF semantic patch to this exact snapshot.
    ///
    /// # Errors
    /// Returns an error when the format vocabulary, source digest, semantic
    /// precondition, operation bounds, or candidate validation fails.
    pub fn apply_durable<Mode>(
        &self,
        patch: &litchi_core::patch::Patch<Mode>,
    ) -> Result<litchi_core::edit::Commit<Self, litchi_core::patch::Patch<Mode>>, crate::edit::Error>
    where
        Mode: Clone,
    {
        crate::edit::apply_durable(self, patch)
    }

    /// Apply a sealed forward-only passive external-reference redaction patch.
    ///
    /// This is separate from ordinary reversible edit replay.  The operation
    /// is source/artifact checked, candidate-parsed, and visible-text checked;
    /// no inverse or repair certification is provided. The returned commit
    /// carries the forward-only patch and the [`crate::redact::Diagnostics`]
    /// captured while replaying it. In particular, callers must inspect the
    /// diagnostics of a best-effort commit before treating the replay as
    /// complete.
    pub fn apply_external_reference_redaction(
        &self,
        patch: &litchi_core::patch::Patch<litchi_core::patch::ForwardOnly>,
    ) -> Result<crate::redact::Commit, crate::edit::Error> {
        crate::redact::apply_forward(self, patch)
    }
}

impl fmt::Debug for Package {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Package")
            .field("paragraphs", &self.paragraph_count())
            .field("tables", &self.tables().len())
            .field("pictures", &self.pictures().len())
            .field("fields", &self.fields().len())
            .field("sections", &self.sections().len())
            .finish_non_exhaustive()
    }
}

/// Owned read-only semantic projection of an immutable RTF package.
///
/// A projection shares its package's model and lazy text cache. It remains valid
/// after the originating package is dropped and never observes later edits.
/// Lifecycle and publication belong to [`Package`].
///
/// ```compile_fail
/// let document = litchi_rtf::Package::parse(r"{\rtf1 readonly}").unwrap().document();
/// document.edit();
/// ```
/// ```compile_fail
/// let document = litchi_rtf::Package::parse(r"{\rtf1 readonly}").unwrap().document();
/// document.to_bytes();
/// ```
/// ```compile_fail
/// let document = litchi_rtf::Document::parse(r"{\rtf1 readonly}").unwrap();
/// ```
#[derive(Clone)]
pub struct Document {
    owner: Package,
}

impl Document {
    /// Every repair the parse applied, in discovery order.
    ///
    /// An empty report means the source conformed and nothing was
    /// substituted. Word-compatible repairs are recorded under every
    /// [`Leniency`](crate::leniency::Leniency), including the default
    /// strict posture.
    #[inline]
    #[must_use]
    pub fn tolerance_report(&self) -> &crate::leniency::ToleranceReport {
        self.owner.tolerance_report()
    }

    /// Borrow the document's plain text.
    ///
    /// The flattened text is materialized at most once per snapshot and is
    /// then shared by every clone.
    #[must_use]
    pub fn text(&self) -> &str {
        self.owner.text()
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
        self.owner.write_text_to(output, options)
    }

    /// Whether the body contains no text.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.owner.is_empty()
    }

    /// Number of logical body paragraphs.
    ///
    /// The parser retains this exact value with the immutable snapshot.
    #[must_use]
    pub fn paragraph_count(&self) -> usize {
        self.owner.paragraph_count()
    }

    /// Borrow the main text story through lazy semantic views.
    #[must_use]
    pub fn body(&self) -> crate::text::Story<'_> {
        self.owner.body()
    }

    /// Borrow the document font resources.
    #[must_use]
    pub fn fonts(&self) -> crate::font::Catalog<'_> {
        self.owner.fonts()
    }

    /// Borrow the document color resources.
    #[must_use]
    pub fn colors(&self) -> crate::color::Palette<'_> {
        self.owner.colors()
    }

    /// Borrow body tables in source order.
    #[must_use]
    pub fn tables(&self) -> &[crate::table::Table<'_>] {
        self.owner.tables()
    }

    /// Borrow pictures in source order.
    #[must_use]
    pub fn pictures(&self) -> &[crate::picture::Picture<'_>] {
        self.owner.pictures()
    }

    /// UTF-8 offset in [`Self::text`] where the picture at `index` is
    /// anchored; see [`Package::picture_body_offset`].
    #[must_use]
    pub fn picture_body_offset(&self, index: usize) -> Option<usize> {
        self.owner.picture_body_offset(index)
    }

    /// Borrow root body shapes in source order.
    #[must_use]
    pub fn shapes(&self) -> &[crate::Shape<'_>] {
        self.owner.shapes()
    }

    /// Borrow fields in source order.
    #[must_use]
    pub fn fields(&self) -> &[crate::field::Field<'_>] {
        self.owner.fields()
    }

    /// Borrow inert SmartTag/factoid ranges in body source order.
    #[must_use]
    pub fn smart_tags(&self) -> &[crate::SmartTag<'_>] {
        self.owner.smart_tags()
    }

    /// Borrow complete inert main-body tracked-move start/end ranges in source order.
    ///
    /// Equal tags with opposite kinds identify the two move locations, but the
    /// bounded model does not expose a cross-location pair index or execute
    /// move/deleted/inserted semantics.
    #[must_use]
    pub fn move_bookmarks(&self) -> &[crate::MoveBookmark<'_>] {
        self.owner.move_bookmarks()
    }

    /// Borrow inert embedded and linked object records in body order.
    #[must_use]
    pub fn objects(&self) -> &[crate::EmbeddedObject<'_>] {
        self.owner.objects()
    }

    /// Produce a bounded, content-free semantic and security inventory.
    ///
    /// The report reuses this immutable snapshot. It never reparses the
    /// source, follows or resolves references, executes fields or objects,
    /// repairs syntax, or mutates the document.
    #[must_use]
    pub fn validation_report(&self) -> crate::ValidationReport {
        self.owner.validation_report()
    }

    /// Alias for [`Self::validation_report`].
    #[must_use]
    pub fn security_report(&self) -> crate::ValidationReport {
        self.owner.security_report()
    }

    /// Borrow list definitions in list-table order.
    #[must_use]
    pub fn lists(&self) -> &[crate::list::List<'_>] {
        self.owner.lists()
    }

    /// Borrow list-instance overrides in override-table order.
    #[must_use]
    pub fn list_overrides(&self) -> &[crate::list::ListOverride] {
        self.owner.list_overrides()
    }

    /// Borrow sections in source order.
    #[must_use]
    pub fn sections(&self) -> &[crate::section::Section<'_>] {
        self.owner.sections()
    }

    /// Borrow named styles in stylesheet order.
    #[must_use]
    pub fn styles(&self) -> &[crate::style::Style<'_>] {
        self.owner.styles()
    }

    /// Borrow inert document metadata.
    #[must_use]
    pub fn info(&self) -> &crate::metadata::Info<'_> {
        self.owner.info()
    }

    /// Borrow inert top-level `nextfile` and `template` metadata without
    /// resolving either name.
    #[must_use]
    pub fn external_references(&self) -> &crate::DocumentExternalReferences<'_> {
        self.owner.external_references()
    }

    /// Borrow annotations in source order.
    #[must_use]
    pub fn annotations(&self) -> &[crate::review::Annotation<'_>] {
        self.owner.annotations()
    }

    /// Borrow footnotes and endnotes in source order.
    #[must_use]
    pub fn notes(&self) -> &[crate::review::Note<'_>] {
        self.owner.notes()
    }

    /// Lazily traverse footnotes in source order.
    #[must_use]
    pub fn footnotes(&self) -> impl DoubleEndedIterator<Item = &crate::review::Note<'_>> + '_ {
        self.owner.footnotes()
    }

    /// Lazily traverse endnotes in source order.
    #[must_use]
    pub fn endnotes(&self) -> impl DoubleEndedIterator<Item = &crate::review::Note<'_>> + '_ {
        self.owner.endnotes()
    }

    /// Borrow tracked revisions in source order.
    #[must_use]
    pub fn revisions(&self) -> &[crate::review::Revision<'_>] {
        self.owner.revisions()
    }

    /// Borrow unsupported syntax retained as bounded inert data.
    #[must_use]
    pub fn opaque(&self) -> &[crate::opaque::Node] {
        self.owner.opaque()
    }
}

impl fmt::Debug for Document {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Document")
            .field("paragraphs", &self.paragraph_count())
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
    use super::Package;
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
        let document = Package::parse(source).unwrap();
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
        let document = Package::parse(r"{\rtf1 one\line two\par three}").unwrap();
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
        let document = Package::parse(r"{\rtf1 one\par two\par three}").unwrap();
        assert_eq!(document.paragraph_count(), 3);

        document
            .edit()
            .replace_paragraph_text(litchi_core::Position::new(1), "changed")
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
