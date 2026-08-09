//! Narrow source-preserving editing for ordinary main-story DOC paragraphs.
//!
//! A Word binary file has independent piece-table, FKP, PLCF, and CFB layers.
//! This API intentionally edits only a paragraph wholly contained in one
//! Unicode piece and only when replacement text has the same UTF-16 length.
//! Consequently no piece descriptor, FKP page, PLCF, FIB length, or table
//! stream changes. Everything outside the replaced `WordDocument` bytes is
//! retained by the package editor. Other edits are rejected with a typed
//! refusal instead of approximating a DOC rewrite.

use crate::package::Error as PackageError;
use crate::tracked_revision::{Limits, RevisionEditor, RevisionKind};
use litchi_core::Position;
use std::sync::Arc;

/// Main-story text visibility used for a review projection.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
pub enum Projection {
    /// Stored text, including both insertion and deletion redline text.
    #[default]
    All,
    /// Text visible after accepting insertion and deletion revisions.
    Accepted,
    /// Text visible after rejecting insertion and deletion revisions.
    Rejected,
}

/// A visible ordinary paragraph in the main document story.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Paragraph {
    position: Position,
    text: String,
}

impl Paragraph {
    /// Zero-based paragraph position in the selected projection.
    ///
    /// Constructing a [`Position`] is infallible. Resolving it against a
    /// snapshot collection is checked by [`Edit::replace_paragraph`].
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Plain inert paragraph text, without the terminating paragraph mark.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}

/// A reason why an edit is outside this intentionally small safe closure.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    /// The selected [`Position`] does not exist in the source body.
    ParagraphNotFound,
    /// A replacement changes UTF-16 length and would require a CLX/PLCF rewrite.
    LengthChange { expected: usize, actual: usize },
    /// The paragraph crosses pieces, which can have distinct encodings and PRMs.
    CrossesPieceBoundary,
    /// The selected paragraph is stored in an ANSI/compressed piece.
    CompressedPiece,
    /// Fields, object markers, cell markers, or other structural controls occur.
    StructuralContent,
    /// The paragraph intersects text-affecting tracked revisions.
    TrackedText,
    /// The requested replacement contains structural controls.
    ReplacementContainsStructuralContent,
    /// The source's review ranges overlap in a way this projection cannot prove.
    AmbiguousReviewRanges,
}

impl std::fmt::Display for Refusal {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::ParagraphNotFound => {
                formatter.write_str("body paragraph position is out of range")
            },
            Self::LengthChange { expected, actual } => write!(
                formatter,
                "replacement has {actual} UTF-16 units; this source paragraph has {expected}"
            ),
            Self::CrossesPieceBoundary => {
                formatter.write_str("body paragraph crosses DOC text pieces")
            },
            Self::CompressedPiece => {
                formatter.write_str("body paragraph is stored in a compressed DOC text piece")
            },
            Self::StructuralContent => {
                formatter.write_str("body paragraph contains DOC structural content")
            },
            Self::TrackedText => formatter.write_str("body paragraph intersects tracked text"),
            Self::ReplacementContainsStructuralContent => {
                formatter.write_str("replacement contains DOC structural content")
            },
            Self::AmbiguousReviewRanges => {
                formatter.write_str("tracked revision ranges overlap ambiguously")
            },
        }
    }
}

/// Failure from a body-text transaction or source-checked patch.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// The DOC/CFB source or its required invariants is invalid.
    Invalid(PackageError),
    /// The request is valid in general but unsafe for this preservation seam.
    Refused(Refusal),
    /// A patch was presented with any snapshot other than its exact source.
    Conflict,
}

impl std::fmt::Display for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Invalid(error) => error.fmt(formatter),
            Self::Refused(refusal) => refusal.fmt(formatter),
            Self::Conflict => formatter.write_str("body-text patch source conflict"),
        }
    }
}

impl std::error::Error for Error {}

/// Immutable, exact-source snapshot for the body-text transaction seam.
#[derive(Clone)]
pub struct Snapshot {
    source: Arc<[u8]>,
    limits: Limits,
}

impl Snapshot {
    /// Opens an owned Word 97+ DOC source after validating its safe edit basis.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when the CFB, Word 97+ FIB, selected table
    /// stream, piece table, or FKP basis cannot support safe editing.
    pub fn open(input: impl Into<Vec<u8>>, limits: Limits) -> Result<Self> {
        let bytes = input.into();
        RevisionEditor::open(bytes.clone(), limits).map_err(Error::Invalid)?;
        Ok(Self {
            source: Arc::from(bytes.into_boxed_slice()),
            limits,
        })
    }

    /// Parses a borrowed DOC source with the default resource limits.
    ///
    /// # Errors
    ///
    /// Returns the same validation failure as [`Self::open`].
    pub fn parse(input: &[u8]) -> Result<Self> {
        Self::open(input.to_vec(), Limits::default())
    }

    /// Opens an owned DOC source with the default resource limits.
    ///
    /// # Errors
    ///
    /// Returns the same validation failure as [`Self::open`].
    pub fn from_bytes(input: Vec<u8>) -> Result<Self> {
        Self::open(input, Limits::default())
    }

    /// Exact CFB source bytes retained for source checks and byte-exact no-ops.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.source
    }

    /// Shared ownership of the exact source allocation.
    #[must_use]
    pub fn bytes_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.source)
    }

    /// Stable first-stage fingerprint for diagnostics and stale-source checks.
    #[must_use]
    pub fn fingerprint(&self) -> u64 {
        fingerprint(&self.source)
    }

    /// Lists ordinary source-body paragraphs under the requested review projection.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] for an invalid source or [`Error::Refused`]
    /// when tracked text ranges overlap ambiguously for the projection.
    pub fn paragraphs(&self, projection: Projection) -> Result<Vec<Paragraph>> {
        let editor = self.editor()?;
        projected_paragraphs(&editor, projection)
    }

    /// Starts a staged same-shape body-paragraph transaction.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] if the retained source no longer validates.
    pub fn edit(&self) -> Result<Edit> {
        Edit::new(self.clone())
    }

    /// Alias for [`Self::edit`].
    ///
    /// # Errors
    ///
    /// Returns the same failure as [`Self::edit`].
    pub fn transaction(&self) -> Result<Edit> {
        self.edit()
    }

    /// Exact source bytes. A snapshot has no implicit serialization step.
    #[must_use]
    pub fn finish(&self) -> Vec<u8> {
        self.source.as_ref().to_vec()
    }

    fn editor(&self) -> Result<RevisionEditor> {
        RevisionEditor::open(self.source.as_ref().to_vec(), self.limits).map_err(Error::Invalid)
    }
}

impl std::fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("bytes", &self.source.len())
            .field("fingerprint", &self.fingerprint())
            .field("limits", &self.limits)
            .finish()
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.source == other.source
    }
}

impl Eq for Snapshot {}

/// Clone-first staged text edit over one immutable source snapshot.
pub struct Edit {
    source: Snapshot,
    editor: RevisionEditor,
}

impl Edit {
    fn new(source: Snapshot) -> Result<Self> {
        let editor = source.editor()?;
        Ok(Self { source, editor })
    }

    /// Immutable source snapshot that authorizes this transaction.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Replaces text in one ordinary source-body paragraph.
    ///
    /// The replacement must have the same UTF-16 length and be confined to a
    /// single Unicode piece. A successful edit does not alter the CLX, FKP,
    /// PLCF, FIB, or table stream.
    ///
    /// `position` is a format-neutral [`Position`]; its membership in this
    /// source body is checked here and an absent paragraph is reported as
    /// [`Refusal::ParagraphNotFound`].
    ///
    /// # Errors
    ///
    /// Returns a typed [`Refusal`] for every operation outside the proven
    /// same-shape closure and [`Error::Invalid`] for a failed package update.
    pub fn replace_paragraph(&mut self, position: Position, replacement: &str) -> Result<()> {
        let paragraphs = source_paragraphs(&self.editor)?;
        let paragraph = paragraphs
            .get(position.get())
            .ok_or(Error::Refused(Refusal::ParagraphNotFound))?;
        if has_structural_content(&paragraph.text) {
            return Err(Error::Refused(Refusal::StructuralContent));
        }
        if has_structural_content(replacement) {
            return Err(Error::Refused(
                Refusal::ReplacementContainsStructuralContent,
            ));
        }
        let expected = paragraph.text.encode_utf16().count();
        let actual = replacement.encode_utf16().count();
        if actual != expected {
            return Err(Error::Refused(Refusal::LengthChange { expected, actual }));
        }
        if !self
            .editor
            .is_unicode_piece_range(paragraph.start_cp, paragraph.end_cp)
        {
            return Err(Error::Refused(piece_refusal(
                &self.editor,
                paragraph.start_cp,
                paragraph.end_cp,
            )));
        }
        if text_revision_intersects(&self.editor, paragraph.start_cp, paragraph.end_cp)? {
            return Err(Error::Refused(Refusal::TrackedText));
        }
        self.editor
            .replace_unicode_text_same_length(paragraph.start_cp, paragraph.end_cp, replacement)
            .map_err(Error::Invalid)
    }

    /// Discards staged changes and returns the original immutable snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    /// Publishes a validated snapshot and its reversible source-checked patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when the rendered candidate cannot be
    /// reopened with the original safety limits.
    pub fn commit(self) -> Result<Commit> {
        let bytes = self.editor.finish().map_err(Error::Invalid)?;
        let snapshot = if bytes == self.source.bytes() {
            self.source.clone()
        } else {
            Snapshot::open(bytes, self.source.limits)?
        };
        let patch = Patch::new(self.source, snapshot.clone());
        Ok(Commit { snapshot, patch })
    }
}

/// Validated commit result for one body-text transaction.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Whether a DOC byte changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.is_noop()
    }

    /// Published post-edit snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Splits a commit into its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// In-memory reversible replacement guarded by exact source bytes.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
    before_fingerprint: u64,
    after_fingerprint: u64,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self {
            before_fingerprint: before.fingerprint(),
            after_fingerprint: after.fingerprint(),
            before,
            after,
        }
    }

    /// Exact source snapshot required for application.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Snapshot produced by the transaction.
    #[must_use]
    pub const fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Fast stale-source precheck; exact bytes remain authoritative.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.before_fingerprint
    }

    /// Target diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.after_fingerprint
    }

    /// Whether this patch preserves the exact artifact.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.bytes() == self.after.bytes()
    }

    /// Applies only to the exact source artifact.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Conflict`] unless `source` has byte-for-byte equality
    /// with this patch's captured source snapshot.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.fingerprint() != self.before_fingerprint || source.bytes() != self.before.bytes()
        {
            return Err(Error::Conflict);
        }
        Ok(if self.is_noop() {
            source.clone()
        } else {
            self.after.clone()
        })
    }

    /// Exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            before_fingerprint: self.after_fingerprint,
            after_fingerprint: self.before_fingerprint,
        }
    }
}

type Result<T> = std::result::Result<T, Error>;

#[derive(Clone)]
struct SourceParagraph {
    start_cp: u32,
    end_cp: u32,
    text: String,
}

fn source_paragraphs(editor: &RevisionEditor) -> Result<Vec<SourceParagraph>> {
    let text = editor.main_story_text().map_err(Error::Invalid)?;
    let mut output = Vec::new();
    let mut start_cp = 0u32;
    let mut start_byte = 0usize;
    let mut cp = 0u32;
    for (byte, character) in text.char_indices() {
        let width = if character.len_utf16() == 1 { 1 } else { 2 };
        let next_cp = cp.checked_add(width).ok_or_else(|| {
            Error::Invalid(PackageError::Corrupted(
                "main-story CP overflow".to_string(),
            ))
        })?;
        if character == '\r' {
            if !editor.is_in_table_at_cp(cp).map_err(Error::Invalid)? {
                output.push(SourceParagraph {
                    start_cp,
                    end_cp: cp,
                    text: text[start_byte..byte].to_string(),
                });
            }
            start_cp = next_cp;
            start_byte = byte + character.len_utf8();
        } else if character == '\u{7}' {
            // A table cell marker is never an ordinary body paragraph.
            start_cp = next_cp;
            start_byte = byte + character.len_utf8();
        }
        cp = next_cp;
    }
    if cp != editor.main_story_cp_len() {
        return Err(Error::Invalid(PackageError::Corrupted(
            "decoded main story has an inconsistent CP length".to_string(),
        )));
    }
    Ok(output)
}

fn projected_paragraphs(editor: &RevisionEditor, projection: Projection) -> Result<Vec<Paragraph>> {
    let source = source_paragraphs(editor)?;
    if projection == Projection::All {
        return Ok(source
            .into_iter()
            .enumerate()
            .map(|(position, paragraph)| Paragraph {
                position: Position::new(position),
                text: paragraph.text,
            })
            .collect());
    }
    let hidden = hidden_ranges(editor, projection)?;
    let mut output = Vec::new();
    for paragraph in source {
        let text = project_text(&paragraph, &hidden)?;
        output.push(Paragraph {
            position: Position::new(output.len()),
            text,
        });
    }
    Ok(output)
}

fn hidden_ranges(editor: &RevisionEditor, projection: Projection) -> Result<Vec<(u32, u32)>> {
    let mut ranges = Vec::new();
    for revision in editor.revisions().map_err(Error::Invalid)? {
        let hide = matches!(
            (projection, revision.kind),
            (
                Projection::Accepted,
                RevisionKind::Deletion | RevisionKind::MoveFrom
            ) | (
                Projection::Rejected,
                RevisionKind::Insertion | RevisionKind::MoveTo
            )
        );
        if hide && revision.start_cp < revision.end_cp {
            ranges.push((revision.start_cp, revision.end_cp));
        }
    }
    ranges.sort_unstable();
    if ranges.windows(2).any(|pair| pair[0].1 > pair[1].0) {
        return Err(Error::Refused(Refusal::AmbiguousReviewRanges));
    }
    Ok(ranges)
}

fn project_text(paragraph: &SourceParagraph, hidden: &[(u32, u32)]) -> Result<String> {
    let mut output = String::new();
    let mut cp = paragraph.start_cp;
    for character in paragraph.text.chars() {
        let width = if character.len_utf16() == 1 { 1 } else { 2 };
        let end = cp.checked_add(width).ok_or_else(|| {
            Error::Invalid(PackageError::Corrupted(
                "projection CP overflow".to_string(),
            ))
        })?;
        if !hidden
            .iter()
            .any(|(start, finish)| *start < end && cp < *finish)
        {
            output.push(character);
        }
        cp = end;
    }
    Ok(output)
}

fn text_revision_intersects(editor: &RevisionEditor, start: u32, end: u32) -> Result<bool> {
    Ok(editor
        .revisions()
        .map_err(Error::Invalid)?
        .into_iter()
        .any(|revision| {
            matches!(
                revision.kind,
                RevisionKind::Insertion
                    | RevisionKind::Deletion
                    | RevisionKind::MoveFrom
                    | RevisionKind::MoveTo
            ) && revision.start_cp < end
                && start < revision.end_cp
        }))
}

fn piece_refusal(editor: &RevisionEditor, start: u32, end: u32) -> Refusal {
    if editor.piece_count_for_range(start, end) > 1 {
        Refusal::CrossesPieceBoundary
    } else {
        Refusal::CompressedPiece
    }
}

fn has_structural_content(text: &str) -> bool {
    text.chars().any(|character| {
        matches!(character, '\r' | '\u{7}' | '\u{13}'..='\u{15}' | '\u{fffc}')
            || (character.is_control() && character != '\t')
    })
}

fn fingerprint(bytes: &[u8]) -> u64 {
    let mut value = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    value
}

#[cfg(test)]
mod tests {
    use super::{Error, Projection, Refusal, Snapshot};
    use crate::tracked_revision::Limits;
    use crate::writer::{CharacterFormatting, ParagraphFormatting, TextRevision, Writer};
    use litchi_core::Position;
    use std::io::Cursor;

    fn doc(paragraphs: &[&str]) -> Vec<u8> {
        let mut writer = Writer::new();
        for paragraph in paragraphs {
            writer
                .add_paragraph_runs(
                    vec![(paragraph.to_string(), CharacterFormatting::default())],
                    ParagraphFormatting::default(),
                )
                .expect("fixture paragraph must be valid");
        }
        let mut output = Cursor::new(Vec::new());
        writer
            .write_to(&mut output)
            .expect("fixture DOC must serialize");
        output.into_inner()
    }

    #[test]
    fn same_shape_body_edit_is_reversible_and_source_checked() {
        let source = Snapshot::parse(&doc(&["alpha", "bravo"])).expect("snapshot");
        assert_eq!(
            source
                .paragraphs(Projection::All)
                .expect("paragraphs")
                .iter()
                .map(|paragraph| paragraph.text())
                .collect::<Vec<_>>(),
            ["alpha", "bravo"]
        );
        assert_eq!(
            source.paragraphs(Projection::All).expect("paragraphs")[0].position(),
            Position::new(0)
        );

        let mut edit = source.edit().expect("edit");
        edit.replace_paragraph(Position::new(0), "omega")
            .expect("same shape edit");
        let commit = edit.commit().expect("commit");
        assert!(commit.changed());
        assert_eq!(
            commit
                .snapshot()
                .paragraphs(Projection::All)
                .expect("changed paragraphs")[0]
                .text(),
            "omega"
        );

        let applied = commit.patch().apply(&source).expect("exact source applies");
        assert_eq!(applied, *commit.snapshot());
        let restored = commit
            .patch()
            .inverse()
            .apply(&applied)
            .expect("inverse applies");
        assert_eq!(restored.bytes(), source.bytes());

        let other = Snapshot::open(doc(&["other"]), Limits::default()).expect("other source");
        assert!(matches!(commit.patch().apply(&other), Err(Error::Conflict)));
    }

    #[test]
    fn length_and_structural_changes_are_refused_before_publication() {
        let source = Snapshot::parse(&doc(&["alpha"])).expect("snapshot");
        let mut edit = source.edit().expect("edit");
        assert!(matches!(
            edit.replace_paragraph(Position::new(0), "longer"),
            Err(Error::Refused(Refusal::LengthChange { .. }))
        ));
        assert!(matches!(
            edit.replace_paragraph(Position::new(0), "a\rpha"),
            Err(Error::Refused(
                Refusal::ReplacementContainsStructuralContent
            ))
        ));
        assert!(matches!(
            edit.replace_paragraph(Position::new(1), "alpha"),
            Err(Error::Refused(Refusal::ParagraphNotFound))
        ));
        let commit = edit.commit().expect("no-op commit");
        assert!(!commit.changed());
        assert_eq!(commit.snapshot().bytes(), source.bytes());
    }

    #[test]
    fn accepted_and_rejected_projections_hide_text_revisions() {
        let mut writer = Writer::new();
        writer
            .add_paragraph_runs(
                vec![
                    ("kept ".to_string(), CharacterFormatting::default()),
                    (
                        "old".to_string(),
                        CharacterFormatting {
                            deletion_revision: Some(TextRevision::new("Reviewer")),
                            ..CharacterFormatting::default()
                        },
                    ),
                    (" new".to_string(), CharacterFormatting::default()),
                ],
                ParagraphFormatting::default(),
            )
            .expect("fixture paragraph");
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).expect("fixture DOC");
        let snapshot = Snapshot::parse(&output.into_inner()).expect("snapshot");
        assert_eq!(
            snapshot.paragraphs(Projection::Accepted).expect("accepted")[0].text(),
            "kept  new"
        );
        assert_eq!(
            snapshot.paragraphs(Projection::Rejected).expect("rejected")[0].text(),
            "kept old new"
        );
    }
}
