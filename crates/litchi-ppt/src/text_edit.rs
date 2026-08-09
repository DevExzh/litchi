//! Source-checked, lossless edits of text already owned by a PPT shape.
//!
//! This deliberately supports only a replacement that leaves the selected
//! `TextCharsAtom` or `TextBytesAtom` byte length unchanged.  `[MS-PPT]`
//! associates character and paragraph formatting, text-range interactions,
//! and special-information runs with UTF-16 positions.  Changing that length
//! without rewriting every such dependent record would be lossy, so it is a
//! typed refusal rather than a best-effort edit.

use std::fmt;
use std::io::Cursor;
use std::sync::Arc;

pub use litchi_core::Position;

use crate::consts::RecordType;
use crate::package::{Error as PackageError, Package, Result as PackageResult};
use crate::presentation::Presentation;
use crate::shapes::ShapeEnum;

const PPT_HEADER_LEN: usize = 8;
const OFFICEART_SP_CONTAINER: u16 = 0xF004;
const OFFICEART_SP: u16 = 0xF00A;
const OFFICEART_CLIENT_TEXTBOX: u16 = 0xF00D;

/// Semantic identity of one existing text-bearing shape.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Target {
    slide: Position,
    shape: Position,
}

impl Target {
    /// Creates a target from zero-based semantic positions in the immutable
    /// presentation and slide shape collections.
    #[must_use]
    pub const fn new(slide: Position, shape: Position) -> Self {
        Self { slide, shape }
    }

    /// The selected slide.
    #[must_use]
    pub const fn slide(self) -> Position {
        self.slide
    }

    /// The selected shape's zero-based position in its slide's source order.
    #[must_use]
    pub const fn shape(self) -> Position {
        self.shape
    }
}

/// A reason why a text edit cannot be published without rewriting unmodeled
/// dependencies.
#[non_exhaustive]
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Refusal {
    /// The source package is signed, encrypted, or otherwise cannot be
    /// republished through the incremental PPT transaction owner.
    UnsupportedSource,
    /// No slide resolved from the public selector.
    SlideNotFound { position: Position },
    /// No shape exists at the selected source-order position.
    ShapeNotFound,
    /// The selected semantic shape position resolved to ambiguous native data.
    AmbiguousShape,
    /// The selected shape has no host `ClientTextbox` payload.
    NoTextbox,
    /// More than one host textbox payload would need a choice.
    AmbiguousTextbox,
    /// The textbox has no editable text atom.
    NoTextAtom,
    /// The textbox has several text atoms and their owner relationship is not
    /// modeled by this focused transaction.
    MultipleTextAtoms,
    /// The replacement cannot be represented by the original text atom's
    /// encoding while preserving its byte length.
    IncompatibleEncoding,
    /// The replacement changes UTF-16 position counts and would invalidate
    /// formatting, interaction, or special-information ranges.
    DependencyClosure,
}

impl fmt::Display for Refusal {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::UnsupportedSource => write!(
                formatter,
                "PPT source cannot be republished through the safe shape-text transaction"
            ),
            Self::SlideNotFound { position } => {
                write!(
                    formatter,
                    "PPT slide position {} was not found",
                    position.get()
                )
            },
            Self::ShapeNotFound => write!(formatter, "PPT shape position was not found"),
            Self::AmbiguousShape => write!(formatter, "PPT shape position is ambiguous"),
            Self::NoTextbox => write!(formatter, "PPT shape has no ClientTextbox payload"),
            Self::AmbiguousTextbox => {
                write!(formatter, "PPT shape has multiple ClientTextbox payloads")
            },
            Self::NoTextAtom => write!(formatter, "PPT shape has no editable text atom"),
            Self::MultipleTextAtoms => write!(formatter, "PPT shape has multiple text atoms"),
            Self::IncompatibleEncoding => write!(
                formatter,
                "replacement text cannot preserve the PPT shape's encoded text atom"
            ),
            Self::DependencyClosure => write!(
                formatter,
                "replacement text changes the selected PPT shape's modeled dependency closure"
            ),
        }
    }
}

/// Error returned by the focused text-edit transaction.
#[derive(Debug)]
pub enum Error {
    /// The package could not be opened, parsed, or republished.
    Package(PackageError),
    /// The requested operation is not proven lossless for this source.
    Refused(Refusal),
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package(error) => error.fmt(formatter),
            Self::Refused(refusal) => refusal.fmt(formatter),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Package(error) => Some(error),
            Self::Refused(_) => None,
        }
    }
}

impl From<PackageError> for Error {
    fn from(error: PackageError) -> Self {
        Self::Package(error)
    }
}

/// Result type for source-checked existing-shape text edits.
pub type Result<T> = std::result::Result<T, Error>;

/// Immutable whole-package snapshot used by [`Transaction`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    bytes: Arc<[u8]>,
}

impl Snapshot {
    /// Opens an exact PPT artifact with default package limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the complete package cannot be opened.
    pub fn parse(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes(bytes.to_vec())
    }

    /// Captures an owned exact PPT artifact after validating that it opens.
    ///
    /// # Errors
    ///
    /// Returns an error when the complete package cannot be opened.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let _ = presentation(&bytes)?;
        Ok(Self {
            bytes: Arc::from(bytes.into_boxed_slice()),
        })
    }

    /// Exact bytes of the complete source or committed package artifact.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Starts a source-checked transaction for one existing shape.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal when the target cannot be resolved to one safe
    /// text atom.
    pub fn edit_text(&self, target: Target) -> Result<Transaction> {
        // The shared incremental owner rejects signed and encrypted CFB
        // envelopes before any candidate is staged. Surface that capability
        // boundary as a typed refusal instead of trying to normalize it.
        crate::embedded::object::Editor::open_records(self.bytes.to_vec())
            .map_err(|_error| Error::Refused(Refusal::UnsupportedSource))?;
        let resolved = resolve(&self.bytes, target)?;
        Ok(Transaction {
            source: self.clone(),
            resolved,
            replacement: None,
        })
    }
}

/// One isolated text replacement staged against an immutable package.
#[derive(Debug, Clone)]
pub struct Transaction {
    source: Snapshot,
    resolved: Resolved,
    replacement: Option<String>,
}

impl Transaction {
    /// The exact text currently stored in the selected shape.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.resolved.text
    }

    /// The selected semantic target.
    #[must_use]
    pub const fn target(&self) -> Target {
        self.resolved.target
    }

    /// Stages a replacement after proving that the dependent text ranges stay
    /// valid. Failed validation leaves the staged candidate untouched.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal when the replacement is not length-preserving.
    pub fn set_text(&mut self, value: impl Into<String>) -> Result<()> {
        let candidate = value.into();
        encode_replacement(&candidate, self.resolved.kind, self.resolved.payload.len())?;
        self.replacement = Some(candidate);
        Ok(())
    }

    /// Publishes the replacement atomically and returns a reversible,
    /// exact-source-checked patch. Equal text reuses the source allocation.
    ///
    /// # Errors
    ///
    /// Returns an error if package publication or source-checked readback
    /// fails, or a typed refusal for an unsafe replacement.
    pub fn commit(self) -> Result<Commit> {
        let replacement = self
            .replacement
            .unwrap_or_else(|| self.resolved.text.clone());
        if replacement == self.resolved.text {
            let patch = Patch::new(self.source.clone(), self.source.clone());
            return Ok(Commit {
                snapshot: self.source,
                patch,
            });
        }

        let encoded = encode_replacement(
            &replacement,
            self.resolved.kind,
            self.resolved.payload.len(),
        )?;
        let target_slide = rewrite_slide(
            &self.resolved.slide_record,
            self.resolved.native_shape_id,
            self.resolved.kind,
            &self.resolved.payload,
            &encoded,
        )?;
        let mut editor =
            crate::embedded::object::Editor::open_records(self.source.bytes.clone().to_vec())?;
        let live = editor.persisted_record(self.resolved.slide_persist_id)?;
        if live != self.resolved.slide_record.as_ref() {
            return Err(PackageError::Corrupted(
                "PPT shape-text transaction source changed before publication".into(),
            )
            .into());
        }
        editor.replace_persisted_record(self.resolved.slide_persist_id, target_slide)?;
        let bytes = editor.finish()?;
        let snapshot = Snapshot::from_bytes(bytes)?;
        let readback = resolve(&snapshot.bytes, self.resolved.target)?;
        if readback.text != replacement {
            return Err(PackageError::Corrupted(
                "published PPT shape text did not round-trip through the selected source shape"
                    .into(),
            )
            .into());
        }
        let patch = Patch::new(self.source, snapshot.clone());
        Ok(Commit { snapshot, patch })
    }

    /// Discards this candidate without changing the source snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }
}

/// A published snapshot and its reversible source-checked patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// The immutable committed package.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// The source-checked patch that produced this snapshot.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consumes the commit into its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A reversible in-memory patch authorized by exact package bytes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Exact source bytes required for forward application.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        self.before.bytes()
    }

    /// Exact target bytes produced by forward application.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        self.after.bytes()
    }

    /// Whether the patch changes no bytes.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.bytes() == self.after.bytes()
    }

    /// Applies only to the exact source artifact used to create this patch.
    ///
    /// # Errors
    ///
    /// Returns an error if `current` is not that exact source artifact.
    pub fn apply(&self, current: &Snapshot) -> Result<Snapshot> {
        if current.bytes() != self.before.bytes() {
            return Err(PackageError::InvalidFormat(
                "PPT shape-text patch source does not match its base artifact".into(),
            )
            .into());
        }
        Ok(self.after.clone())
    }

    /// Returns the exact-source-checked inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self::new(self.after.clone(), self.before.clone())
    }
}

#[derive(Debug, Clone)]
struct Resolved {
    target: Target,
    slide_persist_id: u32,
    native_shape_id: u32,
    slide_record: Arc<[u8]>,
    kind: TextKind,
    payload: Vec<u8>,
    text: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum TextKind {
    Bytes,
    Chars,
}

fn presentation(bytes: &[u8]) -> PackageResult<Presentation> {
    let mut package = Package::from_reader(Cursor::new(bytes.to_vec()))?;
    package.presentation()
}

fn resolve(bytes: &[u8], target: Target) -> Result<Resolved> {
    let presentation = presentation(bytes)?;
    let slides = presentation.slides()?;
    let slide = slides
        .get(target.slide.get())
        .ok_or(Error::Refused(Refusal::SlideNotFound {
            position: target.slide,
        }))?;
    let shape = slide
        .shapes()?
        .get(target.shape.get())
        .ok_or(Error::Refused(Refusal::ShapeNotFound))?;
    let native_shape_id = native_shape_id(shape);
    let slide_record = crate::embedded::object::Editor::open_records(bytes.to_vec())?
        .persisted_record(slide.persist_id())?;
    let text = inspect_slide(&slide_record, native_shape_id)?;
    Ok(Resolved {
        target,
        slide_persist_id: slide.persist_id(),
        native_shape_id,
        slide_record: Arc::from(slide_record.into_boxed_slice()),
        kind: text.kind,
        payload: text.payload,
        text: text.text,
    })
}

fn native_shape_id(selected: &ShapeEnum<'_>) -> u32 {
    match selected {
        ShapeEnum::TextBox(textbox) => textbox.source_shape_id(),
        ShapeEnum::Placeholder(placeholder) => placeholder.source_shape_id(),
        ShapeEnum::AutoShape(auto_shape) => auto_shape.source_shape_id(),
        ShapeEnum::Picture(picture) => picture.properties().id,
        ShapeEnum::Table(table) => table.id(),
        ShapeEnum::Group(group) => group.id(),
        ShapeEnum::Line(line) => line.id(),
    }
}

#[derive(Debug)]
struct TextAtom {
    kind: TextKind,
    payload: Vec<u8>,
    text: String,
}

fn inspect_slide(slide: &[u8], shape_id: u32) -> Result<TextAtom> {
    let (_, consumed) = crate::Record::parse_strict(slide, 0)?;
    if consumed != slide.len() {
        return Err(PackageError::Corrupted("selected slide has trailing bytes".into()).into());
    }
    let drawing = find_ppt_record(slide, RecordType::PPDrawing as u16, 0)?
        .ok_or(Error::Refused(Refusal::ShapeNotFound))?;
    inspect_drawing(drawing, shape_id)
}

fn find_ppt_record(record: &[u8], target: u16, depth: usize) -> PackageResult<Option<&[u8]>> {
    if depth >= 128 {
        return Err(PackageError::Corrupted(
            "PPT record nesting exceeds text-edit limit".into(),
        ));
    }
    if record_type(record)? == target {
        return Ok(Some(record));
    }
    if record_version(record)? != 0xF {
        return Ok(None);
    }
    let mut found = None;
    for child_result in children(record)? {
        let child = child_result?;
        if let Some(candidate) = find_ppt_record(child, target, depth + 1)? {
            if found.is_some() {
                return Err(PackageError::Corrupted(
                    "selected slide has multiple PPDrawing records".into(),
                ));
            }
            found = Some(candidate);
        }
    }
    Ok(found)
}

fn inspect_drawing(drawing: &[u8], shape_id: u32) -> Result<TextAtom> {
    let mut matches = 0usize;
    let mut text = None;
    visit_officeart(
        drawing_payload(drawing)?,
        shape_id,
        0,
        &mut matches,
        &mut |textbox| {
            let atom = inspect_textbox(textbox)?;
            text = Some(atom);
            Ok(())
        },
    )?;
    match matches {
        0 => Err(Error::Refused(Refusal::ShapeNotFound)),
        1 => text.ok_or(Error::Refused(Refusal::NoTextbox)),
        _ => Err(Error::Refused(Refusal::AmbiguousShape)),
    }
}

fn inspect_textbox(textbox: &[u8]) -> Result<TextAtom> {
    let mut atom = None;
    for child_result in children(textbox)? {
        let child = child_result?;
        let atom_kind = match record_type(child)? {
            value if value == RecordType::TextBytesAtom as u16 => Some(TextKind::Bytes),
            value if value == RecordType::TextCharsAtom as u16 => Some(TextKind::Chars),
            _ => None,
        };
        let Some(kind) = atom_kind else { continue };
        if atom.is_some() {
            return Err(Error::Refused(Refusal::MultipleTextAtoms));
        }
        let payload = drawing_payload(child)?.to_vec();
        let text = match kind {
            TextKind::Bytes => payload.iter().map(|byte| char::from(*byte)).collect(),
            TextKind::Chars => decode_utf16(&payload)?,
        };
        atom = Some(TextAtom {
            kind,
            payload,
            text,
        });
    }
    atom.ok_or(Error::Refused(Refusal::NoTextAtom))
}

fn rewrite_slide(
    slide: &[u8],
    shape_id: u32,
    kind: TextKind,
    before: &[u8],
    after: &[u8],
) -> Result<Vec<u8>> {
    let mut drawing_count = 0usize;
    let rewritten = rewrite_ppt_record(slide, 0, &mut |record| {
        if record_type(record)? != RecordType::PPDrawing as u16 {
            return Ok(None);
        }
        drawing_count += 1;
        let mut shapes = 0usize;
        let drawing = rewrite_drawing(record, shape_id, kind, before, after, &mut shapes)?;
        if shapes == 0 {
            return Ok(None);
        }
        if shapes > 1 {
            return Err(Error::Refused(Refusal::AmbiguousShape));
        }
        Ok(Some(drawing))
    })?;
    if drawing_count != 1 {
        return Err(PackageError::Corrupted(
            "selected slide has ambiguous PPDrawing ownership".into(),
        )
        .into());
    }
    rewritten.ok_or(Error::Refused(Refusal::ShapeNotFound))
}

fn rewrite_ppt_record(
    record: &[u8],
    depth: usize,
    visit: &mut impl FnMut(&[u8]) -> Result<Option<Vec<u8>>>,
) -> Result<Option<Vec<u8>>> {
    if depth >= 128 {
        return Err(
            PackageError::Corrupted("PPT record nesting exceeds text-edit limit".into()).into(),
        );
    }
    if let Some(replacement) = visit(record)? {
        return Ok(Some(replacement));
    }
    if record_version(record)? != 0xF {
        return Ok(None);
    }
    let mut changed = false;
    let mut data = Vec::with_capacity(drawing_payload(record)?.len());
    for child_result in children(record)? {
        let child = child_result?;
        if let Some(replacement) = rewrite_ppt_record(child, depth + 1, visit)? {
            changed = true;
            data.extend_from_slice(&replacement);
        } else {
            data.extend_from_slice(child);
        }
    }
    if !changed {
        return Ok(None);
    }
    rebuild(record, &data).map(Some).map_err(Into::into)
}

fn rewrite_drawing(
    drawing: &[u8],
    shape_id: u32,
    kind: TextKind,
    before: &[u8],
    after: &[u8],
    matches: &mut usize,
) -> Result<Vec<u8>> {
    let data = rewrite_officeart(
        drawing_payload(drawing)?,
        shape_id,
        kind,
        before,
        after,
        0,
        matches,
    )?;
    rebuild(drawing, &data).map_err(Into::into)
}

fn rewrite_officeart(
    record: &[u8],
    shape_id: u32,
    kind: TextKind,
    before: &[u8],
    after: &[u8],
    depth: usize,
    matches: &mut usize,
) -> Result<Vec<u8>> {
    if depth >= 128 {
        return Err(
            PackageError::Corrupted("OfficeArt nesting exceeds text-edit limit".into()).into(),
        );
    }
    if record_type(record)? == OFFICEART_SP_CONTAINER && shape_id_of(record)? == Some(shape_id) {
        *matches = matches.saturating_add(1);
        return rewrite_shape_container(record, kind, before, after);
    }
    if record_version(record)? != 0xF {
        return Ok(record.to_vec());
    }
    let mut data = Vec::with_capacity(drawing_payload(record)?.len());
    for child_result in children(record)? {
        let child = child_result?;
        data.extend_from_slice(&rewrite_officeart(
            child,
            shape_id,
            kind,
            before,
            after,
            depth + 1,
            matches,
        )?);
    }
    rebuild(record, &data).map_err(Into::into)
}

fn rewrite_shape_container(
    record: &[u8],
    kind: TextKind,
    before: &[u8],
    after: &[u8],
) -> Result<Vec<u8>> {
    let mut textbox_count = 0usize;
    let mut replaced = false;
    let mut data = Vec::with_capacity(drawing_payload(record)?.len());
    for child_result in children(record)? {
        let child = child_result?;
        if record_type(child)? != OFFICEART_CLIENT_TEXTBOX {
            data.extend_from_slice(child);
            continue;
        }
        textbox_count += 1;
        let rewritten = rewrite_textbox(child, kind, before, after)?;
        replaced = true;
        data.extend_from_slice(&rewritten);
    }
    if textbox_count == 0 {
        return Err(Error::Refused(Refusal::NoTextbox));
    }
    if textbox_count > 1 {
        return Err(Error::Refused(Refusal::AmbiguousTextbox));
    }
    if !replaced {
        return Err(Error::Refused(Refusal::NoTextAtom));
    }
    rebuild(record, &data).map_err(Into::into)
}

fn rewrite_textbox(textbox: &[u8], kind: TextKind, before: &[u8], after: &[u8]) -> Result<Vec<u8>> {
    let mut count = 0usize;
    let mut data = Vec::with_capacity(drawing_payload(textbox)?.len());
    for child_result in children(textbox)? {
        let child = child_result?;
        let record_type = record_type(child)?;
        let is_target = match kind {
            TextKind::Bytes => record_type == RecordType::TextBytesAtom as u16,
            TextKind::Chars => record_type == RecordType::TextCharsAtom as u16,
        };
        if is_target {
            count += 1;
            if drawing_payload(child)? != before {
                return Err(PackageError::Corrupted(
                    "selected text atom changed before publication".into(),
                )
                .into());
            }
            let mut rewritten = child.to_vec();
            rewritten[PPT_HEADER_LEN..].copy_from_slice(after);
            data.extend_from_slice(&rewritten);
        } else {
            data.extend_from_slice(child);
        }
    }
    match count {
        0 => Err(Error::Refused(Refusal::NoTextAtom)),
        1 => rebuild(textbox, &data).map_err(Into::into),
        _ => Err(Error::Refused(Refusal::MultipleTextAtoms)),
    }
}

fn visit_officeart(
    record: &[u8],
    shape_id: u32,
    depth: usize,
    matches: &mut usize,
    visit: &mut impl FnMut(&[u8]) -> Result<()>,
) -> Result<()> {
    if depth >= 128 {
        return Err(
            PackageError::Corrupted("OfficeArt nesting exceeds text-edit limit".into()).into(),
        );
    }
    if record_type(record)? == OFFICEART_SP_CONTAINER && shape_id_of(record)? == Some(shape_id) {
        *matches = matches.saturating_add(1);
        let mut textbox = None;
        for child_result in children(record)? {
            let child = child_result?;
            if record_type(child)? == OFFICEART_CLIENT_TEXTBOX && textbox.replace(child).is_some() {
                return Err(Error::Refused(Refusal::AmbiguousTextbox));
            }
        }
        let selected_textbox = textbox.ok_or(Error::Refused(Refusal::NoTextbox))?;
        visit(selected_textbox)?;
        return Ok(());
    }
    if record_version(record)? == 0xF {
        for child_result in children(record)? {
            let child = child_result?;
            visit_officeart(child, shape_id, depth + 1, matches, visit)?;
        }
    }
    Ok(())
}

fn shape_id_of(record: &[u8]) -> PackageResult<Option<u32>> {
    for child_result in children(record)? {
        let child = child_result?;
        if record_type(child)? == OFFICEART_SP {
            let payload = drawing_payload(child)?;
            let bytes: &[u8; 4] = payload
                .get(..4)
                .and_then(|value| value.try_into().ok())
                .ok_or_else(|| {
                    PackageError::Corrupted("OfficeArt Sp record has no shape identity".into())
                })?;
            return Ok(Some(u32::from_le_bytes(*bytes)));
        }
    }
    Ok(None)
}

fn encode_replacement(value: &str, kind: TextKind, length: usize) -> Result<Vec<u8>> {
    let bytes = match kind {
        TextKind::Bytes => value
            .chars()
            .map(|character| {
                u8::try_from(u32::from(character))
                    .map_err(|_err| Error::Refused(Refusal::IncompatibleEncoding))
            })
            .collect::<Result<Vec<_>>>()?,
        TextKind::Chars => value.encode_utf16().flat_map(u16::to_le_bytes).collect(),
    };
    if bytes.len() != length {
        return Err(Error::Refused(Refusal::DependencyClosure));
    }
    Ok(bytes)
}

fn decode_utf16(bytes: &[u8]) -> Result<String> {
    if !bytes.len().is_multiple_of(2) {
        return Err(Error::Refused(Refusal::IncompatibleEncoding));
    }
    let units = bytes
        .chunks_exact(2)
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
        .collect::<Vec<_>>();
    String::from_utf16(&units).map_err(|_err| Error::Refused(Refusal::IncompatibleEncoding))
}

fn record_type(record: &[u8]) -> PackageResult<u16> {
    let bytes: &[u8; 2] = record
        .get(2..4)
        .and_then(|value| value.try_into().ok())
        .ok_or_else(|| PackageError::Corrupted("truncated text-edit record header".into()))?;
    Ok(u16::from_le_bytes(*bytes))
}

fn record_version(record: &[u8]) -> PackageResult<u16> {
    let bytes: &[u8; 2] = record
        .get(..2)
        .and_then(|value| value.try_into().ok())
        .ok_or_else(|| PackageError::Corrupted("truncated text-edit record header".into()))?;
    Ok(u16::from_le_bytes(*bytes) & 0x000F)
}

fn drawing_payload(record: &[u8]) -> PackageResult<&[u8]> {
    let bytes: &[u8; 4] = record
        .get(4..8)
        .and_then(|value| value.try_into().ok())
        .ok_or_else(|| PackageError::Corrupted("truncated text-edit record length".into()))?;
    let length = usize::try_from(u32::from_le_bytes(*bytes)).map_err(|_err| {
        PackageError::Corrupted("text-edit record length exceeds this platform".into())
    })?;
    let end = PPT_HEADER_LEN
        .checked_add(length)
        .ok_or_else(|| PackageError::Corrupted("text-edit record length overflows".into()))?;
    record
        .get(PPT_HEADER_LEN..end)
        .ok_or_else(|| PackageError::Corrupted("truncated text-edit record payload".into()))
}

fn children(record: &[u8]) -> PackageResult<impl Iterator<Item = PackageResult<&[u8]>>> {
    let payload = drawing_payload(record)?;
    Ok(ChildRecords { payload, offset: 0 })
}

struct ChildRecords<'a> {
    payload: &'a [u8],
    offset: usize,
}

impl<'a> Iterator for ChildRecords<'a> {
    type Item = PackageResult<&'a [u8]>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.offset == self.payload.len() {
            return None;
        }
        let start = self.offset;
        let record = match drawing_payload(self.payload.get(start..)?) {
            Ok(payload) => match PPT_HEADER_LEN.checked_add(payload.len()) {
                Some(length) => self.payload.get(start..start + length),
                None => None,
            },
            Err(error) => return Some(Err(error)),
        };
        if let Some(child_record) = record {
            self.offset = start + child_record.len();
            Some(Ok(child_record))
        } else {
            self.offset = self.payload.len();
            Some(Err(PackageError::Corrupted(
                "truncated text-edit child record".into(),
            )))
        }
    }
}

fn rebuild(record: &[u8], data: &[u8]) -> PackageResult<Vec<u8>> {
    if data.len() != drawing_payload(record)?.len() {
        return Err(PackageError::Corrupted(
            "text edit unexpectedly changed record framing".into(),
        ));
    }
    let mut output = Vec::with_capacity(record.len());
    output.extend_from_slice(&record[..PPT_HEADER_LEN]);
    output.extend_from_slice(data);
    Ok(output)
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::{Error, Position, Refusal, Snapshot, Target};
    use crate::Package;
    use crate::writer::Writer;
    use std::io::Cursor;

    fn fixture(text: &str) -> Vec<u8> {
        let mut writer = Writer::new();
        let slide = writer.add_slide().unwrap();
        writer.add_textbox(slide, 10, 10, 240, 40, text).unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    fn target() -> Target {
        Target::new(Position::new(0), Position::new(0))
    }

    #[test]
    fn source_checked_text_edit_round_trips_and_reverses() {
        let source = Snapshot::from_bytes(fixture("abc")).unwrap();
        let target = target();
        let source_slide = super::resolve(source.bytes(), target).unwrap();
        let mut edit = source.edit_text(target).unwrap();
        assert_eq!(edit.text(), "abc");
        edit.set_text("xyz").unwrap();
        let commit = edit.commit().unwrap();
        let committed_slide = super::resolve(commit.snapshot().bytes(), target).unwrap();
        assert_eq!(
            source_slide.slide_record.len(),
            committed_slide.slide_record.len()
        );
        let changed = source_slide
            .slide_record
            .iter()
            .zip(committed_slide.slide_record.iter())
            .filter(|(before, after)| before != after)
            .count();
        assert_eq!(changed, 3);
        let mut package =
            Package::from_reader(Cursor::new(commit.snapshot().bytes().to_vec())).unwrap();
        let presentation = package.presentation().unwrap();
        assert_eq!(
            presentation.slides().unwrap()[0].shapes().unwrap()[0]
                .text()
                .unwrap(),
            "xyz"
        );
        let undone = commit.patch().inverse().apply(commit.snapshot()).unwrap();
        assert_eq!(undone.bytes(), source.bytes());
        assert!(commit.patch().apply(&undone).is_ok());
    }

    #[test]
    fn length_changing_edit_is_a_typed_refusal_and_keeps_source() {
        let source = Snapshot::from_bytes(fixture("abc")).unwrap();
        let target = target();
        let mut edit = source.edit_text(target).unwrap();
        assert!(matches!(
            edit.set_text("long"),
            Err(Error::Refused(Refusal::DependencyClosure))
        ));
        let commit = edit.commit().unwrap();
        assert_eq!(commit.snapshot().bytes(), source.bytes());
        assert!(commit.patch().is_empty());
    }

    #[test]
    fn patch_rejects_a_different_source() {
        let source = Snapshot::from_bytes(fixture("abc")).unwrap();
        let target = target();
        let mut edit = source.edit_text(target).unwrap();
        edit.set_text("xyz").unwrap();
        let patch = edit.commit().unwrap().patch().clone();
        let other = Snapshot::from_bytes(fixture("def")).unwrap();
        assert!(patch.apply(&other).is_err());
    }

    #[test]
    fn semantic_positions_are_checked_against_the_source() {
        let source = Snapshot::from_bytes(fixture("abc")).unwrap();
        assert!(matches!(
            source.edit_text(Target::new(Position::new(1), Position::new(0))),
            Err(Error::Refused(Refusal::SlideNotFound { position }))
                if position == Position::new(1)
        ));
        assert!(matches!(
            source.edit_text(Target::new(Position::new(0), Position::new(1))),
            Err(Error::Refused(Refusal::ShapeNotFound))
        ));
    }
}
