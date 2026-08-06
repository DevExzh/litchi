//! Detached source-checked edits for media TracksInfo and narration metadata.

use std::ops::Range;
use std::sync::Arc;

use super::model::{Caption, DisplayLocation, MediaMetadata};
use super::tracks_info::{self, Attr, Found};
use super::validation::{validate_caption, validate_metadata};
use crate::{Error, Result};

/// A stable source revision calculated from the owning slide XML bytes.
pub type Revision = u64;

/// A detached, source-preserving media metadata snapshot.
#[derive(Debug, Clone)]
pub struct Snapshot {
    pub(crate) source_part_name: String,
    pub(crate) source_xml: Arc<Vec<u8>>,
    pub(crate) layout: Found,
    pub(crate) metadata: MediaMetadata,
    pub(crate) revision: Revision,
}

impl Snapshot {
    pub(crate) fn from_wire(
        source_part_name: String,
        source_xml: Vec<u8>,
        layout: Found,
        metadata: MediaMetadata,
    ) -> Result<Self> {
        validate_metadata(&metadata)?;
        let source_xml = Arc::new(source_xml);
        let revision = fingerprint(&source_xml);
        Ok(Self {
            source_part_name,
            source_xml,
            layout,
            metadata,
            revision,
        })
    }

    /// Return the selected media shape identity.
    #[inline]
    pub fn key(&self) -> &super::model::MediaKey {
        &self.metadata.key
    }

    /// Borrow the typed metadata.
    #[inline]
    pub fn metadata(&self) -> &MediaMetadata {
        &self.metadata
    }

    /// Return the source revision used by stale-source checks.
    #[inline]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Borrow the exact owning slide XML captured by this snapshot.
    #[inline]
    pub fn source_xml(&self) -> &[u8] {
        self.source_xml.as_slice()
    }

    /// Start an atomic detached edit.
    #[inline]
    pub fn edit(&self) -> Transaction {
        Transaction {
            original: self.clone(),
            working: self.metadata.clone(),
        }
    }
}

/// A bounded edit staged against one media shape.
#[derive(Debug, Clone)]
pub struct Transaction {
    original: Snapshot,
    working: MediaMetadata,
}

impl Transaction {
    /// Borrow the projected typed metadata.
    #[inline]
    pub fn snapshot(&self) -> &MediaMetadata {
        &self.working
    }

    /// Return whether a source-changing edit has been staged.
    #[inline]
    pub fn is_changed(&self) -> bool {
        self.original.metadata != self.working
    }

    /// Change the schema-defined `tracksInfo/@displayLoc` value.
    pub fn set_display_location(&mut self, value: DisplayLocation) -> Result<()> {
        let tracks = self
            .working
            .tracks_info
            .as_mut()
            .ok_or_else(|| invalid("selected media shape has no tracksInfo metadata"))?;
        tracks.display_location = value;
        Ok(())
    }

    /// Replace one caption's identity and display label without touching its
    /// relationship target or opaque media payload.
    pub fn set_caption_identity(
        &mut self,
        index: usize,
        id: impl Into<String>,
        label: impl Into<String>,
    ) -> Result<()> {
        let caption = self.caption_mut(index)?;
        let mut candidate = caption.clone();
        candidate.id = id.into();
        candidate.label = label.into();
        validate_caption(&candidate)?;
        *caption = candidate;
        Ok(())
    }

    /// Replace one caption's optional language metadata.
    pub fn set_caption_language(&mut self, index: usize, language: Option<String>) -> Result<()> {
        let caption = self.caption_mut(index)?;
        let mut candidate = caption.clone();
        candidate.language = language;
        validate_caption(&candidate)?;
        *caption = candidate;
        Ok(())
    }

    /// Set the shape-level narration flag. `None` means the authored `val`
    /// attribute is absent; it is kept distinct from `Some(false)`.
    pub fn set_narration(&mut self, value: Option<bool>) -> Result<()> {
        if self.original.layout.narration.is_none() && value.is_some() {
            return Err(invalid(
                "cannot create isNarration without rewriting the owning extension list",
            ));
        }
        self.working.narration = value;
        Ok(())
    }

    /// Validate and consume this edit into a source-checked commit.
    pub fn commit(self) -> Result<Commit> {
        validate_metadata(&self.working)?;
        let updated = encode_edit(&self.original, &self.working)?;
        let updated = Arc::new(updated);
        let layout = if updated.as_slice() == self.original.source_xml.as_slice() {
            self.original.layout.clone()
        } else {
            tracks_info::discover(updated.as_slice(), &self.working.key)?.ok_or_else(|| {
                invalid("edited media tracks no longer contain the selected shape")
            })?
        };
        let snapshot = Snapshot {
            source_part_name: self.original.source_part_name.clone(),
            source_xml: Arc::clone(&updated),
            layout,
            metadata: self.working,
            revision: fingerprint(&updated),
        };
        let patch = Patch {
            source_part_name: self.original.source_part_name,
            key: snapshot.metadata.key.clone(),
            expected_revision: self.original.revision,
            expected_xml: Arc::clone(&self.original.source_xml),
            updated_xml: Arc::clone(&updated),
        };
        Ok(Commit { snapshot, patch })
    }

    fn caption_mut(&mut self, index: usize) -> Result<&mut Caption> {
        self.working
            .tracks_info
            .as_mut()
            .and_then(|tracks| tracks.captions.get_mut(index))
            .ok_or_else(|| invalid(format!("media caption index {index} is out of bounds")))
    }
}

/// A committed metadata snapshot and its source-checked package patch.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Borrow the projected post-edit snapshot.
    #[inline]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the source-checked patch for package publication.
    #[inline]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Return whether this commit changes the owning slide bytes.
    #[inline]
    pub fn is_changed(&self) -> bool {
        self.patch.is_changed()
    }

    /// Consume the commit and return its patch.
    #[inline]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// A source-checked, atomic replacement of one owning slide XML part.
#[derive(Debug, Clone)]
pub struct Patch {
    pub(crate) source_part_name: String,
    pub(crate) key: super::model::MediaKey,
    pub(crate) expected_revision: Revision,
    pub(crate) expected_xml: Arc<Vec<u8>>,
    pub(crate) updated_xml: Arc<Vec<u8>>,
}

impl Patch {
    /// Return whether the patch is an exact no-op.
    #[inline]
    pub fn is_changed(&self) -> bool {
        self.expected_xml.as_slice() != self.updated_xml.as_slice()
    }

    /// Return the source revision required for publication.
    #[inline]
    pub const fn expected_revision(&self) -> Revision {
        self.expected_revision
    }
}

fn encode_edit(original: &Snapshot, working: &MediaMetadata) -> Result<Vec<u8>> {
    if original.metadata.key != working.key {
        return Err(invalid(
            "media shape identity cannot change in a transaction",
        ));
    }
    let Some(original_tracks) = original.layout.media.tracks_info.as_ref() else {
        if working.tracks_info.is_some() {
            return Err(invalid("cannot create tracksInfo in a bounded transaction"));
        }
        return encode_narration(original, working, Vec::new());
    };
    let Some(working_tracks) = working.tracks_info.as_ref() else {
        return Err(invalid("tracksInfo cannot be removed by this transaction"));
    };
    if original_tracks.tracks.len() != working_tracks.captions.len() {
        return Err(invalid(
            "caption track count cannot change in a transaction",
        ));
    }

    let mut replacements = Vec::new();
    if original_tracks.display_location.value != working_tracks.display_location.token() {
        replacements.push(Replacement::value(
            &original_tracks.display_location,
            working_tracks.display_location.token(),
        ));
    }
    for (wire, caption) in original_tracks.tracks.iter().zip(&working_tracks.captions) {
        if wire.id.value != caption.id {
            replacements.push(Replacement::value(&wire.id, &caption.id));
        }
        if wire.label.value != caption.label {
            replacements.push(Replacement::value(&wire.label, &caption.label));
        }
        match (&wire.language, &caption.language) {
            (Some(old), Some(new)) if old.value != *new => {
                replacements.push(Replacement::value(old, new));
            },
            (Some(old), None) => replacements.push(Replacement::range(old.full_span.clone(), [])),
            (None, Some(new)) => replacements.push(Replacement::insert(
                wire.opening_insert,
                format!(" lang=\"{}\"", escape(new)).into_bytes(),
            )),
            (None, None) | (Some(_), Some(_)) => {},
        }
    }
    encode_narration(original, working, replacements)
}

fn encode_narration(
    original: &Snapshot,
    working: &MediaMetadata,
    mut replacements: Vec<Replacement>,
) -> Result<Vec<u8>> {
    match (&original.layout.narration, working.narration) {
        (Some(wire), Some(value)) => match (&wire.value, value) {
            (Some(old), new) => {
                let encoded = if new { "true" } else { "false" };
                if old.value != encoded {
                    replacements.push(Replacement::value(old, encoded));
                }
            },
            (None, value) => replacements.push(Replacement::insert(
                wire.opening_insert,
                format!(" val=\"{}\"", if value { "true" } else { "false" }).into_bytes(),
            )),
        },
        (Some(wire), None) => {
            if let Some(value) = &wire.value {
                replacements.push(Replacement::range(value.full_span.clone(), []));
            }
        },
        (None, None) => {},
        (None, Some(_)) => {
            return Err(invalid(
                "cannot create isNarration in a bounded transaction",
            ));
        },
    }
    apply_replacements(original.source_xml.as_slice(), replacements)
}

#[derive(Debug)]
struct Replacement {
    range: Range<usize>,
    value: Vec<u8>,
}

impl Replacement {
    fn value(attribute: &Attr, value: &str) -> Self {
        Self::range(attribute.span.clone(), escape(value).into_bytes())
    }

    fn insert(offset: usize, value: Vec<u8>) -> Self {
        Self {
            range: offset..offset,
            value,
        }
    }

    fn range(range: Range<usize>, value: impl Into<Vec<u8>>) -> Self {
        Self {
            range,
            value: value.into(),
        }
    }
}

fn apply_replacements(source: &[u8], mut replacements: Vec<Replacement>) -> Result<Vec<u8>> {
    replacements.sort_by(|left, right| right.range.start.cmp(&left.range.start));
    let mut output = source.to_vec();
    let mut upper = source.len();
    for replacement in replacements {
        if replacement.range.start > replacement.range.end
            || replacement.range.end > source.len()
            || replacement.range.end > upper
        {
            return Err(invalid(
                "media tracks patch ranges overlap or escape the source",
            ));
        }
        output.splice(replacement.range.clone(), replacement.value);
        upper = replacement.range.start;
    }
    Ok(output)
}

fn fingerprint(bytes: &[u8]) -> Revision {
    let mut hash = 0xcbf29ce484222325u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x100000001b3);
    }
    hash
}

fn escape(value: &str) -> String {
    crate::presentation_properties::metadata::escape_xml(value)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
