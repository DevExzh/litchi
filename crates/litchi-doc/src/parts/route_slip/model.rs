//! Semantic Metadata values and their validation rules.

use crate::package::{Error as PackageError, Result};

/// The protection level applied while a document is being routed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum Protection {
    /// No protection.
    Off = 0,
    /// Revision marking cannot be disabled and changes cannot be accepted or
    /// rejected.
    RevisionMark = 1,
    /// Users may add comments but cannot change document content.
    Annotation = 2,
    /// Users may edit form fields and unprotected sections only.
    Form = 3,
}

/// A document-change family described by route-slip protection.
///
/// This is a policy projection only. The DOC crate does not authenticate a
/// caller, execute a route, or silently rewrite `DopBase` or range-protection
/// records when this value is queried.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EditKind {
    /// An unrestricted content edit.
    Content,
    /// A change-tracked revision operation.
    Revision,
    /// An annotation/comment operation.
    Annotation,
    /// An edit confined to a form field or unprotected section.
    FormField,
}

impl Protection {
    pub(crate) const fn raw(self) -> u16 {
        self as u16
    }

    /// Whether this route policy describes the supplied document-change kind.
    #[must_use]
    pub const fn allows(self, change: EditKind) -> bool {
        match self {
            Self::Off => true,
            Self::RevisionMark => matches!(change, EditKind::Revision),
            Self::Annotation => matches!(change, EditKind::Annotation),
            Self::Form => matches!(change, EditKind::FormField),
        }
    }
}

/// The routing order for recipients.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(i16)]
pub enum DeliveryOption {
    /// Send the document to recipients one at a time.
    Serial = 0,
    /// Send the document to all recipients at once.
    Parallel = 1,
}

impl DeliveryOption {
    pub(crate) const fn raw(self) -> i16 {
        self as i16
    }
}

/// A lossless narrow/ANSI byte string.
///
/// The bytes are deliberately not decoded. MS-DOC stores these fields in the
/// system code page of the writer, which is not available from the payload
/// itself and need not be UTF-8. The length and semantic limits are enforced
/// by the owning Metadata field or recipient record.
#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
pub struct NarrowString(Vec<u8>);

impl NarrowString {
    /// Wrap raw narrow bytes without decoding or normalizing them.
    pub fn new(bytes: impl Into<Vec<u8>>) -> Self {
        Self(bytes.into())
    }

    /// Borrow the exact bytes stored in this value.
    #[inline]
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.0
    }

    /// Consume this value and return its exact bytes.
    #[inline]
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.0
    }

    /// Return the byte length of this narrow string.
    #[inline]
    #[must_use]
    pub fn len(&self) -> usize {
        self.0.len()
    }

    /// Whether this narrow string contains no bytes.
    #[inline]
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.0.is_empty()
    }
}

impl AsRef<[u8]> for NarrowString {
    #[inline]
    fn as_ref(&self) -> &[u8] {
        self.as_bytes()
    }
}

impl From<Vec<u8>> for NarrowString {
    fn from(bytes: Vec<u8>) -> Self {
        Self::new(bytes)
    }
}

/// One `Recipient` recipient record (MS-DOC 2.9.233).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Recipient {
    /// Opaque recipient identifier bytes (`rgbEntryId`).
    pub entry_id: Vec<u8>,
    /// Lossless narrow recipient name or e-mail alias (`szName`).
    pub name: NarrowString,
}

impl Recipient {
    /// Construct a recipient record and validate its signed length fields.
    pub fn try_new(entry_id: Vec<u8>, name: NarrowString) -> Result<Self> {
        let value = Self { entry_id, name };
        value.validate()?;
        Ok(value)
    }

    /// Validate the values that will be written as `Recipient`.
    pub fn validate(&self) -> Result<()> {
        if self.entry_id.len() > i16::MAX as usize {
            return Err(corrupted("Recipient entry ID exceeds i16::MAX bytes"));
        }
        if self.name.is_empty() {
            return Err(corrupted("Recipient name must not be empty"));
        }
        if self.name.len() > i16::MAX as usize {
            return Err(corrupted("Recipient name exceeds i16::MAX bytes"));
        }
        Ok(())
    }
}

/// A complete `Metadata` structure (MS-DOC 2.9.232).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Metadata {
    /// Whether the document was sent out for review (`fRouted`).
    pub routed: bool,
    /// Whether the original document is returned after routing
    /// (`fReturnOrig`).
    pub return_original: bool,
    /// Whether status-tracking mail is sent (`fTrackStatus`).
    pub track_status: bool,
    /// Protection applied during routing (`nProtect`).
    pub protection: Protection,
    /// Zero-based current recipient index (`iStage`).
    pub stage: u16,
    /// Recipient delivery order (`delOption`).
    pub delivery: DeliveryOption,
    /// Subject bytes (`szSubject`).
    pub subject: NarrowString,
    /// Message bytes (`szMessage`).
    pub message: NarrowString,
    /// Status bytes (`szStatus`).
    pub status: NarrowString,
    /// Title bytes (`szTitle`).
    pub title: NarrowString,
    /// Ordered recipient records (`rgRouteSlips`).
    pub recipients: Vec<Recipient>,
}

impl Metadata {
    /// Construct a route slip and validate all encoded bounds.
    #[allow(
        clippy::too_many_arguments,
        reason = "matches the fixed MS-DOC structure"
    )]
    pub fn try_new(
        routed: bool,
        return_original: bool,
        track_status: bool,
        protection: Protection,
        stage: u16,
        delivery: DeliveryOption,
        subject: NarrowString,
        message: NarrowString,
        status: NarrowString,
        title: NarrowString,
        recipients: Vec<Recipient>,
    ) -> Result<Self> {
        let value = Self {
            routed,
            return_original,
            track_status,
            protection,
            stage,
            delivery,
            subject,
            message,
            status,
            title,
            recipients,
        };
        value.validate()?;
        Ok(value)
    }

    /// Validate all Metadata scalar, count, and length constraints.
    pub fn validate(&self) -> Result<()> {
        if self.recipients.is_empty() {
            return Err(corrupted("Metadata must contain at least one recipient"));
        }
        if self.recipients.len() > i16::MAX as usize {
            return Err(corrupted("Metadata recipient count exceeds i16::MAX"));
        }
        if usize::from(self.stage) >= self.recipients.len() {
            return Err(corrupted(
                "Metadata stage must be less than the recipient count",
            ));
        }
        validate_short_string(&self.subject, "szSubject")?;
        validate_short_string(&self.message, "szMessage")?;
        validate_short_string(&self.status, "szStatus")?;
        validate_short_string(&self.title, "szTitle")?;
        for (index, recipient) in self.recipients.iter().enumerate() {
            recipient
                .validate()
                .map_err(|error| with_recipient_context(error, index))?;
        }
        Ok(())
    }

    /// Number of recipient records encoded by this route slip.
    #[inline]
    #[must_use]
    pub fn recipient_count(&self) -> usize {
        self.recipients.len()
    }
}

pub(crate) fn validate_short_string(value: &NarrowString, field: &str) -> Result<()> {
    if value.len() >= 256 {
        return Err(corrupted(format!(
            "{field} must contain fewer than 256 ANSI bytes"
        )));
    }
    Ok(())
}

fn with_recipient_context(error: PackageError, index: usize) -> PackageError {
    match error {
        PackageError::Corrupted(message) => {
            corrupted(format!("Metadata recipient {index}: {message}"))
        },
        other => other,
    }
}

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}
