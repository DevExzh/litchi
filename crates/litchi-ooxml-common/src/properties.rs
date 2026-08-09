//! Core document properties shared by every OOXML host.
//!
//! [`Props`] is the concise, owned semantic value. Package operations are
//! deliberately separate: [`read`] distinguishes an absent part from a present
//! but empty properties document, [`write()`] consumes a value, and [`clear`]
//! removes the validated package graph idempotently.

mod graph;
pub mod keyword;
mod read;
pub mod time;
mod write;

pub use read::read;
pub use write::{clear, write};

use crate::Result;
use litchi_core::Metadata;
use litchi_opc::OpcPackage;

use keyword::Item;

pub(super) const CORE_PROPERTIES_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
pub(super) const STRICT_CORE_PROPERTIES_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/package/metadata/core-properties";
pub(super) const STRICT_CORE_PROPERTIES_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/package/relationships/metadata/core-properties";
pub(super) const MAX_PROPERTY_TEXT: usize = 1_048_576;
pub(super) const MAX_XML_BYTES: usize = 20 * 1_048_576;
pub(super) const MAX_XML_EVENTS: usize = 131_072;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum Dialect {
    Transitional,
    Strict,
}

impl Dialect {
    pub(super) fn namespace(self) -> &'static str {
        match self {
            Self::Transitional => {
                "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
            },
            Self::Strict => "http://purl.oclc.org/ooxml/package/metadata/core-properties",
        }
    }
}

/// Lossless mixed content for the `cp:keywords` core property.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Keywords {
    /// Optional language for the mixed keyword content.
    pub lang: Option<keyword::Lang>,
    /// Text and `cp:value` children in document order.
    pub items: Vec<Item>,
}

impl Keywords {
    /// Creates a present but empty keyword value.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Adds text directly inside `cp:keywords`.
    #[must_use]
    pub fn text(mut self, text: impl Into<String>) -> Self {
        self.append_text(text.into());
        self
    }

    /// Adds a structured `cp:value` child.
    #[must_use]
    pub fn value(mut self, value: impl Into<keyword::Value>) -> Self {
        self.items.push(Item::Value(value.into()));
        self
    }

    /// Adds a validated language to the outer `cp:keywords` element.
    #[must_use]
    pub fn lang(mut self, lang: keyword::Lang) -> Self {
        self.lang = Some(lang);
        self
    }

    /// Concatenates all text and value content in document order.
    #[must_use]
    pub fn joined(&self) -> String {
        let capacity = self
            .items
            .iter()
            .map(|item| match item {
                Item::Text(text) => text.len(),
                Item::Value(value) => value.text.len(),
            })
            .sum();
        let mut joined = String::with_capacity(capacity);
        for item in &self.items {
            match item {
                Item::Text(text) => joined.push_str(text),
                Item::Value(value) => joined.push_str(&value.text),
            }
        }
        joined
    }

    /// Returns a direct borrow for the common plain-text form.
    #[must_use]
    pub fn plain(&self) -> Option<&str> {
        if self.lang.is_some() || self.items.len() != 1 {
            return None;
        }
        match &self.items[0] {
            Item::Text(text) => Some(text),
            Item::Value(_) => None,
        }
    }

    pub(super) fn append_text(&mut self, text: String) {
        if text.is_empty() {
            return;
        }
        if let Some(Item::Text(previous)) = self.items.last_mut() {
            previous.push_str(&text);
        } else {
            self.items.push(Item::Text(text));
        }
    }
}

impl From<String> for Keywords {
    fn from(text: String) -> Self {
        Self::new().text(text)
    }
}

impl From<&str> for Keywords {
    fn from(text: &str) -> Self {
        Self::new().text(text)
    }
}

impl std::fmt::Display for Keywords {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(&self.joined())
    }
}

/// Owned OOXML core properties.
///
/// Public fields make struct-update syntax natural while package writes still
/// validate every string before changing the package.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Props {
    /// Document title.
    pub title: Option<String>,
    /// Document subject.
    pub subject: Option<String>,
    /// Document creator or author.
    pub creator: Option<String>,
    /// Lossless keyword mixed content; plain strings remain the common form.
    pub keywords: Option<Keywords>,
    /// Document description.
    pub description: Option<String>,
    /// Stable document identifier.
    pub identifier: Option<String>,
    /// Person or tool that last modified the document.
    pub last_modified_by: Option<String>,
    /// Document category.
    pub category: Option<String>,
    /// Content status, such as `Draft` or `Final`.
    pub content_status: Option<String>,
    /// Revision lexical value. The OPC schema deliberately permits any string.
    pub revision: Option<String>,
    /// Document version.
    pub version: Option<String>,
    /// Document language.
    pub language: Option<String>,
    /// Lossless W3CDTF creation value.
    pub created: Option<time::W3c>,
    /// Lossless W3CDTF last-modification value.
    pub modified: Option<time::W3c>,
    /// Lossless `xsd:dateTime` last-printed value.
    pub last_printed: Option<time::DateTime>,
}

impl Props {
    /// Creates a present but semantically empty properties document.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Sets the document title.
    #[must_use]
    pub fn title(mut self, title: impl Into<String>) -> Self {
        self.title = Some(title.into());
        self
    }

    /// Sets the document subject.
    #[must_use]
    pub fn subject(mut self, subject: impl Into<String>) -> Self {
        self.subject = Some(subject.into());
        self
    }

    /// Sets the document creator or author.
    #[must_use]
    pub fn creator(mut self, creator: impl Into<String>) -> Self {
        self.creator = Some(creator.into());
        self
    }

    /// Sets the document keywords.
    #[must_use]
    pub fn keywords(mut self, keywords: impl Into<Keywords>) -> Self {
        self.keywords = Some(keywords.into());
        self
    }

    /// Sets the document description.
    #[must_use]
    pub fn description(mut self, description: impl Into<String>) -> Self {
        self.description = Some(description.into());
        self
    }

    /// Sets the stable document identifier.
    #[must_use]
    pub fn identifier(mut self, identifier: impl Into<String>) -> Self {
        self.identifier = Some(identifier.into());
        self
    }

    /// Sets who last modified the document.
    #[must_use]
    pub fn last_modified_by(mut self, name: impl Into<String>) -> Self {
        self.last_modified_by = Some(name.into());
        self
    }

    /// Sets the document category.
    #[must_use]
    pub fn category(mut self, category: impl Into<String>) -> Self {
        self.category = Some(category.into());
        self
    }

    /// Sets the content status.
    #[must_use]
    pub fn content_status(mut self, status: impl Into<String>) -> Self {
        self.content_status = Some(status.into());
        self
    }

    /// Sets the revision lexical value without numeric normalization.
    #[must_use]
    pub fn revision(mut self, revision: impl Into<String>) -> Self {
        self.revision = Some(revision.into());
        self
    }

    /// Sets the document version.
    #[must_use]
    pub fn version(mut self, version: impl Into<String>) -> Self {
        self.version = Some(version.into());
        self
    }

    /// Sets the document language.
    #[must_use]
    pub fn language(mut self, language: impl Into<String>) -> Self {
        self.language = Some(language.into());
        self
    }

    /// Sets a validated W3CDTF creation value.
    #[must_use]
    pub fn created(mut self, created: impl Into<time::W3c>) -> Self {
        self.created = Some(created.into());
        self
    }

    /// Sets a validated W3CDTF modification value.
    #[must_use]
    pub fn modified(mut self, modified: impl Into<time::W3c>) -> Self {
        self.modified = Some(modified.into());
        self
    }

    /// Sets a validated `xsd:dateTime` last-printed value.
    #[must_use]
    pub fn last_printed(mut self, last_printed: impl Into<time::DateTime>) -> Self {
        self.last_printed = Some(last_printed.into());
        self
    }

    /// Encodes a standalone Transitional core-properties document.
    ///
    /// Prefer [`write()`] when an OPC package is available because it preserves
    /// an existing Strict package's namespace and relationship family.
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn xml(&self) -> Result<String> {
        write::encode(self, Dialect::Transitional)
    }
}

impl From<Props> for Metadata {
    fn from(props: Props) -> Self {
        let created = props.created.as_ref().and_then(time::W3c::utc);
        let created_local = props.created.as_ref().and_then(time::W3c::local);
        let modified = props.modified.as_ref().and_then(time::W3c::utc);
        let modified_local = props.modified.as_ref().and_then(time::W3c::local);
        let last_printed_time = props.last_printed.as_ref().and_then(time::DateTime::utc);
        let last_printed_local = props.last_printed.as_ref().and_then(time::DateTime::local);
        Self {
            title: props.title,
            subject: props.subject,
            author: props.creator,
            keywords: props.keywords.map(|keywords| keywords.joined()),
            description: props.description,
            identifier: props.identifier,
            language: props.language,
            last_modified_by: props.last_modified_by,
            revision: props.revision,
            category: props.category,
            content_status: props.content_status,
            version: props.version,
            created,
            created_local,
            modified,
            modified_local,
            last_printed_time,
            last_printed_local,
            ..Self::default()
        }
    }
}

/// A host package's authoritative, mutation-tracked core-properties slot.
///
/// This is public so host crates can share the implementation, but ordinary
/// callers should use their DOCX, PPTX, or XLSX package facade.
#[doc(hidden)]
#[derive(Debug)]
pub struct Slot {
    props: Option<Props>,
    dirty: bool,
}

/// A package-staged properties edit tied to its originating slot.
///
/// Dropping this guard keeps the slot dirty. Consuming it through
/// [`Stage::commit`] is the only way to clear the edit intent.
#[doc(hidden)]
#[must_use = "commit only after the staged package has been published, or drop to retry"]
pub struct Stage<'a> {
    slot: &'a mut Slot,
    changed: bool,
}

impl Stage<'_> {
    /// Reports whether staging changed package bytes or graph state.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.changed
    }

    /// Marks the originating slot clean after successful publication.
    pub fn commit(self) {
        self.slot.dirty = false;
    }
}

impl Slot {
    /// Reads and validates the package graph while retaining absence exactly.
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn load(package: &OpcPackage) -> Result<Self> {
        Ok(Self {
            props: read(package)?,
            dirty: false,
        })
    }

    /// Borrows the current value without marking it dirty.
    #[must_use]
    pub fn get(&self) -> Option<&Props> {
        self.props.as_ref()
    }

    /// Mutably borrows an existing value and marks it for validation on flush.
    pub fn get_mut(&mut self) -> Option<&mut Props> {
        let props = self.props.as_mut()?;
        self.dirty = true;
        Some(props)
    }

    /// Moves a present value into the slot and returns the previous value.
    pub fn put(&mut self, props: Props) -> Option<Props> {
        self.dirty = true;
        self.props.replace(props)
    }

    /// Marks the value absent and moves out the previous value.
    pub fn clear(&mut self) -> Option<Props> {
        let previous = self.props.take();
        self.dirty |= previous.is_some();
        previous
    }

    /// Applies a staged write or clear only when this slot was edited.
    ///
    /// A failed flush leaves the slot dirty so the caller can repair the value
    /// and retry. A byte-identical write is a true no-op.
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn flush(&mut self, package: &mut OpcPackage) -> Result<bool> {
        let staged = self.stage(package)?;
        let changed = staged.changed();
        staged.commit();
        Ok(changed)
    }

    /// Applies this slot to a staging package without clearing edit intent.
    ///
    /// Hosts use this before a fallible publication step and call
    /// [`Stage::commit`] only after publication succeeds. The returned guard
    /// is tied to this exact slot, so another slot cannot be committed by
    /// mistake.
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn stage(&mut self, package: &mut OpcPackage) -> Result<Stage<'_>> {
        if !self.dirty {
            return Ok(Stage {
                slot: self,
                changed: false,
            });
        }
        let changed = match self.props.as_ref() {
            Some(props) => write::sync(package, props)?,
            None => clear(package)?,
        };
        Ok(Stage {
            slot: self,
            changed,
        })
    }

    /// Returns whether the host must flush this slot before publication.
    #[must_use]
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn builder_and_struct_update_stay_concise() {
        let first = Props::new()
            .title("Test Document")
            .creator("Ada")
            .revision("007-alpha");
        let second = Props {
            title: Some("Revised".to_owned()),
            ..first
        };

        assert_eq!(second.title.as_deref(), Some("Revised"));
        assert_eq!(second.creator.as_deref(), Some("Ada"));
        assert_eq!(second.revision.as_deref(), Some("007-alpha"));
    }

    #[test]
    fn xml_escapes_text_and_rejects_forbidden_characters() {
        let xml = Props::new()
            .title("A & <B> \"C\"")
            .xml()
            .expect("valid XML");
        assert!(xml.contains("A &amp; &lt;B&gt; &quot;C&quot;"));

        let error = Props::new().title("bad\u{0}text").xml().unwrap_err();
        assert!(matches!(error, crate::Error::Invalid(_)));
    }

    #[test]
    fn slot_tracks_only_real_edit_intent() {
        let package = OpcPackage::new();
        let mut slot = Slot::load(&package).expect("load");
        assert!(slot.get().is_none());
        assert!(!slot.is_dirty());
        assert!(slot.clear().is_none());
        assert!(!slot.is_dirty());

        slot.put(Props::new());
        assert!(slot.is_dirty());
        assert!(slot.get_mut().is_some());

        let mut staged = OpcPackage::new();
        let pending = slot.stage(&mut staged).expect("stage");
        assert!(pending.changed());
        drop(pending);
        assert!(slot.is_dirty());
        slot.stage(&mut staged).expect("restage").commit();
        assert!(!slot.is_dirty());
    }
}
