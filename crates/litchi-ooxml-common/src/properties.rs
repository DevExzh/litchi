//! Document properties shared by every OOXML host: the core
//! (`docProps/core.xml`), application (`docProps/app.xml`) and custom
//! (`docProps/custom.xml`) families.
//!
//! The semantic value is the one cross-format [`Metadata`]. Each family's
//! lexical form stays inside its codec:
//!
//! - the core reader keeps the exact `cp:keywords` mixed content, the
//!   W3CDTF and `xsd:dateTime` lexical forms and the Strict or Transitional
//!   dialect, so a field the caller leaves unchanged is written back exactly;
//! - [`Metadata::keywords`] is the keyword list of that mixed content: text
//!   split at `;` and `,`, each `cp:value` one keyword
//!   ([`litchi_core::metadata::Keywords::list`]). An unchanged list keeps the
//!   source element byte for byte; a changed list is written as text joined
//!   with `, `, a keyword holding a separator as a `cp:value` child
//!   ([`litchi_core::metadata::Keywords::from_list`]);
//! - the application writer ([`app`]) splices only the elements whose value
//!   changes, so the vector properties and extension content survive;
//! - custom properties keep their own typed collection
//!   ([`crate::custom::Props`]).
//!
//! [`read_with_report`] reads the core family alone and distinguishes an
//! absent part from a present but empty one. [`Edit::set_metadata`] stores the
//! core fields of a [`Metadata`] and refuses any field core properties cannot
//! carry. [`Edit::set_custom_props`] preserves a present raw part when its
//! equal semantic collection is staged; [`Edit::remove_custom_props`] is the
//! explicit graph-removal intent. Hosts edit all three families at once through the shared
//! transaction ([`Edit`], [`Commit`], [`Patch`], read through [`Snapshot`]), which each
//! format re-exports as its `metadata` module.

pub mod app;
mod graph;
pub(crate) mod package_graph;
mod read;
pub mod time;
mod transaction;
mod write;

#[cfg(test)]
mod transaction_tests;

pub(crate) use litchi_core::metadata::{Keywords, keywords as keyword};
pub(crate) use read::{
    read_core_prepared_with_report, read_core_source_backed_with_report, read_core_with_report,
};
pub use transaction::{
    Change, Commit, Diagnostics, Edit, EditOptions, Error, Family, FamilyAdmission, LimitKind,
    OwnerAdmission, Patch, Plan, Snapshot, publish_core_metadata,
};
#[cfg(test)]
pub(crate) use write::write_core;

use crate::Result;
use crate::leniency::ToleranceReport;
use litchi_core::metadata::{Field, Metadata, Timestamp};
use litchi_opc::{OpcPackage, SourceBackedPackage};

pub(super) const CORE_PROPERTIES_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
pub(super) const STRICT_CORE_PROPERTIES_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/package/metadata/core-properties";
pub(super) const STRICT_CORE_PROPERTIES_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/package/relationships/metadata/core-properties";
pub(super) const MAX_PROPERTY_TEXT: usize = litchi_core::metadata::keywords::MAX_TEXT_BYTES;
pub(super) const MAX_XML_BYTES: usize = 20 * 1_048_576;
pub(super) const MAX_XML_EVENTS: usize = 131_072;

/// The [`Metadata`] fields the core-properties part stores.
pub const CORE_FIELDS: [Field; 15] = [
    Field::Title,
    Field::Subject,
    Field::Author,
    Field::Keywords,
    Field::Description,
    Field::Identifier,
    Field::Language,
    Field::LastModifiedBy,
    Field::Revision,
    Field::Category,
    Field::ContentStatus,
    Field::Version,
    Field::Created,
    Field::Modified,
    Field::LastPrinted,
];

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Dialect {
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

pub(super) fn collapse_xml_whitespace(value: &str) -> String {
    value
        .split([' ', '\t', '\r', '\n'])
        .filter(|segment| !segment.is_empty())
        .collect::<Vec<_>>()
        .join(" ")
}

/// The lossless codec record of one core-properties part.
///
/// The semantic [`Metadata`] projection is [`Core::metadata`]; a write merges
/// a [`Metadata`] onto the record it was read from ([`Core::merged`]) so the
/// lexical form of every unchanged field survives.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct Core {
    pub(crate) title: Option<String>,
    pub(crate) subject: Option<String>,
    /// `dc:creator`: [`Metadata::author`].
    pub(crate) creator: Option<String>,
    pub(crate) keywords: Option<Keywords>,
    pub(crate) description: Option<String>,
    pub(crate) identifier: Option<String>,
    pub(crate) last_modified_by: Option<String>,
    pub(crate) category: Option<String>,
    pub(crate) content_status: Option<String>,
    pub(crate) revision: Option<String>,
    pub(crate) version: Option<String>,
    pub(crate) language: Option<String>,
    pub(crate) created: Option<time::W3c>,
    pub(crate) modified: Option<time::W3c>,
    pub(crate) last_printed: Option<time::DateTime>,
}

impl Core {
    /// The semantic projection.
    pub(crate) fn metadata(&self) -> Metadata {
        let mut metadata = Metadata::default();
        self.project_into(&mut metadata);
        metadata
    }

    /// Overwrite the core fields of `metadata` with this record's values.
    pub(crate) fn project_into(&self, metadata: &mut Metadata) {
        metadata.title.clone_from(&self.title);
        metadata.subject.clone_from(&self.subject);
        metadata.author.clone_from(&self.creator);
        metadata.keywords = self
            .keywords
            .as_ref()
            .map(Keywords::list)
            .unwrap_or_default();
        metadata.description.clone_from(&self.description);
        metadata.identifier.clone_from(&self.identifier);
        metadata.language.clone_from(&self.language);
        metadata.last_modified_by.clone_from(&self.last_modified_by);
        metadata.revision.clone_from(&self.revision);
        metadata.category.clone_from(&self.category);
        metadata.content_status.clone_from(&self.content_status);
        metadata.version.clone_from(&self.version);
        metadata.created = self.created.as_ref().and_then(w3c_timestamp);
        metadata.modified = self.modified.as_ref().and_then(w3c_timestamp);
        metadata.last_printed = self.last_printed.as_ref().and_then(date_time_timestamp);
    }

    /// Merge the core fields of `target` onto `base`.
    ///
    /// A field whose semantic value equals `base`'s projection keeps `base`'s
    /// exact lexical form (keyword mixed content, a W3CDTF with an offset or
    /// a reduced precision). A changed field takes `target`'s value in the
    /// canonical lexical form.
    pub(crate) fn merged(base: Option<&Self>, target: &Metadata) -> Result<Self> {
        let empty = Self::default();
        let base = base.unwrap_or(&empty);
        let before = base.metadata();
        fn text(
            base: &Option<String>,
            before: &Option<String>,
            target: &Option<String>,
        ) -> Option<String> {
            if before == target {
                base.clone()
            } else {
                target.clone()
            }
        }
        // The list is the split of the mixed content (`Keywords::list`), so
        // an unchanged list keeps the source's exact separators, values and
        // languages; a changed one is written by `Keywords::from_list`.
        let keywords = if before.keywords == target.keywords {
            base.keywords.clone()
        } else if target.keywords.is_empty() {
            None
        } else {
            Some(Keywords::from_list(&target.keywords))
        };
        let created = if before.created == target.created {
            base.created.clone()
        } else {
            target.created.map(timestamp_w3c).transpose()?
        };
        let modified = if before.modified == target.modified {
            base.modified.clone()
        } else {
            target.modified.map(timestamp_w3c).transpose()?
        };
        let last_printed = if before.last_printed == target.last_printed {
            base.last_printed.clone()
        } else {
            target.last_printed.map(timestamp_date_time).transpose()?
        };
        Ok(Self {
            title: text(&base.title, &before.title, &target.title),
            subject: text(&base.subject, &before.subject, &target.subject),
            creator: text(&base.creator, &before.author, &target.author),
            keywords,
            description: text(&base.description, &before.description, &target.description),
            identifier: text(&base.identifier, &before.identifier, &target.identifier),
            last_modified_by: text(
                &base.last_modified_by,
                &before.last_modified_by,
                &target.last_modified_by,
            ),
            category: text(&base.category, &before.category, &target.category),
            content_status: text(
                &base.content_status,
                &before.content_status,
                &target.content_status,
            ),
            revision: text(&base.revision, &before.revision, &target.revision),
            version: text(&base.version, &before.version, &target.version),
            language: text(&base.language, &before.language, &target.language),
            created,
            modified,
            last_printed,
        })
    }
}

#[cfg(test)]
macro_rules! core_text_builders {
    ($($name:ident),* $(,)?) => {
        $(
            pub(crate) fn $name(mut self, value: impl Into<String>) -> Self {
                self.$name = Some(value.into());
                self
            }
        )*
    };
}

#[cfg(test)]
#[allow(
    dead_code,
    reason = "a complete fixture builder; each test uses a subset"
)]
impl Core {
    pub(crate) fn new() -> Self {
        Self::default()
    }

    core_text_builders!(
        title,
        subject,
        creator,
        description,
        identifier,
        last_modified_by,
        category,
        content_status,
        revision,
        version,
        language,
    );

    pub(crate) fn keywords(mut self, keywords: impl Into<Keywords>) -> Self {
        self.keywords = Some(keywords.into());
        self
    }

    pub(crate) fn created(mut self, created: impl Into<time::W3c>) -> Self {
        self.created = Some(created.into());
        self
    }

    pub(crate) fn modified(mut self, modified: impl Into<time::W3c>) -> Self {
        self.modified = Some(modified.into());
        self
    }

    pub(crate) fn last_printed(mut self, last_printed: impl Into<time::DateTime>) -> Self {
        self.last_printed = Some(last_printed.into());
        self
    }

    #[cfg(test)]
    pub(crate) fn xml(&self) -> Result<String> {
        write::encode(self, Dialect::Transitional)
    }
}

fn w3c_timestamp(value: &time::W3c) -> Option<Timestamp> {
    value
        .utc()
        .map(Timestamp::Utc)
        .or_else(|| value.local().map(Timestamp::Unzoned))
}

fn date_time_timestamp(value: &time::DateTime) -> Option<Timestamp> {
    value
        .utc()
        .map(Timestamp::Utc)
        .or_else(|| value.local().map(Timestamp::Unzoned))
}

fn timestamp_w3c(value: Timestamp) -> Result<time::W3c> {
    match value {
        Timestamp::Utc(value) => Ok(time::W3c::from(value)),
        Timestamp::Unzoned(value) => time::W3c::new(Timestamp::Unzoned(value).to_string()),
    }
}

fn timestamp_date_time(value: Timestamp) -> Result<time::DateTime> {
    match value {
        Timestamp::Utc(value) => Ok(time::DateTime::from(value)),
        Timestamp::Unzoned(value) => time::DateTime::new(Timestamp::Unzoned(value).to_string()),
    }
}

/// Refuse a [`Metadata`] value holding a field core properties cannot store.
#[cfg(test)]
fn ensure_core_only(metadata: &Metadata) -> Result<()> {
    match Field::ALL
        .into_iter()
        .find(|field| metadata.is_set(*field) && !CORE_FIELDS.contains(field))
    {
        Some(field) => Err(crate::Error::Invalid(format!(
            "metadata field `{field}` is not stored in core properties; edit it through the package's metadata transaction"
        ))),
        None => Ok(()),
    }
}

/// Refuse a core-only semantic transaction carrying an application field.
fn ensure_core_only_transaction(metadata: &Metadata) -> std::result::Result<(), Error> {
    match Field::ALL
        .into_iter()
        .find(|field| metadata.is_set(*field) && !CORE_FIELDS.contains(field))
    {
        Some(_) => Err(Error::Refused(litchi_core::edit::Refusal::new(
            litchi_core::edit::RefusalKind::UnsupportedProperty,
            "set_core_metadata",
            "the OOXML metadata route does not support this property",
        ))),
        None => Ok(()),
    }
}

/// Read the relationship-selected core properties of an OOXML package,
/// appending every reference-compatible repair to `report`.
///
/// Absence is `Ok(None)`, distinct from a present but empty part. Only the
/// core fields of the returned [`Metadata`] are set. Producer-rule
/// deviations in the part body (OPC M4.3-M4.5, duplicates, unknown elements,
/// invalid dates, non-UTF-8 bytes, a dialect mismatch) and in the
/// relationship graph are repaired per property as [`crate::leniency`]
/// documents, and each repair is recorded in `report`: a conformant package
/// appends nothing. There is deliberately no report-less variant. Ill-formed
/// XML, the Markup Compatibility namespace (M4.2), and limits still fail.
///
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a
/// configured bound, or an underlying XML or package operation fails.
pub fn read_with_report(
    package: &OpcPackage,
    report: &mut ToleranceReport,
) -> Result<Option<Metadata>> {
    Ok(read_core_with_report(package, report)?.map(|core| core.metadata()))
}

/// [`read_with_report`] over a source-backed package: the catalog and
/// relationship graph stay metadata-only until the selected part is read,
/// with the source version and execution checks around it.
///
/// # Errors
///
/// The errors of [`read_with_report`], and a source-change or cancellation
/// error.
pub fn read_source_backed_with_report(
    package: &SourceBackedPackage,
    report: &mut ToleranceReport,
) -> Result<Option<Metadata>> {
    Ok(read_core_source_backed_with_report(package, report)?.map(|core| core.metadata()))
}

/// [`read_with_report`] over a retained effective graph.
///
/// # Errors
///
/// Returns graph, XML, resource, source or execution errors.
pub fn read_prepared_with_report(
    package: &litchi_opc::PreparedTopology,
    report: &mut ToleranceReport,
) -> Result<Option<Metadata>> {
    Ok(read_core_prepared_with_report(package, report)?.map(|core| core.metadata()))
}

/// A host package's authoritative, mutation-tracked core-properties slot.
///
/// This is public so host crates can share the implementation, but ordinary
/// callers should use their DOCX, PPTX, or XLSX package facade. The slot
/// holds the semantic [`Metadata`] and the codec record it was read from,
/// so a flush keeps the lexical form of every unchanged field.
#[doc(hidden)]
#[derive(Debug, Clone)]
pub struct Slot {
    #[cfg(test)]
    source: Option<Core>,
    metadata: Option<Metadata>,
    dirty: bool,
    report: ToleranceReport,
}

/// A package-staged properties edit tied to its originating slot.
///
/// Dropping this guard keeps the slot dirty. Consuming it through
/// [`Stage::commit`] is the only way to clear the edit intent.
#[doc(hidden)]
#[must_use = "commit only after the staged package has been published, or drop to retry"]
#[cfg(test)]
pub struct Stage<'a> {
    slot: &'a mut Slot,
    changed: bool,
    staged: Option<Core>,
}

#[cfg(test)]
impl Stage<'_> {
    /// Reports whether staging changed package bytes or graph state.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.changed
    }

    /// Marks the originating slot clean after successful publication.
    pub fn commit(self) {
        self.slot.dirty = false;
        self.slot.source = self.staged;
    }
}

impl Slot {
    /// Reads and validates the package graph while retaining absence exactly.
    ///
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn load(package: &OpcPackage) -> Result<Self> {
        let mut report = ToleranceReport::default();
        let source = read_core_with_report(package, &mut report)?;
        let metadata = source.as_ref().map(Core::metadata);
        Ok(Self {
            #[cfg(test)]
            source,
            metadata,
            dirty: false,
            report,
        })
    }

    /// The slot of a package whose core part this crate has just written
    /// from `metadata` (`None`: the part is absent), without decoding it
    /// again (record 0377a). The writer's output conforms, so the tolerance
    /// report is empty; debug builds check the round trip against `package`.
    #[must_use]
    pub fn published(package: &OpcPackage, metadata: Option<Metadata>) -> Self {
        #[cfg(debug_assertions)]
        {
            let loaded = Self::load(package).ok();
            debug_assert!(
                loaded
                    .as_ref()
                    .is_some_and(|slot| slot.metadata == metadata && slot.report.is_clean()),
                "a published core part must decode to the metadata it was written from"
            );
        }
        #[cfg(not(debug_assertions))]
        let _ = package;
        Self {
            #[cfg(test)]
            source: None,
            metadata,
            dirty: false,
            report: ToleranceReport::default(),
        }
    }

    /// Reference-compatible repairs applied when this slot was loaded (see
    /// [`crate::leniency`]); empty when the core properties conformed.
    #[must_use]
    pub fn tolerance_report(&self) -> &ToleranceReport {
        &self.report
    }

    /// Borrows the current value without marking it dirty.
    #[must_use]
    pub fn get(&self) -> Option<&Metadata> {
        self.metadata.as_ref()
    }

    /// Mutably borrows an existing value and marks it for validation on flush.
    pub fn get_mut(&mut self) -> Option<&mut Metadata> {
        let metadata = self.metadata.as_mut()?;
        self.dirty = true;
        Some(metadata)
    }

    /// Moves a present value into the slot and returns the previous value.
    pub fn put(&mut self, metadata: Metadata) -> Option<Metadata> {
        self.dirty = true;
        self.metadata.replace(metadata)
    }

    /// Marks the value absent and moves out the previous value.
    pub fn clear(&mut self) -> Option<Metadata> {
        let previous = self.metadata.take();
        self.dirty |= previous.is_some();
        previous
    }

    /// Applies a staged write or clear only when this slot was edited.
    ///
    /// A failed flush leaves the slot dirty so the caller can repair the value
    /// and retry. A byte-identical write is a true no-op, which keeps a signed
    /// package's signature; a change to a signed package is refused by the
    /// shared properties edit.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::Invalid`] for a field core properties cannot
    /// store, [`crate::Error::Signed`] for a change to a signed package, and
    /// an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    #[cfg(test)]
    pub(crate) fn flush(&mut self, package: &mut OpcPackage) -> Result<bool> {
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
    /// mistake. A change to a signed package is refused, as in [`Self::flush`].
    ///
    /// # Errors
    ///
    /// The errors of [`Self::flush`].
    #[cfg(test)]
    pub(crate) fn stage(&mut self, package: &mut OpcPackage) -> Result<Stage<'_>> {
        if !self.dirty {
            let staged = self.source.clone();
            return Ok(Stage {
                slot: self,
                changed: false,
                staged,
            });
        }
        let (changed, staged) = match self.metadata.as_ref() {
            Some(metadata) => {
                ensure_core_only(metadata)?;
                let core = Core::merged(self.source.as_ref(), metadata)?;
                (write::sync(package, &core)?, Some(core))
            },
            None => (write::clear(package)?, None),
        };
        Ok(Stage {
            slot: self,
            changed,
            staged,
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
    fn xml_escapes_text_and_rejects_forbidden_characters() {
        let xml = Core::new().title("A & <B> \"C\"").xml().expect("valid XML");
        assert!(xml.contains("A &amp; &lt;B&gt; &quot;C&quot;"));

        let error = Core::new().title("bad\u{0}text").xml().unwrap_err();
        assert!(matches!(error, crate::Error::Invalid(_)));
    }

    #[test]
    fn merge_keeps_unchanged_lexical_forms_and_takes_changed_values() {
        let base = Core {
            keywords: Some(Keywords::new().text("a").value("b")),
            created: Some(time::W3c::new("2024-01-02T08:04:05+05:00").unwrap()),
            creator: Some("Ada".to_owned()),
            ..Core::default()
        };
        let mut target = base.metadata();
        assert_eq!(target.keywords, ["a", "b"]);
        assert_eq!(target.author.as_deref(), Some("Ada"));
        let unchanged = Core::merged(Some(&base), &target).unwrap();
        assert_eq!(unchanged, base);

        target.keywords = vec!["plain".to_owned()];
        target.author = None;
        let changed = Core::merged(Some(&base), &target).unwrap();
        assert_eq!(changed.keywords, Some(Keywords::from("plain")));
        assert_eq!(changed.creator, None);
        assert_eq!(changed.created, base.created);
    }

    #[test]
    fn complete_edit_refuses_fields_no_family_can_store() {
        let package = OpcPackage::new();
        let metadata = Metadata {
            title: Some("T".to_owned()),
            content_type: Some("application/example".to_owned()),
            ..Metadata::default()
        };
        let mut edit = Edit::new(&package, crate::custom::Host::Word).unwrap();
        edit.set_metadata(metadata);
        let error = edit.commit().unwrap_err();
        assert!(matches!(
            error,
            Error::Refused(refusal)
                if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedProperty
        ));
        assert!(
            read_with_report(&package, &mut ToleranceReport::default())
                .unwrap()
                .is_none()
        );
    }

    #[test]
    fn slot_tracks_only_real_edit_intent() {
        let package = OpcPackage::new();
        let mut slot = Slot::load(&package).expect("load");
        assert!(slot.get().is_none());
        assert!(!slot.is_dirty());
        assert!(slot.clear().is_none());
        assert!(!slot.is_dirty());

        slot.put(Metadata::default());
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
