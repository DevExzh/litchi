//! Atomic ODS metadata edits backed by the common retained-source patcher.

use std::borrow::Cow;

use litchi_core::{Error, Metadata as CoreMetadata, Result};
use litchi_odf_common::{
    XmlStreamLimits,
    core::{MetaXmlPatch, Structure, metadata::Metadata},
};

use super::model::Snapshot;

/// An isolated metadata draft derived from an immutable [`Snapshot`].
pub(crate) struct Transaction<'source> {
    source: Option<&'source str>,
    source_odf: Metadata,
    original: CoreMetadata,
    draft: CoreMetadata,
    remove_part: bool,
    xml_limits: XmlStreamLimits,
}

impl<'source> Transaction<'source> {
    pub(crate) fn from_snapshot(snapshot: &'source Snapshot) -> Self {
        Self {
            source: snapshot.source.as_deref(),
            source_odf: snapshot.odf.clone(),
            original: snapshot.value.clone(),
            draft: snapshot.value.clone(),
            remove_part: false,
            xml_limits: snapshot.xml_limits,
        }
    }

    /// Borrow the current ergonomic metadata draft.
    #[must_use]
    pub(crate) fn metadata(&self) -> &CoreMetadata {
        &self.draft
    }

    /// Replace the supported ergonomic metadata projection.
    ///
    /// ODF fields that the common retained-source patcher does not own must
    /// remain unchanged; silently dropping such edits would violate the
    /// snapshot contract, so they are rejected before any package bytes move.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub(crate) fn replace(&mut self, metadata: CoreMetadata) -> Result<()> {
        validate_supported_change(&self.original, &metadata)?;
        self.draft = metadata;
        self.remove_part = false;
        Ok(())
    }

    #[cfg(test)]
    /// Open a short-lived semantic editor for common CRUD operations.
    pub(crate) fn editor(&mut self) -> Editor<'_, 'source> {
        Editor { transaction: self }
    }

    /// Remove the complete physical `meta.xml` part.
    pub(crate) fn remove(&mut self) {
        self.remove_part = true;
        self.draft = CoreMetadata::default();
    }

    /// XmlCommit the draft into a bounded metadata XML result.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub(crate) fn commit(self) -> Result<XmlCommit<'source>> {
        if self.remove_part {
            return Ok(XmlCommit {
                xml: None,
                changed: self.source.is_some(),
            });
        }
        if supported_equal(&self.original, &self.draft) {
            return Ok(XmlCommit {
                xml: self.source.map(Cow::Borrowed),
                changed: false,
            });
        }

        let source = self
            .source
            .map_or_else(Structure::default_meta_xml, str::to_owned);
        let source_metadata = if self.source.is_some() {
            self.source_odf
        } else {
            Metadata::from_xml(&source, self.xml_limits)?
        };
        let patch = MetaXmlPatch::preserve_all().diff_simple_fields(&source_metadata, &self.draft);
        let xml = litchi_odf_common::core::patch_meta_xml(&source, &patch, self.xml_limits)?
            .unwrap_or_else(Structure::default_meta_xml);
        if xml.len() > super::model::MAX_XML_BYTES {
            return Err(Error::InvalidFormat(
                "patched ODS meta.xml exceeds the size limit".to_string(),
            ));
        }
        Metadata::from_xml(&xml, self.xml_limits)?;
        Ok(XmlCommit {
            xml: Some(Cow::Owned(xml)),
            changed: true,
        })
    }
}

/// A short-lived editor that keeps metadata operations contextual.
pub struct Editor<'transaction, 'source> {
    transaction: &'transaction mut Transaction<'source>,
}

impl Editor<'_, '_> {
    /// Borrow the current draft.
    #[must_use]
    pub fn metadata(&self) -> &CoreMetadata {
        self.transaction.metadata()
    }

    /// Apply a checked update to the ergonomic metadata projection.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn update<F>(&mut self, update: F) -> Result<()>
    where
        F: FnOnce(&mut CoreMetadata) -> Result<()>,
    {
        let mut candidate = self.transaction.draft.clone();
        update(&mut candidate)?;
        self.transaction.replace(candidate)
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_title(&mut self, value: impl Into<String>) -> Result<()> {
        self.update(|metadata| {
            metadata.title = Some(value.into());
            Ok(())
        })
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn clear_title(&mut self) -> Result<()> {
        self.update(|metadata| {
            metadata.title = None;
            Ok(())
        })
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_author(&mut self, value: impl Into<String>) -> Result<()> {
        self.update(|metadata| {
            metadata.author = Some(value.into());
            Ok(())
        })
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn clear_author(&mut self) -> Result<()> {
        self.update(|metadata| {
            metadata.author = None;
            Ok(())
        })
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_subject(&mut self, value: impl Into<String>) -> Result<()> {
        self.update(|metadata| {
            metadata.subject = Some(value.into());
            Ok(())
        })
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn clear_subject(&mut self) -> Result<()> {
        self.update(|metadata| {
            metadata.subject = None;
            Ok(())
        })
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_description(&mut self, value: impl Into<String>) -> Result<()> {
        self.update(|metadata| {
            metadata.description = Some(value.into());
            Ok(())
        })
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn clear_description(&mut self) -> Result<()> {
        self.update(|metadata| {
            metadata.description = None;
            Ok(())
        })
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_keywords(&mut self, keywords: Vec<String>) -> Result<()> {
        self.update(|metadata| {
            metadata.keywords = keywords;
            Ok(())
        })
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn clear_keywords(&mut self) -> Result<()> {
        self.update(|metadata| {
            metadata.keywords.clear();
            Ok(())
        })
    }

    /// Remove the physical metadata part as an explicit CRUD operation.
    pub fn remove(&mut self) {
        self.transaction.remove();
    }
}

/// The result of publishing a metadata transaction.
pub(crate) struct XmlCommit<'source> {
    xml: Option<Cow<'source, str>>,
    changed: bool,
}

impl XmlCommit<'_> {
    #[cfg(test)]
    /// Borrow the new metadata XML, or `None` when the part is absent.
    #[must_use]
    pub(crate) fn xml(&self) -> Option<&str> {
        self.xml.as_deref()
    }

    /// Whether the transaction changed the physical part.
    #[must_use]
    pub(crate) fn changed(&self) -> bool {
        self.changed
    }
}

impl<'source> XmlCommit<'source> {
    pub(crate) fn into_owned_xml(self) -> Option<String> {
        self.xml.map(Cow::into_owned)
    }
}

fn validate_supported_change(original: &CoreMetadata, candidate: &CoreMetadata) -> Result<()> {
    if unsupported_equal(original, candidate) {
        Ok(())
    } else {
        Err(Error::InvalidFormat(
            "ODS metadata edit contains fields not owned by the common retained-source patcher"
                .to_string(),
        ))
    }
}

fn supported_equal(left: &CoreMetadata, right: &CoreMetadata) -> bool {
    left.title == right.title
        && left.subject == right.subject
        && left.author == right.author
        && left.keywords == right.keywords
        && left.description == right.description
}

fn unsupported_equal(left: &CoreMetadata, right: &CoreMetadata) -> bool {
    left.identifier == right.identifier
        && left.language == right.language
        && left.template == right.template
        && left.last_modified_by == right.last_modified_by
        && left.revision == right.revision
        && left.created == right.created
        && left.modified == right.modified
        && left.page_count == right.page_count
        && left.word_count == right.word_count
        && left.character_count == right.character_count
        && left.character_count_with_spaces == right.character_count_with_spaces
        && left.editing_time_minutes == right.editing_time_minutes
        && left.application == right.application
        && left.category == right.category
        && left.company == right.company
        && left.manager == right.manager
        && left.content_status == right.content_status
        && left.content_type == right.content_type
        && left.version == right.version
        && left.last_printed == right.last_printed
        && left.last_backup == right.last_backup
        && left.hyperlink_base == right.hyperlink_base
        && left.security == right.security
        && left.codepage == right.codepage
}
