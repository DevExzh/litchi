//! Failure-atomic semantic edits over one inert `WebPub` snapshot.

use std::sync::Arc;

use crate::{Error, Result};

use super::super::codec::rewrite_preserving_source;
use super::super::model::{WebPageType, WebPub, WebPubRange};
use super::model::{Commit, Patch, Snapshot};

/// A detached transaction over bounded BIFF8 web-publication metadata.
///
/// The source record's conditional topology is fixed: edits can update owned
/// values but cannot add/remove `srcName`, `ref8`, or `crtID`. Every rejected
/// operation leaves the candidate untouched.
#[derive(Clone, Debug)]
pub struct Transaction {
    source: Snapshot,
    candidate: Vec<u8>,
    publication: WebPub,
}

impl Transaction {
    pub(super) fn new(source: Snapshot) -> Self {
        Self {
            candidate: source.bytes().to_vec(),
            publication: source.publication().clone(),
            source,
        }
    }

    /// Returns the immutable source snapshot used for publication checks.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.source
    }

    /// Alias for [`Self::before`].
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        self.before()
    }

    /// Returns the current typed candidate view.
    #[must_use]
    pub const fn publication(&self) -> &WebPub {
        &self.publication
    }

    /// Alias for callers using the BIFF record name.
    #[must_use]
    pub const fn web_pub(&self) -> &WebPub {
        self.publication()
    }

    /// Returns the current typed candidate view using the semantic name.
    #[must_use]
    pub const fn web_publication(&self) -> &WebPub {
        self.publication()
    }

    /// Materializes the current candidate as a validated immutable snapshot.
    pub fn snapshot(&self) -> Result<Snapshot> {
        if self.candidate.as_slice() == self.source.bytes() {
            Ok(self.source.clone())
        } else {
            Snapshot::parse(&self.candidate)
        }
    }

    /// Whether a staged edit changes any payload byte.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.candidate.as_slice() != self.source.bytes()
    }

    /// Replaces the page behavior while retaining all other wire fields.
    pub fn set_page_type(&mut self, value: WebPageType) -> Result<&mut Self> {
        let mut replacement = self.publication.clone();
        replacement.page_type = value;
        self.replace(replacement)
    }

    /// Toggles automatic republishing on save.
    pub fn set_auto_republish(&mut self, value: bool) -> Result<&mut Self> {
        let mut replacement = self.publication.clone();
        replacement.auto_republish = value;
        self.replace(replacement)
    }

    /// Toggles the single-file MHTML publication flag.
    pub fn set_single_file(&mut self, value: bool) -> Result<&mut Self> {
        let mut replacement = self.publication.clone();
        replacement.single_file = value;
        self.replace(replacement)
    }

    /// Replaces the style identifier carried by `nStyleId`.
    pub fn set_style_id(&mut self, value: u32) -> Result<&mut Self> {
        let mut replacement = self.publication.clone();
        replacement.style_id = value;
        self.replace(replacement)
    }

    /// Replaces the inert destination URL or file path.
    ///
    /// The value is metadata only: this method never resolves, opens, or
    /// fetches the destination.
    pub fn set_file_destination(&mut self, value: impl Into<String>) -> Result<&mut Self> {
        let mut replacement = self.publication.clone();
        replacement.file_destination = value.into();
        self.replace(replacement)
    }

    /// Alias for [`Self::set_file_destination`].
    pub fn set_destination(&mut self, value: impl Into<String>) -> Result<&mut Self> {
        self.set_file_destination(value)
    }

    /// Replaces the inert destination bookmark/division identifier.
    pub fn set_div_id(&mut self, value: impl Into<String>) -> Result<&mut Self> {
        let mut replacement = self.publication.clone();
        replacement.div_id = value.into();
        self.replace(replacement)
    }

    /// Replaces the publication title.
    pub fn set_title(&mut self, value: impl Into<String>) -> Result<&mut Self> {
        let mut replacement = self.publication.clone();
        replacement.title = value.into();
        self.replace(replacement)
    }

    /// Replaces the source name when this record already carries `srcName`.
    pub fn set_source_name(&mut self, value: impl Into<String>) -> Result<&mut Self> {
        if self.publication.source_name.is_none() {
            return Err(Error::UnsafeEdit(
                "this WebPub source does not carry editable srcName metadata".into(),
            ));
        }
        let mut replacement = self.publication.clone();
        replacement.source_name = Some(value.into());
        self.replace(replacement)
    }

    /// Replaces the range when this record already carries `ref8`.
    pub fn set_range(&mut self, value: WebPubRange) -> Result<&mut Self> {
        if self.publication.range.is_none() {
            return Err(Error::UnsafeEdit(
                "this WebPub source does not carry editable ref8 metadata".into(),
            ));
        }
        let mut replacement = self.publication.clone();
        replacement.range = Some(value);
        self.replace(replacement)
    }

    /// Replaces the chart shape identifier when this record already carries
    /// `crtID`.
    pub fn set_chart_shape_id(&mut self, value: u32) -> Result<&mut Self> {
        if self.publication.chart_shape_id.is_none() {
            return Err(Error::UnsafeEdit(
                "this WebPub source does not carry editable crtID metadata".into(),
            ));
        }
        let mut replacement = self.publication.clone();
        replacement.chart_shape_id = Some(value);
        self.replace(replacement)
    }

    /// Discards all staged edits and returns the original source snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    /// Validates and publishes the candidate with a reversible source-checked
    /// patch. Failed validation leaves the transaction candidate untouched.
    pub fn commit(self) -> Result<Commit> {
        let source = self.source;
        if self.candidate.as_slice() == source.bytes() {
            let patch = Patch::new(source.clone(), source.clone());
            return Ok(Commit::new(source, patch));
        }
        let snapshot = Snapshot::parse_shared(Arc::from(self.candidate.into_boxed_slice()))?;
        let patch = Patch::new(source, snapshot.clone());
        Ok(Commit::new(snapshot, patch))
    }

    fn replace(&mut self, replacement: WebPub) -> Result<&mut Self> {
        if replacement == self.publication {
            return Ok(self);
        }
        if &replacement == self.source.publication() {
            self.candidate = self.source.bytes().to_vec();
            self.publication = self.source.publication().clone();
            return Ok(self);
        }
        let candidate =
            rewrite_preserving_source(&self.candidate, &self.publication, &replacement)?;
        let publication = WebPub::parse(&candidate)?;
        self.candidate = candidate;
        self.publication = publication;
        Ok(self)
    }
}
