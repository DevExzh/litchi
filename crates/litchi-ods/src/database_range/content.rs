//! Source-bound content XML edits used by [`crate::Builder`].

use crate::model::database_range::Range;
use litchi_core::{Error, Result};

use super::{
    codec,
    model::{Selector, select},
    validation,
};

/// Clone-staged database-range edit over a builder content XML source.
///
/// This is the package-free counterpart of [`super::Edit`]. It keeps the
/// source XML until commit so a failed closure or validation leaves the
/// builder unchanged.
pub struct Draft {
    source: String,
    location: codec::Location,
    original: Option<Vec<Range>>,
    draft: Option<Vec<Range>>,
}

impl Draft {
    pub(crate) fn from_source(source: &str) -> Result<Self> {
        let location = codec::locate(
            source,
            &litchi_odf_common::XmlStreamLimits::owner_default_value(),
        )?;
        let original = if let Some(container) = &location.container {
            let fragment = codec::owner_fragment(source, container)?;
            Some(crate::model::database_range::parse_database_ranges(
                &fragment,
                &litchi_odf_common::XmlStreamLimits::owner_default_value(),
            )?)
        } else {
            None
        };
        validation::validate_snapshot(source, &location, original.as_deref().unwrap_or(&[]))?;
        Ok(Self {
            source: source.to_owned(),
            location: location.clone(),
            draft: original.clone(),
            original,
        })
    }

    /// Borrow the currently staged declarations.
    #[must_use]
    pub fn ranges(&self) -> &[Range] {
        self.draft.as_deref().unwrap_or(&[])
    }

    /// Return whether the candidate has a physical owner.
    #[must_use]
    pub const fn has_owner(&self) -> bool {
        self.draft.is_some()
    }

    /// Replace the complete ordered catalog.
    pub fn replace(&mut self, ranges: Vec<Range>) -> Result<()> {
        let candidate = Some(ranges);
        validation::validate_candidate(&self.location, &self.original, &candidate)?;
        self.draft = candidate;
        Ok(())
    }

    /// Remove the physical owner.
    pub fn remove(&mut self) -> Result<()> {
        validation::validate_candidate(&self.location, &self.original, &None)?;
        self.draft = None;
        Ok(())
    }

    /// Open semantic CRUD verbs over this edit.
    pub fn editor(&mut self) -> DraftEditor<'_> {
        DraftEditor { edit: self }
    }

    /// Publish the candidate content XML.
    pub(crate) fn commit(self) -> Result<DraftCommit> {
        validation::validate_candidate(&self.location, &self.original, &self.draft)?;
        if self.original == self.draft {
            return Ok(DraftCommit {
                source_xml: self.source,
                changed: false,
            });
        }
        let source_xml = codec::replace(
            &self.source,
            &self.location,
            self.draft.as_deref(),
            &litchi_odf_common::XmlStreamLimits::owner_default_value(),
        )?;
        Ok(DraftCommit {
            source_xml,
            changed: true,
        })
    }
}

/// Semantic CRUD verbs over a builder content edit.
pub struct DraftEditor<'edit> {
    edit: &'edit mut Draft,
}

impl DraftEditor<'_> {
    /// Borrow the current staged declarations.
    #[must_use]
    pub fn ranges(&self) -> &[Range] {
        self.edit.ranges()
    }

    /// Add one declaration at the catalog tail.
    pub fn add(&mut self, range: Range) -> Result<()> {
        let mut candidate = self.edit.draft.clone().unwrap_or_default();
        candidate.push(range);
        self.edit.replace(candidate)
    }

    /// Replace one declaration by name or source position.
    pub fn replace<'a, S>(&mut self, selector: S, range: Range) -> Result<()>
    where
        S: Into<Selector<'a>>,
    {
        let mut candidate = self.edit.draft.clone().unwrap_or_default();
        let index = select(&candidate, selector.into())?.ok_or_else(|| {
            Error::InvalidFormat("ODS database-range selector did not match".to_string())
        })?;
        candidate[index] = range;
        self.edit.replace(candidate)
    }

    /// Apply one checked update to a declaration.
    pub fn update<'a, S, F>(&mut self, selector: S, update: F) -> Result<()>
    where
        S: Into<Selector<'a>>,
        F: FnOnce(&mut Range) -> Result<()>,
    {
        let mut candidate = self.edit.draft.clone().unwrap_or_default();
        let index = select(&candidate, selector.into())?.ok_or_else(|| {
            Error::InvalidFormat("ODS database-range selector did not match".to_string())
        })?;
        update(&mut candidate[index])?;
        self.edit.replace(candidate)
    }

    /// Remove one declaration and return it.
    pub fn remove<'a, S>(&mut self, selector: S) -> Result<Range>
    where
        S: Into<Selector<'a>>,
    {
        let mut candidate = self.edit.draft.clone().unwrap_or_default();
        let index = select(&candidate, selector.into())?.ok_or_else(|| {
            Error::InvalidFormat("ODS database-range selector did not match".to_string())
        })?;
        let removed = candidate.remove(index);
        if candidate.is_empty() {
            self.edit.remove()?;
        } else {
            self.edit.replace(candidate)?;
        }
        Ok(removed)
    }

    /// Remove all declarations and the physical owner.
    pub fn clear(&mut self) -> Result<()> {
        self.edit.remove()
    }
}

/// Result of publishing a builder content edit.
pub(crate) struct DraftCommit {
    source_xml: String,
    changed: bool,
}

impl DraftCommit {
    /// Return whether the content XML changed.
    #[must_use]
    pub(crate) const fn changed(&self) -> bool {
        self.changed
    }

    /// Consume the commit into content XML.
    #[must_use]
    pub(crate) fn into_source_xml(self) -> String {
        self.source_xml
    }
}
