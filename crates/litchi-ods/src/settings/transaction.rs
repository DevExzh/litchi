//! Atomic calculation-settings transactions and contextual editors.

use std::borrow::Cow;

use litchi_core::Result;
use litchi_odf_common::calculation::Settings;

use super::codec::replace;
use super::model::Snapshot;

/// An isolated mutable draft derived from an immutable [`Snapshot`].
pub struct Transaction<'xml> {
    pub(crate) source: &'xml str,
    pub(crate) original: Option<Settings>,
    pub(crate) draft: Option<Settings>,
    pub(crate) location: super::codec::Location,
}

impl<'xml> Transaction<'xml> {
    pub(crate) fn from_snapshot(snapshot: &Snapshot<'xml>) -> Self {
        Self {
            source: snapshot.source,
            original: snapshot.calculation.clone(),
            draft: snapshot.calculation.clone(),
            location: snapshot.location.clone(),
        }
    }

    /// Read the current draft without exposing mutable fields outside the
    /// editor boundary.
    #[must_use]
    pub fn calculation(&self) -> Option<&Settings> {
        self.draft.as_ref()
    }

    /// Open a contextual editor for replacement, update, or removal.
    pub fn editor(&mut self) -> Editor<'_, 'xml> {
        Editor { transaction: self }
    }

    /// Replace the complete typed calculation-settings value.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn replace(&mut self, settings: Settings) -> Result<()> {
        settings.validate()?;
        self.draft = Some(settings);
        Ok(())
    }

    /// Remove the calculation-settings element from the package content.
    pub fn remove(&mut self) {
        self.draft = None;
    }

    /// Atomically publish the draft.  A semantically unchanged transaction
    /// borrows the original XML, retaining unknown or vendor XML exactly.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn commit(self) -> Result<Commit<'xml>> {
        if self.draft == self.original {
            return Ok(Commit {
                xml: Cow::Borrowed(self.source),
                calculation: self.draft,
                changed: false,
            });
        }
        if let Some(settings) = &self.draft {
            settings.validate()?;
        }
        let xml = replace(self.source, &self.location, self.draft.as_ref())?;
        Ok(Commit {
            xml: Cow::Owned(xml),
            calculation: self.draft,
            changed: true,
        })
    }
}

/// A short-lived contextual editor borrowing one transaction.
pub struct Editor<'transaction, 'xml> {
    transaction: &'transaction mut Transaction<'xml>,
}

impl Editor<'_, '_> {
    /// Return the current typed draft, if present.
    #[must_use]
    pub fn calculation(&self) -> Option<&Settings> {
        self.transaction.calculation()
    }

    /// Replace the complete typed value after common validation.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn replace(&mut self, settings: Settings) -> Result<()> {
        self.transaction.replace(settings)
    }

    /// Apply a checked in-place update to the current value.  If no value is
    /// present, the editor starts from the common default settings value.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn update<F>(&mut self, update: F) -> Result<()>
    where
        F: FnOnce(&mut Settings) -> Result<()>,
    {
        let mut candidate = self.transaction.draft.clone().unwrap_or_default();
        update(&mut candidate)?;
        self.transaction.replace(candidate)
    }

    /// Remove the owned calculation-settings element.
    pub fn remove(&mut self) {
        self.transaction.remove();
    }
}

/// The result of publishing a calculation-settings transaction.
pub struct Commit<'xml> {
    xml: Cow<'xml, str>,
    calculation: Option<Settings>,
    changed: bool,
}

impl Commit<'_> {
    /// Borrow the resulting content XML.  This is the original source for a
    /// no-op commit and the rebuilt XML for a changed commit.
    #[must_use]
    pub fn content_xml(&self) -> &str {
        &self.xml
    }

    /// Return the typed value that is now represented by the result.
    #[must_use]
    pub fn calculation(&self) -> Option<&Settings> {
        self.calculation.as_ref()
    }

    /// Whether the transaction required an XML rebuild.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.changed
    }
}

impl<'xml> Commit<'xml> {
    /// Consume the result while retaining a borrow for no-op commits.
    #[must_use]
    pub fn into_xml(self) -> Cow<'xml, str> {
        self.xml
    }

    /// Consume the result into package-owned UTF-8 XML.
    #[must_use]
    pub fn into_owned(self) -> String {
        self.xml.into_owned()
    }
}
