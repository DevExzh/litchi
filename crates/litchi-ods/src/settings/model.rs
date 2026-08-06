//! Immutable calculation-settings snapshots bound to one ODS content part.

use litchi_core::{Error, Result};
use litchi_odf_common::calculation::{Settings, parse};

use super::codec::{Location, locate};
use super::transaction::Transaction;

/// An immutable, context-aware view of one ODS `content.xml` part.
///
/// The source XML is borrowed so inspecting a package does not copy its
/// content part.  The typed calculation settings are decoded once, while the
/// original XML remains available for a zero-allocation no-op transaction.
#[derive(Debug)]
pub struct Snapshot<'xml> {
    pub(crate) source: &'xml str,
    pub(crate) calculation: Option<Settings>,
    pub(crate) location: Location,
}

impl<'xml> Snapshot<'xml> {
    /// Decode calculation settings from an ODS `content.xml` document.
    ///
    /// The surrounding document must contain exactly one direct
    /// `office:body/office:spreadsheet` host.  Unknown XML outside the owned
    /// calculation-settings element is not interpreted and remains part of
    /// the borrowed source snapshot.
    pub fn from_content_xml(source: &'xml str) -> Result<Self> {
        let calculation = parse(source)?;
        let location = locate(source)?;
        if calculation.is_some() != location.calculation.is_some() {
            return Err(Error::InvalidFormat(
                "calculation-settings semantic and XML locations disagree".to_string(),
            ));
        }
        Ok(Self {
            source,
            calculation,
            location,
        })
    }

    /// Borrow the original content XML without normalization.
    pub fn content_xml(&self) -> &'xml str {
        self.source
    }

    /// Return the typed calculation settings, if the document declares them.
    pub fn calculation(&self) -> Option<&Settings> {
        self.calculation.as_ref()
    }

    /// Start an isolated transaction against this immutable snapshot.
    pub fn transaction(&self) -> Transaction<'xml> {
        Transaction::from_snapshot(self)
    }
}
