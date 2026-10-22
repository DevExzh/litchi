//! Source-checked transactions for ordered spreadsheet named definitions.

use std::{fmt, sync::Arc};

use litchi_core::{Error, Position, Result};
use litchi_odf_common::XmlStreamLimits;

use crate::{
    codec::names as codec,
    model::names::{Definition, Scope},
    package::Archive,
};

/// An exact semantic key for one named definition.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Key<'a> {
    name: &'a str,
    scope: &'a Scope,
}

impl<'a> Key<'a> {
    /// Construct an exact name-and-scope key.
    #[must_use]
    pub const fn new(name: &'a str, scope: &'a Scope) -> Self {
        Self { name, scope }
    }

    /// Return the exact producer-visible name.
    #[must_use]
    pub const fn name(self) -> &'a str {
        self.name
    }

    /// Return the visibility scope.
    #[must_use]
    pub const fn scope(self) -> &'a Scope {
        self.scope
    }
}

/// A semantic definition selector.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Selector<'a> {
    /// Exact producer-visible name and visibility scope.
    Key(Key<'a>),
    /// Checked zero-based source position.
    Position(Position),
}

impl<'a> From<Key<'a>> for Selector<'a> {
    fn from(value: Key<'a>) -> Self {
        Self::Key(value)
    }
}

impl From<Position> for Selector<'_> {
    fn from(value: Position) -> Self {
        Self::Position(value)
    }
}

impl From<usize> for Selector<'_> {
    fn from(value: usize) -> Self {
        Self::Position(Position::new(value))
    }
}

/// An immutable named-definition owner bound to exact ODS package bytes.
#[derive(Clone)]
pub struct Snapshot {
    source: Arc<[u8]>,
    definitions: Arc<[Definition]>,
    xml_limits: XmlStreamLimits,
}

impl fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("source_bytes", &self.source.len())
            .field("definitions", &self.definitions.len())
            .finish()
    }
}

impl Snapshot {
    /// Parse an owned ODS package and capture its ordered definition catalog.
    ///
    /// # Errors
    ///
    /// Returns an error when the package or named-definition catalog is invalid.
    pub fn from_bytes(source: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_xml_limits(source, XmlStreamLimits::owner_default_value())
    }

    pub(crate) fn from_bytes_with_xml_limits(
        source: Vec<u8>,
        xml_limits: XmlStreamLimits,
    ) -> Result<Self> {
        Self::from_arc_with_xml_limits(Arc::from(source), xml_limits)
    }

    fn from_arc_with_xml_limits(source: Arc<[u8]>, xml_limits: XmlStreamLimits) -> Result<Self> {
        let package = Archive::from_bytes_with_xml_limits(source.as_ref().to_vec(), xml_limits)?;
        let definitions = Arc::from(package.definitions()?);
        Ok(Self {
            source,
            definitions,
            xml_limits,
        })
    }

    /// Borrow the exact package bytes captured by this snapshot.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.source
    }

    /// Borrow the ordered definition catalog.
    #[must_use]
    pub fn definitions(&self) -> &[Definition] {
        &self.definitions
    }

    /// Select one definition by exact key or checked source position.
    ///
    /// # Errors
    ///
    /// Returns an error if a malformed snapshot contains an ambiguous key.
    pub fn get<'a, S>(&self, selector: S) -> Result<Option<&Definition>>
    where
        S: Into<Selector<'a>>,
    {
        select(&self.definitions, selector.into())
            .map(|selected| selected.map(|position| &self.definitions[position]))
    }

    /// Start a clone-staged, failure-atomic edit.
    #[must_use]
    pub(crate) fn edit(&self) -> Edit {
        Edit {
            before: self.clone(),
            draft: self.definitions.to_vec(),
        }
    }
}

/// A clone-staged definition edit derived from one immutable [`Snapshot`].
#[derive(Clone, Debug)]
pub struct Edit {
    before: Snapshot,
    draft: Vec<Definition>,
}

impl Edit {
    /// Borrow the current ordered candidate.
    #[must_use]
    pub fn definitions(&self) -> &[Definition] {
        &self.draft
    }

    /// Append one validated definition.
    ///
    /// # Errors
    ///
    /// Returns an error when the candidate is invalid or duplicates a key.
    pub fn add(&mut self, definition: Definition) -> Result<()> {
        let mut candidate = self.draft.clone();
        candidate.push(definition);
        crate::model::names::validate_collection(&candidate)?;
        self.draft = candidate;
        Ok(())
    }

    /// Replace one selected definition without changing its source position.
    ///
    /// # Errors
    ///
    /// Returns an error when selection or candidate validation fails.
    pub fn replace<'a, S>(
        &mut self,
        selector: S,
        definition: Definition,
    ) -> Result<Option<Definition>>
    where
        S: Into<Selector<'a>>,
    {
        let Some(index) = select(&self.draft, selector.into())? else {
            return Ok(None);
        };
        let mut candidate = self.draft.clone();
        let previous = std::mem::replace(&mut candidate[index], definition);
        crate::model::names::validate_collection(&candidate)?;
        self.draft = candidate;
        Ok(Some(previous))
    }

    /// Remove one selected definition.
    ///
    /// # Errors
    ///
    /// Returns an error when selection is ambiguous.
    pub fn remove<'a, S>(&mut self, selector: S) -> Result<Option<Definition>>
    where
        S: Into<Selector<'a>>,
    {
        let Some(index) = select(&self.draft, selector.into())? else {
            return Ok(None);
        };
        let mut candidate = self.draft.clone();
        let removed = candidate.remove(index);
        crate::model::names::validate_collection(&candidate)?;
        self.draft = candidate;
        Ok(Some(removed))
    }

    /// Move one selected definition to a final checked zero-based position.
    ///
    /// # Errors
    ///
    /// Returns an error when selection is ambiguous or the destination is out of range.
    pub fn move_to<'a, S>(&mut self, selector: S, destination: Position) -> Result<Option<()>>
    where
        S: Into<Selector<'a>>,
    {
        let Some(source) = select(&self.draft, selector.into())? else {
            return Ok(None);
        };
        let destination_index = destination.get();
        if destination_index >= self.draft.len() {
            return Err(bounds(destination_index, self.draft.len()));
        }
        if source == destination_index {
            return Ok(Some(()));
        }
        let mut candidate = self.draft.clone();
        let definition = candidate.remove(source);
        candidate.insert(destination_index, definition);
        crate::model::names::validate_collection(&candidate)?;
        self.draft = candidate;
        Ok(Some(()))
    }

    /// Replace the complete ordered catalog.
    ///
    /// # Errors
    ///
    /// Returns an error when the candidate catalog is invalid.
    pub fn replace_all(&mut self, definitions: Vec<Definition>) -> Result<()> {
        crate::model::names::validate_collection(&definitions)?;
        self.draft = definitions;
        Ok(())
    }

    /// Remove every named definition.
    pub fn clear(&mut self) {
        self.draft.clear();
    }

    /// Discard every staged change in place.
    pub fn reset(&mut self) {
        self.draft = self.before.definitions.to_vec();
    }

    /// Validate, rewrite one compact XML owner, reparse, and publish atomically.
    ///
    /// # Errors
    ///
    /// Returns an error when validation, compact XML enforcement, package rebuilding, or typed
    /// readback fails.
    pub(crate) fn commit(self) -> Result<Commit> {
        crate::model::names::validate_collection(&self.draft)?;
        if self.draft.as_slice() == self.before.definitions() {
            return Ok(Commit::unchanged(self.before));
        }

        let package = Archive::from_bytes_with_xml_limits(
            self.before.source.as_ref().to_vec(),
            self.before.xml_limits,
        )?;
        let content_xml = codec::replace(package.content_xml(), &self.draft)?;
        litchi_odf_common::compact_xml::validate(
            crate::package::compact_prolog(&content_xml).as_bytes(),
            litchi_odf_common::compact_xml::Limits::default(),
        )
        .map_err(Error::from)?;
        let target_package = package.replace_content_xml(&content_xml)?;
        let target = Snapshot::from_bytes_with_xml_limits(
            target_package.into_bytes(),
            self.before.xml_limits,
        )?;
        if target.definitions() != self.draft {
            return Err(Error::InvalidFormat(
                "ODS named-definition typed readback does not match the staged edit".to_string(),
            ));
        }
        let patch = Patch {
            source: self.before.source.clone(),
            target: target.source.clone(),
        };
        Ok(Commit {
            snapshot: target,
            patch,
        })
    }
}

/// An exact-source, reversible named-definition patch.
#[derive(Clone)]
pub(crate) struct Patch {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("source_bytes", &self.source.len())
            .field("target_bytes", &self.target.len())
            .finish()
    }
}

impl Patch {
    /// Return whether this patch changes package bytes.
    #[must_use]
    pub(crate) fn changed(&self) -> bool {
        self.source != self.target
    }
}

/// A validated named-definition publication.
#[derive(Debug)]
pub(crate) struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    fn unchanged(snapshot: Snapshot) -> Self {
        let source = snapshot.source.clone();
        Self {
            snapshot,
            patch: Patch {
                source: source.clone(),
                target: source,
            },
        }
    }

    /// Return whether package bytes changed.
    #[must_use]
    pub(crate) fn changed(&self) -> bool {
        self.patch.changed()
    }

    /// Borrow the resulting immutable snapshot.
    #[must_use]
    pub(crate) const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }
}

fn select(definitions: &[Definition], selector: Selector<'_>) -> Result<Option<usize>> {
    match selector {
        Selector::Position(position) => {
            Ok((position.get() < definitions.len()).then_some(position.get()))
        },
        Selector::Key(key) => {
            let mut matches = definitions.iter().enumerate().filter(|(_, definition)| {
                definition.name() == key.name && definition.scope() == key.scope
            });
            let selected = matches.next().map(|(index, _)| index);
            if selected.is_some() && matches.next().is_some() {
                return Err(Error::InvalidFormat(
                    "ODS named-definition selector is ambiguous".to_string(),
                ));
            }
            Ok(selected)
        },
    }
}

fn bounds(position: usize, length: usize) -> Error {
    Error::InvalidFormat(format!(
        "ODS named-definition position {position} is outside catalog length {length}"
    ))
}
