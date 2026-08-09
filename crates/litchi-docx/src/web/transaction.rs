#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Source-checked web-settings snapshots, semantic edits, and patches.

use std::sync::Arc;

use super::codec;
use super::model::{Border, BorderSide, Conformance, Frameset, Id, Key, Layout, Screen, Settings};
use super::{Error, Result, invalid};

/// One bounded semantic operation supported by the source-preserving codec.
#[derive(Debug, Clone)]
pub(super) enum Edit {
    TargetScreen(Option<Screen>),
    FramesetSize(Option<String>),
    FramesetLayout(Option<Layout>),
    DivBorder {
        id: Id,
        side: BorderSide,
        value: Option<Border>,
    },
}

/// An immutable, cheaply clonable web-settings source snapshot.
#[derive(Debug, Clone)]
pub struct Snapshot {
    xml: Arc<[u8]>,
    settings: Settings,
    conformance: Conformance,
}

impl Snapshot {
    /// Parse and retain a bounded web-settings XML source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn from_xml(xml: impl Into<Vec<u8>>) -> Result<Self> {
        let xml = xml.into();
        let (settings, conformance) = codec::parse(&xml)?;
        Ok(Self {
            xml: Arc::from(xml.into_boxed_slice()),
            settings,
            conformance,
        })
    }

    /// Parse a snapshot whose conformance is already established by its OPC
    /// relationship. Package loading uses this to bind the XML and graph
    /// dialects before exposing the snapshot.
    pub(super) fn from_xml_with_conformance(
        xml: impl Into<Vec<u8>>,
        conformance: Conformance,
    ) -> Result<Self> {
        let snapshot = Self::from_xml(xml)?;
        if snapshot.conformance != conformance {
            return Err(invalid(
                "web-settings XML conformance does not match its relationship",
            ));
        }
        Ok(snapshot)
    }

    /// Borrow the exact authored XML retained by this snapshot.
    #[inline]
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        &self.xml
    }

    /// Borrow the contextual typed web-settings projection.
    #[inline]
    #[must_use]
    pub const fn settings(&self) -> &Settings {
        &self.settings
    }

    /// Return the detected namespace/conformance family.
    #[inline]
    #[must_use]
    pub const fn conformance(&self) -> Conformance {
        self.conformance
    }

    /// Start an isolated source-checked edit.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            base: self.clone(),
            next: self.settings.clone(),
            edits: Vec::new(),
        }
    }
}

/// A web-settings edit that has not yet been published.
#[derive(Debug, Clone)]
pub struct Transaction {
    base: Snapshot,
    next: Settings,
    edits: Vec<Edit>,
}

impl Transaction {
    /// Borrow the projected typed web-settings state.
    #[inline]
    #[must_use]
    pub const fn settings(&self) -> &Settings {
        &self.next
    }

    /// Set or remove the inert target screen metadata.
    pub fn set_target_screen_size(&mut self, value: Option<Screen>) -> &mut Self {
        if self.next.target_screen_size() == value {
            return self;
        }
        match value {
            Some(value) => {
                self.next.set_target_screen_size(value);
            },
            None => {
                self.next.clear_target_screen_size();
            },
        }
        self.record_target_screen(value);
        self
    }

    /// Remove the inert target screen metadata.
    pub fn clear_target_screen_size(&mut self) -> &mut Self {
        self.set_target_screen_size(None)
    }

    /// Set the root frameset size expression. The value is retained as
    /// metadata; no URL or target is fetched or rendered.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_frameset_size(&mut self, value: impl Into<String>) -> Result<&mut Self> {
        let value = value.into();
        if self
            .next
            .frameset()
            .and_then(Frameset::size)
            .is_some_and(|current| current == value)
        {
            return Ok(self);
        }
        if let Some(frameset) = self.next.frameset_mut() {
            frameset.set_size(value.clone())?;
        } else {
            let mut frameset = Frameset::default();
            frameset.set_size(value.clone())?;
            self.next.set_frameset(frameset);
        }
        self.record_frameset_size(Some(value));
        Ok(self)
    }

    /// Remove the root frameset size expression.
    pub fn clear_frameset_size(&mut self) -> &mut Self {
        let Some(frameset) = self.next.frameset_mut() else {
            return self;
        };
        if frameset.size().is_none() {
            return self;
        }
        frameset.clear_size();
        self.drop_transaction_created_frameset();
        self.record_frameset_size(None);
        self
    }

    /// Set or remove the root frameset layout.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_frameset_layout(&mut self, value: Option<Layout>) -> Result<&mut Self> {
        let current = self.next.frameset().and_then(Frameset::layout);
        if current == value {
            return Ok(self);
        }
        if let Some(value) = value {
            if let Some(frameset) = self.next.frameset_mut() {
                frameset.set_layout(value);
            } else {
                let mut frameset = Frameset::default();
                frameset.set_layout(value);
                self.next.set_frameset(frameset);
            }
        } else {
            if let Some(frameset) = self.next.frameset_mut() {
                frameset.clear_layout();
            }
            self.drop_transaction_created_frameset();
        }
        self.record_frameset_layout(value);
        Ok(self)
    }

    /// Remove the explicit root frameset layout.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn clear_frameset_layout(&mut self) -> Result<&mut Self> {
        self.set_frameset_layout(None)
    }

    /// Set or remove one typed border on a uniquely identified HTML division.
    ///
    /// A numeric [`Key::Index`] is resolved once to the division's nonzero
    /// producer ID; the resulting patch is therefore anchored semantically.
    /// The value is inert formatting metadata and never causes target access.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_div_border(
        &mut self,
        key: impl Into<Key>,
        side: BorderSide,
        value: Option<Border>,
    ) -> Result<&mut Self> {
        let key = key.into();
        let id = self
            .next
            .get(key)?
            .ok_or_else(|| invalid("HTML division selector does not exist"))?
            .id();
        let current = self
            .next
            .get(id)?
            .and_then(|div| div.borders())
            .and_then(|borders| border_at(borders, side));
        if current == value.as_ref() {
            return Ok(self);
        }

        let div = self
            .next
            .get(id)?
            .ok_or_else(|| invalid("HTML division selector changed during border edit"))?;
        let mut div = div.clone();
        match value.as_ref() {
            Some(value) => {
                let mut borders = div.borders().cloned().unwrap_or_default();
                set_border(&mut borders, side, value.clone());
                div.set_borders(borders);
            },
            None => {
                if let Some(mut borders) = div.borders().cloned() {
                    clear_border(&mut borders, side);
                    if all_borders_absent(&borders) {
                        div.clear_borders();
                    } else {
                        div.set_borders(borders);
                    }
                }
            },
        }
        self.next.put(div)?;
        self.record_div_border(id, side, value);
        Ok(self)
    }

    /// Validate and publish the source-preserving edit.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn commit(self) -> Result<Commit> {
        // The canonical writer is used only as a complete semantic validation
        // pass. The authored bytes below are produced by the surgical codec.
        let _ = self.next.xml(self.base.conformance)?;
        if self.edits.is_empty() {
            return Ok(Commit {
                patch: Patch {
                    before: self.base.settings.clone(),
                    after: self.base.settings.clone(),
                    before_xml: self.base.xml.clone(),
                    after_xml: self.base.xml.clone(),
                    conformance: self.base.conformance,
                },
                snapshot: self.base,
            });
        }

        let xml = codec::rewrite(self.base.xml_bytes(), &self.edits)?;
        let snapshot = Snapshot::from_xml_with_conformance(xml, self.base.conformance)?;
        if snapshot.settings != self.next {
            return Err(invalid(
                "source-preserving web-settings rewrite did not match its semantic projection",
            ));
        }
        Ok(Commit {
            patch: Patch {
                before: self.base.settings,
                after: self.next,
                before_xml: self.base.xml,
                after_xml: snapshot.xml.clone(),
                conformance: snapshot.conformance,
            },
            snapshot,
        })
    }

    fn drop_transaction_created_frameset(&mut self) {
        if self.base.settings.frameset().is_none()
            && self
                .next
                .frameset()
                .is_some_and(|frameset| *frameset == Frameset::default())
        {
            self.next.clear_frameset();
        }
    }

    fn record_target_screen(&mut self, value: Option<Screen>) {
        let base = self.base.settings.target_screen_size();
        if let Some(index) = self
            .edits
            .iter()
            .position(|edit| matches!(edit, Edit::TargetScreen(_)))
        {
            if base == value {
                self.edits.remove(index);
            } else {
                self.edits[index] = Edit::TargetScreen(value);
            }
        } else if base != value {
            self.edits.push(Edit::TargetScreen(value));
        }
    }

    fn record_frameset_size(&mut self, value: Option<String>) {
        let base = self
            .base
            .settings
            .frameset()
            .and_then(Frameset::size)
            .map(str::to_owned);
        if let Some(index) = self
            .edits
            .iter()
            .position(|edit| matches!(edit, Edit::FramesetSize(_)))
        {
            if base == value {
                self.edits.remove(index);
            } else {
                self.edits[index] = Edit::FramesetSize(value);
            }
        } else if base != value {
            self.edits.push(Edit::FramesetSize(value));
        }
    }

    fn record_frameset_layout(&mut self, value: Option<Layout>) {
        let base = self.base.settings.frameset().and_then(Frameset::layout);
        if let Some(index) = self
            .edits
            .iter()
            .position(|edit| matches!(edit, Edit::FramesetLayout(_)))
        {
            if base == value {
                self.edits.remove(index);
            } else {
                self.edits[index] = Edit::FramesetLayout(value);
            }
        } else if base != value {
            self.edits.push(Edit::FramesetLayout(value));
        }
    }

    fn record_div_border(&mut self, id: Id, side: BorderSide, value: Option<Border>) {
        let base = self
            .base
            .settings
            .get(id)
            .ok()
            .flatten()
            .and_then(|div| div.borders())
            .and_then(|borders| border_at(borders, side))
            .cloned();
        let existing = self.edits.iter().position(|edit| {
            matches!(edit, Edit::DivBorder { id: edit_id, side: edit_side, .. } if *edit_id == id && *edit_side == side)
        });
        if let Some(index) = existing {
            if base == value {
                self.edits.remove(index);
            } else {
                self.edits[index] = Edit::DivBorder { id, side, value };
            }
        } else if base != value {
            self.edits.push(Edit::DivBorder { id, side, value });
        }
    }
}

/// A successful web-settings publication containing its new snapshot and
/// reversible source patch.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Borrow the published snapshot.
    #[inline]
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Move the published snapshot out of the commit.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Borrow the reversible patch.
    #[inline]
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Move the reversible patch out of the commit.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// A source-checked, reversible web-settings patch.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Settings,
    after: Settings,
    before_xml: Arc<[u8]>,
    after_xml: Arc<[u8]>,
    conformance: Conformance,
}

impl Patch {
    /// Borrow the typed source precondition.
    #[inline]
    #[must_use]
    pub const fn before(&self) -> &Settings {
        &self.before
    }

    /// Borrow the typed state produced by the patch.
    #[inline]
    #[must_use]
    pub const fn after(&self) -> &Settings {
        &self.after
    }

    /// Return the source conformance carried by this patch.
    #[inline]
    #[must_use]
    pub const fn conformance(&self) -> Conformance {
        self.conformance
    }

    /// Return the inverse source-checked operation.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            before_xml: self.after_xml.clone(),
            after_xml: self.before_xml.clone(),
            conformance: self.conformance,
        }
    }

    /// Apply only to the exact source bytes and typed state captured by this
    /// patch.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.conformance != self.conformance
            || source.xml.as_ref() != self.before_xml.as_ref()
            || source.settings != self.before
        {
            return Err(Error::InvalidFormat(
                "web-settings patch source does not match its precondition".into(),
            ));
        }
        if self.before == self.after {
            return Ok(source.clone());
        }
        Snapshot::from_xml_with_conformance(self.after_xml.to_vec(), self.conformance)
    }
}

fn border_at(borders: &super::model::Borders, side: BorderSide) -> Option<&Border> {
    match side {
        BorderSide::Top => borders.top(),
        BorderSide::Left => borders.left(),
        BorderSide::Bottom => borders.bottom(),
        BorderSide::Right => borders.right(),
    }
}

fn set_border(borders: &mut super::model::Borders, side: BorderSide, value: Border) {
    match side {
        BorderSide::Top => {
            borders.set_top(value);
        },
        BorderSide::Left => {
            borders.set_left(value);
        },
        BorderSide::Bottom => {
            borders.set_bottom(value);
        },
        BorderSide::Right => {
            borders.set_right(value);
        },
    }
}

fn clear_border(borders: &mut super::model::Borders, side: BorderSide) {
    match side {
        BorderSide::Top => {
            borders.clear_top();
        },
        BorderSide::Left => {
            borders.clear_left();
        },
        BorderSide::Bottom => {
            borders.clear_bottom();
        },
        BorderSide::Right => {
            borders.clear_right();
        },
    }
}

fn all_borders_absent(borders: &super::model::Borders) -> bool {
    borders.top().is_none()
        && borders.left().is_none()
        && borders.bottom().is_none()
        && borders.right().is_none()
}
