//! Package-level `OfficeArtClientData` edits (record 0357).
//!
//! [`Package::edit_client_data`] stages edits of the client-data children of
//! one top-level shape, [`Edit::commit`] publishes a new [`Package`] through
//! one append-only user edit that rewrites only the shape's slide, and
//! [`Package::apply_client_data`] replays a [`Patch`] onto the exact package
//! it was made from. Dropping an [`Edit`] discards it.

use super::codec::OFFICE_ART_CLIENT_DATA_RECORD_TYPE;
use super::transaction::{RecordEdit, Snapshot};
use super::{Change, ClientData, ClientDataChild, Commit, Error, Patch, Result};
use crate::Package;
use crate::record_wire as wire;
use crate::scoped::{self, Base, Failure, Lazy};
use crate::slide_order::{self, records::Owner};
use crate::text_edit::Target;

const OPERATION: &str = "edit_client_data";
const MISSING_CHILD: &str = "the client data has no child at this index";

/// A staged edit of one shape's `OfficeArtClientData` over one immutable
/// [`Package`].
pub struct Edit {
    base: Base,
    target: Target,
    state: Lazy<State>,
}

struct State {
    native: slide_order::Snapshot,
    shape_id: u32,
    slide: Vec<u8>,
    source: Snapshot,
    inner: RecordEdit,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("target", &self.target)
            .field("initialized", &self.state.get().is_some())
            .finish_non_exhaustive()
    }
}

/// The index of the shape's one client-data child, or `None` without one.
fn client_data_index(shape: &[u8]) -> std::result::Result<Option<usize>, Failure> {
    let mut found = None;
    for (index, child) in wire::child_list(shape)?.iter().enumerate() {
        if wire::record_type(child)? == OFFICE_ART_CLIENT_DATA_RECORD_TYPE
            && found.replace(index).is_some()
        {
            return Err(scoped::unsupported_target(
                OPERATION,
                "the shape has several client-data records",
            ));
        }
    }
    Ok(found)
}

impl Edit {
    fn new(package: &Package, target: Target) -> Result<Self> {
        let base = Base::capture(package)?;
        Ok(Self {
            base,
            target,
            state: Lazy::new(),
        })
    }

    fn initialize_state(&self) -> std::result::Result<State, Failure> {
        let native = self.base.native()?;
        let slide = self
            .base
            .owner_record(&native, Owner::Slide(self.target.slide()))?;
        let shape_id = {
            let presentation = self.base.package().presentation()?;
            let slides = presentation.slides()?;
            let selected = slides
                .get(self.target.slide().get())
                .ok_or_else(|| scoped::not_found(OPERATION, "no slide is at this position"))?;
            let shapes = selected.shapes()?;
            let shape = shapes
                .get(self.target.shape().get())
                .ok_or_else(|| scoped::not_found(OPERATION, "the slide has no shape here"))?;
            crate::text_edit::native_shape_id(shape)
        };
        let mut record = None;
        let mut failure = None;
        crate::hyperlink::edit::rewrite_slide_shape(&slide, shape_id, &mut |shape| {
            match client_data_index(shape) {
                Ok(Some(index)) => record = Some(wire::child_list(shape)?[index].to_vec()),
                Ok(None) => {},
                Err(error) => failure = Some(error),
            }
            Ok(shape.to_vec())
        })
        .map_err(Failure::from)?;
        if let Some(failure) = failure {
            return Err(failure);
        }
        let record = record
            .ok_or_else(|| scoped::unsupported_target(OPERATION, "the shape has no client data"))?;
        let source = Snapshot::parse(&record, &super::ClientDataLimits::default())
            .map_err(scoped::publication(OPERATION))?;
        let inner = source.edit();
        Ok(State {
            native,
            shape_id,
            slide,
            source,
            inner,
        })
    }

    fn state(&self) -> std::result::Result<&State, Failure> {
        self.state.get_or_try_init(|| self.initialize_state())
    }

    fn state_mut(&mut self) -> std::result::Result<&mut State, Failure> {
        let base = self.base.clone();
        let target = self.target;
        self.state.get_or_try_init_mut(|| {
            // Rebuild the deferred source state at the first mutable use.
            let native = base.native()?;
            let slide = base.owner_record(&native, Owner::Slide(target.slide()))?;
            let shape_id = {
                let presentation = base.package().presentation()?;
                let slides = presentation.slides()?;
                let selected = slides
                    .get(target.slide().get())
                    .ok_or_else(|| scoped::not_found(OPERATION, "no slide is at this position"))?;
                let shapes = selected.shapes()?;
                let shape = shapes
                    .get(target.shape().get())
                    .ok_or_else(|| scoped::not_found(OPERATION, "the slide has no shape here"))?;
                crate::text_edit::native_shape_id(shape)
            };
            let mut record = None;
            let mut failure = None;
            crate::hyperlink::edit::rewrite_slide_shape(&slide, shape_id, &mut |shape| {
                match client_data_index(shape) {
                    Ok(Some(index)) => record = Some(wire::child_list(shape)?[index].to_vec()),
                    Ok(None) => {},
                    Err(error) => failure = Some(error),
                }
                Ok(shape.to_vec())
            })
            .map_err(Failure::from)?;
            if let Some(failure) = failure {
                return Err(failure);
            }
            let record = record.ok_or_else(|| {
                scoped::unsupported_target(OPERATION, "the shape has no client data")
            })?;
            let source = Snapshot::parse(&record, &super::ClientDataLimits::default())
                .map_err(scoped::publication(OPERATION))?;
            let inner = source.edit();
            Ok::<State, Failure>(State {
                native,
                shape_id,
                slide,
                source,
                inner,
            })
        })
    }

    fn require_child(&self, operation: &'static str, index: usize) -> Result<()> {
        if index < self.state()?.inner.children().len() {
            Ok(())
        } else {
            Err(scoped::not_found(operation, MISSING_CHILD).into())
        }
    }

    /// The shape whose client data is staged.
    #[must_use]
    pub const fn target(&self) -> Target {
        self.target
    }

    /// The staged client data.
    #[must_use = "handle the fallible client-data projection result"]
    pub fn client_data(&self) -> Result<&ClientData> {
        Ok(self.state()?.inner.client_data())
    }

    /// The staged client-data children, in order.
    #[must_use = "handle the fallible client-data children result"]
    pub fn children(&self) -> Result<&[ClientDataChild]> {
        Ok(self.state()?.inner.children())
    }

    /// The staged operations, in order.
    #[must_use = "handle the fallible client-data changes result"]
    pub fn changes(&self) -> Result<&[Change]> {
        Ok(self.state()?.inner.changes())
    }

    /// Insert one child at `index` (at most the child count).
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the index is past the end or
    /// the record grammar refuses the child, leaving the edit unchanged.
    pub fn insert(&mut self, index: usize, child: ClientDataChild) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .insert(index, child)
            .map_err(scoped::staging("client_data.insert"))?)
    }

    /// Append one child after the current children.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// child, leaving the edit unchanged.
    pub fn append(&mut self, child: ClientDataChild) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .append(child)
            .map_err(scoped::staging("client_data.append"))?)
    }

    /// Replace the child at `index`, returning the replaced child.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] with
    /// [`litchi_core::edit::RefusalKind::NotFound`] when no child is at
    /// `index`, and [`super::Error::Refused`] when the record grammar refuses
    /// the replacement, leaving the edit unchanged.
    pub fn replace(
        &mut self,
        index: usize,
        replacement: ClientDataChild,
    ) -> Result<ClientDataChild> {
        self.require_child("client_data.replace", index)?;
        Ok(self
            .state_mut()?
            .inner
            .replace(index, replacement)
            .map_err(scoped::staging("client_data.replace"))?)
    }

    /// Remove the child at `index`, returning it.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] with
    /// [`litchi_core::edit::RefusalKind::NotFound`] when no child is at
    /// `index`, and [`super::Error::Refused`] when the remaining children
    /// are not a valid client data, leaving the edit unchanged.
    pub fn remove(&mut self, index: usize) -> Result<ClientDataChild> {
        self.require_child("client_data.remove", index)?;
        Ok(self
            .state_mut()?
            .inner
            .remove(index)
            .map_err(scoped::staging("client_data.remove"))?)
    }

    /// Publish the staged client data as a new package. An edit with no net
    /// change publishes the source package unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Signed`] for a changed write to a signed
    /// package and [`super::Error::Package`] when the result cannot be
    /// republished.
    pub fn commit(self) -> Result<Commit> {
        self.state
            .get_or_try_init(|| self.initialize_state())
            .map_err(Error::from)?;
        let Self {
            base,
            target,
            state,
        } = self;
        let State {
            native,
            shape_id,
            slide,
            source,
            inner,
        } = state.into_inner().ok_or_else(|| {
            Failure::Package(crate::package::Error::Corrupted(
                "PPT client-data edit state was not initialized".into(),
            ))
        })?;
        let (record, _) = inner
            .commit()
            .map_err(scoped::publication("client_data.commit"))?
            .into_parts();
        if record.bytes() == source.bytes() {
            return Ok(super::published(base.unchanged()?));
        }
        let after = crate::hyperlink::edit::rewrite_slide_shape(&slide, shape_id, &mut |shape| {
            let children = wire::child_list(shape)?;
            let mut data = Vec::with_capacity(shape.len().saturating_add(record.bytes().len()));
            for child in &children {
                if wire::record_type(child)? == OFFICE_ART_CLIENT_DATA_RECORD_TYPE {
                    data.extend_from_slice(record.bytes());
                } else {
                    data.extend_from_slice(child);
                }
            }
            Ok(wire::rebuild(shape, &data)?)
        })
        .map_err(Failure::from)?;
        Ok(super::published(base.publish_owner(
            &native,
            Owner::Slide(target.slide()),
            slide,
            after,
        )?))
    }
}

impl Package {
    /// Start an edit of the `OfficeArtClientData` of one top-level shape.
    ///
    /// # Errors
    ///
    /// An encrypted presentation is refused immediately. Missing slides or
    /// shapes, absent client data, and unreadable records are reported by the
    /// first projection, staging operation, or commit.
    pub fn edit_client_data(&self, target: Target) -> Result<Edit> {
        Edit::new(self, target)
    }

    /// Apply a client-data patch to the exact package it was made from.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::PatchConflict`] when this package is not the
    /// patch's source.
    pub fn apply_client_data(&self, patch: &Patch) -> Result<Commit> {
        super::apply_patch(self, patch)
    }
}
