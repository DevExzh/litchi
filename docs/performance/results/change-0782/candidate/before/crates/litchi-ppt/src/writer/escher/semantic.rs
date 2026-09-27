//! PPT semantic inputs for `OfficeArt` group and shape records.
//!
//! These types model only the bounded group grammar owned by this writer. Host
//! records such as `ClientData` remain on `UserShapeData`, so a group can reuse
//! the same typed shape authoring path without duplicating shape properties.

#![allow(
    dead_code,
    reason = "typed shape authoring helpers retained for pending call sites"
)]

use super::{ChildAnchor, EscherSpgrData, UserShapeData};

/// A shape that is a member of an `OfficeArt` group.
#[derive(Debug, Clone)]
pub(crate) struct ChildShape {
    pub(crate) id: u32,
    pub(crate) anchor: ChildAnchor,
    pub(crate) data: UserShapeData,
}

impl ChildShape {
    /// Creates a child shape with coordinates in its containing group space.
    pub(crate) const fn new(id: u32, anchor: ChildAnchor, data: UserShapeData) -> Self {
        Self { id, anchor, data }
    }
}

/// A nested `OfficeArt` group shape.
#[derive(Debug, Clone)]
pub(crate) struct GroupShape {
    pub(crate) id: u32,
    pub(crate) anchor: Option<ChildAnchor>,
    pub(crate) coordinate_space: EscherSpgrData,
    pub(crate) children: Vec<GroupChild>,
}

impl GroupShape {
    /// Creates a top-level patriarch group.
    pub(crate) fn new(id: u32, coordinate_space: EscherSpgrData) -> Self {
        Self {
            id,
            anchor: None,
            coordinate_space,
            children: Vec::new(),
        }
    }

    /// Creates a nested group whose anchor is expressed in its parent space.
    pub(crate) fn nested(id: u32, anchor: ChildAnchor, coordinate_space: EscherSpgrData) -> Self {
        Self {
            id,
            anchor: Some(anchor),
            coordinate_space,
            children: Vec::new(),
        }
    }

    /// Appends a child shape and returns the updated group.
    pub(crate) fn with_shape(mut self, id: u32, anchor: ChildAnchor, data: UserShapeData) -> Self {
        self.push_shape(id, anchor, data);
        self
    }

    /// Appends a child shape in the group's coordinate system.
    pub(crate) fn push_shape(&mut self, id: u32, anchor: ChildAnchor, data: UserShapeData) {
        self.children
            .push(GroupChild::Shape(Box::new(ChildShape::new(
                id, anchor, data,
            ))));
    }

    /// Appends a nested group and returns the updated group.
    pub(crate) fn with_group(mut self, group: GroupShape) -> Self {
        self.push_group(group);
        self
    }

    /// Appends a nested group whose own anchor is already in this group's
    /// coordinate system.
    pub(crate) fn push_group(&mut self, group: GroupShape) {
        self.children.push(GroupChild::Group(group));
    }

    /// Returns the group identifier.
    pub(crate) const fn id(&self) -> u32 {
        self.id
    }

    /// Returns the immediate children in record order.
    pub(crate) fn children(&self) -> &[GroupChild] {
        &self.children
    }
}

/// One record family contained by an `OfficeArtSpgrContainer`.
#[derive(Debug, Clone)]
pub(crate) enum GroupChild {
    /// A normal child `SpContainer`.
    Shape(Box<ChildShape>),
    /// A nested `SpgrContainer`.
    Group(GroupShape),
}
