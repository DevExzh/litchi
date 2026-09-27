//! `OfficeArt` group/container record family.
//!
//! This is the bounded missing authoring path for nested `SpgrContainer`
//! records described by [MS-ODRAW] sections 2.2.14, 2.2.16, 2.2.38, and 2.2.40.

#![allow(
    dead_code,
    reason = "encoders cover the full [MS-ODRAW] record surface even where no caller exists yet"
)]

use zerocopy::IntoBytes;

use litchi_odraw::shape::Flags;
use litchi_odraw::write::{Sp, record_type, shape_type};

use super::super::{Error, EscherDgData, GroupChild, GroupShape, header_version};
use super::shapes::{create_child_shape_container, create_root_group_shape_container};
use super::validation::validate_group;
use super::wire::EscherBuilder;

#[derive(Clone, Copy, PartialEq, Eq)]
enum GroupRole {
    Patriarch,
    RootMember,
    NestedMember,
}

/// Encodes one top-level group shape, retaining nested group and child order.
pub(crate) fn create_group_shape_container(group: &GroupShape) -> Result<Vec<u8>, Error> {
    validate_group(group)?;
    build_group_container(group, GroupRole::Patriarch)
}

/// Encodes a complete drawing whose root is the supplied group tree.
pub(crate) fn create_dg_container_with_group(
    drawing_id: u32,
    group: &GroupShape,
) -> Result<Vec<u8>, Error> {
    let shape_count = validate_group(group)?;
    let mut drawing = EscherBuilder::new(header_version::CONTAINER, 0, record_type::DG_CONTAINER);

    #[allow(
        clippy::cast_possible_truncation,
        reason = "the `DG` instance field carries the low 12 bits of the drawing id; `EscherHeader::new` masks it"
    )]
    let mut dg = EscherBuilder::new(header_version::DG, drawing_id as u16, record_type::DG);
    dg.add_data(EscherDgData::new(shape_count, drawing_id).as_bytes());
    drawing.add_data(&dg.build()?);
    drawing.add_data(&build_group_container(group, GroupRole::Patriarch)?);
    drawing.build()
}

fn build_group_container(group: &GroupShape, role: GroupRole) -> Result<Vec<u8>, Error> {
    let mut container =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::SPGR_CONTAINER);
    container.add_data(&build_group_shape_header(group, role)?);

    for child in group.children() {
        let bytes = match child {
            GroupChild::Shape(shape) if role == GroupRole::Patriarch => {
                create_root_group_shape_container(shape.id, &shape.data, shape.anchor)?
            },
            GroupChild::Shape(shape) => {
                create_child_shape_container(shape.id, &shape.data, shape.anchor)?
            },
            GroupChild::Group(nested) if role == GroupRole::Patriarch => {
                build_group_container(nested, GroupRole::RootMember)?
            },
            GroupChild::Group(nested) => build_group_container(nested, GroupRole::NestedMember)?,
        };
        container.add_data(&bytes);
    }
    container.build()
}

fn build_group_shape_header(group: &GroupShape, role: GroupRole) -> Result<Vec<u8>, Error> {
    let mut container = EscherBuilder::new(header_version::CONTAINER, 0, record_type::SP_CONTAINER);

    let mut coordinate_space = EscherBuilder::new(header_version::SPGR, 0, record_type::SPGR);
    coordinate_space.add_data(group.coordinate_space.as_bytes());
    container.add_data(&coordinate_space.build()?);

    let flags = Flags::GROUP
        | match role {
            GroupRole::Patriarch => Flags::PATRIARCH,
            GroupRole::RootMember => Flags::HAVE_ANCHOR,
            GroupRole::NestedMember => Flags::CHILD | Flags::HAVE_ANCHOR,
        };
    let mut shape = EscherBuilder::new(
        header_version::SP,
        shape_type::NOT_PRIMITIVE,
        record_type::SP,
    );
    shape.add_data(Sp::with_flags(group.id, flags).as_bytes());
    container.add_data(&shape.build()?);

    if let Some(anchor_data) = group.anchor {
        let anchor_kind = match role {
            GroupRole::Patriarch | GroupRole::RootMember => record_type::CLIENT_ANCHOR,
            GroupRole::NestedMember => record_type::CHILD_ANCHOR,
        };
        let mut anchor = EscherBuilder::new(header_version::SIMPLE, 0, anchor_kind);
        anchor.add_data(anchor_data.as_bytes());
        container.add_data(&anchor.build()?);
    }
    container.build()
}
