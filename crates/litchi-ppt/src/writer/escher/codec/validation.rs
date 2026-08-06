//! Validation shared by OfficeArt shape-family encoders.

#![allow(dead_code)]

use std::collections::HashSet;
use std::io::{Error, ErrorKind};

use super::super::{GroupChild, GroupShape, UserShapeData};

const MAX_GROUP_DEPTH: usize = 64;
const MAX_SHAPE_ID: u32 = 0x3FFF_FFFF;

/// Validates one PPT shape before its record family is encoded.
pub(super) fn validate_user_shape(shape: &UserShapeData) -> Result<(), Error> {
    if shape.adjust_values.len() > 10 {
        return Err(invalid_input(
            "OfficeArt shapes support at most 10 adjustment values",
        ));
    }
    if let Some(geometry) = &shape.freeform_geometry {
        geometry.validate().map_err(|error| {
            Error::new(
                ErrorKind::InvalidInput,
                format!("invalid freeform geometry: {error}"),
            )
        })?;
    }
    Ok(())
}

/// Validates the bounded nested-group grammar and returns its shape count.
pub(super) fn validate_group(root: &GroupShape) -> Result<u32, Error> {
    if root.anchor.is_some() {
        return Err(invalid_input(
            "the top-level OfficeArt group cannot have a ChildAnchor",
        ));
    }
    let mut ids = HashSet::new();
    visit_group(root, 0, &mut ids)
}

fn visit_group(group: &GroupShape, depth: usize, ids: &mut HashSet<u32>) -> Result<u32, Error> {
    if depth >= MAX_GROUP_DEPTH {
        return Err(invalid_input("OfficeArt group nesting exceeds 64 levels"));
    }
    validate_shape_id(group.id, ids)?;

    let mut count = 1_u32;
    for child in &group.children {
        let child_count = match child {
            GroupChild::Shape(shape) => {
                validate_shape_id(shape.id, ids)?;
                validate_user_shape(&shape.data)?;
                1
            },
            GroupChild::Group(nested) => {
                if nested.anchor.is_none() {
                    return Err(invalid_input(
                        "a nested OfficeArt group must have a ChildAnchor",
                    ));
                }
                visit_group(nested, depth + 1, ids)?
            },
        };
        count = count
            .checked_add(child_count)
            .ok_or_else(|| invalid_input("OfficeArt group shape count overflows u32"))?;
    }
    Ok(count)
}

fn validate_shape_id(id: u32, ids: &mut HashSet<u32>) -> Result<(), Error> {
    if id == 0 || id > MAX_SHAPE_ID {
        return Err(invalid_input(
            "OfficeArt shape identifiers must be non-zero 30-bit values",
        ));
    }
    if !ids.insert(id) {
        return Err(invalid_input("OfficeArt shape identifiers must be unique"));
    }
    Ok(())
}

fn invalid_input(message: &'static str) -> Error {
    Error::new(ErrorKind::InvalidInput, message)
}
