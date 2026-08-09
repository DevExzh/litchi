//! Bounds and structural validation for `DrawingML` transforms.

use crate::{Error, Result};

use super::model::{Point, Size, Transform};

const MAX_SERIALIZED_BYTES: usize = 1 << 20;

/// Validate a detached transform before it is written or committed.
///
/// Coordinates, extents, and angles are private checked domains, so this
/// pass focuses on the aggregate output budget and keeps the validation seam
/// explicit for future `DrawingML` transform extensions.
pub fn validate(transform: &Transform) -> Result<()> {
    let mut size = 7usize; // `<a:xfrm/>`
    if let Some(value) = transform.authored_rotation() {
        size = size
            .checked_add(9 + value.value().to_string().len())
            .ok_or_else(|| limit("serialized transform bytes"))?;
    }
    for value in [
        transform.authored_flip_horizontal(),
        transform.authored_flip_vertical(),
    ] {
        if value.is_some() {
            size = size
                .checked_add(9)
                .ok_or_else(|| limit("serialized transform bytes"))?;
        }
    }
    for value in [transform.offset(), transform.child_offset()]
        .into_iter()
        .flatten()
    {
        size = add_point(size, value)?;
    }
    for value in [transform.extent(), transform.child_extent()]
        .into_iter()
        .flatten()
    {
        size = add_size(size, value)?;
    }
    if size > MAX_SERIALIZED_BYTES {
        return Err(limit("serialized transform bytes"));
    }
    Ok(())
}

fn add_point(size: usize, point: &Point) -> Result<usize> {
    size.checked_add(24 + point.x().to_string().len() + point.y().to_string().len())
        .ok_or_else(|| limit("serialized transform bytes"))
}

fn add_size(size: usize, value: Size) -> Result<usize> {
    size.checked_add(25 + value.width().to_string().len() + value.height().to_string().len())
        .ok_or_else(|| limit("serialized transform bytes"))
}

fn limit(resource: &'static str) -> Error {
    Error::Limit {
        resource,
        limit: MAX_SERIALIZED_BYTES,
    }
}
