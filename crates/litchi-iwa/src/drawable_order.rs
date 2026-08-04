//! Typed stacking-order operations for native iWork drawables.

use std::collections::{HashMap, HashSet};
use std::fmt::Display;
use std::hash::Hash;

use prost::Message;

use crate::protobuf::tsp;
use crate::wire::{repeated_length_delimited_payloads, rewrite_repeated_length_delimited_fields};
use crate::{Error, Result};

const BACK_LAYER_INDEX: usize = 0;
const LAYER_STEP: usize = 1;

/// A native Arrange command that changes one drawable's stacking layer.
///
/// iWork orders drawable lists from back to front. The variants correspond to
/// the native Send to Back, Send Backward, Bring Forward, and Bring to Front
/// controls.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum DrawableLayerMove {
    /// Move the drawable to the back-most layer.
    ToBack,
    /// Move the drawable one layer toward the back.
    Backward,
    /// Move the drawable one layer toward the front.
    Forward,
    /// Move the drawable to the front-most layer.
    ToFront,
}

/// Move one identifier within a back-to-front drawable list.
///
/// Returns `None` when the requested move would leave the list unchanged.
pub(crate) fn move_drawable_layer<T>(
    current: &[T],
    drawable_object_id: T,
    movement: DrawableLayerMove,
) -> Result<Option<Vec<T>>>
where
    T: Copy + Eq + Hash + Display,
{
    validate_unique_drawables(current, "current drawable order")?;
    let current_index = current
        .iter()
        .position(|identifier| *identifier == drawable_object_id)
        .ok_or_else(|| {
            Error::ParseError(format!(
                "drawable object {drawable_object_id} is not present in this stacking order"
            ))
        })?;
    let final_index = current
        .len()
        .checked_sub(1)
        .ok_or_else(|| Error::InvalidFormat("drawable order cannot be empty".to_owned()))?;
    let target_index = match movement {
        DrawableLayerMove::ToBack => BACK_LAYER_INDEX,
        DrawableLayerMove::Backward => current_index.saturating_sub(LAYER_STEP),
        DrawableLayerMove::Forward => current_index.saturating_add(LAYER_STEP).min(final_index),
        DrawableLayerMove::ToFront => final_index,
    };
    if target_index == current_index {
        return Ok(None);
    }

    let mut reordered = current.to_vec();
    let drawable_object_id = reordered.remove(current_index);
    reordered.insert(target_index, drawable_object_id);
    Ok(Some(reordered))
}

/// Rewrite one repeated protobuf reference field as an exact order permutation.
///
/// Every raw reference payload is retained verbatim, including unrecognized
/// fields, while unrelated outer fields retain their original positions.
pub(crate) fn reorder_reference_field(
    data: &[u8],
    field_number: u32,
    current: &[u64],
    requested: &[u64],
) -> Result<Vec<u8>> {
    validate_unique_drawables(current, "current drawable order")?;
    validate_exact_permutation(current, requested)?;

    let payloads = repeated_length_delimited_payloads(data, field_number)?;
    if payloads.len() != current.len() {
        return Err(Error::InvalidFormat(format!(
            "protobuf drawable-order field {field_number} has {} raw references but {} decoded references",
            payloads.len(),
            current.len()
        )));
    }

    let mut payload_indexes = HashMap::with_capacity(current.len());
    for (index, (&expected, payload)) in current.iter().zip(&payloads).enumerate() {
        let reference = tsp::Reference::decode(*payload)?;
        if reference.identifier != expected {
            return Err(Error::InvalidFormat(format!(
                "protobuf drawable-order field {field_number} changed during mutation"
            )));
        }
        if payload_indexes.insert(expected, index).is_some() {
            return Err(Error::InvalidFormat(format!(
                "protobuf drawable-order field {field_number} repeats drawable {expected}"
            )));
        }
    }

    let replacements = requested
        .iter()
        .map(|identifier| {
            let index = payload_indexes.get(identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "requested drawable order contains unknown object {identifier}"
                ))
            })?;
            Ok(payloads[*index].to_vec())
        })
        .collect::<Result<Vec<_>>>()?;
    rewrite_repeated_length_delimited_fields(data, field_number, &replacements)
}

pub(crate) fn validate_unique_drawables<T>(order: &[T], label: &str) -> Result<()>
where
    T: Copy + Eq + Hash + Display,
{
    let mut indexes = HashMap::with_capacity(order.len());
    for (index, &identifier) in order.iter().enumerate() {
        if let Some(previous) = indexes.insert(identifier, index) {
            return Err(Error::InvalidFormat(format!(
                "{label} repeats drawable object {identifier} at indexes {previous} and {index}"
            )));
        }
    }
    Ok(())
}

fn validate_exact_permutation<T>(current: &[T], requested: &[T]) -> Result<()>
where
    T: Copy + Eq + Hash + Display,
{
    if current.len() != requested.len() {
        return Err(Error::ParseError(format!(
            "requested drawable order has {} entries but this scope owns {}",
            requested.len(),
            current.len()
        )));
    }
    validate_unique_drawables(requested, "requested drawable order")?;
    let current_members = current.iter().copied().collect::<HashSet<_>>();
    for &identifier in requested {
        if !current_members.contains(&identifier) {
            return Err(Error::ParseError(format!(
                "requested drawable order contains object {identifier} outside this scope"
            )));
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::wire::{append_length_delimited_field, repeated_length_delimited_payloads};

    #[test]
    fn reorder_preserves_reference_payloads_and_unrelated_wire() {
        let mut first = tsp::Reference {
            identifier: 10,
            ..Default::default()
        }
        .encode_to_vec();
        append_unknown_varint(&mut first, 99, 10_001);
        let mut second = tsp::Reference {
            identifier: 20,
            ..Default::default()
        }
        .encode_to_vec();
        append_unknown_varint(&mut second, 99, 20_001);

        let mut original = crate::varint::encode_varint(90_u64 << 3);
        original.extend(crate::varint::encode_varint(1));
        append_length_delimited_field(&mut original, 7, &first).unwrap();
        original.extend(crate::varint::encode_varint(91_u64 << 3));
        original.extend(crate::varint::encode_varint(2));
        append_length_delimited_field(&mut original, 7, &second).unwrap();

        let reordered = reorder_reference_field(&original, 7, &[10, 20], &[20, 10]).unwrap();
        assert_eq!(
            repeated_length_delimited_payloads(&reordered, 7).unwrap(),
            vec![&second[..], &first[..]]
        );
        let restored = reorder_reference_field(&reordered, 7, &[20, 10], &[10, 20]).unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn move_drawable_layer_supports_each_native_arrange_command() {
        assert_eq!(
            move_drawable_layer(&[10, 20, 30, 40], 30, DrawableLayerMove::ToBack).unwrap(),
            Some(vec![30, 10, 20, 40])
        );
        assert_eq!(
            move_drawable_layer(&[10, 20, 30, 40], 30, DrawableLayerMove::Backward).unwrap(),
            Some(vec![10, 30, 20, 40])
        );
        assert_eq!(
            move_drawable_layer(&[10, 20, 30, 40], 20, DrawableLayerMove::Forward).unwrap(),
            Some(vec![10, 30, 20, 40])
        );
        assert_eq!(
            move_drawable_layer(&[10, 20, 30, 40], 20, DrawableLayerMove::ToFront).unwrap(),
            Some(vec![10, 30, 40, 20])
        );
        assert_eq!(
            move_drawable_layer(&[10, 20, 30], 10, DrawableLayerMove::ToBack).unwrap(),
            None
        );
        assert!(move_drawable_layer(&[10, 20], 30, DrawableLayerMove::ToFront).is_err());
    }

    fn append_unknown_varint(data: &mut Vec<u8>, field_number: u32, value: u64) {
        data.extend(crate::varint::encode_varint(u64::from(field_number) << 3));
        data.extend(crate::varint::encode_varint(value));
    }
}
