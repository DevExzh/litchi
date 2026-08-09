use crate::{Error, Record, RecordKind, Result};
use std::collections::HashMap;

use super::model::{Anchor, Array, Id, Prop, Value};
use super::package::Props;
use super::{IS_BLIP, IS_COMPLEX, PROPERTY_ID_MASK};

#[derive(Debug, Clone, Copy)]
struct Descriptor {
    id: Id,
    raw_id: u16,
    blip: bool,
    complex: bool,
    raw_value: i32,
}

impl<'data> Props<'data> {
    /// Parses an Opt-family record while borrowing all complex values.
    ///
    /// # Errors
    ///
    /// Returns `Error::MalformedProperties` when the record is not an
    /// Opt-family table, has the wrong `recVer`, declares truncated or
    /// overlapping headers and complex payloads, repeats a property
    /// identifier, or leaves unclaimed trailing bytes, and
    /// `Error::ArithmeticOverflow` when a computed offset overflows.
    pub fn parse(opt: &Record<'data>) -> Result<Self> {
        if !matches!(
            opt.kind(),
            RecordKind::Opt | RecordKind::SecondaryOpt | RecordKind::TertiaryOpt
        ) {
            return Err(Error::MalformedProperties {
                reason: "record is not an Opt-family property table",
            });
        }
        if opt.version() != 3 {
            return Err(Error::MalformedProperties {
                reason: "Opt-family property table must have recVer 3",
            });
        }

        let count = usize::from(opt.instance());
        let data = opt.data();
        let header_size = count.checked_mul(6).ok_or(Error::ArithmeticOverflow {
            context: "property-header table length",
        })?;
        if header_size > data.len() {
            return Err(Error::MalformedProperties {
                reason: "property-header table exceeds recLen",
            });
        }

        let mut descriptors = Vec::with_capacity(count);
        let mut by_id = HashMap::with_capacity(count);
        for index in 0..count {
            let offset = index.checked_mul(6).ok_or(Error::ArithmeticOverflow {
                context: "property-header offset",
            })?;
            let end = offset.checked_add(6).ok_or(Error::ArithmeticOverflow {
                context: "property-header extent",
            })?;
            let header = data.get(offset..end).ok_or(Error::MalformedProperties {
                reason: "truncated property header",
            })?;
            let raw_opid = u16::from_le_bytes([header[0], header[1]]);
            let raw_id = raw_opid & PROPERTY_ID_MASK;
            let id = Id::from(raw_id);
            if by_id.insert(id, index).is_some() {
                return Err(Error::MalformedProperties {
                    reason: "duplicate property identifier",
                });
            }
            descriptors.push(Descriptor {
                id,
                raw_id,
                blip: raw_opid & IS_BLIP != 0,
                complex: raw_opid & IS_COMPLEX != 0,
                raw_value: i32::from_le_bytes([header[2], header[3], header[4], header[5]]),
            });
        }

        let mut properties = Vec::with_capacity(count);
        let mut complex_offset = header_size;
        for descriptor in descriptors {
            let value = if descriptor.complex {
                let complex_len = usize::try_from(descriptor.raw_value).map_err(|_err| {
                    Error::MalformedProperties {
                        reason: "negative complex-property length",
                    }
                })?;
                let complex_end =
                    complex_offset
                        .checked_add(complex_len)
                        .ok_or(Error::ArithmeticOverflow {
                            context: "complex-property extent",
                        })?;
                let complex =
                    data.get(complex_offset..complex_end)
                        .ok_or(Error::MalformedProperties {
                            reason: "complex property exceeds recLen",
                        })?;
                complex_offset = complex_end;

                if descriptor.id.is_array() {
                    Value::Array(Array::new(complex)?)
                } else {
                    Value::Complex(complex)
                }
            } else {
                Value::Simple(descriptor.raw_value)
            };
            properties.push(Prop {
                id: descriptor.id,
                raw_id: descriptor.raw_id,
                blip: descriptor.blip,
                complex: descriptor.complex,
                raw_value: descriptor.raw_value,
                value,
            });
        }

        if complex_offset != data.len() {
            return Err(Error::MalformedProperties {
                reason: "unclaimed bytes follow the property table",
            });
        }
        Ok(Self { properties, by_id })
    }
}

impl Anchor {
    #[must_use]
    pub fn from_child_anchor(anchor: &Record<'_>) -> Option<Self> {
        if anchor.kind() != RecordKind::ChildAnchor || anchor.data().len() != 16 {
            return None;
        }

        let left = i32::from_le_bytes([
            anchor.data()[0],
            anchor.data()[1],
            anchor.data()[2],
            anchor.data()[3],
        ]);
        let top = i32::from_le_bytes([
            anchor.data()[4],
            anchor.data()[5],
            anchor.data()[6],
            anchor.data()[7],
        ]);
        let right = i32::from_le_bytes([
            anchor.data()[8],
            anchor.data()[9],
            anchor.data()[10],
            anchor.data()[11],
        ]);
        let bottom = i32::from_le_bytes([
            anchor.data()[12],
            anchor.data()[13],
            anchor.data()[14],
            anchor.data()[15],
        ]);

        Some(Self::new(left, top, right, bottom))
    }
}
