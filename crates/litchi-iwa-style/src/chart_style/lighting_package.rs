//! The validation the host applies to a stored 3D lighting package before
//! it names the preset (retired host at commit `27b4840684`, `charts::lighting_3d::
//! validate_native_package`), at the wire level with `prost`'s decode
//! semantics: the last occurrence of a scalar wins, repeated occurrences of
//! a singular submessage merge, and a field with the wrong wire type is a
//! decode error.

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::varint::decode_varint_from_bytes;
use litchi_iwa_common::wire::WireView;

use crate::series::{SymbolError, SymbolResult};

const LENGTH_DELIMITED: u8 = 2;
const VARINT: u8 = 0;
const FIXED32: u8 = 5;

fn invalid<E>(_error: E) -> SymbolError {
    SymbolError::Invalid("an invalid 3D lighting package")
}

fn fields(payload: &[u8]) -> SymbolResult<WireView<'_>> {
    WireView::parse(payload, WireLimits::DEFAULT).map_err(invalid)
}

fn expect(wire_type: u8, expected: u8) -> SymbolResult<()> {
    if wire_type == expected {
        Ok(())
    } else {
        Err(SymbolError::Invalid("an invalid 3D lighting package"))
    }
}

fn float(payload: &[u8]) -> SymbolResult<f32> {
    let bytes: [u8; 4] = payload.try_into().map_err(invalid)?;
    Ok(f32::from_le_bytes(bytes))
}

fn string(payload: &[u8]) -> SymbolResult<&str> {
    std::str::from_utf8(payload).map_err(invalid)
}

/// A `TSCH.Chart3DVectorArchive` merged over every occurrence.
fn merge_vector(vector: &mut [f32; 4], payload: &[u8]) -> SymbolResult<()> {
    for field in fields(payload)?.fields() {
        if let Some(slot) = (field.number() as usize)
            .checked_sub(1)
            .and_then(|index| vector.get_mut(index))
        {
            expect(field.wire_type(), FIXED32)?;
            *slot = float(field.payload())?;
        }
    }
    Ok(())
}

fn finite(vector: &[f32; 4]) -> bool {
    vector.iter().all(|component| component.is_finite())
}

/// A light source sub-message: its vectors (by field number) and floats.
#[derive(Default)]
struct Source {
    present: bool,
    vectors: [[f32; 4]; 2],
    floats: [f32; 2],
}

impl Source {
    /// Merge one occurrence; `vectors` are the vector field numbers (1 or
    /// 1 and 2), floats follow them.
    fn merge(&mut self, payload: &[u8], vectors: u32) -> SymbolResult<()> {
        self.present = true;
        for field in fields(payload)?.fields() {
            let number = field.number();
            if (1..=vectors).contains(&number) {
                expect(field.wire_type(), LENGTH_DELIMITED)?;
                merge_vector(&mut self.vectors[number as usize - 1], field.payload())?;
            } else if vectors == 2 && (3..=4).contains(&number) {
                expect(field.wire_type(), FIXED32)?;
                self.floats[number as usize - 3] = float(field.payload())?;
            }
        }
        Ok(())
    }

    fn valid(&self) -> bool {
        self.vectors.iter().all(finite) && self.floats.iter().all(|value| value.is_finite())
    }
}

/// Whether one `TSCH.Chart3DLightArchive` passes the host's checks.
fn valid_light(payload: &[u8]) -> SymbolResult<bool> {
    let mut name = "";
    let mut colors = [[0.0_f32; 4]; 4];
    let mut intensity = 0.0_f32;
    let mut coordinate_space = 0_u32;
    let mut sources: [Source; 3] = Default::default();
    for field in fields(payload)?.fields() {
        match field.number() {
            1 => {
                expect(field.wire_type(), LENGTH_DELIMITED)?;
                name = string(field.payload())?;
            },
            number @ (2 | 3 | 4 | 6) => {
                expect(field.wire_type(), LENGTH_DELIMITED)?;
                let index = match number {
                    2 => 0,
                    3 => 1,
                    4 => 2,
                    _ => 3,
                };
                merge_vector(&mut colors[index], field.payload())?;
            },
            5 => {
                expect(field.wire_type(), FIXED32)?;
                intensity = float(field.payload())?;
            },
            7 => {
                expect(field.wire_type(), VARINT)?;
                let (value, _) = decode_varint_from_bytes(field.payload()).map_err(invalid)?;
                // `prost` truncates a `uint32` varint.
                coordinate_space = value as u32;
            },
            number @ 9..=11 => {
                expect(field.wire_type(), LENGTH_DELIMITED)?;
                let vectors = if number == 11 { 2 } else { 1 };
                sources[number as usize - 9].merge(field.payload(), vectors)?;
            },
            _ => {},
        }
    }
    let present = sources.iter().filter(|source| source.present).count();
    Ok(!name.is_empty()
        && intensity.is_finite()
        && intensity >= 0.0
        && colors.iter().all(finite)
        && coordinate_space <= 1
        && present == 1
        && sources
            .iter()
            .all(|source| !source.present || source.valid()))
}

/// The package name of a stored `TSCH.Chart3DLightingPackageArchive`
/// that passes the host's checks: a name, at least one light, and every
/// light named, with finite colors, attenuation and source vectors, a
/// finite non-negative intensity, a known coordinate space and exactly one
/// source.
pub(super) fn validated_name(package: &[u8]) -> SymbolResult<&str> {
    let mut name = "";
    let mut lights = 0usize;
    for field in fields(package)?.fields() {
        match field.number() {
            1 => {
                expect(field.wire_type(), LENGTH_DELIMITED)?;
                name = string(field.payload())?;
            },
            2 => {
                expect(field.wire_type(), LENGTH_DELIMITED)?;
                lights += 1;
                if !valid_light(field.payload())? {
                    return Err(SymbolError::Invalid(
                        "a 3D lighting package contains an invalid light",
                    ));
                }
            },
            _ => {},
        }
    }
    if name.is_empty() || lights == 0 {
        return Err(SymbolError::Invalid(
            "a 3D lighting package without a name or a light",
        ));
    }
    Ok(name)
}
