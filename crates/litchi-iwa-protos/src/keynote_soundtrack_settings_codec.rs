//! Focused read-only Keynote soundtrack settings projection.

use crate::buffa_keynote_soundtrack_settings_generated::LitchiIwaProjection as projection;
use buffa::DecodeOptions as BuffaDecodeOptions;

/// Finite decode policy for the soundtrack sidecar.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
}
impl DecodeOptions {
    #[must_use]
    pub const fn new(
        max_message_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
    ) -> Self {
        Self {
            max_message_bytes,
            max_fields,
            max_work_bytes,
            recursion_limit,
        }
    }
}
/// Successful resource consumption for transaction-budget merge.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    media_references: usize,
    media_bytes: usize,
}
impl DecodeReport {
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }
    #[must_use]
    pub const fn media_references(self) -> usize {
        self.media_references
    }
    #[must_use]
    pub const fn media_bytes(self) -> usize {
        self.media_bytes
    }
}
/// Typed limit classification.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DecodeLimit {
    Bytes { observed: usize, maximum: usize },
    Fields { observed: usize, maximum: usize },
    Work { observed: usize, maximum: usize },
    Nesting { observed: u32, maximum: u32 },
}
/// Strict decode failure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    limit: Option<DecodeLimit>,
}
impl DecodeError {
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        self.limit
    }
    const fn plain() -> Self {
        Self { limit: None }
    }
}
impl core::fmt::Display for DecodeError {
    fn fmt(&self, f: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        f.write_str("invalid Keynote soundtrack settings")
    }
}
impl std::error::Error for DecodeError {}

impl DecodeError {
    fn limit(limit: DecodeLimit) -> Self {
        Self { limit: Some(limit) }
    }
}
/// Scalar settings; media identifiers are exposed through a bounded visitor.
#[derive(Debug, Clone, PartialEq)]
pub struct SoundtrackSettingsSnapshot {
    volume: Option<f64>,
    mode_raw: Option<i32>,
}
impl SoundtrackSettingsSnapshot {
    #[must_use]
    pub const fn volume(&self) -> Option<f64> {
        self.volume
    }
    #[must_use]
    pub const fn mode_raw(&self) -> Option<i32> {
        self.mode_raw
    }
}
pub fn decode_soundtrack_settings(
    source: &[u8],
    options: DecodeOptions,
) -> Result<SoundtrackSettingsSnapshot, DecodeError> {
    Ok(decode_soundtrack_settings_with_report(source, options)?.0)
}
pub fn decode_soundtrack_settings_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(SoundtrackSettingsSnapshot, DecodeReport), DecodeError> {
    decode_soundtrack_settings_with_media(source, options, &mut |_identifier| Ok(()))
}
/// Decode scalar settings, cross-check Buffa once, and stream each validated
/// media identifier during the same strict root pass.
pub fn decode_soundtrack_settings_with_media(
    source: &[u8],
    options: DecodeOptions,
    visitor: &mut dyn FnMut(u64) -> Result<(), DecodeError>,
) -> Result<(SoundtrackSettingsSnapshot, DecodeReport), DecodeError> {
    let (strict, report) = strict(source, options, visitor)?;
    let view: projection::SoundtrackArchiveLazyView<'_> = decode(options, source)?;
    if view.volume.map(f64::to_bits) != strict.volume.map(f64::to_bits)
        || view.mode != strict.mode_raw
    {
        return Err(DecodeError::plain());
    }
    Ok((strict, report))
}
/// Compatibility spelling for transaction callers that emphasize the merged
/// resource report returned with the scalar snapshot.
pub fn decode_soundtrack_settings_with_media_report(
    source: &[u8],
    options: DecodeOptions,
    visitor: &mut dyn FnMut(u64) -> Result<(), DecodeError>,
) -> Result<(SoundtrackSettingsSnapshot, DecodeReport), DecodeError> {
    decode_soundtrack_settings_with_media(source, options, visitor)
}
/// Stream each validated native movie-media identifier without retaining an
/// input-width vector. The callback shares this decoder's exact report.
pub fn visit_soundtrack_media_identifiers(
    source: &[u8],
    options: DecodeOptions,
    visitor: &mut dyn FnMut(u64) -> Result<(), DecodeError>,
) -> Result<DecodeReport, DecodeError> {
    Ok(decode_soundtrack_settings_with_media(source, options, visitor)?.1)
}
fn strict(
    source: &[u8],
    options: DecodeOptions,
    visitor: &mut dyn FnMut(u64) -> Result<(), DecodeError>,
) -> Result<(SoundtrackSettingsSnapshot, DecodeReport), DecodeError> {
    if source.len() > options.max_message_bytes {
        return Err(DecodeError::limit(DecodeLimit::Bytes {
            observed: source.len(),
            maximum: options.max_message_bytes,
        }));
    }
    if options.recursion_limit == 0 || options.recursion_limit > 64 {
        return Err(DecodeError::limit(DecodeLimit::Nesting {
            observed: options.recursion_limit,
            maximum: 64,
        }));
    }
    let budget = Budget::new(options);
    budget.message(source.len(), 1)?;
    budget.message(source.len(), 1)?;
    // A complete root decode counts as depth one even when it has no fields.
    budget.depth(1)?;
    let mut input = source;
    let mut volume = None;
    let mut mode_raw = None;
    while let Some(field) = next(&mut input, &budget, 1)? {
        match field.number {
            1 => {
                if volume.is_some() || field.wire != 1 {
                    return Err(DecodeError::plain());
                }
                volume = Some(f64::from_le_bytes(
                    field.fixed64.ok_or(DecodeError::plain())?,
                ));
            },
            2 => {
                if mode_raw.is_some() || field.wire != 0 {
                    return Err(DecodeError::plain());
                }
                mode_raw = Some(int32(field.varint.ok_or(DecodeError::plain())?)?);
            },
            3 => {
                if field.wire != 2 {
                    return Err(DecodeError::plain());
                }
                let payload = field.bytes.ok_or(DecodeError::plain())?;
                budget.message(payload.len(), 1)?;
                budget.media(payload.len())?;
                visitor(data_reference(payload, &budget, 2)?)?;
            },
            _ => {},
        }
    }
    Ok((
        SoundtrackSettingsSnapshot { volume, mode_raw },
        budget.report(),
    ))
}

fn data_reference(source: &[u8], budget: &Budget, depth: u32) -> Result<u64, DecodeError> {
    let mut input = source;
    let mut identifier = None;
    while let Some(field) = next(&mut input, budget, depth)? {
        if field.number == 1 {
            if identifier.is_some() || field.wire != 0 {
                return Err(DecodeError::plain());
            }
            let id = field.varint.ok_or(DecodeError::plain())?;
            if id == 0 {
                return Err(DecodeError::plain());
            }
            identifier = Some(id);
        }
    }
    identifier.ok_or(DecodeError::plain())
}

#[derive(Clone, Copy)]
struct Field<'a> {
    number: u32,
    wire: u8,
    varint: Option<u64>,
    fixed64: Option<[u8; 8]>,
    bytes: Option<&'a [u8]>,
}
fn next<'a>(
    input: &mut &'a [u8],
    budget: &Budget,
    depth: u32,
) -> Result<Option<Field<'a>>, DecodeError> {
    if input.is_empty() {
        return Ok(None);
    }
    budget.field()?;
    budget.depth(depth)?;
    let tag = varint(input)?;
    let number = u32::try_from(tag >> 3).map_err(|_e| DecodeError::plain())?;
    let wire = u8::try_from(tag & 7).map_err(|_e| DecodeError::plain())?;
    if number == 0 || number > 0x1fff_ffff {
        return Err(DecodeError::plain());
    }
    let mut field = Field {
        number,
        wire,
        varint: None,
        fixed64: None,
        bytes: None,
    };
    match wire {
        0 => field.varint = Some(varint(input)?),
        1 => {
            let bytes = take(input, 8)?;
            field.fixed64 = Some(bytes.try_into().map_err(|_e| DecodeError::plain())?);
        },
        2 => {
            let length = usize::try_from(varint(input)?).map_err(|_e| DecodeError::plain())?;
            field.bytes = Some(take(input, length)?);
        },
        5 => {
            let _ = take(input, 4)?;
        },
        3 | 4 => return Err(DecodeError::plain()),
        _ => return Err(DecodeError::plain()),
    }
    Ok(Some(field))
}
fn take<'a>(input: &mut &'a [u8], count: usize) -> Result<&'a [u8], DecodeError> {
    if input.len() < count {
        return Err(DecodeError::plain());
    }
    let (value, rest) = input.split_at(count);
    *input = rest;
    Ok(value)
}
fn varint(input: &mut &[u8]) -> Result<u64, DecodeError> {
    let original = *input;
    let mut value = 0u64;
    for index in 0..10 {
        let byte = *original.get(index).ok_or(DecodeError::plain())?;
        if index == 9 && byte > 1 {
            return Err(DecodeError::plain());
        }
        value |= u64::from(byte & 127) << (index * 7);
        if byte & 128 == 0 {
            let length = if value == 0 {
                1
            } else {
                usize::try_from((64 - value.leading_zeros()).div_ceil(7))
                    .map_err(|_e| DecodeError::plain())?
            };
            if length != index + 1 {
                return Err(DecodeError::plain());
            }
            *input = &original[index + 1..];
            return Ok(value);
        }
    }
    Err(DecodeError::plain())
}
fn int32(value: u64) -> Result<i32, DecodeError> {
    if let Ok(result) = i32::try_from(value) {
        return Ok(result);
    }
    if value >= 0xffff_ffff_8000_0000 {
        return i32::try_from(i64::from_ne_bytes(value.to_ne_bytes()))
            .map_err(|_e| DecodeError::plain());
    }
    Err(DecodeError::plain())
}

struct Budget {
    fields: std::cell::Cell<usize>,
    work: std::cell::Cell<usize>,
    depth: std::cell::Cell<u32>,
    media_references: std::cell::Cell<usize>,
    media_bytes: std::cell::Cell<usize>,
    options: DecodeOptions,
}
impl Budget {
    fn new(options: DecodeOptions) -> Self {
        Self {
            fields: std::cell::Cell::new(0),
            work: std::cell::Cell::new(0),
            depth: std::cell::Cell::new(0),
            media_references: std::cell::Cell::new(0),
            media_bytes: std::cell::Cell::new(0),
            options,
        }
    }
    fn field(&self) -> Result<(), DecodeError> {
        let value = self
            .fields
            .get()
            .checked_add(1)
            .ok_or(DecodeError::plain())?;
        if value > self.options.max_fields {
            return Err(DecodeError::limit(DecodeLimit::Fields {
                observed: value,
                maximum: self.options.max_fields,
            }));
        }
        self.fields.set(value);
        Ok(())
    }
    fn message(&self, bytes: usize, passes: usize) -> Result<(), DecodeError> {
        let add = bytes.checked_mul(passes).ok_or(DecodeError::plain())?;
        let value = self
            .work
            .get()
            .checked_add(add)
            .ok_or(DecodeError::plain())?;
        if value > self.options.max_work_bytes {
            return Err(DecodeError::limit(DecodeLimit::Work {
                observed: value,
                maximum: self.options.max_work_bytes,
            }));
        }
        self.work.set(value);
        Ok(())
    }
    fn depth(&self, value: u32) -> Result<(), DecodeError> {
        if value > self.options.recursion_limit {
            return Err(DecodeError::limit(DecodeLimit::Nesting {
                observed: value,
                maximum: self.options.recursion_limit,
            }));
        }
        self.depth.set(self.depth.get().max(value));
        Ok(())
    }
    fn media(&self, bytes: usize) -> Result<(), DecodeError> {
        self.media_references.set(
            self.media_references
                .get()
                .checked_add(1)
                .ok_or(DecodeError::plain())?,
        );
        self.media_bytes.set(
            self.media_bytes
                .get()
                .checked_add(bytes)
                .ok_or(DecodeError::plain())?,
        );
        Ok(())
    }
    fn report(&self) -> DecodeReport {
        DecodeReport {
            fields: self.fields.get(),
            work_bytes: self.work.get(),
            max_depth: self.depth.get(),
            media_references: self.media_references.get(),
            media_bytes: self.media_bytes.get(),
        }
    }
}

fn decode<'a, T: buffa::LazyMessageView<'a>>(
    options: DecodeOptions,
    source: &'a [u8],
) -> Result<T, DecodeError> {
    BuffaDecodeOptions::new()
        .with_max_message_size(options.max_message_bytes)
        .with_unknown_field_limit(options.max_fields)
        .with_element_memory_limit(0)
        .with_recursion_limit(options.recursion_limit)
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::plain())
}

#[cfg(test)]
mod tests {
    use super::*;
    const SOURCE: [u8; 15] = [9, 0, 0, 0, 0, 0, 0, 0xf0, 0x3f, 0x10, 1, 0x1a, 2, 8, 7];
    const OPTIONS: DecodeOptions = DecodeOptions::new(64, 8, 64, 2);
    #[test]
    fn streams_media_and_cross_checks_scalar_bits() {
        let (snapshot, report) = decode_soundtrack_settings_with_report(&SOURCE, OPTIONS).unwrap();
        assert_eq!(snapshot.volume().map(f64::to_bits), Some(1.0f64.to_bits()));
        assert_eq!(snapshot.mode_raw(), Some(1));
        let mut identifiers = Vec::new();
        let media_report =
            visit_soundtrack_media_identifiers(&SOURCE, OPTIONS, &mut |identifier| {
                identifiers.push(identifier);
                Ok(())
            })
            .unwrap();
        assert_eq!(identifiers, [7]);
        assert_eq!(report.fields(), 4);
        assert_eq!(report.media_references(), 1);
        assert_eq!(report.media_bytes(), 2);
        assert_eq!(report.max_depth(), 2);
        assert_eq!(report.work_bytes(), 32);
        assert_eq!(media_report, report);
    }
    #[test]
    fn distributed_field_and_work_limits_are_exact() {
        assert!(decode_soundtrack_settings(&SOURCE, DecodeOptions::new(15, 4, 32, 2)).is_ok());
        let fields =
            decode_soundtrack_settings(&SOURCE, DecodeOptions::new(64, 3, 64, 2)).unwrap_err();
        assert_eq!(
            fields.resource_limit(),
            Some(DecodeLimit::Fields {
                observed: 4,
                maximum: 3
            })
        );
        let work =
            decode_soundtrack_settings(&SOURCE, DecodeOptions::new(64, 8, 31, 2)).unwrap_err();
        assert_eq!(
            work.resource_limit(),
            Some(DecodeLimit::Work {
                observed: 32,
                maximum: 31
            })
        );
    }
    #[test]
    fn rejects_malformed_media_reference() {
        assert!(decode_soundtrack_settings(&[0x1a, 0], OPTIONS).is_err());
        assert!(decode_soundtrack_settings(&[0x1a, 2, 8, 0], OPTIONS).is_err());
        assert!(decode_soundtrack_settings(&[0x1a, 4, 8, 7, 8, 8], OPTIONS).is_err());
        assert!(decode_soundtrack_settings(&[0x10, 0x81, 0], OPTIONS).is_err());
    }

    #[test]
    fn scalar_and_media_adversarial_boundaries() {
        let (empty, empty_report) = decode_soundtrack_settings_with_report(&[], OPTIONS).unwrap();
        assert_eq!(empty.volume(), None);
        assert_eq!(empty.mode_raw(), None);
        assert_eq!(empty_report.max_depth(), 1);
        assert!(decode_soundtrack_settings(&[8, 1], OPTIONS).is_err());
        assert!(
            decode_soundtrack_settings(
                &[9, 0, 0, 0, 0, 0, 0, 0, 0, 9, 0, 0, 0, 0, 0, 0, 0],
                OPTIONS
            )
            .is_err()
        );
        assert!(decode_soundtrack_settings(&[0x1a, 2, 8, 0, 0x1a, 2, 8, 7], OPTIONS).is_err());
        let ordered = [0x1a, 2, 8, 7, 0x1a, 2, 8, 7];
        let mut ids = Vec::new();
        visit_soundtrack_media_identifiers(&ordered, OPTIONS, &mut |id| {
            ids.push(id);
            Ok(())
        })
        .unwrap();
        assert_eq!(ids, [7, 7]);
        let bytes =
            decode_soundtrack_settings(&SOURCE, DecodeOptions::new(14, 8, 64, 2)).unwrap_err();
        assert_eq!(
            bytes.resource_limit(),
            Some(DecodeLimit::Bytes {
                observed: 15,
                maximum: 14
            })
        );
        let depth =
            decode_soundtrack_settings(&SOURCE, DecodeOptions::new(64, 8, 64, 1)).unwrap_err();
        assert_eq!(
            depth.resource_limit(),
            Some(DecodeLimit::Nesting {
                observed: 2,
                maximum: 1
            })
        );
    }

    #[test]
    fn scalar_presence_bits_and_unknown_wire_forms_are_strict() {
        let negative_zero = [9, 0, 0, 0, 0, 0, 0, 0, 0x80];
        assert_eq!(
            decode_soundtrack_settings(&negative_zero, OPTIONS)
                .unwrap()
                .volume()
                .map(f64::to_bits),
            Some((-0.0f64).to_bits())
        );
        let negative_mode = [
            0x10, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 1,
        ];
        assert_eq!(
            decode_soundtrack_settings(&negative_mode, OPTIONS)
                .unwrap()
                .mode_raw(),
            Some(-1)
        );
        assert!(decode_soundtrack_settings(&[0x90, 0, 1], OPTIONS).is_err());
        assert!(decode_soundtrack_settings(&[0x81, 0, 1], OPTIONS).is_err());
        assert!(decode_soundtrack_settings(&[0x25, 0, 0, 0, 0], OPTIONS).is_ok());
    }
}
