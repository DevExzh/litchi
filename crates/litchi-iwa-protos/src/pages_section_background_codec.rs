//! Strict raw-preserving codec for `TP.SectionArchive.background_fill`.
//!
//! Handwritten canonical wire routing owns validation and publication. The
//! private Buffa lazy projection is used only as an independent borrowed
//! semantic parity oracle after that preflight succeeds.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Wire helpers stay beside the raw-preserving rewrite model."
)]

use core::fmt;

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_pages_section_background_generated::LitchiIwaPagesBackgroundProjection as projection;

const SECTION_BACKGROUND_FIELD: u32 = 30;
const FILL_COLOR_FIELD: u32 = 1;
const FILL_GRADIENT_FIELD: u32 = 2;
const FILL_IMAGE_FIELD: u32 = 3;
const COLOR_MODEL_FIELD: u32 = 1;
const COLOR_RED_FIELD: u32 = 3;
const COLOR_GREEN_FIELD: u32 = 4;
const COLOR_BLUE_FIELD: u32 = 5;
const COLOR_ALPHA_FIELD: u32 = 6;
const COLOR_CYAN_FIELD: u32 = 7;
const COLOR_MAGENTA_FIELD: u32 = 8;
const COLOR_YELLOW_FIELD: u32 = 9;
const COLOR_BLACK_FIELD: u32 = 10;
const COLOR_WHITE_FIELD: u32 = 11;
const COLOR_SPACE_FIELD: u32 = 12;
const RGB_MODEL: u64 = 1;
const SRGB_SPACE: u64 = 1;
const P3_SPACE: u64 = 2;
const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;

/// Finite aggregate policy for one strict decode or rewrite/readback cycle.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_input_bytes: usize,
    max_output_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
}

impl DecodeOptions {
    #[must_use]
    pub const fn new(
        max_input_bytes: usize,
        max_output_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
    ) -> Self {
        Self {
            max_input_bytes,
            max_output_bytes,
            max_fields,
            max_work_bytes,
            recursion_limit,
        }
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_input_bytes)
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Supported native RGB color spaces.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RgbSpace {
    Srgb,
    DisplayP3,
}

/// Borrow-free semantic classification of one section background.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum BackgroundSnapshot {
    None,
    Solid {
        red: f32,
        green: f32,
        blue: f32,
        alpha: f32,
        rgb_space: RgbSpace,
    },
    Unsupported,
}

/// Typed mutation accepted by the raw-preserving codec.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum BackgroundWrite {
    Clear,
    Solid {
        red: f32,
        green: f32,
        blue: f32,
        alpha: f32,
        rgb_space: RgbSpace,
    },
}

/// Exact finite consumption of one strict decode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    pub input_bytes: usize,
    pub fields: usize,
    pub work_bytes: usize,
    pub max_depth: u32,
}

/// Exact finite consumption of one rewrite including strict readback.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteReport {
    pub input_bytes: usize,
    pub output_bytes: usize,
    pub fields: usize,
    pub work_bytes: usize,
    pub max_depth: u32,
    pub changed: bool,
    pub allocations: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum LimitKind {
    InputBytes { observed: usize, maximum: usize },
    OutputBytes { observed: usize, maximum: usize },
    Fields { observed: usize, maximum: usize },
    WorkBytes { observed: usize, maximum: usize },
    Nesting { observed: u32, maximum: u32 },
    Allocation { requested: usize },
}

/// Strict wire, resource, or projection failure.
#[derive(Debug)]
pub struct Error {
    kind: ErrorKind,
}

#[derive(Debug)]
enum ErrorKind {
    Invalid,
    Limited(LimitKind),
    Projection(buffa::DecodeError),
    ProjectionMismatch,
}

impl Error {
    #[must_use]
    pub const fn limit(&self) -> Option<LimitKind> {
        match self.kind {
            ErrorKind::Limited(limit) => Some(limit),
            ErrorKind::Invalid | ErrorKind::Projection(_) | ErrorKind::ProjectionMismatch => None,
        }
    }

    const fn invalid() -> Self {
        Self {
            kind: ErrorKind::Invalid,
        }
    }

    const fn limited(limit: LimitKind) -> Self {
        Self {
            kind: ErrorKind::Limited(limit),
        }
    }
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            ErrorKind::Projection(error) => error.fmt(formatter),
            ErrorKind::ProjectionMismatch => formatter.write_str(
                "Pages section-background strict preflight disagrees with Buffa projection",
            ),
            ErrorKind::Invalid | ErrorKind::Limited(_) => {
                formatter.write_str("invalid Pages section-background payload")
            },
        }
    }
}

impl std::error::Error for Error {}

impl From<buffa::DecodeError> for Error {
    fn from(error: buffa::DecodeError) -> Self {
        Self {
            kind: ErrorKind::Projection(error),
        }
    }
}

/// Strictly decode and independently cross-check a complete SectionArchive.
pub fn decode_section_background_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(BackgroundSnapshot, DecodeReport), Error> {
    let mut budget = Budget::new(source.len(), options)?;
    let strict = inspect_section(source, &mut budget)?;
    cross_check_buffa(source, options, strict.snapshot, &mut budget)?;
    Ok((strict.snapshot, budget.decode_report()))
}

/// Rewrite only field 30, then strictly decode and cross-check the candidate.
pub fn rewrite_section_background_with_report(
    source: &[u8],
    write: BackgroundWrite,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), Error> {
    validate_write(write)?;
    let mut budget = Budget::new(source.len(), options)?;
    let strict = inspect_section(source, &mut budget)?;
    cross_check_buffa(source, options, strict.snapshot, &mut budget)?;
    if strict.snapshot == BackgroundSnapshot::Unsupported {
        return Err(Error::invalid());
    }
    let desired = write.snapshot();
    let changed = strict.snapshot != desired;
    if changed {
        let exact = measure_final_output(source, strict, write, &mut budget)?;
        enforce_output(exact, options)?;
    }
    let candidate = if !changed {
        allocate_copy(source, options, &mut budget)?
    } else {
        let replacement = match write {
            BackgroundWrite::Clear => None,
            BackgroundWrite::Solid {
                red,
                green,
                blue,
                alpha,
                rgb_space,
            } => {
                let fill = match strict.fill {
                    Some(fill) if matches!(strict.snapshot, BackgroundSnapshot::Solid { .. }) => {
                        rewrite_existing_solid(
                            fill,
                            red,
                            green,
                            blue,
                            alpha,
                            rgb_space,
                            options,
                            &mut budget,
                        )?
                    },
                    _ => encode_solid_fill(red, green, blue, alpha, rgb_space, &mut budget)?,
                };
                Some(fill)
            },
        };
        rewrite_selected(
            source,
            SECTION_BACKGROUND_FIELD,
            replacement.as_deref(),
            options,
            &mut budget,
        )?
    };
    let output_bytes = candidate.len();
    let readback = inspect_section(&candidate, &mut budget)?;
    cross_check_buffa(&candidate, options, readback.snapshot, &mut budget)?;
    let actual = readback.snapshot;
    if actual != desired {
        return Err(Error {
            kind: ErrorKind::ProjectionMismatch,
        });
    }
    Ok((candidate, budget.rewrite_report(output_bytes, changed)))
}

impl BackgroundWrite {
    const fn snapshot(self) -> BackgroundSnapshot {
        match self {
            Self::Clear => BackgroundSnapshot::None,
            Self::Solid {
                red,
                green,
                blue,
                alpha,
                rgb_space,
            } => BackgroundSnapshot::Solid {
                red,
                green,
                blue,
                alpha,
                rgb_space,
            },
        }
    }
}

fn validate_write(write: BackgroundWrite) -> Result<(), Error> {
    if let BackgroundWrite::Solid {
        red,
        green,
        blue,
        alpha,
        ..
    } = write
    {
        for value in [red, green, blue, alpha] {
            if !value.is_finite() || !(0.0..=1.0).contains(&value) {
                return Err(Error::invalid());
            }
        }
    }
    Ok(())
}

#[derive(Clone, Copy)]
struct SectionInspection<'a> {
    snapshot: BackgroundSnapshot,
    fill: Option<&'a [u8]>,
}

fn inspect_section<'a>(
    source: &'a [u8],
    budget: &mut Budget,
) -> Result<SectionInspection<'a>, Error> {
    let mut cursor = source;
    let mut fill = None;
    while let Some(field) = next_field(&mut cursor, budget, 0)? {
        if field.number == SECTION_BACKGROUND_FIELD {
            if fill.is_some() || field.wire != 2 {
                return Err(Error::invalid());
            }
            let payload = field.bytes()?;
            budget.preauthorize_nested(payload.len(), 1)?;
            fill = Some(payload);
        }
    }
    let snapshot = match fill {
        None => BackgroundSnapshot::None,
        Some(payload) => inspect_fill(payload, budget)?,
    };
    Ok(SectionInspection { snapshot, fill })
}

fn inspect_fill(source: &[u8], budget: &mut Budget) -> Result<BackgroundSnapshot, Error> {
    let mut cursor = source;
    let mut color = None;
    let mut gradient_seen = false;
    let mut image_seen = false;
    while let Some(field) = next_field(&mut cursor, budget, 1)? {
        match field.number {
            FILL_COLOR_FIELD => {
                if color.is_some() || field.wire != 2 {
                    return Err(Error::invalid());
                }
                let payload = field.bytes()?;
                budget.preauthorize_nested(payload.len(), 2)?;
                color = Some(payload);
            },
            FILL_GRADIENT_FIELD => {
                if gradient_seen || field.wire != 2 {
                    return Err(Error::invalid());
                }
                let payload = field.bytes()?;
                budget.preauthorize_nested(payload.len(), 2)?;
                validate_gradient(payload, budget, 2)?;
                gradient_seen = true;
            },
            FILL_IMAGE_FIELD => {
                if image_seen || field.wire != 2 {
                    return Err(Error::invalid());
                }
                let payload = field.bytes()?;
                budget.preauthorize_nested(payload.len(), 2)?;
                validate_image_fill(payload, budget, 2)?;
                image_seen = true;
            },
            _ => {},
        }
    }
    let color_snapshot = match color {
        Some(color) => Some(inspect_color(color, budget, 2)?),
        None => None,
    };
    if gradient_seen || image_seen || color_snapshot.is_none() {
        return Ok(BackgroundSnapshot::Unsupported);
    }
    Ok(color_snapshot.expect("presence checked"))
}

fn inspect_color(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<BackgroundSnapshot, Error> {
    let mut cursor = source;
    let mut model = None;
    let mut red = None;
    let mut green = None;
    let mut blue = None;
    let mut alpha = None;
    let mut rgb_space = None;
    let mut non_rgb_component = false;
    let mut non_rgb_seen = [false; 5];
    let mut non_rgb_values = [None; 5];
    while let Some(field) = next_field(&mut cursor, budget, depth)? {
        match field.number {
            COLOR_MODEL_FIELD => set_once(&mut model, field.canonical_varint()?)?,
            COLOR_RED_FIELD => set_once(&mut red, field.float()?)?,
            COLOR_GREEN_FIELD => set_once(&mut green, field.float()?)?,
            COLOR_BLUE_FIELD => set_once(&mut blue, field.float()?)?,
            COLOR_ALPHA_FIELD => set_once(&mut alpha, field.float()?)?,
            COLOR_SPACE_FIELD => set_once(&mut rgb_space, field.canonical_varint()?)?,
            COLOR_CYAN_FIELD | COLOR_MAGENTA_FIELD | COLOR_YELLOW_FIELD | COLOR_BLACK_FIELD
            | COLOR_WHITE_FIELD => {
                unique(&mut non_rgb_seen, field.number - COLOR_CYAN_FIELD + 1)?;
                non_rgb_values[(field.number - COLOR_CYAN_FIELD) as usize] = Some(field.float()?);
                non_rgb_component = true;
            },
            _ => {},
        }
    }
    let model = model.ok_or_else(Error::invalid)?;
    let alpha = alpha.unwrap_or(1.0);
    for value in [red, green, blue, Some(alpha)]
        .into_iter()
        .flatten()
        .chain(non_rgb_values.into_iter().flatten())
    {
        if !value.is_finite() || !(0.0..=1.0).contains(&value) {
            return Err(Error::invalid());
        }
    }
    if model != RGB_MODEL || non_rgb_component {
        return Ok(BackgroundSnapshot::Unsupported);
    }
    let red = red.ok_or_else(Error::invalid)?;
    let green = green.ok_or_else(Error::invalid)?;
    let blue = blue.ok_or_else(Error::invalid)?;
    let rgb_space = match rgb_space.unwrap_or(SRGB_SPACE) {
        SRGB_SPACE => RgbSpace::Srgb,
        P3_SPACE => RgbSpace::DisplayP3,
        _ => return Ok(BackgroundSnapshot::Unsupported),
    };
    Ok(BackgroundSnapshot::Solid {
        red,
        green,
        blue,
        alpha,
        rgb_space,
    })
}

fn cross_check_buffa(
    source: &[u8],
    options: DecodeOptions,
    strict: BackgroundSnapshot,
    budget: &mut Budget,
) -> Result<(), Error> {
    budget.charge_work(source.len())?;
    let view: projection::PagesSectionBackgroundArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    let fill = view.background_fill.get()?;
    match (strict, fill) {
        (BackgroundSnapshot::None, None) => Ok(()),
        (BackgroundSnapshot::None, Some(_)) | (_, None) => Err(Error {
            kind: ErrorKind::ProjectionMismatch,
        }),
        (BackgroundSnapshot::Unsupported, Some(_)) => Ok(()),
        (
            BackgroundSnapshot::Solid {
                red,
                green,
                blue,
                alpha,
                rgb_space,
            },
            Some(fill),
        ) => {
            let color = fill.color.get()?.ok_or(Error {
                kind: ErrorKind::ProjectionMismatch,
            })?;
            let projected_space = match color.rgb_space.unwrap_or(1) {
                1 => RgbSpace::Srgb,
                2 => RgbSpace::DisplayP3,
                _ => {
                    return Err(Error {
                        kind: ErrorKind::ProjectionMismatch,
                    });
                },
            };
            if color.model != 1
                || color.red.map(f32::to_bits) != Some(red.to_bits())
                || color.green.map(f32::to_bits) != Some(green.to_bits())
                || color.blue.map(f32::to_bits) != Some(blue.to_bits())
                || color.alpha.unwrap_or(1.0).to_bits() != alpha.to_bits()
                || projected_space != rgb_space
                || fill.gradient.is_some()
                || fill.image.is_some()
            {
                return Err(Error {
                    kind: ErrorKind::ProjectionMismatch,
                });
            }
            Ok(())
        },
    }
}

fn validate_gradient(source: &[u8], budget: &mut Budget, depth: u32) -> Result<(), Error> {
    let mut cursor = source;
    let mut seen = [false; 6];
    while let Some(field) = next_field(&mut cursor, budget, depth)? {
        match field.number {
            1 => {
                if unique_varint(&mut seen, 1, field)? > 1 {
                    return Err(Error::invalid());
                }
            },
            3 => unique_wire(&mut seen, 3, field, 5)?,
            4 => {
                if unique_varint(&mut seen, 4, field)? > 1 {
                    return Err(Error::invalid());
                }
            },
            2 => {
                if field.wire != 2 {
                    return Err(Error::invalid());
                }
                let nested = field.bytes()?;
                budget.preauthorize_nested(nested.len(), depth + 1)?;
                validate_gradient_stop(nested, budget, depth + 1)?;
            },
            5 | 6 => {
                unique(&mut seen, field.number)?;
                if field.wire != 2 {
                    return Err(Error::invalid());
                }
                let nested = field.bytes()?;
                budget.preauthorize_nested(nested.len(), depth + 1)?;
                if field.number == 5 {
                    validate_angle_gradient(nested, budget, depth + 1)?;
                } else {
                    validate_transform_gradient(nested, budget, depth + 1)?;
                }
            },
            _ => {},
        }
    }
    Ok(())
}

fn validate_gradient_stop(source: &[u8], budget: &mut Budget, depth: u32) -> Result<(), Error> {
    let mut cursor = source;
    let mut seen = [false; 3];
    while let Some(field) = next_field(&mut cursor, budget, depth)? {
        match field.number {
            1 => {
                unique(&mut seen, 1)?;
                if field.wire != 2 {
                    return Err(Error::invalid());
                }
                let nested = field.bytes()?;
                budget.preauthorize_nested(nested.len(), depth + 1)?;
                let _ = inspect_color(nested, budget, depth + 1)?;
            },
            2 | 3 => unique_wire(&mut seen, field.number, field, 5)?,
            _ => {},
        }
    }
    Ok(())
}

fn validate_image_fill(source: &[u8], budget: &mut Budget, depth: u32) -> Result<(), Error> {
    let mut cursor = source;
    let mut seen = [false; 9];
    while let Some(field) = next_field(&mut cursor, budget, depth)? {
        match field.number {
            2 => {
                if unique_varint(&mut seen, 2, field)? > 4 {
                    return Err(Error::invalid());
                }
            },
            8 => {
                if unique_varint(&mut seen, 8, field)? > 1 {
                    return Err(Error::invalid());
                }
            },
            1 | 3 | 4 | 5 | 6 | 7 | 9 => {
                unique(&mut seen, field.number)?;
                if field.wire != 2 {
                    return Err(Error::invalid());
                }
                let nested = field.bytes()?;
                budget.preauthorize_nested(nested.len(), depth + 1)?;
                match field.number {
                    1 | 5 => validate_reference(nested, budget, depth + 1)?,
                    3 | 9 => {
                        let _ = inspect_color(nested, budget, depth + 1)?;
                    },
                    4 => validate_point_or_size(nested, budget, depth + 1)?,
                    6 | 7 => validate_data_reference(nested, budget, depth + 1)?,
                    _ => unreachable!("selected ImageFill fields are exhaustive"),
                }
            },
            _ => {},
        }
    }
    Ok(())
}

fn validate_angle_gradient(source: &[u8], budget: &mut Budget, depth: u32) -> Result<(), Error> {
    let mut cursor = source;
    let mut seen = false;
    while let Some(field) = next_field(&mut cursor, budget, depth)? {
        if field.number == 2 {
            if core::mem::replace(&mut seen, true) {
                return Err(Error::invalid());
            }
            let _ = field.float()?;
        }
    }
    Ok(())
}

fn validate_transform_gradient(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<(), Error> {
    let mut cursor = source;
    let mut seen = [false; 3];
    while let Some(field) = next_field(&mut cursor, budget, depth)? {
        if matches!(field.number, 1..=3) {
            unique(&mut seen, field.number)?;
            if field.wire != 2 {
                return Err(Error::invalid());
            }
            let nested = field.bytes()?;
            budget.preauthorize_nested(nested.len(), depth + 1)?;
            validate_point_or_size(nested, budget, depth + 1)?;
        }
    }
    Ok(())
}

fn validate_point_or_size(source: &[u8], budget: &mut Budget, depth: u32) -> Result<(), Error> {
    let mut cursor = source;
    let mut seen = [false; 2];
    while let Some(field) = next_field(&mut cursor, budget, depth)? {
        if matches!(field.number, 1 | 2) {
            unique(&mut seen, field.number)?;
            let _ = field.float()?;
        }
    }
    if seen != [true, true] {
        return Err(Error::invalid());
    }
    Ok(())
}

fn validate_reference(source: &[u8], budget: &mut Budget, depth: u32) -> Result<(), Error> {
    let mut cursor = source;
    let mut seen = [false; 3];
    let mut identifier = None;
    while let Some(field) = next_field(&mut cursor, budget, depth)? {
        match field.number {
            1 => identifier = Some(unique_varint(&mut seen, 1, field)?),
            2 => {
                let _ = unique_varint(&mut seen, 2, field)?;
            },
            3 => {
                if unique_varint(&mut seen, 3, field)? > 1 {
                    return Err(Error::invalid());
                }
            },
            _ => {},
        }
    }
    if identifier.is_none() || identifier == Some(0) {
        return Err(Error::invalid());
    }
    Ok(())
}

fn validate_data_reference(source: &[u8], budget: &mut Budget, depth: u32) -> Result<(), Error> {
    let mut cursor = source;
    let mut identifier = None;
    while let Some(field) = next_field(&mut cursor, budget, depth)? {
        if field.number == 1 {
            set_once(&mut identifier, field.canonical_varint()?)?;
        }
    }
    if identifier.is_none() || identifier == Some(0) {
        return Err(Error::invalid());
    }
    Ok(())
}

fn unique(seen: &mut [bool], number: u32) -> Result<(), Error> {
    let slot = seen
        .get_mut(number.saturating_sub(1) as usize)
        .ok_or_else(Error::invalid)?;
    if core::mem::replace(slot, true) {
        return Err(Error::invalid());
    }
    Ok(())
}

fn unique_wire(seen: &mut [bool], number: u32, field: Field<'_>, wire: u8) -> Result<(), Error> {
    unique(seen, number)?;
    if field.wire != wire {
        return Err(Error::invalid());
    }
    if wire == 0 {
        let _ = field.canonical_varint()?;
    }
    Ok(())
}

fn unique_varint(seen: &mut [bool], number: u32, field: Field<'_>) -> Result<u64, Error> {
    unique(seen, number)?;
    field.canonical_varint()
}

fn set_once<T>(slot: &mut Option<T>, value: T) -> Result<(), Error> {
    if slot.replace(value).is_some() {
        return Err(Error::invalid());
    }
    Ok(())
}

fn rewrite_existing_solid(
    fill: &[u8],
    red: f32,
    green: f32,
    blue: f32,
    alpha: f32,
    rgb_space: RgbSpace,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<Vec<u8>, Error> {
    let color = selected_payload(fill, FILL_COLOR_FIELD, budget, 1)?.ok_or_else(Error::invalid)?;
    let rewritten_color =
        rewrite_color(color, red, green, blue, alpha, rgb_space, options, budget)?;
    rewrite_selected(
        fill,
        FILL_COLOR_FIELD,
        Some(&rewritten_color),
        options,
        budget,
    )
}

fn rewrite_color(
    source: &[u8],
    red: f32,
    green: f32,
    blue: f32,
    alpha: f32,
    rgb_space: RgbSpace,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<Vec<u8>, Error> {
    let exact = measure_rewritten_color(source, alpha, rgb_space, budget)?;
    let mut cursor = source;
    let mut output = Vec::new();
    output
        .try_reserve_exact(exact)
        .map_err(|_| Error::limited(LimitKind::Allocation { requested: exact }))?;
    budget.allocations += 1;
    let mut seen_alpha = false;
    let mut seen_space = false;
    while let Some(field) = next_field(&mut cursor, budget, 2)? {
        let replacement = match field.number {
            COLOR_RED_FIELD => Some(red.to_bits()),
            COLOR_GREEN_FIELD => Some(green.to_bits()),
            COLOR_BLUE_FIELD => Some(blue.to_bits()),
            COLOR_ALPHA_FIELD => {
                seen_alpha = true;
                Some(alpha.to_bits())
            },
            COLOR_SPACE_FIELD => {
                seen_space = true;
                append_varint_field(&mut output, COLOR_SPACE_FIELD, space_value(rgb_space));
                None
            },
            _ => {
                output.extend_from_slice(field.raw);
                None
            },
        };
        if let Some(bits) = replacement {
            append_fixed32_field(&mut output, field.number, bits);
        }
    }
    if !seen_alpha && alpha.to_bits() != 1.0f32.to_bits() {
        append_fixed32_field(&mut output, COLOR_ALPHA_FIELD, alpha.to_bits());
    }
    if !seen_space && rgb_space != RgbSpace::Srgb {
        append_varint_field(&mut output, COLOR_SPACE_FIELD, space_value(rgb_space));
    }
    enforce_output(output.len(), options)?;
    Ok(output)
}

fn encode_solid_fill(
    red: f32,
    green: f32,
    blue: f32,
    alpha: f32,
    rgb_space: RgbSpace,
    budget: &mut Budget,
) -> Result<Vec<u8>, Error> {
    const COLOR_SIZE: usize = 24;
    let required = encode_bytes_field_len(FILL_COLOR_FIELD, COLOR_SIZE);
    enforce_output(required, budget.options)?;
    let mut color = Vec::new();
    color.try_reserve_exact(COLOR_SIZE).map_err(|_| {
        Error::limited(LimitKind::Allocation {
            requested: COLOR_SIZE,
        })
    })?;
    budget.allocations += 1;
    append_varint_field(&mut color, COLOR_MODEL_FIELD, RGB_MODEL);
    append_fixed32_field(&mut color, COLOR_RED_FIELD, red.to_bits());
    append_fixed32_field(&mut color, COLOR_GREEN_FIELD, green.to_bits());
    append_fixed32_field(&mut color, COLOR_BLUE_FIELD, blue.to_bits());
    append_fixed32_field(&mut color, COLOR_ALPHA_FIELD, alpha.to_bits());
    append_varint_field(&mut color, COLOR_SPACE_FIELD, space_value(rgb_space));
    let size = encode_bytes_field_len(FILL_COLOR_FIELD, color.len());
    let mut fill = Vec::new();
    fill.try_reserve_exact(size)
        .map_err(|_| Error::limited(LimitKind::Allocation { requested: size }))?;
    budget.allocations += 1;
    append_bytes_field(&mut fill, FILL_COLOR_FIELD, &color);
    Ok(fill)
}

fn measure_final_output(
    source: &[u8],
    strict: SectionInspection<'_>,
    write: BackgroundWrite,
    budget: &mut Budget,
) -> Result<usize, Error> {
    let old = strict.fill.map_or(0, |fill| {
        encode_bytes_field_len(SECTION_BACKGROUND_FIELD, fill.len())
    });
    let new = match write {
        BackgroundWrite::Clear => 0,
        BackgroundWrite::Solid {
            alpha, rgb_space, ..
        } => {
            let fill_len = match strict.fill {
                Some(fill) if matches!(strict.snapshot, BackgroundSnapshot::Solid { .. }) => {
                    let color = selected_payload(fill, FILL_COLOR_FIELD, budget, 1)?
                        .ok_or_else(Error::invalid)?;
                    let color_len = measure_rewritten_color(color, alpha, rgb_space, budget)?;
                    fill.len()
                        .checked_sub(encode_bytes_field_len(FILL_COLOR_FIELD, color.len()))
                        .and_then(|retained| {
                            retained
                                .checked_add(encode_bytes_field_len(FILL_COLOR_FIELD, color_len))
                        })
                        .ok_or_else(Error::invalid)?
                },
                _ => encode_bytes_field_len(FILL_COLOR_FIELD, 24),
            };
            encode_bytes_field_len(SECTION_BACKGROUND_FIELD, fill_len)
        },
    };
    source
        .len()
        .checked_sub(old)
        .and_then(|retained| retained.checked_add(new))
        .ok_or_else(Error::invalid)
}

fn measure_rewritten_color(
    source: &[u8],
    alpha: f32,
    rgb_space: RgbSpace,
    budget: &mut Budget,
) -> Result<usize, Error> {
    let mut cursor = source;
    let mut exact = 0usize;
    let mut alpha_seen = false;
    let mut space_seen = false;
    while let Some(field) = next_field(&mut cursor, budget, 2)? {
        let length = match field.number {
            COLOR_RED_FIELD | COLOR_GREEN_FIELD | COLOR_BLUE_FIELD => {
                encode_fixed32_field_len(field.number)
            },
            COLOR_ALPHA_FIELD => {
                alpha_seen = true;
                encode_fixed32_field_len(field.number)
            },
            COLOR_SPACE_FIELD => {
                space_seen = true;
                encode_varint_field_len(field.number, space_value(rgb_space))
            },
            _ => field.raw.len(),
        };
        exact = exact.checked_add(length).ok_or_else(Error::invalid)?;
    }
    if !alpha_seen && alpha.to_bits() != 1.0f32.to_bits() {
        exact = exact
            .checked_add(encode_fixed32_field_len(COLOR_ALPHA_FIELD))
            .ok_or_else(Error::invalid)?;
    }
    if !space_seen && rgb_space != RgbSpace::Srgb {
        exact = exact
            .checked_add(encode_varint_field_len(
                COLOR_SPACE_FIELD,
                space_value(rgb_space),
            ))
            .ok_or_else(Error::invalid)?;
    }
    Ok(exact)
}

fn encode_fixed32_field_len(number: u32) -> usize {
    encoded_varint_len((u64::from(number) << 3) | 5) + 4
}

fn encode_varint_field_len(number: u32, value: u64) -> usize {
    encoded_varint_len(u64::from(number) << 3) + encoded_varint_len(value)
}

fn space_value(space: RgbSpace) -> u64 {
    match space {
        RgbSpace::Srgb => SRGB_SPACE,
        RgbSpace::DisplayP3 => P3_SPACE,
    }
}

fn selected_payload<'a>(
    source: &'a [u8],
    number: u32,
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<&'a [u8]>, Error> {
    let mut cursor = source;
    let mut selected = None;
    while let Some(field) = next_field(&mut cursor, budget, depth)? {
        if field.number == number {
            if selected.is_some() {
                return Err(Error::invalid());
            }
            selected = Some(field.bytes()?);
        }
    }
    Ok(selected)
}

fn rewrite_selected(
    source: &[u8],
    selected: u32,
    replacement: Option<&[u8]>,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<Vec<u8>, Error> {
    let replacement_len =
        replacement.map_or(0, |payload| encode_bytes_field_len(selected, payload.len()));
    let mut measure = source;
    let mut selected_bytes = 0usize;
    let mut selected_count = 0usize;
    while let Some(field) = next_field(&mut measure, budget, 0)? {
        if field.number == selected {
            selected_count = selected_count.checked_add(1).ok_or_else(Error::invalid)?;
            selected_bytes = selected_bytes
                .checked_add(field.raw.len())
                .ok_or_else(Error::invalid)?;
        }
    }
    if selected_count > 1 {
        return Err(Error::invalid());
    }
    let retained = source
        .len()
        .checked_sub(selected_bytes)
        .ok_or_else(Error::invalid)?;
    let exact = retained.checked_add(replacement_len).ok_or_else(|| {
        Error::limited(LimitKind::OutputBytes {
            observed: usize::MAX,
            maximum: options.max_output_bytes,
        })
    })?;
    enforce_output(exact, options)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(exact)
        .map_err(|_| Error::limited(LimitKind::Allocation { requested: exact }))?;
    budget.allocations += 1;
    let mut cursor = source;
    let mut inserted = false;
    while let Some(field) = next_field(&mut cursor, budget, 0)? {
        if field.number == selected {
            if !inserted {
                if let Some(payload) = replacement {
                    append_bytes_field(&mut output, selected, payload);
                }
                inserted = true;
            }
        } else {
            output.extend_from_slice(field.raw);
        }
    }
    if !inserted {
        if let Some(payload) = replacement {
            append_bytes_field(&mut output, selected, payload);
        }
    }
    enforce_output(output.len(), options)?;
    Ok(output)
}

fn allocate_copy(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<Vec<u8>, Error> {
    enforce_output(source.len(), options)?;
    let mut output = Vec::new();
    output.try_reserve_exact(source.len()).map_err(|_| {
        Error::limited(LimitKind::Allocation {
            requested: source.len(),
        })
    })?;
    budget.allocations += 1;
    output.extend_from_slice(source);
    Ok(output)
}

fn enforce_output(observed: usize, options: DecodeOptions) -> Result<(), Error> {
    if observed > options.max_output_bytes {
        return Err(Error::limited(LimitKind::OutputBytes {
            observed,
            maximum: options.max_output_bytes,
        }));
    }
    Ok(())
}

fn append_varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
    push_varint(output, u64::from(number) << 3);
    push_varint(output, value);
}

fn append_fixed32_field(output: &mut Vec<u8>, number: u32, bits: u32) {
    push_varint(output, (u64::from(number) << 3) | 5);
    output.extend_from_slice(&bits.to_le_bytes());
}

fn append_bytes_field(output: &mut Vec<u8>, number: u32, payload: &[u8]) {
    push_varint(output, (u64::from(number) << 3) | 2);
    push_varint(output, payload.len() as u64);
    output.extend_from_slice(payload);
}

fn encode_bytes_field_len(number: u32, payload_len: usize) -> usize {
    encoded_varint_len((u64::from(number) << 3) | 2)
        + encoded_varint_len(payload_len as u64)
        + payload_len
}

fn push_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

#[derive(Clone, Copy)]
struct Field<'a> {
    number: u32,
    wire: u8,
    value: Value<'a>,
    raw: &'a [u8],
}

#[derive(Clone, Copy)]
enum Value<'a> {
    Varint(u64, usize),
    Fixed64,
    Bytes(&'a [u8]),
    Group,
    Fixed32(u32),
}
enum ParseItem<'a> {
    Field(Field<'a>),
    EndGroup(u32),
}

impl<'a> Field<'a> {
    fn bytes(self) -> Result<&'a [u8], Error> {
        match self.value {
            Value::Bytes(value) if self.wire == 2 => Ok(value),
            _ => Err(Error::invalid()),
        }
    }
    fn canonical_varint(self) -> Result<u64, Error> {
        match self.value {
            Value::Varint(value, encoded)
                if self.wire == 0 && encoded_varint_len(value) == encoded =>
            {
                Ok(value)
            },
            _ => Err(Error::invalid()),
        }
    }
    fn float(self) -> Result<f32, Error> {
        match self.value {
            Value::Fixed32(bits) if self.wire == 5 => Ok(f32::from_bits(bits)),
            _ => Err(Error::invalid()),
        }
    }
}

fn next_field<'a>(
    source: &mut &'a [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<Field<'a>>, Error> {
    match parse_field(source, budget, depth)? {
        Some(ParseItem::Field(field)) => Ok(Some(field)),
        Some(ParseItem::EndGroup(_)) => Err(Error::invalid()),
        None => Ok(None),
    }
}

fn parse_field<'a>(
    source: &mut &'a [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<ParseItem<'a>>, Error> {
    if source.is_empty() {
        return Ok(None);
    }
    budget.depth(depth)?;
    budget.field()?;
    let original = *source;
    let tag = take_canonical_varint(source)?;
    let number = u32::try_from(tag >> 3).map_err(|_| Error::invalid())?;
    let wire = u8::try_from(tag & 7).map_err(|_| Error::invalid())?;
    if number == 0 || number > MAX_FIELD_NUMBER {
        return Err(Error::invalid());
    }
    let value = match wire {
        0 => {
            let (value, encoded) = take_varint_relaxed(source)?;
            Value::Varint(value, encoded)
        },
        1 => {
            let _ = take(source, 8)?;
            Value::Fixed64
        },
        2 => {
            let length =
                usize::try_from(take_canonical_varint(source)?).map_err(|_| Error::invalid())?;
            Value::Bytes(take(source, length)?)
        },
        3 => {
            let child = depth.checked_add(1).ok_or_else(Error::invalid)?;
            skip_group(source, number, budget, child)?;
            Value::Group
        },
        4 => return Ok(Some(ParseItem::EndGroup(number))),
        5 => {
            let bytes = take(source, 4)?;
            Value::Fixed32(u32::from_le_bytes(bytes.try_into().expect("four bytes")))
        },
        _ => return Err(Error::invalid()),
    };
    let consumed = original.len() - source.len();
    budget.charge_work(consumed)?;
    Ok(Some(ParseItem::Field(Field {
        number,
        wire,
        value,
        raw: &original[..consumed],
    })))
}

fn skip_group(
    source: &mut &[u8],
    expected: u32,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), Error> {
    loop {
        match parse_field(source, budget, depth)? {
            Some(ParseItem::Field(_)) => {},
            Some(ParseItem::EndGroup(number)) if number == expected => return Ok(()),
            Some(ParseItem::EndGroup(_)) | None => return Err(Error::invalid()),
        }
    }
}

fn take<'a>(source: &mut &'a [u8], amount: usize) -> Result<&'a [u8], Error> {
    if source.len() < amount {
        return Err(Error::invalid());
    }
    let (selected, rest) = source.split_at(amount);
    *source = rest;
    Ok(selected)
}

fn take_canonical_varint(source: &mut &[u8]) -> Result<u64, Error> {
    let (value, consumed) = take_varint_relaxed(source)?;
    if encoded_varint_len(value) != consumed {
        return Err(Error::invalid());
    }
    Ok(value)
}

fn take_varint_relaxed(source: &mut &[u8]) -> Result<(u64, usize), Error> {
    let original = *source;
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *original.get(index).ok_or_else(Error::invalid)?;
        if index == 9 && byte > 1 {
            return Err(Error::invalid());
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let consumed = index + 1;
            *source = &original[consumed..];
            return Ok((value, consumed));
        }
    }
    Err(Error::invalid())
}

const fn encoded_varint_len(mut value: u64) -> usize {
    let mut length = 1;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

struct Budget {
    options: DecodeOptions,
    input_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    allocations: usize,
}

impl Budget {
    fn new(input_bytes: usize, options: DecodeOptions) -> Result<Self, Error> {
        if input_bytes > options.max_input_bytes {
            return Err(Error::limited(LimitKind::InputBytes {
                observed: input_bytes,
                maximum: options.max_input_bytes,
            }));
        }
        Ok(Self {
            options,
            input_bytes,
            fields: 0,
            work_bytes: 0,
            max_depth: 0,
            allocations: 0,
        })
    }
    fn preauthorize_nested(&mut self, bytes: usize, depth: u32) -> Result<(), Error> {
        self.depth(depth)?;
        self.charge_work(bytes)
    }
    fn depth(&mut self, observed: u32) -> Result<(), Error> {
        self.max_depth = self.max_depth.max(observed);
        if observed > self.options.recursion_limit {
            return Err(Error::limited(LimitKind::Nesting {
                observed,
                maximum: self.options.recursion_limit,
            }));
        }
        Ok(())
    }
    fn field(&mut self) -> Result<(), Error> {
        self.fields = self.fields.checked_add(1).ok_or_else(Error::invalid)?;
        if self.fields > self.options.max_fields {
            return Err(Error::limited(LimitKind::Fields {
                observed: self.fields,
                maximum: self.options.max_fields,
            }));
        }
        Ok(())
    }
    fn charge_work(&mut self, amount: usize) -> Result<(), Error> {
        self.work_bytes = self
            .work_bytes
            .checked_add(amount)
            .ok_or_else(Error::invalid)?;
        if self.work_bytes > self.options.max_work_bytes {
            return Err(Error::limited(LimitKind::WorkBytes {
                observed: self.work_bytes,
                maximum: self.options.max_work_bytes,
            }));
        }
        Ok(())
    }
    const fn decode_report(&self) -> DecodeReport {
        DecodeReport {
            input_bytes: self.input_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
        }
    }
    const fn rewrite_report(&self, output_bytes: usize, changed: bool) -> RewriteReport {
        RewriteReport {
            input_bytes: self.input_bytes,
            output_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            changed,
            allocations: self.allocations,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn encode_varint_field(number: u32, value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        append_varint_field(&mut output, number, value);
        output
    }

    fn encode_fixed32_field(number: u32, bits: u32) -> Vec<u8> {
        let mut output = Vec::new();
        append_fixed32_field(&mut output, number, bits);
        output
    }

    fn encode_bytes_field(number: u32, payload: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        append_bytes_field(&mut output, number, payload);
        output
    }

    fn solid_fill(red: f32, green: f32, blue: f32, alpha: f32, space: RgbSpace) -> Vec<u8> {
        let mut budget = Budget::new(0, options()).unwrap();
        encode_solid_fill(red, green, blue, alpha, space, &mut budget).unwrap()
    }

    fn options() -> DecodeOptions {
        DecodeOptions::new(1 << 20, 1 << 20, 1 << 16, 1 << 24, 16)
    }

    fn solid(space: RgbSpace) -> BackgroundWrite {
        BackgroundWrite::Solid {
            red: 0.25,
            green: 0.5,
            blue: 0.75,
            alpha: 1.0,
            rgb_space: space,
        }
    }

    fn section_with_fill(fill: &[u8]) -> Vec<u8> {
        let mut source = encode_varint_field(17, 1);
        source.extend_from_slice(&encode_bytes_field(SECTION_BACKGROUND_FIELD, fill));
        source.extend_from_slice(&encode_varint_field(31, 7));
        source
    }

    #[test]
    fn absent_create_decode_clear_and_exact_noop() {
        let source = encode_varint_field(17, 1);
        let (none, _) = decode_section_background_with_report(&source, options()).unwrap();
        assert_eq!(none, BackgroundSnapshot::None);

        let (created, report) =
            rewrite_section_background_with_report(&source, solid(RgbSpace::Srgb), options())
                .unwrap();
        assert!(report.changed);
        let (snapshot, _) = decode_section_background_with_report(&created, options()).unwrap();
        assert_eq!(snapshot, solid(RgbSpace::Srgb).snapshot());

        let (noop, report) =
            rewrite_section_background_with_report(&created, solid(RgbSpace::Srgb), options())
                .unwrap();
        assert!(!report.changed);
        assert_eq!(noop, created);

        let (cleared, _) =
            rewrite_section_background_with_report(&created, BackgroundWrite::Clear, options())
                .unwrap();
        assert_eq!(cleared, source);
    }

    #[test]
    fn p3_roundtrips_exact_float_bits() {
        let source = Vec::new();
        let (candidate, _) =
            rewrite_section_background_with_report(&source, solid(RgbSpace::DisplayP3), options())
                .unwrap();
        let (snapshot, _) = decode_section_background_with_report(&candidate, options()).unwrap();
        assert_eq!(snapshot, solid(RgbSpace::DisplayP3).snapshot());
    }

    #[test]
    fn solid_patch_preserves_unknown_fill_color_and_section_spans() {
        let mut color = encode_varint_field(COLOR_MODEL_FIELD, RGB_MODEL);
        color.extend_from_slice(&encode_fixed32_field(COLOR_RED_FIELD, 0.1f32.to_bits()));
        color.extend_from_slice(&encode_varint_field(99, 123));
        color.extend_from_slice(&encode_fixed32_field(COLOR_GREEN_FIELD, 0.5f32.to_bits()));
        color.extend_from_slice(&encode_fixed32_field(COLOR_BLUE_FIELD, 0.75f32.to_bits()));
        let color_unknown = encode_varint_field(99, 123);
        let mut fill = encode_varint_field(100, 456);
        fill.extend_from_slice(&encode_bytes_field(FILL_COLOR_FIELD, &color));
        let fill_unknown = encode_varint_field(100, 456);
        let source = section_with_fill(&fill);
        let section_unknown = encode_varint_field(31, 7);

        let (candidate, _) =
            rewrite_section_background_with_report(&source, solid(RgbSpace::Srgb), options())
                .unwrap();
        assert!(
            candidate
                .windows(color_unknown.len())
                .any(|span| span == color_unknown)
        );
        assert!(
            candidate
                .windows(fill_unknown.len())
                .any(|span| span == fill_unknown)
        );
        assert!(
            candidate
                .windows(section_unknown.len())
                .any(|span| span == section_unknown)
        );
        let (snapshot, _) = decode_section_background_with_report(&candidate, options()).unwrap();
        assert_eq!(snapshot, solid(RgbSpace::Srgb).snapshot());
    }

    #[test]
    fn valid_gradient_is_unsupported_but_malformed_nested_wire_is_rejected() {
        let gradient = encode_varint_field(1, 0);
        let fill = encode_bytes_field(FILL_GRADIENT_FIELD, &gradient);
        let source = section_with_fill(&fill);
        let (snapshot, _) = decode_section_background_with_report(&source, options()).unwrap();
        assert_eq!(snapshot, BackgroundSnapshot::Unsupported);
        assert!(
            rewrite_section_background_with_report(&source, BackgroundWrite::Clear, options(),)
                .is_err()
        );
        assert!(
            rewrite_section_background_with_report(&source, solid(RgbSpace::Srgb), options(),)
                .is_err()
        );

        let malformed_gradient = [0x12, 0x02, 0x08];
        let fill = encode_bytes_field(FILL_GRADIENT_FIELD, &malformed_gradient);
        assert!(
            decode_section_background_with_report(&section_with_fill(&fill), options()).is_err()
        );
    }

    #[test]
    fn strict_selected_shape_and_channel_validation_fail_closed() {
        let fill = solid_fill(0.1, 0.2, 0.3, 1.0, RgbSpace::Srgb);
        let mut duplicate = section_with_fill(&fill);
        duplicate.extend_from_slice(&encode_bytes_field(SECTION_BACKGROUND_FIELD, &fill));
        assert!(decode_section_background_with_report(&duplicate, options()).is_err());

        let wrong_wire = encode_varint_field(SECTION_BACKGROUND_FIELD, 1);
        assert!(decode_section_background_with_report(&wrong_wire, options()).is_err());

        let nan = solid_fill(f32::NAN, 0.2, 0.3, 1.0, RgbSpace::Srgb);
        assert!(
            decode_section_background_with_report(&section_with_fill(&nan), options()).is_err()
        );
        assert!(
            rewrite_section_background_with_report(
                &[],
                BackgroundWrite::Solid {
                    red: 1.1,
                    green: 0.0,
                    blue: 0.0,
                    alpha: 1.0,
                    rgb_space: RgbSpace::Srgb,
                },
                options(),
            )
            .is_err()
        );
    }

    #[test]
    fn unsupported_nested_known_fields_remain_schema_strict() {
        let fill = encode_bytes_field(FILL_GRADIENT_FIELD, &encode_varint_field(4, 2));
        assert!(
            decode_section_background_with_report(&section_with_fill(&fill), options()).is_err()
        );

        let fill = encode_bytes_field(FILL_IMAGE_FIELD, &encode_varint_field(8, 2));
        assert!(
            decode_section_background_with_report(&section_with_fill(&fill), options()).is_err()
        );

        let fill = encode_bytes_field(FILL_IMAGE_FIELD, &encode_bytes_field(1, &[]));
        assert!(
            decode_section_background_with_report(&section_with_fill(&fill), options()).is_err()
        );

        let malformed_color = encode_varint_field(COLOR_MODEL_FIELD, RGB_MODEL);
        let mut mixed = encode_bytes_field(FILL_COLOR_FIELD, &malformed_color);
        mixed.extend_from_slice(&encode_bytes_field(
            FILL_GRADIENT_FIELD,
            &encode_varint_field(1, 0),
        ));
        assert!(
            decode_section_background_with_report(&section_with_fill(&mixed), options()).is_err()
        );

        let mut duplicate_cmyk = encode_varint_field(COLOR_MODEL_FIELD, 2);
        duplicate_cmyk.extend_from_slice(&encode_fixed32_field(COLOR_CYAN_FIELD, 0.1f32.to_bits()));
        duplicate_cmyk.extend_from_slice(&encode_fixed32_field(COLOR_CYAN_FIELD, 0.2f32.to_bits()));
        let fill = encode_bytes_field(FILL_COLOR_FIELD, &duplicate_cmyk);
        assert!(
            decode_section_background_with_report(&section_with_fill(&fill), options()).is_err()
        );
    }

    #[test]
    fn output_and_work_limits_refuse_before_candidate_escape() {
        let source = encode_varint_field(17, 1);
        let (candidate, _) =
            rewrite_section_background_with_report(&source, solid(RgbSpace::Srgb), options())
                .unwrap();
        let exact = DecodeOptions::new(1 << 20, candidate.len(), 1 << 16, 1 << 24, 16);
        assert!(
            rewrite_section_background_with_report(&source, solid(RgbSpace::Srgb), exact).is_ok()
        );
        let short = DecodeOptions::new(1 << 20, candidate.len() - 1, 1 << 16, 1 << 24, 16);
        assert!(matches!(
            rewrite_section_background_with_report(&source, solid(RgbSpace::Srgb), short)
                .unwrap_err()
                .limit(),
            Some(LimitKind::OutputBytes { .. })
        ));

        let (_, report) = decode_section_background_with_report(&candidate, options()).unwrap();
        let short_work = DecodeOptions::new(1 << 20, 1 << 20, 1 << 16, report.work_bytes - 1, 16);
        assert!(matches!(
            decode_section_background_with_report(&candidate, short_work)
                .unwrap_err()
                .limit(),
            Some(LimitKind::WorkBytes { .. })
        ));
    }

    #[test]
    fn source_projection_never_uses_prost_or_generated_encoding() {
        let source = include_str!("pages_section_background_codec.rs")
            .split_once("#[cfg(test)]")
            .unwrap()
            .0;
        assert!(!source.contains("prost::"));
        assert!(!source.contains("to_owned_message"));
        assert!(!source.contains("encode_to_vec"));
        assert!(!source.contains("try_encode"));
        assert!(!source.contains(".encode("));
    }
}
