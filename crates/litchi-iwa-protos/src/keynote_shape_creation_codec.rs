//! Bounded Buffa authoring for a fresh Keynote rectangle text-box graph.
//!
//! The public values are a deliberately small, typed closure of the native
//! `ShapeInfoArchive` and `StorageArchive` records. Callers establish their
//! package-wide operation budget before entering this module; this codec still
//! performs a complete local preflight before creating nested Buffa views or
//! output buffers.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The typed shape closure is organized around its wire graph."
)]

use std::fmt;

use buffa::ViewEncode as _;

use crate::buffa_keynote_shape_creation_generated::LitchiIwaProjection as projection;

const DEFAULT_MAX_OUTPUT_BYTES: usize = 64 * 1024;
const DEFAULT_MAX_WORK_BYTES: usize = 256 * 1024;
const DEFAULT_MAX_FIELDS: usize = 256;
const DEFAULT_MAX_REFERENCES: usize = 16;
const DEFAULT_MAX_ALLOCATIONS: usize = 64;
const TEXT_BOX_STORAGE_KIND: i32 = 3;
const RECTANGLE_MOVE_TO: i32 = 1;
const RECTANGLE_LINE_TO: i32 = 2;
const RECTANGLE_CLOSE_SUBPATH: i32 = 5;

/// A finite point used by the source-free rectangle path.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Point {
    x: f32,
    y: f32,
}

impl Point {
    /// Construct a point. Finite-value validation occurs at encoding.
    #[must_use]
    pub const fn new(x: f32, y: f32) -> Self {
        Self { x, y }
    }

    /// Return the horizontal coordinate.
    #[must_use]
    pub const fn x(self) -> f32 {
        self.x
    }

    /// Return the vertical coordinate.
    #[must_use]
    pub const fn y(self) -> f32 {
        self.y
    }
}

/// A positive finite rectangle size.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Size {
    width: f32,
    height: f32,
}

impl Size {
    /// Construct a size. Positivity validation occurs at encoding.
    #[must_use]
    pub const fn new(width: f32, height: f32) -> Self {
        Self { width, height }
    }

    /// Return the width.
    #[must_use]
    pub const fn width(self) -> f32 {
        self.width
    }

    /// Return the height.
    #[must_use]
    pub const fn height(self) -> f32 {
        self.height
    }
}

/// Typed input for a source-free native text-box shape-info payload.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ShapeInfoWrite {
    parent_identifier: u64,
    style_identifier: u64,
    title_identifier: u64,
    caption_identifier: u64,
    storage_identifier: u64,
    position: Point,
    size: Size,
}

impl ShapeInfoWrite {
    /// Construct a canonical rectangle-backed text-box payload.
    #[must_use]
    pub const fn new(
        parent_identifier: u64,
        style_identifier: u64,
        title_identifier: u64,
        caption_identifier: u64,
        storage_identifier: u64,
        position: Point,
        size: Size,
    ) -> Self {
        Self {
            parent_identifier,
            style_identifier,
            title_identifier,
            caption_identifier,
            storage_identifier,
            position,
            size,
        }
    }
}

/// Typed input for a source-free native text storage payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StorageWrite<'text> {
    style_sheet_identifier: u64,
    text: &'text str,
}

impl<'text> StorageWrite<'text> {
    /// Construct a text-box storage payload using an existing stylesheet.
    #[must_use]
    pub const fn new(style_sheet_identifier: u64, text: &'text str) -> Self {
        Self {
            style_sheet_identifier,
            text,
        }
    }
}

/// Finite resource policy for one shape or storage encoding.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EncodeOptions {
    max_output_bytes: usize,
    max_references: usize,
    max_fields: usize,
    max_work_bytes: usize,
    max_allocations: usize,
}

impl EncodeOptions {
    /// Construct an explicit finite output/reference/field/work policy.
    #[must_use]
    pub const fn new(
        max_output_bytes: usize,
        max_references: usize,
        max_fields: usize,
        max_work_bytes: usize,
        max_allocations: usize,
    ) -> Self {
        Self {
            max_output_bytes,
            max_references,
            max_fields,
            max_work_bytes,
            max_allocations,
        }
    }

    /// Build the conservative finite default policy.
    #[must_use]
    pub const fn default() -> Self {
        Self::new(
            DEFAULT_MAX_OUTPUT_BYTES,
            DEFAULT_MAX_REFERENCES,
            DEFAULT_MAX_FIELDS,
            DEFAULT_MAX_WORK_BYTES,
            DEFAULT_MAX_ALLOCATIONS,
        )
    }

    /// Replace the output-byte ceiling.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }

    /// Replace the reference ceiling.
    #[must_use]
    pub const fn with_max_references(mut self, maximum: usize) -> Self {
        self.max_references = maximum;
        self
    }

    /// Replace the field ceiling.
    #[must_use]
    pub const fn with_max_fields(mut self, maximum: usize) -> Self {
        self.max_fields = maximum;
        self
    }

    /// Replace the aggregate work ceiling.
    #[must_use]
    pub const fn with_max_work_bytes(mut self, maximum: usize) -> Self {
        self.max_work_bytes = maximum;
        self
    }

    /// Replace the allocation-count ceiling.
    #[must_use]
    pub const fn with_max_allocations(mut self, maximum: usize) -> Self {
        self.max_allocations = maximum;
        self
    }
}

/// Exact finite resource consumption of one shape/storage encoding.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EncodeReport {
    output_bytes: usize,
    references: usize,
    fields: usize,
    work_bytes: usize,
    allocations: usize,
}

impl EncodeReport {
    /// Return the exact payload length.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    /// Return the number of native object references emitted in the payload.
    #[must_use]
    pub const fn references(self) -> usize {
        self.references
    }

    /// Return the number of protobuf fields emitted, including nested fields.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Return the aggregate measurement and encoding work.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Return the exact number of top-level/nested allocations charged.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }
}

/// Encoded payload plus exact resource evidence.
#[derive(Debug, PartialEq, Eq)]
pub struct EncodeOutput {
    bytes: Vec<u8>,
    report: EncodeReport,
}

impl EncodeOutput {
    /// Borrow the encoded payload.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Return the exact resource report.
    #[must_use]
    pub const fn report(&self) -> EncodeReport {
        self.report
    }

    /// Consume the output and return its bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }
}

/// Invalid input rejected before any nested view/output allocation.
#[derive(Debug, Clone, Copy, PartialEq)]
#[non_exhaustive]
pub enum InvalidInput {
    /// A required object identifier was zero.
    ZeroIdentifier(&'static str),
    /// A scalar was not finite.
    NonFinite(&'static str),
    /// A rectangle extent was not positive.
    NonPositiveSize(&'static str),
}

/// Finite resource exceeded by one encoding.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum EncodeLimit {
    /// Output bytes exceed the caller's ceiling.
    OutputBytes { observed: usize, maximum: usize },
    /// References exceed the caller's ceiling.
    References { observed: usize, maximum: usize },
    /// Fields exceed the caller's ceiling.
    Fields { observed: usize, maximum: usize },
    /// Work exceeds the caller's ceiling.
    WorkBytes { observed: usize, maximum: usize },
    /// Allocations exceed the caller's ceiling.
    Allocations { observed: usize, maximum: usize },
}

/// Failure from a typed, bounded shape/storage encoding.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum EncodeError {
    /// Input failed strict scalar/reference validation.
    InvalidInput(InvalidInput),
    /// A finite caller budget was exceeded.
    Resource(EncodeLimit),
    /// The output vector could not be reserved exactly.
    Allocation { amount: usize },
    /// Buffa rejected the generated view or bound.
    Buffa(buffa::EncodeError),
    /// The static and generated measurements disagreed.
    Verification,
}

impl fmt::Display for EncodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidInput(input) => {
                write!(formatter, "invalid Keynote shape input: {input:?}")
            },
            Self::Resource(_) => formatter.write_str("Keynote shape encoding limit exceeded"),
            Self::Allocation { amount } => {
                write!(
                    formatter,
                    "Keynote shape output allocation failed ({amount} bytes)"
                )
            },
            Self::Buffa(error) => error.fmt(formatter),
            Self::Verification => formatter.write_str("Keynote shape encoding verification failed"),
        }
    }
}

impl std::error::Error for EncodeError {}

impl From<buffa::EncodeError> for EncodeError {
    fn from(error: buffa::EncodeError) -> Self {
        Self::Buffa(error)
    }
}

/// Encode a fresh rectangle-backed `TSWP.ShapeInfoArchive` payload.
pub fn measure_shape_info(write: &ShapeInfoWrite) -> Result<EncodeReport, EncodeError> {
    validate_shape(write)?;
    let output_bytes = shape_encoded_len(write).ok_or(EncodeError::Verification)?;
    report(output_bytes, 6, 71, 28)
}

/// Encode a fresh rectangle-backed `TSWP.ShapeInfoArchive` payload.
pub fn encode_shape_info(
    write: &ShapeInfoWrite,
    options: EncodeOptions,
) -> Result<EncodeOutput, EncodeError> {
    let report = measure_shape_info(write)?;
    let output_bytes = report.output_bytes;
    preflight(report, options)?;

    let view = shape_view(write);
    let measured =
        usize::try_from(view.try_encoded_len()?).map_err(|_| EncodeError::Verification)?;
    if measured != output_bytes {
        return Err(EncodeError::Verification);
    }
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(output_bytes)
        .map_err(|_| EncodeError::Allocation {
            amount: output_bytes,
        })?;
    if bytes.capacity() < output_bytes {
        return Err(EncodeError::Allocation {
            amount: output_bytes,
        });
    }
    let encoded = view.try_encode_bounded(
        u32::try_from(options.max_output_bytes).unwrap_or(u32::MAX),
        &mut bytes,
    )?;
    if usize::try_from(encoded).ok() != Some(output_bytes) || bytes.len() != output_bytes {
        return Err(EncodeError::Verification);
    }
    Ok(EncodeOutput { bytes, report })
}

/// Encode a fresh text-box `TSWP.StorageArchive` payload.
pub fn measure_storage(write: &StorageWrite<'_>) -> Result<EncodeReport, EncodeError> {
    validate_storage(write)?;
    let output_bytes = storage_encoded_len(write).ok_or(EncodeError::Verification)?;
    let fields = 5usize;
    let work_bytes = output_bytes
        .checked_mul(3)
        .and_then(|work| work.checked_add(fields))
        .ok_or(EncodeError::Verification)?;
    Ok(EncodeReport {
        output_bytes,
        references: 1,
        fields,
        work_bytes,
        allocations: 3,
    })
}

/// Encode a fresh text-box `TSWP.StorageArchive` payload.
pub fn encode_storage(
    write: &StorageWrite<'_>,
    options: EncodeOptions,
) -> Result<EncodeOutput, EncodeError> {
    let report = measure_storage(write)?;
    let output_bytes = report.output_bytes;
    preflight(report, options)?;
    let view = storage_view(write);
    let measured =
        usize::try_from(view.try_encoded_len()?).map_err(|_| EncodeError::Verification)?;
    if measured != output_bytes {
        return Err(EncodeError::Verification);
    }
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(output_bytes)
        .map_err(|_| EncodeError::Allocation {
            amount: output_bytes,
        })?;
    if bytes.capacity() < output_bytes {
        return Err(EncodeError::Allocation {
            amount: output_bytes,
        });
    }
    let encoded = view.try_encode_bounded(
        u32::try_from(options.max_output_bytes).unwrap_or(u32::MAX),
        &mut bytes,
    )?;
    if usize::try_from(encoded).ok() != Some(output_bytes) || bytes.len() != output_bytes {
        return Err(EncodeError::Verification);
    }
    Ok(EncodeOutput { bytes, report })
}

fn report(
    output_bytes: usize,
    references: usize,
    fields: usize,
    allocations: usize,
) -> Result<EncodeReport, EncodeError> {
    let work_bytes = output_bytes
        .checked_mul(3)
        .and_then(|work| work.checked_add(fields))
        .ok_or(EncodeError::Verification)?;
    Ok(EncodeReport {
        output_bytes,
        references,
        fields,
        work_bytes,
        allocations,
    })
}

fn preflight(report: EncodeReport, options: EncodeOptions) -> Result<(), EncodeError> {
    let limit = [
        (report.output_bytes > options.max_output_bytes).then_some(EncodeLimit::OutputBytes {
            observed: report.output_bytes,
            maximum: options.max_output_bytes,
        }),
        (report.references > options.max_references).then_some(EncodeLimit::References {
            observed: report.references,
            maximum: options.max_references,
        }),
        (report.fields > options.max_fields).then_some(EncodeLimit::Fields {
            observed: report.fields,
            maximum: options.max_fields,
        }),
        (report.work_bytes > options.max_work_bytes).then_some(EncodeLimit::WorkBytes {
            observed: report.work_bytes,
            maximum: options.max_work_bytes,
        }),
        (report.allocations > options.max_allocations).then_some(EncodeLimit::Allocations {
            observed: report.allocations,
            maximum: options.max_allocations,
        }),
    ]
    .into_iter()
    .flatten()
    .next();
    if let Some(limit) = limit {
        return Err(EncodeError::Resource(limit));
    }
    Ok(())
}

fn validate_identifier(identifier: u64, name: &'static str) -> Result<(), EncodeError> {
    if identifier == 0 {
        return Err(EncodeError::InvalidInput(InvalidInput::ZeroIdentifier(
            name,
        )));
    }
    Ok(())
}

fn validate_point(point: Point, prefix: &'static str) -> Result<(), EncodeError> {
    for value in [point.x, point.y] {
        if !value.is_finite() {
            return Err(EncodeError::InvalidInput(InvalidInput::NonFinite(prefix)));
        }
    }
    Ok(())
}

fn validate_size(size: Size, prefix: &'static str) -> Result<(), EncodeError> {
    for value in [size.width, size.height] {
        if !value.is_finite() {
            return Err(EncodeError::InvalidInput(InvalidInput::NonFinite(prefix)));
        }
        if value <= 0.0 {
            return Err(EncodeError::InvalidInput(InvalidInput::NonPositiveSize(
                prefix,
            )));
        }
    }
    Ok(())
}

fn validate_shape(write: &ShapeInfoWrite) -> Result<(), EncodeError> {
    for (identifier, name) in [
        (write.parent_identifier, "parent_identifier"),
        (write.style_identifier, "style_identifier"),
        (write.title_identifier, "title_identifier"),
        (write.caption_identifier, "caption_identifier"),
        (write.storage_identifier, "storage_identifier"),
    ] {
        validate_identifier(identifier, name)?;
    }
    validate_point(write.position, "position")?;
    validate_size(write.size, "size")?;
    Ok(())
}

fn validate_storage(write: &StorageWrite<'_>) -> Result<(), EncodeError> {
    validate_identifier(write.style_sheet_identifier, "style_sheet_identifier")?;
    if write.text.contains('\u{fffc}') {
        return Err(EncodeError::InvalidInput(InvalidInput::NonFinite("text")));
    }
    Ok(())
}

fn reference(identifier: u64) -> projection::ReferenceView<'static> {
    projection::ReferenceView {
        identifier,
        ..Default::default()
    }
}

fn point(value: Point) -> projection::PointView<'static> {
    projection::PointView {
        x: value.x,
        y: value.y,
        ..Default::default()
    }
}

fn size(value: Size) -> projection::SizeView<'static> {
    projection::SizeView {
        width: value.width,
        height: value.height,
        ..Default::default()
    }
}

fn rectangle_path(value: Size) -> projection::PathView<'static> {
    let points = [
        (RECTANGLE_MOVE_TO, Some(Point::new(0.0, 0.0))),
        (RECTANGLE_LINE_TO, Some(Point::new(value.width, 0.0))),
        (
            RECTANGLE_LINE_TO,
            Some(Point::new(value.width, value.height)),
        ),
        (RECTANGLE_LINE_TO, Some(Point::new(0.0, value.height))),
        (RECTANGLE_CLOSE_SUBPATH, None),
        (RECTANGLE_MOVE_TO, Some(Point::new(0.0, 0.0))),
    ];
    let elements = points
        .into_iter()
        .map(|(element_type, point_value)| {
            let points = point_value
                .map(|point_value| buffa::RepeatedView::new(vec![point(point_value)]))
                .unwrap_or_else(|| buffa::RepeatedView::new(Vec::new()));
            projection::PathElementView {
                r#type: element_type,
                points,
                __buffa_required_seen_0: 1,
            }
        })
        .collect::<Vec<_>>();
    projection::PathView {
        elements: buffa::RepeatedView::new(elements),
    }
}

fn shape_view(write: &ShapeInfoWrite) -> projection::ShapeInfoArchiveView<'static> {
    let geometry = projection::GeometryArchiveView {
        position: buffa::MessageFieldView::set(point(write.position)),
        size: buffa::MessageFieldView::set(size(write.size)),
        flags: Some(3),
        angle: Some(0.0),
    };
    let wrap = projection::ExteriorTextWrapArchiveView {
        r#type: Some(4),
        direction: Some(2),
        fit_type: Some(1),
        margin: Some(12.0),
        alpha_threshold: Some(0.5),
        is_html_wrap: Some(false),
        __buffa_phantom: std::marker::PhantomData,
    };
    let drawable = projection::DrawableArchiveView {
        geometry: buffa::MessageFieldView::set(geometry),
        parent: buffa::MessageFieldView::set(reference(write.parent_identifier)),
        exterior_text_wrap: buffa::MessageFieldView::set(wrap),
        locked: Some(false),
        aspect_ratio_locked: Some(false),
        title: buffa::MessageFieldView::set(reference(write.title_identifier)),
        caption: buffa::MessageFieldView::set(reference(write.caption_identifier)),
        title_hidden: Some(false),
        caption_hidden: Some(false),
    };
    let bezier = projection::BezierPathSourceArchiveView {
        natural_size: buffa::MessageFieldView::set(size(write.size)),
        path: buffa::MessageFieldView::set(rectangle_path(write.size)),
    };
    let pathsource = projection::PathSourceArchiveView {
        horizontal_flip: Some(false),
        vertical_flip: Some(false),
        bezier_path_source: buffa::MessageFieldView::set(bezier),
    };
    let shape = projection::ShapeArchiveView {
        super_: buffa::MessageFieldView::set(drawable),
        style: buffa::MessageFieldView::set(reference(write.style_identifier)),
        pathsource: buffa::MessageFieldView::set(pathsource),
        stroke_pattern_offset_distance: Some(0.0),
    };
    projection::ShapeInfoArchiveView {
        super_: buffa::MessageFieldView::set(shape),
        deprecated_storage: buffa::MessageFieldView::set(reference(write.storage_identifier)),
        text_flow: buffa::MessageFieldView::unset(),
        owned_storage: buffa::MessageFieldView::set(reference(write.storage_identifier)),
        is_text_box: Some(true),
    }
}

fn storage_view<'a>(write: &'a StorageWrite<'a>) -> projection::StorageArchiveView<'a> {
    projection::StorageArchiveView {
        kind: Some(TEXT_BOX_STORAGE_KIND),
        style_sheet: buffa::MessageFieldView::set(reference(write.style_sheet_identifier)),
        text: buffa::RepeatedView::new(vec![write.text]),
        in_document: Some(true),
    }
}

fn shape_encoded_len(write: &ShapeInfoWrite) -> Option<usize> {
    let point = fixed32_field(1).checked_add(fixed32_field(2))?;
    let size = point;
    let geometry = len_field(1, point)?
        .checked_add(len_field(2, size)?)?
        .checked_add(varint_field(3, 3))?
        .checked_add(fixed32_field(4))?;
    let wrap = varint_field(1, 4)
        .checked_add(varint_field(2, 2))?
        .checked_add(varint_field(3, 1))?
        .checked_add(fixed32_field(4))?
        .checked_add(fixed32_field(5))?
        .checked_add(bool_field(6))?;
    let drawable = len_field(1, geometry)?
        .checked_add(len_field(2, reference_len(write.parent_identifier))?)?
        .checked_add(len_field(3, wrap)?)?
        .checked_add(bool_field(5))?
        .checked_add(bool_field(7))?
        .checked_add(len_field(10, reference_len(write.title_identifier))?)?
        .checked_add(len_field(11, reference_len(write.caption_identifier))?)?
        .checked_add(bool_field(12))?
        .checked_add(bool_field(13))?;
    let path_element = |element_type: i32, has_point: bool| {
        let value = varint_field(1, i64::from(element_type) as u64);
        if has_point {
            value.checked_add(len_field(2, point)?)
        } else {
            Some(value)
        }
    };
    let path = [
        path_element(RECTANGLE_MOVE_TO, true)?,
        path_element(RECTANGLE_LINE_TO, true)?,
        path_element(RECTANGLE_LINE_TO, true)?,
        path_element(RECTANGLE_LINE_TO, true)?,
        path_element(RECTANGLE_CLOSE_SUBPATH, false)?,
        path_element(RECTANGLE_MOVE_TO, true)?,
    ]
    .into_iter()
    .try_fold(0usize, |length, element| {
        length.checked_add(len_field(1, element)?)
    })?;
    let bezier = len_field(2, size)?.checked_add(len_field(3, path)?)?;
    let pathsource = bool_field(1)
        .checked_add(bool_field(2))?
        .checked_add(len_field(5, bezier)?)?;
    let shape = len_field(1, drawable)?
        .checked_add(len_field(2, reference_len(write.style_identifier))?)?
        .checked_add(len_field(3, pathsource)?)?
        .checked_add(fixed32_field(6))?;
    len_field(1, shape)?
        .checked_add(len_field(2, reference_len(write.storage_identifier))?)?
        .checked_add(len_field(4, reference_len(write.storage_identifier))?)?
        .checked_add(bool_field(6))
}

fn storage_encoded_len(write: &StorageWrite<'_>) -> Option<usize> {
    varint_field(1, TEXT_BOX_STORAGE_KIND as u64)
        .checked_add(len_field(2, reference_len(write.style_sheet_identifier))?)?
        .checked_add(len_field(3, write.text.len())?)?
        .checked_add(bool_field(10))
}

fn key_len(field: u32, wire_type: u8) -> usize {
    varint_size((u64::from(field) << 3) | u64::from(wire_type))
}

fn varint_size(mut value: u64) -> usize {
    let mut size = 1usize;
    while value >= 0x80 {
        value >>= 7;
        size += 1;
    }
    size
}

fn varint_field(field: u32, value: u64) -> usize {
    key_len(field, 0) + varint_size(value)
}

fn bool_field(field: u32) -> usize {
    varint_field(field, 1)
}

fn fixed32_field(field: u32) -> usize {
    key_len(field, 5) + 4
}

fn len_field(field: u32, inner_length: usize) -> Option<usize> {
    key_len(field, 2)
        .checked_add(varint_size(u64::try_from(inner_length).ok()?))?
        .checked_add(inner_length)
}

fn reference_len(identifier: u64) -> usize {
    varint_field(1, identifier)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn shape() -> ShapeInfoWrite {
        ShapeInfoWrite::new(
            1,
            2,
            3,
            4,
            5,
            Point::new(10.0, 20.0),
            Size::new(100.0, 40.0),
        )
    }

    #[test]
    fn shape_and_storage_measurements_are_exact() {
        let shape = encode_shape_info(&shape(), EncodeOptions::default()).unwrap();
        assert_eq!(shape.bytes().len(), shape.report().output_bytes());
        let storage =
            encode_storage(&StorageWrite::new(8, "hello"), EncodeOptions::default()).unwrap();
        assert_eq!(storage.bytes().len(), storage.report().output_bytes());
    }

    #[test]
    fn limits_are_checked_before_output_allocation() {
        let request = shape();
        let report = encode_shape_info(&request, EncodeOptions::default())
            .unwrap()
            .report();
        let error = encode_shape_info(
            &request,
            EncodeOptions::default().with_max_output_bytes(report.output_bytes() - 1),
        )
        .unwrap_err();
        assert!(matches!(
            error,
            EncodeError::Resource(EncodeLimit::OutputBytes { .. })
        ));
    }
}
