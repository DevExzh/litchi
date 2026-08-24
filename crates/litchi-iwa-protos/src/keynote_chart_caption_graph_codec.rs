//! Bounded generated-free encoding for one canonical inline Keynote chart
//! caption graph.
//!
//! The package owner remains responsible for object identifiers, archive
//! metadata, component ownership, PackageMetadata publication, and the chart
//! edge. This module only authors the four fresh protobuf payloads whose
//! schema and defaults are fixed by the native inline caption profile.

use std::fmt;

const CAPTION_STYLE_IDENTIFIER: &str = "captions-0-shapestyle-Object Caption";
const TITLE_STYLE_IDENTIFIER: &str = "captions-0-shapestyle-Object Title";
const GRAPH_OBJECTS: usize = 4;
const GRAPH_DEPTH: u32 = 9;
const GRAPH_FIELDS_WITHOUT_LANGUAGE: usize = 162;
const GRAPH_FIELDS_WITH_LANGUAGE: usize = 166;

/// The native inline child represented by a freshly authored graph.
///
/// The graph shape is shared by chart titles and captions.  The kind is
/// explicit so callers cannot accidentally author a title using the caption
/// style/placement defaults.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CaptionGraphKind {
    Caption,
    Title,
}

impl CaptionGraphKind {
    const fn style_identifier(self) -> &'static str {
        match self {
            Self::Caption => CAPTION_STYLE_IDENTIFIER,
            Self::Title => TITLE_STYLE_IDENTIFIER,
        }
    }

    const fn child_info_kind(self) -> u64 {
        match self {
            Self::Caption => 1,
            Self::Title => 2,
        }
    }

    const fn placement_anchors(self) -> (u64, u64) {
        match self {
            // TSA.CaptionPlacementArchive.caption_anchor_location,
            // TSA.CaptionPlacementArchive.drawable_anchor_location.
            Self::Caption => (7, 1),
            Self::Title => (1, 7),
        }
    }
}

/// Layout profile for a fresh caption graph.
///
/// Only the inline profile is currently admitted by this codec.  Keeping the
/// profile typed leaves the componentized graph variant out of this narrowly
/// bounded authoring seam instead of silently applying inline defaults to it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CaptionGraphProfile {
    Inline,
}

/// Short aliases for callers that use the generic graph terminology.
pub type GraphKind = CaptionGraphKind;
pub type GraphProfile = CaptionGraphProfile;

/// Finite limits for one fresh caption graph encoding.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EncodeOptions {
    max_output_bytes: usize,
    max_text_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    max_depth: u32,
    max_allocations: usize,
}

impl EncodeOptions {
    /// Build an explicit finite encoding policy.
    #[must_use]
    pub const fn new(
        max_output_bytes: usize,
        max_text_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        max_depth: u32,
        max_allocations: usize,
    ) -> Self {
        Self {
            max_output_bytes,
            max_text_bytes,
            max_fields,
            max_work_bytes,
            max_depth,
            max_allocations,
        }
    }

    /// Build a conservative finite policy from one caller-owned text value.
    #[must_use]
    pub fn for_text(text: &str) -> Self {
        let bytes = text.len().max(1);
        Self::new(
            bytes.saturating_add(16 * 1024),
            bytes,
            256,
            bytes.saturating_mul(8).saturating_add(128 * 1024),
            16,
            GRAPH_OBJECTS,
        )
    }

    /// Replace the aggregate output-byte ceiling.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }

    /// Replace the UTF-8 text-byte ceiling.
    #[must_use]
    pub const fn with_max_text_bytes(mut self, maximum: usize) -> Self {
        self.max_text_bytes = maximum;
        self
    }

    /// Replace the authored-field ceiling.
    #[must_use]
    pub const fn with_max_fields(mut self, maximum: usize) -> Self {
        self.max_fields = maximum;
        self
    }

    /// Replace the aggregate measurement-and-write work ceiling.
    #[must_use]
    pub const fn with_max_work_bytes(mut self, maximum: usize) -> Self {
        self.max_work_bytes = maximum;
        self
    }

    /// Replace the nesting ceiling.
    #[must_use]
    pub const fn with_max_depth(mut self, maximum: u32) -> Self {
        self.max_depth = maximum;
        self
    }

    /// Replace the top-level output-allocation ceiling.
    #[must_use]
    pub const fn with_max_allocations(mut self, maximum: usize) -> Self {
        self.max_allocations = maximum;
        self
    }
}

/// One canonical inline Keynote caption graph request.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct CaptionGraphWrite<'source> {
    pub drawable_identifier: u64,
    pub style_identifier: u64,
    pub info_identifier: u64,
    pub storage_identifier: u64,
    pub placement_identifier: u64,
    pub stylesheet_identifier: u64,
    pub paragraph_style_identifier: u64,
    pub drawable_width: f32,
    pub text: &'source str,
    pub language: Option<&'source str>,
}

/// The four canonical protobuf payloads for one inline caption graph.
#[derive(Debug, PartialEq, Eq)]
pub struct CaptionGraphPayloads {
    style: Vec<u8>,
    info: Vec<u8>,
    storage: Vec<u8>,
    placement: Vec<u8>,
}

impl CaptionGraphPayloads {
    #[must_use]
    pub fn style(&self) -> &[u8] {
        &self.style
    }

    #[must_use]
    pub fn info(&self) -> &[u8] {
        &self.info
    }

    #[must_use]
    pub fn storage(&self) -> &[u8] {
        &self.storage
    }

    #[must_use]
    pub fn placement(&self) -> &[u8] {
        &self.placement
    }

    #[must_use]
    pub fn into_parts(self) -> [Vec<u8>; GRAPH_OBJECTS] {
        [self.style, self.info, self.storage, self.placement]
    }
}

/// Exact finite-resource evidence for one encoding.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EncodeReport {
    output_bytes: usize,
    text_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    allocations: usize,
}

impl EncodeReport {
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }

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
    pub const fn allocations(self) -> usize {
        self.allocations
    }
}

/// Verified graph payloads plus their exact report.
#[derive(Debug, PartialEq, Eq)]
pub struct EncodeOutput {
    payloads: CaptionGraphPayloads,
    report: EncodeReport,
}

impl EncodeOutput {
    #[must_use]
    pub const fn payloads(&self) -> &CaptionGraphPayloads {
        &self.payloads
    }

    #[must_use]
    pub const fn report(&self) -> EncodeReport {
        self.report
    }

    #[must_use]
    pub fn into_payloads(self) -> CaptionGraphPayloads {
        self.payloads
    }
}

/// Finite resource exceeded by a graph encoding.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum EncodeLimit {
    OutputBytes { observed: usize, maximum: usize },
    TextBytes { observed: usize, maximum: usize },
    Fields { observed: usize, maximum: usize },
    WorkBytes { observed: usize, maximum: usize },
    Nesting { observed: u32, maximum: u32 },
    Allocations { observed: usize, maximum: usize },
}

/// Content-redacted encoding error.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum EncodeError {
    InvalidInput,
    Resource(EncodeLimit),
    Allocation { amount: usize },
    Verification,
}

impl fmt::Display for EncodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidInput => formatter.write_str("invalid Keynote caption graph input"),
            Self::Resource(_limit) => formatter.write_str("Keynote caption graph limit exceeded"),
            Self::Allocation { .. } => {
                formatter.write_str("Keynote caption graph allocation failed")
            },
            Self::Verification => formatter.write_str("Keynote caption graph verification failed"),
        }
    }
}

impl std::error::Error for EncodeError {}

/// Encode the four canonical inline graph payloads after exact sizing and
/// complete finite-limit preflight.
pub fn encode_caption_graph_with_report(
    write: CaptionGraphWrite<'_>,
    options: EncodeOptions,
) -> Result<EncodeOutput, EncodeError> {
    encode_caption_graph_with_profile_with_report(
        write,
        CaptionGraphKind::Caption,
        CaptionGraphProfile::Inline,
        options,
    )
}

/// Encode a fresh inline title or caption graph with an explicit kind/profile.
pub fn encode_caption_graph_with_profile_with_report(
    write: CaptionGraphWrite<'_>,
    kind: CaptionGraphKind,
    profile: CaptionGraphProfile,
    options: EncodeOptions,
) -> Result<EncodeOutput, EncodeError> {
    match profile {
        CaptionGraphProfile::Inline => encode_inline_graph_with_report(write, kind, options),
    }
}

/// Encode a fresh inline title or caption graph with an explicit kind.
pub fn encode_caption_graph_with_kind_with_report(
    write: CaptionGraphWrite<'_>,
    kind: CaptionGraphKind,
    options: EncodeOptions,
) -> Result<EncodeOutput, EncodeError> {
    encode_caption_graph_with_profile_with_report(write, kind, CaptionGraphProfile::Inline, options)
}

fn encode_inline_graph_with_report(
    write: CaptionGraphWrite<'_>,
    kind: CaptionGraphKind,
    options: EncodeOptions,
) -> Result<EncodeOutput, EncodeError> {
    validate_write(write)?;
    let sizes = [
        style_size(write, kind),
        info_size(write, kind),
        storage_size(write),
        placement_size(kind),
    ];
    let output_bytes = sizes
        .into_iter()
        .try_fold(0usize, usize::checked_add)
        .ok_or(EncodeError::InvalidInput)?;
    let fields = if write.language.is_some() {
        GRAPH_FIELDS_WITH_LANGUAGE
    } else {
        GRAPH_FIELDS_WITHOUT_LANGUAGE
    };
    let work_bytes = output_bytes
        .checked_mul(2)
        .and_then(|work| work.checked_add(fields))
        .ok_or(EncodeError::InvalidInput)?;
    let report = EncodeReport {
        output_bytes,
        text_bytes: write.text.len(),
        fields,
        work_bytes,
        max_depth: GRAPH_DEPTH,
        allocations: GRAPH_OBJECTS,
    };
    preflight(report, options)?;

    let mut style = reserve_exact(sizes[0])?;
    encode_style(&mut style, write, kind);
    let mut info = reserve_exact(sizes[1])?;
    encode_info(&mut info, write, kind);
    let mut storage = reserve_exact(sizes[2])?;
    encode_storage(&mut storage, write);
    let mut placement = reserve_exact(sizes[3])?;
    encode_placement(&mut placement, kind);
    if [style.len(), info.len(), storage.len(), placement.len()] != sizes {
        return Err(EncodeError::Verification);
    }
    Ok(EncodeOutput {
        payloads: CaptionGraphPayloads {
            style,
            info,
            storage,
            placement,
        },
        report,
    })
}

/// Canonical payload of a freshly allocated `TSD.StandinCaptionArchive`.
#[must_use]
pub const fn canonical_standin_payload() -> &'static [u8] {
    &[]
}

fn validate_write(write: CaptionGraphWrite<'_>) -> Result<(), EncodeError> {
    let identifiers = [
        write.drawable_identifier,
        write.style_identifier,
        write.info_identifier,
        write.storage_identifier,
        write.placement_identifier,
        write.stylesheet_identifier,
        write.paragraph_style_identifier,
    ];
    if identifiers.contains(&0)
        || identifiers
            .iter()
            .enumerate()
            .any(|(index, identifier)| identifiers[..index].contains(identifier))
        || !write.drawable_width.is_finite()
        || write.drawable_width <= 0.0
        || write.language.is_some_and(str::is_empty)
    {
        return Err(EncodeError::InvalidInput);
    }
    Ok(())
}

fn preflight(report: EncodeReport, options: EncodeOptions) -> Result<(), EncodeError> {
    if let Some(limit) = [
        (report.output_bytes > options.max_output_bytes).then_some(EncodeLimit::OutputBytes {
            observed: report.output_bytes,
            maximum: options.max_output_bytes,
        }),
        (report.text_bytes > options.max_text_bytes).then_some(EncodeLimit::TextBytes {
            observed: report.text_bytes,
            maximum: options.max_text_bytes,
        }),
        (report.fields > options.max_fields).then_some(EncodeLimit::Fields {
            observed: report.fields,
            maximum: options.max_fields,
        }),
        (report.work_bytes > options.max_work_bytes).then_some(EncodeLimit::WorkBytes {
            observed: report.work_bytes,
            maximum: options.max_work_bytes,
        }),
        (report.max_depth > options.max_depth).then_some(EncodeLimit::Nesting {
            observed: report.max_depth,
            maximum: options.max_depth,
        }),
        (report.allocations > options.max_allocations).then_some(EncodeLimit::Allocations {
            observed: report.allocations,
            maximum: options.max_allocations,
        }),
    ]
    .into_iter()
    .flatten()
    .next()
    {
        return Err(EncodeError::Resource(limit));
    }
    Ok(())
}

fn reserve_exact(amount: usize) -> Result<Vec<u8>, EncodeError> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(amount)
        .map_err(|_error| EncodeError::Allocation { amount })?;
    if output.capacity() != amount {
        return Err(EncodeError::Allocation { amount });
    }
    Ok(output)
}

fn key_size(number: u32, wire: u8) -> usize {
    varint_size((u64::from(number) << 3) | u64::from(wire))
}

fn varint_size(mut value: u64) -> usize {
    let mut size = 1;
    while value >= 0x80 {
        value >>= 7;
        size += 1;
    }
    size
}

fn varint_field_size(number: u32, value: u64) -> usize {
    key_size(number, 0) + varint_size(value)
}

fn fixed32_field_size(number: u32) -> usize {
    key_size(number, 5) + 4
}

fn bytes_field_size(number: u32, size: usize) -> usize {
    key_size(number, 2) + varint_size(size as u64) + size
}

fn put_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn put_key(output: &mut Vec<u8>, number: u32, wire: u8) {
    put_varint(output, (u64::from(number) << 3) | u64::from(wire));
}

fn put_varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
    put_key(output, number, 0);
    put_varint(output, value);
}

fn put_fixed32_field(output: &mut Vec<u8>, number: u32, value: f32) {
    put_key(output, number, 5);
    output.extend_from_slice(&value.to_bits().to_le_bytes());
}

fn put_bytes_field(output: &mut Vec<u8>, number: u32, bytes: &[u8]) {
    put_message_header(output, number, bytes.len());
    output.extend_from_slice(bytes);
}

fn put_message_header(output: &mut Vec<u8>, number: u32, size: usize) {
    put_key(output, number, 2);
    put_varint(output, size as u64);
}

fn reference_size(identifier: u64) -> usize {
    varint_field_size(1, identifier)
}

fn encode_reference(output: &mut Vec<u8>, identifier: u64) {
    put_varint_field(output, 1, identifier);
}

fn style_size(write: CaptionGraphWrite<'_>, kind: CaptionGraphKind) -> usize {
    let style_archive = bytes_field_size(2, kind.style_identifier().len())
        + bytes_field_size(5, reference_size(write.stylesheet_identifier));
    let color_transparent = color_size(0.0);
    let color_opaque = color_size(1.0);
    let fill = bytes_field_size(1, color_transparent);
    let pattern = varint_field_size(1, 2)
        + fixed32_field_size(2)
        + varint_field_size(3, 0)
        + 6 * fixed32_field_size(4);
    let stroke = bytes_field_size(1, color_opaque)
        + fixed32_field_size(2)
        + varint_field_size(3, 0)
        + varint_field_size(4, 0)
        + fixed32_field_size(5)
        + bytes_field_size(6, pattern);
    let shadow = bytes_field_size(1, color_opaque)
        + fixed32_field_size(2)
        + fixed32_field_size(3)
        + varint_field_size(4, 1)
        + fixed32_field_size(5)
        + varint_field_size(6, 0)
        + varint_field_size(7, 0);
    let shape_properties = bytes_field_size(1, fill)
        + bytes_field_size(2, stroke)
        + fixed32_field_size(3)
        + bytes_field_size(4, shadow)
        + bytes_field_size(5, 0);
    let tsd_style = bytes_field_size(1, style_archive)
        + varint_field_size(10, 7)
        + bytes_field_size(11, shape_properties);
    let padding = 4 * fixed32_field_size(1);
    let text_properties = bytes_field_size(6, padding)
        + bytes_field_size(10, reference_size(write.paragraph_style_identifier));
    bytes_field_size(1, tsd_style)
        + varint_field_size(10, 7)
        + bytes_field_size(11, text_properties)
}

fn encode_style(output: &mut Vec<u8>, write: CaptionGraphWrite<'_>, kind: CaptionGraphKind) {
    let style_archive = bytes_field_size(2, kind.style_identifier().len())
        + bytes_field_size(5, reference_size(write.stylesheet_identifier));
    let color_transparent = color_size(0.0);
    let color_opaque = color_size(1.0);
    let fill = bytes_field_size(1, color_transparent);
    let pattern = varint_field_size(1, 2)
        + fixed32_field_size(2)
        + varint_field_size(3, 0)
        + 6 * fixed32_field_size(4);
    let stroke = bytes_field_size(1, color_opaque)
        + fixed32_field_size(2)
        + varint_field_size(3, 0)
        + varint_field_size(4, 0)
        + fixed32_field_size(5)
        + bytes_field_size(6, pattern);
    let shadow = bytes_field_size(1, color_opaque)
        + fixed32_field_size(2)
        + fixed32_field_size(3)
        + varint_field_size(4, 1)
        + fixed32_field_size(5)
        + varint_field_size(6, 0)
        + varint_field_size(7, 0);
    let shape_properties = bytes_field_size(1, fill)
        + bytes_field_size(2, stroke)
        + fixed32_field_size(3)
        + bytes_field_size(4, shadow)
        + bytes_field_size(5, 0);
    let tsd_style = bytes_field_size(1, style_archive)
        + varint_field_size(10, 7)
        + bytes_field_size(11, shape_properties);
    let padding = fixed32_field_size(1)
        + fixed32_field_size(2)
        + fixed32_field_size(3)
        + fixed32_field_size(4);
    let text_properties = bytes_field_size(6, padding)
        + bytes_field_size(10, reference_size(write.paragraph_style_identifier));

    put_message_header(output, 1, tsd_style);
    put_message_header(output, 1, style_archive);
    put_bytes_field(output, 2, kind.style_identifier().as_bytes());
    put_message_header(output, 5, reference_size(write.stylesheet_identifier));
    encode_reference(output, write.stylesheet_identifier);
    put_varint_field(output, 10, 7);
    put_message_header(output, 11, shape_properties);
    put_message_header(output, 1, fill);
    put_message_header(output, 1, color_transparent);
    encode_color(output, 0.0);
    put_message_header(output, 2, stroke);
    put_message_header(output, 1, color_opaque);
    encode_color(output, 1.0);
    put_fixed32_field(output, 2, 1.0);
    put_varint_field(output, 3, 0);
    put_varint_field(output, 4, 0);
    put_fixed32_field(output, 5, 4.0);
    put_message_header(output, 6, pattern);
    put_varint_field(output, 1, 2);
    put_fixed32_field(output, 2, 0.0);
    put_varint_field(output, 3, 0);
    for _ in 0..6 {
        put_fixed32_field(output, 4, 0.0);
    }
    put_fixed32_field(output, 3, 1.0);
    put_message_header(output, 4, shadow);
    put_message_header(output, 1, color_opaque);
    encode_color(output, 1.0);
    put_fixed32_field(output, 2, 315.0);
    put_fixed32_field(output, 3, 5.0);
    put_varint_field(output, 4, 1);
    put_fixed32_field(output, 5, 1.0);
    put_varint_field(output, 6, 0);
    put_varint_field(output, 7, 0);
    put_message_header(output, 5, 0);
    put_varint_field(output, 10, 7);
    put_message_header(output, 11, text_properties);
    put_message_header(output, 6, padding);
    for number in 1..=4 {
        put_fixed32_field(output, number, 4.0);
    }
    put_message_header(output, 10, reference_size(write.paragraph_style_identifier));
    encode_reference(output, write.paragraph_style_identifier);
}

fn color_size(alpha: f32) -> usize {
    let _ = alpha;
    varint_field_size(1, 1)
        + fixed32_field_size(3)
        + fixed32_field_size(4)
        + fixed32_field_size(5)
        + varint_field_size(12, 1)
        + fixed32_field_size(6)
}

fn encode_color(output: &mut Vec<u8>, alpha: f32) {
    put_varint_field(output, 1, 1);
    put_fixed32_field(output, 3, 0.0);
    put_fixed32_field(output, 4, 0.0);
    put_fixed32_field(output, 5, 0.0);
    put_fixed32_field(output, 6, alpha);
    put_varint_field(output, 12, 1);
}

fn info_size(write: CaptionGraphWrite<'_>, kind: CaptionGraphKind) -> usize {
    let point = 2 * fixed32_field_size(1);
    let size = fixed32_field_size(1) + fixed32_field_size(2);
    let geometry = bytes_field_size(1, point)
        + bytes_field_size(2, size)
        + varint_field_size(3, 1)
        + fixed32_field_size(4);
    let wrap = varint_field_size(1, 4)
        + varint_field_size(2, 2)
        + varint_field_size(3, 1)
        + fixed32_field_size(4)
        + fixed32_field_size(5)
        + varint_field_size(6, 0);
    let drawable = bytes_field_size(1, geometry)
        + bytes_field_size(2, reference_size(write.drawable_identifier))
        + bytes_field_size(3, wrap)
        + varint_field_size(5, 0)
        + varint_field_size(7, 0)
        + varint_field_size(12, 0)
        + varint_field_size(13, 0);
    let path = path_source_size(write.drawable_width);
    let shape = bytes_field_size(1, drawable)
        + bytes_field_size(2, reference_size(write.style_identifier))
        + bytes_field_size(3, path)
        + fixed32_field_size(6);
    let shape_info = bytes_field_size(1, shape)
        + bytes_field_size(2, reference_size(write.storage_identifier))
        + bytes_field_size(4, reference_size(write.storage_identifier))
        + varint_field_size(6, 1);
    bytes_field_size(1, shape_info)
        + bytes_field_size(2, reference_size(write.placement_identifier))
        + varint_field_size(3, kind.child_info_kind())
}

fn encode_info(output: &mut Vec<u8>, write: CaptionGraphWrite<'_>, kind: CaptionGraphKind) {
    let point = fixed32_field_size(1) + fixed32_field_size(2);
    let size = fixed32_field_size(1) + fixed32_field_size(2);
    let geometry = bytes_field_size(1, point)
        + bytes_field_size(2, size)
        + varint_field_size(3, 1)
        + fixed32_field_size(4);
    let wrap = varint_field_size(1, 4)
        + varint_field_size(2, 2)
        + varint_field_size(3, 1)
        + fixed32_field_size(4)
        + fixed32_field_size(5)
        + varint_field_size(6, 0);
    let drawable = bytes_field_size(1, geometry)
        + bytes_field_size(2, reference_size(write.drawable_identifier))
        + bytes_field_size(3, wrap)
        + varint_field_size(5, 0)
        + varint_field_size(7, 0)
        + varint_field_size(12, 0)
        + varint_field_size(13, 0);
    let path = path_source_size(write.drawable_width);
    let shape = bytes_field_size(1, drawable)
        + bytes_field_size(2, reference_size(write.style_identifier))
        + bytes_field_size(3, path)
        + fixed32_field_size(6);
    let shape_info = bytes_field_size(1, shape)
        + bytes_field_size(2, reference_size(write.storage_identifier))
        + bytes_field_size(4, reference_size(write.storage_identifier))
        + varint_field_size(6, 1);
    put_message_header(output, 1, shape_info);
    put_message_header(output, 1, shape);
    put_message_header(output, 1, drawable);
    put_message_header(output, 1, geometry);
    put_message_header(output, 1, point);
    put_fixed32_field(output, 1, 0.0);
    put_fixed32_field(output, 2, 0.0);
    put_message_header(output, 2, size);
    put_fixed32_field(output, 1, write.drawable_width);
    put_fixed32_field(output, 2, 0.0);
    put_varint_field(output, 3, 1);
    put_fixed32_field(output, 4, 0.0);
    put_message_header(output, 2, reference_size(write.drawable_identifier));
    encode_reference(output, write.drawable_identifier);
    put_message_header(output, 3, wrap);
    put_varint_field(output, 1, 4);
    put_varint_field(output, 2, 2);
    put_varint_field(output, 3, 1);
    put_fixed32_field(output, 4, 12.0);
    put_fixed32_field(output, 5, 0.5);
    put_varint_field(output, 6, 0);
    put_varint_field(output, 5, 0);
    put_varint_field(output, 7, 0);
    put_varint_field(output, 12, 0);
    put_varint_field(output, 13, 0);
    put_message_header(output, 2, reference_size(write.style_identifier));
    encode_reference(output, write.style_identifier);
    put_message_header(output, 3, path);
    encode_path_source(output, write.drawable_width);
    put_fixed32_field(output, 6, 0.0);
    put_message_header(output, 2, reference_size(write.storage_identifier));
    encode_reference(output, write.storage_identifier);
    put_message_header(output, 4, reference_size(write.storage_identifier));
    encode_reference(output, write.storage_identifier);
    put_varint_field(output, 6, 1);
    put_message_header(output, 2, reference_size(write.placement_identifier));
    encode_reference(output, write.placement_identifier);
    put_varint_field(output, 3, kind.child_info_kind());
}

fn path_source_size(_width: f32) -> usize {
    let size = fixed32_field_size(1) + fixed32_field_size(2);
    let point = fixed32_field_size(1) + fixed32_field_size(2);
    let element_point = varint_field_size(1, 1) + bytes_field_size(2, point);
    let element_line = varint_field_size(1, 2) + bytes_field_size(2, point);
    let element_close = varint_field_size(1, 5);
    let path = bytes_field_size(1, element_point)
        + 3 * bytes_field_size(1, element_line)
        + bytes_field_size(1, element_close)
        + bytes_field_size(1, element_point);
    let bezier = bytes_field_size(2, size) + bytes_field_size(3, path);
    varint_field_size(1, 0) + varint_field_size(2, 0) + bytes_field_size(5, bezier)
}

fn encode_path_source(output: &mut Vec<u8>, width: f32) {
    let size = fixed32_field_size(1) + fixed32_field_size(2);
    let point = fixed32_field_size(1) + fixed32_field_size(2);
    let element_with_point = |kind| varint_field_size(1, kind) + bytes_field_size(2, point);
    let element_close = varint_field_size(1, 5);
    let path = bytes_field_size(1, element_with_point(1))
        + 3 * bytes_field_size(1, element_with_point(2))
        + bytes_field_size(1, element_close)
        + bytes_field_size(1, element_with_point(1));
    let bezier = bytes_field_size(2, size) + bytes_field_size(3, path);
    put_varint_field(output, 1, 0);
    put_varint_field(output, 2, 0);
    put_message_header(output, 5, bezier);
    put_message_header(output, 2, size);
    put_fixed32_field(output, 1, width);
    put_fixed32_field(output, 2, 0.0);
    put_message_header(output, 3, path);
    for (kind, x, y) in [
        (1, 0.0, 0.0),
        (2, 100.0, 0.0),
        (2, 100.0, 100.0),
        (2, 0.0, 100.0),
    ] {
        let element = element_with_point(kind);
        put_message_header(output, 1, element);
        put_varint_field(output, 1, kind);
        put_message_header(output, 2, point);
        put_fixed32_field(output, 1, x);
        put_fixed32_field(output, 2, y);
    }
    put_message_header(output, 1, element_close);
    put_varint_field(output, 1, 5);
    let element = element_with_point(1);
    put_message_header(output, 1, element);
    put_varint_field(output, 1, 1);
    put_message_header(output, 2, point);
    put_fixed32_field(output, 1, 0.0);
    put_fixed32_field(output, 2, 0.0);
}

fn object_attribute_size(with_reference: Option<u64>) -> usize {
    let entry = varint_field_size(1, 0)
        + with_reference.map_or(0, |identifier| {
            bytes_field_size(2, reference_size(identifier))
        });
    bytes_field_size(1, entry)
}

fn encode_object_attribute(output: &mut Vec<u8>, with_reference: Option<u64>) {
    let entry = varint_field_size(1, 0)
        + with_reference.map_or(0, |identifier| {
            bytes_field_size(2, reference_size(identifier))
        });
    put_message_header(output, 1, entry);
    put_varint_field(output, 1, 0);
    if let Some(identifier) = with_reference {
        put_message_header(output, 2, reference_size(identifier));
        encode_reference(output, identifier);
    }
}

fn para_data_size() -> usize {
    let entry = varint_field_size(1, 0) + varint_field_size(2, 0) + varint_field_size(3, 0);
    bytes_field_size(1, entry)
}

fn encode_para_data(output: &mut Vec<u8>) {
    let entry = varint_field_size(1, 0) + varint_field_size(2, 0) + varint_field_size(3, 0);
    put_message_header(output, 1, entry);
    put_varint_field(output, 1, 0);
    put_varint_field(output, 2, 0);
    put_varint_field(output, 3, 0);
}

fn language_size(language: &str) -> usize {
    let entry = varint_field_size(1, 0) + bytes_field_size(2, language.len());
    bytes_field_size(1, entry)
}

fn encode_language(output: &mut Vec<u8>, language: &str) {
    let entry = varint_field_size(1, 0) + bytes_field_size(2, language.len());
    put_message_header(output, 1, entry);
    put_varint_field(output, 1, 0);
    put_bytes_field(output, 2, language.as_bytes());
}

fn storage_size(write: CaptionGraphWrite<'_>) -> usize {
    bytes_field_size(2, reference_size(write.stylesheet_identifier))
        + bytes_field_size(3, write.text.len())
        + varint_field_size(10, 1)
        + bytes_field_size(
            5,
            object_attribute_size(Some(write.paragraph_style_identifier)),
        )
        + bytes_field_size(6, para_data_size())
        + bytes_field_size(14, para_data_size())
        + write
            .language
            .map_or(0, |language| bytes_field_size(19, language_size(language)))
        + bytes_field_size(24, para_data_size())
        + bytes_field_size(28, object_attribute_size(None))
}

fn encode_storage(output: &mut Vec<u8>, write: CaptionGraphWrite<'_>) {
    put_message_header(output, 2, reference_size(write.stylesheet_identifier));
    encode_reference(output, write.stylesheet_identifier);
    put_bytes_field(output, 3, write.text.as_bytes());
    put_message_header(
        output,
        5,
        object_attribute_size(Some(write.paragraph_style_identifier)),
    );
    encode_object_attribute(output, Some(write.paragraph_style_identifier));
    put_message_header(output, 6, para_data_size());
    encode_para_data(output);
    put_varint_field(output, 10, 1);
    put_message_header(output, 14, para_data_size());
    encode_para_data(output);
    if let Some(language) = write.language {
        put_message_header(output, 19, language_size(language));
        encode_language(output, language);
    }
    put_message_header(output, 24, para_data_size());
    encode_para_data(output);
    put_message_header(output, 28, object_attribute_size(None));
    encode_object_attribute(output, None);
}

fn placement_size(kind: CaptionGraphKind) -> usize {
    let (caption_anchor, drawable_anchor) = kind.placement_anchors();
    varint_field_size(1, caption_anchor) + varint_field_size(2, drawable_anchor)
}

fn encode_placement(output: &mut Vec<u8>, kind: CaptionGraphKind) {
    let (caption_anchor, drawable_anchor) = kind.placement_anchors();
    put_varint_field(output, 1, caption_anchor);
    put_varint_field(output, 2, drawable_anchor);
}

#[cfg(test)]
#[allow(deprecated)]
mod tests {
    use super::*;
    use crate::{tsa, tsd, tsp, tss, tswp};
    use prost::Message;

    fn reference(identifier: u64) -> tsp::Reference {
        tsp::Reference {
            identifier,
            ..Default::default()
        }
    }

    fn color(alpha: f32) -> tsp::Color {
        tsp::Color {
            model: 1,
            r: Some(0.0),
            g: Some(0.0),
            b: Some(0.0),
            rgbspace: Some(1),
            a: Some(alpha),
            ..Default::default()
        }
    }

    fn object_attributes(object: Option<tsp::Reference>) -> tswp::ObjectAttributeTable {
        tswp::ObjectAttributeTable {
            entries: vec![tswp::object_attribute_table::ObjectAttribute {
                character_index: 0,
                object,
            }],
        }
    }

    fn para_data() -> tswp::ParaDataAttributeTable {
        tswp::ParaDataAttributeTable {
            entries: vec![tswp::para_data_attribute_table::ParaDataAttribute {
                character_index: 0,
                first: 0,
                second: 0,
            }],
        }
    }

    fn oracle_style(write: CaptionGraphWrite<'_>) -> Vec<u8> {
        oracle_style_kind(write, CaptionGraphKind::Caption)
    }

    fn oracle_style_kind(write: CaptionGraphWrite<'_>, kind: CaptionGraphKind) -> Vec<u8> {
        tswp::ShapeStyleArchive {
            super_: tsd::ShapeStyleArchive {
                super_: tss::StyleArchive {
                    style_identifier: Some(kind.style_identifier().to_owned()),
                    stylesheet: Some(reference(write.stylesheet_identifier)),
                    ..Default::default()
                },
                override_count: Some(7),
                shape_properties: Some(tsd::ShapeStylePropertiesArchive {
                    fill: Some(tsd::FillArchive {
                        color: Some(color(0.0)),
                        ..Default::default()
                    }),
                    stroke: Some(tsd::StrokeArchive {
                        color: Some(color(1.0)),
                        width: Some(1.0),
                        cap: Some(0),
                        join: Some(0),
                        miter_limit: Some(4.0),
                        pattern: Some(tsd::StrokePatternArchive {
                            r#type: Some(2),
                            phase: Some(0.0),
                            count: Some(0),
                            pattern: vec![0.0; 6],
                        }),
                        ..Default::default()
                    }),
                    opacity: Some(1.0),
                    shadow: Some(tsd::ShadowArchive {
                        color: Some(color(1.0)),
                        angle: Some(315.0),
                        offset: Some(5.0),
                        radius: Some(1),
                        opacity: Some(1.0),
                        is_enabled: Some(false),
                        r#type: Some(0),
                        ..Default::default()
                    }),
                    reflection: Some(tsd::ReflectionArchive::default()),
                    ..Default::default()
                }),
            },
            override_count: Some(7),
            shape_properties: Some(tswp::ShapeStylePropertiesArchive {
                padding: Some(tswp::PaddingArchive {
                    left: Some(4.0),
                    top: Some(4.0),
                    right: Some(4.0),
                    bottom: Some(4.0),
                }),
                paragraph_style: Some(reference(write.paragraph_style_identifier)),
                ..Default::default()
            }),
        }
        .encode_to_vec()
    }

    fn oracle_path(width: f32) -> tsd::PathSourceArchive {
        let point = |x, y| tsp::Point { x, y };
        let element = |r#type, points| tsp::path::Element { r#type, points };
        tsd::PathSourceArchive {
            horizontal_flip: Some(false),
            vertical_flip: Some(false),
            bezier_path_source: Some(tsd::BezierPathSourceArchive {
                natural_size: Some(tsp::Size { width, height: 0.0 }),
                path: Some(tsp::Path {
                    elements: vec![
                        element(1, vec![point(0.0, 0.0)]),
                        element(2, vec![point(100.0, 0.0)]),
                        element(2, vec![point(100.0, 100.0)]),
                        element(2, vec![point(0.0, 100.0)]),
                        element(5, Vec::new()),
                        element(1, vec![point(0.0, 0.0)]),
                    ],
                }),
                ..Default::default()
            }),
            ..Default::default()
        }
    }

    fn oracle_info(write: CaptionGraphWrite<'_>) -> Vec<u8> {
        oracle_info_kind(write, CaptionGraphKind::Caption)
    }

    fn oracle_info_kind(write: CaptionGraphWrite<'_>, kind: CaptionGraphKind) -> Vec<u8> {
        tsa::CaptionInfoArchive {
            super_: tswp::ShapeInfoArchive {
                super_: tsd::ShapeArchive {
                    super_: tsd::DrawableArchive {
                        geometry: Some(tsd::GeometryArchive {
                            position: Some(tsp::Point { x: 0.0, y: 0.0 }),
                            size: Some(tsp::Size {
                                width: write.drawable_width,
                                height: 0.0,
                            }),
                            flags: Some(1),
                            angle: Some(0.0),
                        }),
                        parent: Some(reference(write.drawable_identifier)),
                        exterior_text_wrap: Some(tsd::ExteriorTextWrapArchive {
                            r#type: Some(4),
                            direction: Some(2),
                            fit_type: Some(1),
                            margin: Some(12.0),
                            alpha_threshold: Some(0.5),
                            is_html_wrap: Some(false),
                        }),
                        locked: Some(false),
                        aspect_ratio_locked: Some(false),
                        title_hidden: Some(false),
                        caption_hidden: Some(false),
                        ..Default::default()
                    },
                    style: Some(reference(write.style_identifier)),
                    pathsource: Some(oracle_path(write.drawable_width)),
                    stroke_pattern_offset_distance: Some(0.0),
                    ..Default::default()
                },
                deprecated_storage: Some(reference(write.storage_identifier)),
                owned_storage: Some(reference(write.storage_identifier)),
                is_text_box: Some(true),
                ..Default::default()
            },
            placement: Some(reference(write.placement_identifier)),
            child_info_kind: Some(kind.child_info_kind() as i32),
        }
        .encode_to_vec()
    }

    fn oracle_storage(write: CaptionGraphWrite<'_>) -> Vec<u8> {
        tswp::StorageArchive {
            style_sheet: Some(reference(write.stylesheet_identifier)),
            text: vec![write.text.to_owned()],
            in_document: Some(true),
            table_para_style: Some(object_attributes(Some(reference(
                write.paragraph_style_identifier,
            )))),
            table_para_data: Some(para_data()),
            table_para_starts: Some(para_data()),
            table_language: write.language.map(|language| tswp::StringAttributeTable {
                entries: vec![tswp::string_attribute_table::StringAttribute {
                    character_index: 0,
                    object: Some(language.to_owned()),
                }],
            }),
            table_para_bidi: Some(para_data()),
            table_drop_cap_style: Some(object_attributes(None)),
            ..Default::default()
        }
        .encode_to_vec()
    }

    fn oracle_placement() -> Vec<u8> {
        oracle_placement_kind(CaptionGraphKind::Caption)
    }

    fn oracle_placement_kind(kind: CaptionGraphKind) -> Vec<u8> {
        let (caption_anchor, drawable_anchor) = kind.placement_anchors();
        tsa::CaptionPlacementArchive {
            caption_anchor_location: Some(caption_anchor as i32),
            drawable_anchor_location: Some(drawable_anchor as i32),
        }
        .encode_to_vec()
    }

    fn write<'a>(text: &'a str, language: Option<&'a str>) -> CaptionGraphWrite<'a> {
        CaptionGraphWrite {
            drawable_identifier: 100,
            style_identifier: 201,
            info_identifier: 202,
            storage_identifier: 203,
            placement_identifier: 204,
            stylesheet_identifier: 300,
            paragraph_style_identifier: 301,
            drawable_width: 640.5,
            text,
            language,
        }
    }

    #[test]
    fn canonical_payloads_match_generated_oracle() {
        let write = write("Caption — 北区", Some("zh-Hans"));
        let output = encode_caption_graph_with_report(write, EncodeOptions::for_text(write.text))
            .expect("canonical graph");
        assert_eq!(output.payloads().style(), oracle_style(write));
        assert_eq!(output.payloads().info(), oracle_info(write));
        assert_eq!(output.payloads().storage(), oracle_storage(write));
        assert_eq!(output.payloads().placement(), oracle_placement());
        assert_eq!(
            output.report().output_bytes(),
            output
                .payloads()
                .parts_for_test()
                .into_iter()
                .map(<[u8]>::len)
                .sum()
        );
        assert_eq!(output.report().fields(), GRAPH_FIELDS_WITH_LANGUAGE);
        assert_eq!(output.report().allocations(), GRAPH_OBJECTS);
    }

    #[test]
    fn inline_title_profile_matches_generated_oracle_and_anchors() {
        let write = write("Title — 北区", Some("zh-Hans"));
        let output = encode_caption_graph_with_profile_with_report(
            write,
            CaptionGraphKind::Title,
            CaptionGraphProfile::Inline,
            EncodeOptions::for_text(write.text),
        )
        .expect("canonical title graph");
        assert_eq!(
            output.payloads().style(),
            oracle_style_kind(write, CaptionGraphKind::Title)
        );
        assert_eq!(
            output.payloads().info(),
            oracle_info_kind(write, CaptionGraphKind::Title)
        );
        assert_eq!(output.payloads().storage(), oracle_storage(write));
        assert_eq!(
            output.payloads().placement(),
            oracle_placement_kind(CaptionGraphKind::Title)
        );

        let style = tswp::ShapeStyleArchive::decode(output.payloads().style()).expect("style");
        assert_eq!(
            style.super_.super_.style_identifier.as_deref(),
            Some(TITLE_STYLE_IDENTIFIER)
        );
        let info = tsa::CaptionInfoArchive::decode(output.payloads().info()).expect("info");
        assert_eq!(info.child_info_kind, Some(2));
        let placement =
            tsa::CaptionPlacementArchive::decode(output.payloads().placement()).expect("placement");
        assert_eq!(placement.caption_anchor_location, Some(1));
        assert_eq!(placement.drawable_anchor_location, Some(7));
    }

    #[test]
    fn legacy_caption_entry_point_is_exact_inline_caption_profile() {
        let write = write("Caption compatibility", None);
        let legacy = encode_caption_graph_with_report(write, EncodeOptions::for_text(write.text))
            .expect("legacy caption graph");
        let explicit = encode_caption_graph_with_profile_with_report(
            write,
            CaptionGraphKind::Caption,
            CaptionGraphProfile::Inline,
            EncodeOptions::for_text(write.text),
        )
        .expect("explicit caption graph");
        assert_eq!(legacy, explicit);
    }

    #[test]
    fn inline_title_profile_replays_exact_report_limits() {
        let write = write("bounded title", Some("en"));
        let baseline = encode_caption_graph_with_kind_with_report(
            write,
            CaptionGraphKind::Title,
            EncodeOptions::for_text(write.text),
        )
        .expect("baseline title graph");
        let report = baseline.report();
        let exact = EncodeOptions::new(
            report.output_bytes(),
            report.text_bytes(),
            report.fields(),
            report.work_bytes(),
            report.max_depth(),
            report.allocations(),
        );
        assert!(
            encode_caption_graph_with_kind_with_report(write, CaptionGraphKind::Title, exact)
                .is_ok()
        );
        for limited in [
            exact.with_max_output_bytes(report.output_bytes() - 1),
            exact.with_max_text_bytes(report.text_bytes() - 1),
            exact.with_max_fields(report.fields() - 1),
            exact.with_max_work_bytes(report.work_bytes() - 1),
            exact.with_max_depth(report.max_depth() - 1),
            exact.with_max_allocations(report.allocations() - 1),
        ] {
            assert!(matches!(
                encode_caption_graph_with_kind_with_report(write, CaptionGraphKind::Title, limited),
                Err(EncodeError::Resource(_))
            ));
        }
    }

    #[test]
    fn empty_text_and_absent_language_preserve_presence() {
        let write = write("", None);
        let output = encode_caption_graph_with_report(write, EncodeOptions::for_text(write.text))
            .expect("empty active caption graph");
        assert_eq!(output.payloads().storage(), oracle_storage(write));
        let decoded = tswp::StorageArchive::decode(output.payloads().storage()).expect("storage");
        assert_eq!(decoded.text, vec![String::new()]);
        assert!(decoded.table_language.is_none());
        assert_eq!(output.report().fields(), GRAPH_FIELDS_WITHOUT_LANGUAGE);
    }

    #[test]
    fn invalid_identity_geometry_and_language_are_rejected() {
        let baseline = write("caption", None);
        for invalid in [
            CaptionGraphWrite {
                style_identifier: 0,
                ..baseline
            },
            CaptionGraphWrite {
                style_identifier: baseline.stylesheet_identifier,
                ..baseline
            },
            CaptionGraphWrite {
                drawable_width: f32::NAN,
                ..baseline
            },
            CaptionGraphWrite {
                drawable_width: 0.0,
                ..baseline
            },
            CaptionGraphWrite {
                language: Some(""),
                ..baseline
            },
        ] {
            assert_eq!(
                encode_caption_graph_with_report(invalid, EncodeOptions::for_text(invalid.text)),
                Err(EncodeError::InvalidInput)
            );
        }
    }

    #[test]
    fn exact_resource_limits_are_accepted_and_minus_one_refused() {
        let write = write("bounded caption", Some("en"));
        let baseline = encode_caption_graph_with_report(write, EncodeOptions::for_text(write.text))
            .expect("baseline");
        let report = baseline.report();
        let exact = EncodeOptions::new(
            report.output_bytes(),
            report.text_bytes(),
            report.fields(),
            report.work_bytes(),
            report.max_depth(),
            report.allocations(),
        );
        assert!(encode_caption_graph_with_report(write, exact).is_ok());
        for limited in [
            exact.with_max_output_bytes(report.output_bytes() - 1),
            exact.with_max_text_bytes(report.text_bytes() - 1),
            exact.with_max_fields(report.fields() - 1),
            exact.with_max_work_bytes(report.work_bytes() - 1),
            exact.with_max_depth(report.max_depth() - 1),
            exact.with_max_allocations(report.allocations() - 1),
        ] {
            assert!(matches!(
                encode_caption_graph_with_report(write, limited),
                Err(EncodeError::Resource(_))
            ));
        }
        assert_eq!(canonical_standin_payload(), &[]);
        assert_eq!(
            tsd::StandinCaptionArchive::decode(canonical_standin_payload()).expect("stand-in"),
            tsd::StandinCaptionArchive::default()
        );
    }

    trait PayloadSlicesForTest {
        fn parts_for_test(&self) -> [&[u8]; GRAPH_OBJECTS];
    }

    impl PayloadSlicesForTest for CaptionGraphPayloads {
        fn parts_for_test(&self) -> [&[u8]; GRAPH_OBJECTS] {
            [self.style(), self.info(), self.storage(), self.placement()]
        }
    }
}
