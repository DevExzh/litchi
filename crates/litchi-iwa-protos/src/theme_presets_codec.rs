//! Bounded construction of the five `TSS.ThemeArchive` preset extensions.
//!
//! The shared theme envelope is intentionally separate from its application
//! preset payloads.  This module owns the neutral wire construction needed by
//! fresh Pages, Numbers, and Keynote packages: typed family/list selectors,
//! source-borrowed reference slices, and the small TSD fill/shadow records
//! used by the native blank-document templates.  It does not expose generated
//! protobuf messages and it performs all size, field, work, and nesting checks
//! before reserving the single output buffer.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The public selector API is grouped by native family, then by wire helpers."
)]

use core::fmt;

pub use crate::keynote_slide_background_codec::{
    ColorWrite, GradientKind, GradientStopWrite, GradientWrite, RgbSpace,
};
use crate::theme_codec::ThemeReferenceSnapshot;

const DRAWING_FIELD: u32 = 100;
const TEXT_FIELD: u32 = 110;
const CHART_FIELD: u32 = 120;
const TABLE_FIELD: u32 = 200;
const APPLICATION_FIELD: u32 = 210;

const DRAWING_GRADIENT_FILL_FIELD: u32 = 1;
const DRAWING_IMAGE_FILL_FIELD: u32 = 2;
const DRAWING_SHADOW_FIELD: u32 = 3;

const FILL_COLOR_FIELD: u32 = 1;
const FILL_GRADIENT_FIELD: u32 = 2;
const FILL_IMAGE_FIELD: u32 = 3;

const COLOR_MODEL_FIELD: u32 = 1;
const COLOR_RED_FIELD: u32 = 3;
const COLOR_GREEN_FIELD: u32 = 4;
const COLOR_BLUE_FIELD: u32 = 5;
const COLOR_ALPHA_FIELD: u32 = 6;
const COLOR_SPACE_FIELD: u32 = 12;

const GRADIENT_TYPE_FIELD: u32 = 1;
const GRADIENT_STOP_FIELD: u32 = 2;
const GRADIENT_OPACITY_FIELD: u32 = 3;
const GRADIENT_ADVANCED_FIELD: u32 = 4;
const GRADIENT_ANGLE_FIELD: u32 = 5;
const ANGLE_RADIANS_FIELD: u32 = 2;
const STOP_COLOR_FIELD: u32 = 1;
const STOP_FRACTION_FIELD: u32 = 2;
const STOP_INFLECTION_FIELD: u32 = 3;

const IMAGE_TECHNIQUE_FIELD: u32 = 2;
const IMAGE_TINT_FIELD: u32 = 3;
const IMAGE_SIZE_FIELD: u32 = 4;
const IMAGE_UNTAGGED_FIELD: u32 = 8;
const IMAGE_REFERENCE_COLOR_FIELD: u32 = 9;
const SIZE_WIDTH_FIELD: u32 = 1;
const SIZE_HEIGHT_FIELD: u32 = 2;

const SHADOW_COLOR_FIELD: u32 = 1;
const SHADOW_ANGLE_FIELD: u32 = 2;
const SHADOW_OFFSET_FIELD: u32 = 3;
const SHADOW_RADIUS_FIELD: u32 = 4;
const SHADOW_OPACITY_FIELD: u32 = 5;
const SHADOW_ENABLED_FIELD: u32 = 6;
const SHADOW_TYPE_FIELD: u32 = 7;
const SHADOW_DROP_FIELD: u32 = 8;
const SHADOW_CONTACT_FIELD: u32 = 9;
const SHADOW_CURVED_FIELD: u32 = 10;
const CONTACT_HEIGHT_FIELD: u32 = 2;
const CONTACT_OFFSET_FIELD: u32 = 4;
const CURVED_CURVE_FIELD: u32 = 1;

const RGB_MODEL: u64 = 1;
const SRGB_SPACE: u64 = 1;
const P3_SPACE: u64 = 2;
const MAX_RECURSION_LIMIT: u32 = 64;
const MAX_AUTOMATIC_LIMIT: usize = 64 * 1024 * 1024;

/// Preset extension family stored in a shared theme.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ThemePresetFamily {
    /// `TSD.ThemePresetsArchive` at extension field 100.
    Drawing,
    /// `TSWP.ThemePresetsArchive` at extension field 110.
    Text,
    /// `TSCH.ChartPresetsArchive` at extension field 120.
    Chart,
    /// `TST.ThemePresetsArchive` at extension field 200.
    Table,
    /// `TSA.ThemePresetsArchive` at extension field 210.
    Application,
}

impl ThemePresetFamily {
    /// Native `TSS.ThemeArchive` extension field number.
    #[must_use]
    pub const fn field_number(self) -> u32 {
        match self {
            Self::Drawing => DRAWING_FIELD,
            Self::Text => TEXT_FIELD,
            Self::Chart => CHART_FIELD,
            Self::Table => TABLE_FIELD,
            Self::Application => APPLICATION_FIELD,
        }
    }
}

/// Typed repeated-reference list in one preset family.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ThemePresetList {
    DrawingLineStyle,
    DrawingShapeStyle,
    DrawingTextboxStyle,
    DrawingImageStyle,
    DrawingMovieStyle,
    DrawingLine,
    TextListStyle,
    TextTextStyle,
    TextImportedTextStyle,
    TextTocEntryStyle,
    TextTocSettings,
    TextCharacterStyle,
    TextParagraphStyle,
    TextDropcapStyle,
    ChartStyle,
    TableStyle,
    TableCellStroke,
    ApplicationCaptionStyle,
    ApplicationSvgImportStyle,
}

impl ThemePresetList {
    /// Family containing this repeated list.
    #[must_use]
    pub const fn family(self) -> ThemePresetFamily {
        match self {
            Self::DrawingLineStyle
            | Self::DrawingShapeStyle
            | Self::DrawingTextboxStyle
            | Self::DrawingImageStyle
            | Self::DrawingMovieStyle
            | Self::DrawingLine => ThemePresetFamily::Drawing,
            Self::TextListStyle
            | Self::TextTextStyle
            | Self::TextImportedTextStyle
            | Self::TextTocEntryStyle
            | Self::TextTocSettings
            | Self::TextCharacterStyle
            | Self::TextParagraphStyle
            | Self::TextDropcapStyle => ThemePresetFamily::Text,
            Self::ChartStyle => ThemePresetFamily::Chart,
            Self::TableStyle | Self::TableCellStroke => ThemePresetFamily::Table,
            Self::ApplicationCaptionStyle | Self::ApplicationSvgImportStyle => {
                ThemePresetFamily::Application
            },
        }
    }

    const fn field_number(self) -> u32 {
        match self {
            Self::DrawingLineStyle => 4,
            Self::DrawingShapeStyle => 5,
            Self::DrawingTextboxStyle => 6,
            Self::DrawingImageStyle => 7,
            Self::DrawingMovieStyle => 8,
            Self::DrawingLine => 9,
            Self::TextListStyle => 1,
            Self::TextTextStyle => 2,
            Self::TextImportedTextStyle => 3,
            Self::TextTocEntryStyle => 4,
            Self::TextTocSettings => 5,
            Self::TextCharacterStyle => 6,
            Self::TextParagraphStyle => 7,
            Self::TextDropcapStyle => 8,
            Self::ChartStyle => 1,
            Self::TableStyle => 1,
            Self::TableCellStroke => 2,
            Self::ApplicationCaptionStyle => 1,
            Self::ApplicationSvgImportStyle => 2,
        }
    }

    const fn slot(self) -> usize {
        match self {
            Self::DrawingLineStyle => 0,
            Self::DrawingShapeStyle => 1,
            Self::DrawingTextboxStyle => 2,
            Self::DrawingImageStyle => 3,
            Self::DrawingMovieStyle => 4,
            Self::DrawingLine => 5,
            Self::TextListStyle => 6,
            Self::TextTextStyle => 7,
            Self::TextImportedTextStyle => 8,
            Self::TextTocEntryStyle => 9,
            Self::TextTocSettings => 10,
            Self::TextCharacterStyle => 11,
            Self::TextParagraphStyle => 12,
            Self::TextDropcapStyle => 13,
            Self::ChartStyle => 14,
            Self::TableStyle => 15,
            Self::TableCellStroke => 16,
            Self::ApplicationCaptionStyle => 17,
            Self::ApplicationSvgImportStyle => 18,
        }
    }
}

/// Image fill technique used by the native blank-document templates.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ImageFillTechnique {
    NaturalSize,
    Stretch,
    Tile,
    ScaleToFill,
    ScaleToFit,
}

impl ImageFillTechnique {
    const fn value(self) -> u64 {
        match self {
            Self::NaturalSize => 0,
            Self::Stretch => 1,
            Self::Tile => 2,
            Self::ScaleToFill => 3,
            Self::ScaleToFit => 4,
        }
    }
}

/// Native size payload used by an image fill.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct SizeWrite {
    pub width: f32,
    pub height: f32,
}

/// Typed image fill construction. Optional fields retain proto2 presence.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ImageFillWrite {
    pub technique: ImageFillTechnique,
    pub tint: Option<ColorWrite>,
    pub fill_size: Option<SizeWrite>,
    pub interprets_untagged_image_data_as_generic: Option<bool>,
    pub reference_color: Option<ColorWrite>,
}

impl ImageFillWrite {
    /// Build a native image fill with no resource reference.
    #[must_use]
    pub const fn new(technique: ImageFillTechnique) -> Self {
        Self {
            technique,
            tint: None,
            fill_size: None,
            interprets_untagged_image_data_as_generic: None,
            reference_color: None,
        }
    }
}

/// Typed fill payload accepted by drawing preset construction.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum FillWrite<'source> {
    /// An explicit empty `TSD.FillArchive`.
    Empty,
    /// A native RGB color fill.
    Solid(ColorWrite),
    /// A native linear or radial gradient fill.
    Gradient(GradientWrite<'source>),
    /// A native image fill.
    Image(ImageFillWrite),
}

/// Shadow branch written by the Keynote blank-document template.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ShadowKind {
    Drop,
    Contact { height: f32, offset: Option<f32> },
    Curved { curve: f32 },
}

/// Typed shadow construction with all scalar presence retained.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ShadowWrite {
    empty: bool,
    drop_shadow_record: bool,
    pub color: ColorWrite,
    pub angle: f32,
    pub offset: f32,
    pub radius: i32,
    pub opacity: f32,
    pub is_enabled: bool,
    pub kind: ShadowKind,
}

impl ShadowWrite {
    /// Build an explicit empty `TSD.ShadowArchive` record.
    #[must_use]
    pub const fn empty() -> Self {
        Self {
            empty: true,
            drop_shadow_record: false,
            color: ColorWrite::new(0.0, 0.0, 0.0, 1.0, RgbSpace::Srgb),
            angle: 0.0,
            offset: 0.0,
            radius: 0,
            opacity: 0.0,
            is_enabled: false,
            kind: ShadowKind::Drop,
        }
    }

    /// Build a native drop shadow.
    #[must_use]
    pub const fn drop(
        color: ColorWrite,
        angle: f32,
        offset: f32,
        radius: i32,
        opacity: f32,
    ) -> Self {
        Self {
            empty: false,
            drop_shadow_record: false,
            color,
            angle,
            offset,
            radius,
            opacity,
            is_enabled: true,
            kind: ShadowKind::Drop,
        }
    }

    /// Build a native curved shadow.
    #[must_use]
    pub const fn curved(
        color: ColorWrite,
        offset: f32,
        radius: i32,
        opacity: f32,
        curve: f32,
    ) -> Self {
        Self {
            empty: false,
            drop_shadow_record: false,
            color,
            angle: 90.0,
            offset,
            radius,
            opacity,
            is_enabled: true,
            kind: ShadowKind::Curved { curve },
        }
    }

    /// Build a native contact shadow.
    #[must_use]
    pub const fn contact(
        color: ColorWrite,
        offset: f32,
        radius: i32,
        opacity: f32,
        height: f32,
    ) -> Self {
        Self {
            empty: false,
            drop_shadow_record: false,
            color,
            angle: 0.0,
            offset,
            radius,
            opacity,
            is_enabled: true,
            kind: ShadowKind::Contact {
                height,
                offset: None,
            },
        }
    }

    /// Build a native drop shadow with an explicitly present empty nested
    /// `DropShadowArchive` message.
    #[must_use]
    pub const fn drop_with_nested_record(
        color: ColorWrite,
        angle: f32,
        offset: f32,
        radius: i32,
        opacity: f32,
    ) -> Self {
        Self {
            empty: false,
            drop_shadow_record: true,
            color,
            angle,
            offset,
            radius,
            opacity,
            is_enabled: true,
            kind: ShadowKind::Drop,
        }
    }

    /// Build a native contact shadow while retaining its optional nested
    /// contact offset separately from the outer shadow offset.
    #[must_use]
    pub const fn contact_with_nested_offset(
        color: ColorWrite,
        offset: f32,
        radius: i32,
        opacity: f32,
        height: f32,
        nested_offset: f32,
    ) -> Self {
        Self {
            empty: false,
            drop_shadow_record: false,
            color,
            angle: 0.0,
            offset,
            radius,
            opacity,
            is_enabled: true,
            kind: ShadowKind::Contact {
                height,
                offset: Some(nested_offset),
            },
        }
    }
}

/// Finite ceilings for one freshly constructed preset extension.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ThemePresetEncodeOptions {
    pub max_output_bytes: usize,
    pub max_fields: usize,
    pub max_work_bytes: usize,
    pub max_allocations: usize,
    pub max_depth: u32,
}

impl ThemePresetEncodeOptions {
    /// Build explicit finite ceilings.
    #[must_use]
    pub const fn new(
        max_output_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        max_depth: u32,
    ) -> Self {
        Self {
            max_output_bytes,
            max_fields,
            max_work_bytes,
            max_allocations: 1,
            max_depth,
        }
    }

    /// Build a finite profile from an expected output size.
    #[must_use]
    pub fn for_output(output_bytes: usize) -> Self {
        let bytes = output_bytes.max(1);
        Self::new(
            bytes,
            bytes.saturating_mul(8).clamp(1, MAX_AUTOMATIC_LIMIT),
            bytes.saturating_mul(32).clamp(1, MAX_AUTOMATIC_LIMIT),
            16,
        )
    }

    /// Replace the logical allocation ceiling.
    #[must_use]
    pub const fn with_max_allocations(mut self, maximum: usize) -> Self {
        self.max_allocations = maximum;
        self
    }

    /// Replace the encoded output ceiling.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }
}

/// Exact bounded costs measured before a preset extension is encoded.
#[derive(Debug)]
pub struct ThemePresetEncodeRequirements<'source> {
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    allocations: usize,
    max_depth: u32,
    source: ThemePresetArchiveBuilder<'source>,
}

impl ThemePresetEncodeRequirements<'_> {
    /// Exact encoded extension length.
    #[must_use]
    pub const fn output_bytes(&self) -> usize {
        self.output_bytes
    }

    /// Exact measured field visits.
    #[must_use]
    pub const fn fields(&self) -> usize {
        self.fields
    }

    /// Exact measured work, including the output bytes.
    #[must_use]
    pub const fn work_bytes(&self) -> usize {
        self.work_bytes
    }

    /// Logical output allocations performed by [`ThemePresetArchiveBuilder::encode`].
    #[must_use]
    pub const fn allocations(&self) -> usize {
        self.allocations
    }

    /// Maximum nested field depth observed by the preflight.
    #[must_use]
    pub const fn max_depth(&self) -> u32 {
        self.max_depth
    }

    /// Convert this measurement into exact ceilings for [`Self::encode`].
    #[must_use]
    pub const fn exact_options(&self) -> ThemePresetEncodeOptions {
        ThemePresetEncodeOptions {
            max_output_bytes: self.output_bytes,
            max_fields: self.fields,
            max_work_bytes: self.work_bytes,
            max_allocations: if self.allocations == 0 {
                1
            } else {
                self.allocations
            },
            max_depth: if self.max_depth == 0 {
                1
            } else {
                self.max_depth
            },
        }
    }
}

/// Resource category rejected before output allocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ThemePresetLimit {
    OutputBytes,
    Fields,
    WorkBytes,
    Allocations,
    Nesting,
}

/// Error returned by bounded preset construction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ThemePresetError {
    InvalidFormat(&'static str),
    Allocation {
        amount: usize,
    },
    Resource {
        limit: ThemePresetLimit,
        observed: usize,
        maximum: usize,
    },
    FamilyMismatch {
        family: ThemePresetFamily,
        list: ThemePresetList,
    },
}

impl ThemePresetError {
    /// Return the rejected finite resource, if any.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<ThemePresetLimit> {
        match self {
            Self::Resource { limit, .. } => Some(*limit),
            _ => None,
        }
    }

    /// Return the requested destination allocation when reservation failed.
    #[must_use]
    pub const fn allocation_amount(&self) -> Option<usize> {
        match self {
            Self::Allocation { amount } => Some(*amount),
            _ => None,
        }
    }
}

impl fmt::Display for ThemePresetError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidFormat(message) => formatter.write_str(message),
            Self::Allocation { amount } => {
                write!(
                    formatter,
                    "theme preset allocation failed for {amount} bytes"
                )
            },
            Self::Resource {
                limit,
                observed,
                maximum,
            } => write!(
                formatter,
                "theme preset {limit:?} limit exceeded: {observed} > {maximum}"
            ),
            Self::FamilyMismatch { family, list } => {
                write!(
                    formatter,
                    "theme preset list {list:?} does not belong to {family:?}"
                )
            },
        }
    }
}

impl std::error::Error for ThemePresetError {}

/// Borrowed construction state for one complete preset extension payload.
#[derive(Debug, Clone, Copy)]
pub struct ThemePresetArchiveBuilder<'source> {
    family: ThemePresetFamily,
    references: [Option<&'source [ThemeReferenceSnapshot]>; 19],
    gradient_fills: Option<&'source [FillWrite<'source>]>,
    image_fills: Option<&'source [FillWrite<'source>]>,
    shadows: Option<&'source [ShadowWrite]>,
    unknown_fields: Option<&'source [&'source [u8]]>,
}

impl<'source> ThemePresetArchiveBuilder<'source> {
    /// Start a fresh extension for one native family.
    #[must_use]
    pub const fn new(family: ThemePresetFamily) -> Self {
        Self {
            family,
            references: [None; 19],
            gradient_fills: None,
            image_fills: None,
            shadows: None,
            unknown_fields: None,
        }
    }

    /// Set one typed repeated reference list.
    pub fn with_reference_list(
        mut self,
        list: ThemePresetList,
        references: &'source [ThemeReferenceSnapshot],
    ) -> Result<Self, ThemePresetError> {
        if list.family() != self.family {
            return Err(ThemePresetError::FamilyMismatch {
                family: self.family,
                list,
            });
        }
        self.references[list.slot()] = Some(references);
        Ok(self)
    }

    /// Set TSD gradient-fill slots. Solid fills are accepted because Pages
    /// and Numbers use this native field for repeated solid defaults.
    pub fn with_gradient_fill_presets(
        mut self,
        fills: &'source [FillWrite<'source>],
    ) -> Result<Self, ThemePresetError> {
        if self.family != ThemePresetFamily::Drawing {
            return Err(ThemePresetError::InvalidFormat(
                "gradient fills belong to the drawing preset family",
            ));
        }
        self.gradient_fills = Some(fills);
        Ok(self)
    }

    /// Set TSD image-fill slots.
    pub fn with_image_fill_presets(
        mut self,
        fills: &'source [FillWrite<'source>],
    ) -> Result<Self, ThemePresetError> {
        if self.family != ThemePresetFamily::Drawing {
            return Err(ThemePresetError::InvalidFormat(
                "image fills belong to the drawing preset family",
            ));
        }
        self.image_fills = Some(fills);
        Ok(self)
    }

    /// Set TSD shadow slots.
    pub fn with_shadow_presets(
        mut self,
        shadows: &'source [ShadowWrite],
    ) -> Result<Self, ThemePresetError> {
        if self.family != ThemePresetFamily::Drawing {
            return Err(ThemePresetError::InvalidFormat(
                "shadows belong to the drawing preset family",
            ));
        }
        self.shadows = Some(shadows);
        Ok(self)
    }

    /// Set complete, canonically framed unknown extension fields for exact byte preservation.
    #[must_use]
    pub const fn with_unknown_fields(mut self, fields: &'source [&'source [u8]]) -> Self {
        self.unknown_fields = Some(fields);
        self
    }

    /// Measure this extension before allocating its output buffer.
    pub fn encode_requirements(
        self,
        options: ThemePresetEncodeOptions,
    ) -> Result<ThemePresetEncodeRequirements<'source>, ThemePresetError> {
        validate_options(options)?;
        if options.max_allocations < 1 {
            return Err(resource(
                ThemePresetLimit::Allocations,
                1,
                options.max_allocations,
            ));
        }
        let mut accounting = Accounting::new(options);
        let mut output_bytes = 0usize;

        if let Some(fills) = self.gradient_fills {
            for fill in fills {
                let payload_bytes = measure_fill(*fill, &mut accounting, 2)?;
                let field_bytes = measure_length_field(
                    DRAWING_GRADIENT_FILL_FIELD,
                    payload_bytes,
                    &mut accounting,
                    1,
                )?;
                add_bytes(&mut output_bytes, field_bytes, options)?;
                accounting.output_prefix = output_bytes;
            }
        }
        if let Some(fills) = self.image_fills {
            for fill in fills {
                let payload_bytes = measure_fill(*fill, &mut accounting, 2)?;
                let field_bytes = measure_length_field(
                    DRAWING_IMAGE_FILL_FIELD,
                    payload_bytes,
                    &mut accounting,
                    1,
                )?;
                add_bytes(&mut output_bytes, field_bytes, options)?;
                accounting.output_prefix = output_bytes;
            }
        }
        if let Some(shadows) = self.shadows {
            for shadow in shadows {
                let payload_bytes = measure_shadow(*shadow, &mut accounting, 2)?;
                let field_bytes =
                    measure_length_field(DRAWING_SHADOW_FIELD, payload_bytes, &mut accounting, 1)?;
                add_bytes(&mut output_bytes, field_bytes, options)?;
                accounting.output_prefix = output_bytes;
            }
        }
        for (slot, references) in self.references.iter().enumerate() {
            let Some(references) = references else {
                continue;
            };
            let list = list_from_slot(slot).expect("reference slot");
            let payload_bytes =
                measure_references(references, list.field_number(), &mut accounting, 2)?;
            add_bytes(&mut output_bytes, payload_bytes, options)?;
            accounting.output_prefix = output_bytes;
        }
        if let Some(fields) = self.unknown_fields {
            for field in fields {
                measure_unknown_field(field, self.family, &mut accounting, options)?;
                add_bytes(&mut output_bytes, field.len(), options)?;
                accounting.output_prefix = output_bytes;
            }
        }
        if output_bytes > options.max_output_bytes {
            return Err(resource(
                ThemePresetLimit::OutputBytes,
                output_bytes,
                options.max_output_bytes,
            ));
        }
        accounting.charge_work(output_bytes)?;

        Ok(ThemePresetEncodeRequirements {
            output_bytes,
            fields: accounting.fields,
            work_bytes: accounting.work_bytes,
            allocations: usize::from(output_bytes != 0),
            max_depth: accounting.max_depth,
            source: self,
        })
    }

    /// Validate and encode this extension with one output allocation.
    pub fn encode(self, options: ThemePresetEncodeOptions) -> Result<Vec<u8>, ThemePresetError> {
        let requirements = self.encode_requirements(options)?;
        requirements.encode()
    }
}

impl<'source> ThemePresetEncodeRequirements<'source> {
    /// Consume an admitted source token and encode its exact borrowed builder.
    ///
    /// The token is neither cloneable nor copyable, so execution cannot be
    /// redirected to a separately constructed builder after measurement.
    pub fn encode(self) -> Result<Vec<u8>, ThemePresetError> {
        let ThemePresetEncodeRequirements {
            output_bytes,
            source,
            ..
        } = self;

        let mut output = Vec::new();
        output
            .try_reserve_exact(output_bytes)
            .map_err(|_| ThemePresetError::Allocation {
                amount: output_bytes,
            })?;
        if let Some(fills) = source.gradient_fills {
            for fill in fills {
                append_length_field(
                    &mut output,
                    DRAWING_GRADIENT_FILL_FIELD,
                    encoded_fill_len(*fill),
                    |nested| append_fill(nested, *fill),
                );
            }
        }
        if let Some(fills) = source.image_fills {
            for fill in fills {
                append_length_field(
                    &mut output,
                    DRAWING_IMAGE_FILL_FIELD,
                    encoded_fill_len(*fill),
                    |nested| append_fill(nested, *fill),
                );
            }
        }
        if let Some(shadows) = source.shadows {
            for shadow in shadows {
                append_length_field(
                    &mut output,
                    DRAWING_SHADOW_FIELD,
                    encoded_shadow_len(*shadow),
                    |nested| append_shadow(nested, *shadow),
                );
            }
        }
        for (slot, references) in source.references.iter().enumerate() {
            let Some(references) = references else {
                continue;
            };
            append_references(
                &mut output,
                list_from_slot(slot).expect("reference slot").field_number(),
                references,
            );
        }
        if let Some(fields) = source.unknown_fields {
            for field in fields {
                output.extend_from_slice(field);
            }
        }
        debug_assert_eq!(output.len(), output_bytes);
        Ok(output)
    }
}

impl Default for ThemePresetEncodeOptions {
    fn default() -> Self {
        Self::for_output(MAX_AUTOMATIC_LIMIT)
    }
}

#[derive(Debug)]
struct Accounting {
    options: ThemePresetEncodeOptions,
    fields: usize,
    work_bytes: usize,
    output_prefix: usize,
    max_depth: u32,
}

impl Accounting {
    const fn new(options: ThemePresetEncodeOptions) -> Self {
        Self {
            options,
            fields: 0,
            work_bytes: 0,
            output_prefix: 0,
            max_depth: 0,
        }
    }

    fn field(&mut self, encoded_bytes: usize, depth: u32) -> Result<(), ThemePresetError> {
        self.fields = self.fields.checked_add(1).ok_or_else(|| {
            resource(
                ThemePresetLimit::Fields,
                usize::MAX,
                self.options.max_fields,
            )
        })?;
        if self.fields > self.options.max_fields {
            return Err(resource(
                ThemePresetLimit::Fields,
                self.fields,
                self.options.max_fields,
            ));
        }
        self.charge_work(encoded_bytes)?;
        self.check_depth(depth)
    }

    fn check_depth(&mut self, depth: u32) -> Result<(), ThemePresetError> {
        self.max_depth = self.max_depth.max(depth);
        if depth > self.options.max_depth {
            return Err(resource(
                ThemePresetLimit::Nesting,
                depth as usize,
                self.options.max_depth as usize,
            ));
        }
        Ok(())
    }

    fn charge_work(&mut self, bytes: usize) -> Result<(), ThemePresetError> {
        self.work_bytes = self.work_bytes.checked_add(bytes).ok_or_else(|| {
            resource(
                ThemePresetLimit::WorkBytes,
                usize::MAX,
                self.options.max_work_bytes,
            )
        })?;
        if self.work_bytes > self.options.max_work_bytes {
            return Err(resource(
                ThemePresetLimit::WorkBytes,
                self.work_bytes,
                self.options.max_work_bytes,
            ));
        }
        Ok(())
    }
}

fn validate_options(options: ThemePresetEncodeOptions) -> Result<(), ThemePresetError> {
    if options.max_output_bytes > MAX_AUTOMATIC_LIMIT
        || options.max_fields > MAX_AUTOMATIC_LIMIT
        || options.max_work_bytes > MAX_AUTOMATIC_LIMIT
    {
        return Err(ThemePresetError::InvalidFormat(
            "theme preset limits exceed the automatic ceiling",
        ));
    }
    if options.max_depth == 0 || options.max_depth > MAX_RECURSION_LIMIT {
        return Err(resource(
            ThemePresetLimit::Nesting,
            options.max_depth as usize,
            MAX_RECURSION_LIMIT as usize,
        ));
    }
    Ok(())
}

fn measure_references(
    references: &[ThemeReferenceSnapshot],
    field_number: u32,
    accounting: &mut Accounting,
    depth: u32,
) -> Result<usize, ThemePresetError> {
    let minimum_field_bytes = length_field_len(field_number, varint_field_len(1, 0));
    let minimum_bytes = references
        .len()
        .checked_mul(minimum_field_bytes)
        .ok_or_else(overflow)?;
    ensure_minimum_output(minimum_bytes, accounting)?;
    let mut bytes = 0usize;
    for reference in references {
        if reference.identifier == 0 {
            return Err(ThemePresetError::InvalidFormat(
                "TSP.Reference.identifier must be non-zero",
            ));
        }
        let payload = measure_reference(*reference, accounting, depth)?;
        let field = measure_length_field(field_number, payload, accounting, depth - 1)?;
        add_checked(&mut bytes, field)?;
    }
    Ok(bytes)
}

fn measure_reference(
    reference: ThemeReferenceSnapshot,
    accounting: &mut Accounting,
    depth: u32,
) -> Result<usize, ThemePresetError> {
    let mut bytes = 0usize;
    let identifier = measure_varint_field(1, reference.identifier, accounting, depth)?;
    add_checked(&mut bytes, identifier)?;
    if let Some(value) = reference.deprecated_type {
        add_checked(
            &mut bytes,
            measure_varint_field(2, value as i64 as u64, accounting, depth)?,
        )?;
    }
    if let Some(value) = reference.deprecated_is_external {
        add_checked(
            &mut bytes,
            measure_varint_field(3, u64::from(value), accounting, depth)?,
        )?;
    }
    Ok(bytes)
}

fn measure_fill(
    fill: FillWrite<'_>,
    accounting: &mut Accounting,
    depth: u32,
) -> Result<usize, ThemePresetError> {
    match fill {
        FillWrite::Empty => {
            accounting.check_depth(depth)?;
            Ok(0)
        },
        FillWrite::Solid(color) => {
            let payload = measure_color(color, accounting, depth + 1)?;
            measure_length_field(FILL_COLOR_FIELD, payload, accounting, depth)
        },
        FillWrite::Gradient(gradient) => {
            let payload = measure_gradient(gradient, accounting, depth + 1)?;
            measure_length_field(FILL_GRADIENT_FIELD, payload, accounting, depth)
        },
        FillWrite::Image(image) => {
            let payload = measure_image(image, accounting, depth + 1)?;
            measure_length_field(FILL_IMAGE_FIELD, payload, accounting, depth)
        },
    }
}

fn measure_color(
    color: ColorWrite,
    accounting: &mut Accounting,
    depth: u32,
) -> Result<usize, ThemePresetError> {
    validate_color(color)?;
    let mut bytes = 0usize;
    add_checked(
        &mut bytes,
        measure_varint_field(COLOR_MODEL_FIELD, RGB_MODEL, accounting, depth)?,
    )?;
    for (field, value) in [
        (COLOR_RED_FIELD, color.red),
        (COLOR_GREEN_FIELD, color.green),
        (COLOR_BLUE_FIELD, color.blue),
        (COLOR_ALPHA_FIELD, color.alpha),
    ] {
        add_checked(
            &mut bytes,
            measure_fixed32_field(field, value, accounting, depth)?,
        )?;
    }
    add_checked(
        &mut bytes,
        measure_varint_field(
            COLOR_SPACE_FIELD,
            rgb_space_value(color.rgb_space),
            accounting,
            depth,
        )?,
    )?;
    Ok(bytes)
}

fn measure_gradient(
    gradient: GradientWrite<'_>,
    accounting: &mut Accounting,
    depth: u32,
) -> Result<usize, ThemePresetError> {
    validate_gradient_scalars(gradient)?;
    let minimum_color_bytes =
        encoded_color_len(ColorWrite::new(0.0, 0.0, 0.0, 0.0, RgbSpace::Srgb));
    let minimum_stop_bytes = length_field_len(
        GRADIENT_STOP_FIELD,
        length_field_len(STOP_COLOR_FIELD, minimum_color_bytes)
            + fixed32_field_len(STOP_FRACTION_FIELD)
            + fixed32_field_len(STOP_INFLECTION_FIELD),
    );
    let minimum_stops_bytes = gradient
        .stops
        .len()
        .checked_mul(minimum_stop_bytes)
        .ok_or_else(overflow)?;
    let minimum_gradient_bytes = varint_field_len(GRADIENT_TYPE_FIELD, 0)
        .checked_add(minimum_stops_bytes)
        .and_then(|bytes| bytes.checked_add(fixed32_field_len(GRADIENT_OPACITY_FIELD)))
        .and_then(|bytes| bytes.checked_add(varint_field_len(GRADIENT_ADVANCED_FIELD, 0)))
        .and_then(|bytes| {
            bytes.checked_add(length_field_len(
                GRADIENT_ANGLE_FIELD,
                fixed32_field_len(ANGLE_RADIANS_FIELD),
            ))
        })
        .ok_or_else(overflow)?;
    ensure_minimum_output(minimum_gradient_bytes, accounting)?;
    let mut bytes = 0usize;
    add_checked(
        &mut bytes,
        measure_varint_field(
            GRADIENT_TYPE_FIELD,
            match gradient.kind {
                GradientKind::Linear => 0,
                GradientKind::Radial => 1,
            },
            accounting,
            depth,
        )?,
    )?;
    let mut previous_fraction = 0.0;
    let mut simple_inflections = true;
    for (index, stop) in gradient.stops.iter().copied().enumerate() {
        // Admit scalar validation before inspecting this caller-owned record.
        accounting.check_depth(depth + 2)?;
        if accounting.fields >= accounting.options.max_fields {
            return Err(resource(
                ThemePresetLimit::Fields,
                accounting.fields.saturating_add(1),
                accounting.options.max_fields,
            ));
        }
        accounting.charge_work(size_of::<GradientStopWrite>())?;
        validate_gradient_stop(stop, index, &mut previous_fraction, &mut simple_inflections)?;
        let mut stop_bytes = 0usize;
        let color = measure_color(stop.color, accounting, depth + 2)?;
        add_checked(
            &mut stop_bytes,
            measure_length_field(STOP_COLOR_FIELD, color, accounting, depth + 1)?,
        )?;
        add_checked(
            &mut stop_bytes,
            measure_fixed32_field(STOP_FRACTION_FIELD, stop.fraction, accounting, depth + 1)?,
        )?;
        add_checked(
            &mut stop_bytes,
            measure_fixed32_field(
                STOP_INFLECTION_FIELD,
                stop.inflection,
                accounting,
                depth + 1,
            )?,
        )?;
        add_checked(
            &mut bytes,
            measure_length_field(GRADIENT_STOP_FIELD, stop_bytes, accounting, depth)?,
        )?;
    }
    add_checked(
        &mut bytes,
        measure_fixed32_field(GRADIENT_OPACITY_FIELD, gradient.opacity, accounting, depth)?,
    )?;
    add_checked(
        &mut bytes,
        measure_varint_field(
            GRADIENT_ADVANCED_FIELD,
            u64::from(gradient.advanced),
            accounting,
            depth,
        )?,
    )?;
    let angle = measure_fixed32_field(
        ANGLE_RADIANS_FIELD,
        gradient.angle_radians,
        accounting,
        depth + 1,
    )?;
    add_checked(
        &mut bytes,
        measure_length_field(GRADIENT_ANGLE_FIELD, angle, accounting, depth)?,
    )?;
    if !gradient.advanced
        && (gradient.kind != GradientKind::Linear
            || gradient.stops.len() != 2
            || !simple_inflections)
    {
        return Err(ThemePresetError::InvalidFormat(
            "invalid simple gradient shape",
        ));
    }
    Ok(bytes)
}

fn measure_image(
    image: ImageFillWrite,
    accounting: &mut Accounting,
    depth: u32,
) -> Result<usize, ThemePresetError> {
    let mut bytes = 0usize;
    add_checked(
        &mut bytes,
        measure_varint_field(
            IMAGE_TECHNIQUE_FIELD,
            image.technique.value(),
            accounting,
            depth,
        )?,
    )?;
    if let Some(color) = image.tint {
        let payload = measure_color(color, accounting, depth + 1)?;
        add_checked(
            &mut bytes,
            measure_length_field(IMAGE_TINT_FIELD, payload, accounting, depth)?,
        )?;
    }
    if let Some(size) = image.fill_size {
        validate_size(size)?;
        let mut payload = 0usize;
        add_checked(
            &mut payload,
            measure_fixed32_field(SIZE_WIDTH_FIELD, size.width, accounting, depth + 1)?,
        )?;
        add_checked(
            &mut payload,
            measure_fixed32_field(SIZE_HEIGHT_FIELD, size.height, accounting, depth + 1)?,
        )?;
        add_checked(
            &mut bytes,
            measure_length_field(IMAGE_SIZE_FIELD, payload, accounting, depth)?,
        )?;
    }
    if let Some(value) = image.interprets_untagged_image_data_as_generic {
        add_checked(
            &mut bytes,
            measure_varint_field(IMAGE_UNTAGGED_FIELD, u64::from(value), accounting, depth)?,
        )?;
    }
    if let Some(color) = image.reference_color {
        let payload = measure_color(color, accounting, depth + 1)?;
        add_checked(
            &mut bytes,
            measure_length_field(IMAGE_REFERENCE_COLOR_FIELD, payload, accounting, depth)?,
        )?;
    }
    Ok(bytes)
}

fn measure_shadow(
    shadow: ShadowWrite,
    accounting: &mut Accounting,
    depth: u32,
) -> Result<usize, ThemePresetError> {
    if shadow.empty {
        accounting.check_depth(depth)?;
        return Ok(0);
    }
    validate_shadow(shadow)?;
    let mut bytes = 0usize;
    let color = measure_color(shadow.color, accounting, depth + 1)?;
    add_checked(
        &mut bytes,
        measure_length_field(SHADOW_COLOR_FIELD, color, accounting, depth)?,
    )?;
    for (field, value) in [
        (SHADOW_ANGLE_FIELD, shadow.angle),
        (SHADOW_OFFSET_FIELD, shadow.offset),
        (SHADOW_OPACITY_FIELD, shadow.opacity),
    ] {
        add_checked(
            &mut bytes,
            measure_fixed32_field(field, value, accounting, depth)?,
        )?;
    }
    add_checked(
        &mut bytes,
        measure_varint_field(
            SHADOW_RADIUS_FIELD,
            shadow.radius as i64 as u64,
            accounting,
            depth,
        )?,
    )?;
    add_checked(
        &mut bytes,
        measure_varint_field(
            SHADOW_ENABLED_FIELD,
            u64::from(shadow.is_enabled),
            accounting,
            depth,
        )?,
    )?;
    let nested = match shadow.kind {
        ShadowKind::Drop if shadow.drop_shadow_record => {
            accounting.check_depth(depth + 1)?;
            Some((SHADOW_DROP_FIELD, 0usize))
        },
        ShadowKind::Drop => None,
        ShadowKind::Contact { height, offset } => {
            let mut payload = 0usize;
            add_checked(
                &mut payload,
                measure_fixed32_field(CONTACT_HEIGHT_FIELD, height, accounting, depth + 1)?,
            )?;
            if let Some(offset) = offset {
                add_checked(
                    &mut payload,
                    measure_fixed32_field(CONTACT_OFFSET_FIELD, offset, accounting, depth + 1)?,
                )?;
            }
            Some((SHADOW_CONTACT_FIELD, payload))
        },
        ShadowKind::Curved { curve } => Some((
            SHADOW_CURVED_FIELD,
            measure_fixed32_field(CURVED_CURVE_FIELD, curve, accounting, depth + 1)?,
        )),
    };
    add_checked(
        &mut bytes,
        measure_varint_field(
            SHADOW_TYPE_FIELD,
            match shadow.kind {
                ShadowKind::Drop => 0,
                ShadowKind::Contact { .. } => 1,
                ShadowKind::Curved { .. } => 2,
            },
            accounting,
            depth,
        )?,
    )?;
    if let Some(nested) = nested {
        add_checked(
            &mut bytes,
            measure_length_field(nested.0, nested.1, accounting, depth)?,
        )?;
    }
    Ok(bytes)
}

fn measure_length_field(
    number: u32,
    payload: usize,
    accounting: &mut Accounting,
    depth: u32,
) -> Result<usize, ThemePresetError> {
    let bytes = varint_len((u64::from(number) << 3) | 2)
        .checked_add(varint_len(payload as u64))
        .and_then(|value| value.checked_add(payload))
        .ok_or_else(|| {
            resource(
                ThemePresetLimit::OutputBytes,
                usize::MAX,
                accounting.options.max_output_bytes,
            )
        })?;
    accounting.field(bytes, depth)?;
    Ok(bytes)
}

fn measure_varint_field(
    number: u32,
    value: u64,
    accounting: &mut Accounting,
    depth: u32,
) -> Result<usize, ThemePresetError> {
    let bytes = varint_len(u64::from(number) << 3)
        .checked_add(varint_len(value))
        .ok_or_else(|| {
            resource(
                ThemePresetLimit::OutputBytes,
                usize::MAX,
                accounting.options.max_output_bytes,
            )
        })?;
    accounting.field(bytes, depth)?;
    Ok(bytes)
}

fn measure_fixed32_field(
    number: u32,
    value: f32,
    accounting: &mut Accounting,
    depth: u32,
) -> Result<usize, ThemePresetError> {
    let _ = value;
    let bytes = varint_len((u64::from(number) << 3) | 5)
        .checked_add(4)
        .ok_or_else(|| {
            resource(
                ThemePresetLimit::OutputBytes,
                usize::MAX,
                accounting.options.max_output_bytes,
            )
        })?;
    accounting.field(bytes, depth)?;
    Ok(bytes)
}

fn measure_unknown_field(
    source: &[u8],
    family: ThemePresetFamily,
    accounting: &mut Accounting,
    options: ThemePresetEncodeOptions,
) -> Result<(), ThemePresetError> {
    if source.is_empty() {
        return Err(ThemePresetError::InvalidFormat(
            "empty unknown preset field",
        ));
    }
    let (tag, _) = read_varint(source, 0)?;
    let number = u32::try_from(tag >> 3)
        .map_err(|_| ThemePresetError::InvalidFormat("unknown preset field overflows u32"))?;
    if known_field(family, number) {
        return Err(ThemePresetError::InvalidFormat(
            "unknown preset field uses a known field number",
        ));
    }
    let consumed = scan_unknown(source, 1, None, accounting, options)?;
    if consumed != source.len() {
        return Err(ThemePresetError::InvalidFormat(
            "unknown preset entry must contain exactly one field",
        ));
    }
    Ok(())
}

fn known_field(family: ThemePresetFamily, number: u32) -> bool {
    match family {
        ThemePresetFamily::Drawing => {
            number == 1 || number == 2 || number == 3 || (4..=9).contains(&number)
        },
        ThemePresetFamily::Text => (1..=8).contains(&number),
        ThemePresetFamily::Chart => number == 1,
        ThemePresetFamily::Table | ThemePresetFamily::Application => matches!(number, 1 | 2),
    }
}

fn scan_unknown(
    source: &[u8],
    depth: u32,
    expected_group: Option<u32>,
    accounting: &mut Accounting,
    options: ThemePresetEncodeOptions,
) -> Result<usize, ThemePresetError> {
    if depth > options.max_depth {
        return Err(resource(
            ThemePresetLimit::Nesting,
            depth as usize,
            options.max_depth as usize,
        ));
    }
    let mut cursor = 0usize;
    while cursor < source.len() {
        let start = cursor;
        let (tag, after_tag) = read_varint(source, cursor)?;
        cursor = after_tag;
        let number = u32::try_from(tag >> 3)
            .map_err(|_| ThemePresetError::InvalidFormat("unknown field number overflows u32"))?;
        let wire = u8::try_from(tag & 7)
            .map_err(|_| ThemePresetError::InvalidFormat("unknown wire type overflows u8"))?;
        if number == 0 || number > 0x1fff_ffff {
            return Err(ThemePresetError::InvalidFormat(
                "unknown field number is out of range",
            ));
        }
        if wire == 4 {
            if expected_group == Some(number) {
                let span = cursor - start;
                accounting.field(span, depth)?;
                return Ok(cursor);
            }
            return Err(ThemePresetError::InvalidFormat(
                "unexpected unknown end-group",
            ));
        }
        match wire {
            0 => cursor = read_varint(source, cursor)?.1,
            1 => cursor = cursor.checked_add(8).ok_or(overflow())?,
            2 => {
                let (length, after_length) = read_varint(source, cursor)?;
                cursor = after_length
                    .checked_add(usize::try_from(length).map_err(|_| overflow())?)
                    .ok_or(overflow())?;
            },
            3 => {
                let consumed = scan_unknown(
                    &source[cursor..],
                    depth + 1,
                    Some(number),
                    accounting,
                    options,
                )?;
                cursor = cursor.checked_add(consumed).ok_or(overflow())?;
            },
            5 => cursor = cursor.checked_add(4).ok_or(overflow())?,
            _ => return Err(ThemePresetError::InvalidFormat("invalid unknown wire type")),
        }
        if cursor > source.len() {
            return Err(ThemePresetError::InvalidFormat(
                "truncated unknown preset field",
            ));
        }
        accounting.field(cursor - start, depth)?;
    }
    if expected_group.is_some() {
        Err(ThemePresetError::InvalidFormat(
            "unterminated unknown group",
        ))
    } else {
        Ok(cursor)
    }
}

fn append_references(
    output: &mut Vec<u8>,
    field_number: u32,
    references: &[ThemeReferenceSnapshot],
) {
    for reference in references {
        let payload_len = reference_payload_len(*reference);
        append_length_header(output, field_number, payload_len);
        append_reference_payload(output, *reference);
    }
}

fn append_reference_payload(output: &mut Vec<u8>, reference: ThemeReferenceSnapshot) {
    append_varint_field(output, 1, reference.identifier);
    if let Some(value) = reference.deprecated_type {
        append_varint_field(output, 2, value as i64 as u64);
    }
    if let Some(value) = reference.deprecated_is_external {
        append_varint_field(output, 3, u64::from(value));
    }
}

fn reference_payload_len(reference: ThemeReferenceSnapshot) -> usize {
    let mut length = varint_field_len(1, reference.identifier);
    if let Some(value) = reference.deprecated_type {
        length += varint_field_len(2, value as i64 as u64);
    }
    if let Some(value) = reference.deprecated_is_external {
        length += varint_field_len(3, u64::from(value));
    }
    length
}

fn encoded_color_len(color: ColorWrite) -> usize {
    varint_field_len(COLOR_MODEL_FIELD, RGB_MODEL)
        + fixed32_field_len(COLOR_RED_FIELD)
        + fixed32_field_len(COLOR_GREEN_FIELD)
        + fixed32_field_len(COLOR_BLUE_FIELD)
        + fixed32_field_len(COLOR_ALPHA_FIELD)
        + varint_field_len(COLOR_SPACE_FIELD, rgb_space_value(color.rgb_space))
}

fn encoded_gradient_stop_len(stop: GradientStopWrite) -> usize {
    length_field_len(STOP_COLOR_FIELD, encoded_color_len(stop.color))
        + fixed32_field_len(STOP_FRACTION_FIELD)
        + fixed32_field_len(STOP_INFLECTION_FIELD)
}

fn encoded_gradient_len(gradient: GradientWrite<'_>) -> usize {
    let stops = gradient
        .stops
        .iter()
        .map(|stop| length_field_len(GRADIENT_STOP_FIELD, encoded_gradient_stop_len(*stop)))
        .sum::<usize>();
    varint_field_len(
        GRADIENT_TYPE_FIELD,
        match gradient.kind {
            GradientKind::Linear => 0,
            GradientKind::Radial => 1,
        },
    ) + stops
        + fixed32_field_len(GRADIENT_OPACITY_FIELD)
        + varint_field_len(GRADIENT_ADVANCED_FIELD, u64::from(gradient.advanced))
        + length_field_len(GRADIENT_ANGLE_FIELD, fixed32_field_len(ANGLE_RADIANS_FIELD))
}

fn encoded_size_len() -> usize {
    fixed32_field_len(SIZE_WIDTH_FIELD) + fixed32_field_len(SIZE_HEIGHT_FIELD)
}

fn encoded_image_len(image: ImageFillWrite) -> usize {
    let mut length = varint_field_len(IMAGE_TECHNIQUE_FIELD, image.technique.value());
    if let Some(color) = image.tint {
        length += length_field_len(IMAGE_TINT_FIELD, encoded_color_len(color));
    }
    if image.fill_size.is_some() {
        length += length_field_len(IMAGE_SIZE_FIELD, encoded_size_len());
    }
    if let Some(value) = image.interprets_untagged_image_data_as_generic {
        length += varint_field_len(IMAGE_UNTAGGED_FIELD, u64::from(value));
    }
    if let Some(color) = image.reference_color {
        length += length_field_len(IMAGE_REFERENCE_COLOR_FIELD, encoded_color_len(color));
    }
    length
}

fn encoded_fill_len(fill: FillWrite<'_>) -> usize {
    match fill {
        FillWrite::Empty => 0,
        FillWrite::Solid(color) => length_field_len(FILL_COLOR_FIELD, encoded_color_len(color)),
        FillWrite::Gradient(gradient) => {
            length_field_len(FILL_GRADIENT_FIELD, encoded_gradient_len(gradient))
        },
        FillWrite::Image(image) => length_field_len(FILL_IMAGE_FIELD, encoded_image_len(image)),
    }
}

fn encoded_shadow_len(shadow: ShadowWrite) -> usize {
    if shadow.empty {
        return 0;
    }
    let mut length = length_field_len(SHADOW_COLOR_FIELD, encoded_color_len(shadow.color))
        + fixed32_field_len(SHADOW_ANGLE_FIELD)
        + fixed32_field_len(SHADOW_OFFSET_FIELD)
        + varint_field_len(SHADOW_RADIUS_FIELD, shadow.radius as i64 as u64)
        + fixed32_field_len(SHADOW_OPACITY_FIELD)
        + varint_field_len(SHADOW_ENABLED_FIELD, u64::from(shadow.is_enabled));
    let nested = match shadow.kind {
        ShadowKind::Drop if shadow.drop_shadow_record => Some((SHADOW_DROP_FIELD, 0)),
        ShadowKind::Drop => None,
        ShadowKind::Contact { offset, .. } => Some((
            SHADOW_CONTACT_FIELD,
            fixed32_field_len(CONTACT_HEIGHT_FIELD)
                + offset.map_or(0, |_| fixed32_field_len(CONTACT_OFFSET_FIELD)),
        )),
        ShadowKind::Curved { .. } => {
            Some((SHADOW_CURVED_FIELD, fixed32_field_len(CURVED_CURVE_FIELD)))
        },
    };
    length += varint_field_len(
        SHADOW_TYPE_FIELD,
        match shadow.kind {
            ShadowKind::Drop => 0,
            ShadowKind::Contact { .. } => 1,
            ShadowKind::Curved { .. } => 2,
        },
    );
    length + nested.map_or(0, |(field, payload)| length_field_len(field, payload))
}

fn append_fill(output: &mut Vec<u8>, fill: FillWrite<'_>) {
    match fill {
        FillWrite::Empty => {},
        FillWrite::Solid(color) => {
            append_length_field(
                output,
                FILL_COLOR_FIELD,
                encoded_color_len(color),
                |nested| append_color(nested, color),
            );
        },
        FillWrite::Gradient(gradient) => {
            append_length_field(
                output,
                FILL_GRADIENT_FIELD,
                encoded_gradient_len(gradient),
                |nested| append_gradient(nested, gradient),
            );
        },
        FillWrite::Image(image) => {
            append_length_field(
                output,
                FILL_IMAGE_FIELD,
                encoded_image_len(image),
                |nested| append_image(nested, image),
            );
        },
    }
}

fn append_color(output: &mut Vec<u8>, color: ColorWrite) {
    append_varint_field(output, COLOR_MODEL_FIELD, RGB_MODEL);
    append_fixed32_field(output, COLOR_RED_FIELD, color.red);
    append_fixed32_field(output, COLOR_GREEN_FIELD, color.green);
    append_fixed32_field(output, COLOR_BLUE_FIELD, color.blue);
    append_fixed32_field(output, COLOR_ALPHA_FIELD, color.alpha);
    append_varint_field(output, COLOR_SPACE_FIELD, rgb_space_value(color.rgb_space));
}

fn append_gradient(output: &mut Vec<u8>, gradient: GradientWrite<'_>) {
    append_varint_field(
        output,
        GRADIENT_TYPE_FIELD,
        match gradient.kind {
            GradientKind::Linear => 0,
            GradientKind::Radial => 1,
        },
    );
    for stop in gradient.stops {
        append_length_field(
            output,
            GRADIENT_STOP_FIELD,
            encoded_gradient_stop_len(*stop),
            |nested| {
                append_length_field(
                    nested,
                    STOP_COLOR_FIELD,
                    encoded_color_len(stop.color),
                    |color| append_color(color, stop.color),
                );
                append_fixed32_field(nested, STOP_FRACTION_FIELD, stop.fraction);
                append_fixed32_field(nested, STOP_INFLECTION_FIELD, stop.inflection);
            },
        );
    }
    append_fixed32_field(output, GRADIENT_OPACITY_FIELD, gradient.opacity);
    append_varint_field(
        output,
        GRADIENT_ADVANCED_FIELD,
        u64::from(gradient.advanced),
    );
    append_length_field(
        output,
        GRADIENT_ANGLE_FIELD,
        fixed32_field_len(ANGLE_RADIANS_FIELD),
        |nested| append_fixed32_field(nested, ANGLE_RADIANS_FIELD, gradient.angle_radians),
    );
}

fn append_image(output: &mut Vec<u8>, image: ImageFillWrite) {
    append_varint_field(output, IMAGE_TECHNIQUE_FIELD, image.technique.value());
    if let Some(color) = image.tint {
        append_length_field(
            output,
            IMAGE_TINT_FIELD,
            encoded_color_len(color),
            |nested| append_color(nested, color),
        );
    }
    if let Some(size) = image.fill_size {
        append_length_field(output, IMAGE_SIZE_FIELD, encoded_size_len(), |nested| {
            append_fixed32_field(nested, SIZE_WIDTH_FIELD, size.width);
            append_fixed32_field(nested, SIZE_HEIGHT_FIELD, size.height);
        });
    }
    if let Some(value) = image.interprets_untagged_image_data_as_generic {
        append_varint_field(output, IMAGE_UNTAGGED_FIELD, u64::from(value));
    }
    if let Some(color) = image.reference_color {
        append_length_field(
            output,
            IMAGE_REFERENCE_COLOR_FIELD,
            encoded_color_len(color),
            |nested| append_color(nested, color),
        );
    }
}

fn append_shadow(output: &mut Vec<u8>, shadow: ShadowWrite) {
    if shadow.empty {
        return;
    }
    append_length_field(
        output,
        SHADOW_COLOR_FIELD,
        encoded_color_len(shadow.color),
        |nested| append_color(nested, shadow.color),
    );
    append_fixed32_field(output, SHADOW_ANGLE_FIELD, shadow.angle);
    append_fixed32_field(output, SHADOW_OFFSET_FIELD, shadow.offset);
    append_varint_field(output, SHADOW_RADIUS_FIELD, shadow.radius as i64 as u64);
    append_fixed32_field(output, SHADOW_OPACITY_FIELD, shadow.opacity);
    append_varint_field(output, SHADOW_ENABLED_FIELD, u64::from(shadow.is_enabled));
    match shadow.kind {
        ShadowKind::Drop => {
            append_varint_field(output, SHADOW_TYPE_FIELD, 0);
            if shadow.drop_shadow_record {
                append_length_header(output, SHADOW_DROP_FIELD, 0);
            }
        },
        ShadowKind::Contact { height, offset } => {
            append_varint_field(output, SHADOW_TYPE_FIELD, 1);
            append_length_field(
                output,
                SHADOW_CONTACT_FIELD,
                fixed32_field_len(CONTACT_HEIGHT_FIELD)
                    + offset.map_or(0, |_| fixed32_field_len(CONTACT_OFFSET_FIELD)),
                |nested| {
                    append_fixed32_field(nested, CONTACT_HEIGHT_FIELD, height);
                    if let Some(offset) = offset {
                        append_fixed32_field(nested, CONTACT_OFFSET_FIELD, offset);
                    }
                },
            );
        },
        ShadowKind::Curved { curve } => {
            append_varint_field(output, SHADOW_TYPE_FIELD, 2);
            append_length_field(
                output,
                SHADOW_CURVED_FIELD,
                fixed32_field_len(CURVED_CURVE_FIELD),
                |nested| append_fixed32_field(nested, CURVED_CURVE_FIELD, curve),
            );
        },
    }
}

fn append_length_field(
    output: &mut Vec<u8>,
    number: u32,
    payload_len: usize,
    append_payload: impl FnOnce(&mut Vec<u8>),
) {
    append_length_header(output, number, payload_len);
    append_payload(output);
}

fn append_length_header(output: &mut Vec<u8>, number: u32, payload_len: usize) {
    append_varint((u64::from(number) << 3) | 2, output);
    append_varint(payload_len as u64, output);
}

fn append_varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
    append_varint(u64::from(number) << 3, output);
    append_varint(value, output);
}

fn append_fixed32_field(output: &mut Vec<u8>, number: u32, value: f32) {
    append_varint((u64::from(number) << 3) | 5, output);
    output.extend_from_slice(&value.to_bits().to_le_bytes());
}

fn append_varint(mut value: u64, output: &mut Vec<u8>) {
    while value >= 0x80 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn validate_color(color: ColorWrite) -> Result<(), ThemePresetError> {
    for value in [color.red, color.green, color.blue, color.alpha] {
        if !value.is_finite() || !(0.0..=1.0).contains(&value) {
            return Err(ThemePresetError::InvalidFormat(
                "theme preset color component is outside [0,1]",
            ));
        }
    }
    Ok(())
}

fn validate_gradient_scalars(gradient: GradientWrite<'_>) -> Result<(), ThemePresetError> {
    if !gradient.opacity.is_finite()
        || !(0.0..=1.0).contains(&gradient.opacity)
        || !gradient.angle_radians.is_finite()
        || !(0.0..core::f32::consts::TAU).contains(&gradient.angle_radians)
        || gradient.stops.len() < 2
    {
        return Err(ThemePresetError::InvalidFormat("invalid gradient scalar"));
    }
    Ok(())
}

fn validate_gradient_stop(
    stop: GradientStopWrite,
    index: usize,
    previous_fraction: &mut f32,
    simple_inflections: &mut bool,
) -> Result<(), ThemePresetError> {
    validate_color(stop.color)?;
    if !stop.fraction.is_finite()
        || !(0.0..=1.0).contains(&stop.fraction)
        || !stop.inflection.is_finite()
        || !(0.0..=1.0).contains(&stop.inflection)
        || (index != 0 && stop.fraction < *previous_fraction)
    {
        return Err(ThemePresetError::InvalidFormat(
            "invalid gradient stop scalar",
        ));
    }
    *previous_fraction = stop.fraction;
    *simple_inflections &= stop.inflection.to_bits() == 0.5f32.to_bits();
    Ok(())
}

fn validate_size(size: SizeWrite) -> Result<(), ThemePresetError> {
    if size.width.is_finite() && size.height.is_finite() {
        Ok(())
    } else {
        Err(ThemePresetError::InvalidFormat(
            "image fill size is not finite",
        ))
    }
}

fn validate_shadow(shadow: ShadowWrite) -> Result<(), ThemePresetError> {
    validate_color(shadow.color)?;
    for value in [shadow.angle, shadow.offset, shadow.opacity] {
        if !value.is_finite() {
            return Err(ThemePresetError::InvalidFormat(
                "shadow scalar is not finite",
            ));
        }
    }
    if !(0.0..=1.0).contains(&shadow.opacity) {
        return Err(ThemePresetError::InvalidFormat(
            "shadow opacity is outside [0,1]",
        ));
    }
    match shadow.kind {
        ShadowKind::Drop => {},
        ShadowKind::Contact { height, offset } => {
            if !height.is_finite() || offset.is_some_and(|offset| !offset.is_finite()) {
                return Err(ThemePresetError::InvalidFormat(
                    "contact shadow scalar is not finite",
                ));
            }
        },
        ShadowKind::Curved { curve } => {
            if !curve.is_finite() {
                return Err(ThemePresetError::InvalidFormat(
                    "curved shadow scalar is not finite",
                ));
            }
        },
    }
    Ok(())
}

fn rgb_space_value(space: RgbSpace) -> u64 {
    match space {
        RgbSpace::Srgb => SRGB_SPACE,
        RgbSpace::DisplayP3 => P3_SPACE,
    }
}

fn ensure_minimum_output(
    minimum_bytes: usize,
    accounting: &Accounting,
) -> Result<(), ThemePresetError> {
    let aggregate = accounting
        .output_prefix
        .checked_add(minimum_bytes)
        .ok_or_else(overflow)?;
    if aggregate > accounting.options.max_output_bytes {
        return Err(resource(
            ThemePresetLimit::OutputBytes,
            aggregate,
            accounting.options.max_output_bytes,
        ));
    }
    Ok(())
}

fn add_bytes(
    total: &mut usize,
    amount: usize,
    options: ThemePresetEncodeOptions,
) -> Result<(), ThemePresetError> {
    *total = total.checked_add(amount).ok_or_else(overflow)?;
    if *total > options.max_output_bytes {
        return Err(resource(
            ThemePresetLimit::OutputBytes,
            *total,
            options.max_output_bytes,
        ));
    }
    Ok(())
}

fn add_checked(total: &mut usize, amount: usize) -> Result<(), ThemePresetError> {
    *total = total.checked_add(amount).ok_or_else(overflow)?;
    Ok(())
}

fn varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

fn varint_field_len(number: u32, value: u64) -> usize {
    varint_len(u64::from(number) << 3) + varint_len(value)
}

fn fixed32_field_len(number: u32) -> usize {
    varint_len((u64::from(number) << 3) | 5) + 4
}

fn length_field_len(number: u32, payload_len: usize) -> usize {
    varint_len((u64::from(number) << 3) | 2) + varint_len(payload_len as u64) + payload_len
}

fn read_varint(source: &[u8], start: usize) -> Result<(u64, usize), ThemePresetError> {
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *source
            .get(start + index)
            .ok_or(ThemePresetError::InvalidFormat("truncated preset varint"))?;
        if index > 0 && byte == 0 {
            return Err(ThemePresetError::InvalidFormat(
                "non-canonical preset varint",
            ));
        }
        if index == 9 && byte > 1 {
            return Err(ThemePresetError::InvalidFormat("preset varint is too long"));
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            return Ok((value, start + index + 1));
        }
    }
    Err(ThemePresetError::InvalidFormat("preset varint is too long"))
}

fn overflow() -> ThemePresetError {
    ThemePresetError::Resource {
        limit: ThemePresetLimit::OutputBytes,
        observed: usize::MAX,
        maximum: usize::MAX,
    }
}

fn resource(limit: ThemePresetLimit, observed: usize, maximum: usize) -> ThemePresetError {
    ThemePresetError::Resource {
        limit,
        observed,
        maximum,
    }
}

fn list_from_slot(slot: usize) -> Option<ThemePresetList> {
    Some(match slot {
        0 => ThemePresetList::DrawingLineStyle,
        1 => ThemePresetList::DrawingShapeStyle,
        2 => ThemePresetList::DrawingTextboxStyle,
        3 => ThemePresetList::DrawingImageStyle,
        4 => ThemePresetList::DrawingMovieStyle,
        5 => ThemePresetList::DrawingLine,
        6 => ThemePresetList::TextListStyle,
        7 => ThemePresetList::TextTextStyle,
        8 => ThemePresetList::TextImportedTextStyle,
        9 => ThemePresetList::TextTocEntryStyle,
        10 => ThemePresetList::TextTocSettings,
        11 => ThemePresetList::TextCharacterStyle,
        12 => ThemePresetList::TextParagraphStyle,
        13 => ThemePresetList::TextDropcapStyle,
        14 => ThemePresetList::ChartStyle,
        15 => ThemePresetList::TableStyle,
        16 => ThemePresetList::TableCellStroke,
        17 => ThemePresetList::ApplicationCaptionStyle,
        18 => ThemePresetList::ApplicationSvgImportStyle,
        _ => return None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{tsa, tsch, tsd, tsp, tst, tswp};
    use prost::Message as _;

    fn reference(identifier: u64) -> ThemeReferenceSnapshot {
        ThemeReferenceSnapshot::new(identifier)
    }

    #[test]
    fn zero_reference_identifiers_are_rejected_before_preset_output() {
        let error = ThemePresetArchiveBuilder::new(ThemePresetFamily::Text)
            .with_reference_list(ThemePresetList::TextListStyle, &[reference(0)])
            .expect("reference list family")
            .encode(ThemePresetEncodeOptions::for_output(512))
            .expect_err("zero native identifiers must be rejected");
        assert!(matches!(
            error,
            ThemePresetError::InvalidFormat(message) if message.contains("non-zero")
        ));
    }

    fn black() -> ColorWrite {
        ColorWrite::new(0.0, 0.0, 0.0, 1.0, RgbSpace::Srgb)
    }

    fn white() -> ColorWrite {
        ColorWrite::new(1.0, 1.0, 1.0, 1.0, RgbSpace::Srgb)
    }

    #[test]
    fn all_families_and_typed_lists_match_native_prost_bytes() {
        let list = [reference(11), reference(12)];
        let text = ThemePresetArchiveBuilder::new(ThemePresetFamily::Text)
            .with_reference_list(ThemePresetList::TextListStyle, &list)
            .expect("text list")
            .with_reference_list(ThemePresetList::TextParagraphStyle, &list)
            .expect("paragraph list")
            .encode(ThemePresetEncodeOptions::for_output(512))
            .expect("text extension");
        let expected = tswp::ThemePresetsArchive {
            list_style_presets: list
                .iter()
                .map(|value| tsp::Reference {
                    identifier: value.identifier,
                    ..Default::default()
                })
                .collect(),
            paragraph_style_presets: list
                .iter()
                .map(|value| tsp::Reference {
                    identifier: value.identifier,
                    ..Default::default()
                })
                .collect(),
            ..Default::default()
        }
        .encode_to_vec();
        assert_eq!(text, expected);

        for (family, list_kind, expected) in [
            (
                ThemePresetFamily::Chart,
                ThemePresetList::ChartStyle,
                tsch::ChartPresetsArchive {
                    chart_presets: vec![tsp::Reference {
                        identifier: 11,
                        ..Default::default()
                    }],
                }
                .encode_to_vec(),
            ),
            (
                ThemePresetFamily::Table,
                ThemePresetList::TableStyle,
                tst::ThemePresetsArchive {
                    table_style_presets: vec![tsp::Reference {
                        identifier: 11,
                        ..Default::default()
                    }],
                    ..Default::default()
                }
                .encode_to_vec(),
            ),
            (
                ThemePresetFamily::Application,
                ThemePresetList::ApplicationCaptionStyle,
                tsa::ThemePresetsArchive {
                    caption_style_presets: vec![tsp::Reference {
                        identifier: 11,
                        ..Default::default()
                    }],
                    ..Default::default()
                }
                .encode_to_vec(),
            ),
        ] {
            let actual = ThemePresetArchiveBuilder::new(family)
                .with_reference_list(list_kind, &list[..1])
                .expect("family list")
                .encode(ThemePresetEncodeOptions::for_output(128))
                .expect("extension");
            assert_eq!(actual, expected);
        }
    }

    #[test]
    fn drawing_fill_and_shadow_records_match_native_blank_template_bytes() {
        let stops = [
            GradientStopWrite {
                color: white(),
                fraction: 0.0,
                inflection: 0.5,
            },
            GradientStopWrite {
                color: black(),
                fraction: 1.0,
                inflection: 0.5,
            },
        ];
        let gradient = FillWrite::Gradient(GradientWrite {
            kind: GradientKind::Linear,
            stops: &stops,
            opacity: 1.0,
            advanced: false,
            angle_radians: 3.0 * core::f32::consts::FRAC_PI_2,
        });
        let image = FillWrite::Image(ImageFillWrite {
            technique: ImageFillTechnique::Tile,
            tint: Some(white()),
            fill_size: Some(SizeWrite {
                width: 300.0,
                height: 300.0,
            }),
            interprets_untagged_image_data_as_generic: Some(true),
            reference_color: Some(white()),
        });
        let fills = [gradient, image];
        let shadows = [ShadowWrite::drop(black(), 90.0, 2.0, 5, 0.5)];
        let drawing_refs = [reference(99)];
        let actual = ThemePresetArchiveBuilder::new(ThemePresetFamily::Drawing)
            .with_gradient_fill_presets(&fills)
            .expect("gradient fills")
            .with_shadow_presets(&shadows)
            .expect("shadows")
            .with_reference_list(ThemePresetList::DrawingLine, &drawing_refs)
            .expect("drawing references")
            .encode(ThemePresetEncodeOptions::for_output(4096))
            .expect("drawing extension");
        let expected = tsd::ThemePresetsArchive {
            gradient_fill_presets: vec![
                tsd::FillArchive {
                    gradient: Some(tsd::GradientArchive {
                        r#type: Some(0),
                        stops: vec![
                            tsd::gradient_archive::GradientStop {
                                color: Some(tsp::Color {
                                    model: 1,
                                    r: Some(1.0),
                                    g: Some(1.0),
                                    b: Some(1.0),
                                    rgbspace: Some(1),
                                    a: Some(1.0),
                                    ..Default::default()
                                }),
                                fraction: Some(0.0),
                                inflection: Some(0.5),
                            },
                            tsd::gradient_archive::GradientStop {
                                color: Some(tsp::Color {
                                    model: 1,
                                    r: Some(0.0),
                                    g: Some(0.0),
                                    b: Some(0.0),
                                    rgbspace: Some(1),
                                    a: Some(1.0),
                                    ..Default::default()
                                }),
                                fraction: Some(1.0),
                                inflection: Some(0.5),
                            },
                        ],
                        opacity: Some(1.0),
                        advanced_gradient: Some(false),
                        anglegradient: Some(tsd::AngleGradientArchive {
                            gradientangle: Some(3.0 * core::f32::consts::FRAC_PI_2),
                        }),
                        ..Default::default()
                    }),
                    ..Default::default()
                },
                tsd::FillArchive {
                    image: Some(tsd::ImageFillArchive {
                        technique: Some(2),
                        tint: Some(tsp::Color {
                            model: 1,
                            r: Some(1.0),
                            g: Some(1.0),
                            b: Some(1.0),
                            rgbspace: Some(1),
                            a: Some(1.0),
                            ..Default::default()
                        }),
                        fillsize: Some(tsp::Size {
                            width: 300.0,
                            height: 300.0,
                        }),
                        interprets_untagged_image_data_as_generic: Some(true),
                        referencecolor: Some(tsp::Color {
                            model: 1,
                            r: Some(1.0),
                            g: Some(1.0),
                            b: Some(1.0),
                            rgbspace: Some(1),
                            a: Some(1.0),
                            ..Default::default()
                        }),
                        ..Default::default()
                    }),
                    ..Default::default()
                },
            ],
            shadow_presets: vec![tsd::ShadowArchive {
                color: Some(tsp::Color {
                    model: 1,
                    r: Some(0.0),
                    g: Some(0.0),
                    b: Some(0.0),
                    rgbspace: Some(1),
                    a: Some(1.0),
                    ..Default::default()
                }),
                angle: Some(90.0),
                offset: Some(2.0),
                radius: Some(5),
                opacity: Some(0.5),
                is_enabled: Some(true),
                r#type: Some(0),
                ..Default::default()
            }],
            drawing_line_style_presets: vec![tsp::Reference {
                identifier: 99,
                ..Default::default()
            }],
            ..Default::default()
        }
        .encode_to_vec();
        assert_eq!(actual, expected);
    }

    #[test]
    fn empty_and_contact_shadow_records_match_native_bytes() {
        let shadows = [
            ShadowWrite::empty(),
            ShadowWrite::contact(black(), 9.418_75, 30, 0.75, 0.25),
            ShadowWrite::drop_with_nested_record(black(), 90.0, 2.0, 5, 0.5),
        ];
        let actual = ThemePresetArchiveBuilder::new(ThemePresetFamily::Drawing)
            .with_shadow_presets(&shadows)
            .expect("shadows")
            .encode(ThemePresetEncodeOptions::for_output(512))
            .expect("shadow extension");
        let expected = tsd::ThemePresetsArchive {
            shadow_presets: vec![
                tsd::ShadowArchive::default(),
                tsd::ShadowArchive {
                    color: Some(tsp::Color {
                        model: 1,
                        r: Some(0.0),
                        g: Some(0.0),
                        b: Some(0.0),
                        rgbspace: Some(1),
                        a: Some(1.0),
                        ..Default::default()
                    }),
                    angle: Some(0.0),
                    offset: Some(9.418_75),
                    radius: Some(30),
                    opacity: Some(0.75),
                    is_enabled: Some(true),
                    r#type: Some(1),
                    contact_shadow: Some(tsd::ContactShadowArchive {
                        height: Some(0.25),
                        ..Default::default()
                    }),
                    ..Default::default()
                },
                tsd::ShadowArchive {
                    color: Some(tsp::Color {
                        model: 1,
                        r: Some(0.0),
                        g: Some(0.0),
                        b: Some(0.0),
                        rgbspace: Some(1),
                        a: Some(1.0),
                        ..Default::default()
                    }),
                    angle: Some(90.0),
                    offset: Some(2.0),
                    radius: Some(5),
                    opacity: Some(0.5),
                    is_enabled: Some(true),
                    r#type: Some(0),
                    drop_shadow: Some(tsd::DropShadowArchive {}),
                    ..Default::default()
                },
            ],
            ..Default::default()
        }
        .encode_to_vec();
        assert_eq!(actual, expected);
    }

    #[test]
    fn family_mismatch_and_finite_limits_preempt_encoding() {
        let refs = [reference(1)];
        let error = ThemePresetArchiveBuilder::new(ThemePresetFamily::Text)
            .with_reference_list(ThemePresetList::ChartStyle, &refs)
            .expect_err("mismatched family");
        assert!(matches!(error, ThemePresetError::FamilyMismatch { .. }));

        let error = ThemePresetArchiveBuilder::new(ThemePresetFamily::Text)
            .with_reference_list(ThemePresetList::TextListStyle, &refs)
            .expect("list")
            .encode(ThemePresetEncodeOptions::new(1, 128, 4096, 16))
            .expect_err("output ceiling");
        assert_eq!(error.resource_limit(), Some(ThemePresetLimit::OutputBytes));

        let color = ColorWrite::new(f32::NAN, 0.0, 0.0, 1.0, RgbSpace::Srgb);
        let fills = [FillWrite::Solid(color)];
        let error = ThemePresetArchiveBuilder::new(ThemePresetFamily::Drawing)
            .with_gradient_fill_presets(&fills)
            .expect("fills")
            .encode(ThemePresetEncodeOptions::for_output(1024))
            .expect_err("non-finite color");
        assert!(matches!(error, ThemePresetError::InvalidFormat(_)));

        let empty_fill = [FillWrite::Empty];
        let error = ThemePresetArchiveBuilder::new(ThemePresetFamily::Drawing)
            .with_gradient_fill_presets(&empty_fill)
            .expect("empty fill")
            .encode(ThemePresetEncodeOptions::new(64, 16, 256, 1))
            .expect_err("empty fill still has nested message depth");
        assert_eq!(error.resource_limit(), Some(ThemePresetLimit::Nesting));

        let empty_shadow = [ShadowWrite::empty()];
        let error = ThemePresetArchiveBuilder::new(ThemePresetFamily::Drawing)
            .with_shadow_presets(&empty_shadow)
            .expect("empty shadow")
            .encode(ThemePresetEncodeOptions::new(64, 16, 256, 1))
            .expect_err("empty shadow still has nested message depth");
        assert_eq!(error.resource_limit(), Some(ThemePresetLimit::Nesting));
    }

    #[test]
    fn later_gradient_checks_remaining_output_before_stop_validation() {
        let valid = GradientStopWrite {
            color: black(),
            fraction: 0.0,
            inflection: 0.5,
        };
        let mut invalid_stops = [valid; 2];
        invalid_stops[1].color = ColorWrite::new(f32::NAN, 0.0, 0.0, 1.0, RgbSpace::Srgb);
        let fills = [
            FillWrite::Solid(black()),
            FillWrite::Gradient(GradientWrite {
                kind: GradientKind::Linear,
                stops: &invalid_stops,
                opacity: 1.0,
                advanced: true,
                angle_radians: 0.0,
            }),
        ];
        let error = ThemePresetArchiveBuilder::new(ThemePresetFamily::Drawing)
            .with_gradient_fill_presets(&fills)
            .expect("fills")
            .encode(ThemePresetEncodeOptions::new(110, 1000, 10000, 16))
            .expect_err("aggregate output must reject before the invalid stop");
        assert_eq!(error.resource_limit(), Some(ThemePresetLimit::OutputBytes));
    }

    #[test]
    fn gradient_stop_validation_is_interleaved_with_budget_admission() {
        let valid = GradientStopWrite {
            color: black(),
            fraction: 0.0,
            inflection: 0.5,
        };
        let mut stops = [valid; 65];
        stops[64].color = ColorWrite::new(f32::NAN, 0.0, 0.0, 1.0, RgbSpace::Srgb);
        let fills = [FillWrite::Gradient(GradientWrite {
            kind: GradientKind::Linear,
            stops: &stops,
            opacity: 1.0,
            advanced: true,
            angle_radians: 0.0,
        })];
        let error = ThemePresetArchiveBuilder::new(ThemePresetFamily::Drawing)
            .with_gradient_fill_presets(&fills)
            .expect("fills")
            .encode(ThemePresetEncodeOptions::new(4096, 1, 4096, 16))
            .expect_err("field admission must preempt late stop validation");
        assert_eq!(error.resource_limit(), Some(ThemePresetLimit::Fields));
    }

    #[test]
    fn encode_requirements_are_exact_before_output_allocation() {
        let references = [reference(7), reference(9)];
        let builder = ThemePresetArchiveBuilder::new(ThemePresetFamily::Text)
            .with_reference_list(ThemePresetList::TextListStyle, &references)
            .expect("text reference list");
        let options = ThemePresetEncodeOptions::for_output(1024);
        let requirements = builder
            .encode_requirements(options)
            .expect("bounded measurement");
        let expected_length = requirements.output_bytes();
        let exact_options = requirements.exact_options();
        let below_options = exact_options.with_max_output_bytes(expected_length - 1);
        let public_output = builder
            .encode(exact_options)
            .expect("public encode wrapper");
        let output = requirements.encode().expect("exact measured limits");
        assert_eq!(output.len(), expected_length);
        assert_eq!(public_output, output);

        let error = builder
            .encode_requirements(below_options)
            .expect_err("one output byte below the measured extension must fail");
        assert_eq!(error.resource_limit(), Some(ThemePresetLimit::OutputBytes));
    }
}
