//! Caller-selected resource policy for the chart raw codecs.
//!
//! The chart reader, ChartEx reader, chart-style readers, authoring reader and
//! chart writer all use one policy type.  The defaults preserve the historical
//! ceilings of each codec; a caller may only narrow those ceilings.  Keeping
//! the policy at the chart module boundary lets package owners pass their
//! retained chart budget through every raw operation without introducing a
//! family of silently-defaulted helper functions.

use crate::chart::authoring::Spec;
use crate::chart::data::{NumericData, StringData, TitleText};
use crate::chart::model::Chart;
use crate::chart::series::{DataLabel, DataLabels};
use crate::chart::style::{ColorDocument, Document};
use crate::{Error, LimitKind, Result};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;

/// Count the bytes quick-xml will retain for one normalized attribute value.
///
/// `Attribute::decoded_and_normalized_value` allocates when an entity or an
/// XML-whitespace normalization is present.  Calling it merely to discover
/// the length would therefore allocate before the caller can enforce its
/// policy.  This scanner follows quick-xml's XML 1.0 normalization rules for
/// predefined and numeric character references, but retains no decoded
/// string.  The actual decoder is still called by the owner after this
/// admission succeeds, preserving its normal error and value semantics.
pub(crate) fn decoded_attribute_value_bytes(
    raw: &[u8],
    decoder: Decoder,
) -> std::result::Result<usize, String> {
    let decoded = decoder.decode(raw).map_err(|error| error.to_string())?;
    let value = decoded.as_ref();
    let bytes = value.as_bytes();
    let mut length = 0usize;
    let mut index = 0usize;
    while index < bytes.len() {
        let amount = match bytes[index] {
            b'&' => {
                let end = bytes
                    .get(index + 1..)
                    .and_then(|rest| rest.iter().position(|byte| *byte == b';'))
                    .map(|offset| index + 1 + offset)
                    .ok_or_else(|| "unterminated XML entity reference".to_owned())?;
                let entity = value
                    .get(index + 1..end)
                    .ok_or_else(|| "invalid XML entity reference".to_owned())?;
                let amount = if let Some(number) = entity.strip_prefix('#') {
                    character_reference_bytes(number)
                        .ok_or_else(|| "invalid XML character reference".to_owned())?
                } else {
                    quick_xml::escape::resolve_predefined_entity(entity)
                        .map(str::len)
                        .ok_or_else(|| "unrecognized XML entity reference".to_owned())?
                };
                index = end + 1;
                amount
            },
            b'\t' => {
                index += 1;
                1
            },
            b'\r' => {
                index += 1;
                if bytes.get(index) == Some(&b'\n') {
                    index += 1;
                }
                1
            },
            b'\n' => {
                index += 1;
                1
            },
            _ => {
                index += 1;
                1
            },
        };
        length = length
            .checked_add(amount)
            .ok_or_else(|| "decoded XML attribute value length overflow".to_owned())?;
    }
    Ok(length)
}

/// Count the UTF-8 bytes retained by quick-xml for one text event.
///
/// Text admission must use the decoded value rather than the lexical source
/// length.  In particular, line-ending normalization can shrink a source
/// event, while an XML entity is delivered as a separate `GeneralRef` event.
/// Decoding and normalization are therefore measured before the owner asks
/// quick-xml for an owned `Cow`; this keeps a rejected event from first
/// materializing a large normalized string.
pub(crate) fn decoded_text_value_bytes(
    raw: &[u8],
    decoder: Decoder,
    version: XmlVersion,
) -> std::result::Result<usize, String> {
    let decoded = decoder.decode(raw).map_err(|error| error.to_string())?;
    normalized_text_bytes(decoded.as_ref(), version)
}

/// Count the bytes retained by a decoder for a CDATA event. CDATA has no XML
/// entity expansion and the style codec intentionally does not normalize its
/// line endings, so this is kept separate from ordinary text admission.
pub(crate) fn decoded_cdata_value_bytes(
    raw: &[u8],
    decoder: Decoder,
) -> std::result::Result<usize, String> {
    decoder
        .decode(raw)
        .map(|decoded| decoded.len())
        .map_err(|error| error.to_string())
}

/// Count the bytes produced by `String::from_utf8_lossy` without first
/// allocating its replacement string.
pub(crate) fn lossy_utf8_value_bytes(raw: &[u8]) -> usize {
    let mut remaining = raw;
    let mut length = 0usize;
    while !remaining.is_empty() {
        match std::str::from_utf8(remaining) {
            Ok(value) => {
                length = length.saturating_add(value.len());
                break;
            },
            Err(error) => {
                length = length.saturating_add(error.valid_up_to());
                // `from_utf8_lossy` emits one U+FFFD for each malformed
                // sequence, including a truncated sequence at EOF.
                length = length.saturating_add(char::REPLACEMENT_CHARACTER.len_utf8());
                let invalid = error
                    .error_len()
                    .unwrap_or_else(|| remaining.len().saturating_sub(error.valid_up_to()));
                let skip = error.valid_up_to().saturating_add(invalid);
                remaining = remaining.get(skip..).unwrap_or_default();
            },
        }
    }
    length
}

fn normalized_text_bytes(value: &str, version: XmlVersion) -> std::result::Result<usize, String> {
    let mut length = 0usize;
    let mut chars = value.chars().peekable();
    while let Some(character) = chars.next() {
        let amount = if character == '\r' {
            if chars.peek() == Some(&'\n') {
                chars.next();
            } else if matches!(version, XmlVersion::Explicit1_1)
                && chars.peek() == Some(&'\u{0085}')
            {
                chars.next();
            }
            1
        } else if matches!(version, XmlVersion::Explicit1_1)
            && matches!(character, '\u{0085}' | '\u{2028}')
        {
            1
        } else {
            character.len_utf8()
        };
        length = length
            .checked_add(amount)
            .ok_or_else(|| "decoded XML text value length overflow".to_owned())?;
    }
    Ok(length)
}

fn character_reference_bytes(number: &str) -> Option<usize> {
    let (number, radix) = number
        .strip_prefix('x')
        .map_or((number, 10), |number| (number, 16));
    if number.is_empty() || matches!(number.as_bytes().first(), Some(b'+' | b'-')) {
        return None;
    }
    let codepoint = u32::from_str_radix(number, radix).ok()?;
    if codepoint == 0 {
        return None;
    }
    char::from_u32(codepoint).map(char::len_utf8)
}

/// Historical maximum accepted by the MCE-backed classic chart reader.
const MAX_INPUT_BYTES: usize = 256 * 1024 * 1024;
/// Historical maximum emitted by the MCE processor and chart output policy.
const MAX_OUTPUT_BYTES: usize = 512 * 1024 * 1024;
/// Historical classic-chart stream depth and node ceilings.
const MAX_XML_NODES: usize = 1_000_000;
const MAX_XML_DEPTH: usize = 256;
// The ordinary chart reader historically accepted very wide producer tags;
// keep that behavior while still bounding the raw event before semantic
// projection. MCE branches apply their own tighter 4096 ceiling.
const MAX_XML_ATTRIBUTES: usize = 64 * 1024;
const MAX_ATTRIBUTE_VALUE_BYTES: usize = 16 * 1024 * 1024;

/// Historical authoring data ceilings.
const MAX_SERIES: usize = 255;
const MAX_CATEGORIES: usize = 32_000;
const MAX_POINTS: usize = 32_000;
const MAX_TOTAL_POINTS: usize = 256_000;
const MAX_TEXT_BYTES: usize = 32_767;

/// Historical ChartEx parser ceilings.
const MAX_CHARTEX_BYTES: usize = 16 * 1024 * 1024;
const MAX_CHARTEX_NODES: usize = 400_000;
const MAX_CHARTEX_DEPTH: usize = 64;

/// Historical chart-style parser ceilings.
const MAX_STYLE_BYTES: usize = 16 * 1024 * 1024;
const MAX_STYLE_NODES: usize = 100_000;
const MAX_STYLE_DEPTH: usize = 96;
const MAX_STYLE_ATTRIBUTES: usize = 64;
const MAX_STYLE_STRING_BYTES: usize = 4 * 1024 * 1024;
const MAX_STYLE_MODIFIERS: usize = 256;
const MAX_STYLE_COLORS: usize = 4_096;
const MAX_STYLE_VARIATIONS: usize = 4_096;
const MAX_STYLE_TRANSFORMS: usize = 4_096;

/// Bounded resources used by every public chart raw codec entry point.
///
/// Values are public so a host can derive a policy from its own retained
/// package limits.  Each operation validates the policy before parsing or
/// serializing; zero and values above the historical hard ceiling produce
/// [`Error::InvalidLimit`].
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[must_use]
pub struct Limits {
    /// Maximum source bytes accepted by a classic chart reader.
    pub max_input_bytes: usize,
    /// Maximum bytes emitted by a chart writer.
    pub max_output_bytes: usize,
    /// Maximum XML nodes admitted by the classic chart stream reader.
    pub max_xml_nodes: usize,
    /// Maximum XML nesting depth admitted by the classic chart stream reader.
    pub max_xml_depth: usize,
    /// Maximum attributes admitted on one classic-chart element.
    pub max_xml_attributes: usize,
    /// Maximum decoded attribute value retained by a chart reader.
    pub max_attribute_value_bytes: usize,
    /// Maximum chart series retained by authoring projections.
    pub max_series: usize,
    /// Maximum category points retained by authoring projections.
    pub max_categories: usize,
    /// Maximum points in one series or one ChartEx data level.
    pub max_points: usize,
    /// Maximum total numeric points retained by an authoring projection.
    pub max_total_points: usize,
    /// Maximum bytes in one authored text value.
    pub max_text_bytes: usize,
    /// Maximum ChartEx source bytes.
    pub max_chartex_bytes: usize,
    /// Maximum ChartEx XML nodes.
    pub max_chartex_nodes: usize,
    /// Maximum ChartEx XML depth.
    pub max_chartex_depth: usize,
    /// Maximum chart-style source bytes.
    pub max_style_bytes: usize,
    /// Maximum chart-style XML nodes.
    pub max_style_nodes: usize,
    /// Maximum chart-style XML depth.
    pub max_style_depth: usize,
    /// Maximum chart-style attributes on one element.
    pub max_style_attributes: usize,
    /// Maximum retained chart-style string bytes.
    pub max_style_string_bytes: usize,
    /// Maximum modifiers in one chart-style entry.
    pub max_style_modifiers: usize,
    /// Maximum colors in one chart color-style document.
    pub max_style_colors: usize,
    /// Maximum variations in one chart color-style document.
    pub max_style_variations: usize,
    /// Maximum transforms in one chart color-style variation.
    pub max_style_transforms: usize,
}

impl Limits {
    /// The historical policy for all chart raw codecs.
    pub const DEFAULT: Self = Self {
        max_input_bytes: MAX_INPUT_BYTES,
        max_output_bytes: MAX_OUTPUT_BYTES,
        max_xml_nodes: MAX_XML_NODES,
        max_xml_depth: MAX_XML_DEPTH,
        max_xml_attributes: MAX_XML_ATTRIBUTES,
        max_attribute_value_bytes: MAX_ATTRIBUTE_VALUE_BYTES,
        max_series: MAX_SERIES,
        max_categories: MAX_CATEGORIES,
        max_points: MAX_POINTS,
        max_total_points: MAX_TOTAL_POINTS,
        max_text_bytes: MAX_TEXT_BYTES,
        max_chartex_bytes: MAX_CHARTEX_BYTES,
        max_chartex_nodes: MAX_CHARTEX_NODES,
        max_chartex_depth: MAX_CHARTEX_DEPTH,
        max_style_bytes: MAX_STYLE_BYTES,
        max_style_nodes: MAX_STYLE_NODES,
        max_style_depth: MAX_STYLE_DEPTH,
        max_style_attributes: MAX_STYLE_ATTRIBUTES,
        max_style_string_bytes: MAX_STYLE_STRING_BYTES,
        max_style_modifiers: MAX_STYLE_MODIFIERS,
        max_style_colors: MAX_STYLE_COLORS,
        max_style_variations: MAX_STYLE_VARIATIONS,
        max_style_transforms: MAX_STYLE_TRANSFORMS,
    };

    /// Construct the historical policy with a narrower source/output byte
    /// budget.  Other dimensions retain their historical values.
    #[must_use]
    pub const fn new(max_xml_bytes: usize) -> Self {
        let mut limits = Self::DEFAULT;
        limits.max_input_bytes = max_xml_bytes;
        limits.max_output_bytes = max_xml_bytes;
        limits.max_chartex_bytes = if max_xml_bytes < MAX_CHARTEX_BYTES {
            max_xml_bytes
        } else {
            MAX_CHARTEX_BYTES
        };
        limits.max_style_bytes = if max_xml_bytes < MAX_STYLE_BYTES {
            max_xml_bytes
        } else {
            MAX_STYLE_BYTES
        };
        limits
    }

    /// Validate the caller policy against the fixed codec ceilings.
    pub(crate) fn validate(self) -> Result<Self> {
        check(LimitKind::InputBytes, self.max_input_bytes, MAX_INPUT_BYTES)?;
        check(
            LimitKind::OutputBytes,
            self.max_output_bytes,
            MAX_OUTPUT_BYTES,
        )?;
        check(LimitKind::XmlNodes, self.max_xml_nodes, MAX_XML_NODES)?;
        check(LimitKind::XmlDepth, self.max_xml_depth, MAX_XML_DEPTH)?;
        check(
            LimitKind::XmlAttributes,
            self.max_xml_attributes,
            MAX_XML_ATTRIBUTES,
        )?;
        check(
            LimitKind::AttributeValueBytes,
            self.max_attribute_value_bytes,
            MAX_ATTRIBUTE_VALUE_BYTES,
        )?;
        check(LimitKind::ChartSeries, self.max_series, MAX_SERIES)?;
        check(
            LimitKind::ChartCategories,
            self.max_categories,
            MAX_CATEGORIES,
        )?;
        check(LimitKind::ChartPoints, self.max_points, MAX_POINTS)?;
        check(
            LimitKind::ChartPoints,
            self.max_total_points,
            MAX_TOTAL_POINTS,
        )?;
        check(
            LimitKind::ChartTextBytes,
            self.max_text_bytes,
            MAX_TEXT_BYTES,
        )?;
        check(
            LimitKind::ChartexPartBytes,
            self.max_chartex_bytes,
            MAX_CHARTEX_BYTES,
        )?;
        check(
            LimitKind::ChartexXmlNodes,
            self.max_chartex_nodes,
            MAX_CHARTEX_NODES,
        )?;
        check(
            LimitKind::ChartexXmlDepth,
            self.max_chartex_depth,
            MAX_CHARTEX_DEPTH,
        )?;
        check(LimitKind::XmlBytes, self.max_style_bytes, MAX_STYLE_BYTES)?;
        check(LimitKind::XmlNodes, self.max_style_nodes, MAX_STYLE_NODES)?;
        check(LimitKind::XmlDepth, self.max_style_depth, MAX_STYLE_DEPTH)?;
        check(
            LimitKind::XmlAttributes,
            self.max_style_attributes,
            MAX_STYLE_ATTRIBUTES,
        )?;
        check(
            LimitKind::ChartStyleStrings,
            self.max_style_string_bytes,
            MAX_STYLE_STRING_BYTES,
        )?;
        check(
            LimitKind::ChartStyleModifiers,
            self.max_style_modifiers,
            MAX_STYLE_MODIFIERS,
        )?;
        check(
            LimitKind::ChartStyleColors,
            self.max_style_colors,
            MAX_STYLE_COLORS,
        )?;
        check(
            LimitKind::ChartStyleVariations,
            self.max_style_variations,
            MAX_STYLE_VARIATIONS,
        )?;
        check(
            LimitKind::ChartStyleTransforms,
            self.max_style_transforms,
            MAX_STYLE_TRANSFORMS,
        )?;
        Ok(self)
    }

    /// Check the resource counts already materialized in a typed chart.
    pub(crate) fn validate_chart(self, chart: &Chart) -> Result<()> {
        let _ = self.validate()?;
        check_title_text(self, chart.title.as_ref())?;
        if let Some(formats) = chart.pivot_formats.as_ref() {
            check_count(LimitKind::ChartPoints, formats.len(), self.max_points)?;
            let mut indexes = std::collections::HashSet::new();
            for format in formats {
                if !indexes.insert(format.index) {
                    return Err(Error::Invalid(
                        "chart contains duplicate pivot-format index".into(),
                    ));
                }
                if let Some(label) = format.data_label.as_ref() {
                    check_data_label(self, label)?;
                }
            }
        }
        if let Some(language) = &chart.language {
            check_count(
                LimitKind::ChartTextBytes,
                language.len(),
                self.max_text_bytes,
            )?;
        }
        if let Some(source) = &chart.pivot_source {
            check_count(
                LimitKind::ChartTextBytes,
                source.name.len(),
                self.max_text_bytes,
            )?;
        }
        if let Some(external_data) = &chart.external_data {
            if let Some(relationship_id) = &external_data.relationship_id {
                check_count(
                    LimitKind::AttributeValueBytes,
                    relationship_id.len(),
                    self.max_attribute_value_bytes,
                )?;
            }
        }
        if let Some(user_shapes) = &chart.user_shapes {
            if let Some(relationship_id) = &user_shapes.relationship_id {
                check_count(
                    LimitKind::AttributeValueBytes,
                    relationship_id.len(),
                    self.max_attribute_value_bytes,
                )?;
            }
        }
        if let Some(header_footer) = chart
            .print_settings
            .as_ref()
            .and_then(|settings| settings.header_footer.as_ref())
        {
            for value in [
                header_footer.odd_header.as_ref(),
                header_footer.odd_footer.as_ref(),
                header_footer.even_header.as_ref(),
                header_footer.even_footer.as_ref(),
                header_footer.first_header.as_ref(),
                header_footer.first_footer.as_ref(),
            ]
            .into_iter()
            .flatten()
            {
                check_count(LimitKind::ChartTextBytes, value.len(), self.max_text_bytes)?;
            }
        }
        for axis in &chart.plot_area.axes {
            check_title_text(self, axis.common().title.as_ref())?;
            if let Some(number_format) = &axis.common().number_format {
                check_count(
                    LimitKind::ChartTextBytes,
                    number_format.format_code.len(),
                    self.max_text_bytes,
                )?;
            }
        }
        let mut series = 0usize;
        let mut total_points = 0usize;
        for group in &chart.plot_area.type_groups {
            if let Some(labels) = &group.common().data_labels {
                check_data_labels(self, labels)?;
            }
            series = series
                .checked_add(group.common().series.len())
                .ok_or_else(|| Error::limit(LimitKind::ChartSeries, usize::MAX, self.max_series))?;
            if series > self.max_series {
                return Err(Error::limit(
                    LimitKind::ChartSeries,
                    series,
                    self.max_series,
                ));
            }
            for value in &group.common().series {
                check_title_text(self, value.title.as_ref())?;
                if let Some(categories) = &value.categories {
                    check_count(
                        LimitKind::ChartCategories,
                        categories.values.len(),
                        self.max_categories,
                    )?;
                    check_string_data(self, categories)?;
                }
                for data in [
                    value.values.as_ref(),
                    value.x_values.as_ref(),
                    value.y_values.as_ref(),
                    value.bubble_sizes.as_ref(),
                ]
                .into_iter()
                .flatten()
                {
                    check_numeric_data(self, data)?;
                    check_count(LimitKind::ChartPoints, data.values.len(), self.max_points)?;
                    total_points =
                        total_points.checked_add(data.values.len()).ok_or_else(|| {
                            Error::limit(LimitKind::ChartPoints, usize::MAX, self.max_total_points)
                        })?;
                    if total_points > self.max_total_points {
                        return Err(Error::limit(
                            LimitKind::ChartPoints,
                            total_points,
                            self.max_total_points,
                        ));
                    }
                }
                for error_bar in &value.error_bars {
                    for data in [
                        error_bar.plus_values.as_ref(),
                        error_bar.minus_values.as_ref(),
                    ]
                    .into_iter()
                    .flatten()
                    {
                        check_numeric_data(self, data)?;
                        check_count(LimitKind::ChartPoints, data.values.len(), self.max_points)?;
                        total_points =
                            total_points.checked_add(data.values.len()).ok_or_else(|| {
                                Error::limit(
                                    LimitKind::ChartPoints,
                                    usize::MAX,
                                    self.max_total_points,
                                )
                            })?;
                        if total_points > self.max_total_points {
                            return Err(Error::limit(
                                LimitKind::ChartPoints,
                                total_points,
                                self.max_total_points,
                            ));
                        }
                    }
                }
                if let Some(labels) = &value.data_labels {
                    check_data_labels(self, labels)?;
                }
                for trendline in &value.trendlines {
                    if let Some(name) = &trendline.name {
                        check_count(LimitKind::ChartTextBytes, name.len(), self.max_text_bytes)?;
                    }
                    check_title_text(self, trendline.label.as_ref())?;
                    if let Some(number_format) = &trendline.label_number_format {
                        check_count(
                            LimitKind::ChartTextBytes,
                            number_format.format_code.len(),
                            self.max_text_bytes,
                        )?;
                    }
                }
            }
        }
        Ok(())
    }

    /// Check the resource counts retained in an authoring specification.
    pub(crate) fn validate_spec(self, spec: &Spec) -> Result<()> {
        let _ = self.validate()?;
        check_count(LimitKind::ChartSeries, spec.series().len(), self.max_series)?;
        check_count(
            LimitKind::ChartCategories,
            spec.categories().len(),
            self.max_categories,
        )?;
        check_count(
            LimitKind::ChartPoints,
            spec.appearance().colors.len(),
            self.max_points,
        )?;
        if let Some(subtotals) = &spec.appearance().subtotals {
            check_count(LimitKind::ChartPoints, subtotals.len(), self.max_points)?;
        }
        let mut total = 0usize;
        for series in spec.series() {
            check_count(LimitKind::ChartPoints, series.values.len(), self.max_points)?;
            total = total.checked_add(series.values.len()).ok_or_else(|| {
                Error::limit(LimitKind::ChartPoints, usize::MAX, self.max_total_points)
            })?;
            if total > self.max_total_points {
                return Err(Error::limit(
                    LimitKind::ChartPoints,
                    total,
                    self.max_total_points,
                ));
            }
            check_count(
                LimitKind::ChartTextBytes,
                series.name.len(),
                self.max_text_bytes,
            )?;
        }
        for text in [
            spec.title(),
            spec.category_axis_title(),
            spec.value_axis_title(),
        ]
        .into_iter()
        .flatten()
        {
            check_count(LimitKind::ChartTextBytes, text.len(), self.max_text_bytes)?;
        }
        for text in spec.categories() {
            check_count(LimitKind::ChartTextBytes, text.len(), self.max_text_bytes)?;
        }
        for (_, values, categories) in spec.series_references() {
            check_count(LimitKind::ChartTextBytes, values.len(), self.max_text_bytes)?;
            if let Some(categories) = categories {
                check_count(
                    LimitKind::ChartTextBytes,
                    categories.len(),
                    self.max_text_bytes,
                )?;
            }
        }
        Ok(())
    }

    /// Check an authoring spec before the classic-chart projection allocates
    /// its generated X/Y/bubble arrays.  The spec stores one value vector per
    /// series, while scatter and bubble projections materialize two or three
    /// numeric vectors for each source point.
    pub(crate) fn validate_spec_for_chart(self, spec: &Spec) -> Result<()> {
        self.validate_spec(spec)?;
        let multiplier = match spec.variant().kind {
            crate::chart::authoring::Kind::Scatter => 2,
            crate::chart::authoring::Kind::Bubble => 3,
            _ => 1,
        };
        let source_points = spec.series().iter().try_fold(0usize, |total, series| {
            total.checked_add(series.values.len()).ok_or_else(|| {
                Error::limit(LimitKind::ChartPoints, usize::MAX, self.max_total_points)
            })
        })?;
        let projected = source_points.checked_mul(multiplier).ok_or_else(|| {
            Error::limit(LimitKind::ChartPoints, usize::MAX, self.max_total_points)
        })?;
        check_count(LimitKind::ChartPoints, projected, self.max_total_points)
    }

    /// Check the bounded summaries retained by a chart-style document.
    pub(crate) fn validate_style_document(self, document: &Document) -> Result<()> {
        let _ = self.validate()?;
        let mut strings = 0usize;
        for entry in &document.info.entries {
            check_count(
                LimitKind::ChartStyleModifiers,
                entry.modifiers.len(),
                self.max_style_modifiers,
            )?;
            for modifier in &entry.modifiers {
                check_count(LimitKind::ChartStyleModifierBytes, modifier.len(), 256)?;
                strings = strings.checked_add(modifier.len()).ok_or_else(|| {
                    Error::limit(
                        LimitKind::ChartStyleStrings,
                        usize::MAX,
                        self.max_style_string_bytes,
                    )
                })?;
            }
            for reference in [
                &entry.line_reference,
                &entry.fill_reference,
                &entry.effect_reference,
            ] {
                validate_style_reference(self, reference, &mut strings)?;
            }
            validate_font_reference(self, &entry.font_reference, &mut strings)?;
            for payload in [
                entry.shape_properties.as_ref(),
                entry.default_run_properties.as_ref(),
                entry.body_properties.as_ref(),
            ]
            .into_iter()
            .flatten()
            {
                check_count(
                    LimitKind::XmlNodes,
                    payload.child_elements,
                    self.max_style_nodes,
                )?;
                check_count(
                    LimitKind::XmlAttributes,
                    payload.attributes,
                    self.max_style_attributes,
                )?;
            }
        }
        check_count(
            LimitKind::ChartStyleStrings,
            strings,
            self.max_style_string_bytes,
        )
    }

    /// Check the bounded summaries retained by a chart color-style document.
    pub(crate) fn validate_color_document(self, document: &ColorDocument) -> Result<()> {
        let _ = self.validate()?;
        check_count(
            LimitKind::ChartStyleColors,
            document.info.colors.len(),
            self.max_style_colors,
        )?;
        check_count(
            LimitKind::ChartStyleVariations,
            document.info.variations.len(),
            self.max_style_variations,
        )?;
        let mut transforms = 0usize;
        let mut strings = document.info.method.len();
        for color in &document.info.colors {
            if let Some(value) = &color.value {
                strings = strings.saturating_add(value.len());
            }
            for (name, value) in &color.components {
                strings = strings
                    .saturating_add(name.len())
                    .saturating_add(value.len());
            }
            check_style_transforms(self, &color.transforms, &mut transforms, &mut strings)?;
        }
        for variation in &document.info.variations {
            check_style_transforms(self, &variation.transforms, &mut transforms, &mut strings)?;
        }
        check_count(
            LimitKind::ChartStyleStrings,
            strings,
            self.max_style_string_bytes,
        )
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

fn check(kind: LimitKind, value: usize, maximum: usize) -> Result<()> {
    if value == 0 || value > maximum {
        return Err(Error::invalid_limit(kind, value, maximum));
    }
    Ok(())
}

fn check_count(kind: LimitKind, observed: usize, maximum: usize) -> Result<()> {
    if observed > maximum {
        return Err(Error::limit(kind, observed, maximum));
    }
    Ok(())
}

fn check_title_text(limits: Limits, title: Option<&TitleText>) -> Result<()> {
    let Some(title) = title else {
        return Ok(());
    };
    let text = match title {
        TitleText::Literal(rich) => &rich.text,
        TitleText::Reference(reference) => &reference.formula,
    };
    check_count(LimitKind::ChartTextBytes, text.len(), limits.max_text_bytes)
}

fn check_data_label(limits: Limits, label: &DataLabel) -> Result<()> {
    check_title_text(limits, label.text.as_ref())?;
    if let Some(number_format) = &label.number_format {
        check_count(
            LimitKind::ChartTextBytes,
            number_format.format_code.len(),
            limits.max_text_bytes,
        )?;
    }
    if let Some(separator) = &label.separator {
        check_count(
            LimitKind::ChartTextBytes,
            separator.len(),
            limits.max_text_bytes,
        )?;
    }
    Ok(())
}

fn check_data_labels(limits: Limits, labels: &DataLabels) -> Result<()> {
    check_count(
        LimitKind::ChartPoints,
        labels.labels.len(),
        limits.max_points,
    )?;
    let mut indexes = std::collections::HashSet::new();
    for label in &labels.labels {
        if !indexes.insert(label.index) {
            return Err(Error::Invalid(
                "chart data labels contain duplicate point index".into(),
            ));
        }
        check_data_label(limits, label)?;
    }
    if let Some(number_format) = &labels.number_format {
        check_count(
            LimitKind::ChartTextBytes,
            number_format.format_code.len(),
            limits.max_text_bytes,
        )?;
    }
    if let Some(separator) = &labels.separator {
        check_count(
            LimitKind::ChartTextBytes,
            separator.len(),
            limits.max_text_bytes,
        )?;
    }
    Ok(())
}

fn check_string_data(limits: Limits, data: &StringData) -> Result<()> {
    if let Some(source) = &data.source_ref {
        check_count(
            LimitKind::ChartTextBytes,
            source.formula.len(),
            limits.max_text_bytes,
        )?;
    }
    for value in &data.values {
        check_count(
            LimitKind::ChartTextBytes,
            value.len(),
            limits.max_text_bytes,
        )?;
    }
    Ok(())
}

fn check_numeric_data(limits: Limits, data: &NumericData) -> Result<()> {
    if let Some(source) = &data.source_ref {
        check_count(
            LimitKind::ChartTextBytes,
            source.formula.len(),
            limits.max_text_bytes,
        )?;
    }
    if let Some(format) = &data.format_code {
        check_count(
            LimitKind::ChartTextBytes,
            format.len(),
            limits.max_text_bytes,
        )?;
    }
    Ok(())
}

fn validate_style_reference(
    limits: Limits,
    reference: &crate::chart::style::Reference,
    strings: &mut usize,
) -> Result<()> {
    if let Some(color) = &reference.color {
        check_style_color(limits, color, strings)?;
    }
    if let Some(color) = &reference.style_color {
        if let Some(raw) = &color.raw {
            *strings = strings.saturating_add(raw.len());
        }
        let mut transforms = 0;
        check_style_transforms(limits, &color.transforms, &mut transforms, strings)?;
    }
    Ok(())
}

fn validate_font_reference(
    limits: Limits,
    reference: &crate::chart::style::FontReference,
    strings: &mut usize,
) -> Result<()> {
    if let Some(color) = &reference.color {
        check_style_color(limits, color, strings)?;
    }
    if let Some(color) = &reference.style_color {
        if let Some(raw) = &color.raw {
            *strings = strings.saturating_add(raw.len());
        }
        let mut transforms = 0;
        check_style_transforms(limits, &color.transforms, &mut transforms, strings)?;
    }
    Ok(())
}

fn check_style_color(
    limits: Limits,
    color: &crate::chart::style::Color,
    strings: &mut usize,
) -> Result<()> {
    if let Some(value) = &color.value {
        *strings = strings.saturating_add(value.len());
    }
    for (name, value) in &color.components {
        *strings = strings
            .saturating_add(name.len())
            .saturating_add(value.len());
    }
    let mut transforms = 0;
    check_style_transforms(limits, &color.transforms, &mut transforms, strings)
}

fn check_style_transforms(
    limits: Limits,
    transforms: &[crate::chart::style::Transform],
    total: &mut usize,
    strings: &mut usize,
) -> Result<()> {
    *total = total.checked_add(transforms.len()).ok_or_else(|| {
        Error::limit(
            LimitKind::ChartStyleTransforms,
            usize::MAX,
            limits.max_style_transforms,
        )
    })?;
    if *total > limits.max_style_transforms {
        return Err(Error::limit(
            LimitKind::ChartStyleTransforms,
            *total,
            limits.max_style_transforms,
        ));
    }
    for transform in transforms {
        if let Some(value) = &transform.value {
            *strings = strings.saturating_add(value.len());
        }
    }
    Ok(())
}
