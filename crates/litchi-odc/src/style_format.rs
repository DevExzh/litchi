//! Bounded, source-preserving chart formatting properties in `styles.xml`.
//!
//! ODF chart formatting is attached to `style:chart-properties`, rather than
//! to the chart content tree.  This module intentionally covers the small
//! property surface needed by the ODC audit gap.  It parses style names and
//! chart-property start tags only; edits replace one checked attribute span at
//! a time, leaving comments, unknown elements, prefixes, and lexical choices
//! outside that span untouched.

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::{
    collections::{BTreeMap, BTreeSet},
    ops::Range,
};

const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const CHART: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
const XLINK: &[u8] = b"http://www.w3.org/1999/xlink";
const MAX_PROPERTIES_PER_STYLE: usize = 32;

/// A source-backed chart formatting attribute from ODF 1.4.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum ChartStyleProperty {
    /// `chart:logarithmic`.
    Logarithmic,
    /// `chart:minor-logarithmic`.
    MinorLogarithmic,
    /// `chart:tick-marks-major-inner`.
    TickMarksMajorInner,
    /// `chart:tick-marks-major-outer`.
    TickMarksMajorOuter,
    /// `chart:tick-marks-minor-inner`.
    TickMarksMinorInner,
    /// `chart:tick-marks-minor-outer`.
    TickMarksMinorOuter,
    /// `chart:symbol-type`.
    SymbolType,
    /// `chart:symbol-name`.
    SymbolName,
    /// `chart:error-category`.
    ErrorCategory,
    /// `chart:error-percentage`.
    ErrorPercentage,
    /// `chart:error-margin`.
    ErrorMargin,
    /// `chart:error-lower-limit`.
    ErrorLowerLimit,
    /// `chart:error-upper-limit`.
    ErrorUpperLimit,
    /// `chart:error-upper-indicator`.
    ErrorUpperIndicator,
    /// `chart:error-lower-indicator`.
    ErrorLowerIndicator,
    /// `chart:regression-type`.
    RegressionType,
    /// `chart:regression-max-degree`.
    RegressionMaxDegree,
    /// `chart:regression-force-intercept`.
    RegressionForceIntercept,
    /// `chart:regression-intercept-value`.
    RegressionInterceptValue,
    /// `chart:regression-name`.
    RegressionName,
    /// `chart:regression-period`.
    RegressionPeriod,
    /// `chart:regression-moving-type`.
    RegressionMovingType,
}

/// ODF chart symbol rendering mode.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum ChartSymbolType {
    None,
    Automatic,
    NamedSymbol,
    Image,
}

/// ODF named chart symbol.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum ChartSymbolName {
    Square,
    Diamond,
    ArrowDown,
    ArrowUp,
    ArrowRight,
    ArrowLeft,
    BowTie,
    Hourglass,
    Circle,
    Star,
    X,
    Plus,
    Asterisk,
    HorizontalBar,
    VerticalBar,
}

/// ODF chart error-bar category.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum ChartErrorCategory {
    None,
    Variance,
    StandardDeviation,
    Percentage,
    ErrorMargin,
    Constant,
    StandardError,
    CellRange,
}

/// ODF chart regression model.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum ChartRegressionType {
    None,
    Linear,
    Logarithmic,
    MovingAverage,
    Exponential,
    Power,
    Polynomial,
}

/// ODF moving-average regression window placement.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum ChartRegressionMovingType {
    Prior,
    Central,
    AveragedAbscissa,
}

/// A typed lexical value accepted for one chart formatting property.
#[derive(Clone, Debug)]
pub enum ChartStyleValue {
    Boolean(bool),
    Decimal(f64),
    PositiveInteger(u32),
    Text(String),
    SymbolType(ChartSymbolType),
    SymbolName(ChartSymbolName),
    ErrorCategory(ChartErrorCategory),
    RegressionType(ChartRegressionType),
    RegressionMovingType(ChartRegressionMovingType),
}

impl PartialEq for ChartStyleValue {
    fn eq(&self, other: &Self) -> bool {
        match (self, other) {
            (Self::Boolean(left), Self::Boolean(right)) => left == right,
            (Self::Decimal(left), Self::Decimal(right)) => {
                left == right || (left.is_nan() && right.is_nan())
            },
            (Self::PositiveInteger(left), Self::PositiveInteger(right)) => left == right,
            (Self::Text(left), Self::Text(right)) => left == right,
            (Self::SymbolType(left), Self::SymbolType(right)) => left == right,
            (Self::SymbolName(left), Self::SymbolName(right)) => left == right,
            (Self::ErrorCategory(left), Self::ErrorCategory(right)) => left == right,
            (Self::RegressionType(left), Self::RegressionType(right)) => left == right,
            (Self::RegressionMovingType(left), Self::RegressionMovingType(right)) => left == right,
            _ => false,
        }
    }
}

/// One source-bound chart formatting property change.
#[derive(Clone, Debug, PartialEq)]
pub struct ChartStylePropertyChange {
    style_name: String,
    property: ChartStyleProperty,
    before: Option<ChartStyleValue>,
    after: Option<ChartStyleValue>,
}

impl ChartStylePropertyChange {
    pub(crate) fn new(
        style_name: String,
        property: ChartStyleProperty,
        before: Option<ChartStyleValue>,
        after: Option<ChartStyleValue>,
    ) -> Self {
        Self {
            style_name,
            property,
            before,
            after,
        }
    }

    #[must_use]
    pub fn style_name(&self) -> &str {
        &self.style_name
    }

    #[must_use]
    pub const fn property(&self) -> ChartStyleProperty {
        self.property
    }

    #[must_use]
    pub fn before(&self) -> Option<&ChartStyleValue> {
        self.before.as_ref()
    }

    #[must_use]
    pub fn after(&self) -> Option<&ChartStyleValue> {
        self.after.as_ref()
    }

    pub(crate) fn inverse(&self) -> Self {
        Self {
            style_name: self.style_name.clone(),
            property: self.property,
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ValueKind {
    Boolean,
    Decimal,
    PositiveInteger,
    Text,
    SymbolType,
    SymbolName,
    ErrorCategory,
    RegressionType,
    RegressionMovingType,
}

#[derive(Clone, Debug)]
struct ParsedStyles {
    styles: Vec<StyleRecord>,
}

#[derive(Clone, Debug)]
struct StyleRecord {
    name: String,
    properties: Option<ChartProperties>,
}

#[derive(Clone, Debug)]
struct ChartProperties {
    tag: Range<usize>,
    chart_prefix: Option<String>,
    attributes: Vec<AttributeRecord>,
    has_symbol_image: bool,
}

#[derive(Clone, Debug)]
struct AttributeRecord {
    property: ChartStyleProperty,
    value: String,
    value_span: Range<usize>,
    attribute_span: Range<usize>,
}

/// Read one typed property from one chart-family style.
pub(crate) fn read_property(
    xml: &str,
    style_name: &str,
    property: ChartStyleProperty,
    limits: crate::Limits,
) -> Result<Option<ChartStyleValue>> {
    validate_style_name(style_name, limits)?;
    let parsed = parse(xml, limits)?;
    let Some(style) = parsed.styles.iter().find(|style| style.name == style_name) else {
        return Ok(None);
    };
    let Some(properties) = style.properties.as_ref() else {
        return Ok(None);
    };
    properties
        .attributes
        .iter()
        .find(|attribute| attribute.property == property)
        .map(|attribute| decode_value(property, &attribute.value, limits))
        .transpose()
}

pub(crate) fn has_style(xml: &str, style_name: &str, limits: crate::Limits) -> Result<bool> {
    validate_style_name(style_name, limits)?;
    Ok(parse(xml, limits)?
        .styles
        .iter()
        .any(|style| style.name == style_name))
}

/// Determine whether a property patch is owned by the typed chart-formatting
/// surface for this exact source part.
///
/// A missing style selector, missing `style:chart-properties` element, or an
/// insertion that would require inventing a chart namespace prefix is an
/// unmodelled structural change.  Those cases deliberately return `Ok(false)`
/// so package merge can use its lossless whole-part path.  Malformed XML,
/// invalid typed values, and stale source values remain errors and are never
/// downgraded to ownership misses.
/// Determine whether a typed property patch can reproduce the target without
/// taking ownership of modeled child structure in `styles.xml`.
///
/// The typed property surface does not publish structural children such as
/// `chart:symbol-image`.  Both parts are parsed before classification so a
/// valid structural difference becomes a whole-part fallback while malformed
/// XML remains an error.
pub(crate) fn owns_changes_between(
    before: &str,
    after: &str,
    changes: &[ChartStylePropertyChange],
    limits: crate::Limits,
) -> Result<bool> {
    if changes.is_empty() {
        return Ok(false);
    }
    let before_parsed = parse(before, limits)?;
    let after_parsed = parse(after, limits)?;
    if !modeled_structure_matches(&before_parsed, &after_parsed) {
        return Ok(false);
    }
    owns_changes_in_parsed(&before_parsed, changes, limits)
}

fn owns_changes_in_parsed(
    parsed: &ParsedStyles,
    changes: &[ChartStylePropertyChange],
    limits: crate::Limits,
) -> Result<bool> {
    for change in changes {
        let Some(style) = parsed
            .styles
            .iter()
            .find(|style| style.name == change.style_name)
        else {
            return Ok(false);
        };
        let Some(properties) = style.properties.as_ref() else {
            return Ok(false);
        };
        validate_optional_value(change.property, change.after.as_ref(), limits)?;
        let attribute = properties
            .attributes
            .iter()
            .find(|attribute| attribute.property == change.property);
        let current = attribute
            .map(|attribute| decode_value(change.property, &attribute.value, limits))
            .transpose()?;
        if current != change.before {
            return Err(Error::InvalidFormat(format!(
                "ODC chart style property '{}' is stale",
                property_name(change.property)
            )));
        }
        if current == change.after {
            continue;
        }
        if attribute.is_none() && change.after.is_some() && properties.chart_prefix.is_none() {
            return Ok(false);
        }
    }
    Ok(true)
}

/// Check the destination prerequisites before a transfer starts staging edits.
///
/// This accepts a destination whose values are either the patch's source
/// values or its already-applied target values.  It only rejects structural
/// destinations that cannot host the typed operation.
pub(crate) fn can_host_changes(
    xml: &str,
    changes: &[ChartStylePropertyChange],
    limits: crate::Limits,
) -> Result<bool> {
    if changes.is_empty() {
        return Ok(true);
    }
    let parsed = parse(xml, limits)?;
    for change in changes {
        let Some(style) = parsed
            .styles
            .iter()
            .find(|style| style.name == change.style_name)
        else {
            return Ok(false);
        };
        let Some(properties) = style.properties.as_ref() else {
            return Ok(false);
        };
        validate_optional_value(change.property, change.after.as_ref(), limits)?;
        let attribute = properties
            .attributes
            .iter()
            .find(|attribute| attribute.property == change.property);
        if let Some(attribute) = attribute {
            decode_value(change.property, &attribute.value, limits)?;
        } else if change.after.is_some() && properties.chart_prefix.is_none() {
            return Ok(false);
        }
    }
    Ok(true)
}

fn modeled_structure_matches(left: &ParsedStyles, right: &ParsedStyles) -> bool {
    left.styles.len() == right.styles.len()
        && left.styles.iter().zip(&right.styles).all(|(left, right)| {
            left.name == right.name
                && match (&left.properties, &right.properties) {
                    (Some(left), Some(right)) => left.has_symbol_image == right.has_symbol_image,
                    (None, None) => true,
                    _ => false,
                }
        })
}

/// Apply checked property changes to a source styles part.
pub(crate) fn apply_changes(
    xml: &str,
    changes: &[ChartStylePropertyChange],
    limits: crate::Limits,
) -> Result<String> {
    if changes.is_empty() {
        return Ok(xml.to_owned());
    }
    let replacements = checked_replacements(xml, changes, limits)?;
    if replacements.is_empty() {
        return Ok(xml.to_owned());
    }
    let output = apply_replacements(xml, replacements, limits)?;
    parse(&output, limits)?;
    for change in changes {
        let actual = read_property(&output, &change.style_name, change.property, limits)?;
        if actual != change.after {
            return Err(invalid("ODC chart style property failed typed readback"));
        }
    }
    Ok(output)
}

#[derive(Debug)]
pub(crate) struct StyleSplice {
    pub(crate) range: Range<usize>,
    pub(crate) expected: Vec<u8>,
    pub(crate) replacement: Vec<u8>,
}

pub(crate) fn source_splices(
    xml: &str,
    changes: &[ChartStylePropertyChange],
    limits: crate::Limits,
) -> Result<Vec<StyleSplice>> {
    let mut replacements = checked_replacements(xml, changes, limits)?;
    checked_replacement_output_size(xml.len(), &mut replacements, limits)?;
    let mut splices = Vec::new();
    reserve_exact(&mut splices, replacements.len(), "style splice plan")?;
    for replacement in replacements {
        let expected = copy_bytes(
            &xml.as_bytes()[replacement.range.clone()],
            "style splice source proof",
        )?;
        splices.push(StyleSplice {
            expected,
            range: replacement.range,
            replacement: replacement.value,
        });
    }
    Ok(splices)
}

fn checked_replacements(
    xml: &str,
    changes: &[ChartStylePropertyChange],
    limits: crate::Limits,
) -> Result<Vec<Replacement>> {
    if changes.is_empty() {
        return Ok(Vec::new());
    }
    let parsed = parse(xml, limits)?;
    let mut replacements = Vec::new();
    reserve_exact(&mut replacements, changes.len(), "style replacement plan")?;
    for change in changes {
        let Some(style) = parsed
            .styles
            .iter()
            .find(|style| style.name == change.style_name)
        else {
            return Err(invalid("ODC chart style selector is missing"));
        };
        let properties = style
            .properties
            .as_ref()
            .ok_or_else(|| invalid("ODC chart style has no style:chart-properties element"))?;
        validate_optional_value(change.property, change.after.as_ref(), limits)?;
        let current = properties
            .attributes
            .iter()
            .find(|attribute| attribute.property == change.property)
            .map(|attribute| decode_value(change.property, &attribute.value, limits))
            .transpose()?;
        if current != change.before {
            return Err(Error::InvalidFormat(format!(
                "ODC chart style property '{}' is stale",
                property_name(change.property)
            )));
        }
        if current == change.after {
            continue;
        }
        if let Some(attribute) = properties
            .attributes
            .iter()
            .find(|attribute| attribute.property == change.property)
        {
            if let Some(after) = change.after.as_ref() {
                let lexical = encode_value(change.property, after, limits)?;
                push_one(
                    &mut replacements,
                    Replacement {
                        range: attribute.value_span.clone(),
                        value: escape_attribute(&lexical).into_bytes(),
                    },
                    "style replacement",
                )?;
            } else {
                push_one(
                    &mut replacements,
                    Replacement {
                        range: attribute.attribute_span.clone(),
                        value: Vec::new(),
                    },
                    "style replacement",
                )?;
            }
        } else if let Some(after) = change.after.as_ref() {
            let prefix = properties.chart_prefix.as_deref().ok_or_else(|| {
                Error::Unsupported(
                    "cannot add chart formatting without a lossless chart namespace prefix".into(),
                )
            })?;
            let lexical = encode_value(change.property, after, limits)?;
            let offset = insertion_offset(&xml.as_bytes()[properties.tag.clone()])?;
            push_one(
                &mut replacements,
                Replacement {
                    range: properties.tag.start + offset..properties.tag.start + offset,
                    value: format!(
                        " {prefix}:{}=\"{}\"",
                        property_name(change.property),
                        escape_attribute(&lexical)
                    )
                    .into_bytes(),
                },
                "style replacement",
            )?;
        }
    }
    if replacements.is_empty() {
        return Ok(replacements);
    }
    let mut coalesced: Vec<Replacement> = Vec::new();
    reserve_exact(
        &mut coalesced,
        replacements.len(),
        "coalesced style replacement plan",
    )?;
    for replacement in replacements {
        if let Some(previous) = coalesced.last_mut()
            && previous.range.start == previous.range.end
            && replacement.range.start == replacement.range.end
            && previous.range.start == replacement.range.start
        {
            previous
                .value
                .try_reserve(replacement.value.len())
                .map_err(|error| {
                    invalid(format!("style replacement allocation failed: {error}"))
                })?;
            previous.value.extend_from_slice(&replacement.value);
        } else {
            coalesced.push(replacement);
        }
    }
    Ok(coalesced)
}

fn apply_replacements(
    xml: &str,
    mut replacements: Vec<Replacement>,
    limits: crate::Limits,
) -> Result<String> {
    let output_size = checked_replacement_output_size(xml.len(), &mut replacements, limits)?;
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(output_size)
        .map_err(|error| invalid(format!("edited ODC styles allocation failed: {error}")))?;
    let mut cursor = 0usize;
    for replacement in replacements {
        bytes.extend_from_slice(&xml.as_bytes()[cursor..replacement.range.start]);
        bytes.extend_from_slice(&replacement.value);
        cursor = replacement.range.end;
    }
    bytes.extend_from_slice(&xml.as_bytes()[cursor..]);
    String::from_utf8(bytes).map_err(|_error| invalid("ODC edited styles are not UTF-8"))
}

/// Check every source splice and calculate the final part length before any
/// owned output buffer is allocated.  The caller-selected content ceiling is
/// part of this plan because the shared XML publisher only knows its own
/// generic hard ceiling.
pub(crate) fn checked_splice_output_size(
    source_len: usize,
    splices: &[StyleSplice],
    limits: crate::Limits,
) -> Result<usize> {
    let mut removed = 0usize;
    let mut added = 0usize;
    for (index, splice) in splices.iter().enumerate() {
        validate_range(source_len, &splice.range)?;
        for previous in &splices[..index] {
            if ranges_overlap_or_conflict(&previous.range, &splice.range) {
                return Err(invalid("ODC chart style changes contain overlapping spans"));
            }
        }
        removed = removed
            .checked_add(splice.range.end - splice.range.start)
            .ok_or_else(|| invalid("ODC chart style output size overflow"))?;
        added = added
            .checked_add(splice.replacement.len())
            .ok_or_else(|| invalid("ODC chart style output size overflow"))?;
    }
    checked_output_size(source_len, removed, added, limits)
}

fn checked_replacement_output_size(
    source_len: usize,
    replacements: &mut [Replacement],
    limits: crate::Limits,
) -> Result<usize> {
    replacements
        .sort_unstable_by_key(|replacement| (replacement.range.start, replacement.range.end));
    let mut removed = 0usize;
    let mut added = 0usize;
    for (index, replacement) in replacements.iter().enumerate() {
        validate_range(source_len, &replacement.range)?;
        if let Some(previous) = replacements[..index].iter().rev().find(|candidate| {
            candidate.range.start < replacement.range.start
                || candidate.range.end < replacement.range.end
        }) && ranges_overlap_or_conflict(&previous.range, &replacement.range)
        {
            return Err(invalid("ODC chart style changes contain overlapping spans"));
        }
        removed = removed
            .checked_add(replacement.range.end - replacement.range.start)
            .ok_or_else(|| invalid("ODC chart style output size overflow"))?;
        added = added
            .checked_add(replacement.value.len())
            .ok_or_else(|| invalid("ODC chart style output size overflow"))?;
    }
    checked_output_size(source_len, removed, added, limits)
}

fn checked_output_size(
    source_len: usize,
    removed: usize,
    added: usize,
    limits: crate::Limits,
) -> Result<usize> {
    if source_len > limits.max_content_bytes() {
        return Err(invalid(
            "ODC styles exceed the caller-selected content limit",
        ));
    }
    let output_size = source_len
        .checked_sub(removed)
        .and_then(|size| size.checked_add(added))
        .ok_or_else(|| invalid("ODC chart style output size overflow"))?;
    if output_size > limits.max_content_bytes() {
        return Err(invalid(
            "edited ODC styles exceed the caller-selected content limit",
        ));
    }
    Ok(output_size)
}

fn validate_range(source_len: usize, range: &Range<usize>) -> Result<()> {
    if range.start > range.end || range.end > source_len {
        return Err(invalid(
            "ODC chart style change range is outside the source",
        ));
    }
    Ok(())
}

fn ranges_overlap_or_conflict(left: &Range<usize>, right: &Range<usize>) -> bool {
    if left.start == left.end && right.start == right.end {
        return left.start == right.start;
    }
    if left.start == left.end {
        return left.start > right.start && left.start < right.end;
    }
    if right.start == right.end {
        return right.start > left.start && right.start < left.end;
    }
    left.start < right.end && right.start < left.end
}

fn reserve_exact<T>(values: &mut Vec<T>, additional: usize, label: &str) -> Result<()> {
    values
        .try_reserve_exact(additional)
        .map_err(|error| invalid(format!("ODC {label} allocation failed: {error}")))
}

fn push_one<T>(values: &mut Vec<T>, value: T, label: &str) -> Result<()> {
    values
        .try_reserve(1)
        .map_err(|error| invalid(format!("ODC {label} allocation failed: {error}")))?;
    values.push(value);
    Ok(())
}

fn copy_bytes(bytes: &[u8], label: &str) -> Result<Vec<u8>> {
    let mut copy = Vec::new();
    reserve_exact(&mut copy, bytes.len(), label)?;
    copy.extend_from_slice(bytes);
    Ok(copy)
}

pub(crate) fn validate_change_value(
    property: ChartStyleProperty,
    value: Option<&ChartStyleValue>,
    limits: crate::Limits,
) -> Result<()> {
    validate_optional_value(property, value, limits)
}

/// Derive property-level changes when a patch is reconstructed from bytes.
pub(crate) fn changes_between(
    before: &str,
    after: &str,
    limits: crate::Limits,
) -> Result<Vec<ChartStylePropertyChange>> {
    if before == after {
        return Ok(Vec::new());
    }
    let before_parsed = parse(before, limits)?;
    let after_parsed = parse(after, limits)?;
    let names = before_parsed
        .styles
        .iter()
        .chain(&after_parsed.styles)
        .map(|style| style.name.clone())
        .collect::<BTreeSet<_>>();
    let mut changes = Vec::new();
    for style_name in names {
        for property in ALL_PROPERTIES {
            let old = read_property(before, &style_name, *property, limits)?;
            let new = read_property(after, &style_name, *property, limits)?;
            if old != new {
                push_one(
                    &mut changes,
                    ChartStylePropertyChange::new(style_name.clone(), *property, old, new),
                    "style change summary",
                )?;
            }
        }
    }
    Ok(changes)
}

const ALL_PROPERTIES: &[ChartStyleProperty] = &[
    ChartStyleProperty::Logarithmic,
    ChartStyleProperty::MinorLogarithmic,
    ChartStyleProperty::TickMarksMajorInner,
    ChartStyleProperty::TickMarksMajorOuter,
    ChartStyleProperty::TickMarksMinorInner,
    ChartStyleProperty::TickMarksMinorOuter,
    ChartStyleProperty::SymbolType,
    ChartStyleProperty::SymbolName,
    ChartStyleProperty::ErrorCategory,
    ChartStyleProperty::ErrorPercentage,
    ChartStyleProperty::ErrorMargin,
    ChartStyleProperty::ErrorLowerLimit,
    ChartStyleProperty::ErrorUpperLimit,
    ChartStyleProperty::ErrorUpperIndicator,
    ChartStyleProperty::ErrorLowerIndicator,
    ChartStyleProperty::RegressionType,
    ChartStyleProperty::RegressionMaxDegree,
    ChartStyleProperty::RegressionForceIntercept,
    ChartStyleProperty::RegressionInterceptValue,
    ChartStyleProperty::RegressionName,
    ChartStyleProperty::RegressionPeriod,
    ChartStyleProperty::RegressionMovingType,
];

fn parse(xml: &str, limits: crate::Limits) -> Result<ParsedStyles> {
    if xml.len() > limits.max_content_bytes() {
        return Err(invalid(
            "ODC styles exceed the caller-selected content limit",
        ));
    }
    crate::codec::validate_styles(xml, limits)?;
    let mut reader = NsReader::from_reader(xml.as_bytes());
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut styles = Vec::new();
    let mut active_styles: Vec<(usize, usize)> = Vec::new();
    let mut active_properties: Vec<(usize, usize)> = Vec::new();
    let mut active_symbol_images = Vec::new();
    let mut namespaces: Vec<BTreeMap<String, String>> = Vec::new();
    loop {
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_overflow| invalid("ODC styles event offset exceeds this platform"))?;
        let (resolved_namespace, borrowed_event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid ODC styles XML: {error}")))?;
        let namespace_kind = namespace_kind(&resolved_namespace);
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|_overflow| invalid("ODC styles event offset exceeds this platform"))?;
        let event = borrowed_event.into_owned();
        match event {
            Event::Start(element) => {
                let event_depth = checked_depth(depth, limits.max_depth())?;
                let current = namespace_scope(namespaces.last(), &element)?;
                let namespace = namespace_kind;
                let attributes = read_attributes(
                    &reader,
                    &element,
                    event_start..event_end,
                    xml.as_bytes(),
                    limits,
                )?;
                if active_symbol_images
                    .last()
                    .is_some_and(|image_depth| *image_depth < event_depth)
                {
                    return Err(invalid(
                        "ODC chart:symbol-image cannot contain child elements",
                    ));
                }
                if namespace == NamespaceKind::Style && element.local_name().as_ref() == b"style" {
                    let family = attributes
                        .iter()
                        .find(|attribute| attribute.uri == STYLE && attribute.local == b"family")
                        .map(|attribute| attribute.value.as_str());
                    let name = attributes
                        .iter()
                        .find(|attribute| attribute.uri == STYLE && attribute.local == b"name")
                        .map(|attribute| attribute.value.as_str());
                    if family == Some("chart") {
                        let name = name.ok_or_else(|| invalid("chart style has no style:name"))?;
                        validate_style_name(name, limits)?;
                        if styles.iter().any(|style: &StyleRecord| style.name == name) {
                            return Err(invalid("ODC chart styles contain a duplicate style name"));
                        }
                        push_one(
                            &mut styles,
                            StyleRecord {
                                name: name.to_owned(),
                                properties: None,
                            },
                            "chart style collection",
                        )?;
                        let index = styles.len() - 1;
                        push_one(
                            &mut active_styles,
                            (event_depth, index),
                            "active chart style stack",
                        )?;
                    }
                } else if namespace == NamespaceKind::Style
                    && element.local_name().as_ref() == b"chart-properties"
                    && active_styles
                        .last()
                        .is_some_and(|(style_depth, _)| *style_depth + 1 == event_depth)
                {
                    let (_, index) = *active_styles
                        .last()
                        .ok_or_else(|| invalid("ODC chart properties style scope disappeared"))?;
                    let properties = chart_properties(
                        &reader,
                        &element,
                        event_start..event_end,
                        &attributes,
                        chart_prefix(&current),
                        limits,
                    )?;
                    if styles[index].properties.is_some() {
                        return Err(invalid(
                            "ODC chart style contains duplicate style:chart-properties",
                        ));
                    }
                    styles[index].properties = Some(properties);
                    push_one(
                        &mut active_properties,
                        (event_depth, index),
                        "active chart property stack",
                    )?;
                } else if namespace == NamespaceKind::Chart
                    && element.local_name().as_ref() == b"symbol-image"
                    && active_properties
                        .last()
                        .is_some_and(|(property_depth, _)| *property_depth + 1 == event_depth)
                {
                    let (_, index) = *active_properties
                        .last()
                        .ok_or_else(|| invalid("ODC chart properties scope disappeared"))?;
                    if let Some(properties) = styles[index].properties.as_mut() {
                        validate_symbol_image(&attributes)?;
                        if properties.has_symbol_image {
                            return Err(invalid(
                                "ODC chart properties contain duplicate chart:symbol-image",
                            ));
                        }
                        properties.has_symbol_image = true;
                        push_one(
                            &mut active_symbol_images,
                            event_depth,
                            "active symbol-image stack",
                        )?;
                    }
                }
                push_one(&mut namespaces, current, "chart style namespace stack")?;
                depth = event_depth;
            },
            Event::Empty(element) => {
                let event_depth = checked_depth(depth, limits.max_depth())?;
                let current = namespace_scope(namespaces.last(), &element)?;
                let namespace = namespace_kind;
                let attributes = read_attributes(
                    &reader,
                    &element,
                    event_start..event_end,
                    xml.as_bytes(),
                    limits,
                )?;
                if active_symbol_images
                    .last()
                    .is_some_and(|image_depth| *image_depth < event_depth)
                {
                    return Err(invalid(
                        "ODC chart:symbol-image cannot contain child elements",
                    ));
                }
                if namespace == NamespaceKind::Style && element.local_name().as_ref() == b"style" {
                    let family = attributes
                        .iter()
                        .find(|attribute| attribute.uri == STYLE && attribute.local == b"family")
                        .map(|attribute| attribute.value.as_str());
                    let name = attributes
                        .iter()
                        .find(|attribute| attribute.uri == STYLE && attribute.local == b"name")
                        .map(|attribute| attribute.value.as_str());
                    if family == Some("chart") {
                        let name = name.ok_or_else(|| invalid("chart style has no style:name"))?;
                        validate_style_name(name, limits)?;
                        if styles.iter().any(|style: &StyleRecord| style.name == name) {
                            return Err(invalid("ODC chart styles contain a duplicate style name"));
                        }
                        push_one(
                            &mut styles,
                            StyleRecord {
                                name: name.to_owned(),
                                properties: None,
                            },
                            "chart style collection",
                        )?;
                    }
                } else if namespace == NamespaceKind::Style
                    && element.local_name().as_ref() == b"chart-properties"
                    && active_styles
                        .last()
                        .is_some_and(|(style_depth, _)| *style_depth + 1 == event_depth)
                {
                    let (_, index) = *active_styles
                        .last()
                        .ok_or_else(|| invalid("ODC chart properties style scope disappeared"))?;
                    let properties = chart_properties(
                        &reader,
                        &element,
                        event_start..event_end,
                        &attributes,
                        chart_prefix(&current),
                        limits,
                    )?;
                    if styles[index].properties.is_some() {
                        return Err(invalid(
                            "ODC chart style contains duplicate style:chart-properties",
                        ));
                    }
                    validate_symbol_grammar(&properties)?;
                    styles[index].properties = Some(properties);
                } else if namespace == NamespaceKind::Chart
                    && element.local_name().as_ref() == b"symbol-image"
                    && active_properties
                        .last()
                        .is_some_and(|(property_depth, _)| *property_depth + 1 == event_depth)
                {
                    let (_, index) = *active_properties
                        .last()
                        .ok_or_else(|| invalid("ODC chart properties scope disappeared"))?;
                    if let Some(properties) = styles[index].properties.as_mut() {
                        validate_symbol_image(&attributes)?;
                        if properties.has_symbol_image {
                            return Err(invalid(
                                "ODC chart properties contain duplicate chart:symbol-image",
                            ));
                        }
                        properties.has_symbol_image = true;
                    }
                }
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid("ODC styles depth underflow"));
                }
                if active_symbol_images
                    .last()
                    .is_some_and(|image_depth| *image_depth == depth)
                {
                    active_symbol_images.pop();
                }
                if active_styles
                    .last()
                    .is_some_and(|(style_depth, _)| *style_depth == depth)
                {
                    active_styles.pop();
                }
                if active_properties
                    .last()
                    .is_some_and(|(property_depth, _)| *property_depth == depth)
                {
                    let (_, index) = active_properties
                        .pop()
                        .ok_or_else(|| invalid("ODC chart properties scope disappeared"))?;
                    if let Some(properties) = styles[index].properties.as_ref() {
                        validate_symbol_grammar(properties)?;
                    }
                }
                namespaces.pop();
                depth -= 1;
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in ODC styles")),
            Event::Eof => break,
            Event::Text(text)
                if active_symbol_images.last().is_some()
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(invalid("ODC chart:symbol-image cannot contain text"));
            },
            Event::CData(_) | Event::GeneralRef(_) if active_symbol_images.last().is_some() => {
                return Err(invalid("ODC chart:symbol-image cannot contain text"));
            },
            _ => {},
        }
    }
    if depth != 0
        || !active_styles.is_empty()
        || !active_properties.is_empty()
        || !active_symbol_images.is_empty()
        || !namespaces.is_empty()
    {
        return Err(invalid("ODC styles structure is incomplete"));
    }
    Ok(ParsedStyles { styles })
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum NamespaceKind {
    Style,
    Chart,
    Other,
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == STYLE) {
        NamespaceKind::Style
    } else if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == CHART) {
        NamespaceKind::Chart
    } else {
        NamespaceKind::Other
    }
}

#[derive(Clone, Debug)]
struct RawAttribute {
    uri: Vec<u8>,
    local: Vec<u8>,
    value: String,
    value_span: Range<usize>,
    attribute_span: Range<usize>,
}

fn read_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    event: Range<usize>,
    bytes: &[u8],
    limits: crate::Limits,
) -> Result<Vec<RawAttribute>> {
    let raw = bytes
        .get(event.clone())
        .ok_or_else(|| invalid("ODC styles attribute span is outside the source"))?;
    let mut attributes = Vec::new();
    for result in element.attributes().with_checks(true) {
        let attribute =
            result.map_err(|error| invalid(format!("invalid ODC styles attribute: {error}")))?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid ODC styles value: {error}")))?
            .into_owned();
        if value.len() > limits.max_scalar_bytes() {
            return Err(invalid(
                "ODC styles attribute exceeds the caller-selected scalar limit",
            ));
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let uri = match namespace {
            ResolveResult::Bound(Namespace(uri)) => uri.to_vec(),
            _ => Vec::new(),
        };
        let local = local.as_ref().to_vec();
        let qualified = attribute.key.as_ref();
        let (relative_value, relative_attribute) = attribute_spans(raw, qualified)?;
        push_one(
            &mut attributes,
            RawAttribute {
                uri,
                local,
                value,
                value_span: event.start + relative_value.start..event.start + relative_value.end,
                attribute_span: event.start + relative_attribute.start
                    ..event.start + relative_attribute.end,
            },
            "chart style attribute collection",
        )?;
    }
    Ok(attributes)
}

fn chart_properties(
    reader: &NsReader<&[u8]>,
    _element: &BytesStart<'_>,
    tag: Range<usize>,
    attributes: &[RawAttribute],
    chart_prefix: Option<String>,
    limits: crate::Limits,
) -> Result<ChartProperties> {
    let mut parsed = Vec::new();
    for attribute in attributes {
        if attribute.uri != CHART {
            continue;
        }
        let Some(property) = property_from_local(&attribute.local) else {
            continue;
        };
        decode_value(property, &attribute.value, limits)?;
        if parsed
            .iter()
            .any(|candidate: &AttributeRecord| candidate.property == property)
        {
            return Err(invalid(format!(
                "ODC chart style duplicates chart:{}",
                property_name(property)
            )));
        }
        if parsed.len() >= MAX_PROPERTIES_PER_STYLE {
            return Err(invalid("ODC chart style has too many modeled properties"));
        }
        push_one(
            &mut parsed,
            AttributeRecord {
                property,
                value: attribute.value.clone(),
                value_span: attribute.value_span.clone(),
                attribute_span: attribute.attribute_span.clone(),
            },
            "modeled chart property collection",
        )?;
    }
    // The resolver proves that each existing chart attribute is bound to the
    // normative chart URI.  The prefix map is only used when a new attribute
    // must be inserted and is therefore deliberately kept as lexical data.
    let _ = reader;
    Ok(ChartProperties {
        tag,
        chart_prefix,
        attributes: parsed,
        has_symbol_image: false,
    })
}

fn validate_symbol_grammar(properties: &ChartProperties) -> Result<()> {
    let symbol_type = properties
        .attributes
        .iter()
        .find(|attribute| attribute.property == ChartStyleProperty::SymbolType)
        .map(|attribute| parse_symbol_type(&attribute.value))
        .transpose()?;
    let has_symbol_name = properties
        .attributes
        .iter()
        .any(|attribute| attribute.property == ChartStyleProperty::SymbolName);
    let has_symbol_image = properties.has_symbol_image;
    match symbol_type {
        Some(ChartSymbolType::NamedSymbol) if !has_symbol_name => Err(invalid(
            "ODC chart:symbol-type named-symbol requires chart:symbol-name",
        )),
        Some(ChartSymbolType::NamedSymbol) if has_symbol_image => Err(invalid(
            "ODC chart:symbol-image requires chart:symbol-type image",
        )),
        Some(ChartSymbolType::Image) if !has_symbol_image => Err(invalid(
            "ODC chart:symbol-type image requires chart:symbol-image",
        )),
        Some(ChartSymbolType::Image) if has_symbol_name => Err(invalid(
            "ODC chart:symbol-name is incompatible with chart:symbol-type image",
        )),
        Some(ChartSymbolType::None | ChartSymbolType::Automatic)
            if has_symbol_name || has_symbol_image =>
        {
            Err(invalid(
                "chart:symbol-name and chart:symbol-image require their matching symbol type",
            ))
        },
        None if has_symbol_name || has_symbol_image => Err(invalid(
            "chart:symbol-name and chart:symbol-image require chart:symbol-type",
        )),
        _ => Ok(()),
    }
}

fn validate_symbol_image(attributes: &[RawAttribute]) -> Result<()> {
    if attributes
        .iter()
        .filter(|attribute| attribute.uri == XLINK && attribute.local == b"href")
        .count()
        != 1
    {
        return Err(invalid("ODC chart:symbol-image requires xlink:href"));
    }
    Ok(())
}

fn namespace_scope(
    parent: Option<&BTreeMap<String, String>>,
    element: &BytesStart<'_>,
) -> Result<BTreeMap<String, String>> {
    let mut scope = parent.cloned().unwrap_or_default();
    for result in element.attributes().with_checks(true) {
        let attribute =
            result.map_err(|error| invalid(format!("invalid ODC styles namespace: {error}")))?;
        let key = attribute.key.as_ref();
        let prefix = if key == b"xmlns" {
            String::new()
        } else if let Some(prefix) = key.strip_prefix(b"xmlns:") {
            std::str::from_utf8(prefix)
                .map_err(|_error| invalid("ODC styles namespace prefix is not UTF-8"))?
                .to_owned()
        } else {
            continue;
        };
        let value = std::str::from_utf8(attribute.value.as_ref())
            .map_err(|_error| invalid("ODC styles namespace URI is not UTF-8"))?;
        if value.len() > 256 {
            return Err(invalid("ODC styles namespace URI is too long"));
        }
        scope.insert(prefix, value.to_owned());
    }
    Ok(scope)
}

fn chart_prefix(scope: &BTreeMap<String, String>) -> Option<String> {
    scope
        .iter()
        // The XML default namespace never applies to attributes.  A chart
        // attribute therefore needs a non-empty prefix even when the chart
        // URI is also bound as the default namespace for elements.
        .find(|(prefix, uri)| !prefix.is_empty() && uri.as_bytes() == CHART)
        .map(|(prefix, _uri)| prefix.clone())
}

fn property_from_local(local: &[u8]) -> Option<ChartStyleProperty> {
    Some(match local {
        b"logarithmic" => ChartStyleProperty::Logarithmic,
        b"minor-logarithmic" => ChartStyleProperty::MinorLogarithmic,
        b"tick-marks-major-inner" => ChartStyleProperty::TickMarksMajorInner,
        b"tick-marks-major-outer" => ChartStyleProperty::TickMarksMajorOuter,
        b"tick-marks-minor-inner" => ChartStyleProperty::TickMarksMinorInner,
        b"tick-marks-minor-outer" => ChartStyleProperty::TickMarksMinorOuter,
        b"symbol-type" => ChartStyleProperty::SymbolType,
        b"symbol-name" => ChartStyleProperty::SymbolName,
        b"error-category" => ChartStyleProperty::ErrorCategory,
        b"error-percentage" => ChartStyleProperty::ErrorPercentage,
        b"error-margin" => ChartStyleProperty::ErrorMargin,
        b"error-lower-limit" => ChartStyleProperty::ErrorLowerLimit,
        b"error-upper-limit" => ChartStyleProperty::ErrorUpperLimit,
        b"error-upper-indicator" => ChartStyleProperty::ErrorUpperIndicator,
        b"error-lower-indicator" => ChartStyleProperty::ErrorLowerIndicator,
        b"regression-type" => ChartStyleProperty::RegressionType,
        b"regression-max-degree" => ChartStyleProperty::RegressionMaxDegree,
        b"regression-force-intercept" => ChartStyleProperty::RegressionForceIntercept,
        b"regression-intercept-value" => ChartStyleProperty::RegressionInterceptValue,
        b"regression-name" => ChartStyleProperty::RegressionName,
        b"regression-period" => ChartStyleProperty::RegressionPeriod,
        b"regression-moving-type" => ChartStyleProperty::RegressionMovingType,
        _ => return None,
    })
}

fn property_spec(property: ChartStyleProperty) -> (&'static str, ValueKind) {
    match property {
        ChartStyleProperty::Logarithmic => ("logarithmic", ValueKind::Boolean),
        ChartStyleProperty::MinorLogarithmic => ("minor-logarithmic", ValueKind::Boolean),
        ChartStyleProperty::TickMarksMajorInner => ("tick-marks-major-inner", ValueKind::Boolean),
        ChartStyleProperty::TickMarksMajorOuter => ("tick-marks-major-outer", ValueKind::Boolean),
        ChartStyleProperty::TickMarksMinorInner => ("tick-marks-minor-inner", ValueKind::Boolean),
        ChartStyleProperty::TickMarksMinorOuter => ("tick-marks-minor-outer", ValueKind::Boolean),
        ChartStyleProperty::SymbolType => ("symbol-type", ValueKind::SymbolType),
        ChartStyleProperty::SymbolName => ("symbol-name", ValueKind::SymbolName),
        ChartStyleProperty::ErrorCategory => ("error-category", ValueKind::ErrorCategory),
        ChartStyleProperty::ErrorPercentage => ("error-percentage", ValueKind::Decimal),
        ChartStyleProperty::ErrorMargin => ("error-margin", ValueKind::Decimal),
        ChartStyleProperty::ErrorLowerLimit => ("error-lower-limit", ValueKind::Decimal),
        ChartStyleProperty::ErrorUpperLimit => ("error-upper-limit", ValueKind::Decimal),
        ChartStyleProperty::ErrorUpperIndicator => ("error-upper-indicator", ValueKind::Boolean),
        ChartStyleProperty::ErrorLowerIndicator => ("error-lower-indicator", ValueKind::Boolean),
        ChartStyleProperty::RegressionType => ("regression-type", ValueKind::RegressionType),
        ChartStyleProperty::RegressionMaxDegree => {
            ("regression-max-degree", ValueKind::PositiveInteger)
        },
        ChartStyleProperty::RegressionForceIntercept => {
            ("regression-force-intercept", ValueKind::Boolean)
        },
        ChartStyleProperty::RegressionInterceptValue => {
            ("regression-intercept-value", ValueKind::Decimal)
        },
        ChartStyleProperty::RegressionName => ("regression-name", ValueKind::Text),
        ChartStyleProperty::RegressionPeriod => ("regression-period", ValueKind::PositiveInteger),
        ChartStyleProperty::RegressionMovingType => {
            ("regression-moving-type", ValueKind::RegressionMovingType)
        },
    }
}

fn property_name(property: ChartStyleProperty) -> &'static str {
    property_spec(property).0
}

fn decode_value(
    property: ChartStyleProperty,
    lexical: &str,
    limits: crate::Limits,
) -> Result<ChartStyleValue> {
    let (_, kind) = property_spec(property);
    match kind {
        ValueKind::Boolean => match schema_whitespace(lexical) {
            "true" => Ok(ChartStyleValue::Boolean(true)),
            "false" => Ok(ChartStyleValue::Boolean(false)),
            _ => Err(invalid(format!(
                "invalid chart:{} boolean",
                property_name(property)
            ))),
        },
        ValueKind::Decimal => parse_double(lexical, property).map(ChartStyleValue::Decimal),
        ValueKind::PositiveInteger => {
            let value = schema_whitespace(lexical)
                .parse::<u32>()
                .map_err(|_error| {
                    invalid(format!(
                        "invalid chart:{} positive integer",
                        property_name(property)
                    ))
                })?;
            let minimum = match property {
                ChartStyleProperty::RegressionMaxDegree | ChartStyleProperty::RegressionPeriod => 2,
                _ => 1,
            };
            if value < minimum {
                return Err(invalid(format!(
                    "chart:{} must be at least {minimum}",
                    property_name(property)
                )));
            }
            Ok(ChartStyleValue::PositiveInteger(value))
        },
        ValueKind::Text => {
            if lexical.len() > limits.max_scalar_bytes()
                || lexical.chars().any(is_forbidden_xml_character)
            {
                return Err(invalid(format!(
                    "invalid chart:{} text",
                    property_name(property)
                )));
            }
            Ok(ChartStyleValue::Text(lexical.to_owned()))
        },
        ValueKind::SymbolType => parse_symbol_type(lexical).map(ChartStyleValue::SymbolType),
        ValueKind::SymbolName => parse_symbol_name(lexical).map(ChartStyleValue::SymbolName),
        ValueKind::ErrorCategory => {
            parse_error_category(lexical).map(ChartStyleValue::ErrorCategory)
        },
        ValueKind::RegressionType => {
            parse_regression_type(lexical).map(ChartStyleValue::RegressionType)
        },
        ValueKind::RegressionMovingType => {
            parse_regression_moving_type(lexical).map(ChartStyleValue::RegressionMovingType)
        },
    }
}

fn validate_optional_value(
    property: ChartStyleProperty,
    value: Option<&ChartStyleValue>,
    limits: crate::Limits,
) -> Result<()> {
    if let Some(value) = value {
        encode_value(property, value, limits).map(|_| ())
    } else {
        Ok(())
    }
}

fn encode_value(
    property: ChartStyleProperty,
    value: &ChartStyleValue,
    limits: crate::Limits,
) -> Result<String> {
    let (_, kind) = property_spec(property);
    let lexical = match (kind, value) {
        (ValueKind::Boolean, ChartStyleValue::Boolean(value)) => value.to_string(),
        (ValueKind::Decimal, ChartStyleValue::Decimal(value)) => encode_double(*value),
        (ValueKind::PositiveInteger, ChartStyleValue::PositiveInteger(value)) => {
            let minimum = match property {
                ChartStyleProperty::RegressionMaxDegree | ChartStyleProperty::RegressionPeriod => 2,
                _ => 1,
            };
            if *value >= minimum {
                value.to_string()
            } else {
                return Err(invalid(format!(
                    "chart:{} must be at least {minimum}",
                    property_name(property)
                )));
            }
        },
        (ValueKind::Text, ChartStyleValue::Text(value))
            if value.len() <= limits.max_scalar_bytes()
                && !value.chars().any(is_forbidden_xml_character) =>
        {
            value.clone()
        },
        (ValueKind::SymbolType, ChartStyleValue::SymbolType(value)) => {
            symbol_type_lexical(*value).to_owned()
        },
        (ValueKind::SymbolName, ChartStyleValue::SymbolName(value)) => {
            symbol_name_lexical(*value).to_owned()
        },
        (ValueKind::ErrorCategory, ChartStyleValue::ErrorCategory(value)) => {
            error_category_lexical(*value).to_owned()
        },
        (ValueKind::RegressionType, ChartStyleValue::RegressionType(value)) => {
            regression_type_lexical(*value).to_owned()
        },
        (ValueKind::RegressionMovingType, ChartStyleValue::RegressionMovingType(value)) => {
            regression_moving_type_lexical(*value).to_owned()
        },
        _ => {
            return Err(invalid(format!(
                "chart:{} received a value of the wrong type",
                property_name(property)
            )));
        },
    };
    if lexical.len() > limits.max_scalar_bytes() {
        return Err(invalid("ODC chart formatting value exceeds scalar limit"));
    }
    Ok(lexical)
}

fn parse_symbol_type(value: &str) -> Result<ChartSymbolType> {
    match value {
        "none" => Ok(ChartSymbolType::None),
        "automatic" => Ok(ChartSymbolType::Automatic),
        "named-symbol" => Ok(ChartSymbolType::NamedSymbol),
        "image" => Ok(ChartSymbolType::Image),
        _ => Err(invalid("invalid chart:symbol-type value")),
    }
}

const fn symbol_type_lexical(value: ChartSymbolType) -> &'static str {
    match value {
        ChartSymbolType::None => "none",
        ChartSymbolType::Automatic => "automatic",
        ChartSymbolType::NamedSymbol => "named-symbol",
        ChartSymbolType::Image => "image",
    }
}

fn parse_symbol_name(value: &str) -> Result<ChartSymbolName> {
    Ok(match value {
        "square" => ChartSymbolName::Square,
        "diamond" => ChartSymbolName::Diamond,
        "arrow-down" => ChartSymbolName::ArrowDown,
        "arrow-up" => ChartSymbolName::ArrowUp,
        "arrow-right" => ChartSymbolName::ArrowRight,
        "arrow-left" => ChartSymbolName::ArrowLeft,
        "bow-tie" => ChartSymbolName::BowTie,
        "hourglass" => ChartSymbolName::Hourglass,
        "circle" => ChartSymbolName::Circle,
        "star" => ChartSymbolName::Star,
        "x" => ChartSymbolName::X,
        "plus" => ChartSymbolName::Plus,
        "asterisk" => ChartSymbolName::Asterisk,
        "horizontal-bar" => ChartSymbolName::HorizontalBar,
        "vertical-bar" => ChartSymbolName::VerticalBar,
        _ => return Err(invalid("invalid chart:symbol-name value")),
    })
}

fn symbol_name_lexical(value: ChartSymbolName) -> &'static str {
    match value {
        ChartSymbolName::Square => "square",
        ChartSymbolName::Diamond => "diamond",
        ChartSymbolName::ArrowDown => "arrow-down",
        ChartSymbolName::ArrowUp => "arrow-up",
        ChartSymbolName::ArrowRight => "arrow-right",
        ChartSymbolName::ArrowLeft => "arrow-left",
        ChartSymbolName::BowTie => "bow-tie",
        ChartSymbolName::Hourglass => "hourglass",
        ChartSymbolName::Circle => "circle",
        ChartSymbolName::Star => "star",
        ChartSymbolName::X => "x",
        ChartSymbolName::Plus => "plus",
        ChartSymbolName::Asterisk => "asterisk",
        ChartSymbolName::HorizontalBar => "horizontal-bar",
        ChartSymbolName::VerticalBar => "vertical-bar",
    }
}

fn parse_error_category(value: &str) -> Result<ChartErrorCategory> {
    Ok(match value {
        "none" => ChartErrorCategory::None,
        "variance" => ChartErrorCategory::Variance,
        "standard-deviation" => ChartErrorCategory::StandardDeviation,
        "percentage" => ChartErrorCategory::Percentage,
        "error-margin" => ChartErrorCategory::ErrorMargin,
        "constant" => ChartErrorCategory::Constant,
        "standard-error" => ChartErrorCategory::StandardError,
        "cell-range" => ChartErrorCategory::CellRange,
        _ => return Err(invalid("invalid chart:error-category value")),
    })
}

fn error_category_lexical(value: ChartErrorCategory) -> &'static str {
    match value {
        ChartErrorCategory::None => "none",
        ChartErrorCategory::Variance => "variance",
        ChartErrorCategory::StandardDeviation => "standard-deviation",
        ChartErrorCategory::Percentage => "percentage",
        ChartErrorCategory::ErrorMargin => "error-margin",
        ChartErrorCategory::Constant => "constant",
        ChartErrorCategory::StandardError => "standard-error",
        ChartErrorCategory::CellRange => "cell-range",
    }
}

fn parse_regression_type(value: &str) -> Result<ChartRegressionType> {
    Ok(match value {
        "none" => ChartRegressionType::None,
        "linear" => ChartRegressionType::Linear,
        "logarithmic" => ChartRegressionType::Logarithmic,
        "moving-average" => ChartRegressionType::MovingAverage,
        "exponential" => ChartRegressionType::Exponential,
        "power" => ChartRegressionType::Power,
        "polynomial" => ChartRegressionType::Polynomial,
        _ => return Err(invalid("invalid chart:regression-type value")),
    })
}

fn regression_type_lexical(value: ChartRegressionType) -> &'static str {
    match value {
        ChartRegressionType::None => "none",
        ChartRegressionType::Linear => "linear",
        ChartRegressionType::Logarithmic => "logarithmic",
        ChartRegressionType::MovingAverage => "moving-average",
        ChartRegressionType::Exponential => "exponential",
        ChartRegressionType::Power => "power",
        ChartRegressionType::Polynomial => "polynomial",
    }
}

fn parse_regression_moving_type(value: &str) -> Result<ChartRegressionMovingType> {
    Ok(match value {
        "prior" => ChartRegressionMovingType::Prior,
        "central" => ChartRegressionMovingType::Central,
        "averaged-abscissa" => ChartRegressionMovingType::AveragedAbscissa,
        _ => return Err(invalid("invalid chart:regression-moving-type value")),
    })
}

fn regression_moving_type_lexical(value: ChartRegressionMovingType) -> &'static str {
    match value {
        ChartRegressionMovingType::Prior => "prior",
        ChartRegressionMovingType::Central => "central",
        ChartRegressionMovingType::AveragedAbscissa => "averaged-abscissa",
    }
}

fn validate_style_name(style_name: &str, limits: crate::Limits) -> Result<()> {
    let mut characters = style_name.chars();
    let valid_start = characters.next().is_some_and(is_ncname_start_char);
    let valid_rest = characters.all(is_ncname_char);
    if !valid_start || !valid_rest || style_name.len() > limits.max_scalar_bytes() {
        return Err(invalid("ODC chart style name is invalid"));
    }
    Ok(())
}

fn is_ncname_start_char(character: char) -> bool {
    let codepoint = character as u32;
    character == '_'
        || (b'A' as u32..=b'Z' as u32).contains(&codepoint)
        || (b'a' as u32..=b'z' as u32).contains(&codepoint)
        || (0xC0..=0xD6).contains(&codepoint)
        || (0xD8..=0xF6).contains(&codepoint)
        || (0xF8..=0x2FF).contains(&codepoint)
        || (0x370..=0x37D).contains(&codepoint)
        || (0x37F..=0x1FFF).contains(&codepoint)
        || (0x200C..=0x200D).contains(&codepoint)
        || (0x2070..=0x218F).contains(&codepoint)
        || (0x2C00..=0x2FEF).contains(&codepoint)
        || (0x3001..=0xD7FF).contains(&codepoint)
        || (0xF900..=0xFDCF).contains(&codepoint)
        || (0xFDF0..=0xFFFD).contains(&codepoint)
        || (0x10000..=0xEFFFF).contains(&codepoint)
}

fn is_ncname_char(character: char) -> bool {
    let codepoint = character as u32;
    is_ncname_start_char(character)
        || character == '-'
        || character == '.'
        || character == '\u{B7}'
        || (0x300..=0x36F).contains(&codepoint)
        || (0x203F..=0x2040).contains(&codepoint)
}

fn schema_whitespace(value: &str) -> &str {
    value.trim_matches(|character: char| {
        character == ' ' || character == '\t' || character == '\n' || character == '\r'
    })
}

fn parse_double(value: &str, property: ChartStyleProperty) -> Result<f64> {
    let value = schema_whitespace(value);
    match value {
        "INF" => Ok(f64::INFINITY),
        "-INF" => Ok(f64::NEG_INFINITY),
        "NaN" => Ok(f64::NAN),
        _ => {
            let parsed = value.parse::<f64>().map_err(|_error| {
                invalid(format!("invalid chart:{} number", property_name(property)))
            })?;
            if parsed.is_finite() {
                Ok(parsed)
            } else {
                Err(invalid(format!(
                    "invalid chart:{} number",
                    property_name(property)
                )))
            }
        },
    }
}

fn encode_double(value: f64) -> String {
    if value.is_nan() {
        "NaN".to_owned()
    } else if value == f64::INFINITY {
        "INF".to_owned()
    } else if value == f64::NEG_INFINITY {
        "-INF".to_owned()
    } else {
        value.to_string()
    }
}

fn is_forbidden_xml_character(character: char) -> bool {
    !matches!(
        character as u32,
        0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
    )
}

fn checked_depth(depth: usize, limit: usize) -> Result<usize> {
    let next = depth
        .checked_add(1)
        .ok_or_else(|| invalid("ODC styles depth overflow"))?;
    if next > limit {
        return Err(invalid("ODC styles exceed the caller-selected depth limit"));
    }
    Ok(next)
}

fn attribute_spans(tag: &[u8], wanted: &[u8]) -> Result<(Range<usize>, Range<usize>)> {
    let mut cursor = 1usize;
    while cursor < tag.len() && !tag[cursor].is_ascii_whitespace() && tag[cursor] != b'>' {
        cursor += 1;
    }
    while cursor < tag.len() {
        let whitespace = cursor;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag.len() || matches!(tag[cursor], b'/' | b'>') {
            break;
        }
        let name_start = cursor;
        while cursor < tag.len()
            && !tag[cursor].is_ascii_whitespace()
            && !matches!(tag[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if tag.get(cursor) != Some(&b'=') {
            return Err(invalid("ODC styles attribute is missing '='"));
        }
        cursor += 1;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *tag
            .get(cursor)
            .filter(|quote| matches!(quote, b'\'' | b'\"'))
            .ok_or_else(|| invalid("ODC styles attribute is not quoted"))?;
        cursor += 1;
        let value_start = cursor;
        while cursor < tag.len() && tag[cursor] != quote {
            cursor += 1;
        }
        let value_end = cursor;
        cursor = cursor
            .checked_add(1)
            .ok_or_else(|| invalid("ODC styles attribute offset overflow"))?;
        if &tag[name_start..name_end] == wanted {
            return Ok((value_start..value_end, whitespace..cursor));
        }
    }
    Err(invalid("ODC styles attribute span was not found"))
}

fn insertion_offset(tag: &[u8]) -> Result<usize> {
    let close = tag
        .iter()
        .rposition(|byte| *byte == b'>')
        .ok_or_else(|| invalid("ODC chart-properties tag is incomplete"))?;
    Ok(if close > 0 && tag[close - 1] == b'/' {
        close - 1
    } else {
        close
    })
}

#[derive(Debug)]
struct Replacement {
    range: Range<usize>,
    value: Vec<u8>,
}

fn escape_attribute(value: &str) -> String {
    quick_xml::escape::escape(value).into_owned()
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
