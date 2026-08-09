//! Bounded SpreadsheetML/MCE codec for workbook calculation metadata.

use std::borrow::Cow;

use crate::error::{Error, Result, invalid};
use crate::raw::namespace::{SPREADSHEETML_NAMESPACE, STRICT_SPREADSHEETML_NAMESPACE};
use litchi_ooxml_common::mce::{
    Capabilities, Limits as MceLimits, Name, process_markup_compatibility,
};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::features::{Feature, Features};
use super::limits::Limits;
use super::model::{Mode, Properties, ReferenceMode};
use super::rewriter::{Layout, inspect_layout};

#[cfg_attr(
    not(test),
    allow(
        dead_code,
        reason = "the namespace constant is exercised by the test-only strict-conformance parser"
    )
)]
pub(super) const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
pub(super) const CALC_FEATURES_NAMESPACE: &[u8] =
    b"http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures";
pub(super) const CALC_FEATURES_EXTENSION_URI: &str = "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}";

/// One semantic and physical view shared by both calculation-metadata owners.
#[allow(
    dead_code,
    reason = "this lexical parser is retained for the calculation-properties rewriter"
)]
pub(crate) struct Inspection<'a> {
    pub(crate) source: &'a [u8],
    pub(crate) properties: Option<Properties>,
    pub(crate) features: Option<Features>,
    pub(super) layout: Layout,
}

/// Parse the workbook's direct `calcPr` child without executing calculations.
pub fn parse(xml: &[u8]) -> Result<Option<Properties>> {
    parse_with_limits(xml, &Limits::default())
}

/// Parse `calcPr` using caller-supplied resource limits.
pub fn parse_with_limits(xml: &[u8], limits: &Limits) -> Result<Option<Properties>> {
    Ok(inspect_with_policy(xml, limits, true)?.properties)
}

/// Parse the ordered `xcalcf:calcFeatures` payload, if present.
pub fn parse_features(xml: &[u8]) -> Result<Option<Features>> {
    parse_features_with_limits(xml, &Limits::default())
}

/// Parse calculation features using caller-supplied resource limits.
pub fn parse_features_with_limits(xml: &[u8], limits: &Limits) -> Result<Option<Features>> {
    Ok(inspect_with_policy(xml, limits, true)?.features)
}

/// Inspect both owned values and their physical source layout exactly once.
pub(crate) fn inspect<'a>(xml: &'a [u8], limits: &Limits) -> Result<Inspection<'a>> {
    inspect_with_policy(xml, limits, false)
}

fn inspect_with_policy<'a>(
    xml: &'a [u8],
    limits: &Limits,
    strict_calc_attributes: bool,
) -> Result<Inspection<'a>> {
    if xml.len() > limits.max_raw_bytes() {
        return Err(invalid(
            "workbook calculation metadata exceeds raw byte limit",
        ));
    }
    let layout = inspect_layout(xml, limits)?;
    let mut capabilities = Capabilities::default();
    capabilities
        .understand_namespace(String::from_utf8_lossy(CALC_FEATURES_NAMESPACE).into_owned())
        .preserve_extension_element(Name {
            namespace: String::from_utf8_lossy(SPREADSHEETML_NAMESPACE).into_owned(),
            local_name: "ext".into(),
        })
        .preserve_extension_element(Name {
            namespace: String::from_utf8_lossy(STRICT_SPREADSHEETML_NAMESPACE).into_owned(),
            local_name: "ext".into(),
        });
    let mce_limits = MceLimits {
        max_input_bytes: limits.max_raw_bytes(),
        max_output_bytes: limits.max_mce_bytes(),
        max_depth: limits.max_depth(),
        ..MceLimits::default()
    };
    let processed = process_markup_compatibility(xml, &capabilities, &mce_limits)?;
    if processed.xml.len() > limits.max_mce_bytes() {
        return Err(invalid(
            "processed workbook calculation metadata exceeds byte limit",
        ));
    }
    let semantic = parse_semantic(processed.xml.as_ref(), limits, strict_calc_attributes)?;
    Ok(Inspection {
        source: xml,
        properties: semantic.0,
        features: semantic.1,
        layout,
    })
}

fn parse_semantic(
    xml: &[u8],
    limits: &Limits,
    strict_calc_attributes: bool,
) -> Result<(Option<Properties>, Option<Features>)> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut events = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut root_dialect = None;
    let mut last_core_rank = None;
    let mut calc_depth = None;
    let mut feature_depth = None;
    let mut ext_list_depth = None;
    let mut target_ext_depth = None;
    let mut feature_list_depth = None;
    let mut target_feature_list_seen = false;
    let mut properties = None;
    let mut feature_values = Vec::new();
    let mut feature_name_bytes = 0usize;

    loop {
        bump(&mut events, limits.max_events(), "event count")?;
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        reject_unsafe_event(&event)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                if root_closed {
                    return Err(invalid("content follows the closed workbook root"));
                }
                check_attribute_count(&element, limits)?;
                let local = element.local_name();
                let dialect = spreadsheet_dialect(&namespace);
                let core = dialect.is_some() && dialect == root_dialect;
                if depth == 0 {
                    if root_seen || dialect.is_none() || local.as_ref() != b"workbook" {
                        return Err(invalid(
                            "calculation metadata parser requires a workbook root",
                        ));
                    }
                    root_seen = true;
                    root_dialect = dialect;
                } else if depth == 1 && dialect.is_some() && !core {
                    return Err(invalid("workbook child uses a mixed SpreadsheetML dialect"));
                } else if ext_list_depth == Some(depth)
                    && dialect.is_some()
                    && !core
                    && local.as_ref() == b"ext"
                {
                    return Err(invalid(
                        "calculation extension uses a mixed SpreadsheetML dialect",
                    ));
                } else if calc_depth.is_some_and(|value| depth >= value) {
                    return Err(invalid("calcPr is a leaf element"));
                } else if feature_depth.is_some_and(|value| depth >= value) {
                    return Err(invalid("calculation feature is a leaf element"));
                } else if depth == 1 && core && local.as_ref() == b"calcPr" {
                    if properties.is_some() {
                        return Err(invalid("duplicate workbook calcPr element"));
                    }
                    properties = Some(parse_calc_attributes(
                        &element,
                        decoder,
                        &resolver,
                        limits,
                        strict_calc_attributes,
                    )?);
                    calc_depth = Some(depth + 1);
                } else if depth == 1 && core && local.as_ref() == b"extLst" {
                    if ext_list_depth.is_some() {
                        return Err(invalid("duplicate workbook extLst element"));
                    }
                    ext_list_depth = Some(depth + 1);
                } else if ext_list_depth == Some(depth)
                    && core
                    && local.as_ref() == b"ext"
                    && extension_uri(&element, decoder, &resolver, limits)?.as_deref()
                        == Some(CALC_FEATURES_EXTENSION_URI)
                {
                    if target_ext_depth.is_some() || target_feature_list_seen {
                        return Err(invalid("duplicate calculation-features extension"));
                    }
                    target_ext_depth = Some(depth + 1);
                } else if target_ext_depth == Some(depth)
                    && is_xcalcf(&namespace, local.as_ref(), b"calcFeatures")
                {
                    if target_feature_list_seen {
                        return Err(invalid("duplicate calcFeatures payload"));
                    }
                    no_attributes(&element, &resolver, limits, "calcFeatures")?;
                    target_feature_list_seen = true;
                    feature_list_depth = Some(depth + 1);
                } else if feature_list_depth == Some(depth)
                    && is_xcalcf(&namespace, local.as_ref(), b"feature")
                {
                    let feature = parse_feature(&element, decoder, &resolver, limits)?;
                    push_feature(
                        &mut feature_values,
                        &mut feature_name_bytes,
                        feature,
                        limits,
                    )?;
                    feature_depth = Some(depth + 1);
                } else if target_ext_depth.is_some_and(|value| depth >= value) {
                    return Err(invalid(
                        "calculation-features extension has mixed opaque payload",
                    ));
                } else if feature_list_depth.is_some_and(|value| depth >= value) {
                    return Err(invalid("unexpected calcFeatures child"));
                }
                if depth == 1 && core {
                    validate_core_order(local.as_ref(), &mut last_core_rank)?;
                }
                if depth >= limits.max_depth() {
                    return Err(invalid("workbook calculation metadata nesting is too deep"));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("XML depth overflow"))?;
            },
            Event::Empty(element) => {
                if root_closed || depth == 0 {
                    return Err(invalid(
                        "workbook root cannot be empty or followed by content",
                    ));
                }
                check_attribute_count(&element, limits)?;
                let local = element.local_name();
                let dialect = spreadsheet_dialect(&namespace);
                let core = dialect.is_some() && dialect == root_dialect;
                if depth == 1 && dialect.is_some() && !core {
                    return Err(invalid("workbook child uses a mixed SpreadsheetML dialect"));
                }
                if ext_list_depth == Some(depth)
                    && dialect.is_some()
                    && !core
                    && local.as_ref() == b"ext"
                {
                    return Err(invalid(
                        "calculation extension uses a mixed SpreadsheetML dialect",
                    ));
                }
                if depth == 1 && core {
                    validate_core_order(local.as_ref(), &mut last_core_rank)?;
                }
                if calc_depth.is_some_and(|value| depth >= value) {
                    return Err(invalid("calcPr is a leaf element"));
                }
                if feature_depth.is_some_and(|value| depth >= value) {
                    return Err(invalid("calculation feature is a leaf element"));
                }
                if depth == 1 && core && local.as_ref() == b"calcPr" {
                    if properties.is_some() {
                        return Err(invalid("duplicate workbook calcPr element"));
                    }
                    properties = Some(parse_calc_attributes(
                        &element,
                        decoder,
                        &resolver,
                        limits,
                        strict_calc_attributes,
                    )?);
                } else if depth == 1 && core && local.as_ref() == b"extLst" {
                    if ext_list_depth.is_some() {
                        return Err(invalid("duplicate workbook extLst element"));
                    }
                    ext_list_depth = Some(usize::MAX);
                } else if ext_list_depth == Some(depth)
                    && core
                    && local.as_ref() == b"ext"
                    && extension_uri(&element, decoder, &resolver, limits)?.as_deref()
                        == Some(CALC_FEATURES_EXTENSION_URI)
                {
                    return Err(invalid("calculation-features extension has no payload"));
                } else if target_ext_depth == Some(depth)
                    && is_xcalcf(&namespace, local.as_ref(), b"calcFeatures")
                {
                    return Err(invalid("calcFeatures must contain at least one feature"));
                } else if feature_list_depth == Some(depth)
                    && is_xcalcf(&namespace, local.as_ref(), b"feature")
                {
                    let feature = parse_feature(&element, decoder, &resolver, limits)?;
                    push_feature(
                        &mut feature_values,
                        &mut feature_name_bytes,
                        feature,
                        limits,
                    )?;
                } else if target_ext_depth.is_some_and(|value| depth >= value) {
                    return Err(invalid(
                        "calculation-features extension has mixed opaque payload",
                    ));
                } else if feature_list_depth.is_some_and(|value| depth >= value) {
                    return Err(invalid("unexpected calcFeatures child"));
                }
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid("unexpected workbook end element"));
                }
                if calc_depth == Some(depth) {
                    calc_depth = None;
                }
                if feature_depth == Some(depth) {
                    feature_depth = None;
                }
                if feature_list_depth == Some(depth) {
                    feature_list_depth = None;
                }
                if target_ext_depth == Some(depth) {
                    if !target_feature_list_seen || feature_values.is_empty() {
                        return Err(invalid(
                            "calculation-features extension requires nonempty calcFeatures",
                        ));
                    }
                    target_ext_depth = None;
                }
                if ext_list_depth == Some(depth) {
                    ext_list_depth = Some(usize::MAX);
                }
                depth -= 1;
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Text(text) => {
                let non_whitespace = !text.decode().map_err(xml_error)?.trim().is_empty();
                if depth == 0 && non_whitespace {
                    return Err(invalid("text outside workbook root"));
                }
                if calc_depth.is_some_and(|value| depth >= value) && non_whitespace {
                    return Err(invalid("calcPr cannot contain text"));
                }
                if feature_depth.is_some_and(|value| depth >= value) && non_whitespace {
                    return Err(invalid("calculation feature cannot contain text"));
                }
                if target_ext_depth.is_some_and(|value| depth >= value)
                    && feature_list_depth.is_none()
                    && non_whitespace
                {
                    return Err(invalid(
                        "calculation-features extension has mixed opaque payload",
                    ));
                }
            },
            Event::CData(_) if calc_depth.is_some() || feature_depth.is_some() => {
                return Err(invalid("calculation metadata leaf cannot contain CDATA"));
            },
            Event::GeneralRef(_) if calc_depth.is_some() || feature_depth.is_some() => {
                return Err(invalid(
                    "calculation metadata leaf cannot contain entity references",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !root_seen || !root_closed || depth != 0 || calc_depth.is_some() || feature_depth.is_some() {
        return Err(invalid("unterminated workbook XML"));
    }
    let features = if target_feature_list_seen {
        Some(Features::try_from_vec(feature_values)?)
    } else {
        None
    };
    Ok((properties, features))
}

fn parse_calc_attributes(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    limits: &Limits,
    strict: bool,
) -> Result<Properties> {
    check_attribute_count(element, limits)?;
    let mut builder = Properties::builder();
    let mut seen = [false; 13];
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Unbound) {
            let known_local = calc_attribute_index(local.as_ref()).is_some();
            let core_namespace = matches!(
                namespace,
                ResolveResult::Bound(Namespace(value))
                    if value == SPREADSHEETML_NAMESPACE || value == STRICT_SPREADSHEETML_NAMESPACE
            );
            if strict || known_local || core_namespace {
                return Err(invalid(format!(
                    "unknown or spoofed namespaced calcPr attribute '{}'",
                    String::from_utf8_lossy(attribute.key.as_ref()),
                )));
            }
            if !matches!(namespace, ResolveResult::Bound(_)) {
                return Err(invalid(format!(
                    "unbound calcPr attribute prefix in '{}'",
                    String::from_utf8_lossy(attribute.key.as_ref()),
                )));
            }
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(xml_error)?;
        let value = collapse_whitespace(&value);
        let Some((slot, name)) = calc_attribute_index(local.as_ref()) else {
            return Err(invalid(format!(
                "unknown calcPr attribute '{}'",
                String::from_utf8_lossy(local.as_ref()),
            )));
        };
        if std::mem::replace(&mut seen[slot], true) {
            return Err(invalid(format!("duplicate {name} attribute")));
        }
        builder = match slot {
            0 => builder.calculation_id(Some(parse_u32(&value, name)?)),
            1 => builder.calculation_mode(Some(Mode::parse(&value)?)),
            2 => builder.full_calculation_on_load(Some(parse_bool(&value, name)?)),
            3 => builder.reference_mode(Some(ReferenceMode::parse(&value)?)),
            4 => builder.iterative_calculation(Some(parse_bool(&value, name)?)),
            5 => builder.iteration_count(Some(parse_u32(&value, name)?)),
            6 => builder.iteration_delta(Some(parse_delta(&value)?))?,
            7 => builder.full_precision(Some(parse_bool(&value, name)?)),
            8 => builder.calculation_completed(Some(parse_bool(&value, name)?)),
            9 => builder.calculate_on_save(Some(parse_bool(&value, name)?)),
            10 => builder.concurrent_calculation(Some(parse_bool(&value, name)?)),
            11 => builder.concurrent_manual_count(Some(parse_u32(&value, name)?)),
            12 => builder.force_full_calculation(Some(parse_bool(&value, name)?)),
            _ => unreachable!(),
        };
    }
    Ok(builder.build())
}

pub(super) fn calc_attribute_index(name: &[u8]) -> Option<(usize, &'static str)> {
    Some(match name {
        b"calcId" => (0, "calcId"),
        b"calcMode" => (1, "calcMode"),
        b"fullCalcOnLoad" => (2, "fullCalcOnLoad"),
        b"refMode" => (3, "refMode"),
        b"iterate" => (4, "iterate"),
        b"iterateCount" => (5, "iterateCount"),
        b"iterateDelta" => (6, "iterateDelta"),
        b"fullPrecision" => (7, "fullPrecision"),
        b"calcCompleted" => (8, "calcCompleted"),
        b"calcOnSave" => (9, "calcOnSave"),
        b"concurrentCalc" => (10, "concurrentCalc"),
        b"concurrentManualCount" => (11, "concurrentManualCount"),
        b"forceFullCalc" => (12, "forceFullCalc"),
        _ => return None,
    })
}

fn extension_uri(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    limits: &Limits,
) -> Result<Option<String>> {
    check_attribute_count(element, limits)?;
    let mut uri = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Unbound) && local.as_ref() == b"uri" {
            if uri.is_some() {
                return Err(invalid("duplicate ext uri attribute"));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(xml_error)?;
            uri = Some(collapse_whitespace(&value).into_owned());
        }
    }
    Ok(uri)
}

fn parse_feature(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    limits: &Limits,
) -> Result<Feature> {
    check_attribute_count(element, limits)?;
    let mut name = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Unbound) || local.as_ref() != b"name" {
            return Err(invalid(
                "feature accepts only an unqualified name attribute",
            ));
        }
        if name.is_some() {
            return Err(invalid("duplicate feature name attribute"));
        }
        if attribute.value.len() > limits.max_feature_name_bytes() {
            return Err(invalid("calculation feature name exceeds byte limit"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(xml_error)?;
        if value.len() > limits.max_feature_name_bytes() {
            return Err(invalid("calculation feature name exceeds byte limit"));
        }
        name = Some(Feature::new(value.into_owned())?);
    }
    name.ok_or_else(|| invalid("feature requires an unqualified name attribute"))
}

fn push_feature(
    values: &mut Vec<Feature>,
    total: &mut usize,
    feature: Feature,
    limits: &Limits,
) -> Result<()> {
    if values.len() >= limits.max_features() {
        return Err(invalid("calculation feature count exceeds limit"));
    }
    *total = total
        .checked_add(feature.as_str().len())
        .ok_or_else(|| invalid("calculation feature name byte count overflow"))?;
    if *total > limits.max_feature_names_bytes() {
        return Err(invalid(
            "calculation feature names exceed aggregate byte limit",
        ));
    }
    values.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "calculation features",
        source,
    })?;
    values.push(feature);
    Ok(())
}

fn no_attributes(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    limits: &Limits,
    label: &str,
) -> Result<()> {
    check_attribute_count(element, limits)?;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if !is_namespace_declaration(attribute.key.as_ref()) {
            let (namespace, local) = resolver.resolve_attribute(attribute.key);
            return Err(invalid(format!(
                "{label} has unexpected attribute '{{{:?}}}{}'",
                namespace,
                String::from_utf8_lossy(local.as_ref())
            )));
        }
    }
    Ok(())
}

fn check_attribute_count(element: &BytesStart<'_>, limits: &Limits) -> Result<()> {
    let mut count = 0usize;
    for attribute in element.attributes().with_checks(false) {
        attribute.map_err(xml_error)?;
        bump(&mut count, limits.max_attributes(), "attribute count")?;
    }
    Ok(())
}

fn bump(value: &mut usize, max: usize, label: &str) -> Result<()> {
    *value = value
        .checked_add(1)
        .ok_or_else(|| invalid(format!("calculation metadata {label} overflow")))?;
    if *value > max {
        return Err(invalid(format!(
            "calculation metadata exceeds {label} limit"
        )));
    }
    Ok(())
}

fn is_xcalcf(namespace: &ResolveResult<'_>, local: &[u8], expected: &[u8]) -> bool {
    local == expected
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == CALC_FEATURES_NAMESPACE)
}

pub(super) fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!("invalid calcPr {name} boolean '{value}'"))),
    }
}

fn parse_u32(value: &str, name: &str) -> Result<u32> {
    if value.is_empty() || !value.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(invalid(format!("invalid calcPr {name} value '{value}'")));
    }
    value
        .parse::<u32>()
        .map_err(|_| invalid(format!("invalid calcPr {name} value '{value}'")))
}

fn parse_delta(value: &str) -> Result<f64> {
    match value {
        "INF" => Ok(f64::INFINITY),
        "-INF" => Ok(f64::NEG_INFINITY),
        "NaN" => Ok(f64::NAN),
        _ if valid_finite_double_lexical(value) => value
            .parse::<f64>()
            .map_err(|_| invalid(format!("invalid calcPr iterateDelta '{value}'"))),
        _ => Err(invalid(format!("invalid calcPr iterateDelta '{value}'"))),
    }
}

fn valid_finite_double_lexical(value: &str) -> bool {
    let value = value
        .strip_prefix('+')
        .or_else(|| value.strip_prefix('-'))
        .unwrap_or(value);
    let (mantissa, exponent) = value.find(['e', 'E']).map_or((value, None), |index| {
        (&value[..index], Some(&value[index + 1..]))
    });
    if let Some(exponent) = exponent {
        let digits = exponent
            .strip_prefix('+')
            .or_else(|| exponent.strip_prefix('-'))
            .unwrap_or(exponent);
        if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
            return false;
        }
    }
    let mut parts = mantissa.split('.');
    let before = parts.next().unwrap_or_default();
    let after = parts.next();
    if parts.next().is_some() {
        return false;
    }
    before.bytes().all(|byte| byte.is_ascii_digit())
        && after.is_none_or(|digits| digits.bytes().all(|byte| byte.is_ascii_digit()))
        && (!before.is_empty() || after.is_some_and(|digits| !digits.is_empty()))
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Dialect {
    Transitional,
    Strict,
}

fn spreadsheet_dialect(namespace: &ResolveResult<'_>) -> Option<Dialect> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == SPREADSHEETML_NAMESPACE => {
            Some(Dialect::Transitional)
        },
        ResolveResult::Bound(Namespace(value)) if *value == STRICT_SPREADSHEETML_NAMESPACE => {
            Some(Dialect::Strict)
        },
        _ => None,
    }
}

fn validate_core_order(local: &[u8], last: &mut Option<u8>) -> Result<()> {
    let Some(rank) = workbook_child_rank(local) else {
        return Ok(());
    };
    if last.is_some_and(|previous| rank < previous) {
        return Err(invalid("workbook core children are out of schema order"));
    }
    *last = Some(rank);
    Ok(())
}

pub(super) fn workbook_child_rank(local: &[u8]) -> Option<u8> {
    Some(match local {
        b"fileVersion" => 0,
        b"fileSharing" => 1,
        b"workbookPr" => 2,
        b"workbookProtection" => 3,
        b"bookViews" => 4,
        b"sheets" => 5,
        b"functionGroups" => 6,
        b"externalReferences" => 7,
        b"definedNames" => 8,
        b"calcPr" => 9,
        b"oleSize" => 10,
        b"customWorkbookViews" => 11,
        b"pivotCaches" => 12,
        b"smartTagPr" => 13,
        b"smartTagTypes" => 14,
        b"webPublishing" => 15,
        b"fileRecoveryPr" => 16,
        b"webPublishObjects" => 17,
        b"extLst" => 18,
        _ => return None,
    })
}

pub(super) fn collapse_whitespace(value: &str) -> Cow<'_, str> {
    let pieces: Vec<_> = value
        .split([' ', '\t', '\r', '\n'])
        .filter(|piece| !piece.is_empty())
        .collect();
    match pieces.as_slice() {
        [] => Cow::Borrowed(""),
        [only] if only.len() == value.len() => Cow::Borrowed(value),
        _ => Cow::Owned(pieces.join(" ")),
    }
}

pub(super) fn reject_unsafe_event(event: &Event<'_>) -> Result<()> {
    if matches!(event, Event::DocType(_) | Event::PI(_)) {
        return Err(invalid("DTD and processing instructions are rejected"));
    }
    if let Event::GeneralRef(reference) = event {
        let name = reference.decode().map_err(xml_error)?;
        if !matches!(name.as_ref(), "amp" | "lt" | "gt" | "apos" | "quot") && !name.starts_with('#')
        {
            return Err(invalid("custom XML entities are rejected"));
        }
    }
    Ok(())
}

pub(super) fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(format!(
        "invalid workbook calculation metadata XML: {error}"
    )))
}
