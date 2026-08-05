//! Bounded SpreadsheetML/MCE codec for workbook calculation properties.

use crate::error::{Error, Result, invalid};
use crate::raw::namespace::is_spreadsheetml_name;
use litchi_ooxml_common::mce::{Capabilities, Limits, process_markup_compatibility};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::{Mode, Properties, ReferenceMode};

pub(super) const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_EVENTS: usize = 1_000_000;

/// Parse the workbook's direct `calcPr` child without executing calculations.
pub fn parse(xml: &[u8]) -> Result<Option<Properties>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("workbook calcPr XML exceeds size limit"));
    }
    let limits = Limits {
        max_input_bytes: MAX_XML_BYTES,
        max_output_bytes: MAX_XML_BYTES,
        max_depth: MAX_DEPTH,
        ..Limits::default()
    };
    let processed = process_markup_compatibility(xml, &Capabilities::default(), &limits)?;
    if processed.xml.len() > MAX_XML_BYTES {
        return Err(invalid("processed workbook calcPr XML exceeds size limit"));
    }
    let mut reader = NsReader::from_reader(processed.xml.as_ref());
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut events = 0usize;
    let mut root_seen = false;
    let mut leaf_depth = None;
    let mut properties = None;

    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("workbook calcPr XML event count overflow"))?;
        if events > MAX_EVENTS {
            return Err(invalid("workbook calcPr XML exceeds event limit"));
        }
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        reject_unsafe_event(&event)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                let local = element.local_name();
                let core = is_spreadsheetml_name(&namespace, element.name(), local.as_ref());
                if depth == 0 {
                    if root_seen || !core || local.as_ref() != b"workbook" {
                        return Err(invalid("calcPr parser requires a workbook root"));
                    }
                    root_seen = true;
                } else if leaf_depth.is_some_and(|value| depth >= value) {
                    return Err(invalid("calcPr is a leaf element"));
                } else if depth == 1 && core && local.as_ref() == b"calcPr" {
                    if properties.is_some() {
                        return Err(invalid("duplicate workbook calcPr element"));
                    }
                    properties = Some(parse_attributes(&element, decoder, &resolver)?.finish());
                    leaf_depth = Some(depth + 1);
                }
                if depth >= MAX_DEPTH {
                    return Err(invalid("workbook calcPr XML nesting is too deep"));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("XML depth overflow"))?;
            },
            Event::Empty(element) => {
                let local = element.local_name();
                let core = is_spreadsheetml_name(&namespace, element.name(), local.as_ref());
                if depth == 0 {
                    return Err(invalid("workbook root cannot be empty"));
                }
                if leaf_depth.is_some_and(|value| depth >= value) {
                    return Err(invalid("calcPr is a leaf element"));
                }
                if depth == 1 && core && local.as_ref() == b"calcPr" {
                    if properties.is_some() {
                        return Err(invalid("duplicate workbook calcPr element"));
                    }
                    properties = Some(parse_attributes(&element, decoder, &resolver)?.finish());
                }
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid("unexpected workbook end element"));
                }
                if leaf_depth == Some(depth) {
                    leaf_depth = None;
                }
                depth -= 1;
            },
            Event::Text(text)
                if leaf_depth.is_some_and(|value| depth >= value)
                    && !text.decode().map_err(xml_error)?.trim().is_empty() =>
            {
                return Err(invalid("calcPr cannot contain text"));
            },
            Event::CData(_) if leaf_depth.is_some_and(|value| depth >= value) => {
                return Err(invalid("calcPr cannot contain CDATA"));
            },
            Event::GeneralRef(_) if leaf_depth.is_some_and(|value| depth >= value) => {
                return Err(invalid("calcPr cannot contain entity references"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !root_seen || depth != 0 || leaf_depth.is_some() {
        return Err(invalid("unterminated workbook XML"));
    }
    Ok(properties)
}

#[derive(Default)]
struct Builder {
    calculation_id: Option<u32>,
    calculation_mode: Option<Mode>,
    full_calculation_on_load: Option<bool>,
    reference_mode: Option<ReferenceMode>,
    iterative_calculation: Option<bool>,
    iteration_count: Option<u32>,
    iteration_delta: Option<f64>,
    full_precision: Option<bool>,
    calculation_completed: Option<bool>,
    calculate_on_save: Option<bool>,
    concurrent_calculation: Option<bool>,
    concurrent_manual_count: Option<u32>,
    force_full_calculation: Option<bool>,
}

impl Builder {
    fn finish(self) -> Properties {
        Properties::new(
            self.calculation_id.unwrap_or(0),
            self.calculation_mode.unwrap_or_default(),
            self.full_calculation_on_load.unwrap_or(false),
            self.reference_mode.unwrap_or_default(),
            self.iterative_calculation.unwrap_or(false),
            self.iteration_count.unwrap_or(100),
            self.iteration_delta.unwrap_or(0.001),
            self.full_precision.unwrap_or(true),
            self.calculation_completed.unwrap_or(true),
            self.calculate_on_save.unwrap_or(true),
            self.concurrent_calculation.unwrap_or(true),
            self.concurrent_manual_count,
            self.force_full_calculation.unwrap_or(false),
        )
    }
}

fn parse_attributes(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Builder> {
    let mut builder = Builder::default();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Unbound) {
            return Err(invalid(format!(
                "unknown namespaced calcPr attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref()),
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(xml_error)?;
        match local.as_ref() {
            b"calcId" => set_once(
                &mut builder.calculation_id,
                parse_u32(&value, "calcId")?,
                "calcId",
            )?,
            b"calcMode" => set_once(
                &mut builder.calculation_mode,
                Mode::parse(&value)?,
                "calcMode",
            )?,
            b"fullCalcOnLoad" => set_once(
                &mut builder.full_calculation_on_load,
                parse_bool(&value, "fullCalcOnLoad")?,
                "fullCalcOnLoad",
            )?,
            b"refMode" => set_once(
                &mut builder.reference_mode,
                ReferenceMode::parse(&value)?,
                "refMode",
            )?,
            b"iterate" => set_once(
                &mut builder.iterative_calculation,
                parse_bool(&value, "iterate")?,
                "iterate",
            )?,
            b"iterateCount" => set_once(
                &mut builder.iteration_count,
                parse_u32(&value, "iterateCount")?,
                "iterateCount",
            )?,
            b"iterateDelta" => set_once(
                &mut builder.iteration_delta,
                parse_delta(&value)?,
                "iterateDelta",
            )?,
            b"fullPrecision" => set_once(
                &mut builder.full_precision,
                parse_bool(&value, "fullPrecision")?,
                "fullPrecision",
            )?,
            b"calcCompleted" => set_once(
                &mut builder.calculation_completed,
                parse_bool(&value, "calcCompleted")?,
                "calcCompleted",
            )?,
            b"calcOnSave" => set_once(
                &mut builder.calculate_on_save,
                parse_bool(&value, "calcOnSave")?,
                "calcOnSave",
            )?,
            b"concurrentCalc" => set_once(
                &mut builder.concurrent_calculation,
                parse_bool(&value, "concurrentCalc")?,
                "concurrentCalc",
            )?,
            b"concurrentManualCount" => set_once(
                &mut builder.concurrent_manual_count,
                parse_u32(&value, "concurrentManualCount")?,
                "concurrentManualCount",
            )?,
            b"forceFullCalc" => set_once(
                &mut builder.force_full_calculation,
                parse_bool(&value, "forceFullCalc")?,
                "forceFullCalc",
            )?,
            name => {
                return Err(invalid(format!(
                    "unknown calcPr attribute '{}'",
                    String::from_utf8_lossy(name),
                )));
            },
        }
    }
    Ok(builder)
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn set_once<T>(slot: &mut Option<T>, value: T, name: &str) -> Result<()> {
    if slot.is_some() {
        return Err(invalid(format!("duplicate {name} attribute")));
    }
    *slot = Some(value);
    Ok(())
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!("invalid calcPr {name} boolean '{value}'"))),
    }
}

fn parse_u32(value: &str, name: &str) -> Result<u32> {
    value
        .parse::<u32>()
        .map_err(|_| invalid(format!("invalid calcPr {name} value '{value}'")))
}

fn parse_delta(value: &str) -> Result<f64> {
    let parsed = value
        .parse::<f64>()
        .map_err(|_| invalid(format!("invalid calcPr iterateDelta '{value}'")))?;
    if !parsed.is_finite() || parsed < 0.0 {
        return Err(invalid(
            "calcPr iterateDelta must be finite and non-negative",
        ));
    }
    Ok(parsed)
}

fn reject_unsafe_event(event: &Event<'_>) -> Result<()> {
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

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(format!(
        "invalid workbook calcPr XML: {error}"
    )))
}
