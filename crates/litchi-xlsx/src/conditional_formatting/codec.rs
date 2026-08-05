//! Bounded SpreadsheetML/MCE codec and differential-format association.

use crate::color::Rgb;
use crate::{Error, Result};

use super::model::{
    Association, Axis, Color, ColorRole, ColorScale, Component, DataBar, Differential,
    DifferentialRef, Direction, Formatting, IconSet, IconSet14, Icons, Kind, NamedColor,
    NumberFormat, Operator, Payload, Period, Range, Rule, Source, Value, ValueKind,
};

use litchi_ooxml_common::mce::{Capabilities, Limits, Name, process_markup_compatibility};

use quick_xml::encoding::Decoder;

use quick_xml::events::{BytesStart, Event};

use quick_xml::name::ResolveResult;

use quick_xml::reader::NsReader;

use quick_xml::{Writer, XmlVersion};

use std::collections::{HashMap, HashSet};

use smallvec::SmallVec;
use std::fmt;

const CORE: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";

const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";

const X14: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

const XM: &[u8] = b"http://schemas.microsoft.com/office/excel/2006/main";

const MAX_FORMULA_BYTES: usize = 1024 * 1024;

const MAX_FRAGMENT_BYTES: usize = 16 * 1024 * 1024;

const MAX_RULES: usize = 1_000_000;

#[derive(Debug)]
pub(crate) struct Captured {
    pub(crate) source: Source,
    pub(crate) prefix: Vec<u8>,
    pub(crate) bytes: Vec<u8>,
}

pub fn parse_conditional_formattings(
    xml: &[u8],
    differential_format_count: usize,
) -> Result<Vec<Formatting>> {
    let mut capabilities = Capabilities::default();
    capabilities.preserve_extension_element(Name {
        namespace: String::from_utf8_lossy(X14).into_owned(),
        local_name: "conditionalFormattings".to_owned(),
    });
    let _validated = process_markup_compatibility(xml, &capabilities, &Limits::default())?;
    // Parse the selected vocabulary from the source bytes so formula character data remains
    // byte-for-byte semantically opaque; the generic MCE writer may normalize XML entities.
    let captured = capture_conditional_formatting(xml)?;
    let mut values = Vec::with_capacity(captured.len());
    let mut total_rules = 0usize;
    for fragment in captured {
        let value = parse_container(&fragment)?;
        total_rules = total_rules
            .checked_add(value.rules.len())
            .ok_or_else(|| invalid("conditional-formatting rule count overflow"))?;
        if total_rules > MAX_RULES {
            return Err(invalid("too many conditional-formatting rules"));
        }
        values.push(value);
    }
    validate_and_associate(&mut values, differential_format_count)?;
    Ok(values)
}

pub fn parse_differential_formats(xml: &[u8]) -> Result<Vec<Differential>> {
    let processed = litchi_ooxml_common::mce::process_ooxml(xml)?;
    let Some(fragment) = capture_first(processed.as_ref(), CORE, STRICT, b"dxfs")? else {
        return Ok(Vec::new());
    };
    let wrapped = wrap(&fragment.prefix, &fragment.bytes);
    let mut reader = NsReader::from_reader(wrapped.as_slice());
    let mut expected = None;
    let mut values = Vec::new();
    let mut capture: Option<(usize, Vec<u8>, Writer<Vec<u8>>)> = None;
    let mut depth = 0usize;
    loop {
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        if let Some((capture_depth, _, writer)) = capture.as_mut() {
            writer.write_event(event.clone()).map_err(xml_error)?;
            match event {
                Event::Start(_) => *capture_depth += 1,
                Event::End(_) => *capture_depth -= 1,
                _ => {},
            }
            if *capture_depth == 0 {
                let (_, _, writer) = capture
                    .take()
                    .ok_or_else(|| invalid("missing differential-format capture state"))?;
                let raw = writer.into_inner();
                if raw.len() > MAX_FRAGMENT_BYTES {
                    return Err(invalid("differential format is too large"));
                }
                values.push(parse_dxf(&raw)?);
            }
            continue;
        }
        match event {
            Event::Start(element) => {
                depth += 1;
                if spreadsheet(&namespace) && element.local_name().as_ref() == b"dxfs" {
                    expected = optional_u32(&element, b"count", decoder)?;
                } else if spreadsheet(&namespace) && element.local_name().as_ref() == b"dxf" {
                    depth -= 1;
                    let mut writer = Writer::new(Vec::new());
                    writer
                        .write_event(Event::Start(element))
                        .map_err(xml_error)?;
                    capture = Some((1, Vec::new(), writer));
                }
            },
            Event::Empty(element)
                if spreadsheet(&namespace) && element.local_name().as_ref() == b"dxf" =>
            {
                let mut writer = Writer::new(Vec::new());
                writer
                    .write_event(Event::Empty(element))
                    .map_err(xml_error)?;
                values.push(parse_dxf(&writer.into_inner())?);
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("invalid dxfs nesting"))?
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || capture.is_some() {
        return Err(invalid("unterminated dxfs XML"));
    }
    if expected.is_some_and(|count| count as usize != values.len()) {
        return Err(invalid("dxfs count does not match dxf elements"));
    }
    Ok(values)
}

/// In-flight capture state for a single `conditionalFormatting` element.
type CaptureState = Option<(usize, Source, Vec<u8>, Writer<Vec<u8>>)>;

pub(crate) fn capture_conditional_formatting(xml: &[u8]) -> Result<Vec<Captured>> {
    let mut reader = NsReader::from_reader(xml);
    let mut values = Vec::new();
    let mut capture: CaptureState = None;
    loop {
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        if let Some((depth, _, _, writer)) = capture.as_mut() {
            writer.write_event(event.clone()).map_err(xml_error)?;
            match event {
                Event::Start(_) => *depth += 1,
                Event::End(_) => *depth -= 1,
                _ => {},
            }
            if *depth == 0 {
                let (_, source, prefix, writer) = capture
                    .take()
                    .ok_or_else(|| invalid("missing conditional-formatting capture state"))?;
                values.push(Captured {
                    source,
                    prefix,
                    bytes: writer.into_inner(),
                });
            }
            continue;
        }
        match event {
            Event::Start(element)
                if element.local_name().as_ref() == b"conditionalFormatting"
                    && (spreadsheet(&namespace) || exact(&namespace, X14)) =>
            {
                let source = if exact(&namespace, X14) {
                    Source::Office2010
                } else {
                    Source::Core
                };
                let prefix = prefix(element.name().as_ref());
                let mut writer = Writer::new(Vec::new());
                writer
                    .write_event(Event::Start(element))
                    .map_err(xml_error)?;
                capture = Some((1, source, prefix, writer));
            },
            Event::Empty(element)
                if element.local_name().as_ref() == b"conditionalFormatting"
                    && (spreadsheet(&namespace) || exact(&namespace, X14)) =>
            {
                let source = if exact(&namespace, X14) {
                    Source::Office2010
                } else {
                    Source::Core
                };
                let prefix = prefix(element.name().as_ref());
                let mut writer = Writer::new(Vec::new());
                writer
                    .write_event(Event::Empty(element))
                    .map_err(xml_error)?;
                values.push(Captured {
                    source,
                    prefix,
                    bytes: writer.into_inner(),
                });
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if capture.is_some() {
        return Err(invalid("unterminated conditionalFormatting"));
    }
    Ok(values)
}

fn parse_container(fragment: &Captured) -> Result<Formatting> {
    let wrapped = wrap(&fragment.prefix, &fragment.bytes);
    let mut reader = NsReader::from_reader(wrapped.as_slice());
    let mut ranges = Vec::new();
    let mut pivot = false;
    let mut rules = Vec::new();
    let mut capture: Option<(usize, Vec<String>, Writer<Vec<u8>>)> = None;
    let mut captured_formula: Option<String> = None;
    let mut sqref_text: Option<(usize, String)> = None;
    let mut depth = 0usize;
    loop {
        let decoder = reader.decoder();
        let event_start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let event_end = reader.buffer_position() as usize;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        if let Some((capture_depth, formulas, writer)) = capture.as_mut() {
            if matches!(&event, Event::Start(element)
                if *capture_depth == 1 && ((spreadsheet(&namespace) && element.local_name().as_ref() == b"formula")
                    || (exact(&namespace, XM) && element.local_name().as_ref() == b"f")))
            {
                captured_formula = Some(String::new());
            } else if let Event::Text(_) = &event
                && let Some(value) = captured_formula.as_mut()
            {
                value.push_str(&decode_raw_text(&wrapped[event_start..event_end])?);
                if value.len() > MAX_FORMULA_BYTES {
                    return Err(invalid("conditional-formatting formula is too large"));
                }
            } else if let Event::GeneralRef(_) = &event
                && let Some(value) = captured_formula.as_mut()
            {
                value.push_str(&decode_raw_text(&wrapped[event_start..event_end])?);
                if value.len() > MAX_FORMULA_BYTES {
                    return Err(invalid("conditional-formatting formula is too large"));
                }
            } else if let Event::CData(text) = &event
                && let Some(value) = captured_formula.as_mut()
            {
                value.push_str(&text.decode().map_err(xml_error)?);
                if value.len() > MAX_FORMULA_BYTES {
                    return Err(invalid("conditional-formatting formula is too large"));
                }
            } else if matches!(&event, Event::End(element)
                if *capture_depth == 2 && ((spreadsheet(&namespace) && element.local_name().as_ref() == b"formula")
                    || (exact(&namespace, XM) && element.local_name().as_ref() == b"f")))
            {
                formulas.push(
                    captured_formula
                        .take()
                        .ok_or_else(|| invalid("formula end without matching start"))?,
                );
                if formulas.len() > 3 {
                    return Err(invalid("cfRule has more than three formulas"));
                }
            }
            writer.write_event(event.clone()).map_err(xml_error)?;
            match event {
                Event::Start(_) => *capture_depth += 1,
                Event::End(_) => *capture_depth -= 1,
                _ => {},
            }
            if *capture_depth == 0 {
                let (_, formulas, writer) = capture
                    .take()
                    .ok_or_else(|| invalid("missing cfRule capture state"))?;
                let mut rule = parse_rule(&writer.into_inner(), fragment.source)?;
                rule.formulas = formulas.into_iter().collect();
                validate_rule_shape(&rule)?;
                rules.push(rule);
            }
            continue;
        }
        match event {
            Event::Start(element) => {
                depth += 1;
                if element.local_name().as_ref() == b"conditionalFormatting"
                    && ((fragment.source == Source::Core && spreadsheet(&namespace))
                        || (fragment.source == Source::Office2010 && exact(&namespace, X14)))
                {
                    if fragment.source == Source::Core {
                        let raw = required_attr(&element, b"sqref", decoder)?;
                        ranges = parse_sqref(&raw)?;
                        pivot = optional_bool(&element, b"pivot", decoder)?.unwrap_or(false);
                    }
                } else if element.local_name().as_ref() == b"cfRule"
                    && ((fragment.source == Source::Core && spreadsheet(&namespace))
                        || (fragment.source == Source::Office2010 && exact(&namespace, X14)))
                {
                    let mut writer = Writer::new(Vec::new());
                    writer
                        .write_event(Event::Start(element))
                        .map_err(xml_error)?;
                    capture = Some((1, Vec::new(), writer));
                } else if fragment.source == Source::Office2010
                    && exact(&namespace, XM)
                    && element.local_name().as_ref() == b"sqref"
                {
                    sqref_text = Some((depth, String::new()));
                }
            },
            Event::Text(text) => {
                if let Some((_, value)) = sqref_text.as_mut() {
                    value.push_str(&decode_text(&text)?);
                }
            },
            Event::CData(text) => {
                if let Some((_, value)) = sqref_text.as_mut() {
                    value.push_str(&text.decode().map_err(xml_error)?);
                }
            },
            Event::Empty(element)
                if element.local_name().as_ref() == b"cfRule"
                    && ((fragment.source == Source::Core && spreadsheet(&namespace))
                        || (fragment.source == Source::Office2010 && exact(&namespace, X14))) =>
            {
                let mut writer = Writer::new(Vec::new());
                writer
                    .write_event(Event::Empty(element))
                    .map_err(xml_error)?;
                rules.push(parse_rule(&writer.into_inner(), fragment.source)?);
            },
            Event::End(element) => {
                if sqref_text
                    .as_ref()
                    .is_some_and(|(target, _)| *target == depth)
                    && exact(&namespace, XM)
                    && element.local_name().as_ref() == b"sqref"
                {
                    let (_, raw) = sqref_text
                        .take()
                        .ok_or_else(|| invalid("sqref end without matching start"))?;
                    ranges = parse_sqref(&raw)?;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("invalid conditionalFormatting nesting"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if capture.is_some() || sqref_text.is_some() {
        return Err(invalid("unterminated conditionalFormatting child"));
    }
    if ranges.is_empty() {
        return Err(invalid("conditionalFormatting has no sqref"));
    }
    Ok(Formatting {
        ranges,
        pivot,
        rules,
    })
}

fn parse_rule(raw: &[u8], source: Source) -> Result<Rule> {
    if raw.len() > MAX_FRAGMENT_BYTES {
        return Err(invalid("conditional-formatting rule is too large"));
    }
    let wrapped = wrap(if source == Source::Core { b"" } else { b"x14" }, raw);
    let inline_dxf = capture_first(&wrapped, CORE, STRICT, b"dxf")?
        .map(|value| parse_dxf(&value.bytes))
        .transpose()?;
    let mut reader = NsReader::from_reader(wrapped.as_slice());
    let mut rule = Rule {
        source,
        rule_type: None,
        priority: None,
        differential_format: inline_dxf.map(DifferentialRef::Inline),
        formulas: SmallVec::new(),
        stop_if_true: false,
        above_average: true,
        equal_average: false,
        percent: false,
        bottom: false,
        operator: None,
        text: None,
        time_period: None,
        rank: None,
        standard_deviations: None,
        payload: None,
        extension_id: None,
        extension_association: Association::Independent,
    };
    let mut depth = 0usize;
    let mut text_target: Option<(usize, TextTarget, String)> = None;
    let mut payload = PayloadBuilder::new(source);
    loop {
        let decoder = reader.decoder();
        let event_start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let event_end = reader.buffer_position() as usize;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                depth += 1;
                let name = element.local_name();
                if name.as_ref() == b"cfRule"
                    && ((source == Source::Core && spreadsheet(&namespace))
                        || (source == Source::Office2010 && exact(&namespace, X14)))
                {
                    parse_rule_attributes(&element, decoder, &mut rule)?;
                } else if (spreadsheet(&namespace) && name.as_ref() == b"formula")
                    || (exact(&namespace, XM) && name.as_ref() == b"f")
                {
                    let target = if exact(&namespace, XM) && payload.active.is_some() {
                        TextTarget::CfvoFormula
                    } else {
                        TextTarget::Formula
                    };
                    text_target = Some((depth, target, String::new()));
                } else if exact(&namespace, X14) && name.as_ref() == b"id" {
                    text_target = Some((depth, TextTarget::ExtensionId, String::new()));
                } else if (spreadsheet(&namespace) || exact(&namespace, X14))
                    && name.as_ref() == b"colorScale"
                {
                    payload.begin(PayloadKind::ColorScale, &element, decoder)?;
                } else if (spreadsheet(&namespace) || exact(&namespace, X14))
                    && name.as_ref() == b"dataBar"
                {
                    payload.begin(PayloadKind::DataBar, &element, decoder)?;
                } else if (spreadsheet(&namespace) || exact(&namespace, X14))
                    && name.as_ref() == b"iconSet"
                {
                    payload.begin(PayloadKind::IconSet, &element, decoder)?;
                } else if (spreadsheet(&namespace) || exact(&namespace, X14))
                    && name.as_ref() == b"cfvo"
                {
                    payload.cfvo(&element, decoder)?;
                } else if (spreadsheet(&namespace) || exact(&namespace, X14))
                    && is_color_element(name.as_ref())
                {
                    payload.color(name.as_ref(), &element, decoder)?;
                }
            },
            Event::Empty(element) => {
                let name = element.local_name();
                if (spreadsheet(&namespace) || exact(&namespace, X14)) && name.as_ref() == b"cfvo" {
                    payload.cfvo(&element, decoder)?;
                } else if (spreadsheet(&namespace) || exact(&namespace, X14))
                    && is_color_element(name.as_ref())
                {
                    payload.color(name.as_ref(), &element, decoder)?;
                }
            },
            Event::Text(_) => {
                if let Some((_, _, value)) = text_target.as_mut() {
                    value.push_str(&decode_raw_text(&wrapped[event_start..event_end])?);
                    if value.len() > MAX_FORMULA_BYTES {
                        return Err(invalid("conditional-formatting formula is too large"));
                    }
                }
            },
            Event::GeneralRef(_) => {
                if let Some((_, _, value)) = text_target.as_mut() {
                    value.push_str(&decode_raw_text(&wrapped[event_start..event_end])?);
                    if value.len() > MAX_FORMULA_BYTES {
                        return Err(invalid("conditional-formatting formula is too large"));
                    }
                }
            },
            Event::CData(text) => {
                if let Some((_, _, value)) = text_target.as_mut() {
                    value.push_str(&text.decode().map_err(xml_error)?);
                }
            },
            Event::End(element) => {
                if text_target
                    .as_ref()
                    .is_some_and(|(target, _, _)| *target == depth)
                {
                    let (_, target, value) = text_target
                        .take()
                        .ok_or_else(|| invalid("text end without matching start"))?;
                    match target {
                        TextTarget::Formula => {
                            if rule.formulas.len() == 3 {
                                return Err(invalid("cfRule has more than three formulas"));
                            }
                            rule.formulas.push(value);
                        },
                        TextTarget::ExtensionId => rule.extension_id = Some(value),
                        TextTarget::CfvoFormula => payload.set_last_formula(value)?,
                    }
                }
                if spreadsheet(&namespace) || exact(&namespace, X14) {
                    match element.local_name().as_ref() {
                        b"colorScale" | b"dataBar" | b"iconSet" => payload.end()?,
                        _ => {},
                    }
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("invalid cfRule nesting"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if source == Source::Core && rule.priority.is_none() {
        return Err(invalid("core cfRule is missing priority"));
    }
    if let Some(index) = optional_u32_from_raw(raw, b"dxfId")? {
        if rule.differential_format.is_some() {
            return Err(invalid(
                "cfRule has indexed and inline differential formats",
            ));
        }
        rule.differential_format = Some(DifferentialRef::StylesIndex(index));
    }
    rule.payload = payload.finish()?;
    validate_rule_shape(&rule)?;
    Ok(rule)
}

fn parse_rule_attributes(
    element: &BytesStart<'_>,
    decoder: Decoder,
    rule: &mut Rule,
) -> Result<()> {
    rule.rule_type = optional_attr(element, b"type", decoder)?
        .map(|value| {
            value
                .parse::<Kind>()
                .map_err(|error| invalid(error.to_string()))
        })
        .transpose()?;
    rule.priority = optional_attr(element, b"priority", decoder)?
        .map(|value| {
            value
                .parse::<i32>()
                .map_err(|_| invalid(format!("invalid conditional-formatting priority '{value}'")))
        })
        .transpose()?;
    if rule.priority.is_some_and(|value| value <= 0) {
        return Err(invalid("conditional-formatting priority must be positive"));
    }
    rule.stop_if_true = optional_bool(element, b"stopIfTrue", decoder)?.unwrap_or(false);
    rule.above_average = optional_bool(element, b"aboveAverage", decoder)?.unwrap_or(true);
    rule.equal_average = optional_bool(element, b"equalAverage", decoder)?.unwrap_or(false);
    rule.percent = optional_bool(element, b"percent", decoder)?.unwrap_or(false);
    rule.bottom = optional_bool(element, b"bottom", decoder)?.unwrap_or(false);
    rule.operator = optional_attr(element, b"operator", decoder)?
        .map(|value| {
            value
                .parse::<Operator>()
                .map_err(|error| invalid(error.to_string()))
        })
        .transpose()?;
    rule.text = optional_attr(element, b"text", decoder)?;
    rule.time_period = optional_attr(element, b"timePeriod", decoder)?
        .map(|value| {
            value
                .parse::<Period>()
                .map_err(|error| invalid(error.to_string()))
        })
        .transpose()?;
    rule.rank = optional_u32(element, b"rank", decoder)?;
    rule.standard_deviations = optional_attr(element, b"stdDev", decoder)?
        .map(|value| {
            value
                .parse::<i32>()
                .map_err(|_| invalid(format!("invalid stdDev '{value}'")))
        })
        .transpose()?;
    rule.extension_id = optional_attr(element, b"id", decoder)?;
    Ok(())
}

#[derive(Debug, Clone, Copy)]
enum TextTarget {
    Formula,
    ExtensionId,
    CfvoFormula,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PayloadKind {
    ColorScale,
    DataBar,
    IconSet,
}

struct PayloadBuilder {
    source: Source,
    active: Option<PayloadKind>,
    seen: Option<PayloadKind>,
    thresholds: Vec<Value>,
    colors: Vec<NamedColor>,
    data_attrs: Option<(u32, u32, bool, bool, bool, Direction, Axis)>,
    icon_attrs: Option<IconAttrs>,
}

enum IconAttrs {
    Core(IconSet, bool, bool, bool),
    Office2010(IconSet14, bool, bool, bool),
}

impl PayloadBuilder {
    fn new(source: Source) -> Self {
        Self {
            source,
            active: None,
            seen: None,
            thresholds: Vec::new(),
            colors: Vec::new(),
            data_attrs: None,
            icon_attrs: None,
        }
    }

    fn begin(
        &mut self,
        kind: PayloadKind,
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<()> {
        if self.active.is_some() || self.seen.is_some() {
            return Err(invalid("cfRule has multiple visual payloads"));
        }
        self.active = Some(kind);
        self.seen = Some(kind);
        if kind == PayloadKind::DataBar {
            let min = optional_u32(element, b"minLength", decoder)?.unwrap_or(10);
            let max = optional_u32(element, b"maxLength", decoder)?.unwrap_or(90);
            if min > max || max > 100 {
                return Err(invalid("invalid data-bar length bounds"));
            }
            let border = optional_bool(element, b"border", decoder)?;
            let gradient = optional_bool(element, b"gradient", decoder)?;
            let direction = optional_attr(element, b"direction", decoder)?;
            let axis = optional_attr(element, b"axisPosition", decoder)?;
            if self.source == Source::Core
                && (border.is_some() || gradient.is_some() || direction.is_some() || axis.is_some())
            {
                return Err(invalid("core dataBar contains Office 2010-only attributes"));
            }
            self.data_attrs = Some((
                min,
                max,
                optional_bool(element, b"showValue", decoder)?.unwrap_or(true),
                border.unwrap_or(false),
                gradient.unwrap_or(true),
                direction
                    .as_deref()
                    .unwrap_or("context")
                    .parse::<Direction>()
                    .map_err(|error| invalid(error.to_string()))?,
                axis.as_deref()
                    .unwrap_or("automatic")
                    .parse::<Axis>()
                    .map_err(|error| invalid(error.to_string()))?,
            ));
        } else if kind == PayloadKind::IconSet {
            let raw = optional_attr(element, b"iconSet", decoder)?
                .unwrap_or_else(|| "3TrafficLights1".into());
            let show = optional_bool(element, b"showValue", decoder)?.unwrap_or(true);
            let percent = optional_bool(element, b"percent", decoder)?.unwrap_or(true);
            let reverse = optional_bool(element, b"reverse", decoder)?.unwrap_or(false);
            self.icon_attrs = Some(match self.source {
                Source::Core => IconAttrs::Core(
                    raw.parse::<IconSet>()
                        .map_err(|error| invalid(error.to_string()))?,
                    show,
                    percent,
                    reverse,
                ),
                Source::Office2010 => IconAttrs::Office2010(
                    raw.parse::<IconSet14>()
                        .map_err(|error| invalid(error.to_string()))?,
                    show,
                    percent,
                    reverse,
                ),
            });
        }
        Ok(())
    }
    fn end(&mut self) -> Result<()> {
        if self.active.take().is_none() {
            return Err(invalid("unexpected visual payload end"));
        }
        Ok(())
    }
    fn cfvo(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        if self.active.is_none() {
            return Err(invalid("cfvo outside visual payload"));
        }
        let raw = required_attr(element, b"type", decoder)?;
        let kind = raw
            .parse::<ValueKind>()
            .map_err(|error| invalid(error.to_string()))?;
        if self.source == Source::Core && !kind.is_core() {
            return Err(invalid(format!(
                "core conditional formatting does not support CFVO kind '{raw}'"
            )));
        }
        self.thresholds.push(Value {
            kind,
            value: optional_attr(element, b"val", decoder)?,
            formula: None,
            greater_than_or_equal: optional_bool(element, b"gte", decoder)?.unwrap_or(true),
        });
        Ok(())
    }
    fn set_last_formula(&mut self, value: String) -> Result<()> {
        let last = self
            .thresholds
            .last_mut()
            .ok_or_else(|| invalid("CFVO formula without CFVO"))?;
        last.formula = Some(value);
        Ok(())
    }
    fn color(&mut self, role: &[u8], element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        if self.active.is_none() {
            return Err(invalid("color outside visual payload"));
        }
        let role = std::str::from_utf8(role)
            .map_err(xml_error)?
            .parse::<ColorRole>()
            .map_err(|error| invalid(error.to_string()))?;
        if self.source == Source::Core && role != ColorRole::Color {
            return Err(invalid(format!(
                "core conditional formatting does not support '{}' colors",
                role.as_str()
            )));
        }
        self.colors.push(NamedColor {
            role,
            color: Color {
                rgb: optional_attr(element, b"rgb", decoder)?
                    .map(|value| {
                        value
                            .parse::<Rgb>()
                            .map_err(|error| invalid(error.to_string()))
                    })
                    .transpose()?,
                indexed: optional_u32(element, b"indexed", decoder)?,
                theme: optional_u32(element, b"theme", decoder)?,
                tint: optional_attr(element, b"tint", decoder)?
                    .map(|v| v.parse::<f64>().map_err(|_| invalid("invalid color tint")))
                    .transpose()?,
                automatic: optional_bool(element, b"auto", decoder)?,
            },
        });
        Ok(())
    }
    fn finish(self) -> Result<Option<Payload>> {
        if self.active.is_some() {
            return Err(invalid("unterminated visual payload"));
        }
        Ok(match self.seen {
            None => None,
            Some(PayloadKind::ColorScale) => {
                if self.thresholds.len() < 2 || self.thresholds.len() != self.colors.len() {
                    return Err(invalid(
                        "color scale thresholds and colors must have matching cardinality of at least two",
                    ));
                }
                Some(Payload::ColorScale(ColorScale {
                    thresholds: self.thresholds,
                    colors: self.colors.into_iter().map(|value| value.color).collect(),
                }))
            },
            Some(PayloadKind::DataBar) => {
                if self.thresholds.len() != 2 {
                    return Err(invalid("data bar must contain exactly two thresholds"));
                }
                let (min, max, show, border, gradient, direction, axis) = self
                    .data_attrs
                    .ok_or_else(|| invalid("data-bar payload has no attributes"))?;
                Some(Payload::DataBar(DataBar {
                    thresholds: self.thresholds,
                    colors: self.colors,
                    min_length: min,
                    max_length: max,
                    show_value: show,
                    border,
                    gradient,
                    direction,
                    axis_position: axis,
                }))
            },
            Some(PayloadKind::IconSet) => match self
                .icon_attrs
                .ok_or_else(|| invalid("icon-set payload has no attributes"))?
            {
                IconAttrs::Core(set, show, percent, reverse) => {
                    if usize::from(set.len()) != self.thresholds.len() {
                        return Err(invalid(
                            "icon-set threshold count does not match its icon cardinality",
                        ));
                    }
                    Some(Payload::IconSet(Icons {
                        set,
                        thresholds: self.thresholds,
                        show_value: show,
                        percent,
                        reverse,
                    }))
                },
                IconAttrs::Office2010(set, show, percent, reverse) => {
                    if !(3..=5).contains(&self.thresholds.len())
                        || set
                            .len()
                            .is_some_and(|value| usize::from(value) != self.thresholds.len())
                    {
                        return Err(invalid(
                            "Office 2010 icon-set threshold count does not match its icon cardinality",
                        ));
                    }
                    Some(Payload::IconSet14(Icons {
                        set,
                        thresholds: self.thresholds,
                        show_value: show,
                        percent,
                        reverse,
                    }))
                },
            },
        })
    }
}

fn validate_rule_shape(rule: &Rule) -> Result<()> {
    match (&rule.rule_type, &rule.payload) {
        (Some(Kind::ColorScale), Some(Payload::ColorScale(_)))
        | (Some(Kind::DataBar), Some(Payload::DataBar(_)))
        | (Some(Kind::IconSet), Some(Payload::IconSet(_) | Payload::IconSet14(_))) => {},
        (Some(Kind::ColorScale | Kind::DataBar | Kind::IconSet), _) => {
            return Err(invalid(
                "visual cfRule type is missing its matching payload",
            ));
        },
        (_, Some(_)) => return Err(invalid("non-visual cfRule contains a visual payload")),
        _ => {},
    }
    if matches!(rule.payload, Some(Payload::IconSet14(_))) && rule.source != Source::Office2010 {
        return Err(invalid("Office 2010 icon set used in a core rule"));
    }
    if matches!(rule.payload, Some(Payload::IconSet(_))) && rule.source != Source::Core {
        return Err(invalid("core icon set used in an Office 2010 rule"));
    }
    if matches!(
        rule.operator,
        Some(Operator::Between | Operator::NotBetween)
    ) && rule.formulas.len() != 2
    {
        return Err(invalid("between cfRule requires two formulas"));
    }
    Ok(())
}

fn validate_and_associate(values: &mut [Formatting], dxf_count: usize) -> Result<()> {
    let mut priorities = HashSet::new();
    let mut ids = HashMap::new();
    for (ci, container) in values.iter().enumerate() {
        for (ri, rule) in container.rules.iter().enumerate() {
            if let Some(priority) = rule.priority
                && !priorities.insert(priority)
            {
                return Err(invalid(format!(
                    "duplicate conditional-formatting priority {priority}"
                )));
            }
            if let Some(DifferentialRef::StylesIndex(id)) = rule.differential_format
                && id as usize >= dxf_count
            {
                return Err(invalid(format!(
                    "conditional-formatting dxfId {id} is out of range"
                )));
            }
            if rule.source == Source::Core
                && let Some(id) = rule.extension_id.as_ref()
            {
                ids.insert(id.clone(), (ci, ri, rule.priority));
            }
        }
    }
    for container in values.iter_mut() {
        for rule in &mut container.rules {
            if rule.source == Source::Office2010 && rule.priority.is_none() {
                rule.extension_association =
                    match rule.extension_id.as_ref().and_then(|id| ids.get(id)) {
                        Some((_, _, Some(priority))) => Association::EnhancesCore {
                            priority: *priority,
                        },
                        _ => Association::UnmatchedIgnored,
                    };
            }
        }
    }
    Ok(())
}

fn parse_dxf(raw: &[u8]) -> Result<Differential> {
    let fragment = Captured {
        source: Source::Core,
        prefix: Vec::new(),
        bytes: raw.to_vec(),
    };
    let wrapped = wrap(&fragment.prefix, &fragment.bytes);
    let mut value = Differential {
        raw_xml: raw.to_vec().into_boxed_slice(),
        ..Default::default()
    };
    for name in [
        b"numFmt".as_slice(),
        b"font",
        b"fill",
        b"border",
        b"alignment",
        b"protection",
        b"extLst",
    ] {
        let fragments = capture_all(&wrapped, CORE, STRICT, name)?;
        if fragments.len() > 1 {
            return Err(invalid(format!(
                "dxf has duplicate {}",
                String::from_utf8_lossy(name)
            )));
        }
        let Some(component) = fragments.into_iter().next() else {
            continue;
        };
        if name == b"numFmt" {
            let component_xml = wrap(&component.prefix, &component.bytes);
            let mut reader = NsReader::from_reader(component_xml.as_slice());
            loop {
                let decoder = reader.decoder();
                let event = reader.read_event().map_err(xml_error)?.into_owned();
                if let Event::Start(ref e) | Event::Empty(ref e) = event
                    && e.local_name().as_ref() == b"numFmt"
                {
                    value.number_format = Some(NumberFormat {
                        id: required_attr(e, b"numFmtId", decoder)?
                            .parse()
                            .map_err(|_| invalid("invalid dxf numFmtId"))?,
                        code: required_attr(e, b"formatCode", decoder)?,
                        raw_xml: component.bytes.into_boxed_slice(),
                    });
                    break;
                }
                if matches!(event, Event::Eof) {
                    break;
                }
            }
        } else {
            let component = Component {
                raw_xml: component.bytes.into_boxed_slice(),
            };
            match name {
                b"font" => value.font = Some(component),
                b"fill" => value.fill = Some(component),
                b"border" => value.border = Some(component),
                b"alignment" => value.alignment = Some(component),
                b"protection" => value.protection = Some(component),
                _ => value.extensions.push(component),
            }
        }
    }
    Ok(value)
}

fn capture_first(xml: &[u8], ns1: &[u8], ns2: &[u8], name: &[u8]) -> Result<Option<Captured>> {
    Ok(capture_all(xml, ns1, ns2, name)?.into_iter().next())
}
fn capture_all(xml: &[u8], ns1: &[u8], ns2: &[u8], name: &[u8]) -> Result<Vec<Captured>> {
    let mut reader = NsReader::from_reader(xml);
    let mut out = Vec::new();
    let mut cap: Option<(usize, Vec<u8>, Writer<Vec<u8>>)> = None;
    loop {
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        if let Some((depth, _, writer)) = cap.as_mut() {
            writer.write_event(event.clone()).map_err(xml_error)?;
            match event {
                Event::Start(_) => *depth += 1,
                Event::End(_) => *depth -= 1,
                _ => {},
            }
            if *depth == 0 {
                let (_, prefix, writer) = cap
                    .take()
                    .ok_or_else(|| invalid("missing XML fragment capture state"))?;
                out.push(Captured {
                    source: Source::Core,
                    prefix,
                    bytes: writer.into_inner(),
                });
            }
            continue;
        }
        match event {
            Event::Start(e)
                if (exact(&namespace, ns1) || exact(&namespace, ns2))
                    && e.local_name().as_ref() == name =>
            {
                let p = prefix(e.name().as_ref());
                let mut w = Writer::new(Vec::new());
                w.write_event(Event::Start(e)).map_err(xml_error)?;
                cap = Some((1, p, w));
            },
            Event::Empty(e)
                if (exact(&namespace, ns1) || exact(&namespace, ns2))
                    && e.local_name().as_ref() == name =>
            {
                let p = prefix(e.name().as_ref());
                let mut w = Writer::new(Vec::new());
                w.write_event(Event::Empty(e)).map_err(xml_error)?;
                out.push(Captured {
                    source: Source::Core,
                    prefix: p,
                    bytes: w.into_inner(),
                });
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if cap.is_some() {
        return Err(invalid("unterminated XML fragment"));
    }
    Ok(out)
}

fn wrap(prefix: &[u8], fragment: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(fragment.len() + 512);
    out.extend_from_slice(br#"<root xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main""#);
    if !prefix.is_empty() && !matches!(prefix, b"s" | b"x" | b"x14" | b"xm") {
        out.extend_from_slice(b" xmlns:");
        out.extend_from_slice(prefix);
        out.extend_from_slice(b"=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
    }
    out.push(b'>');
    out.extend_from_slice(fragment);
    out.extend_from_slice(b"</root>");
    out
}
fn prefix(name: &[u8]) -> Vec<u8> {
    name.iter()
        .position(|b| *b == b':')
        .map_or_else(Vec::new, |i| name[..i].to_vec())
}
fn spreadsheet(ns: &ResolveResult<'_>) -> bool {
    exact(ns, CORE) || exact(ns, STRICT)
}
fn exact(ns: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(ns,ResolveResult::Bound(value)if value.as_ref()==expected)
}
fn is_color_element(name: &[u8]) -> bool {
    matches!(
        name,
        b"color"
            | b"fillColor"
            | b"borderColor"
            | b"negativeFillColor"
            | b"negativeBorderColor"
            | b"axisColor"
    )
}

fn parse_sqref(raw: &str) -> Result<Vec<Range>> {
    if raw.trim().is_empty() {
        return Err(invalid("conditionalFormatting sqref is empty"));
    }
    raw.split_whitespace()
        .map(|area| {
            let mut parts = area.split(':');
            let a = parts
                .next()
                .ok_or_else(|| invalid("conditional-formatting range has no first cell"))?;
            let b = parts.next();
            if parts.next().is_some() || !valid_cell(a) || b.is_some_and(|v| !valid_cell(v)) {
                return Err(invalid(format!(
                    "invalid conditional-formatting range '{area}'"
                )));
            }
            Ok(Range::from_raw(area.to_owned()))
        })
        .collect()
}
fn valid_cell(raw: &str) -> bool {
    let raw = raw.as_bytes();
    let mut i = 0;
    while i < raw.len() && raw[i] == b'$' {
        i += 1;
    }
    let start = i;
    while i < raw.len() && raw[i].is_ascii_alphabetic() {
        i += 1;
    }
    if i == start {
        return false;
    }
    let mut col = 0u32;
    for b in &raw[start..i] {
        col = col
            .saturating_mul(26)
            .saturating_add(u32::from(b.to_ascii_uppercase() - b'A' + 1));
    }
    if col == 0 || col > 16384 {
        return false;
    }
    if i < raw.len() && raw[i] == b'$' {
        i += 1;
    }
    let Ok(row) = std::str::from_utf8(&raw[i..])
        .ok()
        .unwrap_or("")
        .parse::<u32>()
    else {
        return false;
    };
    (1..=1_048_576).contains(&row)
}
fn optional_attr(e: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<Option<String>> {
    let mut value = None;
    for a in e.attributes().with_checks(true) {
        let a = a.map_err(xml_error)?;
        if a.key.local_name().as_ref() == name {
            if value.is_some() {
                return Err(invalid("duplicate XML attribute"));
            }
            value = Some(
                a.decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .map_err(xml_error)?
                    .into_owned(),
            );
        }
    }
    Ok(value)
}
fn required_attr(e: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<String> {
    optional_attr(e, name, decoder)?.ok_or_else(|| {
        invalid(format!(
            "missing required '{}' attribute",
            String::from_utf8_lossy(name)
        ))
    })
}
fn optional_u32(e: &BytesStart<'_>, name: &[u8], d: Decoder) -> Result<Option<u32>> {
    optional_attr(e, name, d)?
        .map(|v| {
            v.parse()
                .map_err(|_| invalid(format!("invalid unsigned integer '{v}'")))
        })
        .transpose()
}
fn optional_bool(e: &BytesStart<'_>, name: &[u8], d: Decoder) -> Result<Option<bool>> {
    optional_attr(e, name, d)?
        .map(|v| match v.as_str() {
            "1" | "true" => Ok(true),
            "0" | "false" => Ok(false),
            _ => Err(invalid(format!("invalid boolean '{v}'"))),
        })
        .transpose()
}
fn optional_u32_from_raw(raw: &[u8], name: &[u8]) -> Result<Option<u32>> {
    let wrapped = wrap(b"", raw);
    let mut r = NsReader::from_reader(wrapped.as_slice());
    loop {
        let d = r.decoder();
        match r.read_event().map_err(xml_error)? {
            Event::Start(e) if e.local_name().as_ref() == b"cfRule" => {
                return optional_u32(&e, name, d);
            },
            Event::Eof => return Ok(None),
            _ => {},
        }
    }
}
fn decode_text(text: &quick_xml::events::BytesText<'_>) -> Result<String> {
    decode_raw_text(text.as_ref())
}
fn decode_raw_text(bytes: &[u8]) -> Result<String> {
    let raw = std::str::from_utf8(bytes).map_err(xml_error)?;
    let Some(mut cursor) = raw.find('&') else {
        return Ok(raw.to_owned());
    };
    let mut value = String::with_capacity(raw.len());
    value.push_str(&raw[..cursor]);
    while cursor < raw.len() {
        let tail = &raw[cursor..];
        let end = tail
            .find(';')
            .ok_or_else(|| invalid("unterminated XML entity in formula"))?;
        let entity = &tail[1..end];
        match entity {
            "amp" => value.push('&'),
            "lt" => value.push('<'),
            "gt" => value.push('>'),
            "quot" => value.push('"'),
            "apos" => value.push('\''),
            _ if entity.starts_with("#x") => {
                let scalar = u32::from_str_radix(&entity[2..], 16)
                    .map_err(|_| invalid("invalid hexadecimal XML character reference"))?;
                value.push(
                    char::from_u32(scalar)
                        .ok_or_else(|| invalid("invalid XML character reference"))?,
                );
            },
            _ if entity.starts_with('#') => {
                let scalar = entity[1..]
                    .parse::<u32>()
                    .map_err(|_| invalid("invalid decimal XML character reference"))?;
                value.push(
                    char::from_u32(scalar)
                        .ok_or_else(|| invalid("invalid XML character reference"))?,
                );
            },
            _ => return Err(invalid(format!("unknown XML entity '&{entity};'"))),
        }
        let consumed = cursor + end + 1;
        match raw[consumed..].find('&') {
            Some(next) => {
                value.push_str(&raw[consumed..consumed + next]);
                cursor = consumed + next;
            },
            None => {
                value.push_str(&raw[consumed..]);
                break;
            },
        }
    }
    Ok(value)
}
fn xml_error(error: impl fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}
fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
