//! `SpreadsheetML` data-validation XML parsing.

use super::super::model::{
    Collection, Formula, ListSource, Source, Sqref, Validation, ValidationErrorStyle,
    ValidationImeMode, ValidationOperator, ValidationType,
};
use super::super::{
    MAX_DEPTH, MAX_EVENTS, MAX_FORMULA_BYTES, MAX_FRAGMENT_BYTES, MAX_NODES, MAX_VALIDATIONS,
    MAX_XML_BYTES, X12AC, X14, XM, XR,
};
use super::validation::{validate_data_validation_collections, validate_formula_cardinality};
use super::wire::{
    Captured, append_limited_text, bounded_attr, capture_collections, decode_flags, encode_flags,
    exact, invalid, optional_attr, optional_bool, optional_u32, parse_sqref, required_attr,
    reserve_vec, source_ns, spreadsheet, sqref_flags, uid_attr, wrap, xml_error,
};
use crate::error::Result;
use litchi_ooxml_common::mce::{Capabilities, Limits, Name, process_markup_compatibility};
use litchi_ooxml_common::xml::decode_xml_reference;
use quick_xml::Writer;
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

pub fn parse_data_validation_collections(xml: &[u8]) -> Result<Vec<Collection>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("data-validation worksheet XML is too large"));
    }
    let mut capabilities = Capabilities::default();
    capabilities
        .understand_namespace(String::from_utf8_lossy(X14).into_owned())
        .understand_namespace(String::from_utf8_lossy(XM).into_owned())
        .understand_namespace(String::from_utf8_lossy(X12AC).into_owned())
        .understand_namespace(String::from_utf8_lossy(XR).into_owned());
    capabilities.preserve_extension_element(Name {
        namespace: String::from_utf8_lossy(X14).into_owned(),
        local_name: "dataValidations".into(),
    });
    let limits = Limits {
        max_input_bytes: MAX_XML_BYTES,
        max_output_bytes: MAX_XML_BYTES,
        max_depth: MAX_DEPTH,
        ..Limits::default()
    };
    let validated = process_markup_compatibility(xml, &capabilities, &limits)?;
    if validated.xml.len() > MAX_XML_BYTES {
        return Err(invalid("processed data-validation XML is too large"));
    }
    let selected = if validated.report.alternate_content_count == 0 {
        xml
    } else {
        validated.xml.as_ref()
    };
    let fragments = capture_collections(selected)?;
    let mut values = Vec::new();
    reserve_vec(&mut values, fragments.len(), "data-validation collections")?;
    let mut count = 0usize;
    for fragment in fragments {
        let value = parse_collection(&fragment)?;
        count = count
            .checked_add(value.validations.len())
            .ok_or_else(|| invalid("data-validation count overflow"))?;
        if count > MAX_VALIDATIONS {
            return Err(invalid("too many data validations"));
        }
        values.push(value);
    }
    validate_data_validation_collections(&values)?;
    Ok(values)
}

/// In-flight capture state for a single `dataValidation` element.
fn parse_collection(fragment: &Captured) -> Result<Collection> {
    let wrapped = wrap(&fragment.prefix, &fragment.bytes)?;
    let mut reader = NsReader::from_reader(wrapped.as_slice());
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut root_depth = None;
    let mut closed = false;
    let mut expected = None;
    let mut disable = false;
    let mut x_window = None;
    let mut y_window = None;
    let mut validations = Vec::new();
    let mut capture: Option<(usize, Writer<Vec<u8>>)> = None;
    let mut events = 0usize;
    let mut nodes = 0usize;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("dataValidations event count overflow"))?;
        if events > MAX_EVENTS {
            return Err(invalid("dataValidations exceeds event limit"));
        }
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        if matches!(&event, Event::Eof) {
            if capture.is_some() || depth != 0 || !closed {
                return Err(invalid("unterminated dataValidations"));
            }
            break;
        }
        if matches!(&event, Event::Start(_) | Event::Empty(_) | Event::End(_)) {
            nodes = nodes
                .checked_add(1)
                .ok_or_else(|| invalid("dataValidations node count overflow"))?;
            if nodes > MAX_NODES {
                return Err(invalid("dataValidations exceeds node limit"));
            }
        }
        if let Some((capture_depth, writer)) = capture.as_mut() {
            writer.write_event(event.clone()).map_err(xml_error)?;
            if writer.get_ref().len() > MAX_FRAGMENT_BYTES {
                return Err(invalid("dataValidation rule is too large"));
            }
            match event {
                Event::Start(_) => {
                    if *capture_depth >= MAX_DEPTH {
                        return Err(invalid("dataValidation nesting is too deep"));
                    }
                    *capture_depth = capture_depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("dataValidation nesting overflow"))?;
                },
                Event::End(_) => {
                    *capture_depth = capture_depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("invalid dataValidation nesting"))?;
                },
                _ => {},
            }
            if *capture_depth == 0 {
                let Some((_, writer)) = capture.take() else {
                    return Err(invalid("dataValidation capture state disappeared"));
                };
                if validations.len() >= MAX_VALIDATIONS {
                    return Err(invalid("too many data validations"));
                }
                let raw = writer.into_inner();
                let value = parse_rule(&raw, fragment.source)?;
                reserve_vec(&mut validations, 1, "data-validation rules")?;
                validations.push(value);
            }
            continue;
        }
        match event {
            Event::Start(element)
                if element.local_name().as_ref() == b"dataValidations"
                    && source_ns(fragment.source, &namespace) =>
            {
                if root_depth.is_some() {
                    return Err(invalid("nested dataValidations"));
                }
                if closed || depth >= MAX_DEPTH {
                    return Err(invalid("dataValidations nesting is too deep"));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("dataValidations nesting overflow"))?;
                root_depth = Some(depth);
                expected = optional_u32(&element, b"count", decoder)?;
                if expected.is_some_and(|value| value as usize > MAX_VALIDATIONS) {
                    return Err(invalid("too many data validations"));
                }
                disable = optional_bool(&element, b"disablePrompts", decoder)?.unwrap_or(false);
                x_window = optional_u32(&element, b"xWindow", decoder)?;
                y_window = optional_u32(&element, b"yWindow", decoder)?;
            },
            Event::Start(element)
                if element.local_name().as_ref() == b"dataValidation"
                    && source_ns(fragment.source, &namespace)
                    && root_depth == Some(depth) =>
            {
                if closed {
                    return Err(invalid("content follows dataValidations"));
                }
                let mut writer = Writer::new(Vec::new());
                writer
                    .write_event(Event::Start(element))
                    .map_err(xml_error)?;
                if writer.get_ref().len() > MAX_FRAGMENT_BYTES {
                    return Err(invalid("dataValidation rule is too large"));
                }
                capture = Some((1, writer));
            },
            Event::Empty(element)
                if element.local_name().as_ref() == b"dataValidation"
                    && source_ns(fragment.source, &namespace)
                    && root_depth == Some(depth) =>
            {
                if closed || validations.len() >= MAX_VALIDATIONS {
                    return Err(invalid("too many data validations"));
                }
                let mut writer = Writer::new(Vec::new());
                writer
                    .write_event(Event::Empty(element))
                    .map_err(xml_error)?;
                if writer.get_ref().len() > MAX_FRAGMENT_BYTES {
                    return Err(invalid("dataValidation rule is too large"));
                }
                let raw = writer.into_inner();
                let value = parse_rule(&raw, fragment.source)?;
                reserve_vec(&mut validations, 1, "data-validation rules")?;
                validations.push(value);
            },
            Event::Start(_) => {
                if depth >= MAX_DEPTH {
                    return Err(invalid("dataValidations nesting is too deep"));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("dataValidations nesting overflow"))?;
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected dataValidations closing element"));
                }
                if root_depth == Some(depth)
                    && element.local_name().as_ref() == b"dataValidations"
                    && source_ns(fragment.source, &namespace)
                {
                    closed = true;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("invalid dataValidations nesting"))?;
            },
            Event::Text(value) => {
                if !value.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("dataValidations must not contain text"));
                }
            },
            Event::CData(_) | Event::GeneralRef(_) => {
                return Err(invalid("dataValidations must not contain character data"));
            },
            Event::Decl(_) => {
                return Err(invalid(
                    "XML declarations are not allowed in dataValidations",
                ));
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Empty(_) => {},
            Event::Comment(_) => {},
            Event::Eof => return Err(invalid("unexpected EOF in dataValidations")),
        }
    }
    if validations.is_empty() {
        return Err(invalid("dataValidations must contain at least one rule"));
    }
    if expected.is_some_and(|v| v as usize != validations.len()) {
        return Err(invalid("dataValidations count does not match its children"));
    }
    Ok(Collection {
        source: fragment.source,
        disable_prompts: disable,
        x_window,
        y_window,
        declared_count: expected,
        validations,
    })
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum TextTarget {
    Formula1,
    Formula2,
    Sqref,
    List,
}

fn text_target_matches(
    target: TextTarget,
    source: Source,
    namespace: &ResolveResult<'_>,
    local: &[u8],
) -> bool {
    match target {
        TextTarget::Formula1 => {
            (source == Source::Core && local == b"formula1" && source_ns(source, namespace))
                || (source == Source::Office2010 && local == b"f" && exact(namespace, XM))
        },
        TextTarget::Formula2 => {
            (source == Source::Core && local == b"formula2" && source_ns(source, namespace))
                || (source == Source::Office2010 && local == b"f" && exact(namespace, XM))
        },
        TextTarget::Sqref => local == b"sqref" && exact(namespace, XM),
        TextTarget::List => local == b"list" && exact(namespace, X12AC),
    }
}

fn parse_rule(raw: &[u8], source: Source) -> Result<Validation> {
    if raw.len() > MAX_FRAGMENT_BYTES {
        return Err(invalid("dataValidation rule is too large"));
    }
    let wrapped = wrap(if source == Source::Core { b"" } else { b"x14" }, raw)?;
    let mut reader = NsReader::from_reader(wrapped.as_slice());
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut rule_depth = None;
    let mut closed = false;
    let mut order = 0u8;
    let mut kind = ValidationType::None;
    let mut operator = ValidationOperator::Between;
    let mut error_style = ValidationErrorStyle::Stop;
    let mut ime = ValidationImeMode::NoControl;
    let mut allow_blank = false;
    let mut show_drop_down = false;
    let mut show_input = false;
    let mut show_error = false;
    let (mut error_title, mut error, mut prompt_title, mut prompt, mut uid) =
        (None, None, None, None, None);
    let (mut formula1, mut formula2, mut sqref): (
        Option<ListSource>,
        Option<Formula>,
        Option<Sqref>,
    ) = (None, None, None);
    let mut wrapper: Option<(u8, usize, bool)> = None;
    let mut text: Option<(usize, TextTarget, String)> = None;
    let mut events = 0usize;
    let mut nodes = 0usize;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("dataValidation event count overflow"))?;
        if events > MAX_EVENTS {
            return Err(invalid("dataValidation exceeds event limit"));
        }
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        if matches!(&event, Event::Eof) {
            break;
        }
        if matches!(&event, Event::Start(_) | Event::Empty(_) | Event::End(_)) {
            nodes = nodes
                .checked_add(1)
                .ok_or_else(|| invalid("dataValidation node count overflow"))?;
            if nodes > MAX_NODES {
                return Err(invalid("dataValidation exceeds node limit"));
            }
        }
        match event {
            Event::Start(element) => {
                if closed {
                    return Err(invalid("content follows dataValidation"));
                }
                let local = element.local_name();
                if local.as_ref() == b"dataValidation"
                    && source_ns(source, &namespace)
                    && rule_depth.is_none()
                {
                    if depth >= MAX_DEPTH {
                        return Err(invalid("dataValidation nesting is too deep"));
                    }
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("dataValidation nesting overflow"))?;
                    rule_depth = Some(depth);
                    kind = ValidationType::parse(
                        optional_attr(&element, b"type", decoder)?
                            .as_deref()
                            .unwrap_or("none"),
                    )?;
                    operator = ValidationOperator::parse(
                        optional_attr(&element, b"operator", decoder)?
                            .as_deref()
                            .unwrap_or("between"),
                    )?;
                    error_style = ValidationErrorStyle::parse(
                        optional_attr(&element, b"errorStyle", decoder)?
                            .as_deref()
                            .unwrap_or("stop"),
                    )?;
                    ime = ValidationImeMode::parse(
                        optional_attr(&element, b"imeMode", decoder)?
                            .as_deref()
                            .unwrap_or("noControl"),
                    )?;
                    allow_blank = optional_bool(&element, b"allowBlank", decoder)?.unwrap_or(false);
                    show_drop_down =
                        optional_bool(&element, b"showDropDown", decoder)?.unwrap_or(false);
                    show_input =
                        optional_bool(&element, b"showInputMessage", decoder)?.unwrap_or(false);
                    show_error =
                        optional_bool(&element, b"showErrorMessage", decoder)?.unwrap_or(false);
                    error_title = bounded_attr(&element, b"errorTitle", decoder, 32)?;
                    error = bounded_attr(&element, b"error", decoder, 225)?;
                    prompt_title = bounded_attr(&element, b"promptTitle", decoder, 32)?;
                    prompt = bounded_attr(&element, b"prompt", decoder, 255)?;
                    uid = uid_attr(&element, decoder, &resolver)?;
                    if source == Source::Core {
                        sqref = Some(parse_sqref(
                            &required_attr(&element, b"sqref", decoder)?,
                            false,
                            false,
                            false,
                            false,
                        )?);
                    }
                } else if rule_depth == Some(depth)
                    && source_ns(source, &namespace)
                    && matches!(local.as_ref(), b"formula1" | b"formula2")
                {
                    let number = if local.as_ref() == b"formula1" { 1 } else { 2 };
                    if number < order {
                        return Err(invalid("dataValidation children are out of order"));
                    }
                    order = number;
                    if wrapper.is_some() {
                        return Err(invalid("nested data-validation formula wrapper"));
                    }
                    let target_depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("dataValidation nesting overflow"))?;
                    wrapper = Some((number, target_depth, false));
                    if target_depth > MAX_DEPTH {
                        return Err(invalid("dataValidation nesting is too deep"));
                    }
                    depth = target_depth;
                    if source == Source::Core {
                        text = Some((
                            depth,
                            if number == 1 {
                                TextTarget::Formula1
                            } else {
                                TextTarget::Formula2
                            },
                            String::new(),
                        ));
                    }
                } else if source == Source::Office2010
                    && wrapper.is_some()
                    && rule_depth.is_some_and(|value| depth == value + 1)
                    && exact(&namespace, XM)
                    && local.as_ref() == b"f"
                {
                    let Some(wrapper_state) = wrapper.as_mut() else {
                        return Err(invalid("data-validation formula outside its wrapper"));
                    };
                    if wrapper_state.2 {
                        return Err(invalid("formula wrapper must contain exactly one value"));
                    }
                    wrapper_state.2 = true;
                    depth += 1;
                    let target = if wrapper_state.0 == 1 {
                        TextTarget::Formula1
                    } else {
                        TextTarget::Formula2
                    };
                    text = Some((depth, target, String::new()));
                } else if source == Source::Office2010
                    && wrapper.is_some()
                    && rule_depth.is_some_and(|value| depth == value + 1)
                    && exact(&namespace, X12AC)
                    && local.as_ref() == b"list"
                {
                    let Some(wrapper_state) = wrapper.as_mut() else {
                        return Err(invalid("quoted validation list is outside its wrapper"));
                    };
                    if wrapper_state.0 != 1 {
                        return Err(invalid("quoted validation list is only valid in formula1"));
                    }
                    if wrapper_state.2 {
                        return Err(invalid("formula wrapper must contain exactly one value"));
                    }
                    wrapper_state.2 = true;
                    depth += 1;
                    text = Some((depth, TextTarget::List, String::new()));
                } else if source == Source::Office2010
                    && rule_depth == Some(depth)
                    && exact(&namespace, XM)
                    && local.as_ref() == b"sqref"
                {
                    if order > 3 {
                        return Err(invalid("dataValidation children are out of order"));
                    }
                    order = 3;
                    let flags = sqref_flags(&element, decoder)?;
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("dataValidation nesting overflow"))?;
                    if depth > MAX_DEPTH {
                        return Err(invalid("dataValidation nesting is too deep"));
                    }
                    text = Some((depth, TextTarget::Sqref, encode_flags(flags)));
                } else {
                    if depth >= MAX_DEPTH {
                        return Err(invalid("dataValidation nesting is too deep"));
                    }
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("dataValidation nesting overflow"))?;
                }
            },
            Event::Text(value) => {
                if let Some((_, _, buffer)) = text.as_mut() {
                    let decoded = value.decode().map_err(xml_error)?;
                    append_limited_text(
                        buffer,
                        &decoded,
                        MAX_FORMULA_BYTES,
                        "data-validation text",
                    )?;
                } else if !value.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("dataValidation contains unexpected text"));
                }
            },
            Event::CData(value) => {
                if let Some((_, _, buffer)) = text.as_mut() {
                    let decoded = value.decode().map_err(xml_error)?;
                    append_limited_text(
                        buffer,
                        &decoded,
                        MAX_FORMULA_BYTES,
                        "data-validation text",
                    )?;
                } else {
                    return Err(invalid("dataValidation contains unexpected CDATA"));
                }
            },
            Event::GeneralRef(value) => {
                if let Some((_, _, buffer)) = text.as_mut() {
                    let decoded = decode_xml_reference(&value)?;
                    append_limited_text(
                        buffer,
                        &decoded,
                        MAX_FORMULA_BYTES,
                        "data-validation text",
                    )?;
                } else {
                    return Err(invalid("dataValidation contains unexpected entity text"));
                }
            },
            Event::Empty(element) => {
                if closed {
                    return Err(invalid("content follows dataValidation"));
                }
                let local = element.local_name();
                if source == Source::Core
                    && rule_depth == Some(depth)
                    && spreadsheet(&namespace)
                    && matches!(local.as_ref(), b"formula1" | b"formula2")
                {
                    let number = if local.as_ref() == b"formula1" { 1 } else { 2 };
                    if number < order {
                        return Err(invalid("dataValidation children are out of order"));
                    }
                    order = number;
                    if number == 1 {
                        if formula1
                            .replace(ListSource::Formula(Formula(String::new())))
                            .is_some()
                        {
                            return Err(invalid("duplicate formula1"));
                        }
                    } else if formula2.replace(Formula(String::new())).is_some() {
                        return Err(invalid("duplicate formula2"));
                    }
                } else if source == Source::Office2010
                    && wrapper.is_some()
                    && rule_depth.is_some_and(|value| depth == value + 1)
                    && exact(&namespace, XM)
                    && local.as_ref() == b"f"
                {
                    let Some(wrapper_state) = wrapper.as_mut() else {
                        return Err(invalid("formula value is outside its wrapper"));
                    };
                    if wrapper_state.2 {
                        return Err(invalid("formula wrapper must contain exactly one value"));
                    }
                    let number = wrapper_state.0;
                    wrapper_state.2 = true;
                    if number == 1 {
                        if formula1
                            .replace(ListSource::Formula(Formula(String::new())))
                            .is_some()
                        {
                            return Err(invalid("duplicate formula1"));
                        }
                    } else if formula2.replace(Formula(String::new())).is_some() {
                        return Err(invalid("duplicate formula2"));
                    }
                } else if source == Source::Office2010
                    && wrapper.is_some()
                    && rule_depth.is_some_and(|value| depth == value + 1)
                    && exact(&namespace, X12AC)
                    && local.as_ref() == b"list"
                {
                    let Some(wrapper_state) = wrapper.as_mut() else {
                        return Err(invalid("data-validation list source outside its wrapper"));
                    };
                    if wrapper_state.0 != 1 || wrapper_state.2 {
                        return Err(invalid("invalid quoted-list formula wrapper"));
                    }
                    wrapper_state.2 = true;
                    if formula1
                        .replace(ListSource::QuotedList(String::new()))
                        .is_some()
                    {
                        return Err(invalid("duplicate formula1 source"));
                    }
                } else if source == Source::Office2010
                    && rule_depth == Some(depth)
                    && exact(&namespace, XM)
                    && local.as_ref() == b"sqref"
                {
                    return Err(invalid("data-validation sqref is empty"));
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected dataValidation closing element"));
                }
                if text.as_ref().is_some_and(|(target, _, _)| *target == depth) {
                    let Some((_, target, value)) = text.take() else {
                        return Err(invalid("data-validation text state disappeared"));
                    };
                    if !text_target_matches(
                        target,
                        source,
                        &namespace,
                        element.local_name().as_ref(),
                    ) {
                        return Err(invalid("invalid data-validation text closing element"));
                    }
                    match target {
                        TextTarget::Formula1 => {
                            if formula1.is_some() {
                                return Err(invalid("duplicate formula1"));
                            }
                            formula1 = Some(ListSource::Formula(Formula(value)));
                        },
                        TextTarget::Formula2 => {
                            if formula2.is_some() {
                                return Err(invalid("duplicate formula2"));
                            }
                            formula2 = Some(Formula(value));
                        },
                        TextTarget::List => {
                            if formula1.is_some() {
                                return Err(invalid("duplicate formula1 source"));
                            }
                            formula1 = Some(ListSource::QuotedList(value));
                        },
                        TextTarget::Sqref => {
                            let (flags, value) = decode_flags(&value)?;
                            sqref = Some(parse_sqref(&value, flags.0, flags.1, flags.2, flags.3)?);
                        },
                    }
                }
                if wrapper
                    .as_ref()
                    .is_some_and(|(_, target, _)| *target == depth)
                {
                    let Some((number, _, seen)) = wrapper.take() else {
                        return Err(invalid("data-validation wrapper state disappeared"));
                    };
                    let expected = if number == 1 {
                        b"formula1"
                    } else {
                        b"formula2"
                    };
                    if !source_ns(source, &namespace) || element.local_name().as_ref() != expected {
                        return Err(invalid("invalid data-validation wrapper closing element"));
                    }
                    if source == Source::Office2010 && !seen {
                        return Err(invalid(
                            "x14 formula wrapper must contain exactly one value",
                        ));
                    }
                }
                if rule_depth == Some(depth)
                    && element.local_name().as_ref() == b"dataValidation"
                    && source_ns(source, &namespace)
                {
                    closed = true;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("invalid dataValidation nesting"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Decl(_) => {
                return Err(invalid(
                    "XML declarations are not allowed in dataValidation",
                ));
            },
            Event::Comment(_) => {},
            Event::Eof => return Err(invalid("unexpected EOF in dataValidation")),
        }
    }
    if !closed || wrapper.is_some() || text.is_some() || depth != 0 {
        return Err(invalid("unterminated dataValidation"));
    }
    let sqref = sqref.ok_or_else(|| invalid("dataValidation is missing sqref"))?;
    validate_formula_cardinality(kind, operator, &formula1, &formula2)?;
    Ok(Validation {
        source,
        validation_type: kind,
        operator,
        error_style,
        ime_mode: ime,
        allow_blank,
        show_drop_down,
        show_input_message: show_input,
        show_error_message: show_error,
        error_title,
        error,
        prompt_title,
        prompt,
        formula1,
        formula2,
        sqref,
        uid,
    })
}
