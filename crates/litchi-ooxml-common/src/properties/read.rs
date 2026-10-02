//! OPC core-properties metadata extraction: tolerant reads (LENIENCY_AUDIT
//! 8.1), strict decoding for the writer.

use super::{
    CORE_PROPERTIES_NAMESPACE, Dialect, Keywords, MAX_PROPERTY_TEXT, MAX_XML_BYTES, MAX_XML_EVENTS,
    Props, STRICT_CORE_PROPERTIES_NAMESPACE, graph,
    keyword::{self, Item},
    time,
};
use crate::leniency::{CorePropertiesDefect, ToleranceReport};
use crate::{Error, Result};
use litchi_opc::{OpcPackage, SourceBackedPackage};

use crate::xml::attributes::BytesStartExt as _;
use crate::xml::decode_xml_reference;
use quick_xml::{
    XmlVersion,
    events::{BytesEnd, BytesStart, Event},
    name::{Namespace, QName, ResolveResult},
    reader::NsReader,
};

const DUBLIN_CORE_NAMESPACE: &[u8] = b"http://purl.org/dc/elements/1.1/";
const DUBLIN_CORE_TERMS_NAMESPACE: &[u8] = b"http://purl.org/dc/terms/";
const XSI_NAMESPACE: &[u8] = b"http://www.w3.org/2001/XMLSchema-instance";
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const MARKUP_COMPATIBILITY_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/markup-compatibility/2006";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum CoreProperty {
    Title,
    Subject,
    Author,
    Keywords,
    Description,
    Identifier,
    Language,
    LastModifiedBy,
    Revision,
    Category,
    ContentStatus,
    Version,
    Created,
    Modified,
    LastPrinted,
}

#[derive(Debug)]
enum Active {
    Text(CoreProperty, String),
    Keywords {
        value: Keywords,
        child: Option<keyword::Value>,
        bytes: usize,
    },
}

impl CoreProperty {
    const COUNT: usize = Self::LastPrinted as usize + 1;

    fn index(self) -> usize {
        self as usize
    }

    fn requires_w3cdtf(self) -> bool {
        matches!(self, Self::Created | Self::Modified)
    }
}

/// Extract relationship-selected core metadata from an OOXML package,
/// appending every reference-compatible repair to `report`.
///
/// Absence is valid and remains distinguishable from a present empty document.
/// Producer-rule deviations in the part body (OPC M4.3-M4.5, duplicates,
/// unknown elements, invalid dates, non-UTF-8 bytes, a dialect mismatch) and
/// in the relationship graph are repaired per property as
/// [`crate::leniency`] documents, and each repair is recorded in `report`: a
/// conformant package appends nothing. There is deliberately no report-less
/// variant — repaired properties cannot be obtained without their report
/// (the tolerance contract). Ill-formed XML, the Markup Compatibility
/// namespace (M4.2), and limits still fail.
///
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn read_with_report(
    package: &OpcPackage,
    report: &mut ToleranceReport,
) -> Result<Option<Props>> {
    let graph = graph::inspect_tolerant(package, report)?;
    let Some(part_name) = graph.part else {
        return Ok(None);
    };
    let core_part = package.get_part(&part_name)?;
    decode_selected(core_part.blob(), part_name.as_str(), graph.dialect, report).map(Some)
}

/// Decode the relationship-selected core-properties payload, repairing and
/// recording a non-UTF-8 payload and a dialect mismatch (LENIENCY_AUDIT 8.1).
fn decode_selected(
    bytes: &[u8],
    part: &str,
    relationship_dialect: Option<Dialect>,
    report: &mut ToleranceReport,
) -> Result<Props> {
    let mut recorder = Recorder {
        report,
        part,
        strict: false,
    };
    let xml = match std::str::from_utf8(bytes) {
        Ok(xml) => std::borrow::Cow::Borrowed(xml),
        Err(_) => {
            recorder.record(CorePropertiesDefect::NonUtf8, None, None)?;
            String::from_utf8_lossy(bytes)
        },
    };
    let (props, dialect) = decode_reporting(&xml, &mut recorder)?;
    if let Some(relationship_dialect) = relationship_dialect
        && relationship_dialect != dialect
    {
        recorder.record(
            CorePropertiesDefect::DialectMismatch,
            None,
            Some(match relationship_dialect {
                Dialect::Strict => super::STRICT_CORE_PROPERTIES_RELATIONSHIP,
                Dialect::Transitional => litchi_opc::constants::relationship_type::CORE_PROPERTIES,
            }),
        )?;
    }
    Ok(props)
}

/// Where core-properties repairs are recorded, and for which part.
///
/// A strict recorder refuses instead of repairing: the writer uses it so
/// content the typed model cannot carry is never silently overwritten.
struct Recorder<'a> {
    report: &'a mut ToleranceReport,
    part: &'a str,
    strict: bool,
}

impl Recorder<'_> {
    fn record(
        &mut self,
        defect: CorePropertiesDefect,
        item: Option<&str>,
        value: Option<&str>,
    ) -> Result<()> {
        if self.strict {
            return Err(Error::Invalid(format!(
                "non-conformant core properties ({}): {}",
                item.unwrap_or("part"),
                defect.repair()
            )));
        }
        self.report.record(defect, self.part, item, value);
        Ok(())
    }
}

/// Extract relationship-selected core metadata from a source-backed OPC
/// package, appending every reference-compatible repair to `report`.
///
/// The catalog and relationship graph remain metadata-only until the
/// relationship-selected core-properties part is read. This is the
/// positional equivalent of [`read_with_report`]: it applies the same graph
/// and dialect repairs (recording them), while retaining the source
/// package's version and execution-policy checks around the selected payload
/// read.
///
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a
/// configured bound, the source changes, execution is cancelled, or an
/// underlying source/package operation fails.
pub fn read_source_backed_with_report(
    package: &SourceBackedPackage,
    report: &mut ToleranceReport,
) -> Result<Option<Props>> {
    package.check_execution()?;
    package.source_version()?;
    let graph = graph::inspect_source(package, report)?;
    let Some(part_name) = graph.part else {
        package.check_execution()?;
        package.source_version()?;
        return Ok(None);
    };
    let core_part = package.part(&part_name)?;
    let data = core_part.data()?;
    package.check_execution()?;
    let props = decode_selected(data.as_bytes(), part_name.as_str(), graph.dialect, report)?;
    package.check_execution()?;
    package.source_version()?;
    Ok(Some(props))
}

/// Decode core-properties XML strictly: every deviation [`read_with_report`]
/// repairs is
/// an error here.
///
/// The writer uses this before replacing a payload, so content the typed
/// model cannot carry (an unknown element, a duplicate) is never silently
/// discarded by a write.
pub(super) fn decode(xml: &str) -> Result<(Props, Dialect)> {
    let mut report = ToleranceReport::default();
    decode_reporting(
        xml,
        &mut Recorder {
            report: &mut report,
            part: "",
            strict: true,
        },
    )
}

/// Decode core-properties XML, repairing the producer-rule deviations of
/// LENIENCY_AUDIT 8.1 and recording each into `recorder`.
///
/// Word and LibreOffice look properties up one by one and ignore what they do
/// not understand, so a deviation costs at most the deviating property:
/// unknown elements and duplicate properties are skipped, unexpected
/// attributes ignored, a missing W3CDTF `xsi:type` accepted, and an
/// unparseable date dropped. XML well-formedness, the document element, DTDs,
/// the Markup Compatibility namespace (OPC M4.2), keyword structure, and the
/// byte/event/text limits stay hard errors.
fn decode_reporting(xml: &str, recorder: &mut Recorder<'_>) -> Result<(Props, Dialect)> {
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::Limit {
            resource: "core-properties XML bytes",
            max: MAX_XML_BYTES,
            actual: xml.len(),
        });
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut buffer = Vec::new();
    let mut props = Props::new();
    let mut seen = [false; CoreProperty::COUNT];
    let mut root_namespace: Option<Vec<u8>> = None;
    let mut root_closed = false;
    let mut active: Option<Active> = None;
    let mut events = 0usize;
    // Depth of a skipped unknown or duplicate property subtree; 0 when not
    // skipping. Bounded by the event limit.
    let mut skip_depth = 0usize;

    loop {
        events = events.checked_add(1).ok_or(Error::Limit {
            resource: "core-properties XML events",
            max: MAX_XML_EVENTS,
            actual: usize::MAX,
        })?;
        if events > MAX_XML_EVENTS {
            return Err(Error::Limit {
                resource: "core-properties XML events",
                max: MAX_XML_EVENTS,
                actual: events,
            });
        }
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Xml(format!("core properties XML error: {error}")))?;
        let namespace = bound_namespace(&namespace).map(<[u8]>::to_vec);
        let namespace = namespace.as_deref();
        let event = event.into_owned();
        reject_markup_compatibility(namespace)?;

        if skip_depth > 0 {
            match event {
                Event::Start(_) => skip_depth = skip_depth.saturating_add(1),
                Event::End(_) => skip_depth -= 1,
                Event::DocType(_) => {
                    return Err(Error::Invalid(
                        "DTD is not allowed in core properties".to_string(),
                    ));
                },
                Event::Eof => break,
                _ => {},
            }
            buffer.clear();
            continue;
        }

        match event {
            Event::Start(element) if root_namespace.is_none() => {
                validate_root(namespace, &element, &reader, recorder)?;
                root_namespace = Some(
                    namespace
                        .ok_or_else(|| Error::Invalid("coreProperties has no namespace".into()))?
                        .to_vec(),
                );
            },
            Event::Empty(element) if root_namespace.is_none() => {
                validate_root(namespace, &element, &reader, recorder)?;
                root_namespace = Some(
                    namespace
                        .ok_or_else(|| Error::Invalid("coreProperties has no namespace".into()))?
                        .to_vec(),
                );
                root_closed = true;
            },
            Event::Start(element) if !root_closed && active.is_none() => {
                let Some(property) = admit_property(namespace, &element, &mut seen, recorder)?
                else {
                    skip_depth = 1;
                    buffer.clear();
                    continue;
                };
                active = Some(if property == CoreProperty::Keywords {
                    Active::Keywords {
                        value: Keywords {
                            lang: keyword_lang(&reader, &element, recorder)?,
                            items: Vec::new(),
                        },
                        child: None,
                        bytes: 0,
                    }
                } else {
                    validate_property_attributes(&reader, &element, property, recorder)?;
                    Active::Text(property, String::new())
                });
            },
            Event::Empty(element) if !root_closed && active.is_none() => {
                let Some(property) = admit_property(namespace, &element, &mut seen, recorder)?
                else {
                    buffer.clear();
                    continue;
                };
                if property == CoreProperty::Keywords {
                    props.keywords = Some(Keywords {
                        lang: keyword_lang(&reader, &element, recorder)?,
                        items: Vec::new(),
                    });
                } else {
                    validate_property_attributes(&reader, &element, property, recorder)?;
                    apply_property(&mut props, property, String::new(), recorder)?;
                }
            },
            Event::Start(element) if active.is_some() => {
                start_keyword_value(
                    &mut active,
                    namespace,
                    root_namespace.as_deref(),
                    &reader,
                    &element,
                    false,
                    recorder,
                )?;
            },
            Event::Empty(element) if active.is_some() => {
                start_keyword_value(
                    &mut active,
                    namespace,
                    root_namespace.as_deref(),
                    &reader,
                    &element,
                    true,
                    recorder,
                )?;
            },
            Event::Start(_) | Event::Empty(_) => {
                return Err(Error::Invalid(
                    "coreProperties must be the only document element".to_string(),
                ));
            },
            Event::Text(value) => {
                let decoded = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                let decoded = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                push_property_text(&mut active, &decoded)?;
            },
            Event::CData(value) => {
                let decoded = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                push_property_text(&mut active, &decoded)?;
            },
            Event::GeneralRef(reference) => {
                push_property_text(&mut active, &decode_xml_reference(&reference)?)?;
            },
            Event::End(element) if active.is_some() => {
                close_active(
                    &mut active,
                    namespace,
                    root_namespace.as_deref(),
                    &element,
                    &mut props,
                    recorder,
                )?;
            },
            Event::End(element) if !root_closed => {
                if element.local_name().as_ref() != b"coreProperties"
                    || namespace != root_namespace.as_deref()
                {
                    return Err(Error::Invalid(
                        "malformed coreProperties root element".to_string(),
                    ));
                }
                root_closed = true;
            },
            Event::End(_) => {
                return Err(Error::Invalid(
                    "unexpected element after coreProperties".to_string(),
                ));
            },
            Event::DocType(_) => {
                return Err(Error::Invalid(
                    "DTD is not allowed in core properties".to_string(),
                ));
            },
            Event::Eof => break,
            Event::Comment(_) | Event::Decl(_) | Event::PI(_) => {},
        }
        buffer.clear();
    }
    if root_namespace.is_none() || !root_closed || active.is_some() {
        return Err(Error::Invalid(
            "unterminated or missing coreProperties root element".to_string(),
        ));
    }
    let dialect = match root_namespace.as_deref() {
        Some(CORE_PROPERTIES_NAMESPACE) => Dialect::Transitional,
        Some(STRICT_CORE_PROPERTIES_NAMESPACE) => Dialect::Strict,
        _ => {
            return Err(Error::Invalid(
                "unsupported core-properties namespace".to_owned(),
            ));
        },
    };
    Ok((props, dialect))
}

fn validate_root(
    namespace: Option<&[u8]>,
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    recorder: &mut Recorder<'_>,
) -> Result<()> {
    if element.local_name().as_ref() != b"coreProperties"
        || !matches!(
            namespace,
            Some(CORE_PROPERTIES_NAMESPACE | STRICT_CORE_PROPERTIES_NAMESPACE)
        )
    {
        return Err(Error::Invalid(
            "expected coreProperties in the OPC core-properties namespace".to_string(),
        ));
    }
    validate_attributes(reader, element, None, recorder)
}

/// Admit a property start tag: `None` (recorded) when the element is not a
/// known core property or repeats one, so the caller skips its subtree.
fn admit_property(
    namespace: Option<&[u8]>,
    element: &BytesStart<'_>,
    seen: &mut [bool; CoreProperty::COUNT],
    recorder: &mut Recorder<'_>,
) -> Result<Option<CoreProperty>> {
    let Some(property) = parse_property(namespace, element.local_name().as_ref()) else {
        recorder.record(
            CorePropertiesDefect::UnknownElement,
            Some(&String::from_utf8_lossy(element.name().as_ref())),
            None,
        )?;
        return Ok(None);
    };
    if std::mem::replace(&mut seen[property.index()], true) {
        recorder.record(
            CorePropertiesDefect::DuplicateProperty,
            Some(property_str(property)),
            None,
        )?;
        return Ok(None);
    }
    Ok(Some(property))
}

fn parse_property(namespace: Option<&[u8]>, local: &[u8]) -> Option<CoreProperty> {
    let property = match (namespace, local) {
        (Some(DUBLIN_CORE_NAMESPACE), b"title") => CoreProperty::Title,
        (Some(DUBLIN_CORE_NAMESPACE), b"subject") => CoreProperty::Subject,
        (Some(DUBLIN_CORE_NAMESPACE), b"creator") => CoreProperty::Author,
        (Some(DUBLIN_CORE_NAMESPACE), b"description") => CoreProperty::Description,
        (Some(DUBLIN_CORE_NAMESPACE), b"identifier") => CoreProperty::Identifier,
        (Some(DUBLIN_CORE_NAMESPACE), b"language") => CoreProperty::Language,
        (Some(DUBLIN_CORE_TERMS_NAMESPACE), b"created") => CoreProperty::Created,
        (Some(DUBLIN_CORE_TERMS_NAMESPACE), b"modified") => CoreProperty::Modified,
        (Some(CORE_PROPERTIES_NAMESPACE | STRICT_CORE_PROPERTIES_NAMESPACE), b"keywords") => {
            CoreProperty::Keywords
        },
        (Some(CORE_PROPERTIES_NAMESPACE | STRICT_CORE_PROPERTIES_NAMESPACE), b"lastModifiedBy") => {
            CoreProperty::LastModifiedBy
        },
        (Some(CORE_PROPERTIES_NAMESPACE | STRICT_CORE_PROPERTIES_NAMESPACE), b"revision") => {
            CoreProperty::Revision
        },
        (Some(CORE_PROPERTIES_NAMESPACE | STRICT_CORE_PROPERTIES_NAMESPACE), b"category") => {
            CoreProperty::Category
        },
        (Some(CORE_PROPERTIES_NAMESPACE | STRICT_CORE_PROPERTIES_NAMESPACE), b"contentStatus") => {
            CoreProperty::ContentStatus
        },
        (Some(CORE_PROPERTIES_NAMESPACE | STRICT_CORE_PROPERTIES_NAMESPACE), b"version") => {
            CoreProperty::Version
        },
        (Some(CORE_PROPERTIES_NAMESPACE | STRICT_CORE_PROPERTIES_NAMESPACE), b"lastPrinted") => {
            CoreProperty::LastPrinted
        },
        // OPC M4.3 permits only dcterms:created and dcterms:modified; any
        // other term, like any unknown element, is skipped by the caller.
        _ => return None,
    };
    Some(property)
}

fn validate_property_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    property: CoreProperty,
    recorder: &mut Recorder<'_>,
) -> Result<()> {
    validate_attributes(reader, element, Some(property), recorder)
}

fn keyword_lang(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    recorder: &mut Recorder<'_>,
) -> Result<Option<keyword::Lang>> {
    let mut language = None;
    for attribute in element.checked_attributes() {
        let attribute =
            attribute.map_err(|error| Error::Xml(format!("invalid keyword attribute: {error}")))?;
        let raw_key = attribute.key.as_ref();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?;
        if matches!(raw_key, b"xmlns") || raw_key.starts_with(b"xmlns:") {
            if value.as_bytes() == MARKUP_COMPATIBILITY_NAMESPACE {
                return Err(Error::Invalid(
                    "OPC M4.2 forbids the Markup Compatibility namespace".to_string(),
                ));
            }
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = bound_namespace(&namespace);
        reject_markup_compatibility(namespace)?;
        if namespace == Some(XML_NAMESPACE) && local.as_ref() == b"lang" {
            if language.is_some() {
                return Err(Error::Invalid(
                    "keyword contains duplicate xml:lang attributes".to_owned(),
                ));
            }
            language = Some(keyword::Lang::new(value.into_owned())?);
            continue;
        }
        recorder.record(
            CorePropertiesDefect::UnexpectedAttribute,
            Some(&String::from_utf8_lossy(element.local_name().as_ref())),
            Some(&String::from_utf8_lossy(raw_key)),
        )?;
    }
    Ok(language)
}

fn start_keyword_value(
    active: &mut Option<Active>,
    namespace: Option<&[u8]>,
    root_namespace: Option<&[u8]>,
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    empty: bool,
    recorder: &mut Recorder<'_>,
) -> Result<()> {
    let Some(Active::Keywords {
        value,
        child,
        bytes: _,
    }) = active.as_mut()
    else {
        return Err(Error::Invalid(
            "only cp:keywords may contain child elements".to_owned(),
        ));
    };
    if child.is_some() || namespace != root_namespace || element.local_name().as_ref() != b"value" {
        return Err(Error::Invalid(
            "cp:keywords permits only non-nested cp:value children".to_owned(),
        ));
    }
    let child_value = keyword::Value {
        text: String::new(),
        lang: keyword_lang(reader, element, recorder)?,
    };
    if empty {
        value.items.push(Item::Value(child_value));
    } else {
        *child = Some(child_value);
    }
    Ok(())
}

fn close_active(
    active: &mut Option<Active>,
    namespace: Option<&[u8]>,
    root_namespace: Option<&[u8]>,
    element: &BytesEnd<'_>,
    props: &mut Props,
    recorder: &mut Recorder<'_>,
) -> Result<()> {
    if matches!(active, Some(Active::Keywords { child: Some(_), .. })) {
        if namespace != root_namespace || element.local_name().as_ref() != b"value" {
            return Err(Error::Invalid(
                "malformed cp:value element nesting".to_owned(),
            ));
        }
        let Some(Active::Keywords { value, child, .. }) = active.as_mut() else {
            return Err(Error::Invalid("missing keyword state".to_owned()));
        };
        let child = child
            .take()
            .ok_or_else(|| Error::Invalid("missing cp:value state".to_owned()))?;
        value.items.push(Item::Value(child));
        return Ok(());
    }

    let finished = active
        .take()
        .ok_or_else(|| Error::Invalid("core property has no active state".to_owned()))?;
    match finished {
        Active::Text(property, value) => {
            if parse_property(namespace, element.local_name().as_ref()) != Some(property) {
                return Err(Error::Invalid(
                    "malformed core property element nesting".to_owned(),
                ));
            }
            apply_property(props, property, value, recorder)
        },
        Active::Keywords {
            value,
            child: None,
            bytes: _,
        } => {
            if namespace != root_namespace || element.local_name().as_ref() != b"keywords" {
                return Err(Error::Invalid(
                    "malformed cp:keywords element nesting".to_owned(),
                ));
            }
            props.keywords = Some(value);
            Ok(())
        },
        Active::Keywords { child: Some(_), .. } => Err(Error::Invalid(
            "unterminated cp:value in cp:keywords".to_owned(),
        )),
    }
}

fn validate_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    property: Option<CoreProperty>,
    recorder: &mut Recorder<'_>,
) -> Result<()> {
    let owner = property.map_or("coreProperties", property_str);
    let mut has_w3cdtf = false;
    // Every attribute OPC M4.4/M4.5 or the schema rejects is a producer-rule
    // deviation: ignored and recorded (LENIENCY_AUDIT 8.1).
    let unexpected = |recorder: &mut Recorder<'_>, raw_key: &[u8]| {
        recorder.record(
            CorePropertiesDefect::UnexpectedAttribute,
            Some(owner),
            Some(&String::from_utf8_lossy(raw_key)),
        )
    };
    for attribute in element.checked_attributes() {
        let attribute = attribute
            .map_err(|error| Error::Xml(format!("invalid core property attribute: {error}")))?;
        let raw_key = attribute.key.as_ref();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?;
        if matches!(raw_key, b"xmlns") || raw_key.starts_with(b"xmlns:") {
            if value.as_bytes() == MARKUP_COMPATIBILITY_NAMESPACE {
                return Err(Error::Invalid(
                    "OPC M4.2 forbids the Markup Compatibility namespace".to_string(),
                ));
            }
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = bound_namespace(&namespace);
        reject_markup_compatibility(namespace)?;
        if namespace == Some(XSI_NAMESPACE) && local.as_ref() == b"type" {
            let (value_namespace, value_local) =
                reader.resolver().resolve_element(QName(value.as_bytes()));
            let permitted = property.is_some_and(CoreProperty::requires_w3cdtf)
                && !has_w3cdtf
                && value.as_ref() == "dcterms:W3CDTF"
                && bound_namespace(&value_namespace) == Some(DUBLIN_CORE_TERMS_NAMESPACE)
                && value_local.as_ref() == b"W3CDTF";
            if !permitted {
                unexpected(recorder, raw_key)?;
            }
            // A misspelled type still declares one; the date is read as
            // W3CDTF either way, so only the attribute is recorded.
            has_w3cdtf = true;
            continue;
        }
        // OPC M4.4 forbids xml:lang; it and any other attribute are ignored.
        unexpected(recorder, raw_key)?;
    }
    if property.is_some_and(CoreProperty::requires_w3cdtf) && !has_w3cdtf {
        recorder.record(CorePropertiesDefect::MissingDateType, Some(owner), None)?;
    }
    Ok(())
}

fn reject_markup_compatibility(namespace: Option<&[u8]>) -> Result<()> {
    if namespace == Some(MARKUP_COMPATIBILITY_NAMESPACE) {
        return Err(Error::Invalid(
            "OPC M4.2 forbids the Markup Compatibility namespace".to_string(),
        ));
    }
    Ok(())
}

fn push_property_text(active: &mut Option<Active>, value: &str) -> Result<()> {
    let Some(active) = active.as_mut() else {
        if value.trim().is_empty() {
            return Ok(());
        }
        return Err(Error::Invalid(
            "non-whitespace text outside a core property".to_string(),
        ));
    };
    match active {
        Active::Text(_, text) => append_bounded(text, value),
        Active::Keywords {
            value: keywords,
            child,
            bytes,
        } => {
            *bytes = bytes.checked_add(value.len()).ok_or(Error::Limit {
                resource: "core property text bytes",
                max: MAX_PROPERTY_TEXT,
                actual: usize::MAX,
            })?;
            if *bytes > MAX_PROPERTY_TEXT {
                return Err(Error::Limit {
                    resource: "core property text bytes",
                    max: MAX_PROPERTY_TEXT,
                    actual: *bytes,
                });
            }
            if let Some(child) = child.as_mut() {
                child.text.push_str(value);
            } else {
                keywords.append_text(value.to_owned());
            }
            Ok(())
        },
    }
}

fn append_bounded(text: &mut String, value: &str) -> Result<()> {
    let length = text.len().checked_add(value.len()).ok_or(Error::Limit {
        resource: "core property text bytes",
        max: MAX_PROPERTY_TEXT,
        actual: usize::MAX,
    })?;
    if length > MAX_PROPERTY_TEXT {
        return Err(Error::Limit {
            resource: "core property text bytes",
            max: MAX_PROPERTY_TEXT,
            actual: length,
        });
    }
    text.push_str(value);
    Ok(())
}

fn apply_property(
    props: &mut Props,
    property: CoreProperty,
    value: String,
    recorder: &mut Recorder<'_>,
) -> Result<()> {
    match property {
        CoreProperty::Title => props.title = Some(value),
        CoreProperty::Subject => props.subject = Some(value),
        CoreProperty::Author => props.creator = Some(value),
        CoreProperty::Keywords => {
            return Err(Error::Invalid(
                "keyword mixed content bypassed its typed parser".to_owned(),
            ));
        },
        CoreProperty::Description => props.description = Some(value),
        CoreProperty::Identifier => props.identifier = Some(value),
        CoreProperty::Language => props.language = Some(value),
        CoreProperty::LastModifiedBy => props.last_modified_by = Some(value),
        CoreProperty::Revision => props.revision = Some(value),
        CoreProperty::Category => props.category = Some(value),
        CoreProperty::ContentStatus => props.content_status = Some(value),
        CoreProperty::Version => props.version = Some(value),
        // An unparseable date costs only that property (LENIENCY_AUDIT 8.1).
        CoreProperty::Created | CoreProperty::Modified => match time::W3c::new(value.clone()) {
            Ok(parsed) if property == CoreProperty::Created => props.created = Some(parsed),
            Ok(parsed) => props.modified = Some(parsed),
            Err(_) => recorder.record(
                CorePropertiesDefect::InvalidValue,
                Some(property_str(property)),
                Some(&value),
            )?,
        },
        CoreProperty::LastPrinted => match time::DateTime::new(value.clone()) {
            Ok(parsed) => props.last_printed = Some(parsed),
            Err(_) => recorder.record(
                CorePropertiesDefect::InvalidValue,
                Some(property_str(property)),
                Some(&value),
            )?,
        },
    }
    Ok(())
}

fn property_str(property: CoreProperty) -> &'static str {
    std::str::from_utf8(property_name(property)).unwrap_or("property")
}

fn property_name(property: CoreProperty) -> &'static [u8] {
    match property {
        CoreProperty::Title => b"title",
        CoreProperty::Subject => b"subject",
        CoreProperty::Author => b"creator",
        CoreProperty::Keywords => b"keywords",
        CoreProperty::Description => b"description",
        CoreProperty::Identifier => b"identifier",
        CoreProperty::Language => b"language",
        CoreProperty::LastModifiedBy => b"lastModifiedBy",
        CoreProperty::Revision => b"revision",
        CoreProperty::Category => b"category",
        CoreProperty::ContentStatus => b"contentStatus",
        CoreProperty::Version => b"version",
        CoreProperty::Created => b"created",
        CoreProperty::Modified => b"modified",
        CoreProperty::LastPrinted => b"lastPrinted",
    }
}

fn bound_namespace<'a>(namespace: &'a ResolveResult<'a>) -> Option<&'a [u8]> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => Some(*value),
        ResolveResult::Unbound | ResolveResult::Unknown(_) => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn read(package: &OpcPackage) -> Result<Option<Props>> {
        read_with_report(package, &mut ToleranceReport::default())
    }

    fn read_source_backed(package: &SourceBackedPackage) -> Result<Option<Props>> {
        read_source_backed_with_report(package, &mut ToleranceReport::default())
    }
    use crate::Props;
    use crate::properties::STRICT_CORE_PROPERTIES_RELATIONSHIP;
    use chrono::{DateTime, Utc};
    use litchi_core::OwnedSource;
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};
    use litchi_opc::{PackURI, SourceBackedPackage, part::BlobPart};
    use std::sync::Arc;

    fn package_with_core(path: &str, content_type: &str, xml: &[u8]) -> OpcPackage {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new(path).unwrap(),
            content_type.to_string(),
            xml.to_vec(),
        )));
        package.relate_to(path.trim_start_matches('/'), rt::CORE_PROPERTIES);
        package
    }

    fn poi_package(name: &str) -> OpcPackage {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/openxml4j")
            .join(name);
        OpcPackage::from_bytes(&std::fs::read(path).unwrap()).unwrap()
    }

    fn source_package_with_core(xml: &[u8], relationship: &str) -> SourceBackedPackage {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/docProps/core.xml").unwrap(),
            ct::OPC_CORE_PROPERTIES.to_owned(),
            xml.to_vec(),
        )));
        package.relate_to("docProps/core.xml", relationship);
        let mut bytes = Vec::new();
        package.to_stream(&mut bytes).unwrap();
        SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap()
    }

    fn source_package(package: &OpcPackage) -> SourceBackedPackage {
        let mut bytes = Vec::new();
        package.to_stream(&mut bytes).unwrap();
        SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap()
    }

    fn error_family(error: &Error) -> &'static str {
        match error {
            Error::Opc(_) => "opc",
            Error::Patch(_) => "patch",
            Error::Xml(_) => "xml",
            Error::Missing(_) => "missing",
            Error::ContentType { .. } => "content-type",
            Error::Relationship(_) => "relationship",
            Error::Invalid(_) => "invalid",
            Error::Limit { .. } => "limit",
            Error::Allocation { .. } => "allocation",
            Error::Uri(_) => "uri",
            #[cfg(feature = "vba-inspection")]
            Error::Vba(_) => "vba",
            Error::SpreadsheetXmlMaps(_) => "spreadsheet-xml-maps",
            Error::Mce(_) => "mce",
            Error::Decode(_) => "decode",
            Error::Io(_) => "io",
            Error::Fmt(_) => "fmt",
        }
    }

    #[test]
    fn selects_only_relationship_target_and_allows_absence() {
        let valid = br#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Target</dc:title></cp:coreProperties>"#;
        let decoy = br#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Decoy</dc:title></cp:coreProperties>"#;
        let mut package = package_with_core("/custom/real.xml", ct::OPC_CORE_PROPERTIES, valid);
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/docProps/core.xml").unwrap(),
            "application/xml".to_owned(),
            decoy.to_vec(),
        )));
        assert_eq!(
            read(&package).unwrap().unwrap().title.as_deref(),
            Some("Target")
        );
        assert!(read(&OpcPackage::new()).unwrap().is_none());
    }

    #[test]
    fn source_reader_matches_eager_for_present_absent_and_strict_properties() {
        let valid = br#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Target</dc:title></cp:coreProperties>"#;
        let eager = package_with_core("/docProps/core.xml", ct::OPC_CORE_PROPERTIES, valid);
        let mut bytes = Vec::new();
        eager.to_stream(&mut bytes).unwrap();
        let source = SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();
        assert_eq!(read_source_backed(&source).unwrap(), read(&eager).unwrap());

        let empty = OpcPackage::new();
        let mut bytes = Vec::new();
        empty.to_stream(&mut bytes).unwrap();
        let source = SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();
        assert_eq!(read_source_backed(&source).unwrap(), read(&empty).unwrap());

        let strict_xml = br#"<cp:coreProperties xmlns:cp="http://purl.oclc.org/ooxml/package/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Strict</dc:title></cp:coreProperties>"#;
        let source = source_package_with_core(strict_xml, STRICT_CORE_PROPERTIES_RELATIONSHIP);
        assert_eq!(
            read_source_backed(&source)
                .unwrap()
                .unwrap()
                .title
                .as_deref(),
            Some("Strict")
        );
    }

    #[test]
    fn source_reader_matches_eager_graph_failures() {
        let mut external = OpcPackage::new();
        external.relate_to_external("https://example.test/core.xml", rt::CORE_PROPERTIES);

        let mut dangling = OpcPackage::new();
        dangling.relate_to("missing.xml", rt::CORE_PROPERTIES);

        let wrong = package_with_core(
            "/docProps/core.xml",
            ct::WML_DOCUMENT,
            br#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>"#,
        );

        for package in [&external, &wrong] {
            let eager = read(package).unwrap_err();
            let source = source_package(package);
            let deferred = read_source_backed(&source).unwrap_err();
            assert_eq!(error_family(&eager), error_family(&deferred));
        }

        // LENIENCY_AUDIT 8.2: both readers repair a dangling relationship the
        // same way and record the same defect.
        let mut eager_report = ToleranceReport::default();
        assert!(
            read_with_report(&dangling, &mut eager_report)
                .unwrap()
                .is_none()
        );
        let source = source_package(&dangling);
        let mut source_report = ToleranceReport::default();
        assert!(
            read_source_backed_with_report(&source, &mut source_report)
                .unwrap()
                .is_none()
        );
        assert_eq!(eager_report, source_report);
        assert_eq!(eager_report.total(), 1);
    }

    #[test]
    fn rejects_external_dangling_and_wrong_content_type_relationships() {
        let mut external = OpcPackage::new();
        external.relate_to_external("https://example.test/core.xml", rt::CORE_PROPERTIES);
        assert!(matches!(read(&external), Err(Error::Relationship(_))));

        let mut dangling = OpcPackage::new();
        dangling.relate_to("missing.xml", rt::CORE_PROPERTIES);
        let mut report = ToleranceReport::default();
        assert!(read_with_report(&dangling, &mut report).unwrap().is_none());
        assert_eq!(
            report
                .iter()
                .map(|t| (t.defect(), t.part(), t.value()))
                .collect::<Vec<_>>(),
            vec![(
                CorePropertiesDefect::DanglingRelationship.into(),
                "/",
                Some("missing.xml")
            )]
        );
        // The writer still refuses to act on the dangling graph.
        assert!(matches!(
            super::super::write(&mut dangling, Props::new().title("T")),
            Err(Error::Missing(_))
        ));

        let wrong = package_with_core(
            "/core.xml",
            ct::WML_DOCUMENT,
            br#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>"#,
        );
        assert!(matches!(read(&wrong), Err(Error::ContentType { .. })));
    }

    /// LENIENCY_AUDIT 8.2: an orphaned part, a second core-typed part, and a
    /// relationship owned by a part are read past and recorded.
    #[test]
    fn orphaned_duplicate_and_misplaced_graph_anomalies_are_recorded() {
        let xml = br#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Selected</dc:title></cp:coreProperties>"#;
        let mut orphan = OpcPackage::new();
        orphan.add_part(Box::new(BlobPart::new(
            PackURI::new("/docProps/core.xml").unwrap(),
            ct::OPC_CORE_PROPERTIES.to_owned(),
            xml.to_vec(),
        )));
        let mut report = ToleranceReport::default();
        assert!(read_with_report(&orphan, &mut report).unwrap().is_none());
        assert_eq!(
            report
                .iter()
                .map(|t| (t.defect(), t.part()))
                .collect::<Vec<_>>(),
            vec![(
                CorePropertiesDefect::OrphanedPart.into(),
                "/docProps/core.xml"
            )]
        );
        assert!(read(&orphan).unwrap().is_none());

        let mut duplicate = package_with_core("/docProps/core.xml", ct::OPC_CORE_PROPERTIES, xml);
        duplicate.add_part(Box::new(BlobPart::new(
            PackURI::new("/docProps/core2.xml").unwrap(),
            ct::OPC_CORE_PROPERTIES.to_owned(),
            xml.to_vec(),
        )));
        let owner = PackURI::new("/docProps/core2.xml").unwrap();
        duplicate
            .get_part_mut(&owner)
            .unwrap()
            .rels_mut()
            .add_relationship(
                rt::CORE_PROPERTIES.to_owned(),
                "core.xml".to_owned(),
                "rIdX".to_owned(),
                false,
            );
        let mut report = ToleranceReport::default();
        let props = read_with_report(&duplicate, &mut report).unwrap().unwrap();
        assert_eq!(props.title.as_deref(), Some("Selected"));
        let mut recorded: Vec<_> = report
            .iter()
            .map(|t| (t.defect(), t.part().to_owned(), t.item().map(str::to_owned)))
            .collect();
        recorded.sort_by(|a, b| a.1.cmp(&b.1).then(a.2.cmp(&b.2)));
        assert!(recorded.contains(&(
            CorePropertiesDefect::DuplicatePart.into(),
            "/docProps/core2.xml".to_owned(),
            None
        )));
        assert!(recorded.contains(&(
            CorePropertiesDefect::MisplacedRelationship.into(),
            "/docProps/core2.xml".to_owned(),
            Some("rIdX".to_owned())
        )));
        assert_eq!(recorded.len(), 2);
    }

    #[test]
    fn parses_all_schema_fields_without_normalizing_lexical_values() {
        let xml = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>T</dc:title><dc:subject>S</dc:subject><dc:creator>A</dc:creator><dc:description>D</dc:description><dc:identifier>I</dc:identifier><dc:language>en-GB</dc:language><cp:keywords>K</cp:keywords><cp:lastModifiedBy>M</cp:lastModifiedBy><cp:revision>007</cp:revision><cp:category>C</cp:category><cp:contentStatus>Draft</cp:contentStatus><cp:version>2.0</cp:version><dcterms:created xsi:type="dcterms:W3CDTF">2006-10-13T18:06:00.123+03:00</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2007-06-20T07:59:00-13:00</dcterms:modified><cp:lastPrinted>2008-01-02T03:04:05Z</cp:lastPrinted></cp:coreProperties>"#;
        let (props, dialect) = decode(xml).unwrap();
        assert_eq!(dialect, Dialect::Transitional);
        assert_eq!(props.identifier.as_deref(), Some("I"));
        assert_eq!(props.language.as_deref(), Some("en-GB"));
        assert_eq!(props.version.as_deref(), Some("2.0"));
        assert_eq!(props.keywords.as_ref().and_then(Keywords::plain), Some("K"));
        assert_eq!(props.revision.as_deref(), Some("007"));
        assert_eq!(
            props.created.as_ref().map(time::W3c::as_str),
            Some("2006-10-13T18:06:00.123+03:00")
        );
        assert_eq!(
            props.modified.as_ref().map(time::W3c::as_str),
            Some("2007-06-20T07:59:00-13:00")
        );
        assert_eq!(
            props.last_printed.as_ref().map(time::DateTime::as_str),
            Some("2008-01-02T03:04:05Z")
        );
    }

    #[test]
    fn shared_writer_round_trips_through_strict_parser() {
        let created = DateTime::parse_from_rfc3339("2024-03-04T05:06:07+02:00")
            .unwrap()
            .with_timezone(&Utc);
        let properties = Props::new()
            .title("A & B")
            .creator("Writer")
            .identifier("urn:test:42")
            .language("en-US")
            .revision("12")
            .version("3.0")
            .created(created)
            .last_printed(created);
        let xml = properties.xml().unwrap();
        let (parsed, dialect) = decode(&xml).unwrap();
        assert_eq!(dialect, Dialect::Transitional);
        assert_eq!(parsed.title.as_deref(), Some("A & B"));
        assert_eq!(parsed.creator.as_deref(), Some("Writer"));
        assert_eq!(parsed.identifier.as_deref(), Some("urn:test:42"));
        assert_eq!(parsed.language.as_deref(), Some("en-US"));
        assert_eq!(parsed.revision.as_deref(), Some("12"));
        assert_eq!(parsed.version.as_deref(), Some("3.0"));
        assert_eq!(
            parsed.created.as_ref().map(time::W3c::as_str),
            Some("2024-03-04T03:06:07Z")
        );
        assert_eq!(
            parsed.last_printed.as_ref().map(time::DateTime::as_str),
            Some("2024-03-04T03:06:07Z")
        );
    }

    /// Decode `xml`, returning the properties and the recorded repairs as
    /// `(defect, item)` pairs.
    fn decode_recorded(xml: &str) -> (Props, Vec<(CorePropertiesDefect, Option<String>)>) {
        let mut report = ToleranceReport::default();
        let (props, _) = decode_reporting(
            xml,
            &mut Recorder {
                report: &mut report,
                part: "/docProps/core.xml",
                strict: false,
            },
        )
        .unwrap();
        let recorded = report
            .iter()
            .map(|tolerated| {
                let crate::leniency::Defect::CoreProperties(defect) = tolerated.defect() else {
                    panic!("unexpected defect family");
                };
                assert_eq!(tolerated.part(), "/docProps/core.xml");
                (defect, tolerated.item().map(str::to_owned))
            })
            .collect();
        (props, recorded)
    }

    #[test]
    fn rejects_wrong_root_dtd_and_entities() {
        assert!(decode("<props/>").is_err());
        assert!(decode("<!DOCTYPE x><x/>").is_err());
        let package = poi_package("CorePropertiesHasEntities.ooxml");
        assert!(read(&package).is_err());
    }

    /// LENIENCY_AUDIT 8.1: duplicates keep the first value, unknown elements
    /// (including other dcterms terms) are skipped with their subtree, and an
    /// unparseable date drops only that property — each recorded.
    #[test]
    fn duplicates_unknown_elements_and_invalid_dates_degrade_per_property() {
        let xml = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:x="urn:x"><dc:title>A</dc:title><dc:title>B<x:b/></dc:title><cp:contentType>text/plain</cp:contentType><x:ext><x:deep><dc:title>C</dc:title></x:deep></x:ext><dcterms:abstract>s</dcterms:abstract><x:empty/><cp:lastPrinted>2026-08-03</cp:lastPrinted><dc:creator>Me</dc:creator></cp:coreProperties>"#;
        let (props, recorded) = decode_recorded(xml);
        assert_eq!(props.title.as_deref(), Some("A"));
        assert_eq!(props.creator.as_deref(), Some("Me"));
        assert!(props.last_printed.is_none());
        let item = |name: &str| Some(name.to_owned());
        assert_eq!(
            recorded,
            vec![
                (CorePropertiesDefect::DuplicateProperty, item("title")),
                (CorePropertiesDefect::UnknownElement, item("cp:contentType")),
                (CorePropertiesDefect::UnknownElement, item("x:ext")),
                (
                    CorePropertiesDefect::UnknownElement,
                    item("dcterms:abstract")
                ),
                (CorePropertiesDefect::UnknownElement, item("x:empty")),
                (CorePropertiesDefect::InvalidValue, item("lastPrinted")),
            ]
        );

        let conformant = Props::new().title("A").creator("Me").xml().unwrap();
        assert!(decode_recorded(&conformant).1.is_empty());
    }

    /// LENIENCY_AUDIT 8.1: OPC M4.4/M4.5 attribute rules are producer rules —
    /// a missing W3CDTF type is accepted, and `xml:lang`, misplaced or
    /// misspelled `xsi:type`, and foreign attributes are ignored; recorded.
    #[test]
    fn attribute_deviations_are_ignored_and_recorded() {
        let xml = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:x="urn:x" x:root="1"><dc:subject xml:lang="en">S</dc:subject><dcterms:created>2026-01-02T03:04:05Z</dcterms:created><dcterms:modified xsi:type="dcterms:Other">2026-01-03</dcterms:modified><dc:title xsi:type="dcterms:W3CDTF">T</dc:title><cp:keywords x:k="1">k</cp:keywords></cp:coreProperties>"#;
        let (props, recorded) = decode_recorded(xml);
        assert_eq!(props.subject.as_deref(), Some("S"));
        assert_eq!(props.title.as_deref(), Some("T"));
        assert_eq!(
            props.created.as_ref().map(time::W3c::as_str),
            Some("2026-01-02T03:04:05Z")
        );
        assert_eq!(
            props.modified.as_ref().map(time::W3c::as_str),
            Some("2026-01-03")
        );
        assert_eq!(
            props.keywords.as_ref().map(Keywords::joined).as_deref(),
            Some("k")
        );
        let item = |name: &str| Some(name.to_owned());
        assert_eq!(
            recorded,
            vec![
                (
                    CorePropertiesDefect::UnexpectedAttribute,
                    item("coreProperties")
                ),
                (CorePropertiesDefect::UnexpectedAttribute, item("subject")),
                (CorePropertiesDefect::MissingDateType, item("created")),
                (CorePropertiesDefect::UnexpectedAttribute, item("modified")),
                (CorePropertiesDefect::UnexpectedAttribute, item("title")),
                (CorePropertiesDefect::UnexpectedAttribute, item("keywords")),
            ]
        );
    }

    /// LENIENCY_AUDIT 8.1: a non-UTF-8 payload is decoded lossily and a
    /// dialect mismatch is accepted; both are recorded on the read path.
    #[test]
    fn non_utf8_payloads_and_dialect_mismatches_are_recorded() {
        let mut latin1 = br#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>Jos"#.to_vec();
        latin1.push(0xe9);
        latin1.extend_from_slice(b"</dc:creator></cp:coreProperties>");
        let package = package_with_core("/docProps/core.xml", ct::OPC_CORE_PROPERTIES, &latin1);
        let mut report = ToleranceReport::default();
        let props = read_with_report(&package, &mut report).unwrap().unwrap();
        assert_eq!(props.creator.as_deref(), Some("Jos\u{fffd}"));
        assert_eq!(
            report.iter().map(|t| t.defect()).collect::<Vec<_>>(),
            vec![CorePropertiesDefect::NonUtf8.into()]
        );
        assert_eq!(report.defects()[0].part(), "/docProps/core.xml");

        let strict_xml = br#"<cp:coreProperties xmlns:cp="http://purl.oclc.org/ooxml/package/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>T</dc:title></cp:coreProperties>"#;
        let package = package_with_core("/docProps/core.xml", ct::OPC_CORE_PROPERTIES, strict_xml);
        let mut report = ToleranceReport::default();
        let props = read_with_report(&package, &mut report).unwrap().unwrap();
        assert_eq!(props.title.as_deref(), Some("T"));
        assert_eq!(
            report
                .iter()
                .map(|t| (t.defect(), t.value()))
                .collect::<Vec<_>>(),
            vec![(
                CorePropertiesDefect::DialectMismatch.into(),
                Some(rt::CORE_PROPERTIES)
            )]
        );
        let source = source_package(&package);
        let mut source_report = ToleranceReport::default();
        read_source_backed_with_report(&source, &mut source_report).unwrap();
        assert_eq!(source_report, report);
    }

    #[test]
    fn revision_is_a_lossless_string_including_empty_values() {
        for value in ["", "seven", "0007", " 7 "] {
            let xml = Props::new().revision(value).xml().unwrap();
            let (parsed, _) = decode(&xml).unwrap();
            assert_eq!(parsed.revision.as_deref(), Some(value));
        }
    }

    #[test]
    fn keywords_preserve_mixed_content_and_language_annotations() {
        let xml = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"><cp:keywords xml:lang="">lead<cp:value xml:lang="en-CA">colour</cp:value>;<cp:value xml:lang="fr">couleur</cp:value>tail</cp:keywords></cp:coreProperties>"#;
        let (props, _) = decode(xml).unwrap();
        let expected = Keywords::new()
            .lang(keyword::Lang::new("").unwrap())
            .text("lead")
            .value(keyword::Value::new("colour").lang(keyword::Lang::new("en-CA").unwrap()))
            .text(";")
            .value(keyword::Value::new("couleur").lang(keyword::Lang::new("fr").unwrap()))
            .text("tail");
        assert_eq!(props.keywords.as_ref(), Some(&expected));
        assert_eq!(
            props.keywords.as_ref().unwrap().joined(),
            "leadcolour;couleurtail"
        );

        let encoded = Props::new().keywords(expected.clone()).xml().unwrap();
        assert!(encoded.contains(r#"<cp:keywords xml:lang="">"#));
        assert!(encoded.contains(r#"<cp:value xml:lang="en-CA">colour</cp:value>"#));
        assert_eq!(decode(&encoded).unwrap().0.keywords, Some(expected));
    }

    #[test]
    fn rejects_invalid_keyword_children_and_language_placement() {
        let nested = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"><cp:keywords><cp:value><cp:value>x</cp:value></cp:value></cp:keywords></cp:coreProperties>"#;
        assert!(decode(nested).is_err());
        let wrong_child = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:x="urn:test"><cp:keywords><x:value>x</x:value></cp:keywords></cp:coreProperties>"#;
        assert!(decode(wrong_child).is_err());
    }

    #[test]
    fn w3cdtf_supports_partial_and_unzoned_forms_but_xsi_type_is_literal() {
        let xml = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dcterms:created xsi:type="dcterms:W3CDTF">2026</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2026-08-03T04:05:06</dcterms:modified><cp:lastPrinted>2026-08-03T04:05:06</cp:lastPrinted></cp:coreProperties>"#;
        let (props, _) = decode(xml).unwrap();
        assert_eq!(props.created.as_ref().map(time::W3c::as_str), Some("2026"));
        assert_eq!(
            props.modified.as_ref().map(time::W3c::as_str),
            Some("2026-08-03T04:05:06")
        );
        assert_eq!(
            props.last_printed.as_ref().map(time::DateTime::as_str),
            Some("2026-08-03T04:05:06")
        );

        // A prefix alias is not the literal `dcterms:W3CDTF`: recorded, and the
        // date is still read.
        let alias = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:t="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dcterms:created xsi:type="t:W3CDTF">2026</dcterms:created></cp:coreProperties>"#;
        let (props, recorded) = decode_recorded(alias);
        assert_eq!(props.created.as_ref().map(time::W3c::as_str), Some("2026"));
        assert_eq!(
            recorded,
            vec![(
                CorePropertiesDefect::UnexpectedAttribute,
                Some("created".to_owned())
            )]
        );
    }

    #[test]
    fn rejects_property_text_beyond_the_declared_budget() {
        let value = "x".repeat(MAX_PROPERTY_TEXT + 1);
        let xml = format!(
            r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>{value}</dc:title></cp:coreProperties>"#
        );
        assert!(matches!(
            decode(&xml),
            Err(Error::Limit {
                resource: "core property text bytes",
                max: MAX_PROPERTY_TEXT,
                actual,
            }) if actual == MAX_PROPERTY_TEXT + 1
        ));
    }

    #[test]
    fn matches_poi_success_timezone_and_m4_conformance_fixtures() {
        let success = read(&poi_package("OPCCompliance_CoreProperties_SUCCESS.docx"))
            .unwrap()
            .unwrap();
        assert_eq!(success.title.as_deref(), Some("MyTitle"));
        assert_eq!(success.revision.as_deref(), Some("2"));

        let timezone = read(&poi_package(
            "OPCCompliance_CoreProperties_AlternateTimezones.docx",
        ))
        .unwrap()
        .unwrap();
        assert_eq!(
            timezone.created.as_ref().map(time::W3c::as_str),
            Some("2006-10-13T18:06:00.123+03:00")
        );
        assert_eq!(
            timezone.modified.as_ref().map(time::W3c::as_str),
            Some("2007-06-20T07:59:00-13:00")
        );

        // OPC M4.2 (Markup Compatibility) stays a hard error.
        assert!(
            read(&poi_package(
                "OPCCompliance_CoreProperties_DoNotUseCompatibilityMarkupFAIL.docx"
            ))
            .is_err()
        );
        // LENIENCY_AUDIT 8.1: POI's M4.3-M4.5 failure fixtures are producer
        // deviations Word and LibreOffice open; each is repaired and recorded.
        for (name, expected) in [
            (
                "OPCCompliance_CoreProperties_DCTermsNamespaceLimitedUseFAIL.docx",
                CorePropertiesDefect::UnknownElement,
            ),
            (
                "OPCCompliance_CoreProperties_UnauthorizedXMLLangAttributeFAIL.docx",
                CorePropertiesDefect::UnexpectedAttribute,
            ),
            (
                "OPCCompliance_CoreProperties_LimitedXSITypeAttribute_NotPresentFAIL.docx",
                CorePropertiesDefect::MissingDateType,
            ),
            (
                "OPCCompliance_CoreProperties_LimitedXSITypeAttribute_PresentWithUnauthorizedValueFAIL.docx",
                CorePropertiesDefect::UnexpectedAttribute,
            ),
        ] {
            let mut report = ToleranceReport::default();
            let props = read_with_report(&poi_package(name), &mut report)
                .unwrap_or_else(|error| panic!("{name}: {error}"));
            assert!(props.is_some(), "{name}");
            assert!(
                report.iter().any(|t| t.defect() == expected.into()),
                "{name}: {report:?}"
            );
        }
        let mut report = ToleranceReport::default();
        read_with_report(
            &poi_package("OPCCompliance_CoreProperties_SUCCESS.docx"),
            &mut report,
        )
        .unwrap();
        assert!(report.is_clean(), "{report:?}");
    }

    #[test]
    fn poi_no_core_is_empty_and_multiple_relationships_fail_during_package_load() {
        let empty = read(&poi_package("OPCCompliance_NoCoreProperties.xlsx")).unwrap();
        assert!(empty.is_none());
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(
            "../../test-data/poi/test-data/openxml4j/OPCCompliance_CoreProperties_OnlyOneCorePropertiesPartFAIL.docx",
        );
        assert!(OpcPackage::from_bytes(&std::fs::read(path).unwrap()).is_err());
    }
}
