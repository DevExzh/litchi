//! Bounded XML codec for ODF calculation settings.

use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use std::num::NonZeroUsize;

use super::model::{Iteration, IterationStatus, NullDate, Settings};
use crate::Error as OdfError;
use crate::core::{ResolvedReader, XmlStreamLimits};
use crate::limits::RawXmlLimitKind;

type OdfResult<T> = std::result::Result<T, OdfError>;
use crate::leniency::{ContentDefect, ToleranceReport};

const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";

/// Serialize optional calculation settings as an `office:calculation-settings` fragment.
///
/// # Errors
///
/// Returns an error when a lexical setting is malformed or exceeds a resource
/// limit.
pub fn write(out: &mut String, settings_option: Option<&Settings>) -> Result<()> {
    let Some(settings) = settings_option else {
        return Ok(());
    };
    settings.validate()?;
    out.push_str("<table:calculation-settings");
    write_optional_bool(out, "table:case-sensitive", settings.case_sensitive);
    write_optional_bool(out, "table:precision-as-shown", settings.precision_as_shown);
    write_optional_bool(
        out,
        "table:search-criteria-must-apply-to-whole-cell",
        settings.search_criteria_must_apply_to_whole_cell,
    );
    write_optional_bool(
        out,
        "table:automatic-find-labels",
        settings.automatic_find_labels,
    );
    write_optional_bool(
        out,
        "table:use-regular-expressions",
        settings.use_regular_expressions,
    );
    write_optional_bool(out, "table:use-wildcards", settings.use_wildcards);
    if let Some(year) = settings.null_year {
        write_attribute(out, "table:null-year", &year.get().to_string());
    }
    if settings.null_date.is_none() && settings.iteration.is_none() {
        out.push_str("/>");
        return Ok(());
    }
    out.push('>');
    if let Some(null_date) = &settings.null_date {
        out.push_str("<table:null-date");
        if null_date.value_type_date {
            write_attribute(out, "table:value-type", "date");
        }
        if let Some(value) = &null_date.date_value {
            write_attribute(out, "table:date-value", value);
        }
        out.push_str("/>");
    }
    if let Some(iteration) = &settings.iteration {
        out.push_str("<table:iteration");
        if let Some(status) = iteration.status {
            write_attribute(out, "table:status", status.as_str());
        }
        if let Some(steps) = iteration.steps {
            write_attribute(out, "table:steps", &steps.get().to_string());
        }
        if let Some(value) = &iteration.maximum_difference {
            write_attribute(out, "table:maximum-difference", value);
        }
        out.push_str("/>");
    }
    out.push_str("</table:calculation-settings>");
    Ok(())
}

/// Parse a bounded calculation-settings XML fragment under the caller's raw
/// XML policy.
///
/// The policy bounds the source bytes, nesting depth, events, attributes,
/// token bytes and namespace declarations while the fragment is read; an
/// exceeded dimension is returned as
/// [`LimitKind::RawXml`](crate::LimitKind::RawXml) with its observation.
/// Family owners' fused parsers drive the same scan through
/// [`CalculationHandler`] under their retained reader policy.
///
/// # Errors
///
/// Returns an error for malformed XML, invalid settings, or an exceeded
/// resource limit.
pub fn parse(xml: &str, limits: XmlStreamLimits) -> OdfResult<Option<Settings>> {
    parse_with_report(xml, limits).map(|(settings, _tolerance)| settings)
}

/// [`parse`], also returning the repairs applied while parsing.
///
/// An unknown or foreign-namespace child of `table:calculation-settings` is
/// skipped with its whole subtree and a `table:null-date` that follows
/// `table:iteration` is accepted, each recorded as a
/// [`ContentDefect`], exactly as LibreOffice's
/// `ScXMLCalculationSettingsContext` ignores children it does not handle.
///
/// # Errors
///
/// Returns an error for malformed XML, invalid settings, or exceeded resource
/// limits.
pub fn parse_with_report(
    xml: &str,
    limits: XmlStreamLimits,
) -> OdfResult<(Option<Settings>, ToleranceReport)> {
    let mut reader = ResolvedReader::from_xml(xml, limits)?;
    let mut buf = Vec::new();
    let mut scan = SettingsScan::default();
    loop {
        let (namespace, event) = reader.read_resolved_event_into(&mut buf)?;
        let is_table = is_namespace(&namespace);
        let is_office =
            matches!(&event, Event::Start(_)) && is_namespace_uri(&namespace, OFFICE_NAMESPACE);
        scan.on_event(
            is_table,
            is_office,
            &event,
            reader.resolver(),
            reader.decoder(),
        )?;
        if matches!(event, Event::Eof) {
            break;
        }
        scan.track_depth(&event);
        buf.clear();
    }
    Ok(scan.finish()?)
}

/// Enforce the pre-parse byte-size limit of the calculation-settings source.
///
/// Internal plumbing for family owners' fused parsers, not a supported API.
#[doc(hidden)]
pub fn validate_size(xml: &str) -> Result<()> {
    if xml.len() > super::MAX_XML_BYTES {
        return Err(raw_limit(
            RawXmlLimitKind::InputBytes,
            xml.len(),
            super::MAX_XML_BYTES,
        ));
    }
    Ok(())
}

/// The fused handler's historical envelope reports the same typed raw XML
/// dimension that [`ResolvedReader`] reports for an explicit policy, so a
/// fused open and the standalone [`parse`] agree on every bound.
fn raw_limit(kind: RawXmlLimitKind, observed: usize, maximum: usize) -> Error {
    Error::from(OdfError::Limit {
        kind: crate::LimitKind::RawXml(kind),
        observed: u64::try_from(observed).unwrap_or(u64::MAX),
        maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
    })
}

/// Classify the resolved namespace as `(is_table, is_office)` for
/// [`CalculationHandler::on_event`]; the resolved value borrows the reader
/// mutably, so callers classify it immediately after the read exactly where
/// the historical loop body classified it.
///
/// Internal plumbing for family owners' fused parsers, not a supported API.
#[doc(hidden)]
pub fn classify(namespace: &ResolveResult<'_>) -> (bool, bool) {
    (
        is_namespace(namespace),
        is_namespace_uri(namespace, OFFICE_NAMESPACE),
    )
}

/// Where the scan is relative to the document's single
/// `table:calculation-settings` element.
///
/// ```text
/// Outside --Start(calculation-settings)--> Open(Children)
/// Outside --Empty(calculation-settings)--> Closed
/// Open(Children) --End(calculation-settings)--> Closed
/// ```
///
/// A second `table:calculation-settings` element is an error in `Open` and
/// in `Closed`. Reaching the end of the document in `Open` is an error.
#[derive(Debug, Default)]
enum SettingsState {
    /// No settings element has started.
    #[default]
    Outside,
    /// Inside the settings element; `phase` says where.
    Open {
        settings: Settings,
        phase: OpenPhase,
    },
    /// The settings element has ended.
    Closed(Settings),
}

/// Where the scan is inside the open `table:calculation-settings` element.
///
/// ```text
/// Children --Start(null-date | iteration)--> Child
/// Child    --End-->                          Children
/// Children --Start(any other element)-->     Skipping { depth: 1 }
/// Skipping --Start-->                        Skipping { depth + 1 }
/// Skipping --End-->                          Skipping { depth - 1 }, or Children from depth 1
/// ```
///
/// An element inside `Child` is an error; empty elements never change the
/// phase.
#[derive(Debug, Clone, Copy)]
enum OpenPhase {
    /// Directly inside the settings element.
    Children,
    /// Inside a `table:null-date` or `table:iteration` written with an end tag.
    Child,
    /// Inside an unknown child subtree being skipped, `depth` (at
    /// least one) open elements deep.
    Skipping { depth: usize },
}

/// The calculation-settings scan shared by [`parse_with_report`] and
/// [`CalculationHandler`]: both apply the same checks and error messages at
/// the same events; each driver adds its own resource limits.
#[derive(Debug, Default)]
struct SettingsScan {
    state: SettingsState,
    depth: usize,
    document_body_depth: Option<usize>,
    tolerance: ToleranceReport,
}

impl SettingsScan {
    /// Process one resolved event. `is_office` is read for `Start` events
    /// only. The caller updates the depth afterwards with
    /// [`Self::track_depth`].
    #[inline]
    fn on_event(
        &mut self,
        is_table: bool,
        is_office: bool,
        event: &Event<'_>,
        resolver: &NamespaceResolver,
        decoder: quick_xml::encoding::Decoder,
    ) -> Result<()> {
        if let Event::Start(element) = event
            && is_office
            && matches!(
                element.local_name().as_ref(),
                b"chart" | b"drawing" | b"presentation" | b"spreadsheet" | b"text"
            )
        {
            self.document_body_depth = Some(self.depth);
        }
        if let Event::Start(element) | Event::Empty(element) = event
            && is_table
            && element.local_name().as_ref() == b"calculation-settings"
        {
            let is_document_body_child = self
                .document_body_depth
                .is_some_and(|value| self.depth == value + 1);
            if !is_document_body_child {
                return Err(Error::InvalidFormat(
                    "table:calculation-settings must be a direct office document-body child"
                        .to_string(),
                ));
            }
            if !matches!(self.state, SettingsState::Outside) {
                return Err(Error::InvalidFormat(
                    "duplicate table:calculation-settings".to_string(),
                ));
            }
            let settings = parse_settings_attributes(resolver, decoder, element)?;
            self.state = if matches!(event, Event::Start(_)) {
                SettingsState::Open {
                    settings,
                    phase: OpenPhase::Children,
                }
            } else {
                SettingsState::Closed(settings)
            };
            return Ok(());
        }
        let closed = match &mut self.state {
            SettingsState::Open { settings, phase } => Self::on_open_event(
                settings,
                phase,
                &mut self.tolerance,
                is_table,
                event,
                resolver,
                decoder,
            )?,
            SettingsState::Outside | SettingsState::Closed(_) => false,
        };
        if closed {
            self.state = match std::mem::take(&mut self.state) {
                SettingsState::Open { settings, .. } => SettingsState::Closed(settings),
                state @ (SettingsState::Outside | SettingsState::Closed(_)) => state,
            };
        }
        Ok(())
    }

    /// Process one event inside the open settings element and report whether
    /// it was the element's end tag.
    fn on_open_event(
        settings: &mut Settings,
        phase: &mut OpenPhase,
        tolerance: &mut ToleranceReport,
        is_table: bool,
        event: &Event<'_>,
        resolver: &NamespaceResolver,
        decoder: quick_xml::encoding::Decoder,
    ) -> Result<bool> {
        match (*phase, event) {
            (OpenPhase::Skipping { depth }, Event::Start(_)) => {
                *phase = OpenPhase::Skipping {
                    depth: depth.saturating_add(1),
                };
            },
            (OpenPhase::Skipping { .. }, Event::Empty(_) | Event::Text(_) | Event::CData(_)) => {},
            (OpenPhase::Skipping { .. }, Event::GeneralRef(reference)) => {
                skipped_reference(reference)?;
            },
            (OpenPhase::Skipping { depth }, Event::End(_)) => {
                *phase = if depth > 1 {
                    OpenPhase::Skipping { depth: depth - 1 }
                } else {
                    OpenPhase::Children
                };
            },
            (OpenPhase::Child, Event::Start(_) | Event::Empty(_)) => {
                return Err(Error::InvalidFormat(
                    "calculation setting children must be empty".to_string(),
                ));
            },
            (OpenPhase::Children, Event::Start(element)) => {
                *phase =
                    if settings_child(settings, resolver, decoder, is_table, element, tolerance)? {
                        OpenPhase::Child
                    } else {
                        OpenPhase::Skipping { depth: 1 }
                    };
            },
            (OpenPhase::Children, Event::Empty(element)) => {
                settings_child(settings, resolver, decoder, is_table, element, tolerance)?;
            },
            (OpenPhase::Child, Event::End(_)) => *phase = OpenPhase::Children,
            (OpenPhase::Children, Event::End(element)) => {
                return Ok(is_table && element.local_name().as_ref() == b"calculation-settings");
            },
            (OpenPhase::Children | OpenPhase::Child, Event::Text(text)) => {
                let text_content = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid calculation settings text: {error}"))
                })?;
                if !text_content.trim().is_empty() {
                    return Err(Error::InvalidFormat(
                        "table:calculation-settings cannot contain text".to_string(),
                    ));
                }
            },
            (OpenPhase::Children | OpenPhase::Child, Event::CData(text)) => {
                let cdata_content = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid calculation settings CDATA: {error}"))
                })?;
                if !cdata_content.trim().is_empty() {
                    return Err(Error::InvalidFormat(
                        "table:calculation-settings cannot contain CDATA".to_string(),
                    ));
                }
            },
            (OpenPhase::Children | OpenPhase::Child, Event::GeneralRef(_)) => {
                return Err(Error::InvalidFormat(
                    "table:calculation-settings cannot contain entity references".to_string(),
                ));
            },
            (
                OpenPhase::Children | OpenPhase::Child | OpenPhase::Skipping { .. },
                Event::Eof | Event::Comment(_) | Event::Decl(_) | Event::PI(_) | Event::DocType(_),
            ) => {},
        }
        Ok(false)
    }

    /// Follow the element nesting of one processed event.
    #[inline]
    fn track_depth(&mut self, event: &Event<'_>) {
        match event {
            Event::Start(_) => self.depth = self.depth.saturating_add(1),
            Event::End(_) => self.depth = self.depth.saturating_sub(1),
            Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_)
            | Event::Eof => {},
        }
    }

    /// Run the post-stream checks and return the parsed settings with the
    /// repairs applied while parsing.
    fn finish(self) -> Result<(Option<Settings>, ToleranceReport)> {
        let settings = match self.state {
            SettingsState::Outside => None,
            SettingsState::Open { .. } => {
                return Err(Error::InvalidFormat(
                    "unterminated table:calculation-settings".to_string(),
                ));
            },
            SettingsState::Closed(settings) => {
                settings.validate()?;
                Some(settings)
            },
        };
        Ok((settings, self.tolerance))
    }
}

/// Streaming event handler holding the [`parse`] scan state.
///
/// Internal plumbing for family owners' fused parsers, not a supported API:
/// a fused open parse drives one shared tokenizer through this handler while
/// [`parse`] drives the same scan under its caller's reader policy; both apply
/// the same checks and error messages at the same events.
#[doc(hidden)]
#[derive(Debug, Default)]
pub struct CalculationHandler {
    scan: SettingsScan,
    events: usize,
}

impl CalculationHandler {
    /// Process one resolved event whose namespace was classified with
    /// [`classify`], with the reader's resolver and decoder passed in exactly
    /// where the historical loop body consulted them.
    pub fn on_event(
        &mut self,
        is_table: bool,
        is_office: bool,
        event: &Event<'_>,
        resolver: &NamespaceResolver,
        decoder: quick_xml::encoding::Decoder,
    ) -> Result<()> {
        self.events = self.events.saturating_add(1);
        if self.events > super::MAX_EVENTS {
            return Err(raw_limit(
                RawXmlLimitKind::Events,
                self.events,
                super::MAX_EVENTS,
            ));
        }
        self.scan
            .on_event(is_table, is_office, event, resolver, decoder)?;
        if matches!(event, Event::Start(_)) && self.scan.depth >= super::MAX_DEPTH {
            return Err(raw_limit(
                RawXmlLimitKind::Depth,
                self.scan.depth.saturating_add(1),
                super::MAX_DEPTH,
            ));
        }
        self.scan.track_depth(event);
        Ok(())
    }

    /// Run the post-stream checks and return the parsed settings.
    pub fn finish(self) -> Result<Option<Settings>> {
        self.finish_with_report()
            .map(|(settings, _tolerance)| settings)
    }

    /// [`Self::finish`], also returning the repairs applied while parsing
    /// (see [`parse_with_report`]).
    pub fn finish_with_report(self) -> Result<(Option<Settings>, ToleranceReport)> {
        self.scan.finish()
    }
}

fn parse_settings_attributes(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
) -> Result<Settings> {
    Ok(Settings {
        case_sensitive: optional_bool(resolver, decoder, element, b"case-sensitive")?,
        precision_as_shown: optional_bool(resolver, decoder, element, b"precision-as-shown")?,
        search_criteria_must_apply_to_whole_cell: optional_bool(
            resolver,
            decoder,
            element,
            b"search-criteria-must-apply-to-whole-cell",
        )?,
        automatic_find_labels: optional_bool(resolver, decoder, element, b"automatic-find-labels")?,
        use_regular_expressions: optional_bool(
            resolver,
            decoder,
            element,
            b"use-regular-expressions",
        )?,
        use_wildcards: optional_bool(resolver, decoder, element, b"use-wildcards")?,
        null_year: optional_positive(resolver, decoder, element, b"null-year")?,
        null_date: None,
        iteration: None,
    })
}

/// Apply one child of `table:calculation-settings`, returning whether it
/// was a known setting.
///
/// An unknown or foreign-namespace child is recorded as
/// [`ContentDefect::CalculationSettingsUnknownChild`] and returns `false`;
/// the caller skips its subtree, the way LibreOffice's
/// `ScXMLCalculationSettingsContext::createFastChildContext` returns no
/// context for it.
fn settings_child(
    settings: &mut Settings,
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    is_table: bool,
    element: &BytesStart<'_>,
    tolerance: &mut ToleranceReport,
) -> Result<bool> {
    if !is_table || !matches!(element.local_name().as_ref(), b"null-date" | b"iteration") {
        tolerance.record(
            ContentDefect::CalculationSettingsUnknownChild,
            u16::from(is_table),
        );
        return Ok(false);
    }
    parse_settings_child(settings, resolver, decoder, element, tolerance)?;
    Ok(true)
}

/// Check an entity reference inside a skipped settings subtree: only the
/// predefined XML entities and character references are accepted, so
/// skipping never relaxes entity safety.
fn skipped_reference(reference: &[u8]) -> Result<()> {
    if matches!(reference, b"amp" | b"lt" | b"gt" | b"apos" | b"quot")
        || reference.first() == Some(&b'#')
    {
        Ok(())
    } else {
        Err(Error::InvalidFormat(
            "table:calculation-settings cannot contain entity references".to_string(),
        ))
    }
}

fn parse_settings_child(
    settings: &mut Settings,
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
    tolerance: &mut ToleranceReport,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"null-date" => {
            if settings.null_date.is_some() {
                return Err(Error::InvalidFormat(
                    "duplicate table:null-date".to_string(),
                ));
            }
            if settings.iteration.is_some() {
                // ODF schema order is null-date then iteration; LibreOffice
                // reads either order.
                tolerance.record(ContentDefect::CalculationSettingsReordered, 0);
            }
            let value_type = optional_attribute(resolver, decoder, element, b"value-type")?;
            if value_type.as_deref().is_some_and(|value| value != "date") {
                return Err(Error::InvalidFormat(
                    "table:null-date value type must be 'date'".to_string(),
                ));
            }
            settings.null_date = Some(NullDate {
                value_type_date: value_type.is_some(),
                date_value: optional_attribute(resolver, decoder, element, b"date-value")?,
            });
        },
        b"iteration" => {
            if settings.iteration.is_some() {
                return Err(Error::InvalidFormat(
                    "duplicate table:iteration".to_string(),
                ));
            }
            settings.iteration = Some(Iteration {
                status: optional_attribute(resolver, decoder, element, b"status")?
                    .map(|value| IterationStatus::parse(&value))
                    .transpose()?,
                steps: optional_positive(resolver, decoder, element, b"steps")?,
                maximum_difference: optional_attribute(
                    resolver,
                    decoder,
                    element,
                    b"maximum-difference",
                )?,
            });
        },
        _ => {
            return Err(Error::InvalidFormat(
                "unsupported calculation settings child".to_string(),
            ));
        },
    }
    Ok(())
}

fn optional_bool(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<bool>> {
    optional_attribute(resolver, decoder, element, name)?
        .map(|value| match value.as_str() {
            "true" | "1" => Ok(true),
            "false" | "0" => Ok(false),
            _ => Err(Error::InvalidFormat(format!(
                "invalid Boolean value '{value}'"
            ))),
        })
        .transpose()
}

fn optional_positive(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<NonZeroUsize>> {
    optional_attribute(resolver, decoder, element, name)?
        .map(|value| {
            value
                .parse::<usize>()
                .ok()
                .and_then(NonZeroUsize::new)
                .ok_or_else(|| Error::InvalidFormat(format!("invalid positive integer '{value}'")))
        })
        .transpose()
}

fn optional_attribute(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<String>> {
    for raw_attribute in element.attributes() {
        let attribute = raw_attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
        if attribute.value.len() > super::MAX_ATTRIBUTE_BYTES {
            return Err(Error::InvalidFormat(
                "calculation settings attribute exceeds the size limit".to_string(),
            ));
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if is_namespace(&namespace) && local.as_ref() == name {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map(|value| Some(value.into_owned()))
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")));
        }
    }
    Ok(None)
}

fn is_namespace(namespace: &ResolveResult<'_>) -> bool {
    is_namespace_uri(namespace, TABLE_NAMESPACE)
}

fn is_namespace_uri(namespace: &ResolveResult<'_>, uri: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == uri)
}

fn write_optional_bool(out: &mut String, name: &str, value: Option<bool>) {
    if let Some(boolean_value) = value {
        write_attribute(out, name, if boolean_value { "true" } else { "false" });
    }
}

fn write_attribute(out: &mut String, name: &str, value: &str) {
    out.push(' ');
    out.push_str(name);
    out.push_str("=\"");
    out.push_str(&escape_xml(value));
    out.push('"');
}
