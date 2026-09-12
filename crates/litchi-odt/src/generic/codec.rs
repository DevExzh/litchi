//! Bounded MIME, flat-document, and XML-part codecs for the generic owner.

use super::Family;
use super::flat::FlatMutationBudget;
use crate::constants;
use crate::core::ResolvedReader;
use litchi_core::{Error, ExecutionContext, ExecutionError, Resource, Result};
use quick_xml::events::Event;

pub(super) fn validate_flat_document_with_context(
    xml: &str,
    family: Family,
    context: &ExecutionContext,
) -> Result<()> {
    validate_flat_document_inner(xml, family, EventBudget::Context(context))
}

pub(super) fn validate_flat_document_with_budget(
    xml: &str,
    family: Family,
    budget: &FlatMutationBudget,
) -> Result<()> {
    validate_flat_document_inner(xml, family, EventBudget::Mutation(budget))
}

enum EventBudget<'a> {
    Context(&'a ExecutionContext),
    Mutation(&'a FlatMutationBudget),
}

impl EventBudget<'_> {
    fn event(&self, depth: usize, maximum_depth: &mut usize) -> Result<()> {
        match self {
            Self::Context(context) => {
                context.check().map_err(map_execution_error)?;
                context
                    .consume(Resource::Work, 1)
                    .map_err(map_execution_error)?;
                context
                    .consume(Resource::Objects, 1)
                    .map_err(map_execution_error)?;
                Self::observe_context_depth(context, depth, maximum_depth)
            },
            Self::Mutation(budget) => budget.event(depth),
        }
    }

    fn observe(&self, depth: usize, maximum_depth: &mut usize) -> Result<()> {
        match self {
            Self::Context(context) => Self::observe_context_depth(context, depth, maximum_depth),
            Self::Mutation(budget) => budget.observe_depth(depth),
        }
    }

    fn observe_context_depth(
        context: &ExecutionContext,
        depth: usize,
        maximum_depth: &mut usize,
    ) -> Result<()> {
        if depth <= *maximum_depth {
            return Ok(());
        }
        context.check().map_err(map_execution_error)?;
        context
            .consume(
                Resource::Depth,
                u64::try_from(depth - *maximum_depth).map_err(|_| {
                    Error::InvalidFormat(
                        "flat OpenDocument XML depth exceeds platform limits".into(),
                    )
                })?,
            )
            .map_err(map_execution_error)?;
        *maximum_depth = depth;
        Ok(())
    }
}

fn validate_flat_document_inner(
    xml: &str,
    family: Family,
    event_budget: EventBudget<'_>,
) -> Result<()> {
    const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";

    let mut reader = ResolvedReader::from_xml(xml);
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut body_seen = false;
    let mut family_body_seen = false;
    let mut maximum_depth = 0usize;
    let expected_body = match family {
        Family::Text => b"text".as_slice(),
        Family::Spreadsheet => b"spreadsheet".as_slice(),
        Family::Presentation => b"presentation".as_slice(),
        Family::Drawing => b"drawing".as_slice(),
        Family::Chart => b"chart".as_slice(),
        Family::Formula => b"formula".as_slice(),
        Family::Image => b"image".as_slice(),
        Family::Master | Family::Web | Family::Database => {
            unreachable!()
        },
    };
    loop {
        event_budget.event(depth, &mut maximum_depth)?;
        let (namespace, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!("invalid flat OpenDocument XML: {error}"))
        })?;
        match event {
            Event::Start(element) => {
                let is_office = crate::elements::xml::normalized_namespace_matches(
                    &namespace,
                    OFFICE_NAMESPACE,
                    "flat OpenDocument",
                )?;
                if depth == 0 {
                    if root_seen || !is_office || element.local_name().as_ref() != b"document" {
                        return Err(Error::InvalidFormat(
                            "flat OpenDocument must contain one office:document root".to_string(),
                        ));
                    }
                    root_seen = true;
                } else if depth == 1 && is_office && element.local_name().as_ref() == b"body" {
                    body_seen = true;
                } else if depth == 2 && is_office && element.local_name().as_ref() == expected_body
                {
                    family_body_seen = true;
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("flat OpenDocument nesting overflow".to_string())
                })?;
                event_budget.observe(depth, &mut maximum_depth)?;
            },
            Event::Empty(element) => {
                let is_office = crate::elements::xml::normalized_namespace_matches(
                    &namespace,
                    OFFICE_NAMESPACE,
                    "flat OpenDocument",
                )?;
                if depth == 0 {
                    return Err(Error::InvalidFormat(
                        "flat OpenDocument root cannot be empty".to_string(),
                    ));
                }
                event_budget.observe(depth + 1, &mut maximum_depth)?;
                if depth == 1 && is_office && element.local_name().as_ref() == b"body" {
                    body_seen = true;
                } else if depth == 2 && is_office && element.local_name().as_ref() == expected_body
                {
                    family_body_seen = true;
                }
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("unexpected flat OpenDocument closing tag".to_string())
                })?;
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Text(text) if depth == 0 && !text.iter().all(u8::is_ascii_whitespace) => {
                return Err(Error::InvalidFormat(
                    "text is not allowed outside the flat OpenDocument root".to_string(),
                ));
            },
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 => {
                return Err(Error::InvalidFormat(
                    "content is not allowed outside the flat OpenDocument root".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !root_seen || !root_closed || depth != 0 || !body_seen || !family_body_seen {
        return Err(Error::InvalidFormat(
            "flat OpenDocument is missing its complete family-specific body".to_string(),
        ));
    }
    Ok(())
}

fn map_execution_error(error: ExecutionError) -> Error {
    match error {
        ExecutionError::ResourceLimit(limit) => Error::ResourceLimit(limit),
        ExecutionError::Cancelled => Error::Other("flat OpenDocument operation cancelled".into()),
        other => Error::Other(format!("flat OpenDocument execution failed: {other}")),
    }
}

pub(super) fn decode_xml_part(bytes: Vec<u8>, path: &str) -> Result<String> {
    String::from_utf8(bytes)
        .map_err(|_error| Error::InvalidFormat(format!("invalid UTF-8 in {path}")))
}

pub(super) fn classify_mimetype(mimetype: &str) -> Option<(Family, bool)> {
    Some(match mimetype {
        constants::ODF_TEXT => (Family::Text, false),
        constants::ODF_TEXT_TEMPLATE => (Family::Text, true),
        constants::ODF_SPREADSHEET => (Family::Spreadsheet, false),
        constants::ODF_SPREADSHEET_TEMPLATE => (Family::Spreadsheet, true),
        constants::ODF_PRESENTATION => (Family::Presentation, false),
        constants::ODF_PRESENTATION_TEMPLATE => (Family::Presentation, true),
        constants::ODF_DRAWING => (Family::Drawing, false),
        constants::ODF_DRAWING_TEMPLATE => (Family::Drawing, true),
        constants::ODF_CHART => (Family::Chart, false),
        constants::ODF_CHART_TEMPLATE => (Family::Chart, true),
        constants::ODF_FORMULA => (Family::Formula, false),
        constants::ODF_FORMULA_TEMPLATE => (Family::Formula, true),
        constants::ODF_IMAGE => (Family::Image, false),
        constants::ODF_IMAGE_TEMPLATE => (Family::Image, true),
        constants::ODF_MASTER => (Family::Master, false),
        constants::ODF_MASTER_TEMPLATE => (Family::Master, true),
        constants::ODF_WEB => (Family::Web, true),
        constants::ODF_DATABASE => (Family::Database, false),
        _ => return None,
    })
}
