//! Structured data extraction from iWork documents.
//!
//! Archive traversal is kept in focused private adapters. The public result
//! remains composed entirely of the semantic values owned by the leaf crates:
//! [`litchi_numbers::Table`], [`litchi_keynote::Slide`], and
//! [`litchi_pages::Section`].

mod keynote;
mod numbers;
mod pages;
mod text;

use crate::Result;
use crate::application::Application;
use crate::bundle::Bundle;
use crate::detect::detect_application_from_document;
use crate::object_index::ObjectIndex;
use litchi_keynote::Slide;
use litchi_numbers::Table;
use litchi_pages::Section;
use prost::Message;

pub use litchi_iwa_structured::StructuredData;

/// Extract Numbers tables as canonical leaf-owned semantic values.
pub(crate) fn extract_tables(bundle: &Bundle, object_index: &ObjectIndex) -> Result<Vec<Table>> {
    numbers::extract(bundle, object_index)
}

/// Extract Keynote slides as canonical leaf-owned semantic values.
pub(crate) fn extract_slides(bundle: &Bundle, object_index: &ObjectIndex) -> Result<Vec<Slide>> {
    keynote::extract(bundle, object_index)
}

/// Extract the main Pages body section as a canonical leaf-owned value.
pub(crate) fn extract_sections(
    bundle: &Bundle,
    object_index: &ObjectIndex,
) -> Result<Vec<Section>> {
    pages::extract(bundle, object_index)
}

/// Extract all supported structured data from one document snapshot.
pub(crate) fn extract_all(bundle: &Bundle, object_index: &ObjectIndex) -> Result<StructuredData> {
    let application = structured_application(bundle)?;
    let (tables, slides, sections) = match application {
        Some(Application::Numbers) => (
            extract_tables(bundle, object_index)?,
            Vec::new(),
            Vec::new(),
        ),
        Some(Application::Keynote) => (
            Vec::new(),
            extract_slides(bundle, object_index)?,
            Vec::new(),
        ),
        Some(Application::Pages) => (
            Vec::new(),
            Vec::new(),
            extract_sections(bundle, object_index)?,
        ),
        Some(Application::Common) | None => (Vec::new(), Vec::new(), Vec::new()),
    };

    StructuredData::from_parts(tables, slides, sections).map_err(|error| {
        crate::Error::InvalidFormat(format!("invalid structured snapshot: {error}"))
    })
}

/// Determine which leaf extractor owns the document root before traversing it.
///
/// A package with no document archive is a valid empty/unsupported input for
/// the aggregate API. Once a document archive and root object are present,
/// malformed root evidence is rejected instead of being downgraded to three
/// empty semantic collections. Focused leaf extractors remain strict even
/// when called directly.
fn structured_application(bundle: &Bundle) -> Result<Option<Application>> {
    let Some(archive) = bundle.get_archive("Index/Document.iwa") else {
        return Ok(None);
    };
    let root = archive.object(1).ok_or_else(|| {
        crate::Error::InvalidFormat("structured document root object 1 is missing".to_owned())
    })?;

    let mut application = bundle
        .metadata()
        .detected_application()
        .and_then(|value| value.parse::<Application>().ok())
        .filter(|application| application.is_concrete());

    for candidate in root
        .messages
        .iter()
        .filter_map(|message| detect_application_from_document(&message.data))
    {
        merge_application(&mut application, candidate)?;
    }

    // Pages' root message type is unambiguous even when its protobuf payload
    // is the valid zero-byte encoding of an empty DocumentArchive. The strict
    // Pages extractor performs the actual decode and reference checks.
    if application.is_none() && root.messages.iter().any(|message| message.type_ == 10_000) {
        application = Some(Application::Pages);
    }

    // A minimal but valid Keynote root can contain only its show reference.
    // The full detector deliberately requires shared-document evidence, so
    // retain this narrow fallback for empty presentations while still leaving
    // malformed/ambiguous type-1 roots to the typed error below.
    if application.is_none() {
        application = root
            .messages
            .iter()
            .filter(|message| message.type_ == 1)
            .find_map(|message| {
                crate::protobuf::kn::DocumentArchive::decode(message.data.as_slice())
                    .ok()
                    .filter(|document| document.show.identifier != 0)
                    .map(|_| Application::Keynote)
            });
    }

    if application.is_none() && root.messages.iter().any(|message| !message.data.is_empty()) {
        return Err(crate::Error::InvalidFormat(
            "structured document root has no recognized iWork application payload".to_owned(),
        ));
    }

    Ok(application)
}

fn merge_application(current: &mut Option<Application>, candidate: Application) -> Result<()> {
    if !candidate.is_concrete() {
        return Ok(());
    }
    if let Some(previous) = *current
        && previous != candidate
    {
        return Err(crate::Error::InvalidFormat(format!(
            "structured document root contains conflicting applications: {previous} and {candidate}"
        )));
    }
    *current = Some(candidate);
    Ok(())
}

#[cfg(test)]
mod tests;
