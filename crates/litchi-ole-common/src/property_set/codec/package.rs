//! CFB/package integration and metadata projection for Property Sets.

use super::super::model::*;
use super::binary::{filetime_to_date, filetime_to_duration};
use super::support::allocation;
use litchi_cfb::{OleError, OleFile};
use std::io::{Read, Seek};

pub(super) fn try_path_refs(path: &[String]) -> Result<Vec<&str>, OleError> {
    let mut refs = try_vec_with_capacity(path.len(), "property-set stream path")?;
    refs.extend(path.iter().map(String::as_str));
    Ok(refs)
}

/// Read and project OLE Property Set streams from an opened compound file.
///
/// This is an extension trait because the CFB container owns OleFile.
pub trait PropertySetReader {
    /// Strictly parse a Property Set stream at path.
    fn property_set_stream(&mut self, path: &[&str]) -> Result<Stream, OleError>;

    /// Parse standard SummaryInformation and DocumentSummaryInformation metadata.
    fn get_metadata(&mut self) -> Result<Metadata, OleError>;
}

impl<R: Read + Seek> PropertySetReader for OleFile<R> {
    /// Strictly parse a Property Set stream at path.
    fn property_set_stream(&mut self, path: &[&str]) -> Result<Stream, OleError> {
        let data = self.open_stream(path)?;
        Stream::parse(&data)
    }

    /// Parse standard metadata. Missing streams are optional; malformed streams are errors.
    fn get_metadata(&mut self) -> Result<Metadata, OleError> {
        let mut metadata = Metadata::default();
        match PropertySetReader::property_set_stream(self, &["\u{0005}SummaryInformation"]) {
            Ok(stream) => {
                let section = stream
                    .sections
                    .first()
                    .ok_or_else(|| invalid("SummaryInformation has no section"))?;
                extract_summary_info(&mut metadata, section)?;
            },
            Err(OleError::StreamNotFound) => {},
            Err(error) => return Err(error),
        }
        match PropertySetReader::property_set_stream(self, &["\u{0005}DocumentSummaryInformation"])
        {
            Ok(stream) => {
                let section = stream
                    .sections
                    .first()
                    .ok_or_else(|| invalid("DocumentSummaryInformation has no section"))?;
                extract_document_summary_info(&mut metadata, section)?;
                for custom_section in stream.sections.iter().skip(1) {
                    metadata
                        .custom_properties
                        .try_reserve(custom_section.dictionary.len())
                        .map_err(|source| allocation("custom properties", source))?;
                    for (name, value) in custom_section.named_properties() {
                        if metadata.custom_properties.contains_key(name) {
                            return Err(invalid(format!(
                                "Duplicate custom property name '{name}'"
                            )));
                        }
                        let name = try_clone_string(name, "custom property name")?;
                        let value = try_clone_property_value(value)?;
                        metadata.custom_properties.insert(name, value);
                    }
                }
            },
            Err(OleError::StreamNotFound) => {},
            Err(error) => return Err(error),
        }
        Ok(metadata)
    }
}

fn extract_summary_info(metadata: &mut Metadata, section: &Section) -> Result<(), OleError> {
    if let Some(codepage) = section.codepage {
        metadata.codepage = Some(u32::from(codepage.id()));
    }
    metadata.title = extract_string(section.property(2))?;
    metadata.subject = extract_string(section.property(3))?;
    metadata.author = extract_string(section.property(4))?;
    metadata.keywords = extract_string(section.property(5))?;
    metadata.comments = extract_string(section.property(6))?;
    metadata.template = extract_string(section.property(7))?;
    metadata.last_saved_by = extract_string(section.property(8))?;
    metadata.revision_number = extract_string(section.property(9))?;
    if let Some(Value::Filetime(value)) = section.property(10) {
        metadata.edit_time = filetime_to_duration(*value);
    }
    if let Some(Value::Filetime(value)) = section.property(11) {
        metadata.last_printed_time = filetime_to_date(*value);
    }
    if let Some(Value::Filetime(value)) = section.property(12) {
        metadata.create_time = filetime_to_date(*value);
    }
    if let Some(Value::Filetime(value)) = section.property(13) {
        metadata.last_saved_time = filetime_to_date(*value);
    }
    metadata.num_pages = section.property(14).and_then(nonnegative_i4);
    metadata.num_words = section.property(15).and_then(nonnegative_i4);
    metadata.num_chars = section.property(16).and_then(nonnegative_i4);
    metadata.creating_application = extract_string(section.property(18))?;
    if let Some(Value::I4(value)) = section.property(19) {
        metadata.security = Some(*value as u32);
    }
    Ok(())
}

fn extract_document_summary_info(
    metadata: &mut Metadata,
    section: &Section,
) -> Result<(), OleError> {
    if metadata.codepage.is_none() {
        metadata.codepage = section.codepage.map(|page| u32::from(page.id()));
    }
    metadata.category = extract_string(section.property(2))?;
    metadata.manager = extract_string(section.property(14))?;
    metadata.company = extract_string(section.property(15))?;
    Ok(())
}

fn extract_string(value: Option<&Value>) -> Result<Option<String>, OleError> {
    let Some(value) = value else {
        return Ok(None);
    };
    match value {
        Value::Bstr(value) | Value::Lpstr(value) | Value::Lpwstr(value) => {
            Ok(Some(try_clone_string(value, "metadata string")?))
        },
        _ => Ok(None),
    }
}

fn nonnegative_i4(value: &Value) -> Option<u32> {
    match value {
        Value::I4(value) => u32::try_from(*value).ok(),
        _ => None,
    }
}
