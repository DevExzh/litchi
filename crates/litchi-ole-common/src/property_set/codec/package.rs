//! CFB/package integration and metadata projection for Property Sets.

use super::super::binding::Binding;
use super::super::model::{
    Metadata, Section, Stream, Value, invalid, try_clone_property_value, try_clone_string,
    try_vec_with_capacity,
};
use super::binary::{filetime_to_date, filetime_to_duration};
use super::support::allocation;
use litchi_cfb::{OleError, OleFile, SharedOleFile};
use std::io::{Read, Seek};

/// Read and project OLE Property Set streams from an opened compound file.
///
/// This is an extension trait because the CFB container owns `OleFile`.
pub trait PropertySetReader {
    /// Parse standard `SummaryInformation` and `DocumentSummaryInformation` metadata.
    ///
    /// # Errors
    ///
    /// Returns an error if a present standard stream cannot be read or parsed,
    /// or if its metadata cannot be projected.
    fn get_metadata(&mut self) -> Result<Metadata, OleError> {
        project_metadata(|binding| self.property_set(binding))
    }

    /// Strictly parse a standard or GUID-derived Property Set binding.
    ///
    /// # Errors
    ///
    /// Returns an error if the bound stream cannot be opened or is not a valid
    /// Property Set stream.
    fn property_set(&mut self, binding: Binding) -> Result<Stream, OleError> {
        let name = binding.name();
        self.property_set_stream(&[name.as_str()])
    }

    /// Strictly parse a Property Set stream at path.
    ///
    /// # Errors
    ///
    /// Returns an error if the stream cannot be opened or is not a valid
    /// Property Set stream.
    fn property_set_stream(&mut self, path: &[&str]) -> Result<Stream, OleError>;
}

/// Read property sets through an immutable positional CFB view.
///
/// [`PropertySetReader`] predates [`SharedOleFile`] and keeps its mutable
/// receiver for compatibility with cursor-backed [`OleFile`] callers. A
/// shared CFB view has no cursor to mutate, so this additive trait exposes the
/// same checked metadata projection without requiring an `Arc` unwrap or a
/// second eager CFB parse.
pub trait SharedPropertySetReader {
    /// Parse standard `SummaryInformation` and `DocumentSummaryInformation`
    /// metadata from the positional source.
    ///
    /// # Errors
    ///
    /// Returns an error if a present standard stream cannot be read or parsed,
    /// or if its metadata cannot be projected.
    fn get_metadata(&self) -> Result<Metadata, OleError> {
        project_metadata(|binding| self.property_set(binding))
    }

    /// Strictly parse a standard or GUID-derived Property Set binding.
    ///
    /// # Errors
    ///
    /// Returns an error if the bound stream cannot be opened or is not a valid
    /// Property Set stream.
    fn property_set(&self, binding: Binding) -> Result<Stream, OleError> {
        let name = binding.name();
        self.property_set_stream(&[name.as_str()])
    }

    /// Strictly parse a Property Set stream at path.
    ///
    /// # Errors
    ///
    /// Returns an error if the stream cannot be opened or is not a valid
    /// Property Set stream.
    fn property_set_stream(&self, path: &[&str]) -> Result<Stream, OleError>;
}

impl<R: Read + Seek> PropertySetReader for OleFile<R> {
    /// Strictly parse a Property Set stream at path.
    fn property_set_stream(&mut self, path: &[&str]) -> Result<Stream, OleError> {
        let data = self.open_stream(path)?;
        Stream::parse(&data)
    }
}

impl SharedPropertySetReader for SharedOleFile {
    fn property_set_stream(&self, path: &[&str]) -> Result<Stream, OleError> {
        let data = self.open_stream(path)?;
        Stream::parse(&data)
    }
}

fn project_metadata<F>(mut property_set: F) -> Result<Metadata, OleError>
where
    F: FnMut(Binding) -> Result<Stream, OleError>,
{
    let mut metadata = Metadata::default();
    match property_set(Binding::SummaryInformation) {
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
    match property_set(Binding::DocumentSummaryInformation) {
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
                for (property_name, property_value) in custom_section.named_properties() {
                    if metadata.custom_properties.contains_key(property_name) {
                        return Err(invalid(format!(
                            "Duplicate custom property name '{property_name}'"
                        )));
                    }
                    let owned_name = try_clone_string(property_name, "custom property name")?;
                    let owned_value = try_clone_property_value(property_value)?;
                    metadata.custom_properties.insert(owned_name, owned_value);
                }
            }
        },
        Err(OleError::StreamNotFound) => {},
        Err(error) => return Err(error),
    }
    Ok(metadata)
}

pub(super) fn try_path_refs(path: &[String]) -> Result<Vec<&str>, OleError> {
    let mut refs = try_vec_with_capacity(path.len(), "property-set stream path")?;
    refs.extend(path.iter().map(String::as_str));
    Ok(refs)
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
        metadata.security = Some(u32::from_ne_bytes(value.to_ne_bytes()));
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

fn extract_string(candidate: Option<&Value>) -> Result<Option<String>, OleError> {
    let Some(property) = candidate else {
        return Ok(None);
    };
    if let Value::Bstr(text) | Value::Lpstr(text) | Value::Lpwstr(text) = property {
        Ok(Some(try_clone_string(text, "metadata string")?))
    } else {
        Ok(None)
    }
}

fn nonnegative_i4(property: &Value) -> Option<u32> {
    if let Value::I4(count) = property {
        u32::try_from(*count).ok()
    } else {
        None
    }
}
