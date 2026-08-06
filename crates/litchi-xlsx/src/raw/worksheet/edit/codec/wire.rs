//! XML wire primitives for the lossless worksheet snapshot editor.
//!
//! These helpers deliberately operate on captured names and attributes rather
//! than rebuilding a DOM.  The snapshot layer can therefore replace only the
//! spans it owns while the package layer keeps every unrelated byte intact.

use litchi_core::xml::escape_xml;
use litchi_sheet::{COLUMNS, Column};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::BytesStart;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::super::super::required_u32;
use super::MCE;
use super::snapshot::{Attribute, Tag};
use crate::error::{Result, invalid};

pub(crate) fn write_tag(
    output: &mut Vec<u8>,
    tag: &Tag,
    empty: bool,
    removed: &[&str],
    appended: &[(&str, String)],
) {
    output.extend_from_slice(b"<");
    output.extend_from_slice(tag.name.as_bytes());
    for attribute in &tag.attributes {
        if removed.iter().any(|name| *name == attribute.name.as_ref()) {
            continue;
        }
        write_attribute(output, &attribute.name, &attribute.value);
    }
    for (name, value) in appended {
        write_attribute(output, name, value);
    }
    if empty {
        output.extend_from_slice(b"/>");
    } else {
        output.extend_from_slice(b">");
    }
}

pub(crate) fn write_attribute(output: &mut Vec<u8>, name: &str, value: &str) {
    output.extend_from_slice(b" ");
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    output.extend_from_slice(escape_xml(value).as_bytes());
    output.extend_from_slice(b"\"");
}

pub(crate) fn write_close(output: &mut Vec<u8>, name: &str) {
    output.extend_from_slice(b"</");
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b">");
}

pub(crate) fn tag(element: &BytesStart<'_>, decoder: Decoder) -> Result<Tag> {
    let name = std::str::from_utf8(element.name().as_ref())
        .map_err(|error| invalid(format!("worksheet element name is not UTF-8: {error}")))?
        .to_owned();
    let mut attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| invalid(format!("worksheet attribute name is not UTF-8: {error}")))?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| invalid(error.to_string()))?
            .into_owned();
        attributes.push(Attribute {
            name: name.into_boxed_str(),
            value: value.into_boxed_str(),
        });
    }
    Ok(Tag {
        name: name.into_boxed_str(),
        attributes: attributes.into_boxed_slice(),
    })
}

pub(crate) fn sibling_name(name: &str, local: &str) -> String {
    name.split_once(':').map_or_else(
        || local.to_owned(),
        |(prefix, _)| format!("{prefix}:{local}"),
    )
}

pub(crate) fn column_range(element: &BytesStart<'_>, decoder: Decoder) -> Result<(Column, Column)> {
    let min = required_u32(element, b"min", decoder, "worksheet column minimum")?;
    let max = required_u32(element, b"max", decoder, "worksheet column maximum")?;
    if min == 0 || min > max || max > COLUMNS {
        return Err(invalid(format!(
            "invalid worksheet column range '{min}:{max}' during edit"
        )));
    }
    Ok((Column::new(min - 1)?, Column::new(max - 1)?))
}

pub(crate) fn is_mce_name(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> bool {
    element.name().local_name().as_ref() == local
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == MCE)
}

pub(crate) fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid("worksheet XML position does not fit usize"))
}
