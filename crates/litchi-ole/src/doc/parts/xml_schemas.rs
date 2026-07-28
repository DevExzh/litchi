//! Word 2003 XML schema reference tables (`Hplxsdr`) and the custom XML
//! save transform (`fcCustomXForm`).
//!
//! The `Hplxsdr` (MS-DOC 2.9.117) records the XML schema definitions a
//! document references: each `XSDR` (MS-DOC 2.9.352) carries the schema URI,
//! an optional expansion-pack manifest location, and the element and
//! attribute name string tables that namespace the `TIQ` name references
//! (MS-DOC 2.9.325) of structured document tags. `fcCustomXForm` points at a
//! UTF-16 path of the XML stylesheet Word applies when saving the document
//! in XML format.
//!
//! All structures are parsed as inert metadata: no schema or stylesheet is
//! fetched, resolved, or applied, and no document content is modified.

use super::fib::FileInformationBlock;
use crate::doc::package::{DocError, Result};

/// Table-pointer index of `fcHplxsdr`/`lcbHplxsdr`.
const HPLXSDR: usize = 136;
/// Table-pointer index of `fcCustomXForm`/`lcbCustomXForm`.
const CUSTOM_XFORM: usize = 140;

/// `fExtend` value of an extended STTB (MS-DOC 2.2.4).
const STTB_F_EXTEND: u16 = 0xFFFF;
/// Fixed header of an extended STTB with a 4-byte `cData`: `fExtend`,
/// `cData`, and `cbExtra` (MS-DOC 2.2.4).
const STTB_HEADER_LEN: usize = 8;
/// `cXSDR` and the 4-byte STTB `cData` fields are signed integers whose
/// minimum value is zero, so the sign bit must be clear.
const MAX_SIGNED_COUNT: u32 = 0x7FFF_FFFF;
/// Minimum size of one `XSDR`: two empty length-prefixed strings and two
/// empty STTB headers (MS-DOC 2.9.352).
const MIN_XSDR_LEN: usize = 2 + 2 + STTB_HEADER_LEN + STTB_HEADER_LEN;
/// Maximum byte length of the `fcCustomXForm` path array (MS-DOC
/// FibRgFcLcb2007).
const MAX_CUSTOM_XFORM_BYTES: u32 = 4168;

/// A single XML schema definition reference (`XSDR`, MS-DOC 2.9.352).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XmlSchemaReference {
    /// `wzURI`: the URI of the schema definition.
    pub uri: String,
    /// `wzManifestLocation`: the URI of the expansion-pack manifest the
    /// schema was loaded through, or empty when none was used.
    pub manifest_location: String,
    /// `sttbElements`: the element names of the schema, in table order.
    pub elements: Vec<String>,
    /// `sttbAttributes`: the attribute names of the schema, in table order.
    pub attributes: Vec<String>,
}

/// The XML schema definition references of a document (`Hplxsdr`, MS-DOC
/// 2.9.117), in `rgxsdr` order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentXmlSchemas {
    schemas: Vec<XmlSchemaReference>,
}

impl DocumentXmlSchemas {
    /// Parse the `Hplxsdr` addressed by the FIB, or `None` when the document
    /// carries none.
    pub fn parse(
        fib: &FileInformationBlock,
        table_stream: &[u8],
    ) -> Result<Option<DocumentXmlSchemas>> {
        let Some(data) = optional_slice(fib, table_stream, HPLXSDR, "Hplxsdr")? else {
            return Ok(None);
        };
        if data.len() < 4 {
            return Err(corrupted("Hplxsdr is truncated"));
        }
        let count = read_u32(data, 0, "Hplxsdr cXSDR")?;
        if count > MAX_SIGNED_COUNT {
            return Err(corrupted("Hplxsdr cXSDR is negative"));
        }
        let count = usize::try_from(count).map_err(|_| corrupted("Hplxsdr count exceeds usize"))?;
        if count > (data.len() - 4) / MIN_XSDR_LEN {
            return Err(corrupted("Hplxsdr byte length does not match its count"));
        }

        let mut schemas = Vec::with_capacity(count);
        let mut offset = 4usize;
        for _ in 0..count {
            let (schema, size) = parse_xsdr(&data[offset..])?;
            schemas.push(schema);
            offset += size;
        }
        if offset != data.len() {
            return Err(corrupted("Hplxsdr contains trailing bytes"));
        }
        Ok(Some(Self { schemas }))
    }

    /// All schema definition references in `rgxsdr` order.
    pub fn schemas(&self) -> &[XmlSchemaReference] {
        &self.schemas
    }

    /// Resolve a `TIQ` name reference against the element string table of
    /// the addressed schema, or `None` when either index is out of range.
    ///
    /// Per MS-DOC 2.9.325 step 4, the `TIQ` of an `FSDAP` (a structured tag
    /// attribute) names a string in `sttbElements`.
    pub fn element_name(&self, schema_index: u32, name_index: u32) -> Option<&str> {
        self.schemas
            .get(usize::try_from(schema_index).ok()?)?
            .elements
            .get(usize::try_from(name_index).ok()?)
            .map(String::as_str)
    }

    /// Resolve a `TIQ` name reference against the attribute string table of
    /// the addressed schema, or `None` when either index is out of range.
    ///
    /// Per MS-DOC 2.9.325 step 4, the `TIQ` of an `SDTI` (a structured tag
    /// node) names a string in `sttbAttributes`.
    pub fn attribute_name(&self, schema_index: u32, name_index: u32) -> Option<&str> {
        self.schemas
            .get(usize::try_from(schema_index).ok()?)?
            .attributes
            .get(usize::try_from(name_index).ok()?)
            .map(String::as_str)
    }
}

/// Parse the custom XML save transform path (`fcCustomXForm`): the full path
/// and file name of the XML stylesheet Word applies when saving the document
/// in XML format, or `None` when the document carries none.
///
/// The path is inert: it is exposed verbatim and never opened, resolved, or
/// applied.
pub fn parse_custom_xml_transform(
    fib: &FileInformationBlock,
    table_stream: &[u8],
) -> Result<Option<String>> {
    let Some((offset, length)) = fib.get_table_pointer(CUSTOM_XFORM) else {
        return Ok(None);
    };
    if length == 0 {
        return Ok(None);
    }
    if length > MAX_CUSTOM_XFORM_BYTES || length % 2 != 0 {
        return Err(corrupted(
            "fcCustomXForm length exceeds 4168 bytes or is not even",
        ));
    }
    let start =
        usize::try_from(offset).map_err(|_| corrupted("fcCustomXForm offset exceeds usize"))?;
    let end = start
        .checked_add(usize::try_from(length).map_err(|_| {
            corrupted("fcCustomXForm length exceeds usize")
        })?)
        .ok_or_else(|| corrupted("fcCustomXForm range overflows"))?;
    let data = table_stream
        .get(start..end)
        .ok_or_else(|| corrupted("fcCustomXForm extends beyond the table stream"))?;
    let mut units = data
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
        .collect::<Vec<_>>();
    // Producers commonly terminate the path array with a null code unit;
    // the spec defines the array by its byte length alone.
    if units.last() == Some(&0) {
        units.pop();
    }
    String::from_utf16(&units)
        .map(Some)
        .map_err(|_| corrupted("fcCustomXForm is invalid UTF-16"))
}

/// Parse one `XSDR`, returning the schema reference and the consumed byte
/// count.
fn parse_xsdr(data: &[u8]) -> Result<(XmlSchemaReference, usize)> {
    let mut offset = 0usize;
    let uri = parse_length_prefixed_string(data, &mut offset, "XSDR wzURI")?;
    let manifest_location =
        parse_length_prefixed_string(data, &mut offset, "XSDR wzManifestLocation")?;
    let elements = parse_string_table(data, &mut offset, "XSDR sttbElements")?;
    let attributes = parse_string_table(data, &mut offset, "XSDR sttbAttributes")?;
    Ok((
        XmlSchemaReference {
            uri,
            manifest_location,
            elements,
            attributes,
        },
        offset,
    ))
}

/// Parse a 16-bit length-prefixed UTF-16 string that is not null-terminated
/// (MS-DOC 2.9.352), advancing `offset` past it.
fn parse_length_prefixed_string(
    data: &[u8],
    offset: &mut usize,
    field: &str,
) -> Result<String> {
    let chars = usize::from(read_u16(data, *offset, field)?);
    let start = *offset + 2;
    let end = start
        .checked_add(chars.checked_mul(2).ok_or_else(|| {
            corrupted(format!("{field} byte length overflows"))
        })?)
        .ok_or_else(|| corrupted(format!("{field} range overflows")))?;
    let bytes = data
        .get(start..end)
        .ok_or_else(|| corrupted(format!("{field} is truncated")))?;
    *offset = end;
    decode_utf16(bytes, field)
}

/// Parse an extended STTB with a 4-byte `cData` (MS-DOC 2.2.4), advancing
/// `offset` past it. Per-entry extra data (`cbExtra`) is skipped verbatim.
fn parse_string_table(data: &[u8], offset: &mut usize, name: &str) -> Result<Vec<String>> {
    if data.len() < *offset + STTB_HEADER_LEN {
        return Err(corrupted(format!("{name} is truncated")));
    }
    if read_u16(data, *offset, "STTB fExtend")? != STTB_F_EXTEND {
        return Err(corrupted(format!("{name} is not an extended STTB")));
    }
    let count = read_u32(data, *offset + 2, "STTB cData")?;
    if count > MAX_SIGNED_COUNT {
        return Err(corrupted(format!("{name} cData is negative")));
    }
    let count = usize::try_from(count).map_err(|_| corrupted(format!("{name} count exceeds usize")))?;
    let extra = usize::from(read_u16(data, *offset + 6, "STTB cbExtra")?);
    let minimum_entry = 2usize
        .checked_add(extra)
        .ok_or_else(|| corrupted(format!("{name} entry size overflows")))?;
    if count > (data.len() - *offset - STTB_HEADER_LEN) / minimum_entry {
        return Err(corrupted(format!(
            "{name} byte length does not match its count"
        )));
    }

    let mut cursor = *offset + STTB_HEADER_LEN;
    let mut strings = Vec::with_capacity(count);
    for _ in 0..count {
        let chars = usize::from(read_u16(data, cursor, "STTB cchData")?);
        let start = cursor + 2;
        let end = start
            .checked_add(chars.checked_mul(2).ok_or_else(|| {
                corrupted(format!("{name} string byte length overflows"))
            })?)
            .ok_or_else(|| corrupted(format!("{name} string range overflows")))?;
        let bytes = data
            .get(start..end)
            .ok_or_else(|| corrupted(format!("{name} string is truncated")))?;
        strings.push(decode_utf16(bytes, name)?);
        cursor = end
            .checked_add(extra)
            .ok_or_else(|| corrupted(format!("{name} extra data range overflows")))?;
        if cursor > data.len() {
            return Err(corrupted(format!("{name} extra data is truncated")));
        }
    }
    *offset = cursor;
    Ok(strings)
}

fn decode_utf16(bytes: &[u8], field: &str) -> Result<String> {
    let units = bytes
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
        .collect::<Vec<_>>();
    String::from_utf16(&units).map_err(|_| corrupted(format!("{field} is invalid UTF-16")))
}

fn optional_slice<'a>(
    fib: &FileInformationBlock,
    table_stream: &'a [u8],
    index: usize,
    name: &str,
) -> Result<Option<&'a [u8]>> {
    let Some((offset, length)) = fib.get_table_pointer(index) else {
        return Ok(None);
    };
    if length == 0 {
        return Ok(None);
    }
    let start =
        usize::try_from(offset).map_err(|_| corrupted(format!("{name} offset exceeds usize")))?;
    let length =
        usize::try_from(length).map_err(|_| corrupted(format!("{name} length exceeds usize")))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted(format!("{name} range overflows")))?;
    table_stream
        .get(start..end)
        .map(Some)
        .ok_or_else(|| corrupted(format!("{name} extends beyond the table stream")))
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    litchi_core::binary::read_u16_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    litchi_core::binary::read_u32_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn corrupted(message: impl Into<String>) -> DocError {
    DocError::Corrupted(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    const FIB_POINTERS: usize = 145;

    fn set_fib_pointer(fib: &mut [u8], index: usize, offset: u32, length: u32) {
        let declared = u16::from_le_bytes([fib[152], fib[153]]);
        let count = declared.max(u16::try_from(index + 1).unwrap());
        fib[152..154].copy_from_slice(&count.to_le_bytes());
        let start = 154 + index * 8;
        fib[start..start + 4].copy_from_slice(&offset.to_le_bytes());
        fib[start + 4..start + 8].copy_from_slice(&length.to_le_bytes());
    }

    fn bare_fib() -> Vec<u8> {
        let mut fib_data = vec![0; 154 + FIB_POINTERS * 8];
        fib_data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        fib_data[2..4].copy_from_slice(&0x010Cu16.to_le_bytes());
        fib_data[152..154].copy_from_slice(&(FIB_POINTERS as u16).to_le_bytes());
        fib_data
    }

    fn utf16(text: &str) -> Vec<u8> {
        text.encode_utf16().flat_map(u16::to_le_bytes).collect()
    }

    fn length_prefixed(text: &str) -> Vec<u8> {
        let mut data = (text.encode_utf16().count() as u16).to_le_bytes().to_vec();
        data.extend_from_slice(&utf16(text));
        data
    }

    fn string_table(strings: &[&str], extra: u16) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&STTB_F_EXTEND.to_le_bytes());
        data.extend_from_slice(&(strings.len() as u32).to_le_bytes());
        data.extend_from_slice(&extra.to_le_bytes());
        for string in strings {
            data.extend_from_slice(&(string.encode_utf16().count() as u16).to_le_bytes());
            data.extend_from_slice(&utf16(string));
            data.extend(std::iter::repeat_n(0, usize::from(extra)));
        }
        data
    }

    fn xsdr(uri: &str, manifest: &str, elements: &[&str], attributes: &[&str]) -> Vec<u8> {
        let mut data = length_prefixed(uri);
        data.extend_from_slice(&length_prefixed(manifest));
        data.extend_from_slice(&string_table(elements, 0));
        data.extend_from_slice(&string_table(attributes, 0));
        data
    }

    fn hplxsdr(entries: &[Vec<u8>]) -> Vec<u8> {
        let mut data = (entries.len() as u32).to_le_bytes().to_vec();
        for entry in entries {
            data.extend_from_slice(entry);
        }
        data
    }

    /// Build a FIB plus table stream holding two schema references.
    fn fixture() -> (FileInformationBlock, Vec<u8>) {
        let mut fib_data = bare_fib();
        let table = hplxsdr(&[
            xsdr("urn:one", "", &["root", "child"], &["attr"]),
            xsdr("urn:two", "urn:manifest", &["only"], &[]),
        ]);
        set_fib_pointer(&mut fib_data, HPLXSDR, 0, table.len() as u32);
        (
            FileInformationBlock::parse(&fib_data).unwrap(),
            table,
        )
    }

    #[test]
    fn parses_schema_references() {
        let (fib, table) = fixture();
        let parsed = DocumentXmlSchemas::parse(&fib, &table)
            .unwrap()
            .expect("schemas present");
        assert_eq!(
            parsed.schemas(),
            [
                XmlSchemaReference {
                    uri: "urn:one".to_string(),
                    manifest_location: String::new(),
                    elements: vec!["root".to_string(), "child".to_string()],
                    attributes: vec!["attr".to_string()],
                },
                XmlSchemaReference {
                    uri: "urn:two".to_string(),
                    manifest_location: "urn:manifest".to_string(),
                    elements: vec!["only".to_string()],
                    attributes: Vec::new(),
                },
            ]
        );
        assert_eq!(parsed.element_name(0, 1), Some("child"));
        assert_eq!(parsed.attribute_name(0, 0), Some("attr"));
        assert_eq!(parsed.element_name(1, 0), Some("only"));
        assert_eq!(parsed.element_name(0, 2), None);
        assert_eq!(parsed.element_name(2, 0), None);
        assert_eq!(parsed.attribute_name(1, 0), None);
    }

    #[test]
    fn absent_table_yields_none() {
        let fib = FileInformationBlock::parse(&bare_fib()).unwrap();
        assert!(DocumentXmlSchemas::parse(&fib, &[]).unwrap().is_none());
    }

    #[test]
    fn parses_empty_schema_list() {
        let mut fib_data = bare_fib();
        let table = hplxsdr(&[]);
        set_fib_pointer(&mut fib_data, HPLXSDR, 0, table.len() as u32);
        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        let parsed = DocumentXmlSchemas::parse(&fib, &table)
            .unwrap()
            .expect("table present");
        assert!(parsed.schemas().is_empty());
    }

    #[test]
    fn skips_sttb_extra_data() {
        let mut entry = length_prefixed("urn:x");
        entry.extend_from_slice(&length_prefixed(""));
        entry.extend_from_slice(&string_table(&["el"], 3));
        entry.extend_from_slice(&string_table(&[], 0));
        let mut fib_data = bare_fib();
        let table = hplxsdr(&[entry]);
        set_fib_pointer(&mut fib_data, HPLXSDR, 0, table.len() as u32);
        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        let parsed = DocumentXmlSchemas::parse(&fib, &table)
            .unwrap()
            .expect("schemas present");
        assert_eq!(parsed.element_name(0, 0), Some("el"));
    }

    #[test]
    fn rejects_malformed_tables() {
        let (fib, table) = fixture();

        // Negative cXSDR.
        let mut negative = table.clone();
        negative[0..4].copy_from_slice(&(-1i32).to_le_bytes());
        assert!(DocumentXmlSchemas::parse(&fib, &negative).is_err());

        // Declared count exceeds what the byte length can hold.
        let mut inflated = table.clone();
        inflated[0..4].copy_from_slice(&3u32.to_le_bytes());
        assert!(DocumentXmlSchemas::parse(&fib, &inflated).is_err());

        // Trailing bytes after the last XSDR.
        let mut trailing = table.clone();
        trailing.push(0);
        let mut fib_data = fib.raw_data().to_vec();
        set_fib_pointer(&mut fib_data, HPLXSDR, 0, trailing.len() as u32);
        let trailing_fib = FileInformationBlock::parse(&fib_data).unwrap();
        assert!(DocumentXmlSchemas::parse(&trailing_fib, &trailing).is_err());

        // Non-extended STTB for the element table.
        let mut entry = length_prefixed("urn:x");
        entry.extend_from_slice(&length_prefixed(""));
        let mut bad_sttb = string_table(&["el"], 0);
        bad_sttb[0..2].copy_from_slice(&1u16.to_le_bytes());
        entry.extend_from_slice(&bad_sttb);
        entry.extend_from_slice(&string_table(&[], 0));
        assert!(DocumentXmlSchemas::parse(&fib, &hplxsdr(&[entry])).is_err());

        // Invalid UTF-16 in the URI (lone surrogate).
        let mut bad_uri = 1u16.to_le_bytes().to_vec();
        bad_uri.extend_from_slice(&0xD800u16.to_le_bytes());
        bad_uri.extend_from_slice(&length_prefixed(""));
        bad_uri.extend_from_slice(&string_table(&[], 0));
        bad_uri.extend_from_slice(&string_table(&[], 0));
        assert!(DocumentXmlSchemas::parse(&fib, &hplxsdr(&[bad_uri])).is_err());

        // Truncated table.
        let truncated = &table[..table.len() - 1];
        assert!(DocumentXmlSchemas::parse(&fib, truncated).is_err());

        // Pointer extending beyond the table stream.
        let mut fib_data = bare_fib();
        set_fib_pointer(&mut fib_data, HPLXSDR, 0, (table.len() + 1) as u32);
        let out_of_bounds = FileInformationBlock::parse(&fib_data).unwrap();
        assert!(DocumentXmlSchemas::parse(&out_of_bounds, &table).is_err());
    }

    #[test]
    fn parses_custom_xml_transform_path() {
        let mut fib_data = bare_fib();
        let table = utf16("C:\\transforms\\save.xsl");
        set_fib_pointer(&mut fib_data, CUSTOM_XFORM, 0, table.len() as u32);
        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        assert_eq!(
            parse_custom_xml_transform(&fib, &table).unwrap().as_deref(),
            Some("C:\\transforms\\save.xsl")
        );

        // A trailing null code unit is stripped.
        let mut terminated = utf16("save.xsl");
        terminated.extend_from_slice(&[0, 0]);
        let mut fib_data = bare_fib();
        set_fib_pointer(&mut fib_data, CUSTOM_XFORM, 0, terminated.len() as u32);
        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        assert_eq!(
            parse_custom_xml_transform(&fib, &terminated)
                .unwrap()
                .as_deref(),
            Some("save.xsl")
        );
    }

    #[test]
    fn absent_custom_xml_transform_yields_none() {
        let fib = FileInformationBlock::parse(&bare_fib()).unwrap();
        assert!(parse_custom_xml_transform(&fib, &[]).unwrap().is_none());
    }

    #[test]
    fn rejects_malformed_custom_xml_transform() {
        // Odd byte length.
        let mut fib_data = bare_fib();
        set_fib_pointer(&mut fib_data, CUSTOM_XFORM, 0, 3);
        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        assert!(parse_custom_xml_transform(&fib, &[0; 8]).is_err());

        // Length beyond the 4168-byte limit.
        let mut fib_data = bare_fib();
        set_fib_pointer(&mut fib_data, CUSTOM_XFORM, 0, MAX_CUSTOM_XFORM_BYTES + 2);
        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        assert!(parse_custom_xml_transform(&fib, &[0; 8]).is_err());

        // Invalid UTF-16.
        let mut fib_data = bare_fib();
        let table = 0xD800u16.to_le_bytes().to_vec();
        set_fib_pointer(&mut fib_data, CUSTOM_XFORM, 0, table.len() as u32);
        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        assert!(parse_custom_xml_transform(&fib, &table).is_err());
    }
}
