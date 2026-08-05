use super::codec::{
    CUSTOM_XFORM, HPLXSDR, MAX_CUSTOM_XFORM_BYTES, STTB_F_EXTEND, parse_custom_xml_transform,
};
use super::model::{Collection, Reference};
use crate::parts::fib::FileInformationBlock;

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
    (FileInformationBlock::parse(&fib_data).unwrap(), table)
}

#[test]
fn parses_schema_references() {
    let (fib, table) = fixture();
    let parsed = Collection::parse(&fib, &table)
        .unwrap()
        .expect("schemas present");
    assert_eq!(
        parsed.schemas(),
        [
            Reference {
                uri: "urn:one".to_string(),
                manifest_location: String::new(),
                elements: vec!["root".to_string(), "child".to_string()],
                attributes: vec!["attr".to_string()],
            },
            Reference {
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
    assert!(Collection::parse(&fib, &[]).unwrap().is_none());
}

#[test]
fn parses_empty_schema_list() {
    let mut fib_data = bare_fib();
    let table = hplxsdr(&[]);
    set_fib_pointer(&mut fib_data, HPLXSDR, 0, table.len() as u32);
    let fib = FileInformationBlock::parse(&fib_data).unwrap();
    let parsed = Collection::parse(&fib, &table)
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
    let parsed = Collection::parse(&fib, &table)
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
    assert!(Collection::parse(&fib, &negative).is_err());

    // Declared count exceeds what the byte length can hold.
    let mut inflated = table.clone();
    inflated[0..4].copy_from_slice(&3u32.to_le_bytes());
    assert!(Collection::parse(&fib, &inflated).is_err());

    // Trailing bytes after the last XSDR.
    let mut trailing = table.clone();
    trailing.push(0);
    let mut fib_data = fib.raw_data().to_vec();
    set_fib_pointer(&mut fib_data, HPLXSDR, 0, trailing.len() as u32);
    let trailing_fib = FileInformationBlock::parse(&fib_data).unwrap();
    assert!(Collection::parse(&trailing_fib, &trailing).is_err());

    // Non-extended STTB for the element table.
    let mut entry = length_prefixed("urn:x");
    entry.extend_from_slice(&length_prefixed(""));
    let mut bad_sttb = string_table(&["el"], 0);
    bad_sttb[0..2].copy_from_slice(&1u16.to_le_bytes());
    entry.extend_from_slice(&bad_sttb);
    entry.extend_from_slice(&string_table(&[], 0));
    assert!(Collection::parse(&fib, &hplxsdr(&[entry])).is_err());

    // Invalid UTF-16 in the URI (lone surrogate).
    let mut bad_uri = 1u16.to_le_bytes().to_vec();
    bad_uri.extend_from_slice(&0xD800u16.to_le_bytes());
    bad_uri.extend_from_slice(&length_prefixed(""));
    bad_uri.extend_from_slice(&string_table(&[], 0));
    bad_uri.extend_from_slice(&string_table(&[], 0));
    assert!(Collection::parse(&fib, &hplxsdr(&[bad_uri])).is_err());

    // Truncated table.
    let truncated = &table[..table.len() - 1];
    assert!(Collection::parse(&fib, truncated).is_err());

    // Pointer extending beyond the table stream.
    let mut fib_data = bare_fib();
    set_fib_pointer(&mut fib_data, HPLXSDR, 0, (table.len() + 1) as u32);
    let out_of_bounds = FileInformationBlock::parse(&fib_data).unwrap();
    assert!(Collection::parse(&out_of_bounds, &table).is_err());
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
