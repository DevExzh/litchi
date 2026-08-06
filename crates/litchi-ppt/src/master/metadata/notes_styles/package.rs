//! OPC package construction and bounded package inspection.

use super::model::{MAX_PACKAGE_BYTES, MAX_PARTS, MAX_XML_BYTES, Package, Styles};
use crate::package::{Error, Result};
use crate::slide_round_trip::parse_embedded_xml_package;
use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::{OpcPackage, PackURI, XmlPart};
use std::io::Cursor;

const RECORD_NAME: &str = "RoundTripNotesMasterTextStyles12Atom";
const PART_NAME: &str = "/drs/slideMasters/slideMaster1.xml";
const PRESENTATIONML_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/presentationml/2006/main";

pub(super) fn from_xml(xml: &[u8]) -> Result<Styles> {
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "notes-master text styles XML exceeds {MAX_XML_BYTES} bytes"
        )));
    }

    let part_name = PackURI::new(PART_NAME).map_err(|error| {
        Error::InvalidFormat(format!(
            "notes-master text styles part name is invalid: {error}"
        ))
    })?;
    let mut package = OpcPackage::new();
    package.add_part(Box::new(XmlPart::new(
        part_name,
        content_type::PML_SLIDE_MASTER.to_string(),
        xml.to_vec(),
    )));
    package.rels_mut().add_relationship(
        relationship_type::SLIDE_MASTER.to_string(),
        PART_NAME.to_string(),
        "rId1".to_string(),
        false,
    );

    let mut output = Cursor::new(Vec::new());
    package.to_stream(&mut output).map_err(|error| {
        Error::InvalidFormat(format!(
            "notes-master text styles package write failed: {error}"
        ))
    })?;
    from_bytes(output.into_inner())
}

pub(super) fn from_bytes(data: Vec<u8>) -> Result<Styles> {
    if data.is_empty() {
        return Err(Error::Corrupted(
            "notes-master text styles package is empty".into(),
        ));
    }
    if data.len() > MAX_PACKAGE_BYTES {
        return Err(Error::InvalidFormat(format!(
            "notes-master text styles package exceeds {MAX_PACKAGE_BYTES} bytes"
        )));
    }

    let summary = parse_embedded_xml_package(
        &data,
        RECORD_NAME,
        content_type::PML_SLIDE_MASTER,
        PRESENTATIONML_NAMESPACE,
        b"txStyles",
    )?;
    if summary.part_count > MAX_PARTS {
        return Err(Error::InvalidFormat(format!(
            "notes-master text styles package contains more than {MAX_PARTS} parts"
        )));
    }

    // The package parser owns ZIP and OPC framing.  Re-open the bounded
    // package here to apply this owner’s aggregate part-size budget before the
    // bytes enter a snapshot.
    let package = OpcPackage::from_bytes(&data).map_err(|error| {
        Error::Corrupted(format!(
            "{RECORD_NAME} contains an invalid OPC package: {error}"
        ))
    })?;
    for part in package.iter_parts() {
        if part.blob().len() > MAX_XML_BYTES {
            return Err(Error::InvalidFormat(format!(
                "{RECORD_NAME} part {} exceeds {MAX_XML_BYTES} bytes",
                part.partname()
            )));
        }
    }

    Ok(Styles::from_package_model(Package::new(
        data,
        summary.part_count,
        summary.xml_part_name,
    )))
}
