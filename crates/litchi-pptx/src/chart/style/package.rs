//! OPC graph ownership for Office 2013 chart-style companion parts.

use litchi_opc::OpcPackage;
use litchi_opc::part::Part as OpcPart;

use super::codec;
use super::model::{ColorDocument, Document};
use super::{ColorPart, Part};
use crate::{Error, Result};

/// Resolve the chart-style siblings related to a ChartEx part.
pub fn discover(
    package: &OpcPackage,
    source: &dyn OpcPart,
) -> Result<(Option<Document>, Option<ColorDocument>)> {
    let mut style = None;
    let mut colors = None;
    for relationship in source.rels().iter() {
        let (expected, label) = if relationship.reltype() == codec::STYLE_RELATIONSHIP_TYPE
            || relationship.reltype()
                == "http://purl.oclc.org/ooxml/officeDocument/relationships/chartStyle"
        {
            (codec::STYLE_CONTENT_TYPE, "chart style")
        } else if relationship.reltype() == codec::COLOR_RELATIONSHIP_TYPE
            || relationship.reltype()
                == "http://purl.oclc.org/ooxml/officeDocument/relationships/chartColorStyle"
        {
            (codec::COLOR_CONTENT_TYPE, "chart color style")
        } else {
            continue;
        };
        if relationship.is_external() {
            return Err(Error::Relationship(format!(
                "external {label} relationships are not loaded"
            )));
        }
        if relationship.target_ref().starts_with('/')
            || relationship.target_ref().contains("..")
            || relationship.target_ref().ends_with('/')
        {
            return Err(Error::Relationship(format!(
                "invalid {label} relationship target"
            )));
        }
        let target = relationship.target_partname()?;
        if target.base_uri() != source.partname().base_uri() {
            return Err(Error::Relationship(format!(
                "{label} part is not a sibling of the ChartEx part"
            )));
        }
        let target_part = package
            .get_part(&target)
            .map_err(|_| Error::PartNotFound(format!("{label} target is missing")))?;
        if target_part.content_type() != expected {
            return Err(Error::ContentType {
                expected: expected.to_owned(),
                actual: target_part.content_type().to_owned(),
            });
        }
        if expected == codec::STYLE_CONTENT_TYPE {
            if style.is_some() {
                return Err(Error::Relationship(
                    "multiple chart style relationships".to_owned(),
                ));
            }
            style = Some(Part::from_part(target_part)?.parse()?);
        } else {
            if colors.is_some() {
                return Err(Error::Relationship(
                    "multiple chart color style relationships".to_owned(),
                ));
            }
            colors = Some(ColorPart::from_part(target_part)?.parse()?);
        }
    }
    Ok((style, colors))
}
