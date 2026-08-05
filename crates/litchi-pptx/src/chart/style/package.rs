//! OPC graph ownership for Office 2013 chart-style companion parts.

use litchi_opc::OpcPackage;
use litchi_opc::part::Part as OpcPart;

use super::{ColorDocument, Document};
use crate::{Error, Result};

pub const STYLE_CONTENT_TYPE: &str = "application/vnd.ms-office.chartstyle+xml";
pub const COLOR_CONTENT_TYPE: &str = "application/vnd.ms-office.chartcolorstyle+xml";
pub const STYLE_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2012/relationships/chartStyle";
pub const COLOR_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2012/relationships/chartColorStyle";

/// Borrowed chart-style companion part owned by a PPTX package.
pub struct Part<'a> {
    pub(crate) part: &'a dyn OpcPart,
}

/// Borrowed chart-color-style companion part owned by a PPTX package.
pub struct ColorPart<'a> {
    pub(crate) part: &'a dyn OpcPart,
}

impl<'a> Part<'a> {
    /// Validate and borrow one chart-style companion part.
    pub fn from_part(part: &'a dyn OpcPart) -> Result<Self> {
        if part.content_type() != STYLE_CONTENT_TYPE {
            return Err(Error::ContentType {
                expected: STYLE_CONTENT_TYPE.to_owned(),
                actual: part.content_type().to_owned(),
            });
        }
        Ok(Self { part })
    }

    /// Parse the XML owned by this package part with the shared DrawingML codec.
    pub fn parse(&self) -> Result<Document> {
        Ok(litchi_drawingml::chart::style::parse(self.part.blob())?)
    }

    /// The underlying OPC part.
    pub fn part(&self) -> &'a dyn OpcPart {
        self.part
    }
}

impl<'a> ColorPart<'a> {
    /// Validate and borrow one chart-color-style companion part.
    pub fn from_part(part: &'a dyn OpcPart) -> Result<Self> {
        if part.content_type() != COLOR_CONTENT_TYPE {
            return Err(Error::ContentType {
                expected: COLOR_CONTENT_TYPE.to_owned(),
                actual: part.content_type().to_owned(),
            });
        }
        Ok(Self { part })
    }

    /// Parse the XML owned by this package part with the shared DrawingML codec.
    pub fn parse(&self) -> Result<ColorDocument> {
        Ok(litchi_drawingml::chart::style::parse_color(
            self.part.blob(),
        )?)
    }

    /// The underlying OPC part.
    pub fn part(&self) -> &'a dyn OpcPart {
        self.part
    }
}

/// Resolve the chart-style siblings related to a ChartEx part.
pub fn discover(
    package: &OpcPackage,
    source: &dyn OpcPart,
) -> Result<(Option<Document>, Option<ColorDocument>)> {
    let mut style = None;
    let mut colors = None;
    for relationship in source.rels().iter() {
        let (expected, label) = if relationship.reltype() == STYLE_RELATIONSHIP_TYPE
            || relationship.reltype()
                == "http://purl.oclc.org/ooxml/officeDocument/relationships/chartStyle"
        {
            (STYLE_CONTENT_TYPE, "chart style")
        } else if relationship.reltype() == COLOR_RELATIONSHIP_TYPE
            || relationship.reltype()
                == "http://purl.oclc.org/ooxml/officeDocument/relationships/chartColorStyle"
        {
            (COLOR_CONTENT_TYPE, "chart color style")
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
        if expected == STYLE_CONTENT_TYPE {
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
