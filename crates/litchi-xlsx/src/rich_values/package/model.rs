//! Physical rich-value package snapshots.

use crate::error::Result;
use crate::rich_values::codec;
use crate::rich_values::model::{
    ArrayData, Bags, Link, Opaque, RichValueData, RichValueRels, Structures,
};

/// The XLSX content type that identifies one rich-value family part.
pub const RICH_VALUE_DATA_CONTENT_TYPE: &str = "application/vnd.ms-excel.rdRichValue+xml";
/// The XLSX content type that identifies rich-value structure definitions.
pub const RICH_VALUE_STRUCTURE_CONTENT_TYPE: &str =
    "application/vnd.ms-excel.rdRichValueStructure+xml";
/// The XLSX content type that identifies rich-value arrays.
pub const RICH_VALUE_ARRAY_CONTENT_TYPE: &str = "application/vnd.ms-excel.rdArray+xml";
/// The XLSX content type that identifies rich styles.
pub const RICH_STYLE_CONTENT_TYPE: &str = "application/vnd.ms-excel.richStyles+xml";
/// The XLSX content type that identifies supporting property-bag data.
pub const SUPPORTING_PROPERTY_BAG_CONTENT_TYPE: &str =
    "application/vnd.ms-excel.rdSupportingPropertyBag+xml";
/// The XLSX content type that identifies supporting property-bag structures.
pub const SUPPORTING_PROPERTY_BAG_STRUCTURE_CONTENT_TYPE: &str =
    "application/vnd.ms-excel.rdSupportingPropertyBagStructure+xml";
/// The XLSX content type that identifies rich-value type information.
pub const RICH_VALUE_TYPES_CONTENT_TYPE: &str = "application/vnd.ms-excel.rdRichValuetypes+xml";
/// The XLSX content type that identifies web-image supporting rich data.
pub const WEB_IMAGE_CONTENT_TYPE: &str = "application/vnd.ms-excel.rdrichvaluewebimage+xml";
/// The XLSX content type that identifies rich-value relationship metadata.
pub const RICH_VALUE_RELATIONSHIPS_CONTENT_TYPE: &str = "application/vnd.ms-excel.richvaluerel+xml";
/// The XLSX content type that identifies feature property bags.
pub const FEATURE_PROPERTY_BAG_CONTENT_TYPE: &str =
    "application/vnd.ms-excel.featurepropertybag+xml";

/// The relationship type associated with rich-value data.
pub const RICH_VALUE_DATA_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue";
/// The relationship type associated with rich-value structures.
pub const RICH_VALUE_STRUCTURE_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure";
/// The relationship type associated with rich-value arrays.
pub const RICH_VALUE_ARRAY_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2017/06/relationships/rdArray";
/// The relationship type associated with rich styles.
pub const RICH_STYLE_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2017/06/relationships/richStyles";
/// The relationship type associated with supporting property-bag data.
pub const SUPPORTING_PROPERTY_BAG_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2017/06/relationships/rdSupportingPropertyBag";
/// The relationship type associated with supporting property-bag structures.
pub const SUPPORTING_PROPERTY_BAG_STRUCTURE_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2017/06/relationships/rdSupportingPropertyBagStructure";
/// The relationship type associated with rich-value types.
pub const RICH_VALUE_TYPES_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes";
/// The relationship type associated with web-image supporting rich data.
pub const WEB_IMAGE_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2020/07/relationships/rdRichValueWebImage";
/// The relationship type associated with rich-value relationships.
pub const RICH_VALUE_RELATIONSHIPS_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel";
/// The relationship type associated with feature property bags.
pub const FEATURE_PROPERTY_BAG_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2022/11/relationships/FeaturePropertyBag";

/// A recognized rich-value part kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Kind {
    Data,
    Structures,
    Arrays,
    Styles,
    SupportingData,
    SupportingStructures,
    Types,
    WebImages,
    Relationships,
    FeatureBags,
}

impl Kind {
    /// Classify an OPC content type without inspecting its payload.
    pub fn from_content_type(content_type: &str) -> Option<Self> {
        match content_type {
            RICH_VALUE_DATA_CONTENT_TYPE => Some(Self::Data),
            RICH_VALUE_STRUCTURE_CONTENT_TYPE => Some(Self::Structures),
            RICH_VALUE_ARRAY_CONTENT_TYPE => Some(Self::Arrays),
            RICH_STYLE_CONTENT_TYPE => Some(Self::Styles),
            SUPPORTING_PROPERTY_BAG_CONTENT_TYPE => Some(Self::SupportingData),
            SUPPORTING_PROPERTY_BAG_STRUCTURE_CONTENT_TYPE => Some(Self::SupportingStructures),
            RICH_VALUE_TYPES_CONTENT_TYPE => Some(Self::Types),
            WEB_IMAGE_CONTENT_TYPE => Some(Self::WebImages),
            RICH_VALUE_RELATIONSHIPS_CONTENT_TYPE => Some(Self::Relationships),
            FEATURE_PROPERTY_BAG_CONTENT_TYPE => Some(Self::FeatureBags),
            _ => None,
        }
    }

    /// Return this kind's OPC content type.
    pub const fn content_type(self) -> &'static str {
        match self {
            Self::Data => RICH_VALUE_DATA_CONTENT_TYPE,
            Self::Structures => RICH_VALUE_STRUCTURE_CONTENT_TYPE,
            Self::Arrays => RICH_VALUE_ARRAY_CONTENT_TYPE,
            Self::Styles => RICH_STYLE_CONTENT_TYPE,
            Self::SupportingData => SUPPORTING_PROPERTY_BAG_CONTENT_TYPE,
            Self::SupportingStructures => SUPPORTING_PROPERTY_BAG_STRUCTURE_CONTENT_TYPE,
            Self::Types => RICH_VALUE_TYPES_CONTENT_TYPE,
            Self::WebImages => WEB_IMAGE_CONTENT_TYPE,
            Self::Relationships => RICH_VALUE_RELATIONSHIPS_CONTENT_TYPE,
            Self::FeatureBags => FEATURE_PROPERTY_BAG_CONTENT_TYPE,
        }
    }
}

/// A typed rich-value document or an inert, bounded document outside this slice.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Document {
    Data(RichValueData),
    Structures(Structures),
    Arrays(ArrayData),
    Relationships(RichValueRels),
    FeatureBags(Bags),
    Opaque(Opaque),
}

impl Document {
    /// Return the typed kind represented by this document.
    pub const fn kind(&self) -> Option<Kind> {
        match self {
            Self::Data(_) => Some(Kind::Data),
            Self::Structures(_) => Some(Kind::Structures),
            Self::Arrays(_) => Some(Kind::Arrays),
            Self::Relationships(_) => Some(Kind::Relationships),
            Self::FeatureBags(_) => Some(Kind::FeatureBags),
            Self::Opaque(_) => None,
        }
    }
}

/// One recognized XLSX rich-value part and its outgoing relationship edges.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Part {
    pub(crate) name: String,
    pub(crate) kind: Kind,
    pub(crate) document: Document,
    pub(crate) relationships: Vec<Link>,
}

impl Part {
    /// The stable OPC part name.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// The recognized content-family kind.
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// The typed or inert document owned by this part.
    pub fn document(&self) -> &Document {
        &self.document
    }

    /// All outgoing relationships from this part, including external edges.
    pub fn relationships(&self) -> &[Link] {
        &self.relationships
    }

    /// Serialize this document through the bounded rich-values codec.
    pub fn xml(&self) -> Result<Vec<u8>> {
        codec::write_part(self.kind, &self.document)
    }
}

/// A deterministic snapshot of rich-value parts and the complete OPC edge list.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Package {
    pub(crate) relationships: Vec<Link>,
    pub(crate) topology: Vec<Link>,
    pub(crate) parts: Vec<Part>,
}

impl Package {
    /// Package-root relationships, retained without resolving external targets.
    pub fn relationships(&self) -> &[Link] {
        &self.relationships
    }

    /// Every package-root and part relationship in deterministic order.
    pub fn topology(&self) -> &[Link] {
        &self.topology
    }

    /// Recognized rich-value family parts in deterministic name order.
    pub fn parts(&self) -> &[Part] {
        &self.parts
    }

    /// Find the first part of a recognized kind.
    pub fn part(&self, kind: Kind) -> Option<&Part> {
        self.parts.iter().find(|part| part.kind == kind)
    }
}
