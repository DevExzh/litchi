//! Typed semantic values for the MS-XLSX rich-data extension family.

use crate::error::Result;

/// Bounded XML retained when this owner does not interpret a producer
/// extension. The bytes are never executed or resolved as a relationship.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Opaque {
    pub(crate) xml: Vec<u8>,
}

impl Opaque {
    /// Adopt one complete XML element or document after bounded validation.
    pub fn new(xml: impl Into<Vec<u8>>) -> Result<Self> {
        let xml = xml.into();
        super::codec::validate_fragment(&xml)?;
        Ok(Self { xml })
    }

    pub(crate) fn from_serialized(xml: Vec<u8>) -> Self {
        Self { xml }
    }

    /// The retained, inert XML bytes.
    #[must_use]
    pub fn xml(&self) -> &[u8] {
        &self.xml
    }
}

/// A rich-value key's wire data type.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum ValueType {
    Number,
    Integer,
    Boolean,
    Error,
    Text,
    RichValue,
    Array,
    SupportingBag,
    Unknown(String),
}

impl ValueType {
    pub(crate) fn parse(value: &str) -> Self {
        match value {
            "d" => Self::Number,
            "i" => Self::Integer,
            "b" => Self::Boolean,
            "e" => Self::Error,
            "s" => Self::Text,
            "r" => Self::RichValue,
            "a" => Self::Array,
            "spb" => Self::SupportingBag,
            other => Self::Unknown(other.to_owned()),
        }
    }

    pub(crate) fn token(&self) -> &str {
        match self {
            Self::Number => "d",
            Self::Integer => "i",
            Self::Boolean => "b",
            Self::Error => "e",
            Self::Text => "s",
            Self::RichValue => "r",
            Self::Array => "a",
            Self::SupportingBag => "spb",
            Self::Unknown(value) => value,
        }
    }
}

/// The fallback scalar kind used by one rich value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FallbackType {
    Boolean,
    Number,
    Error,
    Text,
}

impl FallbackType {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "b" => Ok(Self::Boolean),
            "n" => Ok(Self::Number),
            "e" => Ok(Self::Error),
            "s" => Ok(Self::Text),
            _ => Err(super::invalid(format!(
                "invalid rich-value fallback type '{value}'"
            ))),
        }
    }

    pub(crate) fn token(self) -> &'static str {
        match self {
            Self::Boolean => "b",
            Self::Number => "n",
            Self::Error => "e",
            Self::Text => "s",
        }
    }
}

/// One key and its declared rich-value data type.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Key {
    pub name: String,
    pub value_type: ValueType,
}

/// One rich-value structure definition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Structure {
    pub type_name: String,
    pub keys: Vec<Key>,
    pub opaque: Vec<Opaque>,
}

/// The `rvStructures` part model.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Structures {
    pub values: Vec<Structure>,
    pub extension_list: Option<Opaque>,
    pub opaque: Vec<Opaque>,
}

/// A rich-value fallback string and its typed presentation kind.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Fallback {
    pub value_type: FallbackType,
    pub value: String,
}

/// One value aligned by position with the keys in its structure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RichValue {
    pub structure: u32,
    pub fallback: Option<Fallback>,
    pub values: Vec<String>,
    pub opaque: Vec<Opaque>,
}

/// The `rvData` part model.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RichValueData {
    pub values: Vec<RichValue>,
    pub extension_list: Option<Opaque>,
    pub opaque: Vec<Opaque>,
}

/// A rich-array element's wire data type.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum ArrayValueType {
    Number,
    Integer,
    Boolean,
    Error,
    Text,
    RichValue,
    Array,
    Unknown(String),
}

impl ArrayValueType {
    pub(crate) fn parse(value: &str) -> Self {
        match value {
            "d" => Self::Number,
            "i" => Self::Integer,
            "b" => Self::Boolean,
            "e" => Self::Error,
            "s" => Self::Text,
            "r" => Self::RichValue,
            "a" => Self::Array,
            other => Self::Unknown(other.to_owned()),
        }
    }

    pub(crate) fn token(&self) -> &str {
        match self {
            Self::Number => "d",
            Self::Integer => "i",
            Self::Boolean => "b",
            Self::Error => "e",
            Self::Text => "s",
            Self::RichValue => "r",
            Self::Array => "a",
            Self::Unknown(value) => value,
        }
    }
}

/// One scalar in a rich array.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ArrayValue {
    pub value_type: ArrayValueType,
    pub value: String,
}

/// One row-major rich array.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Array {
    pub rows: u32,
    pub columns: u32,
    pub values: Vec<ArrayValue>,
    pub opaque: Vec<Opaque>,
}

/// The `arrayData` part model.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ArrayData {
    pub values: Vec<Array>,
    pub extension_list: Option<Opaque>,
    pub opaque: Vec<Opaque>,
}

/// A feature-property-bag domain named by MS-XLSX section 2.3.9.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum BagType {
    XfComplements,
    XfComplement,
    XfControls,
    Checkbox,
    DxfComplements,
    Unknown(String),
}

impl BagType {
    pub(crate) fn parse(value: &str) -> Self {
        match value {
            "XFComplements" => Self::XfComplements,
            "XFComplement" => Self::XfComplement,
            "XFControls" => Self::XfControls,
            "Checkbox" => Self::Checkbox,
            "DXFComplements" => Self::DxfComplements,
            other => Self::Unknown(other.to_owned()),
        }
    }

    pub(crate) fn token(&self) -> &str {
        match self {
            Self::XfComplements => "XFComplements",
            Self::XfComplement => "XFComplement",
            Self::XfControls => "XFControls",
            Self::Checkbox => "Checkbox",
            Self::DxfComplements => "DXFComplements",
            Self::Unknown(value) => value,
        }
    }
}

/// One typed scalar or nested reference inside an array property.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum PropertyValue {
    Bag(u32),
    Integer(String),
    Text(String),
    Boolean(bool),
    Decimal(String),
    Relationship(String),
    Unknown(Opaque),
}

/// One typed key/value entry in a feature property bag.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Property {
    Array {
        key: String,
        values: Vec<PropertyValue>,
    },
    Bag {
        key: String,
        index: u32,
    },
    Integer {
        key: String,
        value: String,
    },
    Text {
        key: String,
        value: String,
    },
    Boolean {
        key: String,
        value: bool,
    },
    Decimal {
        key: String,
        value: String,
    },
    Relationship {
        key: String,
        id: String,
    },
    Unknown(Opaque),
}

impl Property {
    /// The semantic key, when this is a typed feature property.
    #[must_use]
    pub fn key(&self) -> Option<&str> {
        match self {
            Self::Array { key, .. }
            | Self::Bag { key, .. }
            | Self::Integer { key, .. }
            | Self::Text { key, .. }
            | Self::Boolean { key, .. }
            | Self::Decimal { key, .. }
            | Self::Relationship { key, .. } => Some(key),
            Self::Unknown(_) => None,
        }
    }
}

/// One feature property bag and its optional producer metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Bag {
    pub bag_type: BagType,
    pub ext_ref: Option<String>,
    pub bag_extension: Option<u32>,
    pub attribute: Option<String>,
    pub properties: Vec<Property>,
    pub opaque: Vec<Opaque>,
}

impl Bag {
    /// Find the first typed property with the requested producer key.
    #[must_use]
    pub fn property(&self, key: &str) -> Option<&Property> {
        self.properties
            .iter()
            .find(|property| property.key() == Some(key))
    }

    /// Project a Checkbox bag into its inert semantic state.
    pub fn checkbox(&self) -> Result<Option<Checkbox>> {
        if !matches!(&self.bag_type, BagType::Checkbox) {
            return Ok(None);
        }
        let state = match self.property("default") {
            None => CheckboxState::Unchecked,
            Some(Property::Integer { value, .. }) => CheckboxState::parse(value)?,
            Some(_) => {
                return Err(super::invalid(
                    "Checkbox default must be an integer property",
                ));
            },
        };
        Ok(Some(Checkbox { default: state }))
    }
}

/// The `FeaturePropertyBags` part model.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Bags {
    pub count: Option<u32>,
    pub bag_extensions: Vec<Opaque>,
    pub values: Vec<Bag>,
    pub extension_list: Option<Opaque>,
    pub opaque: Vec<Opaque>,
}

impl Bags {
    /// Find a bag by its zero-based package-local index.
    #[must_use]
    pub fn get(&self, index: usize) -> Option<&Bag> {
        self.values.get(index)
    }

    /// Return the typed checkbox projection at one bag index.
    pub fn checkbox(&self, index: usize) -> Result<Option<Checkbox>> {
        self.get(index)
            .ok_or_else(|| super::invalid("feature property bag index is out of range"))?
            .checkbox()
    }
}

/// The inert checkbox default state defined by MS-XLSX section 2.3.9.4.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CheckboxState {
    Unchecked,
    Checked,
    Empty,
}

impl CheckboxState {
    fn parse(value: &str) -> Result<Self> {
        match value.trim() {
            "0" => Ok(Self::Unchecked),
            "1" => Ok(Self::Checked),
            "2" => Ok(Self::Empty),
            _ => Err(super::invalid("Checkbox default must be 0, 1, or 2")),
        }
    }
}

/// Typed semantic projection of a Checkbox feature property bag.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Checkbox {
    pub default: CheckboxState,
}

/// The `richValueRels` part model.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RichValueRels {
    pub ids: Vec<String>,
    pub extension_list: Option<Opaque>,
    pub opaque: Vec<Opaque>,
}

/// A typed `xfComplement` extension fragment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XfComplement {
    pub index: u32,
    pub opaque: Vec<Opaque>,
}

/// A typed `DXFComplement` extension fragment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DxfComplement {
    pub index: u32,
    pub opaque: Vec<Opaque>,
}

/// The target mode of one retained OPC relationship edge.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Mode {
    Internal,
    External,
}

/// One relationship edge retained by the rich-values package snapshot.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Link {
    pub source: Option<String>,
    pub id: String,
    pub relationship_type: String,
    pub target: String,
    pub resolved_target: Option<String>,
    pub mode: Mode,
}
