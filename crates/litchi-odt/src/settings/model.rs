//! Typed OpenDocument configuration settings values.

/// Semantic contents of one `office:settings` element.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Settings {
    /// Top-level configuration sets in document order.
    pub sets: Vec<ConfigSet>,
}

/// A named `config:config-item-set`.
#[derive(Debug, Clone, PartialEq)]
pub struct ConfigSet {
    pub name: String,
    pub children: Vec<ConfigNode>,
}

/// A setting node retained in document order.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum ConfigNode {
    Item(ConfigItem),
    Set(ConfigSet),
    IndexedMap(ConfigMap),
    NamedMap(ConfigMap),
}

/// One typed `config:config-item`.
#[derive(Debug, Clone, PartialEq)]
pub struct ConfigItem {
    pub name: String,
    pub value: ConfigValue,
}

/// Entries of an indexed or named configuration map.
#[derive(Debug, Clone, PartialEq)]
pub struct ConfigMap {
    pub name: String,
    pub entries: Vec<ConfigMapEntry>,
}

/// One map entry. Names are required in named maps and absent in indexed maps.
#[derive(Debug, Clone, PartialEq)]
pub struct ConfigMapEntry {
    pub name: Option<String>,
    pub children: Vec<ConfigNode>,
}

/// The scalar types defined for OpenDocument configuration items.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum ConfigValue {
    Boolean(bool),
    Short(i16),
    Int(i32),
    Long(i64),
    Double(f64),
    String(String),
    DateTime(String),
    Base64Binary(Vec<u8>),
}
