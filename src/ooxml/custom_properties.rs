//! OOXML custom document properties.
//!
//! This module provides functionality to add, read, modify, and remove custom
//! document properties from Office Open XML (OOXML) documents.
//!
//! Custom properties are stored in the "docProps/custom.xml" part of OOXML packages
//! and allow users to attach arbitrary metadata with typed values to documents.
//!
//! # Supported Property Types
//!
//! - **String** (`lpwstr` in OOXML)
//! - **Integer** (`i4` in OOXML) - 32-bit signed integer
//! - **Long** (`i8` in OOXML) - 64-bit signed integer  
//! - **Float** (`r4` in OOXML) - 32-bit floating point
//! - **Double** (`r8` in OOXML) - 64-bit floating point
//! - **Boolean** (`bool` in OOXML)
//! - **DateTime** (`filetime` in OOXML)
//!
//! # Example Usage
//!
//! ```rust,no_run
//! use litchi::ooxml::custom_properties::{CustomProperties, PropertyValue};
//!
//! let mut props = CustomProperties::new();
//!
//! // Add various types of properties
//! props.add_property("ProjectName", PropertyValue::String("MyProject".to_string()));
//! props.add_property("Version", PropertyValue::Integer(42));
//! props.add_property("Budget", PropertyValue::Double(12345.67));
//! props.add_property("IsApproved", PropertyValue::Boolean(true));
//!
//! // Get property value
//! if let Some(value) = props.get_property("Version") {
//!     println!("Version: {:?}", value);
//! }
//!
//! // Modify property
//! props.set_property("Version", PropertyValue::Integer(43));
//!
//! // Remove property
//! props.remove_property("Budget");
//!
//! // List all property names
//! for name in props.property_names() {
//!     println!("Property: {}", name);
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::constants::content_type as ct;
use crate::ooxml::opc::{OpcPackage, PackURI};
use chrono::{DateTime, Utc};
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::collections::HashMap;
use std::io::Cursor;

/// Fixed GUID format ID for custom properties as per OOXML specification.
///
/// All custom properties must use this format ID.
const FORMAT_ID: &str = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

/// XML namespace for custom properties.
const CUSTOM_PROPERTIES_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";

/// VTypes namespace for variant types.
const VTYPES_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

/// A custom document property value.
///
/// Represents the different types of values that can be stored in custom properties.
#[derive(Debug, Clone, PartialEq)]
pub enum PropertyValue {
    /// String value (lpwstr in OOXML)
    String(String),
    /// 32-bit signed integer (i4 in OOXML)
    Integer(i32),
    /// 64-bit signed integer (i8 in OOXML)
    Long(i64),
    /// 32-bit floating point (r4 in OOXML)
    Float(f32),
    /// 64-bit floating point (r8 in OOXML)
    Double(f64),
    /// Boolean value (bool in OOXML)
    Boolean(bool),
    /// DateTime value (filetime in OOXML)
    DateTime(DateTime<Utc>),
}

impl PropertyValue {
    /// Get the OOXML element name for this property type.
    fn element_name(&self) -> &'static str {
        match self {
            PropertyValue::String(_) => "lpwstr",
            PropertyValue::Integer(_) => "i4",
            PropertyValue::Long(_) => "i8",
            PropertyValue::Float(_) => "r4",
            PropertyValue::Double(_) => "r8",
            PropertyValue::Boolean(_) => "bool",
            PropertyValue::DateTime(_) => "filetime",
        }
    }

    /// Convert the value to its string representation for XML.
    fn to_xml_string(&self) -> String {
        match self {
            PropertyValue::String(s) => s.clone(),
            PropertyValue::Integer(i) => i.to_string(),
            PropertyValue::Long(l) => l.to_string(),
            PropertyValue::Float(f) => f.to_string(),
            PropertyValue::Double(d) => d.to_string(),
            PropertyValue::Boolean(b) => b.to_string(),
            PropertyValue::DateTime(dt) => {
                // FILETIME is Windows epoch (100-nanosecond intervals since 1601-01-01)
                const WINDOWS_EPOCH_OFFSET: i64 = 116_444_736_000_000_000;
                let unix_nanos = dt.timestamp_nanos_opt().unwrap_or(0);
                let windows_filetime = (unix_nanos / 100) + WINDOWS_EPOCH_OFFSET;
                windows_filetime.to_string()
            },
        }
    }

    /// Parse a property value from XML text content.
    fn from_xml_string(element: &str, text: &str) -> Result<Self> {
        match element {
            "lpwstr" => Ok(PropertyValue::String(text.to_string())),
            "i4" => text
                .parse::<i32>()
                .map(PropertyValue::Integer)
                .map_err(|e| OoxmlError::InvalidFormat(format!("Invalid i4 value: {}", e))),
            "i8" => text
                .parse::<i64>()
                .map(PropertyValue::Long)
                .map_err(|e| OoxmlError::InvalidFormat(format!("Invalid i8 value: {}", e))),
            "r4" => text
                .parse::<f32>()
                .map(PropertyValue::Float)
                .map_err(|e| OoxmlError::InvalidFormat(format!("Invalid r4 value: {}", e))),
            "r8" => text
                .parse::<f64>()
                .map(PropertyValue::Double)
                .map_err(|e| OoxmlError::InvalidFormat(format!("Invalid r8 value: {}", e))),
            "bool" => {
                let val = text.to_lowercase();
                match val.as_str() {
                    "true" | "1" => Ok(PropertyValue::Boolean(true)),
                    "false" | "0" => Ok(PropertyValue::Boolean(false)),
                    _ => Err(OoxmlError::InvalidFormat(format!(
                        "Invalid bool value: {}",
                        text
                    ))),
                }
            },
            "filetime" => {
                // Parse Windows FILETIME to DateTime
                const WINDOWS_EPOCH_OFFSET: i64 = 116_444_736_000_000_000;
                let filetime = text.parse::<i64>().map_err(|e| {
                    OoxmlError::InvalidFormat(format!("Invalid filetime value: {}", e))
                })?;
                let unix_nanos = (filetime - WINDOWS_EPOCH_OFFSET) * 100;
                DateTime::from_timestamp(
                    unix_nanos / 1_000_000_000,
                    (unix_nanos % 1_000_000_000) as u32,
                )
                .map(PropertyValue::DateTime)
                .ok_or_else(|| {
                    OoxmlError::InvalidFormat(format!("Invalid filetime value: {}", text))
                })
            },
            _ => Err(OoxmlError::InvalidFormat(format!(
                "Unknown property type: {}",
                element
            ))),
        }
    }
}

/// A single custom property with name, value, and internal ID.
#[derive(Debug, Clone)]
struct CustomProperty {
    /// Property name
    name: String,
    /// Property value
    value: PropertyValue,
    /// Internal property ID (pid attribute)
    pid: i32,
}

/// Collection of custom document properties.
///
/// This struct manages custom properties for OOXML documents, providing
/// methods to add, retrieve, modify, and remove properties.
///
/// Custom properties are stored in the `docProps/custom.xml` part of the
/// OOXML package and can contain various typed values.
#[derive(Debug, Clone, Default)]
pub struct CustomProperties {
    /// Map of property names to properties
    properties: HashMap<String, CustomProperty>,
    /// Next available property ID
    next_pid: i32,
}

impl CustomProperties {
    /// Create a new empty custom properties collection.
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi::ooxml::custom_properties::CustomProperties;
    ///
    /// let props = CustomProperties::new();
    /// ```
    pub fn new() -> Self {
        Self {
            properties: HashMap::new(),
            next_pid: 2, // PIDs start at 2 per OOXML spec
        }
    }

    /// Add a new custom property.
    ///
    /// If a property with the same name already exists, it will be replaced
    /// and the old value returned.
    ///
    /// # Arguments
    ///
    /// * `name` - The property name
    /// * `value` - The property value
    ///
    /// # Returns
    ///
    /// The previous value if a property with this name existed, or `None`.
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi::ooxml::custom_properties::{CustomProperties, PropertyValue};
    ///
    /// let mut props = CustomProperties::new();
    /// props.add_property("Author", PropertyValue::String("John Doe".to_string()));
    /// props.add_property("Version", PropertyValue::Integer(1));
    /// ```
    pub fn add_property(
        &mut self,
        name: impl Into<String>,
        value: PropertyValue,
    ) -> Option<PropertyValue> {
        let name = name.into();

        // If property exists, keep its PID, otherwise allocate new one
        let pid = if let Some(existing) = self.properties.get(&name) {
            existing.pid
        } else {
            let pid = self.next_pid;
            self.next_pid += 1;
            pid
        };

        let property = CustomProperty {
            name: name.clone(),
            value,
            pid,
        };

        self.properties.insert(name, property).map(|p| p.value)
    }

    /// Set a property value (alias for `add_property`).
    ///
    /// This method has the same behavior as `add_property` but uses a more
    /// intuitive name for updating existing properties.
    pub fn set_property(
        &mut self,
        name: impl Into<String>,
        value: PropertyValue,
    ) -> Option<PropertyValue> {
        self.add_property(name, value)
    }

    /// Get a property value by name.
    ///
    /// # Arguments
    ///
    /// * `name` - The property name to look up
    ///
    /// # Returns
    ///
    /// A reference to the property value if found, or `None`.
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi::ooxml::custom_properties::{CustomProperties, PropertyValue};
    ///
    /// let mut props = CustomProperties::new();
    /// props.add_property("Version", PropertyValue::Integer(42));
    ///
    /// if let Some(PropertyValue::Integer(ver)) = props.get_property("Version") {
    ///     println!("Version: {}", ver);
    /// }
    /// ```
    pub fn get_property(&self, name: &str) -> Option<&PropertyValue> {
        self.properties.get(name).map(|p| &p.value)
    }

    /// Remove a property by name.
    ///
    /// # Arguments
    ///
    /// * `name` - The property name to remove
    ///
    /// # Returns
    ///
    /// The removed property value if it existed, or `None`.
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi::ooxml::custom_properties::{CustomProperties, PropertyValue};
    ///
    /// let mut props = CustomProperties::new();
    /// props.add_property("TempData", PropertyValue::String("temp".to_string()));
    /// let removed = props.remove_property("TempData");
    /// assert!(removed.is_some());
    /// ```
    pub fn remove_property(&mut self, name: &str) -> Option<PropertyValue> {
        self.properties.remove(name).map(|p| p.value)
    }

    /// Check if a property with the given name exists.
    ///
    /// # Arguments
    ///
    /// * `name` - The property name to check
    ///
    /// # Returns
    ///
    /// `true` if a property with this name exists, `false` otherwise.
    pub fn contains(&self, name: &str) -> bool {
        self.properties.contains_key(name)
    }

    /// Get the number of custom properties.
    pub fn len(&self) -> usize {
        self.properties.len()
    }

    /// Check if the collection is empty.
    pub fn is_empty(&self) -> bool {
        self.properties.is_empty()
    }

    /// Get an iterator over all property names.
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi::ooxml::custom_properties::{CustomProperties, PropertyValue};
    ///
    /// let mut props = CustomProperties::new();
    /// props.add_property("Name", PropertyValue::String("Test".to_string()));
    /// props.add_property("Version", PropertyValue::Integer(1));
    ///
    /// for name in props.property_names() {
    ///     println!("Property: {}", name);
    /// }
    /// ```
    pub fn property_names(&self) -> impl Iterator<Item = &str> {
        self.properties.keys().map(|s| s.as_str())
    }

    /// Get an iterator over all properties (name and value pairs).
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi::ooxml::custom_properties::{CustomProperties, PropertyValue};
    ///
    /// let mut props = CustomProperties::new();
    /// props.add_property("Name", PropertyValue::String("Test".to_string()));
    ///
    /// for (name, value) in props.iter() {
    ///     println!("{}: {:?}", name, value);
    /// }
    /// ```
    pub fn iter(&self) -> impl Iterator<Item = (&str, &PropertyValue)> {
        self.properties
            .iter()
            .map(|(name, prop)| (name.as_str(), &prop.value))
    }

    /// Clear all custom properties.
    pub fn clear(&mut self) {
        self.properties.clear();
        self.next_pid = 2;
    }

    /// Generate XML content for the custom properties.
    ///
    /// This creates the XML structure for `docProps/custom.xml`.
    pub fn to_xml(&self) -> Result<String> {
        let mut writer = Writer::new(Cursor::new(Vec::new()));

        // XML declaration
        writer
            .write_event(Event::Decl(quick_xml::events::BytesDecl::new(
                "1.0",
                Some("UTF-8"),
                Some("yes"),
            )))
            .map_err(|e| OoxmlError::Xml(format!("Failed to write XML declaration: {}", e)))?;

        // Root <Properties> element
        let mut properties_elem = BytesStart::new("Properties");
        properties_elem.push_attribute(("xmlns", CUSTOM_PROPERTIES_NS));
        properties_elem.push_attribute(("xmlns:vt", VTYPES_NS));

        writer
            .write_event(Event::Start(properties_elem))
            .map_err(|e| OoxmlError::Xml(format!("Failed to write Properties element: {}", e)))?;

        // Sort properties by PID for consistent output
        let mut sorted_props: Vec<_> = self.properties.values().collect();
        sorted_props.sort_by_key(|p| p.pid);

        // Write each property
        for prop in sorted_props {
            let mut property_elem = BytesStart::new("property");
            property_elem.push_attribute(("fmtid", FORMAT_ID));
            property_elem.push_attribute(("pid", prop.pid.to_string().as_str()));
            property_elem.push_attribute(("name", prop.name.as_str()));

            writer
                .write_event(Event::Start(property_elem))
                .map_err(|e| OoxmlError::Xml(format!("Failed to write property element: {}", e)))?;

            // Write value element
            let value_elem_name = format!("vt:{}", prop.value.element_name());
            let value_start = BytesStart::new(&value_elem_name);
            writer
                .write_event(Event::Start(value_start))
                .map_err(|e| OoxmlError::Xml(format!("Failed to write value element: {}", e)))?;

            // Write value text
            let value_text = prop.value.to_xml_string();
            writer
                .write_event(Event::Text(BytesText::new(&value_text)))
                .map_err(|e| OoxmlError::Xml(format!("Failed to write value text: {}", e)))?;

            // Close value element
            writer
                .write_event(Event::End(BytesEnd::new(&value_elem_name)))
                .map_err(|e| OoxmlError::Xml(format!("Failed to close value element: {}", e)))?;

            // Close property element
            writer
                .write_event(Event::End(BytesEnd::new("property")))
                .map_err(|e| OoxmlError::Xml(format!("Failed to close property element: {}", e)))?;
        }

        // Close root element
        writer
            .write_event(Event::End(BytesEnd::new("Properties")))
            .map_err(|e| OoxmlError::Xml(format!("Failed to close Properties element: {}", e)))?;

        let result = writer.into_inner().into_inner();
        String::from_utf8(result)
            .map_err(|e| OoxmlError::Xml(format!("Invalid UTF-8 in generated XML: {}", e)))
    }

    /// Parse custom properties from XML content.
    ///
    /// # Arguments
    ///
    /// * `xml` - The XML content from `docProps/custom.xml`
    ///
    /// # Returns
    ///
    /// A `CustomProperties` instance populated with the parsed properties.
    pub fn from_xml(xml: &str) -> Result<Self> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut properties = HashMap::new();
        let mut max_pid = 1;

        // Current property being parsed
        let mut current_name: Option<String> = None;
        let mut current_pid: Option<i32> = None;

        loop {
            match reader.read_event() {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    let local_name = e.local_name();
                    let name_str = std::str::from_utf8(local_name.as_ref()).map_err(|e| {
                        OoxmlError::Xml(format!("Invalid UTF-8 in element name: {}", e))
                    })?;

                    match name_str {
                        "property" => {
                            // Parse property attributes
                            for attr in e.attributes() {
                                let attr = attr.map_err(|e| {
                                    OoxmlError::Xml(format!("Failed to parse attribute: {}", e))
                                })?;
                                let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                                let value = std::str::from_utf8(&attr.value).unwrap_or("");

                                match key {
                                    "name" => current_name = Some(value.to_string()),
                                    "pid" => {
                                        if let Ok(pid) = value.parse::<i32>() {
                                            current_pid = Some(pid);
                                            if pid > max_pid {
                                                max_pid = pid;
                                            }
                                        }
                                    },
                                    _ => {},
                                }
                            }
                        },
                        other
                            if other.starts_with("vt:")
                                || other.starts_with("lpwstr")
                                || other.starts_with("i4")
                                || other.starts_with("i8")
                                || other.starts_with("r4")
                                || other.starts_with("r8")
                                || other.starts_with("bool")
                                || other.starts_with("filetime") =>
                        {
                            // This is a value type element
                            let type_name = other.strip_prefix("vt:").unwrap_or(other);

                            // Read the text content
                            if let Ok(Event::Text(text)) = reader.read_event() {
                                let text_content =
                                    std::str::from_utf8(text.as_ref()).map_err(|e| {
                                        OoxmlError::Xml(format!("Invalid UTF-8 in text: {}", e))
                                    })?;

                                // Parse the value
                                if let (Some(name), Some(pid)) = (&current_name, current_pid) {
                                    let value =
                                        PropertyValue::from_xml_string(type_name, text_content)?;
                                    let property = CustomProperty {
                                        name: name.clone(),
                                        value,
                                        pid,
                                    };
                                    properties.insert(name.clone(), property);

                                    // Reset current property state
                                    current_name = None;
                                    current_pid = None;
                                }
                            }
                        },
                        _ => {},
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(format!("XML parsing error: {}", e))),
                _ => {},
            }
        }

        Ok(Self {
            properties,
            next_pid: max_pid + 1,
        })
    }
}

/// Extract custom properties from an OOXML package.
///
/// # Arguments
///
/// * `package` - The OOXML package to extract custom properties from
///
/// # Returns
///
/// A `CustomProperties` instance, which may be empty if no custom properties exist.
///
/// # Example
///
/// ```rust,no_run
/// use litchi::ooxml::OpcPackage;
/// use litchi::ooxml::custom_properties::extract_custom_properties;
///
/// let package = OpcPackage::open("document.docx")?;
/// let custom_props = extract_custom_properties(&package)?;
///
/// for (name, value) in custom_props.iter() {
///     println!("{}: {:?}", name, value);
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub fn extract_custom_properties(package: &OpcPackage) -> Result<CustomProperties> {
    // Find the custom properties part
    match find_custom_properties_part(package) {
        Ok(part) => {
            let xml_content = std::str::from_utf8(part.blob()).map_err(|e| {
                OoxmlError::Xml(format!("Invalid UTF-8 in custom properties: {}", e))
            })?;
            CustomProperties::from_xml(xml_content)
        },
        Err(_) => {
            // No custom properties part found, return empty collection
            Ok(CustomProperties::new())
        },
    }
}

/// Find the custom properties part in an OOXML package.
fn find_custom_properties_part(package: &OpcPackage) -> Result<&dyn crate::ooxml::opc::part::Part> {
    // Try the standard location first
    let standard_uri = PackURI::new("/docProps/custom.xml")
        .map_err(|e| OoxmlError::Other(format!("Invalid custom properties URI: {}", e)))?;

    if let Ok(part) = package.get_part(&standard_uri)
        && part.content_type() == ct::OFC_CUSTOM_PROPERTIES
    {
        return Ok(part);
    }

    // Fallback: search through all parts for custom properties content type
    for part in package.iter_parts() {
        if part.content_type() == ct::OFC_CUSTOM_PROPERTIES {
            return Ok(part);
        }
    }

    Err(OoxmlError::PartNotFound(
        "Custom properties part not found".to_string(),
    ))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_property_value_conversion() {
        let string_val = PropertyValue::String("test".to_string());
        assert_eq!(string_val.element_name(), "lpwstr");
        assert_eq!(string_val.to_xml_string(), "test");

        let int_val = PropertyValue::Integer(42);
        assert_eq!(int_val.element_name(), "i4");
        assert_eq!(int_val.to_xml_string(), "42");

        let bool_val = PropertyValue::Boolean(true);
        assert_eq!(bool_val.element_name(), "bool");
        assert_eq!(bool_val.to_xml_string(), "true");
    }

    #[test]
    fn test_custom_properties_add_get() {
        let mut props = CustomProperties::new();

        props.add_property("Name", PropertyValue::String("Test".to_string()));
        props.add_property("Version", PropertyValue::Integer(1));

        assert_eq!(props.len(), 2);
        assert!(props.contains("Name"));
        assert!(props.contains("Version"));

        match props.get_property("Name") {
            Some(PropertyValue::String(s)) => assert_eq!(s, "Test"),
            _ => panic!("Expected string value"),
        }

        match props.get_property("Version") {
            Some(PropertyValue::Integer(i)) => assert_eq!(*i, 1),
            _ => panic!("Expected integer value"),
        }
    }

    #[test]
    fn test_custom_properties_remove() {
        let mut props = CustomProperties::new();
        props.add_property("Test", PropertyValue::Integer(42));

        assert!(props.contains("Test"));

        let removed = props.remove_property("Test");
        assert!(removed.is_some());
        assert!(!props.contains("Test"));
        assert_eq!(props.len(), 0);
    }

    #[test]
    fn test_custom_properties_replace() {
        let mut props = CustomProperties::new();
        props.add_property("Value", PropertyValue::Integer(1));

        let old = props.add_property("Value", PropertyValue::Integer(2));
        assert!(matches!(old, Some(PropertyValue::Integer(1))));

        match props.get_property("Value") {
            Some(PropertyValue::Integer(i)) => assert_eq!(*i, 2),
            _ => panic!("Expected updated integer value"),
        }
    }

    #[test]
    fn test_custom_properties_xml_generation() {
        let mut props = CustomProperties::new();
        props.add_property(
            "StringProp",
            PropertyValue::String("test value".to_string()),
        );
        props.add_property("IntProp", PropertyValue::Integer(123));
        props.add_property("BoolProp", PropertyValue::Boolean(true));

        let xml = props.to_xml().unwrap();

        assert!(xml.contains(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#));
        assert!(xml.contains("Properties"));
        assert!(xml.contains(r#"name="StringProp""#));
        assert!(xml.contains(r#"name="IntProp""#));
        assert!(xml.contains(r#"name="BoolProp""#));
        assert!(xml.contains("test value"));
        assert!(xml.contains("123"));
        assert!(xml.contains("true"));
    }

    #[test]
    fn test_custom_properties_xml_parsing() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" 
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="TestString">
        <vt:lpwstr>Hello World</vt:lpwstr>
    </property>
    <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="TestInt">
        <vt:i4>42</vt:i4>
    </property>
    <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="TestBool">
        <vt:bool>true</vt:bool>
    </property>
    <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="TestDouble">
        <vt:r8>3.14159</vt:r8>
    </property>
</Properties>"#;

        let props = CustomProperties::from_xml(xml).unwrap();

        assert_eq!(props.len(), 4);
        assert!(props.contains("TestString"));
        assert!(props.contains("TestInt"));
        assert!(props.contains("TestBool"));
        assert!(props.contains("TestDouble"));

        match props.get_property("TestString") {
            Some(PropertyValue::String(s)) => assert_eq!(s, "Hello World"),
            _ => panic!("Expected string value"),
        }

        match props.get_property("TestInt") {
            Some(PropertyValue::Integer(i)) => assert_eq!(*i, 42),
            _ => panic!("Expected integer value"),
        }

        match props.get_property("TestBool") {
            Some(PropertyValue::Boolean(b)) => assert!(*b),
            _ => panic!("Expected boolean value"),
        }

        match props.get_property("TestDouble") {
            Some(PropertyValue::Double(d)) => assert!((*d - std::f64::consts::PI).abs() < 0.001),
            _ => panic!("Expected double value"),
        }
    }

    #[test]
    fn test_custom_properties_roundtrip() {
        let mut props = CustomProperties::new();
        props.add_property("Name", PropertyValue::String("John Doe".to_string()));
        props.add_property("Age", PropertyValue::Integer(30));
        props.add_property("Score", PropertyValue::Double(98.5));
        props.add_property("Active", PropertyValue::Boolean(true));

        let xml = props.to_xml().unwrap();
        let parsed = CustomProperties::from_xml(&xml).unwrap();

        assert_eq!(parsed.len(), 4);
        assert_eq!(
            parsed.get_property("Name"),
            Some(&PropertyValue::String("John Doe".to_string()))
        );
        assert_eq!(
            parsed.get_property("Age"),
            Some(&PropertyValue::Integer(30))
        );
        assert_eq!(
            parsed.get_property("Active"),
            Some(&PropertyValue::Boolean(true))
        );
    }

    #[test]
    fn test_property_names_iterator() {
        let mut props = CustomProperties::new();
        props.add_property("Prop1", PropertyValue::Integer(1));
        props.add_property("Prop2", PropertyValue::Integer(2));
        props.add_property("Prop3", PropertyValue::Integer(3));

        let names: Vec<_> = props.property_names().collect();
        assert_eq!(names.len(), 3);
        assert!(names.contains(&"Prop1"));
        assert!(names.contains(&"Prop2"));
        assert!(names.contains(&"Prop3"));
    }

    #[test]
    fn test_iter() {
        let mut props = CustomProperties::new();
        props.add_property("A", PropertyValue::Integer(1));
        props.add_property("B", PropertyValue::String("test".to_string()));

        let items: Vec<_> = props.iter().collect();
        assert_eq!(items.len(), 2);
    }

    #[test]
    fn test_clear() {
        let mut props = CustomProperties::new();
        props.add_property("Test1", PropertyValue::Integer(1));
        props.add_property("Test2", PropertyValue::Integer(2));

        assert_eq!(props.len(), 2);

        props.clear();

        assert_eq!(props.len(), 0);
        assert!(props.is_empty());
    }
}
