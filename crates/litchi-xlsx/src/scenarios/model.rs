//! Typed worksheet what-if scenario values.

use crate::error::{Result, invalid};

pub(crate) const TRANSITIONAL_MAIN: &str =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(crate) const STRICT_MAIN: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
pub(crate) const MAX_SCENARIOS: usize = 65_535;
pub(crate) const MAX_INPUT_CELLS: usize = 65_536;
pub(crate) const MAX_SQREF_ITEMS: usize = 32_767;
pub(crate) const MAX_XSTRING_CHARS: usize = 32_767;
pub(crate) const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
pub(crate) const MAX_DEPTH: usize = 256;
pub(crate) const MAX_EVENTS: usize = 1_000_000;
pub(crate) const MAX_ROW: u32 = 1_048_576;
pub(crate) const MAX_COLUMN: u32 = 16_384;
pub(crate) const MAX_UNKNOWN_ATTRIBUTES: usize = 65_536;
pub(crate) const MAX_UNKNOWN_ELEMENTS: usize = 65_536;
pub(crate) const MAX_UNKNOWN_BYTES: usize = MAX_XML_BYTES;

/// Namespace form used when serializing a scenarios fragment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    pub(crate) fn main_namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_MAIN,
            Self::Strict => STRICT_MAIN,
        }
    }
}

/// A validated A1 cell reference (ST_CellRef).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellReference(String);

impl CellReference {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_cell_reference(&value, "scenario cell reference")?;
        Ok(Self(value))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// A validated A1 cell or rectangular range reference from scenarios sqref.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RangeReference(String);

impl RangeReference {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_range_reference(&value, "scenarios sqref item")?;
        Ok(Self(value))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// A namespaced attribute that this owner does not interpret.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UnknownAttribute {
    qualified_name: Box<str>,
    value: Box<str>,
    pub(crate) namespace: Option<NamespaceBinding>,
}

impl UnknownAttribute {
    pub fn name(&self) -> &str {
        &self.qualified_name
    }

    pub fn value(&self) -> &str {
        &self.value
    }

    pub(crate) fn from_decoded(
        qualified_name: String,
        value: String,
        namespace: Option<NamespaceBinding>,
    ) -> Result<Self> {
        if qualified_name.is_empty() || qualified_name.len() > MAX_XML_BYTES {
            return Err(invalid("unknown scenario attribute name is out of bounds"));
        }
        if value.len() > MAX_XML_BYTES {
            return Err(invalid("unknown scenario attribute value is out of bounds"));
        }
        Ok(Self {
            qualified_name: qualified_name.into_boxed_str(),
            value: value.into_boxed_str(),
            namespace,
        })
    }
}

/// One self-contained element that this owner does not interpret.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UnknownElement {
    xml: Box<[u8]>,
    pub(crate) namespaces: Vec<NamespaceBinding>,
}

impl UnknownElement {
    /// Construct an inert, bounded element fragment.
    ///
    /// The fragment is emitted without interpreting its content. Parsed
    /// values are validated by the worksheet codec before reaching this model.
    pub fn new(xml: impl Into<Vec<u8>>) -> Result<Self> {
        Self::from_parts(xml.into(), Vec::new())
    }

    pub fn as_xml(&self) -> &[u8] {
        &self.xml
    }

    pub(crate) fn from_parts(xml: Vec<u8>, namespaces: Vec<NamespaceBinding>) -> Result<Self> {
        if xml.is_empty() || xml.len() > MAX_XML_BYTES {
            return Err(invalid("unknown scenario element is out of bounds"));
        }
        if namespaces.len() > MAX_UNKNOWN_ATTRIBUTES {
            return Err(invalid(
                "unknown scenario namespace bindings exceed safety limit",
            ));
        }
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(xml.len())
            .map_err(|_| invalid("unknown scenario element allocation failed"))?;
        owned.extend_from_slice(&xml);
        Ok(Self {
            xml: owned.into_boxed_slice(),
            namespaces,
        })
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct NamespaceBinding {
    pub(crate) prefix: Box<str>,
    pub(crate) uri: Box<str>,
}

impl NamespaceBinding {
    pub(crate) fn new(prefix: String, uri: String) -> Self {
        Self {
            prefix: prefix.into_boxed_str(),
            uri: uri.into_boxed_str(),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ChildOrder {
    Known(usize),
    Unknown(usize),
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub(crate) struct OpaqueFields {
    pub(crate) attributes: Vec<UnknownAttribute>,
    pub(crate) elements: Vec<UnknownElement>,
    pub(crate) order: Vec<ChildOrder>,
    pub(crate) retained_bytes: usize,
}

impl OpaqueFields {
    pub(crate) fn push_attribute(&mut self, value: UnknownAttribute) -> Result<()> {
        if self.attributes.len() >= MAX_UNKNOWN_ATTRIBUTES {
            return Err(invalid("unknown scenario attributes exceed safety limit"));
        }
        let size = value
            .name()
            .len()
            .checked_add(value.value().len())
            .and_then(|size| {
                value.namespace.as_ref().map_or(Some(size), |namespace| {
                    size.checked_add(namespace.prefix.len())
                        .and_then(|size| size.checked_add(namespace.uri.len()))
                })
            })
            .ok_or_else(|| invalid("unknown scenario attribute size overflow"))?;
        let retained = self
            .retained_bytes
            .checked_add(size)
            .ok_or_else(|| invalid("unknown scenario data size overflow"))?;
        if retained > MAX_UNKNOWN_BYTES {
            return Err(invalid("unknown scenario data exceeds safety limit"));
        }
        self.attributes
            .try_reserve(1)
            .map_err(|_| invalid("unknown scenario attribute allocation failed"))?;
        self.attributes.push(value);
        self.retained_bytes = retained;
        Ok(())
    }

    pub(crate) fn push_element(&mut self, value: UnknownElement) -> Result<usize> {
        if self.elements.len() >= MAX_UNKNOWN_ELEMENTS {
            return Err(invalid("unknown scenario elements exceed safety limit"));
        }
        let size = value.xml.len();
        let retained = self
            .retained_bytes
            .checked_add(size)
            .ok_or_else(|| invalid("unknown scenario data size overflow"))?;
        if retained > MAX_UNKNOWN_BYTES {
            return Err(invalid("unknown scenario data exceeds safety limit"));
        }
        self.elements
            .try_reserve(1)
            .map_err(|_| invalid("unknown scenario element allocation failed"))?;
        self.order
            .try_reserve(1)
            .map_err(|_| invalid("unknown scenario child-order allocation failed"))?;
        let index = self.elements.len();
        self.elements.push(value);
        self.order.push(ChildOrder::Unknown(index));
        self.retained_bytes = retained;
        Ok(index)
    }

    pub(crate) fn push_known(&mut self, index: usize) -> Result<()> {
        self.order
            .try_reserve(1)
            .map_err(|_| invalid("unknown scenario child-order allocation failed"))?;
        self.order.push(ChildOrder::Known(index));
        Ok(())
    }

    pub(crate) fn has_unknown(&self) -> bool {
        !self.attributes.is_empty() || !self.elements.is_empty()
    }
}

/// One substitute value assignment (CT_InputCells).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct InputCell {
    pub(crate) reference: CellReference,
    pub(crate) deleted: bool,
    pub(crate) undone: bool,
    pub(crate) value: String,
    pub(crate) number_format_id: Option<u32>,
    pub(crate) opaque: Option<Box<OpaqueFields>>,
}

impl InputCell {
    pub fn new(reference: CellReference, value: impl Into<String>) -> Result<Self> {
        let value = checked_xstring(value.into(), "inputCells val")?;
        Ok(Self {
            reference,
            deleted: false,
            undone: false,
            value,
            number_format_id: None,
            opaque: None,
        })
    }

    pub fn with_deleted(mut self, value: bool) -> Self {
        self.deleted = value;
        self
    }

    pub fn with_undone(mut self, value: bool) -> Self {
        self.undone = value;
        self
    }

    pub fn with_number_format_id(mut self, value: u32) -> Self {
        self.number_format_id = Some(value);
        self
    }

    pub fn reference(&self) -> &CellReference {
        &self.reference
    }

    pub fn deleted(&self) -> bool {
        self.deleted
    }

    pub fn undone(&self) -> bool {
        self.undone
    }

    pub fn value(&self) -> &str {
        &self.value
    }

    pub fn number_format_id(&self) -> Option<u32> {
        self.number_format_id
    }

    /// Unknown attributes retained on the inputCells element.
    pub fn unknown_attributes(&self) -> &[UnknownAttribute] {
        self.opaque
            .as_deref()
            .map_or(&[], |opaque| opaque.attributes.as_slice())
    }

    /// Unknown child elements retained inertly and in document order.
    pub fn unknown_elements(&self) -> &[UnknownElement] {
        self.opaque
            .as_deref()
            .map_or(&[], |opaque| opaque.elements.as_slice())
    }

    pub(crate) fn opaque_mut(&mut self) -> &mut OpaqueFields {
        self.opaque
            .get_or_insert_with(|| Box::new(OpaqueFields::default()))
            .as_mut()
    }
}

/// One named what-if scenario (CT_Scenario).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Scenario {
    pub(crate) name: String,
    pub(crate) locked: bool,
    pub(crate) hidden: bool,
    pub(crate) count: Option<u32>,
    pub(crate) user: Option<String>,
    pub(crate) comment: Option<String>,
    pub(crate) input_cells: Vec<InputCell>,
    pub(crate) opaque: Option<Box<OpaqueFields>>,
}

impl Scenario {
    pub fn new(name: impl Into<String>) -> Result<Self> {
        let name = checked_xstring(name.into(), "scenario name")?;
        Ok(Self {
            name,
            locked: false,
            hidden: false,
            count: None,
            user: None,
            comment: None,
            input_cells: Vec::new(),
            opaque: None,
        })
    }

    pub fn with_locked(mut self, value: bool) -> Self {
        self.locked = value;
        self
    }

    pub fn with_hidden(mut self, value: bool) -> Self {
        self.hidden = value;
        self
    }

    pub fn with_count(mut self, value: u32) -> Self {
        self.count = Some(value);
        self
    }

    pub fn with_user(mut self, value: impl Into<String>) -> Result<Self> {
        self.user = Some(checked_xstring(value.into(), "scenario user")?);
        Ok(self)
    }

    pub fn with_comment(mut self, value: impl Into<String>) -> Result<Self> {
        self.comment = Some(checked_xstring(value.into(), "scenario comment")?);
        Ok(self)
    }

    pub fn with_input_cells(mut self, value: Vec<InputCell>) -> Result<Self> {
        if value.len() > MAX_INPUT_CELLS {
            return Err(invalid(format!(
                "scenario inputCells exceeds safety limit {MAX_INPUT_CELLS}"
            )));
        }
        self.input_cells = value;
        if let Some(opaque) = self.opaque.as_mut() {
            opaque.order.clear();
            opaque
                .order
                .try_reserve(self.input_cells.len())
                .map_err(|_| invalid("scenario child-order allocation failed"))?;
            for index in 0..self.input_cells.len() {
                opaque.order.push(ChildOrder::Known(index));
            }
            for index in 0..opaque.elements.len() {
                opaque.order.push(ChildOrder::Unknown(index));
            }
        }
        Ok(self)
    }

    pub fn name(&self) -> &str {
        &self.name
    }

    pub fn locked(&self) -> bool {
        self.locked
    }

    pub fn hidden(&self) -> bool {
        self.hidden
    }

    pub fn count(&self) -> Option<u32> {
        self.count
    }

    pub fn user(&self) -> Option<&str> {
        self.user.as_deref()
    }

    pub fn comment(&self) -> Option<&str> {
        self.comment.as_deref()
    }

    pub fn input_cells(&self) -> &[InputCell] {
        &self.input_cells
    }

    /// Unknown attributes retained on the scenario element.
    pub fn unknown_attributes(&self) -> &[UnknownAttribute] {
        self.opaque
            .as_deref()
            .map_or(&[], |opaque| opaque.attributes.as_slice())
    }

    /// Unknown child elements retained inertly and in document order.
    pub fn unknown_elements(&self) -> &[UnknownElement] {
        self.opaque
            .as_deref()
            .map_or(&[], |opaque| opaque.elements.as_slice())
    }
}

/// The worksheet scenarios collection in document order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Collection {
    pub(crate) current: Option<u32>,
    pub(crate) show: Option<u32>,
    pub(crate) ranges: Vec<RangeReference>,
    pub(crate) scenarios: Vec<Scenario>,
    pub(crate) opaque: Option<Box<OpaqueFields>>,
}

impl Collection {
    pub fn new(scenarios: Vec<Scenario>) -> Result<Self> {
        if scenarios.is_empty() {
            return Err(invalid("scenarios requires at least one scenario"));
        }
        if scenarios.len() > MAX_SCENARIOS {
            return Err(invalid(format!(
                "scenarios exceeds safety limit {MAX_SCENARIOS}"
            )));
        }
        Ok(Self {
            current: None,
            show: None,
            ranges: Vec::new(),
            scenarios,
            opaque: None,
        })
    }

    pub fn with_current(mut self, value: u32) -> Self {
        self.current = Some(value);
        self
    }

    pub fn with_show(mut self, value: u32) -> Self {
        self.show = Some(value);
        self
    }

    pub fn with_ranges(mut self, value: Vec<RangeReference>) -> Result<Self> {
        if value.len() > MAX_SQREF_ITEMS {
            return Err(invalid(format!(
                "scenarios sqref exceeds safety limit {MAX_SQREF_ITEMS}"
            )));
        }
        self.ranges = value;
        Ok(self)
    }

    pub fn current(&self) -> Option<u32> {
        self.current
    }

    pub fn show(&self) -> Option<u32> {
        self.show
    }

    pub fn ranges(&self) -> &[RangeReference] {
        &self.ranges
    }

    pub fn scenarios(&self) -> &[Scenario] {
        &self.scenarios
    }

    /// Unknown attributes retained on the scenarios element.
    pub fn unknown_attributes(&self) -> &[UnknownAttribute] {
        self.opaque
            .as_deref()
            .map_or(&[], |opaque| opaque.attributes.as_slice())
    }

    /// Unknown child elements retained inertly and in document order.
    pub fn unknown_elements(&self) -> &[UnknownElement] {
        self.opaque
            .as_deref()
            .map_or(&[], |opaque| opaque.elements.as_slice())
    }
}

pub(crate) fn checked_xstring(value: String, name: &str) -> Result<String> {
    if value.chars().count() > MAX_XSTRING_CHARS {
        return Err(invalid(format!(
            "{name} exceeds {MAX_XSTRING_CHARS} characters"
        )));
    }
    Ok(value)
}

pub(crate) fn validate_range_reference(value: &str, name: &str) -> Result<()> {
    let mut parts = value.split(':');
    let first = parts.next().unwrap_or_default();
    let second = parts.next();
    if parts.next().is_some() || first.is_empty() || second.is_some_and(str::is_empty) {
        return Err(invalid(format!("invalid {name} '{value}'")));
    }
    validate_cell_reference(first, name)?;
    if let Some(second) = second {
        validate_cell_reference(second, name)?;
    }
    Ok(())
}

pub(crate) fn validate_cell_reference(value: &str, name: &str) -> Result<()> {
    let bytes = value.as_bytes();
    let mut index = usize::from(bytes.first() == Some(&b'$'));
    let column_start = index;
    while bytes.get(index).is_some_and(u8::is_ascii_alphabetic) {
        index += 1;
    }
    if index == column_start || index - column_start > 3 {
        return Err(invalid(format!("invalid {name} '{value}'")));
    }
    let mut column = 0u32;
    for byte in &bytes[column_start..index] {
        column = column * 26 + u32::from(byte.to_ascii_uppercase() - b'A' + 1);
    }
    if column == 0 || column > MAX_COLUMN {
        return Err(invalid(format!(
            "{name} column is out of range in '{value}'"
        )));
    }
    if bytes.get(index) == Some(&b'$') {
        index += 1;
    }
    let row_start = index;
    while bytes.get(index).is_some_and(u8::is_ascii_digit) {
        index += 1;
    }
    if index == row_start || index != bytes.len() {
        return Err(invalid(format!("invalid {name} '{value}'")));
    }
    let row = value[row_start..]
        .parse::<u32>()
        .map_err(|_| invalid(format!("invalid {name} row in '{value}'")))?;
    if row == 0 || row > MAX_ROW {
        return Err(invalid(format!("{name} row is out of range in '{value}'")));
    }
    Ok(())
}
