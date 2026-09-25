//! Semantic values for `x14:formControlPr`.

use std::fmt;
use std::mem::size_of;
use std::ops::Deref;
use std::sync::Arc;

use litchi_core::Reservation;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesRef, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use crate::source_payload::SourcePayload;

use super::{
    MAX_FORMULA_BYTES, MAX_ITEM_VALUE_BYTES, MAX_OPAQUE_BYTES, MAX_RETAINED_BYTES, MAX_XML_DEPTH,
    MAX_XML_EVENTS, Result, allocation, invalid,
};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

/// A schema enum value that keeps a bounded unknown lexical token on read.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum KnownOrUnknown<T> {
    /// A value admitted by the local schema enumeration.
    Known(T),
    /// A bounded token retained for lossless source-backed preservation.
    Unknown(String),
}

impl<T> KnownOrUnknown<T> {
    /// Return the known value, if the source token was admitted.
    #[must_use]
    pub const fn known(&self) -> Option<&T> {
        match self {
            Self::Known(value) => Some(value),
            Self::Unknown(_) => None,
        }
    }

    /// Return the retained unknown lexical token, if any.
    #[must_use]
    pub const fn unknown(&self) -> Option<&String> {
        match self {
            Self::Known(_) => None,
            Self::Unknown(value) => Some(value),
        }
    }
}

macro_rules! token_enum {
    ($(#[$meta:meta])* $name:ident { $($variant:ident => $wire:literal),+ $(,)? }) => {
        $(#[$meta])*
        #[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
        pub enum $name {
            $($variant),+
        }

        impl $name {
            pub(crate) fn parse_token(value: &str) -> Option<Self> {
                match value {
                    $($wire => Some(Self::$variant),)+
                    _ => None,
                }
            }

            pub(crate) const fn wire(self) -> &'static str {
                match self {
                    $(Self::$variant => $wire,)+
                }
            }
        }

        impl fmt::Display for $name {
            fn fmt(&self, output: &mut fmt::Formatter<'_>) -> fmt::Result {
                output.write_str((*self).wire())
            }
        }
    };
}

token_enum! {
    /// Form-control object type (`ST_ObjectType`).
    ObjectType {
        Button => "Button",
        CheckBox => "CheckBox",
        Drop => "Drop",
        GBox => "GBox",
        Label => "Label",
        List => "List",
        Radio => "Radio",
        Scroll => "Scroll",
        Spin => "Spin",
        EditBox => "EditBox",
        Dialog => "Dialog",
    }
}

token_enum! {
    /// Check-box/radio state (`ST_Checked`).
    Checked {
        Unchecked => "Unchecked",
        Checked => "Checked",
        Mixed => "Mixed",
    }
}

token_enum! {
    /// Drop-down style (`ST_DropStyle`).
    DropStyle {
        Combo => "combo",
        ComboEdit => "comboedit",
        Simple => "simple",
    }
}

token_enum! {
    /// List selection mode (`ST_SelType`).
    SelectionType {
        Single => "single",
        Multi => "multi",
        Extended => "extended",
    }
}

token_enum! {
    /// Edit-box validation mode (`ST_EditValidation`).
    EditValidation {
        Text => "text",
        Integer => "integer",
        Number => "number",
        Reference => "reference",
        Formula => "formula",
    }
}

token_enum! {
    /// Horizontal text alignment (`ST_TextHAlign`).
    TextHAlign {
        Left => "left",
        Center => "center",
        Right => "right",
        Justify => "justify",
        Distributed => "distributed",
    }
}

token_enum! {
    /// Vertical text alignment (`ST_TextVAlign`).
    TextVAlign {
        Top => "top",
        Center => "center",
        Bottom => "bottom",
        Justify => "justify",
        Distributed => "distributed",
    }
}

/// A source-bound `ST_Formula` lexical value.
///
/// The local schema represents this as a string.  The wrapper keeps that
/// lexical value inert and bounded; reference/formula interpretation belongs
/// to a later worksheet semantic owner.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct FormControlFormula {
    value: String,
    source_only: bool,
}

impl FormControlFormula {
    /// Construct a non-empty formula/reference lexical value.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_xml_text(&value, "form-control formula")?;
        if value.is_empty() {
            return Err(invalid("form-control formula cannot be empty"));
        }
        if value.len() > MAX_FORMULA_BYTES {
            return Err(super::limit(
                "form-control formula bytes",
                value.len(),
                MAX_FORMULA_BYTES,
            ));
        }
        validate_formula_authoring(&value)?;
        Ok(Self {
            value,
            source_only: false,
        })
    }

    /// Construct a formula from source lexical text without interpreting it.
    pub(crate) fn from_source(value: String) -> Result<Self> {
        validate_xml_text(&value, "form-control formula")?;
        if value.is_empty() {
            return Err(invalid("form-control formula cannot be empty"));
        }
        if value.len() > MAX_FORMULA_BYTES {
            return Err(super::limit(
                "form-control formula bytes",
                value.len(),
                MAX_FORMULA_BYTES,
            ));
        }
        Ok(Self {
            source_only: validate_formula_authoring(&value).is_err(),
            value,
        })
    }

    /// Return the original decoded lexical value.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.value
    }

    /// Consume the wrapper and return the lexical value.
    #[must_use]
    pub fn into_string(self) -> String {
        self.value
    }
}

impl AsRef<str> for FormControlFormula {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl fmt::Display for FormControlFormula {
    fn fmt(&self, output: &mut fmt::Formatter<'_>) -> fmt::Result {
        output.write_str(self.as_str())
    }
}

impl FormControlFormula {
    pub(crate) const fn source_only(&self) -> bool {
        self.source_only
    }

    pub(crate) fn mark_source_only(&mut self) {
        self.source_only = true;
    }
}

/// A bounded namespace binding retained for detached extension fragments.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct NamespaceBinding {
    prefix: Arc<str>,
    uri: Arc<str>,
}

impl NamespaceBinding {
    pub(crate) fn new(prefix: String, uri: String) -> Self {
        Self {
            prefix: Arc::from(prefix),
            uri: Arc::from(uri),
        }
    }

    /// Prefix, or the empty string for the default binding.
    #[must_use]
    pub fn prefix(&self) -> &str {
        &self.prefix
    }

    /// Namespace URI.
    #[must_use]
    pub fn uri(&self) -> &str {
        &self.uri
    }
}

/// Exact opaque XML retained for an admitted `extLst` subtree.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OpaqueXml {
    pub(crate) source: SourcePayload,
    pub(crate) range: std::ops::Range<usize>,
    pub(crate) namespaces: Arc<[NamespaceBinding]>,
    pub(crate) diagnostic: bool,
}

impl OpaqueXml {
    pub(crate) fn from_range(
        source: SourcePayload,
        range: std::ops::Range<usize>,
        namespaces: Arc<[NamespaceBinding]>,
    ) -> Result<Self> {
        if range.start > range.end || range.end > source.len() {
            return Err(invalid("opaque XML range lies outside the source part"));
        }
        Ok(Self {
            source,
            range,
            namespaces,
            diagnostic: false,
        })
    }

    pub(crate) fn from_range_with_diagnostic(
        source: SourcePayload,
        range: std::ops::Range<usize>,
        namespaces: Arc<[NamespaceBinding]>,
        diagnostic: bool,
    ) -> Result<Self> {
        if range.start > range.end || range.end > source.len() {
            return Err(invalid("opaque XML range lies outside the source part"));
        }
        Ok(Self {
            source,
            range,
            namespaces,
            diagnostic,
        })
    }

    /// Construct an owned opaque XML fragment for a detached model.
    pub fn new(xml: impl Into<Vec<u8>>) -> Result<Self> {
        let xml = xml.into();
        if xml.len() > MAX_OPAQUE_BYTES {
            return Err(super::limit(
                "opaque form-control bytes",
                xml.len(),
                MAX_OPAQUE_BYTES,
            ));
        }
        validate_text_bytes(&xml, "opaque XML")?;
        validate_opaque_fragment(&xml)?;
        let source = SourcePayload::Owned(Arc::new(xml));
        let end = source.len();
        Ok(Self {
            source,
            range: 0..end,
            namespaces: Arc::from([]),
            diagnostic: false,
        })
    }

    /// Borrow the exact opaque source bytes.
    #[must_use]
    pub fn xml(&self) -> &[u8] {
        &self.source.as_bytes()[self.range.clone()]
    }

    /// Return namespace bindings captured from the owning root.
    #[must_use]
    pub fn namespaces(&self) -> &[NamespaceBinding] {
        &self.namespaces
    }

    /// Whether the retained subtree had a namespace or extension-token
    /// diagnostic that prevents typed extension authoring.
    #[must_use]
    pub const fn has_diagnostic(&self) -> bool {
        self.diagnostic
    }

    pub(crate) fn byte_len(&self) -> usize {
        self.range.end - self.range.start
    }
}

/// One unknown root attribute retained as opaque lexical XML.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OpaqueAttribute {
    pub(crate) lexical: Arc<[u8]>,
    pub(crate) name: Arc<[u8]>,
}

impl OpaqueAttribute {
    pub(crate) fn from_lexical(name: &[u8], lexical: &[u8]) -> Result<Self> {
        let mut name_copy = Vec::new();
        name_copy
            .try_reserve_exact(name.len())
            .map_err(|source| allocation("form-control unknown attribute name", source))?;
        name_copy.extend_from_slice(name);
        let mut lexical_copy = Vec::new();
        lexical_copy
            .try_reserve_exact(lexical.len())
            .map_err(|source| allocation("form-control unknown attribute", source))?;
        lexical_copy.extend_from_slice(lexical);
        Ok(Self {
            lexical: Arc::from(lexical_copy),
            name: Arc::from(name_copy),
        })
    }

    /// Attribute qualified name as it appeared in the source.
    #[must_use]
    pub fn name(&self) -> &[u8] {
        &self.name
    }

    /// Exact attribute token bytes, excluding surrounding whitespace.
    #[must_use]
    pub fn lexical(&self) -> &[u8] {
        &self.lexical
    }
}

/// One list item (`x14:item/@val`).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Item {
    value: String,
    pub(crate) raw: Option<OpaqueXml>,
    pub(crate) opaque_attributes: bool,
    pub(crate) raw_value: Option<String>,
}

impl Item {
    /// Construct an item from its XML Schema `string` value.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_item_value(&value, MAX_ITEM_VALUE_BYTES)?;
        Ok(Self {
            value,
            raw: None,
            opaque_attributes: false,
            raw_value: None,
        })
    }

    /// Item value with XML entities decoded.
    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
    }

    /// Alias matching the wire attribute name.
    #[must_use]
    pub fn val(&self) -> &str {
        self.value()
    }

    /// Replace the item value.
    pub fn set_value(&mut self, value: impl Into<String>) -> Result<()> {
        let value = value.into();
        validate_item_value(&value, MAX_ITEM_VALUE_BYTES)?;
        self.value = value;
        Ok(())
    }

    /// Alias matching the wire attribute name.
    pub fn set_val(&mut self, value: impl Into<String>) -> Result<()> {
        self.set_value(value)
    }

    /// Whether a source item carries opaque attributes or child markup.
    #[must_use]
    pub const fn has_opaque_markup(&self) -> bool {
        self.opaque_attributes
    }
}

impl TryFrom<&str> for Item {
    type Error = super::FormControlError;

    fn try_from(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

/// Ordered `itemLst` contents and its optional opaque extension list.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ItemList {
    items: Vec<Item>,
    extension_list: Option<OpaqueXml>,
    namespaces: Arc<[NamespaceBinding]>,
    unknown_attributes: Vec<OpaqueAttribute>,
}

impl ItemList {
    /// Create a list from ordered items.
    pub fn new(values: impl IntoIterator<Item = Item>) -> Result<Self> {
        let mut items = Vec::new();
        let mut retained_payload = 0usize;
        for item in values {
            if items.len() >= super::MAX_ITEMS {
                return Err(super::limit(
                    "item count",
                    items.len() + 1,
                    super::MAX_ITEMS,
                ));
            }
            let item_payload = item
                .value
                .capacity()
                .checked_add(item.raw_value.as_ref().map_or(0, String::capacity))
                .and_then(|value| {
                    value.checked_add(item.raw.as_ref().map_or(0, OpaqueXml::byte_len))
                })
                .ok_or_else(|| invalid("form-control item retained bytes overflow"))?;
            retained_payload = retained_payload
                .checked_add(item_payload)
                .ok_or_else(|| invalid("form-control item retained bytes overflow"))?;
            let required = items
                .len()
                .checked_add(1)
                .ok_or_else(|| invalid("form-control item count overflow"))?;
            let target_capacity = if required > items.capacity() {
                required.max(
                    items
                        .capacity()
                        .max(1)
                        .checked_mul(2)
                        .ok_or_else(|| invalid("form-control item capacity overflow"))?,
                )
            } else {
                items.capacity()
            };
            let vector_bytes = target_capacity
                .checked_mul(size_of::<Item>())
                .ok_or_else(|| invalid("form-control item vector bytes overflow"))?;
            let retained = retained_payload
                .checked_add(vector_bytes)
                .ok_or_else(|| invalid("form-control item retained bytes overflow"))?;
            if retained > MAX_RETAINED_BYTES {
                return Err(super::limit(
                    "retained form-control item bytes",
                    retained,
                    MAX_RETAINED_BYTES,
                ));
            }
            items
                .try_reserve_exact(target_capacity.saturating_sub(items.capacity()))
                .map_err(|source| allocation("form-control items", source))?;
            items.push(item);
        }
        Ok(Self {
            items,
            extension_list: None,
            namespaces: Arc::from([]),
            unknown_attributes: Vec::new(),
        })
    }

    pub(crate) fn from_parts_with_namespace(
        items: Vec<Item>,
        extension_list: Option<OpaqueXml>,
        namespaces: Arc<[NamespaceBinding]>,
        unknown_attributes: Vec<OpaqueAttribute>,
    ) -> Self {
        Self {
            items,
            extension_list,
            namespaces,
            unknown_attributes,
        }
    }

    /// Ordered item values.
    #[must_use]
    pub fn items(&self) -> &[Item] {
        &self.items
    }

    /// Number of items.
    #[must_use]
    pub fn len(&self) -> usize {
        self.items.len()
    }

    /// Whether the list contains no items.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.items.is_empty()
    }

    /// Optional `itemLst/extLst` payload.
    #[must_use]
    pub const fn extension_list(&self) -> Option<&OpaqueXml> {
        self.extension_list.as_ref()
    }

    /// Set the opaque `itemLst/extLst` payload.
    pub fn set_extension_list(&mut self, value: Option<OpaqueXml>) {
        self.extension_list = value;
    }

    /// Namespace bindings that were in scope on the source `itemLst`.
    #[must_use]
    pub fn namespaces(&self) -> &[NamespaceBinding] {
        &self.namespaces
    }

    /// Unknown `itemLst` attributes retained as exact lexical tokens.
    #[must_use]
    pub fn unknown_attributes(&self) -> &[OpaqueAttribute] {
        &self.unknown_attributes
    }
}

/// Selector vocabulary reserved for the eventual worksheet owner.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum ControlSelector<'a> {
    /// Checked zero-based ordinal among effective worksheet controls.
    Position(usize),
    /// Exact control name.
    Name(&'a str),
}

impl<'a> ControlSelector<'a> {
    pub const fn position(value: usize) -> Self {
        Self::Position(value)
    }

    pub const fn name(value: &'a str) -> Self {
        Self::Name(value)
    }
}

/// A typed source-independent scalar field name.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ScalarField {
    ObjectType,
    Checked,
    Colored,
    DropLines,
    DropStyle,
    Dx,
    FirstButton,
    FmlaGroup,
    FmlaLink,
    FmlaRange,
    FmlaTxbx,
    Horiz,
    Inc,
    JustLastX,
    LockText,
    Max,
    Min,
    MultiSel,
    NoThreeD,
    NoThreeD2,
    Page,
    Sel,
    SelType,
    TextHAlign,
    TextVAlign,
    Val,
    WidthMin,
    EditVal,
    MultiLine,
    VerticalBar,
    PasswordEdit,
}

/// Typed value accepted by a scalar source splice.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum ScalarValue {
    ObjectType(ObjectType),
    Checked(Checked),
    DropStyle(DropStyle),
    SelectionType(SelectionType),
    TextHAlign(TextHAlign),
    TextVAlign(TextVAlign),
    EditValidation(EditValidation),
    Boolean(bool),
    Unsigned(u32),
    Formula(FormControlFormula),
    String(String),
}

/// A clone-cheap optional value used by the source-backed leaf model.
///
/// The value allocation is immutable and replaced as a whole.  This keeps a
/// cloned `Properties` snapshot from copying an existing string, item list, or
/// opaque value when a caller edits the clone; only the newly supplied value is
/// allocated.  The parser mutates its own handles while they are unique.
#[derive(Clone, Debug)]
pub(crate) struct SharedOption<T>(Option<Arc<T>>);

impl<T> SharedOption<T> {
    fn new(value: Option<T>) -> Self {
        Self(value.map(Arc::new))
    }

    pub(crate) fn as_ref(&self) -> Option<&T> {
        self.0.as_deref()
    }

    pub(crate) fn as_deref(&self) -> Option<&T::Target>
    where
        T: Deref,
    {
        self.0.as_deref().map(Deref::deref)
    }

    pub(crate) fn is_some(&self) -> bool {
        self.0.is_some()
    }

    pub(crate) fn cloned(&self) -> Option<T>
    where
        T: Clone,
    {
        self.as_ref().cloned()
    }

    pub(crate) fn replace(&mut self, value: Option<T>) {
        self.0 = value.map(Arc::new);
    }
}

impl<T: PartialEq> PartialEq<Option<T>> for SharedOption<T> {
    fn eq(&self, other: &Option<T>) -> bool {
        self.as_ref() == other.as_ref()
    }
}

/// A clone-cheap immutable vector handle used for parser-owned collections.
#[derive(Clone, Debug)]
pub(crate) struct SharedVec<T>(Arc<Vec<T>>);

impl<T> SharedVec<T> {
    fn new(value: Vec<T>) -> Self {
        Self(Arc::new(value))
    }

    pub(crate) fn as_slice(&self) -> &[T] {
        self.0.as_slice()
    }

    /// Replace the vector contents, reusing the existing unique `Arc` when
    /// possible.  The parser constructs its `Properties` value privately, so
    /// its backing handles are unique while it is assembling the model.  A
    /// shared handle can still be replaced by allocating a fresh backing; the
    /// caller can inspect `is_unique` first and precharge that allocation.
    pub(crate) fn replace(&mut self, value: Vec<T>) {
        if let Some(current) = Arc::get_mut(&mut self.0) {
            *current = value;
        } else {
            self.0 = Arc::new(value);
        }
    }

    pub(crate) fn is_unique(&self) -> bool {
        Arc::strong_count(&self.0) == 1
    }

    pub(crate) fn as_mut_vec(&mut self) -> &mut Vec<T>
    where
        T: Clone,
    {
        Arc::make_mut(&mut self.0)
    }
}

impl<T> Deref for SharedVec<T> {
    type Target = Vec<T>;

    fn deref(&self) -> &Self::Target {
        &self.0
    }
}

/// Detached form-control property model.
#[derive(Clone, Debug)]
pub struct Properties {
    pub(crate) object_type: SharedOption<KnownOrUnknown<ObjectType>>,
    pub(crate) checked: SharedOption<KnownOrUnknown<Checked>>,
    pub(crate) colored: Option<bool>,
    pub(crate) drop_lines: Option<u32>,
    pub(crate) drop_style: SharedOption<KnownOrUnknown<DropStyle>>,
    pub(crate) dx: Option<u32>,
    pub(crate) first_button: Option<bool>,
    pub(crate) fmla_group: SharedOption<FormControlFormula>,
    pub(crate) fmla_link: SharedOption<FormControlFormula>,
    pub(crate) fmla_range: SharedOption<FormControlFormula>,
    pub(crate) fmla_txbx: SharedOption<FormControlFormula>,
    pub(crate) horiz: Option<bool>,
    pub(crate) inc: Option<u32>,
    pub(crate) just_last_x: Option<bool>,
    pub(crate) lock_text: Option<bool>,
    pub(crate) max: Option<u32>,
    pub(crate) min: Option<u32>,
    pub(crate) multi_sel: SharedOption<String>,
    pub(crate) no_three_d: Option<bool>,
    pub(crate) no_three_d2: Option<bool>,
    pub(crate) page: Option<u32>,
    pub(crate) sel: Option<u32>,
    pub(crate) seltype: SharedOption<KnownOrUnknown<SelectionType>>,
    pub(crate) text_h_align: SharedOption<KnownOrUnknown<TextHAlign>>,
    pub(crate) text_v_align: SharedOption<KnownOrUnknown<TextVAlign>>,
    pub(crate) val: Option<u32>,
    pub(crate) width_min: Option<u32>,
    pub(crate) edit_val: SharedOption<KnownOrUnknown<EditValidation>>,
    pub(crate) multi_line: Option<bool>,
    pub(crate) vertical_bar: Option<bool>,
    pub(crate) password_edit: Option<bool>,
    pub(crate) item_list: SharedOption<ItemList>,
    pub(crate) root_extension_list: SharedOption<OpaqueXml>,
    pub(crate) unknown_attributes: SharedVec<OpaqueAttribute>,
    pub(crate) namespaces: Arc<[NamespaceBinding]>,
    pub(crate) lexical_attributes: SharedVec<LexicalAttribute>,
    pub(crate) source: Option<SourcePayload>,
    pub(crate) retained_lease: Option<Arc<Reservation>>,
    pub(crate) source_dirty: bool,
    pub(crate) source_diagnostics: bool,
}

/// Conservative fixed storage bound for one retained `Properties` value.
/// The individual shared value/vector allocations are charged when created;
/// this covers the model's fixed handles and its owner-side `Arc` header.
pub(crate) const PROPERTIES_STORAGE_BYTES: usize = size_of::<Properties>() + 2 * size_of::<usize>();

const SHARED_ARC_HEADER_BYTES: usize = 2 * size_of::<usize>();
pub(crate) const SHARED_VEC_STORAGE_BYTES: usize = SHARED_ARC_HEADER_BYTES + size_of::<Vec<()>>();
pub(crate) const RETAINED_LEASE_STORAGE_BYTES: usize =
    SHARED_ARC_HEADER_BYTES + size_of::<Reservation>();

/// One recognized attribute's original lexical token.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct LexicalAttribute {
    pub(crate) name: Arc<[u8]>,
    pub(crate) raw_value: Arc<[u8]>,
    pub(crate) decoded_value: String,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SemanticState {
    object_type: Option<KnownOrUnknown<ObjectType>>,
    checked: Option<KnownOrUnknown<Checked>>,
    colored: Option<bool>,
    drop_lines: Option<u32>,
    drop_style: Option<KnownOrUnknown<DropStyle>>,
    dx: Option<u32>,
    first_button: Option<bool>,
    fmla_group: Option<FormControlFormula>,
    fmla_link: Option<FormControlFormula>,
    fmla_range: Option<FormControlFormula>,
    fmla_txbx: Option<FormControlFormula>,
    horiz: Option<bool>,
    inc: Option<u32>,
    just_last_x: Option<bool>,
    lock_text: Option<bool>,
    max: Option<u32>,
    min: Option<u32>,
    multi_sel: Option<String>,
    no_three_d: Option<bool>,
    no_three_d2: Option<bool>,
    page: Option<u32>,
    sel: Option<u32>,
    seltype: Option<KnownOrUnknown<SelectionType>>,
    text_h_align: Option<KnownOrUnknown<TextHAlign>>,
    text_v_align: Option<KnownOrUnknown<TextVAlign>>,
    val: Option<u32>,
    width_min: Option<u32>,
    edit_val: Option<KnownOrUnknown<EditValidation>>,
    multi_line: Option<bool>,
    vertical_bar: Option<bool>,
    password_edit: Option<bool>,
    item_list: Option<ItemList>,
    root_extension_list: Option<OpaqueXml>,
    unknown_attributes: Vec<OpaqueAttribute>,
}

impl Default for Properties {
    fn default() -> Self {
        Self::new()
    }
}

impl PartialEq for Properties {
    fn eq(&self, other: &Self) -> bool {
        self.semantic_state() == other.semantic_state()
    }
}

impl Eq for Properties {}

impl Properties {
    /// Create an empty valid `formControlPr` model.
    #[must_use]
    pub fn new() -> Self {
        Self {
            object_type: SharedOption::new(None),
            checked: SharedOption::new(None),
            colored: None,
            drop_lines: None,
            drop_style: SharedOption::new(None),
            dx: None,
            first_button: None,
            fmla_group: SharedOption::new(None),
            fmla_link: SharedOption::new(None),
            fmla_range: SharedOption::new(None),
            fmla_txbx: SharedOption::new(None),
            horiz: None,
            inc: None,
            just_last_x: None,
            lock_text: None,
            max: None,
            min: None,
            multi_sel: SharedOption::new(None),
            no_three_d: None,
            no_three_d2: None,
            page: None,
            sel: None,
            seltype: SharedOption::new(None),
            text_h_align: SharedOption::new(None),
            text_v_align: SharedOption::new(None),
            val: None,
            width_min: None,
            edit_val: SharedOption::new(None),
            multi_line: None,
            vertical_bar: None,
            password_edit: None,
            item_list: SharedOption::new(None),
            root_extension_list: SharedOption::new(None),
            unknown_attributes: SharedVec::new(Vec::new()),
            namespaces: Arc::from([]),
            lexical_attributes: SharedVec::new(Vec::new()),
            source: None,
            retained_lease: None,
            source_dirty: false,
            source_diagnostics: false,
        }
    }

    /// Return the exact source bytes when this model came from a part.
    #[must_use]
    pub fn source_bytes(&self) -> Option<&[u8]> {
        self.source.as_ref().map(SourcePayload::as_bytes)
    }

    pub(crate) fn source_payload(&self) -> Option<SourcePayload> {
        self.source.clone()
    }

    /// Return the optional object type, including an unknown source token.
    #[must_use]
    pub fn object_type(&self) -> Option<&KnownOrUnknown<ObjectType>> {
        self.object_type.as_ref()
    }

    /// Return the optional checked state, including an unknown source token.
    #[must_use]
    pub fn checked(&self) -> Option<&KnownOrUnknown<Checked>> {
        self.checked.as_ref()
    }

    /// Return authored `colored` presence.
    #[must_use]
    pub const fn colored(&self) -> Option<bool> {
        self.colored
    }

    /// Effective `colored` value (`false` when omitted).
    #[must_use]
    pub fn effective_colored(&self) -> bool {
        self.colored.unwrap_or(false)
    }

    /// Return authored `dropLines` presence.
    #[must_use]
    pub const fn drop_lines(&self) -> Option<u32> {
        self.drop_lines
    }

    /// Effective `dropLines` value (`8` when omitted).
    #[must_use]
    pub fn effective_drop_lines(&self) -> u32 {
        self.drop_lines.unwrap_or(8)
    }

    /// Return the optional drop style.
    #[must_use]
    pub fn drop_style(&self) -> Option<&KnownOrUnknown<DropStyle>> {
        self.drop_style.as_ref()
    }

    /// Return authored `dx` presence.
    #[must_use]
    pub const fn dx(&self) -> Option<u32> {
        self.dx
    }

    /// Effective `dx` value (`80` when omitted).
    #[must_use]
    pub fn effective_dx(&self) -> u32 {
        self.dx.unwrap_or(80)
    }

    /// Return authored `firstButton` presence.
    #[must_use]
    pub const fn first_button(&self) -> Option<bool> {
        self.first_button
    }

    /// Effective `firstButton` value (`false` when omitted).
    #[must_use]
    pub fn effective_first_button(&self) -> bool {
        self.first_button.unwrap_or(false)
    }

    /// Return the formula linking a group box.
    #[must_use]
    pub fn fmla_group(&self) -> Option<&FormControlFormula> {
        self.fmla_group.as_ref()
    }

    /// Return the formula linking a control.
    #[must_use]
    pub fn fmla_link(&self) -> Option<&FormControlFormula> {
        self.fmla_link.as_ref()
    }

    /// Return the list/drop-down source range formula.
    #[must_use]
    pub fn fmla_range(&self) -> Option<&FormControlFormula> {
        self.fmla_range.as_ref()
    }

    /// Return the label/edit-box source formula.
    #[must_use]
    pub fn fmla_txbx(&self) -> Option<&FormControlFormula> {
        self.fmla_txbx.as_ref()
    }

    /// Return authored `horiz` presence.
    #[must_use]
    pub const fn horiz(&self) -> Option<bool> {
        self.horiz
    }

    /// Effective `horiz` value (`false` when omitted).
    #[must_use]
    pub fn effective_horiz(&self) -> bool {
        self.horiz.unwrap_or(false)
    }

    /// Return authored `inc` presence.
    #[must_use]
    pub const fn inc(&self) -> Option<u32> {
        self.inc
    }

    /// Effective `inc` value (`1` when omitted).
    #[must_use]
    pub fn effective_inc(&self) -> u32 {
        self.inc.unwrap_or(1)
    }

    /// Return authored `justLastX` presence.
    #[must_use]
    pub const fn just_last_x(&self) -> Option<bool> {
        self.just_last_x
    }

    /// Effective `justLastX` value (`false` when omitted).
    #[must_use]
    pub fn effective_just_last_x(&self) -> bool {
        self.just_last_x.unwrap_or(false)
    }

    /// Return authored `lockText` presence.
    #[must_use]
    pub const fn lock_text(&self) -> Option<bool> {
        self.lock_text
    }

    /// Effective `lockText` value (`false` when omitted).
    #[must_use]
    pub fn effective_lock_text(&self) -> bool {
        self.lock_text.unwrap_or(false)
    }

    /// Return authored `max` presence.
    #[must_use]
    pub const fn max(&self) -> Option<u32> {
        self.max
    }

    /// Return authored `min` presence.
    #[must_use]
    pub const fn min(&self) -> Option<u32> {
        self.min
    }

    /// Effective `min` value (`0` when omitted).
    #[must_use]
    pub fn effective_min(&self) -> u32 {
        self.min.unwrap_or(0)
    }

    /// Return authored `multiSel` presence.
    #[must_use]
    pub fn multi_sel(&self) -> Option<&str> {
        self.multi_sel.as_deref()
    }

    /// Return authored `noThreeD` presence.
    #[must_use]
    pub const fn no_three_d(&self) -> Option<bool> {
        self.no_three_d
    }

    /// Effective `noThreeD` value (`false` when omitted).
    #[must_use]
    pub fn effective_no_three_d(&self) -> bool {
        self.no_three_d.unwrap_or(false)
    }

    /// Return authored `noThreeD2` presence.
    #[must_use]
    pub const fn no_three_d2(&self) -> Option<bool> {
        self.no_three_d2
    }

    /// Effective `noThreeD2` value (`false` when omitted).
    #[must_use]
    pub fn effective_no_three_d2(&self) -> bool {
        self.no_three_d2.unwrap_or(false)
    }

    /// Return authored `page` presence.
    #[must_use]
    pub const fn page(&self) -> Option<u32> {
        self.page
    }

    /// Return authored selected index presence.
    #[must_use]
    pub const fn sel(&self) -> Option<u32> {
        self.sel
    }

    /// Return the optional selection type.
    #[must_use]
    pub fn seltype(&self) -> Option<&KnownOrUnknown<SelectionType>> {
        self.seltype.as_ref()
    }

    /// Effective selection type (`single` when omitted).
    #[must_use]
    pub fn effective_seltype(&self) -> KnownOrUnknown<SelectionType> {
        self.seltype
            .cloned()
            .unwrap_or(KnownOrUnknown::Known(SelectionType::Single))
    }

    /// Alias using a word-separated spelling of the wire name.
    #[must_use]
    pub fn effective_sel_type(&self) -> KnownOrUnknown<SelectionType> {
        self.effective_seltype()
    }

    /// Return the optional horizontal alignment.
    #[must_use]
    pub fn text_h_align(&self) -> Option<&KnownOrUnknown<TextHAlign>> {
        self.text_h_align.as_ref()
    }

    /// Effective horizontal alignment (`left` when omitted).
    #[must_use]
    pub fn effective_text_h_align(&self) -> KnownOrUnknown<TextHAlign> {
        self.text_h_align
            .cloned()
            .unwrap_or(KnownOrUnknown::Known(TextHAlign::Left))
    }

    /// Return the optional vertical alignment.
    #[must_use]
    pub fn text_v_align(&self) -> Option<&KnownOrUnknown<TextVAlign>> {
        self.text_v_align.as_ref()
    }

    /// Effective vertical alignment (`top` when omitted).
    #[must_use]
    pub fn effective_text_v_align(&self) -> KnownOrUnknown<TextVAlign> {
        self.text_v_align
            .cloned()
            .unwrap_or(KnownOrUnknown::Known(TextVAlign::Top))
    }

    /// Return authored `val` presence.
    #[must_use]
    pub const fn val(&self) -> Option<u32> {
        self.val
    }

    /// Effective `val` value (`0` when omitted).
    #[must_use]
    pub fn effective_val(&self) -> u32 {
        self.val.unwrap_or(0)
    }

    /// Return authored `widthMin` presence.
    #[must_use]
    pub const fn width_min(&self) -> Option<u32> {
        self.width_min
    }

    /// Return the optional edit validation mode.
    #[must_use]
    pub fn edit_val(&self) -> Option<&KnownOrUnknown<EditValidation>> {
        self.edit_val.as_ref()
    }

    /// Effective edit validation (`text` when omitted).
    #[must_use]
    pub fn effective_edit_val(&self) -> KnownOrUnknown<EditValidation> {
        self.edit_val
            .cloned()
            .unwrap_or(KnownOrUnknown::Known(EditValidation::Text))
    }

    /// Return authored `multiLine` presence.
    #[must_use]
    pub const fn multi_line(&self) -> Option<bool> {
        self.multi_line
    }

    /// Effective `multiLine` value (`false` when omitted).
    #[must_use]
    pub fn effective_multi_line(&self) -> bool {
        self.multi_line.unwrap_or(false)
    }

    /// Return authored `verticalBar` presence.
    #[must_use]
    pub const fn vertical_bar(&self) -> Option<bool> {
        self.vertical_bar
    }

    /// Effective `verticalBar` value (`false` when omitted).
    #[must_use]
    pub fn effective_vertical_bar(&self) -> bool {
        self.vertical_bar.unwrap_or(false)
    }

    /// Return authored `passwordEdit` presence.
    #[must_use]
    pub const fn password_edit(&self) -> Option<bool> {
        self.password_edit
    }

    /// Effective `passwordEdit` value (`false` when omitted).
    #[must_use]
    pub fn effective_password_edit(&self) -> bool {
        self.password_edit.unwrap_or(false)
    }

    /// Return the optional ordered item list.
    #[must_use]
    pub fn item_list(&self) -> Option<&ItemList> {
        self.item_list.as_ref()
    }

    /// Return the root opaque extension list.
    #[must_use]
    pub fn root_extension_list(&self) -> Option<&OpaqueXml> {
        self.root_extension_list.as_ref()
    }

    /// Return unknown root attributes.
    #[must_use]
    pub fn unknown_attributes(&self) -> &[OpaqueAttribute] {
        self.unknown_attributes.as_slice()
    }

    /// Return the root namespace context retained for detached writing.
    #[must_use]
    pub fn namespaces(&self) -> &[NamespaceBinding] {
        &self.namespaces
    }

    /// Set an optional known object type.
    pub fn set_object_type(&mut self, value: Option<ObjectType>) {
        let value = value.map(KnownOrUnknown::Known);
        if self.object_type.as_ref() != value.as_ref() {
            self.mark_changed();
            self.object_type.replace(value);
        }
    }

    /// Set an optional known checked state.
    pub fn set_checked(&mut self, value: Option<Checked>) {
        let value = value.map(KnownOrUnknown::Known);
        if self.checked.as_ref() != value.as_ref() {
            self.mark_changed();
            self.checked.replace(value);
        }
    }

    pub fn set_colored(&mut self, value: Option<bool>) {
        if self.colored != value {
            self.mark_changed();
            self.colored = value;
        }
    }

    pub fn set_drop_lines(&mut self, value: Option<u32>) -> Result<()> {
        check_bounded("dropLines", value)?;
        if self.drop_lines != value {
            self.mark_changed();
            self.drop_lines = value;
        }
        Ok(())
    }

    pub fn set_drop_style(&mut self, value: Option<DropStyle>) {
        let value = value.map(KnownOrUnknown::Known);
        if self.drop_style.as_ref() != value.as_ref() {
            self.mark_changed();
            self.drop_style.replace(value);
        }
    }

    pub fn set_dx(&mut self, value: Option<u32>) {
        if self.dx != value {
            self.mark_changed();
            self.dx = value;
        }
    }

    pub fn set_first_button(&mut self, value: Option<bool>) {
        if self.first_button != value {
            self.mark_changed();
            self.first_button = value;
        }
    }

    pub fn set_fmla_group(&mut self, value: Option<FormControlFormula>) {
        if self.fmla_group.as_ref() != value.as_ref() {
            self.mark_changed();
            self.fmla_group.replace(value);
        }
    }

    pub fn set_fmla_link(&mut self, value: Option<FormControlFormula>) {
        if self.fmla_link.as_ref() != value.as_ref() {
            self.mark_changed();
            self.fmla_link.replace(value);
        }
    }

    pub fn set_fmla_range(&mut self, value: Option<FormControlFormula>) {
        if self.fmla_range.as_ref() != value.as_ref() {
            self.mark_changed();
            self.fmla_range.replace(value);
        }
    }

    pub fn set_fmla_txbx(&mut self, value: Option<FormControlFormula>) {
        if self.fmla_txbx.as_ref() != value.as_ref() {
            self.mark_changed();
            self.fmla_txbx.replace(value);
        }
    }

    pub fn set_horiz(&mut self, value: Option<bool>) {
        if self.horiz != value {
            self.mark_changed();
            self.horiz = value;
        }
    }

    pub fn set_inc(&mut self, value: Option<u32>) -> Result<()> {
        check_bounded("inc", value)?;
        if self.inc != value {
            self.mark_changed();
            self.inc = value;
        }
        Ok(())
    }

    pub fn set_just_last_x(&mut self, value: Option<bool>) {
        if self.just_last_x != value {
            self.mark_changed();
            self.just_last_x = value;
        }
    }

    pub fn set_lock_text(&mut self, value: Option<bool>) {
        if self.lock_text != value {
            self.mark_changed();
            self.lock_text = value;
        }
    }

    pub fn set_max(&mut self, value: Option<u32>) -> Result<()> {
        check_bounded("max", value)?;
        if self.max != value {
            self.mark_changed();
            self.max = value;
        }
        Ok(())
    }

    pub fn set_min(&mut self, value: Option<u32>) -> Result<()> {
        check_bounded("min", value)?;
        if self.min != value {
            self.mark_changed();
            self.min = value;
        }
        Ok(())
    }

    pub fn set_multi_sel(&mut self, value: Option<String>) -> Result<()> {
        if let Some(value) = value.as_deref() {
            if value.len() > MAX_OPAQUE_BYTES {
                return Err(super::limit(
                    "multiSel bytes",
                    value.len(),
                    MAX_OPAQUE_BYTES,
                ));
            }
            validate_multi_selection(value)?;
        }
        if self.multi_sel.as_ref() != value.as_ref() {
            self.mark_changed();
            self.multi_sel.replace(value);
        }
        Ok(())
    }

    pub fn set_no_three_d(&mut self, value: Option<bool>) {
        if self.no_three_d != value {
            self.mark_changed();
            self.no_three_d = value;
        }
    }

    pub fn set_no_three_d2(&mut self, value: Option<bool>) {
        if self.no_three_d2 != value {
            self.mark_changed();
            self.no_three_d2 = value;
        }
    }

    pub fn set_page(&mut self, value: Option<u32>) -> Result<()> {
        check_bounded("page", value)?;
        if self.page != value {
            self.mark_changed();
            self.page = value;
        }
        Ok(())
    }

    pub fn set_sel(&mut self, value: Option<u32>) {
        if self.sel != value {
            self.mark_changed();
            self.sel = value;
        }
    }

    pub fn set_seltype(&mut self, value: Option<SelectionType>) {
        let value = value.map(KnownOrUnknown::Known);
        if self.seltype.as_ref() != value.as_ref() {
            self.mark_changed();
            self.seltype.replace(value);
        }
    }

    pub fn set_text_h_align(&mut self, value: Option<TextHAlign>) {
        let value = value.map(KnownOrUnknown::Known);
        if self.text_h_align.as_ref() != value.as_ref() {
            self.mark_changed();
            self.text_h_align.replace(value);
        }
    }

    pub fn set_text_v_align(&mut self, value: Option<TextVAlign>) {
        let value = value.map(KnownOrUnknown::Known);
        if self.text_v_align.as_ref() != value.as_ref() {
            self.mark_changed();
            self.text_v_align.replace(value);
        }
    }

    pub fn set_val(&mut self, value: Option<u32>) {
        if self.val != value {
            self.mark_changed();
            self.val = value;
        }
    }

    pub fn set_width_min(&mut self, value: Option<u32>) {
        if self.width_min != value {
            self.mark_changed();
            self.width_min = value;
        }
    }

    pub fn set_edit_val(&mut self, value: Option<EditValidation>) {
        let value = value.map(KnownOrUnknown::Known);
        if self.edit_val.as_ref() != value.as_ref() {
            self.mark_changed();
            self.edit_val.replace(value);
        }
    }

    pub fn set_multi_line(&mut self, value: Option<bool>) {
        if self.multi_line != value {
            self.mark_changed();
            self.multi_line = value;
        }
    }

    pub fn set_vertical_bar(&mut self, value: Option<bool>) {
        if self.vertical_bar != value {
            self.mark_changed();
            self.vertical_bar = value;
        }
    }

    pub fn set_password_edit(&mut self, value: Option<bool>) {
        if self.password_edit != value {
            self.mark_changed();
            self.password_edit = value;
        }
    }

    /// Replace the optional ordered item list.
    pub fn set_item_list(&mut self, value: Option<ItemList>) -> Result<()> {
        if let Some(value) = value.as_ref() {
            check_item_list(value, super::Limits::default())?;
        }
        if self.item_list.as_ref() != value.as_ref() {
            self.mark_changed();
            self.item_list.replace(value);
        }
        Ok(())
    }

    /// Set the root opaque extension list.
    pub fn set_root_extension_list(&mut self, value: Option<OpaqueXml>) {
        if self.root_extension_list.as_ref() != value.as_ref() {
            self.mark_changed();
            self.root_extension_list.replace(value);
        }
    }

    /// Return a typed scalar snapshot for source-backed editing.
    pub fn scalar(&self, field: ScalarField) -> Option<ScalarValue> {
        match field {
            ScalarField::ObjectType => self.object_type.as_ref().and_then(|v| match v {
                KnownOrUnknown::Known(v) => Some(ScalarValue::ObjectType(*v)),
                KnownOrUnknown::Unknown(_) => None,
            }),
            ScalarField::Checked => self.checked.as_ref().and_then(|v| match v {
                KnownOrUnknown::Known(v) => Some(ScalarValue::Checked(*v)),
                KnownOrUnknown::Unknown(_) => None,
            }),
            ScalarField::Colored => self.colored.map(ScalarValue::Boolean),
            ScalarField::DropLines => self.drop_lines.map(ScalarValue::Unsigned),
            ScalarField::DropStyle => self.drop_style.as_ref().and_then(|v| match v {
                KnownOrUnknown::Known(v) => Some(ScalarValue::DropStyle(*v)),
                KnownOrUnknown::Unknown(_) => None,
            }),
            ScalarField::Dx => self.dx.map(ScalarValue::Unsigned),
            ScalarField::FirstButton => self.first_button.map(ScalarValue::Boolean),
            ScalarField::FmlaGroup => self.fmla_group.cloned().map(ScalarValue::Formula),
            ScalarField::FmlaLink => self.fmla_link.cloned().map(ScalarValue::Formula),
            ScalarField::FmlaRange => self.fmla_range.cloned().map(ScalarValue::Formula),
            ScalarField::FmlaTxbx => self.fmla_txbx.cloned().map(ScalarValue::Formula),
            ScalarField::Horiz => self.horiz.map(ScalarValue::Boolean),
            ScalarField::Inc => self.inc.map(ScalarValue::Unsigned),
            ScalarField::JustLastX => self.just_last_x.map(ScalarValue::Boolean),
            ScalarField::LockText => self.lock_text.map(ScalarValue::Boolean),
            ScalarField::Max => self.max.map(ScalarValue::Unsigned),
            ScalarField::Min => self.min.map(ScalarValue::Unsigned),
            ScalarField::MultiSel => self.multi_sel.cloned().map(ScalarValue::String),
            ScalarField::NoThreeD => self.no_three_d.map(ScalarValue::Boolean),
            ScalarField::NoThreeD2 => self.no_three_d2.map(ScalarValue::Boolean),
            ScalarField::Page => self.page.map(ScalarValue::Unsigned),
            ScalarField::Sel => self.sel.map(ScalarValue::Unsigned),
            ScalarField::SelType => self.seltype.as_ref().and_then(|v| match v {
                KnownOrUnknown::Known(v) => Some(ScalarValue::SelectionType(*v)),
                KnownOrUnknown::Unknown(_) => None,
            }),
            ScalarField::TextHAlign => self.text_h_align.as_ref().and_then(|v| match v {
                KnownOrUnknown::Known(v) => Some(ScalarValue::TextHAlign(*v)),
                KnownOrUnknown::Unknown(_) => None,
            }),
            ScalarField::TextVAlign => self.text_v_align.as_ref().and_then(|v| match v {
                KnownOrUnknown::Known(v) => Some(ScalarValue::TextVAlign(*v)),
                KnownOrUnknown::Unknown(_) => None,
            }),
            ScalarField::Val => self.val.map(ScalarValue::Unsigned),
            ScalarField::WidthMin => self.width_min.map(ScalarValue::Unsigned),
            ScalarField::EditVal => self.edit_val.as_ref().and_then(|v| match v {
                KnownOrUnknown::Known(v) => Some(ScalarValue::EditValidation(*v)),
                KnownOrUnknown::Unknown(_) => None,
            }),
            ScalarField::MultiLine => self.multi_line.map(ScalarValue::Boolean),
            ScalarField::VerticalBar => self.vertical_bar.map(ScalarValue::Boolean),
            ScalarField::PasswordEdit => self.password_edit.map(ScalarValue::Boolean),
        }
    }

    pub(crate) fn semantic_state(&self) -> SemanticState {
        SemanticState {
            object_type: self.object_type.cloned(),
            checked: self.checked.cloned(),
            colored: self.colored,
            drop_lines: self.drop_lines,
            drop_style: self.drop_style.cloned(),
            dx: self.dx,
            first_button: self.first_button,
            fmla_group: self.fmla_group.cloned(),
            fmla_link: self.fmla_link.cloned(),
            fmla_range: self.fmla_range.cloned(),
            fmla_txbx: self.fmla_txbx.cloned(),
            horiz: self.horiz,
            inc: self.inc,
            just_last_x: self.just_last_x,
            lock_text: self.lock_text,
            max: self.max,
            min: self.min,
            multi_sel: self.multi_sel.cloned(),
            no_three_d: self.no_three_d,
            no_three_d2: self.no_three_d2,
            page: self.page,
            sel: self.sel,
            seltype: self.seltype.cloned(),
            text_h_align: self.text_h_align.cloned(),
            text_v_align: self.text_v_align.cloned(),
            val: self.val,
            width_min: self.width_min,
            edit_val: self.edit_val.cloned(),
            multi_line: self.multi_line,
            vertical_bar: self.vertical_bar,
            password_edit: self.password_edit,
            item_list: self.item_list.cloned(),
            root_extension_list: self.root_extension_list.cloned(),
            unknown_attributes: self.unknown_attributes.as_slice().to_vec(),
        }
    }

    pub(crate) fn source_unchanged(&self) -> bool {
        self.source.is_some() && !self.source_dirty
    }

    pub(crate) fn mark_source(
        &mut self,
        source: SourcePayload,
        namespaces: Arc<[NamespaceBinding]>,
    ) {
        self.namespaces = namespaces;
        self.source = Some(source);
        self.source_dirty = false;
        self.source_diagnostics = true;
    }

    pub(crate) fn mark_metadata(&mut self, namespaces: Arc<[NamespaceBinding]>) {
        self.namespaces = namespaces;
        self.source_dirty = false;
        self.source_diagnostics = true;
    }

    pub(crate) fn mark_retained_lease(&mut self, lease: Option<Arc<Reservation>>) {
        self.retained_lease = lease;
    }

    pub(crate) fn mark_changed(&mut self) {
        if self.source.is_some() {
            self.source_dirty = true;
        }
    }

    pub(crate) fn validate_with_limits(&self, limits: super::Limits) -> Result<()> {
        self.validate_with_limits_mode(limits, true)
    }

    pub(crate) fn validate_parse_with_limits(&self, limits: super::Limits) -> Result<()> {
        self.validate_with_limits_mode(limits, false)
    }

    fn validate_with_limits_mode(
        &self,
        limits: super::Limits,
        check_applicability: bool,
    ) -> Result<()> {
        self.validate_retained_with_limits(limits)?;
        check_bounded("dropLines", self.drop_lines)?;
        check_bounded("inc", self.inc)?;
        check_bounded("max", self.max)?;
        check_bounded("min", self.min)?;
        check_bounded("page", self.page)?;
        check_item_list_opt(self.item_list.as_ref(), limits)?;
        if let Some(value) = self.multi_sel.as_deref() {
            if value.len() > limits.max_opaque_bytes() {
                return Err(super::limit(
                    "multiSel bytes",
                    value.len(),
                    limits.max_opaque_bytes(),
                ));
            }
            validate_multi_selection(value)?;
        }
        if let (Some(min), Some(max)) = (self.min, self.max) {
            if min > max {
                return Err(invalid("form-control min exceeds max"));
            }
        }
        let mut opaque = 0usize;
        for attribute in self.unknown_attributes.as_slice() {
            opaque = opaque
                .checked_add(attribute.lexical.len())
                .ok_or_else(|| invalid("unknown form-control attribute bytes overflow"))?;
        }
        if let Some(value) = self.root_extension_list.as_ref() {
            opaque = opaque
                .checked_add(value.byte_len())
                .ok_or_else(|| invalid("root extension bytes overflow"))?;
        }
        if let Some(item_list) = self.item_list.as_ref() {
            for attribute in &item_list.unknown_attributes {
                opaque = opaque
                    .checked_add(attribute.lexical.len())
                    .ok_or_else(|| invalid("item-list attribute bytes overflow"))?;
            }
            for item in &item_list.items {
                opaque = opaque
                    .checked_add(item.raw.as_ref().map_or(0, OpaqueXml::byte_len))
                    .ok_or_else(|| invalid("opaque item bytes overflow"))?;
            }
            if let Some(value) = item_list.extension_list.as_ref() {
                opaque = opaque
                    .checked_add(value.byte_len())
                    .ok_or_else(|| invalid("item extension bytes overflow"))?;
            }
        }
        if opaque > limits.max_opaque_bytes() {
            return Err(super::limit(
                "opaque form-control bytes",
                opaque,
                limits.max_opaque_bytes(),
            ));
        }
        if !self.source_diagnostics
            && [
                self.fmla_group.as_ref(),
                self.fmla_link.as_ref(),
                self.fmla_range.as_ref(),
                self.fmla_txbx.as_ref(),
            ]
            .into_iter()
            .flatten()
            .any(FormControlFormula::source_only)
        {
            return Err(invalid(
                "source-only form-control formula cannot be authored",
            ));
        }
        for (field, formula) in [
            (ScalarField::FmlaGroup, self.fmla_group.as_ref()),
            (ScalarField::FmlaLink, self.fmla_link.as_ref()),
            (ScalarField::FmlaRange, self.fmla_range.as_ref()),
            (ScalarField::FmlaTxbx, self.fmla_txbx.as_ref()),
        ]
        .into_iter()
        .filter_map(|(field, formula)| formula.map(|formula| (field, formula)))
        {
            if !formula.source_only() {
                validate_formula_for_field(field, formula.as_str())?;
            }
        }
        if check_applicability {
            if let Some(object_type) = self.object_type.as_ref().and_then(KnownOrUnknown::known) {
                validate_applicability(self, *object_type)?;
            }
        }
        Ok(())
    }

    pub(crate) fn validate_retained_with_limits(&self, limits: super::Limits) -> Result<()> {
        let mut retained = 0usize;
        add_retained(&mut retained, PROPERTIES_STORAGE_BYTES, limits)?;
        add_retained(&mut retained, 2 * SHARED_VEC_STORAGE_BYTES, limits)?;
        add_shared_option_storage(&mut retained, &self.object_type, limits)?;
        add_shared_option_storage(&mut retained, &self.checked, limits)?;
        add_shared_option_storage(&mut retained, &self.drop_style, limits)?;
        add_shared_option_storage(&mut retained, &self.fmla_group, limits)?;
        add_shared_option_storage(&mut retained, &self.fmla_link, limits)?;
        add_shared_option_storage(&mut retained, &self.fmla_range, limits)?;
        add_shared_option_storage(&mut retained, &self.fmla_txbx, limits)?;
        add_shared_option_storage(&mut retained, &self.multi_sel, limits)?;
        add_shared_option_storage(&mut retained, &self.seltype, limits)?;
        add_shared_option_storage(&mut retained, &self.text_h_align, limits)?;
        add_shared_option_storage(&mut retained, &self.text_v_align, limits)?;
        add_shared_option_storage(&mut retained, &self.edit_val, limits)?;
        add_shared_option_storage(&mut retained, &self.item_list, limits)?;
        add_shared_option_storage(&mut retained, &self.root_extension_list, limits)?;
        if self.retained_lease.is_some() {
            add_retained(&mut retained, RETAINED_LEASE_STORAGE_BYTES, limits)?;
        }
        add_retained(
            &mut retained,
            self.source
                .as_ref()
                .map_or(0, SourcePayload::retained_storage_bytes),
            limits,
        )?;
        // The root namespace Arc header is covered by the fixed Properties
        // storage charge; distinct child contexts carry their own headers.
        add_namespace_arc_storage(&mut retained, &self.namespaces, limits, false)?;
        for binding in &*self.namespaces {
            add_retained(&mut retained, binding.prefix.len(), limits)?;
            add_retained(&mut retained, binding.uri.len(), limits)?;
        }
        for attribute in self.lexical_attributes.as_slice() {
            add_retained(&mut retained, attribute.name.len(), limits)?;
            add_retained(&mut retained, attribute.raw_value.len(), limits)?;
            add_retained(&mut retained, attribute.decoded_value.len(), limits)?;
        }
        for attribute in self.unknown_attributes.as_slice() {
            add_retained(&mut retained, attribute.name.len(), limits)?;
            add_retained(&mut retained, attribute.lexical.len(), limits)?;
        }
        add_unknown_enum_retained(&mut retained, self.object_type.as_ref(), limits)?;
        add_unknown_enum_retained(&mut retained, self.checked.as_ref(), limits)?;
        add_unknown_enum_retained(&mut retained, self.drop_style.as_ref(), limits)?;
        add_unknown_enum_retained(&mut retained, self.seltype.as_ref(), limits)?;
        add_unknown_enum_retained(&mut retained, self.text_h_align.as_ref(), limits)?;
        add_unknown_enum_retained(&mut retained, self.text_v_align.as_ref(), limits)?;
        add_unknown_enum_retained(&mut retained, self.edit_val.as_ref(), limits)?;
        for formula in [
            self.fmla_group.as_ref(),
            self.fmla_link.as_ref(),
            self.fmla_range.as_ref(),
            self.fmla_txbx.as_ref(),
        ]
        .into_iter()
        .flatten()
        {
            add_retained(&mut retained, formula.value.len(), limits)?;
        }
        if let Some(value) = self.multi_sel.as_ref() {
            add_retained(&mut retained, value.len(), limits)?;
        }
        let source = self.source.as_ref();
        if let Some(item_list) = self.item_list.as_ref() {
            for attribute in &item_list.unknown_attributes {
                add_retained(&mut retained, attribute.name.len(), limits)?;
                add_retained(&mut retained, attribute.lexical.len(), limits)?;
            }
            add_namespace_context(
                &mut retained,
                item_list.namespaces(),
                &self.namespaces,
                limits,
            )?;
            for item in &item_list.items {
                add_retained(&mut retained, item.value.len(), limits)?;
                if let Some(value) = item.raw_value.as_ref() {
                    add_retained(&mut retained, value.len(), limits)?;
                }
                if let Some(raw) = item.raw.as_ref() {
                    add_opaque_retained(&mut retained, raw, source, limits)?;
                    add_namespace_context(
                        &mut retained,
                        raw.namespaces(),
                        item_list.namespaces(),
                        limits,
                    )?;
                }
            }
            if let Some(value) = item_list.extension_list.as_ref() {
                add_opaque_retained(&mut retained, value, source, limits)?;
                add_namespace_context(
                    &mut retained,
                    value.namespaces(),
                    item_list.namespaces(),
                    limits,
                )?;
            }
        }
        if let Some(value) = self.root_extension_list.as_ref() {
            add_opaque_retained(&mut retained, value, source, limits)?;
            add_namespace_context(&mut retained, value.namespaces(), &self.namespaces, limits)?;
        }
        // The parser charges its source-backed vector growth while it is
        // reserving each layout/model slot.  A detached model, or a parsed
        // model after a caller mutation, can instead have freshly grown
        // vectors that the parser never saw; include those live capacities
        // before a bounded write is allowed to proceed.  Unchanged owned
        // parses keep the parser's one-pass ledger, which also preserves the
        // exact source-plus-namespace boundary for a no-op parse.
        if self.source.is_none() || self.source_dirty {
            for value in [
                self.fmla_group.as_ref(),
                self.fmla_link.as_ref(),
                self.fmla_range.as_ref(),
                self.fmla_txbx.as_ref(),
            ]
            .into_iter()
            .flatten()
            {
                add_string_capacity(&mut retained, &value.value, limits)?;
            }
            if let Some(value) = self.multi_sel.as_ref() {
                add_string_capacity(&mut retained, value, limits)?;
            }
            for value in self
                .object_type
                .as_ref()
                .and_then(KnownOrUnknown::unknown)
                .into_iter()
                .chain(self.checked.as_ref().and_then(KnownOrUnknown::unknown))
                .chain(self.drop_style.as_ref().and_then(KnownOrUnknown::unknown))
                .chain(self.seltype.as_ref().and_then(KnownOrUnknown::unknown))
                .chain(self.text_h_align.as_ref().and_then(KnownOrUnknown::unknown))
                .chain(self.text_v_align.as_ref().and_then(KnownOrUnknown::unknown))
                .chain(self.edit_val.as_ref().and_then(KnownOrUnknown::unknown))
            {
                add_string_capacity(&mut retained, value, limits)?;
            }
            add_retained_capacity(
                &mut retained,
                self.unknown_attributes.capacity(),
                size_of::<OpaqueAttribute>(),
                limits,
            )?;
            add_retained_capacity(
                &mut retained,
                self.lexical_attributes.capacity(),
                size_of::<LexicalAttribute>(),
                limits,
            )?;
            if let Some(item_list) = self.item_list.as_ref() {
                for item in &item_list.items {
                    add_string_capacity(&mut retained, &item.value, limits)?;
                    if let Some(value) = item.raw_value.as_ref() {
                        add_string_capacity(&mut retained, value, limits)?;
                    }
                }
                add_retained_capacity(
                    &mut retained,
                    item_list.items.capacity(),
                    size_of::<Item>(),
                    limits,
                )?;
                add_retained_capacity(
                    &mut retained,
                    item_list.unknown_attributes.capacity(),
                    size_of::<OpaqueAttribute>(),
                    limits,
                )?;
            }
        }
        Ok(())
    }

    pub(crate) fn validate_scalar_edit(
        &self,
        field: ScalarField,
        value: Option<&ScalarValue>,
        limits: super::Limits,
    ) -> Result<()> {
        if field == ScalarField::MultiSel {
            if let Some(ScalarValue::String(value)) = value {
                if value.len() > limits.max_opaque_bytes() {
                    return Err(super::limit(
                        "multiSel bytes",
                        value.len(),
                        limits.max_opaque_bytes(),
                    ));
                }
            }
        }
        let object_type = if field == ScalarField::ObjectType {
            value.and_then(|value| match value {
                ScalarValue::ObjectType(value) => Some(*value),
                _ => None,
            })
        } else {
            self.object_type
                .as_ref()
                .and_then(KnownOrUnknown::known)
                .copied()
        };
        let Some(object_type) = object_type else {
            return Ok(());
        };
        for candidate in scalar_applicability_fields() {
            let present = if *candidate == field {
                value.is_some()
            } else {
                self.scalar_present(*candidate)
            };
            if present && !scalar_applies(*candidate, object_type) {
                return Err(invalid(format!(
                    "form-control attribute {candidate:?} is inapplicable to {object_type}"
                )));
            }
        }
        let checked_mixed = if field == ScalarField::Checked {
            matches!(value, Some(ScalarValue::Checked(Checked::Mixed)))
        } else {
            matches!(
                self.checked.as_ref(),
                Some(KnownOrUnknown::Known(Checked::Mixed))
            )
        };
        if checked_mixed && object_type != ObjectType::CheckBox {
            return Err(invalid("checked=Mixed applies only to CheckBox controls"));
        }
        let multi_sel_present = if field == ScalarField::MultiSel {
            value.is_some()
        } else {
            self.multi_sel.is_some()
        };
        let multi_sel_type = if field == ScalarField::SelType {
            matches!(
                value,
                Some(ScalarValue::SelectionType(SelectionType::Multi))
            )
        } else {
            matches!(
                self.seltype.as_ref(),
                Some(KnownOrUnknown::Known(SelectionType::Multi))
            )
        };
        if multi_sel_present && !multi_sel_type {
            return Err(invalid("multiSel requires seltype=multi"));
        }
        let sel = if field == ScalarField::Sel {
            value.and_then(|value| match value {
                ScalarValue::Unsigned(value) => Some(*value),
                _ => None,
            })
        } else {
            self.sel
        };
        if let (Some(sel), Some(item_list)) = (sel, self.item_list.as_ref()) {
            if sel != 0 && sel as usize > item_list.items.len() {
                return Err(invalid("sel exceeds the authored item count"));
            }
        }
        let min = if field == ScalarField::Min {
            value.and_then(|value| match value {
                ScalarValue::Unsigned(value) => Some(*value),
                _ => None,
            })
        } else {
            self.min
        };
        let max = if field == ScalarField::Max {
            value.and_then(|value| match value {
                ScalarValue::Unsigned(value) => Some(*value),
                _ => None,
            })
        } else {
            self.max
        };
        if let (Some(min), Some(max)) = (min, max) {
            if min > max {
                return Err(invalid("form-control min exceeds max"));
            }
        }
        Ok(())
    }

    fn scalar_present(&self, field: ScalarField) -> bool {
        match field {
            ScalarField::Checked => self.checked.is_some(),
            ScalarField::Colored => self.colored.is_some(),
            ScalarField::DropLines => self.drop_lines.is_some(),
            ScalarField::DropStyle => self.drop_style.is_some(),
            ScalarField::Dx => self.dx.is_some(),
            ScalarField::FirstButton => self.first_button.is_some(),
            ScalarField::FmlaGroup => self.fmla_group.is_some(),
            ScalarField::FmlaLink => self.fmla_link.is_some(),
            ScalarField::FmlaRange => self.fmla_range.is_some(),
            ScalarField::FmlaTxbx => self.fmla_txbx.is_some(),
            ScalarField::Horiz => self.horiz.is_some(),
            ScalarField::Inc => self.inc.is_some(),
            ScalarField::LockText => self.lock_text.is_some(),
            ScalarField::Max => self.max.is_some(),
            ScalarField::Min => self.min.is_some(),
            ScalarField::MultiSel => self.multi_sel.is_some(),
            ScalarField::NoThreeD => self.no_three_d.is_some(),
            ScalarField::NoThreeD2 => self.no_three_d2.is_some(),
            ScalarField::Page => self.page.is_some(),
            ScalarField::Sel => self.sel.is_some(),
            ScalarField::SelType => self.seltype.is_some(),
            ScalarField::Val => self.val.is_some(),
            ScalarField::WidthMin => self.width_min.is_some(),
            ScalarField::EditVal => self.edit_val.is_some(),
            ScalarField::MultiLine => self.multi_line.is_some(),
            ScalarField::VerticalBar => self.vertical_bar.is_some(),
            ScalarField::PasswordEdit => self.password_edit.is_some(),
            ScalarField::ObjectType
            | ScalarField::JustLastX
            | ScalarField::TextHAlign
            | ScalarField::TextVAlign => false,
        }
    }
}

/// Draft placeholder for the future worksheet graph owner.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct FormControlDraft {
    /// Optional semantic control name.
    pub name: Option<String>,
    /// Typed form-control properties.
    pub properties: Properties,
}

/// Read-only leaf view placeholder for the future worksheet owner.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FormControl<'a> {
    /// Selector used by the eventual worksheet owner.
    pub selector: ControlSelector<'a>,
    /// Optional source control name.
    pub name: Option<&'a str>,
    /// Typed properties payload.
    pub properties: Properties,
}

fn check_bounded(name: &'static str, value: Option<u32>) -> Result<()> {
    if let Some(value) = value {
        if value > 30_000 {
            return Err(super::limit(name, value as usize, 30_000));
        }
    }
    Ok(())
}

fn add_retained(total: &mut usize, amount: usize, limits: super::Limits) -> Result<()> {
    *total = total
        .checked_add(amount)
        .ok_or_else(|| invalid("form-control retained bytes overflow"))?;
    if *total > limits.max_retained_bytes() {
        return Err(super::limit(
            "retained form-control bytes",
            *total,
            limits.max_retained_bytes(),
        ));
    }
    Ok(())
}

fn add_shared_option_storage<T>(
    total: &mut usize,
    value: &SharedOption<T>,
    limits: super::Limits,
) -> Result<()> {
    if value.is_some() {
        add_retained(
            total,
            SHARED_ARC_HEADER_BYTES
                .checked_add(size_of::<T>())
                .ok_or_else(|| invalid("form-control shared option storage overflow"))?,
            limits,
        )?;
    }
    Ok(())
}

fn add_retained_capacity(
    total: &mut usize,
    capacity: usize,
    element_size: usize,
    limits: super::Limits,
) -> Result<()> {
    let bytes = capacity
        .checked_mul(element_size)
        .ok_or_else(|| invalid("form-control retained capacity overflow"))?;
    add_retained(total, bytes, limits)
}

fn add_string_capacity(total: &mut usize, value: &String, limits: super::Limits) -> Result<()> {
    add_retained(total, value.capacity().saturating_sub(value.len()), limits)
}

fn add_namespace_arc_storage(
    total: &mut usize,
    context: &[NamespaceBinding],
    limits: super::Limits,
    include_header: bool,
) -> Result<()> {
    if context.is_empty() {
        return Ok(());
    }
    let payload = context
        .len()
        .checked_mul(size_of::<NamespaceBinding>())
        .ok_or_else(|| invalid("form-control namespace Arc storage overflow"))?;
    let bytes = payload
        .checked_add(if include_header {
            SHARED_ARC_HEADER_BYTES
        } else {
            0
        })
        .ok_or_else(|| invalid("form-control namespace Arc storage overflow"))?;
    add_retained(total, bytes, limits)
}

fn add_unknown_enum_retained<T>(
    total: &mut usize,
    value: Option<&KnownOrUnknown<T>>,
    limits: super::Limits,
) -> Result<()> {
    if let Some(KnownOrUnknown::Unknown(value)) = value {
        add_retained(total, value.len(), limits)?;
    }
    Ok(())
}

fn add_namespace_context(
    total: &mut usize,
    context: &[NamespaceBinding],
    root: &[NamespaceBinding],
    limits: super::Limits,
) -> Result<()> {
    if !context.is_empty() && context.as_ptr() != root.as_ptr() {
        add_namespace_arc_storage(total, context, limits, true)?;
    }
    for binding in context {
        if root.iter().any(|value| {
            Arc::ptr_eq(&value.prefix, &binding.prefix) && Arc::ptr_eq(&value.uri, &binding.uri)
        }) {
            continue;
        }
        add_retained(total, binding.prefix.len(), limits)?;
        add_retained(total, binding.uri.len(), limits)?;
    }
    Ok(())
}

fn add_opaque_retained(
    total: &mut usize,
    value: &OpaqueXml,
    source: Option<&SourcePayload>,
    limits: super::Limits,
) -> Result<()> {
    if value.byte_len() > limits.max_opaque_bytes() {
        return Err(super::limit(
            "opaque form-control bytes",
            value.byte_len(),
            limits.max_opaque_bytes(),
        ));
    }
    // `OpaqueXml` retains a clone of the same `SourcePayload` as the model.
    // SourcePayload intentionally keeps managed PartData opaque, so compare
    // the borrowed slice identity instead of attempting an Arc escape.  The
    // full source slice is used for both values, making this allocation test
    // independent of the retained opaque range.
    let shares_source = source
        .is_some_and(|source| std::ptr::eq::<[u8]>(source.as_bytes(), value.source.as_bytes()));
    if !shares_source {
        add_retained(total, value.byte_len(), limits)?;
    }
    Ok(())
}

fn check_item_list_opt(value: Option<&ItemList>, limits: super::Limits) -> Result<()> {
    if let Some(value) = value {
        check_item_list(value, limits)?;
    }
    Ok(())
}

fn check_item_list(value: &ItemList, limits: super::Limits) -> Result<()> {
    if value.items.len() > limits.max_items() {
        return Err(super::limit(
            "item count",
            value.items.len(),
            limits.max_items(),
        ));
    }
    let mut opaque = 0usize;
    for item in &value.items {
        validate_item_value(&item.value, limits.max_item_value_bytes())?;
        opaque = opaque
            .checked_add(item.raw.as_ref().map_or(0, OpaqueXml::byte_len))
            .ok_or_else(|| invalid("opaque item bytes overflow"))?;
    }
    if let Some(value) = value.extension_list.as_ref() {
        opaque = opaque
            .checked_add(value.byte_len())
            .ok_or_else(|| invalid("opaque item extension bytes overflow"))?;
    }
    if opaque > limits.max_opaque_bytes() {
        return Err(super::limit(
            "opaque item/list bytes",
            opaque,
            limits.max_opaque_bytes(),
        ));
    }
    Ok(())
}

fn validate_item_value(value: &str, maximum: usize) -> Result<()> {
    if value.len() > maximum {
        return Err(super::limit("item value bytes", value.len(), maximum));
    }
    validate_xml_text(value, "item value")
}

pub(crate) fn validate_xml_text(value: &str, what: &'static str) -> Result<()> {
    validate_text_bytes(value.as_bytes(), what)
}

fn validate_opaque_fragment(xml: &[u8]) -> Result<()> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut events = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("opaque XML event count overflow"))?;
        if events > MAX_XML_EVENTS {
            return Err(super::limit(
                "opaque XML event count",
                events,
                MAX_XML_EVENTS,
            ));
        }
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("invalid opaque XML: {error}")))?;
        let resolver = reader.resolver();
        let decoder = reader.decoder();
        match event {
            Event::Start(element) => {
                if root_closed || (depth == 0 && root_seen) {
                    return Err(invalid("opaque XML fragment has multiple roots"));
                }
                validate_opaque_attributes(&element, resolver, decoder)?;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("opaque XML depth overflow"))?;
                if depth > MAX_XML_DEPTH {
                    return Err(super::limit("opaque XML depth", depth, MAX_XML_DEPTH));
                }
                if depth == 1 {
                    root_seen = true;
                }
            },
            Event::Empty(element) => {
                if root_closed || (depth == 0 && root_seen) {
                    return Err(invalid("opaque XML fragment has multiple roots"));
                }
                validate_opaque_attributes(&element, resolver, decoder)?;
                let element_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("opaque XML depth overflow"))?;
                if element_depth > MAX_XML_DEPTH {
                    return Err(super::limit(
                        "opaque XML depth",
                        element_depth,
                        MAX_XML_DEPTH,
                    ));
                }
                if depth == 0 {
                    root_seen = true;
                    root_closed = true;
                }
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid("opaque XML fragment has an unexpected end"));
                }
                depth -= 1;
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Text(text) => {
                if depth == 0 && !is_xml_whitespace(text.as_ref()) {
                    return Err(invalid("opaque XML fragment has top-level text"));
                }
            },
            Event::CData(_) => {
                if depth == 0 {
                    return Err(invalid("opaque XML fragment has top-level CDATA"));
                }
            },
            Event::GeneralRef(reference) => {
                if depth == 0 {
                    return Err(invalid("opaque XML fragment has a top-level entity"));
                }
                validate_opaque_reference(&reference)?;
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::Decl(_) => return Err(invalid("opaque XML declarations are not admitted")),
            Event::DocType(_) => return Err(invalid("opaque XML DOCTYPE is not admitted")),
            Event::Eof => break,
        }
    }
    if !root_seen || !root_closed || depth != 0 {
        return Err(invalid("opaque XML fragment has no complete root"));
    }
    Ok(())
}

fn validate_opaque_attributes(
    element: &quick_xml::events::BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
) -> Result<()> {
    let attributes = element.checked_attributes();
    let mut count = 0usize;
    for attribute in attributes {
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid("opaque XML attribute count overflow"))?;
        if count > super::MAX_ATTRIBUTES {
            return Err(super::limit(
                "opaque XML attribute count",
                count,
                super::MAX_ATTRIBUTES,
            ));
        }
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid opaque XML attribute: {error}")))?;
        attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| invalid(format!("invalid opaque XML attribute value: {error}")))?;
        if attribute.key.as_namespace_binding().is_none()
            && attribute.key.prefix().is_some()
            && !matches!(
                resolver.resolve_attribute(attribute.key).0,
                ResolveResult::Bound(_)
            )
        {
            return Err(invalid(
                "opaque XML attribute uses an unbound namespace prefix",
            ));
        }
    }
    let name = element.name();
    if name.prefix().is_some()
        && !matches!(resolver.resolve_element(name).0, ResolveResult::Bound(_))
    {
        return Err(invalid(
            "opaque XML element uses an unbound namespace prefix",
        ));
    }
    Ok(())
}

fn validate_opaque_reference(reference: &BytesRef<'_>) -> Result<()> {
    let value = reference
        .decode()
        .map_err(|error| invalid(format!("invalid opaque XML entity: {error}")))?;
    let value = value.as_ref();
    if let Some(value) = value
        .strip_prefix("#x")
        .or_else(|| value.strip_prefix("#X"))
    {
        let codepoint = u32::from_str_radix(value, 16)
            .map_err(|_| invalid("invalid opaque XML hexadecimal character reference"))?;
        let character = char::from_u32(codepoint)
            .ok_or_else(|| invalid("opaque XML character reference is outside Unicode"))?;
        if !is_xml_10_char(character) {
            return Err(invalid("opaque XML character reference is XML-invalid"));
        }
    } else if let Some(value) = value.strip_prefix('#') {
        let codepoint = value
            .parse::<u32>()
            .map_err(|_| invalid("invalid opaque XML decimal character reference"))?;
        let character = char::from_u32(codepoint)
            .ok_or_else(|| invalid("opaque XML character reference is outside Unicode"))?;
        if !is_xml_10_char(character) {
            return Err(invalid("opaque XML character reference is XML-invalid"));
        }
    } else if !matches!(value, "amp" | "lt" | "gt" | "apos" | "quot") {
        return Err(invalid("undeclared opaque XML entity reference"));
    }
    Ok(())
}

fn is_xml_whitespace(value: &[u8]) -> bool {
    value
        .iter()
        .all(|character| matches!(character, b' ' | b'\t' | b'\r' | b'\n'))
}

fn validate_text_bytes(value: &[u8], what: &'static str) -> Result<()> {
    let value = std::str::from_utf8(value)
        .map_err(|error| invalid(format!("{what} is not UTF-8: {error}")))?;
    for character in value.chars() {
        if !is_xml_10_char(character) {
            return Err(invalid(format!("{what} contains an XML-invalid character")));
        }
    }
    Ok(())
}

pub(crate) const fn is_xml_10_char(character: char) -> bool {
    matches!(
        character,
        '\u{9}'
            | '\u{a}'
            | '\u{d}'
            | '\u{20}'..='\u{d7ff}'
            | '\u{e000}'..='\u{fffd}'
            | '\u{10000}'..='\u{10ffff}'
    )
}

pub(crate) fn validate_multi_selection(value: &str) -> Result<()> {
    if value.is_empty() {
        return Ok(());
    }
    for token in value.split(',') {
        let token = token.trim_matches([' ', '\t', '\r', '\n']);
        let index = token
            .parse::<u32>()
            .map_err(|_| invalid("multiSel must be a comma-delimited one-based index list"))?;
        if index == 0 {
            return Err(invalid("multiSel indices are one-based"));
        }
    }
    Ok(())
}

fn validate_formula_authoring(value: &str) -> Result<()> {
    if value.starts_with('#') {
        return Err(invalid("form-control formula error tokens are source-only"));
    }
    validate_reference_characters(value)?;
    validate_reference_shape(value)
}

pub(crate) fn validate_formula_for_field(field: ScalarField, value: &str) -> Result<()> {
    if !matches!(field, ScalarField::FmlaGroup | ScalarField::FmlaLink) {
        return Ok(());
    }
    let body = reference_bang(value)
        .map(|index| &value[index + 1..])
        .unwrap_or(value);
    if let Some(relative) = top_level_delimiter(body, b',') {
        return Err(invalid(format!(
            "{field:?} formula cannot contain a union: byte {}",
            relative
        )));
    }
    if top_level_delimiter(body, b':').is_some() {
        return Err(invalid(format!(
            "{field:?} formula must be a single cell reference"
        )));
    }
    Ok(())
}

fn validate_reference_characters(value: &str) -> Result<()> {
    let mut quoted = false;
    let mut brackets = 0usize;
    let mut characters = value.chars().peekable();
    while let Some(character) = characters.next() {
        if quoted {
            if character == '\'' {
                if characters.next_if_eq(&'\'').is_none() {
                    quoted = false;
                }
            } else if character == '\u{0}' || character == '\u{7f}' || character.is_control() {
                return Err(invalid(
                    "form-control formula contains invalid reference syntax",
                ));
            }
            continue;
        }
        match character {
            '\'' => quoted = true,
            '[' => {
                brackets = brackets
                    .checked_add(1)
                    .ok_or_else(|| invalid("form-control formula qualifier depth overflow"))?;
            },
            ']' => {
                if brackets == 0 {
                    return Err(invalid("form-control formula has an unmatched ]"));
                }
                brackets -= 1;
            },
            character if brackets != 0 => {
                if character == '\u{0}' || character == '\u{7f}' || character.is_control() {
                    return Err(invalid(
                        "form-control formula contains invalid reference syntax",
                    ));
                }
            },
            '$' | ':' | '!' | ',' | '.' | '_' | '\\' | '-' | '+' => {},
            character if character.is_ascii_alphanumeric() || character.is_alphabetic() => {},
            _ => {
                return Err(invalid(
                    "form-control formula contains invalid reference syntax",
                ));
            },
        }
    }
    if quoted || brackets != 0 {
        return Err(invalid(
            "form-control formula has an unbalanced quoted sheet name",
        ));
    }
    Ok(())
}

fn validate_reference_shape(value: &str) -> Result<()> {
    let mut area_start = 0usize;
    while let Some(relative) = top_level_delimiter(&value[area_start..], b',') {
        let index = area_start
            .checked_add(relative)
            .ok_or_else(|| invalid("form-control formula area offset overflow"))?;
        // A comma inside a quoted sheet name or an external-book qualifier is
        // part of that qualifier.  `top_level_delimiter` only reports a
        // separator outside both forms.
        validate_reference_area(&value[area_start..index])?;
        area_start = index
            .checked_add(1)
            .ok_or_else(|| invalid("form-control formula area offset overflow"))?;
    }
    validate_reference_area(&value[area_start..])?;
    Ok(())
}

fn reference_bang(value: &str) -> Option<usize> {
    top_level_delimiter(value, b'!')
}

fn validate_reference_area(value: &str) -> Result<()> {
    if value.is_empty() {
        return Err(invalid("form-control formula has an empty reference area"));
    }
    let body = if let Some(index) = reference_bang(value) {
        let qualifier = &value[..index];
        if qualifier.is_empty() {
            return Err(invalid("form-control formula has an empty sheet qualifier"));
        }
        validate_reference_qualifier(qualifier)?;
        &value[index + 1..]
    } else {
        value
    };
    if body.is_empty() {
        return Err(invalid("form-control formula has an empty cell reference"));
    }
    if let Some(index) = top_level_delimiter(body, b':') {
        if top_level_delimiter(&body[index + 1..], b':').is_some() {
            return Err(invalid(
                "form-control formula has too many range separators",
            ));
        }
        validate_reference_atom(&body[..index])?;
        validate_reference_atom(&body[index + 1..])?;
    } else {
        validate_reference_atom(body)?;
    }
    Ok(())
}

fn top_level_delimiter(value: &str, delimiter: u8) -> Option<usize> {
    let bytes = value.as_bytes();
    let mut index = 0usize;
    let mut quoted = false;
    let mut brackets = 0usize;
    while index < bytes.len() {
        match bytes[index] {
            b'\'' => {
                if quoted && bytes.get(index + 1) == Some(&b'\'') {
                    index += 2;
                    continue;
                }
                quoted = !quoted;
            },
            b'[' if !quoted => brackets = brackets.saturating_add(1),
            b']' if !quoted => brackets = brackets.saturating_sub(1),
            byte if !quoted && brackets == 0 && byte == delimiter => return Some(index),
            _ => {},
        }
        index += 1;
    }
    None
}

fn validate_reference_qualifier(value: &str) -> Result<()> {
    let mut quoted = false;
    let mut bracket_depth = 0usize;
    let mut meaningful = false;
    let mut characters = value.chars().peekable();
    while let Some(character) = characters.next() {
        if quoted {
            if character == '\'' {
                if characters.next_if_eq(&'\'').is_none() {
                    quoted = false;
                }
            } else {
                meaningful = true;
            }
            continue;
        }
        match character {
            '\'' => quoted = true,
            '[' => {
                bracket_depth = bracket_depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("form-control formula qualifier depth overflow"))?;
                meaningful = true;
            },
            ']' => {
                if bracket_depth == 0 {
                    return Err(invalid("form-control formula has an unmatched ]"));
                }
                bracket_depth -= 1;
            },
            character if bracket_depth != 0 => {
                if character == '\u{0}' || character == '\u{7f}' || character.is_control() {
                    return Err(invalid(
                        "form-control formula has an invalid sheet qualifier",
                    ));
                }
                meaningful = true;
            },
            character
                if character.is_ascii_alphanumeric()
                    || character.is_alphabetic()
                    || matches!(character, '_' | '.' | ':' | '-' | '\\' | '$') =>
            {
                meaningful = true;
            },
            _ => {
                return Err(invalid(
                    "form-control formula has an invalid sheet qualifier",
                ));
            },
        }
    }
    if quoted || bracket_depth != 0 || !meaningful {
        return Err(invalid(
            "form-control formula has an invalid sheet qualifier",
        ));
    }
    Ok(())
}

fn validate_reference_atom(value: &str) -> Result<()> {
    if value.is_empty() {
        return Err(invalid("form-control formula has an empty reference atom"));
    }
    if is_r1c1_reference(value) || is_a1_reference(value) || is_defined_name(value) {
        return Ok(());
    }
    Err(invalid(
        "form-control formula is not an A1, R1C1, or defined-name reference",
    ))
}

fn is_a1_reference(value: &str) -> bool {
    let bytes = value.as_bytes();
    let mut cursor = 0usize;
    if bytes.first() == Some(&b'$') {
        cursor += 1;
    }
    let column_start = cursor;
    while bytes.get(cursor).is_some_and(u8::is_ascii_alphabetic) {
        cursor += 1;
    }
    if cursor == column_start {
        return false;
    }
    if bytes.get(cursor) == Some(&b'$') {
        cursor += 1;
    }
    let row_start = cursor;
    while bytes.get(cursor).is_some_and(u8::is_ascii_digit) {
        cursor += 1;
    }
    row_start != cursor && cursor == bytes.len()
}

fn is_r1c1_reference(value: &str) -> bool {
    let bytes = value.as_bytes();
    let mut cursor = 0usize;
    if !matches!(bytes.first(), Some(b'R' | b'r')) {
        return false;
    }
    if !consume_r1c1_component(bytes, &mut cursor) {
        return false;
    }
    if !matches!(bytes.get(cursor), Some(b'C' | b'c')) {
        return false;
    }
    cursor += 1;
    consume_r1c1_component(bytes, &mut cursor) && cursor == bytes.len()
}

fn consume_r1c1_component(bytes: &[u8], cursor: &mut usize) -> bool {
    if bytes.get(*cursor) == Some(&b'[') {
        *cursor += 1;
        if matches!(bytes.get(*cursor), Some(b'-' | b'+')) {
            *cursor += 1;
        }
        let start = *cursor;
        while bytes.get(*cursor).is_some_and(u8::is_ascii_digit) {
            *cursor += 1;
        }
        if start == *cursor || bytes.get(*cursor) != Some(&b']') {
            return false;
        }
        *cursor += 1;
        return true;
    }
    while bytes.get(*cursor).is_some_and(u8::is_ascii_digit) {
        *cursor += 1;
    }
    true
}

fn is_defined_name(value: &str) -> bool {
    let mut characters = value.chars();
    let Some(first) = characters.next() else {
        return false;
    };
    if !(first == '_' || first == '\\' || first.is_alphabetic()) {
        return false;
    }
    characters.all(|character| {
        character == '_' || character == '\\' || character == '.' || character.is_alphanumeric()
    })
}

fn scalar_applies(field: ScalarField, object_type: ObjectType) -> bool {
    match field {
        ScalarField::ObjectType
        | ScalarField::JustLastX
        | ScalarField::TextHAlign
        | ScalarField::TextVAlign => true,
        ScalarField::Checked => matches!(object_type, ObjectType::CheckBox | ObjectType::Radio),
        ScalarField::Colored
        | ScalarField::DropLines
        | ScalarField::DropStyle
        | ScalarField::WidthMin => object_type == ObjectType::Drop,
        ScalarField::Dx => matches!(
            object_type,
            ObjectType::List | ObjectType::Scroll | ObjectType::Spin | ObjectType::Drop
        ),
        ScalarField::FirstButton => object_type == ObjectType::Radio,
        ScalarField::FmlaGroup => object_type == ObjectType::GBox,
        ScalarField::FmlaLink => matches!(
            object_type,
            ObjectType::CheckBox
                | ObjectType::Radio
                | ObjectType::Scroll
                | ObjectType::Spin
                | ObjectType::Drop
                | ObjectType::List
        ),
        ScalarField::FmlaRange | ScalarField::Sel => {
            matches!(object_type, ObjectType::List | ObjectType::Drop)
        },
        ScalarField::FmlaTxbx => matches!(object_type, ObjectType::Label | ObjectType::EditBox),
        ScalarField::Horiz
        | ScalarField::Inc
        | ScalarField::Max
        | ScalarField::Min
        | ScalarField::Page => matches!(object_type, ObjectType::Scroll | ObjectType::Spin),
        ScalarField::LockText => matches!(
            object_type,
            ObjectType::Button | ObjectType::Radio | ObjectType::CheckBox | ObjectType::Label
        ),
        ScalarField::MultiSel | ScalarField::SelType => object_type == ObjectType::List,
        ScalarField::NoThreeD => matches!(
            object_type,
            ObjectType::CheckBox
                | ObjectType::Radio
                | ObjectType::GBox
                | ObjectType::Scroll
                | ObjectType::Drop
                | ObjectType::List
                | ObjectType::Spin
        ),
        ScalarField::NoThreeD2 => matches!(object_type, ObjectType::Drop | ObjectType::List),
        ScalarField::Val => matches!(
            object_type,
            ObjectType::Scroll | ObjectType::Spin | ObjectType::List | ObjectType::Drop
        ),
        ScalarField::EditVal
        | ScalarField::MultiLine
        | ScalarField::VerticalBar
        | ScalarField::PasswordEdit => object_type == ObjectType::EditBox,
    }
}

fn scalar_applicability_fields() -> &'static [ScalarField] {
    &[
        ScalarField::Checked,
        ScalarField::Colored,
        ScalarField::DropLines,
        ScalarField::DropStyle,
        ScalarField::Dx,
        ScalarField::FirstButton,
        ScalarField::FmlaGroup,
        ScalarField::FmlaLink,
        ScalarField::FmlaRange,
        ScalarField::FmlaTxbx,
        ScalarField::Horiz,
        ScalarField::Inc,
        ScalarField::LockText,
        ScalarField::Max,
        ScalarField::Min,
        ScalarField::MultiSel,
        ScalarField::NoThreeD,
        ScalarField::NoThreeD2,
        ScalarField::Page,
        ScalarField::Sel,
        ScalarField::SelType,
        ScalarField::Val,
        ScalarField::WidthMin,
        ScalarField::EditVal,
        ScalarField::MultiLine,
        ScalarField::VerticalBar,
        ScalarField::PasswordEdit,
    ]
}

fn validate_applicability(value: &Properties, object_type: ObjectType) -> Result<()> {
    for (field, present) in [
        (ScalarField::Checked, value.checked.is_some()),
        (ScalarField::Colored, value.colored.is_some()),
        (ScalarField::DropLines, value.drop_lines.is_some()),
        (ScalarField::DropStyle, value.drop_style.is_some()),
        (ScalarField::Dx, value.dx.is_some()),
        (ScalarField::FirstButton, value.first_button.is_some()),
        (ScalarField::FmlaGroup, value.fmla_group.is_some()),
        (ScalarField::FmlaLink, value.fmla_link.is_some()),
        (ScalarField::FmlaRange, value.fmla_range.is_some()),
        (ScalarField::FmlaTxbx, value.fmla_txbx.is_some()),
        (ScalarField::Horiz, value.horiz.is_some()),
        (ScalarField::Inc, value.inc.is_some()),
        (ScalarField::LockText, value.lock_text.is_some()),
        (ScalarField::Max, value.max.is_some()),
        (ScalarField::Min, value.min.is_some()),
        (ScalarField::MultiSel, value.multi_sel.is_some()),
        (ScalarField::NoThreeD, value.no_three_d.is_some()),
        (ScalarField::NoThreeD2, value.no_three_d2.is_some()),
        (ScalarField::Page, value.page.is_some()),
        (ScalarField::Sel, value.sel.is_some()),
        (ScalarField::SelType, value.seltype.is_some()),
        (ScalarField::Val, value.val.is_some()),
        (ScalarField::WidthMin, value.width_min.is_some()),
        (ScalarField::EditVal, value.edit_val.is_some()),
        (ScalarField::MultiLine, value.multi_line.is_some()),
        (ScalarField::VerticalBar, value.vertical_bar.is_some()),
        (ScalarField::PasswordEdit, value.password_edit.is_some()),
    ] {
        if present && !scalar_applies(field, object_type) {
            return Err(invalid(format!(
                "form-control attribute {field:?} is inapplicable to {object_type}"
            )));
        }
    }
    if value.item_list.is_some() && !matches!(object_type, ObjectType::List | ObjectType::Drop) {
        return Err(invalid("itemLst applies only to List and Drop controls"));
    }
    if matches!(
        value.checked.as_ref(),
        Some(KnownOrUnknown::Known(Checked::Mixed))
    ) && object_type != ObjectType::CheckBox
    {
        return Err(invalid("checked=Mixed applies only to CheckBox controls"));
    }
    if value.multi_sel.is_some()
        && !matches!(
            value.seltype.as_ref(),
            Some(KnownOrUnknown::Known(SelectionType::Multi))
        )
    {
        return Err(invalid("multiSel requires seltype=multi"));
    }
    if let Some(sel) = value.sel {
        if let Some(item_list) = value.item_list.as_ref() {
            if sel as usize > item_list.items.len() && sel != 0 {
                return Err(invalid("sel exceeds the authored item count"));
            }
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_core::{
        Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as BudgetLimits,
        Resource,
    };
    use std::num::{NonZeroU64, NonZeroUsize};

    #[test]
    fn properties_clone_shares_backing_until_mutation() {
        let mut original = Properties::new();
        original.set_colored(Some(true));
        original.set_object_type(Some(ObjectType::Button));
        let clone = original.clone();
        assert!(Arc::ptr_eq(
            original
                .object_type
                .0
                .as_ref()
                .expect("original object type"),
            clone.object_type.0.as_ref().expect("cloned object type")
        ));

        let mut changed = clone;
        changed.set_object_type(Some(ObjectType::CheckBox));
        assert!(!Arc::ptr_eq(
            original
                .object_type
                .0
                .as_ref()
                .expect("original object type"),
            changed.object_type.0.as_ref().expect("changed object type")
        ));
        assert_eq!(original.colored(), Some(true));
        assert_eq!(
            original.object_type(),
            Some(&KnownOrUnknown::Known(ObjectType::Button))
        );
        assert_eq!(
            changed.object_type(),
            Some(&KnownOrUnknown::Known(ObjectType::CheckBox))
        );
    }

    #[test]
    fn opaque_xml_clone_keeps_generated_source_budget_alive() {
        let budget = Budget::root(
            "form-control-opaque-clone-test",
            BudgetLimits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let (cancellation_source, cancellation) = CancellationSource::pair();
        let context = ExecutionContext::new(
            budget.clone(),
            cancellation,
            ExecutionLimits::new(
                NonZeroUsize::new(1).expect("one worker"),
                NonZeroUsize::new(1).expect("one in-flight task"),
                NonZeroU64::new(u64::MAX).expect("in-flight bytes"),
                0,
            )
            .expect("execution limits"),
        );
        let hold =
            super::super::budget::reserve_generated(Some(&context), 3, 1, "opaque clone test")
                .expect("generated lease")
                .expect("context supplies generated lease");
        let source = SourcePayload::owned_budgeted(Arc::new(vec![b'<', b'x', b'>']), Some(hold));
        let opaque = OpaqueXml::from_range(
            source,
            0..3,
            Arc::from(Vec::<NamespaceBinding>::new().into_boxed_slice()),
        )
        .expect("opaque source range");
        let cloned = opaque.clone();
        drop(opaque);
        assert!(budget.used(Resource::Memory) > 0);
        assert!(budget.used(Resource::Objects) > 0);
        drop(cloned);
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(budget.used(Resource::Objects), 0);
        drop(cancellation_source);
    }

    #[test]
    fn shared_vec_replacement_reuses_unique_backing() {
        let mut values = SharedVec::new(Vec::<u8>::new());
        let original = Arc::as_ptr(&values.0);
        values.replace(vec![1, 2, 3]);
        assert!(std::ptr::eq(original, Arc::as_ptr(&values.0)));
        assert_eq!(values.as_slice(), &[1, 2, 3]);

        let mut shared = values.clone();
        let shared_original = Arc::as_ptr(&shared.0);
        shared.replace(vec![4]);
        assert!(!std::ptr::eq(shared_original, Arc::as_ptr(&shared.0)));
        assert_eq!(values.as_slice(), &[1, 2, 3]);
        assert_eq!(shared.as_slice(), &[4]);
    }
}
