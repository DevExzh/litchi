//! Bounded lexical codec and source-local splices for `formControlPr`.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::match_same_arms,
    clippy::too_many_lines,
    reason = "the parser and writer are kept in schema and source order"
)]

use std::mem::size_of;
use std::ops::Deref;
use std::ops::Range;
use std::sync::Arc;

use litchi_core::xml::ReaderOrigin;
use litchi_core::{ExecutionContext, ExecutionError, Reservation, Resource};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesRef, BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;

use crate::source_payload::SourcePayload;

use super::model::{
    Checked, DropStyle, EditValidation, FormControlFormula, Item, ItemList, KnownOrUnknown,
    LexicalAttribute, NamespaceBinding, ObjectType, OpaqueAttribute, OpaqueXml,
    PROPERTIES_STORAGE_BYTES, Properties, RETAINED_LEASE_STORAGE_BYTES, SHARED_VEC_STORAGE_BYTES,
    ScalarField, ScalarValue, SelectionType, SharedOption, SharedVec, TextHAlign, TextVAlign,
    is_xml_10_char, validate_multi_selection, validate_xml_text,
};
use super::{FORM_CONTROL_NAMESPACE, Limits, Result, allocation, invalid, limit};

const ROOT: &[u8] = b"formControlPr";
const ITEM_LIST: &[u8] = b"itemLst";
const ITEM: &[u8] = b"item";
const EXT_LIST: &[u8] = b"extLst";
const ARC_HEADER_BYTES: usize = 2 * size_of::<usize>();

/// Parse one complete control-properties XML part.
pub fn parse(xml: &[u8]) -> Result<Properties> {
    parse_with_limits(xml, &Limits::default())
}

/// Parse one complete control-properties XML part under a caller policy.
pub fn parse_with_limits(xml: &[u8], limits: &Limits) -> Result<Properties> {
    let source = owned_source(xml, *limits)?;
    let inspected = inspect_with_limits_owned(xml, *limits, Some(source), None)?;
    Ok(inspected.properties)
}

/// Parse one source-backed control-properties part while retaining the
/// caller-provided payload handle.
///
/// The source payload may be a managed OPC [`litchi_opc::PartData`] handle.  Parsing only
/// borrows its bytes and stores clones of that handle in the properties model
/// and any retained opaque ranges; it never calls `into_arc()` and therefore
/// does not detach a managed reservation.
#[allow(dead_code, reason = "called by the source-backed form-control owner")]
pub(crate) fn parse_source_with_limits(
    source: SourcePayload,
    limits: &Limits,
) -> Result<Properties> {
    parse_source_with_limits_and_context(source, limits, None)
}

/// Parse one source-backed part while charging parser-owned semantic memory to
/// the caller's execution context. The source payload itself is assumed to be
/// already budgeted by its owner (as it is for managed OPC `PartData`), so the
/// retained lease covers only parser allocations and is attached to the
/// returned model for its full clone-shared lifetime. One unit of cumulative
/// `Work` is also consumed for each XML event when a context is supplied.
#[allow(dead_code, reason = "called by the source-backed form-control owner")]
pub(crate) fn parse_source_with_limits_and_context(
    source: SourcePayload,
    limits: &Limits,
    context: Option<&ExecutionContext>,
) -> Result<Properties> {
    let xml = source.as_bytes();
    let inspected = inspect_with_limits_owned(xml, *limits, Some(source.clone()), context)?;
    Ok(inspected.properties)
}

/// Inspect a source part while retaining a borrowed source view for splices.
pub fn inspect(xml: &[u8]) -> Result<SourceView<'_>> {
    inspect_with_limits(xml, Limits::default())
}

/// Inspect a source part under a caller policy.
pub fn inspect_with_limits(xml: &[u8], limits: Limits) -> Result<SourceView<'_>> {
    let inspected = inspect_with_limits_owned(xml, limits, None, None)?;
    Ok(SourceView {
        source: xml,
        properties: inspected.properties,
        layout: inspected.layout,
        limits,
    })
}

/// A borrowed source-bound properties part.
pub struct SourceView<'a> {
    source: &'a [u8],
    properties: Properties,
    layout: Layout,
    limits: Limits,
}

/// Borrowed typed projection paired with its original source member.
///
/// The projection keeps source bytes borrowed by the view.  Passing it to
/// [`write()`] therefore performs an exact no-op without materializing another
/// copy of the complete part.  The source-range accessors expose opaque item
/// and extension bytes without copying them; use [`parse`] when an owned
/// detached model including those opaque fragments is required.
pub struct SourceProperties<'a> {
    properties: &'a Properties,
    source: &'a [u8],
    layout: &'a Layout,
}

impl<'a> Deref for SourceProperties<'a> {
    type Target = Properties;

    fn deref(&self) -> &Self::Target {
        self.properties
    }
}

impl SourceProperties<'_> {
    /// Borrow the exact root `extLst` source range, when present.
    #[must_use]
    pub fn root_extension_xml(&self) -> Option<&[u8]> {
        self.layout
            .root_extension
            .as_ref()
            .and_then(|range| self.source.get(range.clone()))
    }

    /// Borrow the exact `itemLst/extLst` source range, when present.
    #[must_use]
    pub fn item_list_extension_xml(&self) -> Option<&[u8]> {
        self.layout
            .item_list
            .as_ref()
            .and_then(|list| list.extension.as_ref())
            .and_then(|range| self.source.get(range.clone()))
    }

    /// Borrow one exact source item range by ordered item position.
    #[must_use]
    pub fn item_xml(&self, index: usize) -> Option<&[u8]> {
        self.layout
            .item_list
            .as_ref()
            .and_then(|list| list.items.get(index))
            .and_then(|item| self.source.get(item.span.clone()))
    }
}

/// Values accepted by the detached/source-backed writer.
pub trait SourceWritable {
    fn write_with_limits(&self, limits: &Limits) -> Result<Vec<u8>>;
}

impl SourceWritable for Properties {
    fn write_with_limits(&self, limits: &Limits) -> Result<Vec<u8>> {
        write_properties(self, limits)
    }
}

impl SourceWritable for &Properties {
    fn write_with_limits(&self, limits: &Limits) -> Result<Vec<u8>> {
        write_properties(self, limits)
    }
}

impl<'a> SourceWritable for SourceProperties<'a> {
    fn write_with_limits(&self, limits: &Limits) -> Result<Vec<u8>> {
        copy_bounded(
            self.source,
            limits.max_output_bytes(),
            "form-control no-op output",
        )
    }
}

impl<'a> SourceWritable for &SourceProperties<'a> {
    fn write_with_limits(&self, limits: &Limits) -> Result<Vec<u8>> {
        copy_bounded(
            self.source,
            limits.max_output_bytes(),
            "form-control no-op output",
        )
    }
}

impl<'a> SourceView<'a> {
    /// Original source bytes.
    #[must_use]
    pub const fn source(&self) -> &'a [u8] {
        self.source
    }

    /// Parsed typed properties.
    #[must_use]
    pub fn properties(&self) -> SourceProperties<'_> {
        SourceProperties {
            properties: &self.properties,
            source: self.source,
            layout: &self.layout,
        }
    }

    /// Borrow the exact root `extLst` source range, when present.
    #[must_use]
    pub fn root_extension_xml(&self) -> Option<&'a [u8]> {
        self.layout
            .root_extension
            .as_ref()
            .and_then(|range| self.source.get(range.clone()))
    }

    /// Borrow the exact `itemLst/extLst` source range, when present.
    #[must_use]
    pub fn item_list_extension_xml(&self) -> Option<&'a [u8]> {
        self.layout
            .item_list
            .as_ref()
            .and_then(|list| list.extension.as_ref())
            .and_then(|range| self.source.get(range.clone()))
    }

    /// Borrow one exact source item range by ordered item position.
    #[must_use]
    pub fn item_xml(&self, index: usize) -> Option<&'a [u8]> {
        self.layout
            .item_list
            .as_ref()
            .and_then(|list| list.items.get(index))
            .and_then(|item| self.source.get(item.span.clone()))
    }

    /// Replace one scalar attribute by source-range splice.
    pub fn replace_scalar(
        &self,
        field: ScalarField,
        value: Option<ScalarValue>,
    ) -> Result<Vec<u8>> {
        replace_scalar_from_parts(
            self.source,
            field,
            value,
            self.limits,
            &self.properties,
            &self.layout,
        )
    }

    /// Replace ordered list items by source-range splice.
    pub fn replace_items(&self, values: &[Item]) -> Result<Vec<u8>> {
        replace_items_from_parts(
            self.source,
            values,
            self.limits,
            &self.properties,
            &self.layout,
        )
    }

    /// Insert one item at a checked zero-based position.
    pub fn insert_item(&self, index: usize, value: Item) -> Result<Vec<u8>> {
        insert_item_from_parts(
            self.source,
            index,
            value,
            self.limits,
            &self.properties,
            &self.layout,
        )
    }

    /// Remove one item at a checked zero-based position.
    pub fn remove_item(&self, index: usize) -> Result<Vec<u8>> {
        remove_item_from_parts(
            self.source,
            index,
            self.limits,
            &self.properties,
            &self.layout,
        )
    }
}

/// Write a detached or unchanged source-backed properties model.
pub fn write<P: SourceWritable>(properties: P) -> Result<Vec<u8>> {
    write_with_limits(properties, &Limits::default())
}

/// Write a detached or unchanged source-backed properties model under limits.
pub fn write_with_limits<P: SourceWritable>(properties: P, limits: &Limits) -> Result<Vec<u8>> {
    properties.write_with_limits(limits)
}

fn write_properties(properties: &Properties, limits: &Limits) -> Result<Vec<u8>> {
    if properties.source_unchanged() {
        let source = properties
            .source_bytes()
            .ok_or_else(|| invalid("source-backed properties lost their source bytes"))?;
        return copy_bounded(
            source,
            limits.max_output_bytes(),
            "form-control no-op output",
        );
    }
    properties.validate_with_limits(*limits)?;
    write_detached(properties, *limits)
}

/// Replace one scalar attribute by a source-local splice.
pub fn replace_scalar(
    xml: &[u8],
    field: ScalarField,
    value: Option<ScalarValue>,
) -> Result<Vec<u8>> {
    replace_scalar_with_limits(xml, field, value, Limits::default())
}

/// Alias for [`replace_scalar`].
pub fn set_scalar(xml: &[u8], field: ScalarField, value: Option<ScalarValue>) -> Result<Vec<u8>> {
    replace_scalar(xml, field, value)
}

/// Replace ordered list items by source-local splice.
pub fn replace_items(xml: &[u8], values: &[Item]) -> Result<Vec<u8>> {
    replace_items_with_limits(xml, values, Limits::default())
}

/// Insert one item at a checked zero-based position.
pub fn insert_item(xml: &[u8], index: usize, value: Item) -> Result<Vec<u8>> {
    insert_item_with_limits(xml, index, value, Limits::default())
}

/// Remove one item at a checked zero-based position.
pub fn remove_item(xml: &[u8], index: usize) -> Result<Vec<u8>> {
    remove_item_with_limits(xml, index, Limits::default())
}

fn replace_scalar_with_limits(
    xml: &[u8],
    field: ScalarField,
    value: Option<ScalarValue>,
    limits: Limits,
) -> Result<Vec<u8>> {
    let inspected = inspect_with_limits(xml, limits)?;
    replace_scalar_from_parts(
        xml,
        field,
        value,
        limits,
        &inspected.properties,
        &inspected.layout,
    )
}

fn replace_scalar_from_parts(
    xml: &[u8],
    field: ScalarField,
    value: Option<ScalarValue>,
    limits: Limits,
    properties: &Properties,
    layout: &Layout,
) -> Result<Vec<u8>> {
    validate_scalar_value(field, value.as_ref())?;
    properties.validate_scalar_edit(field, value.as_ref(), limits)?;
    if scalar_value_matches(properties, field, value.as_ref()) {
        return copy_bounded(
            xml,
            limits.max_output_bytes(),
            "form-control scalar no-op output",
        );
    }
    let attribute = layout
        .attributes
        .iter()
        .find(|attribute| attribute.field == Some(field));
    if let Some(value) = value.as_ref() {
        let encoded_len = scalar_xml_value_len(field, value, limits)?;
        let output_len = match attribute {
            Some(attribute) => xml
                .len()
                .checked_sub(attribute.value.end - attribute.value.start)
                .and_then(|value| value.checked_add(encoded_len))
                .ok_or_else(|| invalid("form-control scalar output size overflow"))?,
            None => xml
                .len()
                .checked_add(field.wire_name().len())
                .and_then(|value| value.checked_add(encoded_len + 4))
                .ok_or_else(|| invalid("form-control scalar output size overflow"))?,
        };
        if output_len > limits.max_output_bytes() {
            return Err(limit(
                "generated output bytes",
                output_len,
                limits.max_output_bytes(),
            ));
        }
    }
    let replacement = value
        .as_ref()
        .map(|value| scalar_xml_value(field, value, limits))
        .transpose()?;
    match (attribute, replacement) {
        (Some(attribute), Some(replacement)) => {
            splice_ranges(xml, &[(attribute.value.clone(), replacement)], limits)
        },
        (Some(attribute), None) => {
            splice_ranges(xml, &[(attribute.token.clone(), Vec::new())], limits)
        },
        (None, Some(replacement)) => {
            let insertion = layout
                .root_attribute_insertion
                .ok_or_else(|| invalid("form-control root has no opening-tag insertion point"))?;
            let mut bytes = Vec::new();
            let attribute_len = field
                .wire_name()
                .len()
                .checked_add(replacement.len())
                .and_then(|value| value.checked_add(4))
                .ok_or_else(|| invalid("form-control scalar attribute length overflow"))?;
            bytes
                .try_reserve_exact(attribute_len)
                .map_err(|source| allocation("form-control scalar attribute", source))?;
            bytes.push(b' ');
            bytes.extend_from_slice(field.wire_name().as_bytes());
            bytes.extend_from_slice(b"=\"");
            bytes.extend_from_slice(&replacement);
            bytes.push(b'"');
            splice_ranges(xml, &[(insertion..insertion, bytes)], limits)
        },
        (None, None) => copy_bounded(
            xml,
            limits.max_output_bytes(),
            "form-control scalar no-op output",
        ),
    }
}

fn scalar_value_matches(
    properties: &Properties,
    field: ScalarField,
    value: Option<&ScalarValue>,
) -> bool {
    fn known_enum_matches<T: PartialEq>(
        current: Option<&KnownOrUnknown<T>>,
        value: Option<&T>,
    ) -> bool {
        match (current, value) {
            (None, None) => true,
            (Some(KnownOrUnknown::Known(current)), Some(value)) => current == value,
            _ => false,
        }
    }

    match field {
        ScalarField::ObjectType => known_enum_matches(
            properties.object_type.as_ref(),
            value.and_then(|value| match value {
                ScalarValue::ObjectType(value) => Some(value),
                _ => None,
            }),
        ),
        ScalarField::Checked => known_enum_matches(
            properties.checked.as_ref(),
            value.and_then(|value| match value {
                ScalarValue::Checked(value) => Some(value),
                _ => None,
            }),
        ),
        ScalarField::Colored => {
            properties.colored
                == value.and_then(|v| match v {
                    ScalarValue::Boolean(value) => Some(*value),
                    _ => None,
                })
        },
        ScalarField::DropLines => {
            properties.drop_lines
                == value.and_then(|v| match v {
                    ScalarValue::Unsigned(value) => Some(*value),
                    _ => None,
                })
        },
        ScalarField::DropStyle => known_enum_matches(
            properties.drop_style.as_ref(),
            value.and_then(|value| match value {
                ScalarValue::DropStyle(value) => Some(value),
                _ => None,
            }),
        ),
        ScalarField::Dx => {
            properties.dx
                == value.and_then(|v| match v {
                    ScalarValue::Unsigned(value) => Some(*value),
                    _ => None,
                })
        },
        ScalarField::FirstButton => {
            properties.first_button
                == value.and_then(|v| match v {
                    ScalarValue::Boolean(value) => Some(*value),
                    _ => None,
                })
        },
        ScalarField::FmlaGroup => {
            properties.fmla_group.as_ref()
                == value.and_then(|v| match v {
                    ScalarValue::Formula(value) => Some(value),
                    _ => None,
                })
        },
        ScalarField::FmlaLink => {
            properties.fmla_link.as_ref()
                == value.and_then(|v| match v {
                    ScalarValue::Formula(value) => Some(value),
                    _ => None,
                })
        },
        ScalarField::FmlaRange => {
            properties.fmla_range.as_ref()
                == value.and_then(|v| match v {
                    ScalarValue::Formula(value) => Some(value),
                    _ => None,
                })
        },
        ScalarField::FmlaTxbx => {
            properties.fmla_txbx.as_ref()
                == value.and_then(|v| match v {
                    ScalarValue::Formula(value) => Some(value),
                    _ => None,
                })
        },
        ScalarField::Horiz => {
            properties.horiz
                == value.and_then(|v| match v {
                    ScalarValue::Boolean(value) => Some(*value),
                    _ => None,
                })
        },
        ScalarField::Inc => {
            properties.inc
                == value.and_then(|v| match v {
                    ScalarValue::Unsigned(value) => Some(*value),
                    _ => None,
                })
        },
        ScalarField::JustLastX => {
            properties.just_last_x
                == value.and_then(|v| match v {
                    ScalarValue::Boolean(value) => Some(*value),
                    _ => None,
                })
        },
        ScalarField::LockText => {
            properties.lock_text
                == value.and_then(|v| match v {
                    ScalarValue::Boolean(value) => Some(*value),
                    _ => None,
                })
        },
        ScalarField::Max => {
            properties.max
                == value.and_then(|v| match v {
                    ScalarValue::Unsigned(value) => Some(*value),
                    _ => None,
                })
        },
        ScalarField::Min => {
            properties.min
                == value.and_then(|v| match v {
                    ScalarValue::Unsigned(value) => Some(*value),
                    _ => None,
                })
        },
        ScalarField::MultiSel => {
            properties.multi_sel.as_ref()
                == value.and_then(|v| match v {
                    ScalarValue::String(value) => Some(value),
                    _ => None,
                })
        },
        ScalarField::NoThreeD => {
            properties.no_three_d
                == value.and_then(|v| match v {
                    ScalarValue::Boolean(value) => Some(*value),
                    _ => None,
                })
        },
        ScalarField::NoThreeD2 => {
            properties.no_three_d2
                == value.and_then(|v| match v {
                    ScalarValue::Boolean(value) => Some(*value),
                    _ => None,
                })
        },
        ScalarField::Page => {
            properties.page
                == value.and_then(|v| match v {
                    ScalarValue::Unsigned(value) => Some(*value),
                    _ => None,
                })
        },
        ScalarField::Sel => {
            properties.sel
                == value.and_then(|v| match v {
                    ScalarValue::Unsigned(value) => Some(*value),
                    _ => None,
                })
        },
        ScalarField::SelType => known_enum_matches(
            properties.seltype.as_ref(),
            value.and_then(|value| match value {
                ScalarValue::SelectionType(value) => Some(value),
                _ => None,
            }),
        ),
        ScalarField::TextHAlign => known_enum_matches(
            properties.text_h_align.as_ref(),
            value.and_then(|value| match value {
                ScalarValue::TextHAlign(value) => Some(value),
                _ => None,
            }),
        ),
        ScalarField::TextVAlign => known_enum_matches(
            properties.text_v_align.as_ref(),
            value.and_then(|value| match value {
                ScalarValue::TextVAlign(value) => Some(value),
                _ => None,
            }),
        ),
        ScalarField::Val => {
            properties.val
                == value.and_then(|v| match v {
                    ScalarValue::Unsigned(value) => Some(*value),
                    _ => None,
                })
        },
        ScalarField::WidthMin => {
            properties.width_min
                == value.and_then(|v| match v {
                    ScalarValue::Unsigned(value) => Some(*value),
                    _ => None,
                })
        },
        ScalarField::EditVal => known_enum_matches(
            properties.edit_val.as_ref(),
            value.and_then(|value| match value {
                ScalarValue::EditValidation(value) => Some(value),
                _ => None,
            }),
        ),
        ScalarField::MultiLine => {
            properties.multi_line
                == value.and_then(|v| match v {
                    ScalarValue::Boolean(value) => Some(*value),
                    _ => None,
                })
        },
        ScalarField::VerticalBar => {
            properties.vertical_bar
                == value.and_then(|v| match v {
                    ScalarValue::Boolean(value) => Some(*value),
                    _ => None,
                })
        },
        ScalarField::PasswordEdit => {
            properties.password_edit
                == value.and_then(|v| match v {
                    ScalarValue::Boolean(value) => Some(*value),
                    _ => None,
                })
        },
    }
}

fn replace_items_with_limits(xml: &[u8], values: &[Item], limits: Limits) -> Result<Vec<u8>> {
    let inspected = inspect_with_limits(xml, limits)?;
    replace_items_from_parts(
        xml,
        values,
        limits,
        &inspected.properties,
        &inspected.layout,
    )
}

fn replace_items_from_parts(
    xml: &[u8],
    values: &[Item],
    limits: Limits,
    properties: &Properties,
    layout: &Layout,
) -> Result<Vec<u8>> {
    if values.len() > limits.max_items() {
        return Err(limit("item count", values.len(), limits.max_items()));
    }
    let current = properties
        .item_list
        .as_ref()
        .map(|list| list.items())
        .unwrap_or(&[]);
    if items_equal_by_value(current, values) {
        return copy_bounded(
            xml,
            limits.max_output_bytes(),
            "form-control item no-op output",
        );
    }
    for (index, old) in current.iter().enumerate() {
        let Some(new) = values.get(index) else {
            break;
        };
        if old.opaque_attributes {
            if old.value() != new.value() {
                return Err(invalid(
                    "item-list rewrite would discard opaque item attributes or markup",
                ));
            }
        }
    }
    for value in values {
        validate_item_for_splice(value, limits)?;
    }
    validate_item_replacement(properties, values, limits)?;
    if let Some(list) = layout.item_list.as_ref()
        && !list.empty
        && !list.items.is_empty()
    {
        return replace_nonempty_item_list_source(
            xml,
            values,
            current,
            list,
            layout.root_prefix.as_deref(),
            limits,
        );
    }
    let prefix = layout.root_prefix.as_deref();
    let item_name_len = qname_len(prefix, ITEM)?;
    let item_xml_len = items_output_len_with_name_len(item_name_len, values, limits)?;
    preflight_item_xml_replacement(xml, item_xml_len, limits, layout)?;
    let item_xml = write_items(prefix, values, limits)?;
    replace_item_xml_from_parts(xml, item_xml, limits, layout)
}

fn preflight_item_xml_replacement(
    xml: &[u8],
    item_xml_len: usize,
    limits: Limits,
    layout: &Layout,
) -> Result<usize> {
    let output_len = match layout.item_list.as_ref() {
        Some(list) if list.empty => {
            let name_len = qname_len(layout.root_prefix.as_deref(), ITEM_LIST)?;
            let closing_len = name_len
                .checked_add(4)
                .ok_or_else(|| invalid("form-control itemLst output size overflow"))?;
            xml.len()
                .checked_sub(2)
                .and_then(|value| value.checked_add(item_xml_len))
                .and_then(|value| value.checked_add(closing_len))
                .ok_or_else(|| invalid("form-control itemLst output size overflow"))?
        },
        Some(list) if list.items.is_empty() => xml
            .len()
            .checked_add(item_xml_len)
            .ok_or_else(|| invalid("form-control itemLst output size overflow"))?,
        Some(list) => {
            let mut removed = 0usize;
            for item in &list.items {
                removed = removed
                    .checked_add(item.span.end - item.span.start)
                    .ok_or_else(|| invalid("form-control item output size overflow"))?;
            }
            xml.len()
                .checked_sub(removed)
                .and_then(|value| value.checked_add(item_xml_len))
                .ok_or_else(|| invalid("form-control item output size overflow"))?
        },
        None if item_xml_len == 0 => xml.len(),
        None => {
            let prefix = layout.root_prefix.as_deref();
            let name_len = qname_len(prefix, ITEM_LIST)?;
            let wrapper_len = name_len
                .checked_mul(2)
                .and_then(|value| value.checked_add(item_xml_len + 5))
                .ok_or_else(|| invalid("form-control itemLst output size overflow"))?;
            let root_open = &layout.root_open;
            let is_empty_root = xml
                .get(root_open.clone())
                .is_some_and(|value| value.ends_with(b"/>"));
            if is_empty_root {
                let root_name_len = qname_len(prefix, ROOT)?;
                let replacement_len = 1usize
                    .checked_add(wrapper_len)
                    .and_then(|value| value.checked_add(root_name_len + 3))
                    .ok_or_else(|| invalid("form-control root output size overflow"))?;
                xml.len()
                    .checked_sub(root_open.end - root_open.start)
                    .and_then(|value| value.checked_add(replacement_len))
                    .ok_or_else(|| invalid("form-control root output size overflow"))?
            } else {
                xml.len()
                    .checked_add(wrapper_len)
                    .ok_or_else(|| invalid("form-control itemLst output size overflow"))?
            }
        },
    };
    if output_len > limits.max_output_bytes() {
        return Err(limit(
            "generated output bytes",
            output_len,
            limits.max_output_bytes(),
        ));
    }
    Ok(output_len)
}

fn replace_nonempty_item_list_source(
    xml: &[u8],
    values: &[Item],
    current: &[Item],
    layout: &ItemListLayout,
    prefix: Option<&[u8]>,
    limits: Limits,
) -> Result<Vec<u8>> {
    if current.len() != layout.items.len() || current.is_empty() {
        return Err(invalid("form-control item layout count mismatch"));
    }
    let name_len = qname_len(prefix, ITEM)?;
    if name_len > limits.max_output_bytes() {
        return Err(limit(
            "generated item name",
            name_len,
            limits.max_output_bytes(),
        ));
    }
    let common = current.len().min(values.len());
    let mut output_len = xml.len();
    for index in 0..common {
        if current[index].value() == values[index].value() {
            continue;
        }
        if current[index].opaque_attributes {
            return Err(invalid(
                "item-list rewrite would discard opaque item attributes or markup",
            ));
        }
        let replacement_len =
            generated_item_len_for_item_with_name_len(name_len, &values[index], limits)?;
        let old_len = layout.items[index]
            .span
            .end
            .checked_sub(layout.items[index].span.start)
            .ok_or_else(|| invalid("form-control item source span underflow"))?;
        output_len = output_len
            .checked_sub(old_len)
            .and_then(|value| value.checked_add(replacement_len))
            .ok_or_else(|| invalid("form-control item output size overflow"))?;
    }
    for item in layout.items.iter().skip(values.len()) {
        let old_len = item
            .span
            .end
            .checked_sub(item.span.start)
            .ok_or_else(|| invalid("form-control item source span underflow"))?;
        output_len = output_len
            .checked_sub(old_len)
            .ok_or_else(|| invalid("form-control item output size underflow"))?;
    }
    let extra_start = current.len().min(values.len());
    let extra = values.get(extra_start..).unwrap_or_default();
    let extra_len = if extra.is_empty() {
        0
    } else {
        items_output_len_with_name_len(name_len, extra, limits)?
    };
    output_len = output_len
        .checked_add(extra_len)
        .ok_or_else(|| invalid("form-control item output size overflow"))?;
    if output_len > limits.max_output_bytes() {
        return Err(limit(
            "generated output bytes",
            output_len,
            limits.max_output_bytes(),
        ));
    }

    // All replacements are emitted directly into the one pre-sized final
    // candidate.  This preserves every untouched lexical span while avoiding
    // one temporary item buffer per changed item and a second copy in the
    // splice result.
    let name = qname(prefix, ITEM)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("form-control source output", source))?;
    let mut cursor = 0usize;
    for (index, item_layout) in layout.items.iter().enumerate() {
        if item_layout.span.start > item_layout.span.end || item_layout.span.end > xml.len() {
            return Err(invalid("form-control item source span lies outside source"));
        }
        output.extend_from_slice(&xml[cursor..item_layout.span.start]);
        if let Some(value) = values.get(index) {
            if current[index].value() == value.value() {
                output.extend_from_slice(&xml[item_layout.span.clone()]);
            } else {
                append_generated_item(&mut output, &name, value, limits)?;
            }
        }
        cursor = item_layout.span.end;
    }
    if !extra.is_empty() {
        append_items_direct(&mut output, &name, extra, limits)?;
    }
    output.extend_from_slice(&xml[cursor..]);
    Ok(output)
}

fn splice_nonempty_item_list_edit(
    xml: &[u8],
    layout: &ItemListLayout,
    prefix: Option<&[u8]>,
    limits: Limits,
    insertion: Option<(usize, &Item)>,
    removal: Option<usize>,
) -> Result<Vec<u8>> {
    if insertion.is_some() == removal.is_some() {
        return Err(invalid(
            "form-control item edit must insert or remove exactly one item",
        ));
    }
    let (range, output_len) = if let Some((index, item)) = insertion {
        if index > layout.items.len() {
            return Err(invalid(
                "form-control item insertion index is out of bounds",
            ));
        }
        let name_len = qname_len(prefix, ITEM)?;
        if name_len > limits.max_output_bytes() {
            return Err(limit(
                "generated item name",
                name_len,
                limits.max_output_bytes(),
            ));
        }
        let replacement_len = generated_item_len_for_item_with_name_len(name_len, item, limits)?;
        let insertion = if index == layout.items.len() {
            layout
                .items
                .last()
                .map_or(layout.span.end, |value| value.span.end)
        } else {
            layout.items[index].span.start
        };
        let output_len = xml
            .len()
            .checked_add(replacement_len)
            .ok_or_else(|| invalid("form-control item output size overflow"))?;
        if output_len > limits.max_output_bytes() {
            return Err(limit(
                "generated output bytes",
                output_len,
                limits.max_output_bytes(),
            ));
        }
        (insertion..insertion, output_len)
    } else if let Some(index) = removal {
        let Some(item) = layout.items.get(index) else {
            return Err(invalid("form-control item removal index is out of bounds"));
        };
        let removed = item
            .span
            .end
            .checked_sub(item.span.start)
            .ok_or_else(|| invalid("form-control item source span underflow"))?;
        let output_len = xml
            .len()
            .checked_sub(removed)
            .ok_or_else(|| invalid("form-control item output size underflow"))?;
        (item.span.clone(), output_len)
    } else {
        return Err(invalid("form-control item edit is empty"));
    };
    if output_len > limits.max_output_bytes() {
        return Err(limit(
            "generated output bytes",
            output_len,
            limits.max_output_bytes(),
        ));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("form-control source output", source))?;
    output.extend_from_slice(&xml[..range.start]);
    if let Some((_, item)) = insertion {
        let name = qname(prefix, ITEM)?;
        append_generated_item(&mut output, &name, item, limits)?;
    }
    output.extend_from_slice(&xml[range.end..]);
    Ok(output)
}

fn replace_item_xml_from_parts(
    xml: &[u8],
    item_xml: Vec<u8>,
    limits: Limits,
    layout: &Layout,
) -> Result<Vec<u8>> {
    match layout.item_list.as_ref() {
        Some(list) => {
            if list.empty {
                let prefix = layout.root_prefix.as_deref();
                let name_len = qname_len(prefix, ITEM_LIST)?;
                let replacement_len = name_len
                    .checked_add(item_xml.len())
                    .and_then(|value| value.checked_add(4))
                    .ok_or_else(|| invalid("itemLst output size overflow"))?;
                let insertion =
                    list.span.end.checked_sub(2).ok_or_else(|| {
                        invalid("form-control itemLst self-closing span underflow")
                    })?;
                let output_len = xml
                    .len()
                    .checked_sub(list.span.end - insertion)
                    .and_then(|value| value.checked_add(replacement_len))
                    .ok_or_else(|| invalid("form-control itemLst output size overflow"))?;
                if output_len > limits.max_output_bytes() {
                    return Err(limit(
                        "form-control itemLst output",
                        output_len,
                        limits.max_output_bytes(),
                    ));
                }
                let name = qname(prefix, ITEM_LIST)?;
                let mut replacement = Vec::new();
                replacement
                    .try_reserve_exact(replacement_len)
                    .map_err(|source| allocation("form-control itemLst output", source))?;
                replacement.extend_from_slice(b">");
                replacement.extend_from_slice(&item_xml);
                replacement.extend_from_slice(b"</");
                replacement.extend_from_slice(&name);
                replacement.push(b'>');
                splice_ranges(xml, &[(insertion..list.span.end, replacement)], limits)
            } else if list.items.is_empty() {
                let insertion = list
                    .first_content
                    .ok_or_else(|| invalid("empty itemLst has no insertion point"))?;
                splice_ranges(xml, &[(insertion..insertion, item_xml)], limits)
            } else {
                let mut edits = Vec::new();
                edits
                    .try_reserve_exact(list.items.len().saturating_add(1))
                    .map_err(|source| allocation("form-control item edits", source))?;
                for item in &list.items {
                    edits.push((item.span.clone(), Vec::new()));
                }
                let insertion = list.items[0].span.start;
                edits.push((insertion..insertion, item_xml));
                splice_ranges(xml, &edits, limits)
            }
        },
        None if item_xml.is_empty() => copy_bounded(
            xml,
            limits.max_output_bytes(),
            "form-control item no-op output",
        ),
        None => {
            let prefix = layout.root_prefix.as_deref();
            let name_len = qname_len(prefix, ITEM_LIST)?;
            let wrapper_len = name_len
                .checked_mul(2)
                .and_then(|value| value.checked_add(item_xml.len() + 5))
                .ok_or_else(|| invalid("itemLst output size overflow"))?;
            let root_open = &layout.root_open;
            let is_empty_root = xml
                .get(root_open.clone())
                .is_some_and(|value| value.ends_with(b"/>"));
            let replacement_len = if is_empty_root {
                let root_name_len = qname_len(prefix, ROOT)?;
                1usize
                    .checked_add(wrapper_len)
                    .and_then(|value| value.checked_add(root_name_len + 3))
                    .ok_or_else(|| invalid("form-control root output size overflow"))?
            } else {
                wrapper_len
            };
            if replacement_len > limits.max_output_bytes() {
                return Err(limit(
                    "form-control itemLst output",
                    replacement_len,
                    limits.max_output_bytes(),
                ));
            }
            let output_len = if is_empty_root {
                xml.len()
                    .checked_sub(root_open.end - root_open.start)
                    .and_then(|value| value.checked_add(replacement_len))
                    .ok_or_else(|| invalid("form-control root output size overflow"))?
            } else {
                xml.len()
                    .checked_add(replacement_len)
                    .ok_or_else(|| invalid("form-control itemLst output size overflow"))?
            };
            if output_len > limits.max_output_bytes() {
                return Err(limit(
                    "form-control itemLst output",
                    output_len,
                    limits.max_output_bytes(),
                ));
            }
            let name = qname(prefix, ITEM_LIST)?;
            let mut wrapper = Vec::new();
            wrapper
                .try_reserve_exact(wrapper_len)
                .map_err(|source| allocation("form-control itemLst", source))?;
            wrapper.extend_from_slice(b"<");
            wrapper.extend_from_slice(&name);
            wrapper.push(b'>');
            wrapper.extend_from_slice(&item_xml);
            wrapper.extend_from_slice(b"</");
            wrapper.extend_from_slice(&name);
            wrapper.push(b'>');
            if is_empty_root {
                let root_name = qname(prefix, ROOT)?;
                let mut replacement = Vec::new();
                replacement
                    .try_reserve_exact(replacement_len)
                    .map_err(|source| allocation("form-control root output", source))?;
                replacement.extend_from_slice(b">");
                replacement.extend_from_slice(&wrapper);
                replacement.extend_from_slice(b"</");
                replacement.extend_from_slice(&root_name);
                replacement.push(b'>');
                let slash = root_open
                    .end
                    .checked_sub(2)
                    .ok_or_else(|| invalid("form-control root self-closing span underflow"))?;
                splice_ranges(xml, &[(slash..root_open.end, replacement)], limits)
            } else {
                let insertion = layout
                    .root_child_insertion
                    .ok_or_else(|| invalid("form-control root has no child insertion point"))?;
                splice_ranges(xml, &[(insertion..insertion, wrapper)], limits)
            }
        },
    }
}

fn insert_item_with_limits(
    xml: &[u8],
    index: usize,
    value: Item,
    limits: Limits,
) -> Result<Vec<u8>> {
    let inspected = inspect_with_limits(xml, limits)?;
    insert_item_from_parts(
        xml,
        index,
        value,
        limits,
        &inspected.properties,
        &inspected.layout,
    )
}

fn remove_item_with_limits(xml: &[u8], index: usize, limits: Limits) -> Result<Vec<u8>> {
    let inspected = inspect_with_limits(xml, limits)?;
    remove_item_from_parts(xml, index, limits, &inspected.properties, &inspected.layout)
}

fn insert_item_from_parts(
    xml: &[u8],
    index: usize,
    value: Item,
    limits: Limits,
    properties: &Properties,
    layout: &Layout,
) -> Result<Vec<u8>> {
    let Some(item_list) = properties.item_list.as_ref() else {
        if index != 0 {
            return Err(invalid(
                "form-control item insertion index is out of bounds",
            ));
        }
        validate_item_for_splice(&value, limits)?;
        validate_item_edit(properties, true, 1, limits)?;
        let item_name_len = qname_len(layout.root_prefix.as_deref(), ITEM)?;
        let item_xml_len =
            generated_item_len_for_item_with_name_len(item_name_len, &value, limits)?;
        preflight_item_xml_replacement(xml, item_xml_len, limits, layout)?;
        let item_xml = write_items(layout.root_prefix.as_deref(), &[value], limits)?;
        return replace_item_xml_from_parts(xml, item_xml, limits, layout);
    };
    if index > item_list.items().len() {
        return Err(invalid(
            "form-control item insertion index is out of bounds",
        ));
    }
    validate_item_for_splice(&value, limits)?;
    validate_item_edit(properties, true, item_list.items().len() + 1, limits)?;
    if let Some(list) = layout.item_list.as_ref()
        && !list.empty
        && !list.items.is_empty()
    {
        return splice_nonempty_item_list_edit(
            xml,
            list,
            layout.root_prefix.as_deref(),
            limits,
            Some((index, &value)),
            None,
        );
    }
    let item_name_len = qname_len(layout.root_prefix.as_deref(), ITEM)?;
    let item_xml_len = generated_item_len_for_item_with_name_len(item_name_len, &value, limits)?;
    preflight_item_xml_replacement(xml, item_xml_len, limits, layout)?;
    let item_xml = write_items_with_source_edit(
        layout.root_prefix.as_deref(),
        item_list.items(),
        layout
            .item_list
            .as_ref()
            .ok_or_else(|| invalid("form-control item-list layout missing"))?,
        xml,
        limits,
        Some((index, &value)),
        None,
    )?;
    replace_item_xml_from_parts(xml, item_xml, limits, layout)
}

fn remove_item_from_parts(
    xml: &[u8],
    index: usize,
    limits: Limits,
    properties: &Properties,
    layout: &Layout,
) -> Result<Vec<u8>> {
    let Some(item_list) = properties.item_list.as_ref() else {
        return Err(invalid("form-control item removal index is out of bounds"));
    };
    if index >= item_list.items().len() {
        return Err(invalid("form-control item removal index is out of bounds"));
    }
    validate_item_edit(properties, true, item_list.items().len() - 1, limits)?;
    if let Some(list) = layout.item_list.as_ref()
        && !list.empty
        && !list.items.is_empty()
    {
        return splice_nonempty_item_list_edit(
            xml,
            list,
            layout.root_prefix.as_deref(),
            limits,
            None,
            Some(index),
        );
    }
    let item_xml = write_items_with_source_edit(
        layout.root_prefix.as_deref(),
        item_list.items(),
        layout
            .item_list
            .as_ref()
            .ok_or_else(|| invalid("form-control item-list layout missing"))?,
        xml,
        limits,
        None,
        Some(index),
    )?;
    replace_item_xml_from_parts(xml, item_xml, limits, layout)
}

#[derive(Clone, Debug)]
struct Inspected {
    properties: Properties,
    layout: Layout,
}

#[derive(Clone, Debug)]
struct Layout {
    root_open: Range<usize>,
    root_attribute_insertion: Option<usize>,
    root_child_insertion: Option<usize>,
    root_prefix: Option<Box<[u8]>>,
    attributes: Vec<AttributeLayout>,
    item_list: Option<ItemListLayout>,
    root_extension: Option<Range<usize>>,
}

#[derive(Clone, Debug)]
struct AttributeLayout {
    field: Option<ScalarField>,
    token: Range<usize>,
    value: Range<usize>,
}

#[derive(Clone, Debug)]
struct ItemListLayout {
    span: Range<usize>,
    items: Vec<ItemLayout>,
    first_content: Option<usize>,
    extension: Option<Range<usize>>,
    empty: bool,
}

#[derive(Clone, Debug)]
struct ItemLayout {
    span: Range<usize>,
    opaque: bool,
}

fn owned_source(xml: &[u8], limits: Limits) -> Result<SourcePayload> {
    if xml.len() > limits.max_part_bytes() {
        return Err(limit(
            "source part bytes",
            xml.len(),
            limits.max_part_bytes(),
        ));
    }
    if xml.len() > limits.max_retained_bytes() {
        return Err(limit(
            "retained source bytes",
            xml.len(),
            limits.max_retained_bytes(),
        ));
    }
    let mut source_bytes = Vec::new();
    source_bytes
        .try_reserve_exact(xml.len())
        .map_err(|source| allocation("form-control source bytes", source))?;
    source_bytes.extend_from_slice(xml);
    Ok(SourcePayload::Owned(Arc::new(source_bytes)))
}

fn inspect_with_limits_owned(
    xml: &[u8],
    limits: Limits,
    source: Option<SourcePayload>,
    context: Option<&ExecutionContext>,
) -> Result<Inspected> {
    if xml.len() > limits.max_part_bytes() {
        return Err(limit(
            "source part bytes",
            xml.len(),
            limits.max_part_bytes(),
        ));
    }
    if let Some(context) = context {
        context.check().map_err(map_execution_error)?;
    }
    validate_xml_source_characters(xml)?;
    if source
        .as_ref()
        .is_some_and(|source| source.len() != xml.len())
    {
        return Err(invalid("form-control source payload length mismatch"));
    }
    if source.is_some() && xml.len() > limits.max_retained_bytes() {
        return Err(limit(
            "retained source bytes",
            xml.len(),
            limits.max_retained_bytes(),
        ));
    }
    let mut retained = RetainedBudget::new(limits.max_retained_bytes(), context.cloned());
    if source.is_some() {
        // The source payload is already owned by the caller's source cache or
        // by the standalone `Owned` copy. It contributes to the leaf's local
        // retained cap, but managed source bytes must not be charged a second
        // time to the shared execution budget.
        retained.charge_source(xml.len(), "retained source bytes")?;
    }
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut events = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut root_open = None::<Range<usize>>;
    let mut root_attribute_insertion = None;
    let mut root_close_opening = None;
    let mut root_child_insertion = None;
    let mut root_prefix = None::<Box<[u8]>>;
    let mut root_namespaces = Vec::<NamespaceBinding>::new();
    let mut namespace_context = None::<Arc<[NamespaceBinding]>>;
    let mut lexical_attributes = Vec::<LexicalAttribute>::new();
    let mut attributes = Vec::<AttributeLayout>::new();
    let mut item_list = None::<ItemListLayout>;
    let mut item_list_namespace_context = None::<Arc<[NamespaceBinding]>>;
    let mut item_values = Vec::<Item>::new();
    let mut item_list_unknown_attributes = Vec::<OpaqueAttribute>::new();
    let mut item_list_extension = None::<OpaqueXml>;
    let mut root_extension = None::<OpaqueXml>;
    let mut item_list_extension_span = None::<Range<usize>>;
    let mut root_extension_span = None::<Range<usize>>;
    let mut opaque_bytes = 0usize;
    // Attribute/entity decoding is transient parser scratch.  It is bounded
    // separately from the retained model budget and from the candidate
    // output ceiling; a temporary URI decode must not consume output quota.
    let mut scratch = ScratchBudget::new(limits.max_opaque_bytes(), context.cloned());
    let mut open_item = None::<OpenItem>;
    let mut open_item_list = false;
    let mut open_opaque = None::<OpenOpaque>;
    let mut child_rank = 0u8;
    let mut item_child_rank = 0u8;
    retained.charge(PROPERTIES_STORAGE_BYTES, "form-control properties backing")?;
    retained.charge(
        2usize
            .checked_mul(SHARED_VEC_STORAGE_BYTES)
            .ok_or_else(|| invalid("form-control shared vector storage overflow"))?,
        "form-control shared vector handles",
    )?;
    let mut properties = Properties::new();

    loop {
        retained.check_execution()?;
        retained.charge_work(1)?;
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("form-control XML event count overflow"))?;
        if events > limits.max_events() {
            return Err(limit("XML event count", events, limits.max_events()));
        }
        let start = position(&reader, origin)?;
        let event = reader.read_event().map_err(xml_error)?;
        let end = position(&reader, origin)?;
        let resolver = reader.resolver();

        if let Some(opaque) = open_opaque.as_mut() {
            match event {
                Event::Start(element) => {
                    check_opaque_attribute_budget(&element, limits)?;
                    if depth == opaque.base_depth + 1 {
                        opaque.child_seen = true;
                        opaque.child_open_depth = Some(depth + 1);
                        if !validate_extension_element(
                            &element,
                            resolver,
                            reader.decoder(),
                            limits,
                            &mut scratch,
                        )? {
                            opaque.diagnostic = true;
                        }
                    }
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("form-control XML depth overflow"))?;
                    check_depth(depth, limits)?;
                },
                Event::Empty(element) => {
                    check_opaque_attribute_budget(&element, limits)?;
                    if depth == opaque.base_depth + 1 {
                        opaque.child_seen = true;
                        if !validate_extension_element(
                            &element,
                            resolver,
                            reader.decoder(),
                            limits,
                            &mut scratch,
                        )? {
                            opaque.diagnostic = true;
                        }
                    } else if depth <= opaque.base_depth {
                        opaque.diagnostic = true;
                    }
                    let empty_depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("form-control XML depth overflow"))?;
                    check_depth(empty_depth, limits)?;
                },
                Event::End(_) => {
                    if opaque.child_open_depth == Some(depth) {
                        opaque.child_open_depth = None;
                    } else if depth == opaque.base_depth + 1 && !opaque.child_seen {
                        // An empty extension list is valid.  Whitespace and
                        // comments do not create a child, so no diagnostic is
                        // needed here.
                    }
                    if depth == opaque.base_depth + 1 {
                        let span = opaque.start..end;
                        charge_opaque_range(&mut opaque_bytes, &span, limits)?;
                        let retained = source_for_range(
                            &source,
                            xml,
                            span.clone(),
                            limits,
                            &mut retained,
                            "opaque extension",
                        )?;
                        if let Some(retained) = retained {
                            let opaque_namespace_context = match opaque.owner {
                                OpaqueOwner::Root => namespace_context.as_ref(),
                                OpaqueOwner::ItemList => item_list_namespace_context.as_ref(),
                            }
                            .ok_or_else(|| invalid("form-control namespace context missing"))?;
                            let value = OpaqueXml::from_range_with_diagnostic(
                                retained.source,
                                retained.range,
                                Arc::clone(opaque_namespace_context),
                                opaque.diagnostic,
                            )?;
                            if value.byte_len() > limits.max_opaque_bytes() {
                                return Err(limit(
                                    "opaque extension bytes",
                                    value.byte_len(),
                                    limits.max_opaque_bytes(),
                                ));
                            }
                            match opaque.owner {
                                OpaqueOwner::Root => root_extension = Some(value),
                                OpaqueOwner::ItemList => item_list_extension = Some(value),
                            }
                        } else {
                            match opaque.owner {
                                OpaqueOwner::Root => root_extension_span = Some(span),
                                OpaqueOwner::ItemList => item_list_extension_span = Some(span),
                            }
                        }
                        open_opaque = None;
                    }
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("form-control XML depth underflow"))?;
                },
                Event::Text(text) => {
                    if depth == opaque.base_depth + 1 && !is_xml_whitespace(text.as_ref()) {
                        opaque.diagnostic = true;
                    }
                },
                Event::CData(_) => {
                    if depth == opaque.base_depth + 1 {
                        opaque.diagnostic = true;
                    }
                },
                Event::GeneralRef(reference) => {
                    validate_general_reference(&reference)?;
                },
                Event::Eof => return Err(invalid("unterminated opaque extension list")),
                Event::DocType(_) => {
                    return Err(invalid("DOCTYPE is not admitted in formControlPr"));
                },
                _ => {},
            }
            continue;
        }

        match event {
            Event::Start(element) => {
                let (namespace, local) = resolve_element(resolver, element.name())?;
                if depth == 0 {
                    if root_seen
                        || root_closed
                        || !namespace_uri_matches(namespace, FORM_CONTROL_NAMESPACE.as_bytes())
                        || local != ROOT
                    {
                        return Err(invalid(
                            "form-control part requires one x14:formControlPr root",
                        ));
                    }
                    root_seen = true;
                    root_open = Some(start..end);
                    root_attribute_insertion = Some(root_close_insertion(xml, start, end)?);
                    root_prefix = element_prefix(element.name(), &mut retained)?;
                    parse_root_attributes(
                        &element,
                        resolver,
                        reader.decoder(),
                        SourceElement {
                            source: xml,
                            start,
                            end,
                        },
                        limits,
                        &mut retained,
                        &mut properties,
                        &mut attributes,
                        &mut lexical_attributes,
                        &mut root_namespaces,
                        &mut opaque_bytes,
                    )?;
                    namespace_context = Some(finish_root_namespace_context(
                        &mut root_namespaces,
                        &mut retained,
                    )?);
                    depth = 1;
                    check_depth(depth, limits)?;
                } else if depth == 1 {
                    if root_child_insertion.is_none() {
                        root_child_insertion = Some(start);
                    }
                    if !namespace_uri_matches(namespace, FORM_CONTROL_NAMESPACE.as_bytes()) {
                        return Err(invalid("form-control root child uses a foreign namespace"));
                    }
                    match local {
                        ITEM_LIST => {
                            if child_rank > 1 || item_list.is_some() {
                                return Err(invalid(
                                    "duplicate or out-of-order form-control itemLst",
                                ));
                            }
                            child_rank = 1;
                            item_child_rank = 0;
                            item_list_unknown_attributes.clear();
                            parse_item_list_attributes(
                                &element,
                                resolver,
                                SourceElement {
                                    source: xml,
                                    start,
                                    end,
                                },
                                limits,
                                &mut retained,
                                &mut opaque_bytes,
                                &mut item_list_unknown_attributes,
                            )?;
                            let root_namespace_context = namespace_context
                                .as_ref()
                                .ok_or_else(|| invalid("form-control namespace context missing"))?;
                            item_list_namespace_context = Some(namespace_context_for_element(
                                &element,
                                reader.decoder(),
                                root_namespace_context,
                                limits,
                                &mut retained,
                            )?);
                            open_item_list = true;
                            item_list = Some(ItemListLayout {
                                span: start..start,
                                items: Vec::new(),
                                first_content: None,
                                extension: None,
                                empty: false,
                            });
                            depth = 2;
                            check_depth(depth, limits)?;
                        },
                        EXT_LIST => {
                            if child_rank > 2
                                || root_extension.is_some()
                                || root_extension_span.is_some()
                            {
                                return Err(invalid(
                                    "duplicate or out-of-order form-control extLst",
                                ));
                            }
                            child_rank = 2;
                            let diagnostic = extension_list_container_diagnostic(&element, limits)?;
                            open_opaque = Some(OpenOpaque {
                                owner: OpaqueOwner::Root,
                                start,
                                base_depth: depth,
                                child_seen: false,
                                child_open_depth: None,
                                diagnostic,
                            });
                            depth = 2;
                            check_depth(depth, limits)?;
                        },
                        _ => return Err(invalid("unexpected form-control root child")),
                    }
                } else if open_item_list && depth == 2 {
                    if !namespace_uri_matches(namespace, FORM_CONTROL_NAMESPACE.as_bytes()) {
                        return Err(invalid(
                            "form-control itemLst child uses a foreign namespace",
                        ));
                    }
                    match local {
                        ITEM => {
                            if item_child_rank == 3 || open_item.is_some() {
                                return Err(invalid("itemLst contains an item after extLst"));
                            }
                            if item_values.len() >= limits.max_items() {
                                return Err(limit(
                                    "item count",
                                    item_values.len().saturating_add(1),
                                    limits.max_items(),
                                ));
                            }
                            // Item source spans retain their own namespace
                            // declarations.  The inherited item-list scope
                            // is shared by the private opaque owner instead
                            // of flattening every item's local scope.
                            let item_namespace_context =
                                Arc::clone(item_list_namespace_context.as_ref().ok_or_else(
                                    || invalid("form-control itemLst namespace context missing"),
                                )?);
                            reserve_retained_slots(
                                &mut item_values,
                                1,
                                &mut retained,
                                "form-control items",
                            )?;
                            if let Some(layout) = item_list.as_mut() {
                                reserve_retained_slots(
                                    &mut layout.items,
                                    1,
                                    &mut retained,
                                    "form-control item layout",
                                )?;
                            }
                            let (item, opaque) = parse_item_attributes(
                                &element,
                                resolver,
                                reader.decoder(),
                                limits,
                                &mut retained,
                            )?;
                            item_values.push(item);
                            if let Some(layout) = item_list.as_mut() {
                                if layout.first_content.is_none() {
                                    layout.first_content = Some(start);
                                }
                                layout.items.push(ItemLayout {
                                    span: start..end,
                                    opaque,
                                });
                            }
                            let item_index = item_values.len().saturating_sub(1);
                            open_item = Some(OpenItem {
                                index: item_index,
                                start,
                                depth: 2,
                                opaque,
                                namespace_context: item_namespace_context,
                            });
                            depth = 3;
                            check_depth(depth, limits)?;
                        },
                        EXT_LIST => {
                            if item_child_rank == 3 {
                                return Err(invalid("duplicate itemLst extLst"));
                            }
                            item_child_rank = 3;
                            if let Some(layout) = item_list.as_mut() {
                                if layout.first_content.is_none() {
                                    layout.first_content = Some(start);
                                }
                            }
                            let diagnostic = extension_list_container_diagnostic(&element, limits)?;
                            open_opaque = Some(OpenOpaque {
                                owner: OpaqueOwner::ItemList,
                                start,
                                base_depth: depth,
                                child_seen: false,
                                child_open_depth: None,
                                diagnostic,
                            });
                            depth = 3;
                            check_depth(depth, limits)?;
                        },
                        _ => return Err(invalid("unexpected itemLst child")),
                    }
                } else if open_item.is_some() {
                    return Err(invalid("x14:item must be empty"));
                } else {
                    return Err(invalid("unexpected nested form-control element"));
                }
            },
            Event::Empty(element) => {
                let (namespace, local) = resolve_element(resolver, element.name())?;
                if depth == 0 {
                    if root_seen
                        || root_closed
                        || !namespace_uri_matches(namespace, FORM_CONTROL_NAMESPACE.as_bytes())
                        || local != ROOT
                    {
                        return Err(invalid(
                            "form-control part requires one x14:formControlPr root",
                        ));
                    }
                    root_seen = true;
                    root_closed = true;
                    root_open = Some(start..end);
                    root_attribute_insertion = Some(root_close_insertion(xml, start, end)?);
                    root_prefix = element_prefix(element.name(), &mut retained)?;
                    check_depth(1, limits)?;
                    parse_root_attributes(
                        &element,
                        resolver,
                        reader.decoder(),
                        SourceElement {
                            source: xml,
                            start,
                            end,
                        },
                        limits,
                        &mut retained,
                        &mut properties,
                        &mut attributes,
                        &mut lexical_attributes,
                        &mut root_namespaces,
                        &mut opaque_bytes,
                    )?;
                    namespace_context = Some(finish_root_namespace_context(
                        &mut root_namespaces,
                        &mut retained,
                    )?);
                } else if depth == 1 {
                    if root_child_insertion.is_none() {
                        root_child_insertion = Some(start);
                    }
                    if !namespace_uri_matches(namespace, FORM_CONTROL_NAMESPACE.as_bytes()) {
                        return Err(invalid("form-control root child uses a foreign namespace"));
                    }
                    match local {
                        ITEM_LIST => {
                            if child_rank > 1 || item_list.is_some() {
                                return Err(invalid(
                                    "duplicate or out-of-order form-control itemLst",
                                ));
                            }
                            child_rank = 1;
                            item_child_rank = 0;
                            item_list_unknown_attributes.clear();
                            parse_item_list_attributes(
                                &element,
                                resolver,
                                SourceElement {
                                    source: xml,
                                    start,
                                    end,
                                },
                                limits,
                                &mut retained,
                                &mut opaque_bytes,
                                &mut item_list_unknown_attributes,
                            )?;
                            let root_namespace_context = namespace_context
                                .as_ref()
                                .ok_or_else(|| invalid("form-control namespace context missing"))?;
                            item_list_namespace_context = Some(namespace_context_for_element(
                                &element,
                                reader.decoder(),
                                root_namespace_context,
                                limits,
                                &mut retained,
                            )?);
                            item_list = Some(ItemListLayout {
                                span: start..end,
                                items: Vec::new(),
                                first_content: Some(end.saturating_sub(2)),
                                extension: None,
                                empty: true,
                            });
                            check_depth(2, limits)?;
                        },
                        EXT_LIST => {
                            if child_rank > 2
                                || root_extension.is_some()
                                || root_extension_span.is_some()
                            {
                                return Err(invalid(
                                    "duplicate or out-of-order form-control extLst",
                                ));
                            }
                            child_rank = 2;
                            check_depth(2, limits)?;
                            let diagnostic = extension_list_container_diagnostic(&element, limits)?;
                            let span = start..end;
                            charge_opaque_range(&mut opaque_bytes, &span, limits)?;
                            let retained = source_for_range(
                                &source,
                                xml,
                                span.clone(),
                                limits,
                                &mut retained,
                                "opaque extension",
                            )?;
                            if let Some(retained) = retained {
                                let value = OpaqueXml::from_range_with_diagnostic(
                                    retained.source,
                                    retained.range,
                                    Arc::clone(namespace_context.as_ref().ok_or_else(|| {
                                        invalid("form-control namespace context missing")
                                    })?),
                                    diagnostic,
                                )?;
                                if value.byte_len() > limits.max_opaque_bytes() {
                                    return Err(limit(
                                        "opaque extension bytes",
                                        value.byte_len(),
                                        limits.max_opaque_bytes(),
                                    ));
                                }
                                root_extension = Some(value);
                            } else {
                                root_extension_span = Some(span);
                            }
                        },
                        _ => return Err(invalid("unexpected form-control root child")),
                    }
                } else if open_item_list && depth == 2 {
                    if !namespace_uri_matches(namespace, FORM_CONTROL_NAMESPACE.as_bytes()) {
                        return Err(invalid(
                            "form-control itemLst child uses a foreign namespace",
                        ));
                    }
                    match local {
                        ITEM => {
                            if item_child_rank == 3 {
                                return Err(invalid("itemLst contains an item after extLst"));
                            }
                            if item_values.len() >= limits.max_items() {
                                return Err(limit(
                                    "item count",
                                    item_values.len().saturating_add(1),
                                    limits.max_items(),
                                ));
                            }
                            check_depth(3, limits)?;
                            // As with the start-tag path, preserve any item
                            // local declarations in the source span and
                            // share the inherited item-list scope.
                            let item_namespace_context =
                                Arc::clone(item_list_namespace_context.as_ref().ok_or_else(
                                    || invalid("form-control itemLst namespace context missing"),
                                )?);
                            reserve_retained_slots(
                                &mut item_values,
                                1,
                                &mut retained,
                                "form-control items",
                            )?;
                            if let Some(layout) = item_list.as_mut() {
                                reserve_retained_slots(
                                    &mut layout.items,
                                    1,
                                    &mut retained,
                                    "form-control item layout",
                                )?;
                            }
                            let (mut item, opaque) = parse_item_attributes(
                                &element,
                                resolver,
                                reader.decoder(),
                                limits,
                                &mut retained,
                            )?;
                            if opaque {
                                let span = start..end;
                                charge_opaque_range(&mut opaque_bytes, &span, limits)?;
                                let item_source = source_for_range(
                                    &source,
                                    xml,
                                    span,
                                    limits,
                                    &mut retained,
                                    "item source",
                                )?;
                                if let Some(item_source) = item_source {
                                    item.raw = Some(OpaqueXml::from_range(
                                        item_source.source,
                                        item_source.range,
                                        Arc::clone(&item_namespace_context),
                                    )?);
                                }
                            }
                            item_values.push(item);
                            if let Some(layout) = item_list.as_mut() {
                                if layout.first_content.is_none() {
                                    layout.first_content = Some(start);
                                }
                                layout.items.push(ItemLayout {
                                    span: start..end,
                                    opaque,
                                });
                            }
                        },
                        EXT_LIST => {
                            if item_child_rank == 3 {
                                return Err(invalid("duplicate itemLst extLst"));
                            }
                            item_child_rank = 3;
                            check_depth(3, limits)?;
                            let diagnostic = extension_list_container_diagnostic(&element, limits)?;
                            if let Some(layout) = item_list.as_mut() {
                                if layout.first_content.is_none() {
                                    layout.first_content = Some(start);
                                }
                            }
                            let span = start..end;
                            charge_opaque_range(&mut opaque_bytes, &span, limits)?;
                            let retained = source_for_range(
                                &source,
                                xml,
                                span.clone(),
                                limits,
                                &mut retained,
                                "opaque extension",
                            )?;
                            if let Some(retained) = retained {
                                let value = OpaqueXml::from_range_with_diagnostic(
                                    retained.source,
                                    retained.range,
                                    Arc::clone(item_list_namespace_context.as_ref().ok_or_else(
                                        || invalid("form-control namespace context missing"),
                                    )?),
                                    diagnostic,
                                )?;
                                if value.byte_len() > limits.max_opaque_bytes() {
                                    return Err(limit(
                                        "opaque extension bytes",
                                        value.byte_len(),
                                        limits.max_opaque_bytes(),
                                    ));
                                }
                                item_list_extension = Some(value);
                            } else if let Some(layout) = item_list.as_mut() {
                                layout.extension = Some(span);
                            }
                        },
                        _ => return Err(invalid("unexpected itemLst child")),
                    }
                } else if open_item.is_some() {
                    return Err(invalid("x14:item must be empty"));
                } else {
                    return Err(invalid("unexpected nested form-control element"));
                }
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid("unexpected form-control end element"));
                }
                if let Some(item) = open_item.take() {
                    if depth != item.depth + 1 {
                        return Err(invalid("mismatched form-control item depth"));
                    }
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("form-control XML depth underflow"))?;
                    if let Some(layout) = item_list.as_mut() {
                        if let Some(last) = layout.items.last_mut() {
                            last.span.end = end;
                            last.opaque |= item.opaque;
                        }
                    }
                    if item.opaque {
                        if let Some(value) = item_values.get_mut(item.index) {
                            let span = item.start..end;
                            charge_opaque_range(&mut opaque_bytes, &span, limits)?;
                            let item_source = source_for_range(
                                &source,
                                xml,
                                span,
                                limits,
                                &mut retained,
                                "item source",
                            )?;
                            if let Some(item_source) = item_source {
                                value.raw = Some(OpaqueXml::from_range(
                                    item_source.source,
                                    item_source.range,
                                    Arc::clone(&item.namespace_context),
                                )?);
                            }
                        }
                    }
                } else if open_item_list && depth == 2 {
                    open_item_list = false;
                    if let Some(layout) = item_list.as_mut() {
                        layout.span.end = end;
                        layout.empty = false;
                    }
                    depth = 1;
                } else if depth == 1 {
                    root_closed = true;
                    root_close_opening = Some(start);
                    if root_child_insertion.is_none() {
                        root_child_insertion = Some(start);
                    }
                    depth = 0;
                } else {
                    return Err(invalid("unexpected form-control closing element"));
                }
            },
            Event::Text(text) => {
                if open_item.is_some() {
                    if !is_xml_whitespace(text.as_ref()) {
                        return Err(invalid("x14:item cannot contain character content"));
                    }
                    if let Some(item) = open_item.as_mut() {
                        item.opaque = true;
                    }
                } else if !is_xml_whitespace(text.as_ref()) {
                    return Err(invalid(
                        "form-control XML has non-whitespace top-level text",
                    ));
                }
            },
            Event::CData(_) => {
                if open_item.is_some() {
                    return Err(invalid("x14:item cannot contain CDATA"));
                }
                return Err(invalid("CDATA is not admitted outside opaque extLst"));
            },
            Event::Comment(_) | Event::PI(_) => {
                if let Some(item) = open_item.as_mut() {
                    item.opaque = true;
                }
            },
            Event::Decl(_) => {},
            Event::DocType(_) => return Err(invalid("DOCTYPE is not admitted in formControlPr")),
            Event::GeneralRef(_) => return Err(invalid("entity references are not admitted")),
            Event::Eof => break,
        }
    }
    if !root_seen || !root_closed || depth != 0 || open_item.is_some() || open_opaque.is_some() {
        return Err(invalid("form-control XML does not contain one closed root"));
    }
    if item_values.len() > limits.max_items() {
        return Err(limit("item count", item_values.len(), limits.max_items()));
    }
    if let Some(layout) = item_list.as_mut() {
        if layout.first_content.is_none() {
            if !layout.empty {
                let name_len = qname_len(root_prefix.as_deref(), ITEM_LIST)?;
                let closing_len = name_len
                    .checked_add(3)
                    .ok_or_else(|| invalid("itemLst closing span overflow"))?;
                layout.first_content = layout.span.end.checked_sub(closing_len);
            }
        }
    }
    if let Some(layout) = item_list.as_mut() {
        layout.extension = layout.extension.take().or(item_list_extension_span);
    }
    let parsed_item_list = item_list.as_ref().map(|_| {
        ItemList::from_parts_with_namespace(
            item_values,
            item_list_extension,
            item_list_namespace_context
                .clone()
                .unwrap_or_else(|| Arc::from([])),
            item_list_unknown_attributes,
        )
    });
    replace_shared_option(
        &mut properties.item_list,
        parsed_item_list,
        &mut retained,
        "form-control item-list backing",
    )?;
    replace_shared_option(
        &mut properties.root_extension_list,
        root_extension,
        &mut retained,
        "form-control opaque root backing",
    )?;
    let namespace_context =
        namespace_context.ok_or_else(|| invalid("form-control namespace context missing"))?;
    replace_shared_vec(
        &mut properties.lexical_attributes,
        lexical_attributes,
        &mut retained,
        "form-control lexical attribute backing",
    )?;
    if source.is_some() {
        properties.mark_source(
            source.ok_or_else(|| invalid("form-control source ownership missing"))?,
            namespace_context,
        );
    } else {
        properties.mark_metadata(namespace_context);
    }
    properties.validate_parse_with_limits(limits)?;
    properties.mark_retained_lease(retained.finish_lease()?);
    let root_open = root_open.ok_or_else(|| invalid("form-control root source span missing"))?;
    if root_child_insertion.is_none() {
        root_child_insertion = root_close_opening.or_else(|| Some(root_open.end.saturating_sub(1)));
    }
    Ok(Inspected {
        properties,
        layout: Layout {
            root_open,
            root_attribute_insertion,
            root_child_insertion,
            root_prefix,
            attributes,
            item_list,
            root_extension: root_extension_span,
        },
    })
}

#[derive(Clone, Debug)]
struct OpenItem {
    index: usize,
    start: usize,
    depth: usize,
    opaque: bool,
    namespace_context: Arc<[NamespaceBinding]>,
}

#[derive(Clone, Copy, Debug)]
struct OpenOpaque {
    owner: OpaqueOwner,
    start: usize,
    base_depth: usize,
    child_seen: bool,
    child_open_depth: Option<usize>,
    diagnostic: bool,
}

#[derive(Clone, Copy, Debug)]
enum OpaqueOwner {
    Root,
    ItemList,
}

fn position(reader: &NsReader<&[u8]>, origin: ReaderOrigin) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| invalid("form-control source position overflow"))
}

struct RetainedRange {
    source: SourcePayload,
    range: Range<usize>,
}

#[derive(Clone, Copy, Debug)]
struct SourceElement<'a> {
    source: &'a [u8],
    start: usize,
    end: usize,
}

struct RetainedBudget {
    used: usize,
    maximum: usize,
    temporary_used: usize,
    namespace_temporary_used: usize,
    lease: Option<RetainedLeaseBuilder>,
}

struct RetainedLeaseBuilder {
    context: ExecutionContext,
    reservation: Option<Reservation>,
    temporary: Option<Reservation>,
    namespace_temporary: Option<Reservation>,
}

fn reserve_retained_slots<T>(
    values: &mut Vec<T>,
    additional: usize,
    retained: &mut RetainedBudget,
    resource: &'static str,
) -> Result<()> {
    retained.reserve_vec_growth(values, additional, resource)
}

fn reserve_namespace_slots<T>(
    values: &mut Vec<T>,
    additional: usize,
    retained: &mut RetainedBudget,
    resource: &'static str,
) -> Result<()> {
    let required = values
        .len()
        .checked_add(additional)
        .ok_or_else(|| invalid("form-control namespace binding count overflow"))?;
    let old_capacity = values.capacity();
    if required <= old_capacity {
        return Ok(());
    }
    let doubled = old_capacity
        .max(1)
        .checked_mul(2)
        .ok_or_else(|| invalid("form-control namespace binding capacity overflow"))?;
    let target = required.max(doubled);
    let target_bytes = target
        .checked_mul(size_of::<T>())
        .ok_or_else(|| invalid("form-control namespace binding storage overflow"))?;
    retained.check_peak(target_bytes, resource)?;
    retained.reserve_namespace_temporary(target_bytes, resource)?;
    if let Err(source) = values.try_reserve_exact(target - old_capacity) {
        retained.release_namespace_temporary();
        return Err(allocation(resource, source));
    }
    let actual_bytes = values
        .capacity()
        .checked_mul(size_of::<T>())
        .ok_or_else(|| invalid("form-control namespace binding storage overflow"))?;
    if values.capacity() < required {
        retained.release_namespace_temporary();
        return Err(invalid(
            "form-control namespace binding capacity below request",
        ));
    }
    if actual_bytes > target_bytes {
        let extra = actual_bytes
            .checked_sub(target_bytes)
            .ok_or_else(|| invalid("form-control namespace binding storage underflow"))?;
        retained.check_peak(extra, resource)?;
        if let Err(error) = retained.reserve_namespace_temporary(extra, resource) {
            retained.release_namespace_temporary();
            return Err(error);
        }
    }
    // The Vec's spare capacity is temporary: the final root context is an
    // Arc slice and retains exactly its length.  Rebuild the temporary lease
    // from the allocator's actual capacity before the next growth.
    retained.release_namespace_temporary();
    retained.reserve_namespace_temporary(actual_bytes, resource)
}

/// Reserve a fresh parser-temporary vector and keep its actual allocation in
/// the temporary peak ledger.  The caller must release the temporary ledger
/// after dropping the vector, then charge only the final retained allocation.
fn reserve_temporary_slots<T>(
    values: &mut Vec<T>,
    required: usize,
    retained: &mut RetainedBudget,
    resource: &'static str,
) -> Result<usize> {
    let old_capacity = values.capacity();
    if required <= old_capacity {
        return old_capacity
            .checked_mul(size_of::<T>())
            .ok_or_else(|| invalid("form-control temporary slot storage overflow"));
    }
    let target_bytes = required
        .checked_mul(size_of::<T>())
        .ok_or_else(|| invalid("form-control temporary slot storage overflow"))?;
    retained.check_peak(target_bytes, resource)?;
    retained.reserve_temporary(target_bytes)?;
    if let Err(source) = values.try_reserve_exact(required - old_capacity) {
        retained.release_temporary();
        return Err(allocation(resource, source));
    }
    let actual_capacity = values.capacity();
    if actual_capacity < required {
        retained.release_temporary();
        return Err(invalid(
            "form-control temporary vector capacity below request",
        ));
    }
    let actual_bytes = actual_capacity
        .checked_mul(size_of::<T>())
        .ok_or_else(|| invalid("form-control temporary slot storage overflow"))?;
    if actual_bytes > target_bytes {
        let extra = actual_bytes
            .checked_sub(target_bytes)
            .ok_or_else(|| invalid("form-control temporary slot storage underflow"))?;
        retained.check_peak(extra, resource)?;
        if let Err(error) = retained.reserve_temporary(extra) {
            retained.release_temporary();
            return Err(error);
        }
    }
    Ok(actual_bytes)
}

impl RetainedBudget {
    fn new(maximum: usize, context: Option<ExecutionContext>) -> Self {
        Self {
            used: 0,
            maximum,
            temporary_used: 0,
            namespace_temporary_used: 0,
            lease: context.map(|context| RetainedLeaseBuilder {
                context,
                reservation: None,
                temporary: None,
                namespace_temporary: None,
            }),
        }
    }

    fn charge(&mut self, amount: usize, resource: &'static str) -> Result<()> {
        self.charge_inner(amount, resource, true)
    }

    fn charge_source(&mut self, amount: usize, resource: &'static str) -> Result<()> {
        self.charge_inner(amount, resource, false)
    }

    fn reserve_vec_growth<T>(
        &mut self,
        values: &mut Vec<T>,
        additional: usize,
        resource: &'static str,
    ) -> Result<()> {
        let required = values
            .len()
            .checked_add(additional)
            .ok_or_else(|| invalid("form-control retained slot count overflow"))?;
        let old_capacity = values.capacity();
        if required <= old_capacity {
            return Ok(());
        }
        let doubled = old_capacity
            .max(1)
            .checked_mul(2)
            .ok_or_else(|| invalid("form-control retained slot capacity overflow"))?;
        let target = required.max(doubled);
        let old_bytes = old_capacity
            .checked_mul(size_of::<T>())
            .ok_or_else(|| invalid("form-control retained slot storage overflow"))?;
        let target_bytes = target
            .checked_mul(size_of::<T>())
            .ok_or_else(|| invalid("form-control retained slot storage overflow"))?;

        // The old allocation is already part of `used`. Reserve the complete
        // requested new allocation separately so the shared execution budget
        // and local retained cap see the old+new reallocation peak.
        self.check_peak(target_bytes, resource)?;
        self.reserve_temporary(target_bytes)?;
        if let Err(source) = values.try_reserve_exact(target - old_capacity) {
            self.release_temporary();
            return Err(allocation(resource, source));
        }
        let actual_capacity = values.capacity();
        if actual_capacity < required {
            self.release_temporary();
            return Err(invalid(
                "form-control retained vector capacity below request",
            ));
        }
        let actual_bytes = actual_capacity
            .checked_mul(size_of::<T>())
            .ok_or_else(|| invalid("form-control retained slot storage overflow"))?;
        if actual_bytes > target_bytes {
            let extra = actual_bytes
                .checked_sub(target_bytes)
                .ok_or_else(|| invalid("form-control retained slot storage underflow"))?;
            self.check_peak(actual_bytes, resource)?;
            self.reserve_temporary(extra)?;
        }
        self.release_temporary();

        // `try_reserve_exact` has released the old allocation before it
        // returns. Retain only the actual final capacity in the live ledger;
        // reconciling against `capacity()` avoids trusting the requested
        // target when an allocator grows beyond it.
        let additional_bytes = actual_bytes
            .checked_sub(old_bytes)
            .ok_or_else(|| invalid("form-control retained slot storage underflow"))?;
        self.charge(additional_bytes, resource)
    }

    fn check_peak(&self, additional: usize, resource: &'static str) -> Result<()> {
        let peak = self
            .used
            .checked_add(self.temporary_used)
            .ok_or_else(|| invalid("form-control retained allocation length overflow"))?
            .checked_add(self.namespace_temporary_used)
            .ok_or_else(|| invalid("form-control retained allocation length overflow"))?
            .checked_add(additional)
            .ok_or_else(|| invalid("form-control retained allocation length overflow"))?;
        if peak > self.maximum {
            return Err(limit(resource, peak, self.maximum));
        }
        Ok(())
    }

    fn reserve_temporary(&mut self, amount: usize) -> Result<()> {
        let temporary_used = self
            .temporary_used
            .checked_add(amount)
            .ok_or_else(|| invalid("form-control temporary allocation length overflow"))?;
        let peak = self
            .used
            .checked_add(temporary_used)
            .and_then(|value| value.checked_add(self.namespace_temporary_used))
            .ok_or_else(|| invalid("form-control retained allocation length overflow"))?;
        if peak > self.maximum {
            return Err(limit(
                "form-control temporary retained allocation",
                peak,
                self.maximum,
            ));
        }
        if let Some(lease) = self.lease.as_mut() {
            lease.reserve_temporary(amount)?;
        }
        self.temporary_used = temporary_used;
        Ok(())
    }

    fn release_temporary(&mut self) {
        self.temporary_used = 0;
        if let Some(lease) = self.lease.as_mut() {
            lease.release_temporary();
        }
    }

    fn reserve_namespace_temporary(&mut self, amount: usize, resource: &'static str) -> Result<()> {
        let namespace_temporary_used = self
            .namespace_temporary_used
            .checked_add(amount)
            .ok_or_else(|| invalid("form-control namespace temporary storage overflow"))?;
        let peak = self
            .used
            .checked_add(self.temporary_used)
            .and_then(|value| value.checked_add(namespace_temporary_used))
            .ok_or_else(|| invalid("form-control retained allocation length overflow"))?;
        if peak > self.maximum {
            return Err(limit(resource, peak, self.maximum));
        }
        if let Some(lease) = self.lease.as_mut() {
            lease.reserve_namespace_temporary(amount)?;
        }
        self.namespace_temporary_used = namespace_temporary_used;
        Ok(())
    }

    fn release_namespace_temporary(&mut self) {
        self.namespace_temporary_used = 0;
        if let Some(lease) = self.lease.as_mut() {
            lease.release_namespace_temporary();
        }
    }

    fn charge_inner(
        &mut self,
        amount: usize,
        resource: &'static str,
        charge_execution: bool,
    ) -> Result<()> {
        let used = self
            .used
            .checked_add(amount)
            .ok_or_else(|| invalid("form-control retained allocation length overflow"))?;
        let peak = used
            .checked_add(self.temporary_used)
            .and_then(|value| value.checked_add(self.namespace_temporary_used))
            .ok_or_else(|| invalid("form-control retained allocation length overflow"))?;
        if peak > self.maximum {
            return Err(limit(resource, peak, self.maximum));
        }
        if charge_execution {
            if let Some(lease) = self.lease.as_mut() {
                lease.reserve(amount)?;
            }
        }
        self.used = used;
        Ok(())
    }

    fn check_execution(&self) -> Result<()> {
        self.lease
            .as_ref()
            .map_or(Ok(()), RetainedLeaseBuilder::check)
    }

    fn charge_work(&self, amount: usize) -> Result<()> {
        let Some(lease) = self.lease.as_ref() else {
            return Ok(());
        };
        let amount = u64::try_from(amount)
            .map_err(|_| invalid("form-control execution work charge overflow"))?;
        lease
            .context
            .consume(Resource::Work, amount)
            .map_err(map_execution_error)
    }

    fn finish_lease(mut self) -> Result<Option<Arc<Reservation>>> {
        let has_reservation = self
            .lease
            .as_ref()
            .is_some_and(|lease| lease.reservation.is_some());
        if !has_reservation {
            return Ok(None);
        }
        // `Arc::new` allocates an Arc header in addition to the reservation;
        // reserve that retained storage before publishing the lease handle.
        self.charge(
            RETAINED_LEASE_STORAGE_BYTES,
            "form-control retained execution lease",
        )?;
        let reservation = self
            .lease
            .as_mut()
            .and_then(|lease| lease.reservation.take())
            .ok_or_else(|| invalid("form-control retained execution lease disappeared"))?;
        Ok(Some(Arc::new(reservation)))
    }
}

impl RetainedLeaseBuilder {
    fn reserve(&mut self, amount: usize) -> Result<()> {
        if amount == 0 {
            return Ok(());
        }
        let amount = u64::try_from(amount)
            .map_err(|_| invalid("form-control execution memory charge overflow"))?;
        let reservation = self
            .context
            .reserve(Resource::Memory, amount)
            .map_err(map_execution_error)?;
        if let Some(existing) = self.reservation.as_mut() {
            if let Err(other) = existing.try_merge(reservation) {
                drop(other);
                return Err(invalid(
                    "form-control execution memory reservations used different budget chains",
                ));
            }
        } else {
            self.reservation = Some(reservation);
        }
        Ok(())
    }

    fn reserve_temporary(&mut self, amount: usize) -> Result<()> {
        if amount == 0 {
            return Ok(());
        }
        let amount = u64::try_from(amount)
            .map_err(|_| invalid("form-control execution temporary memory charge overflow"))?;
        let reservation = self
            .context
            .reserve(Resource::Memory, amount)
            .map_err(map_execution_error)?;
        if let Some(existing) = self.temporary.as_mut() {
            if let Err(other) = existing.try_merge(reservation) {
                drop(other);
                return Err(invalid(
                    "form-control temporary reservations used different budget chains",
                ));
            }
        } else {
            self.temporary = Some(reservation);
        }
        Ok(())
    }

    fn release_temporary(&mut self) {
        drop(self.temporary.take());
    }

    fn reserve_namespace_temporary(&mut self, amount: usize) -> Result<()> {
        if amount == 0 {
            return Ok(());
        }
        let amount = u64::try_from(amount)
            .map_err(|_| invalid("form-control namespace temporary memory charge overflow"))?;
        let reservation = self
            .context
            .reserve(Resource::Memory, amount)
            .map_err(map_execution_error)?;
        if let Some(existing) = self.namespace_temporary.as_mut() {
            if let Err(other) = existing.try_merge(reservation) {
                drop(other);
                return Err(invalid(
                    "form-control namespace temporary reservations used different budget chains",
                ));
            }
        } else {
            self.namespace_temporary = Some(reservation);
        }
        Ok(())
    }

    fn release_namespace_temporary(&mut self) {
        drop(self.namespace_temporary.take());
    }

    fn check(&self) -> Result<()> {
        self.context.check().map_err(map_execution_error)
    }
}

/// A short-lived parser scratch budget.
///
/// This is deliberately independent of `Limits::max_output_bytes`: output is
/// the size of the final candidate, while this counter bounds a temporary
/// decoded attribute value which is released before the next event.
struct ScratchBudget {
    used: usize,
    maximum: usize,
    context: Option<ExecutionContext>,
    reservation: Option<Reservation>,
}

impl ScratchBudget {
    fn new(maximum: usize, context: Option<ExecutionContext>) -> Self {
        Self {
            used: 0,
            maximum,
            context,
            reservation: None,
        }
    }

    fn charge(&mut self, amount: usize, resource: &'static str) -> Result<()> {
        self.used = self
            .used
            .checked_add(amount)
            .ok_or_else(|| invalid("form-control scratch allocation length overflow"))?;
        if self.used > self.maximum {
            return Err(limit(resource, self.used, self.maximum));
        }
        if let Some(context) = self.context.as_ref() {
            if self.reservation.is_some() {
                return Err(invalid(
                    "form-control scratch allocations overlap unexpectedly",
                ));
            }
            let amount = u64::try_from(amount)
                .map_err(|_| invalid("form-control scratch memory charge overflow"))?;
            self.reservation = Some(
                context
                    .reserve(Resource::Memory, amount)
                    .map_err(map_execution_error)?,
            );
        }
        Ok(())
    }

    fn release(&mut self, amount: usize) {
        self.used = self.used.saturating_sub(amount);
        drop(self.reservation.take());
    }
}

fn map_execution_error(error: ExecutionError) -> super::FormControlError {
    error.into()
}

fn replace_shared_option<T>(
    slot: &mut SharedOption<T>,
    value: Option<T>,
    retained: &mut RetainedBudget,
    resource: &'static str,
) -> Result<()> {
    if value.is_some() {
        retained.charge(
            ARC_HEADER_BYTES
                .checked_add(size_of::<T>())
                .ok_or_else(|| invalid("form-control shared option storage overflow"))?,
            resource,
        )?;
    }
    slot.replace(value);
    Ok(())
}

fn replace_shared_vec<T>(
    slot: &mut SharedVec<T>,
    value: Vec<T>,
    retained: &mut RetainedBudget,
    resource: &'static str,
) -> Result<()> {
    // Parsing starts with unique empty vector handles.  Replacing one of
    // those handles can move the completed Vec into the existing Arc without
    // allocating another Arc header, so the retained ledger must charge only
    // when a shared caller forces a new backing handle.  This keeps the
    // boundary charge equal to the live model while still precharging the
    // fallback allocation before publishing it.
    if !slot.is_unique() {
        retained.charge(SHARED_VEC_STORAGE_BYTES, resource)?;
    }
    slot.replace(value);
    Ok(())
}

fn source_for_range(
    source: &Option<SourcePayload>,
    xml: &[u8],
    range: Range<usize>,
    limits: Limits,
    retained: &mut RetainedBudget,
    resource: &'static str,
) -> Result<Option<RetainedRange>> {
    if range.start > range.end || range.end > xml.len() {
        return Err(invalid("form-control retained range lies outside source"));
    }
    let length = range.end - range.start;
    if length > limits.max_opaque_bytes() {
        return Err(limit(resource, length, limits.max_opaque_bytes()));
    }
    if let Some(source) = source {
        return Ok(Some(RetainedRange {
            source: source.clone(),
            range,
        }));
    }
    let _ = (retained, resource);
    Ok(None)
}

fn check_depth(depth: usize, limits: Limits) -> Result<()> {
    if depth > limits.max_depth() {
        return Err(limit("XML depth", depth, limits.max_depth()));
    }
    Ok(())
}

fn charge_opaque_range(total: &mut usize, range: &Range<usize>, limits: Limits) -> Result<()> {
    let length = range
        .end
        .checked_sub(range.start)
        .ok_or_else(|| invalid("form-control opaque range length underflow"))?;
    charge_opaque_length(total, length, limits, "aggregate opaque form-control bytes")
}

fn charge_opaque_length(
    total: &mut usize,
    length: usize,
    limits: Limits,
    resource: &'static str,
) -> Result<()> {
    *total = total
        .checked_add(length)
        .ok_or_else(|| invalid("form-control aggregate opaque bytes overflow"))?;
    if *total > limits.max_opaque_bytes() {
        return Err(limit(resource, *total, limits.max_opaque_bytes()));
    }
    Ok(())
}

fn xml_error(error: quick_xml::Error) -> super::FormControlError {
    invalid(format!("form-control XML error: {error}"))
}

fn validate_xml_source_characters(xml: &[u8]) -> Result<()> {
    let source = std::str::from_utf8(xml)
        .map_err(|error| invalid(format!("form-control XML is not UTF-8: {error}")))?;
    if source.chars().any(|character| !is_xml_10_char(character)) {
        return Err(invalid(
            "form-control XML contains an XML-invalid character",
        ));
    }
    Ok(())
}

fn validate_general_reference(reference: &BytesRef<'_>) -> Result<()> {
    let value = reference
        .decode()
        .map_err(|error| invalid(format!("invalid form-control entity reference: {error}")))?;
    let value = value.as_ref();
    if let Some(value) = value
        .strip_prefix("#x")
        .or_else(|| value.strip_prefix("#X"))
    {
        let codepoint = u32::from_str_radix(value, 16)
            .map_err(|_| invalid("invalid hexadecimal XML character reference"))?;
        let character = char::from_u32(codepoint)
            .ok_or_else(|| invalid("XML character reference is outside Unicode"))?;
        if !is_xml_10_char(character) {
            return Err(invalid("XML character reference is invalid in XML 1.0"));
        }
    } else if let Some(value) = value.strip_prefix('#') {
        let codepoint = value
            .parse::<u32>()
            .map_err(|_| invalid("invalid decimal XML character reference"))?;
        let character = char::from_u32(codepoint)
            .ok_or_else(|| invalid("XML character reference is outside Unicode"))?;
        if !is_xml_10_char(character) {
            return Err(invalid("XML character reference is invalid in XML 1.0"));
        }
    } else if !matches!(value, "amp" | "lt" | "gt" | "apos" | "quot") {
        return Err(invalid("undeclared XML entity reference is not admitted"));
    }
    Ok(())
}

fn resolve_element<'namespace, 'name>(
    resolver: &'namespace NamespaceResolver,
    name: QName<'name>,
) -> Result<(&'namespace [u8], &'name [u8])> {
    let (namespace, local) = resolver.resolve_element(name);
    match namespace {
        ResolveResult::Bound(Namespace(value)) => Ok((value, local.into_inner())),
        ResolveResult::Unbound => Err(invalid("form-control element has no namespace")),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "form-control element uses unbound namespace prefix '{}',",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

/// Compare a namespace URI after XML attribute-value normalization.
///
/// `quick_xml::NamespaceResolver` intentionally exposes declaration values in
/// lexical form, while XML namespace identity expands character/entity
/// references first.  The admitted SpreadsheetML URI is ASCII, so this
/// comparator decodes one scalar at a time without allocating a normalized
/// copy of an attacker-controlled declaration.
fn namespace_uri_matches(raw: &[u8], wanted: &[u8]) -> bool {
    if raw == wanted {
        return true;
    }
    let mut raw_index = 0usize;
    let mut wanted_index = 0usize;
    while raw_index < raw.len() && wanted_index < wanted.len() {
        let value = if raw[raw_index] == b'&' {
            let entity_start = raw_index + 1;
            let Some(relative_end) = raw[entity_start..].iter().position(|byte| *byte == b';')
            else {
                return false;
            };
            let entity_end = entity_start + relative_end;
            let Some(scalar) = namespace_entity_scalar(&raw[entity_start..entity_end]) else {
                return false;
            };
            raw_index = entity_end + 1;
            scalar
        } else {
            let value = raw[raw_index] as u32;
            raw_index += 1;
            value
        };
        if value > u32::from(u8::MAX) || wanted[wanted_index] != value as u8 {
            return false;
        }
        wanted_index += 1;
    }
    raw_index == raw.len() && wanted_index == wanted.len()
}

fn namespace_entity_scalar(entity: &[u8]) -> Option<u32> {
    match entity {
        b"amp" => Some(u32::from(b'&')),
        b"lt" => Some(u32::from(b'<')),
        b"gt" => Some(u32::from(b'>')),
        b"apos" => Some(u32::from(b'\'')),
        b"quot" => Some(u32::from(b'"')),
        _ if entity.first() == Some(&b'#') => {
            let (radix, digits) = match entity.get(1) {
                Some(b'x' | b'X') => (16, &entity[2..]),
                Some(_) => (10, &entity[1..]),
                None => return None,
            };
            if digits.is_empty() {
                return None;
            }
            let mut value = 0u32;
            for digit in digits {
                let digit = match radix {
                    16 => match digit {
                        b'0'..=b'9' => u32::from(*digit - b'0'),
                        b'a'..=b'f' => u32::from(*digit - b'a' + 10),
                        b'A'..=b'F' => u32::from(*digit - b'A' + 10),
                        _ => return None,
                    },
                    _ => match digit {
                        b'0'..=b'9' => u32::from(*digit - b'0'),
                        _ => return None,
                    },
                };
                value = value.checked_mul(radix)?.checked_add(digit)?;
            }
            (value <= 0x10_FFFF && !(0xD800..=0xDFFF).contains(&value) && value != 0)
                .then_some(value)
        },
        _ => None,
    }
}

fn element_prefix(name: QName<'_>, retained: &mut RetainedBudget) -> Result<Option<Box<[u8]>>> {
    let (_, prefix) = name.decompose();
    prefix
        .map(|value| {
            let value = value.into_inner();
            retained.charge(value.len(), "form-control root prefix")?;
            owned_bytes(value, "form-control element prefix").map(Vec::into_boxed_slice)
        })
        .transpose()
}

fn finish_root_namespace_context(
    namespaces: &mut Vec<NamespaceBinding>,
    retained: &mut RetainedBudget,
) -> Result<Arc<[NamespaceBinding]>> {
    let storage = namespaces
        .len()
        .checked_mul(size_of::<NamespaceBinding>())
        .ok_or_else(|| invalid("form-control namespace binding storage overflow"))?;
    // Keep the parsing Vec reservation live while the exact-length Arc slice
    // is charged and published.  `Arc::from(Vec)` may allocate a new slice
    // before the source Vec is dropped, so this is a real old-plus-new peak.
    retained.charge(storage, "form-control namespace binding storage")?;
    let context = Arc::from(std::mem::take(namespaces));
    retained.release_namespace_temporary();
    Ok(context)
}

fn parse_root_attributes(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: Decoder,
    source: SourceElement<'_>,
    limits: Limits,
    retained: &mut RetainedBudget,
    properties: &mut Properties,
    layouts: &mut Vec<AttributeLayout>,
    lexical: &mut Vec<LexicalAttribute>,
    namespaces: &mut Vec<NamespaceBinding>,
    opaque_bytes: &mut usize,
) -> Result<()> {
    let mut attributes_binding = element.attributes();
    let attributes = attributes_binding.with_checks(true);
    let mut count = 0usize;
    for attribute in attributes {
        let attribute =
            attribute.map_err(|error| invalid(format!("form-control attribute error: {error}")))?;
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid("form-control attribute count overflow"))?;
        if count > limits.max_attributes() {
            return Err(limit("attribute count", count, limits.max_attributes()));
        }
        if let Some(binding) = attribute.key.as_namespace_binding() {
            let prefix = match binding {
                quick_xml::name::PrefixDeclaration::Default => &[][..],
                quick_xml::name::PrefixDeclaration::Named(value) => value,
            };
            retained.charge(prefix.len(), "form-control namespace prefix")?;
            retained.charge(attribute.value.len(), "form-control namespace URI")?;
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| invalid(format!("invalid form-control namespace URI: {error}")))?;
            validate_xml_text(value.as_ref(), "form-control namespace URI")?;
            if value.len() > limits.max_opaque_bytes() {
                return Err(limit(
                    "namespace URI bytes",
                    value.len(),
                    limits.max_opaque_bytes(),
                ));
            }
            reserve_namespace_slots(namespaces, 1, retained, "form-control namespace bindings")?;
            namespaces.push(NamespaceBinding::new(
                String::from_utf8(prefix.to_vec())
                    .map_err(|error| invalid(format!("invalid namespace prefix: {error}")))?,
                value.into_owned(),
            ));
            continue;
        }
        // Resolve every qualified attribute, including unknown attributes, so
        // an unbound prefix cannot be retained as an apparently opaque token.
        if attribute.key.prefix().is_some() {
            match resolver.resolve_attribute(attribute.key).0 {
                ResolveResult::Bound(_) => {},
                ResolveResult::Unbound | ResolveResult::Unknown(_) => {
                    return Err(invalid(
                        "form-control attribute uses an unbound namespace prefix",
                    ));
                },
            }
        }
        let key = attribute.key.as_ref();
        if attribute.value.len() > limits.max_opaque_bytes() {
            return Err(limit(
                "attribute lexical value bytes",
                attribute.value.len(),
                limits.max_opaque_bytes(),
            ));
        }
        let field = ScalarField::from_wire_name(key);
        let span = find_attribute_span(source.source, source.start, source.end, key);
        let (token, value_span) =
            span.ok_or_else(|| invalid("form-control attribute source span missing"))?;
        if field.is_none() || attribute.key.prefix().is_some() {
            charge_opaque_length(
                opaque_bytes,
                token.end.saturating_sub(token.start),
                limits,
                "unknown form-control root attribute",
            )?;
            retained.charge(key.len(), "form-control unknown attribute name")?;
            retained.charge(
                token.end.saturating_sub(token.start),
                "form-control unknown attribute",
            )?;
            reserve_retained_slots(
                properties.unknown_attributes.as_mut_vec(),
                1,
                retained,
                "form-control unknown attributes",
            )?;
            properties
                .unknown_attributes
                .as_mut_vec()
                .push(OpaqueAttribute::from_lexical(
                    key,
                    &source.source[token.clone()],
                )?);
        } else if let Some(field) = field {
            retained.charge(key.len(), "form-control attribute name")?;
            retained.charge(
                attribute.value.len(),
                "form-control attribute lexical value",
            )?;
            retained.charge(
                attribute.value.len(),
                "form-control attribute decoded value",
            )?;
            if scalar_retains_string(field) {
                // The lexical projection is retained separately from the
                // semantic string/formula (and token collapse may need the
                // same upper bound transiently).  Charge that second value
                // before decoding/assignment can allocate it.
                retained.charge(attribute.value.len(), "form-control scalar semantic value")?;
            }
            let decoded = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| {
                    invalid(format!("invalid form-control attribute value: {error}"))
                })?;
            validate_xml_text(decoded.as_ref(), "form-control attribute value")?;
            if decoded.len() > limits.max_opaque_bytes() {
                return Err(limit(
                    "attribute value bytes",
                    decoded.len(),
                    limits.max_opaque_bytes(),
                ));
            }
            assign_parsed_scalar(properties, field, decoded.as_ref(), retained)?;
            reserve_retained_slots(
                lexical,
                1,
                retained,
                "form-control attribute lexical values",
            )?;
            lexical.push(LexicalAttribute {
                name: arc_bytes(key, "form-control attribute name")?,
                raw_value: arc_bytes(
                    attribute.value.as_ref(),
                    "form-control attribute lexical value",
                )?,
                decoded_value: decoded.into_owned(),
            });
        }
        reserve_retained_slots(layouts, 1, retained, "form-control attribute layout")?;
        layouts.push(AttributeLayout {
            field,
            token,
            value: value_span,
        });
    }
    // The field parser above is intentionally independent of source offsets.
    // Source ranges are repaired by the complete opening-tag scanner below.
    let _ = decoder;
    Ok(())
}

fn parse_item_list_attributes(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    source: SourceElement<'_>,
    limits: Limits,
    retained: &mut RetainedBudget,
    opaque_bytes: &mut usize,
    unknown: &mut Vec<OpaqueAttribute>,
) -> Result<()> {
    let mut count = 0usize;
    let mut attributes_binding = element.attributes();
    let attributes = attributes_binding.with_checks(true);
    for attribute in attributes {
        let attribute = attribute
            .map_err(|error| invalid(format!("form-control itemLst attribute error: {error}")))?;
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid("form-control itemLst attribute count overflow"))?;
        if count > limits.max_attributes() {
            return Err(limit(
                "itemLst attribute count",
                count,
                limits.max_attributes(),
            ));
        }
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        if attribute.key.prefix().is_some() {
            match resolver.resolve_attribute(attribute.key).0 {
                ResolveResult::Bound(_) => {},
                ResolveResult::Unbound | ResolveResult::Unknown(_) => {
                    return Err(invalid(
                        "form-control itemLst uses an unbound attribute prefix",
                    ));
                },
            }
        }
        let key = attribute.key.as_ref();
        let (token, _) = find_attribute_span(source.source, source.start, source.end, key)
            .ok_or_else(|| invalid("form-control itemLst attribute source span missing"))?;
        let lexical = source
            .source
            .get(token.clone())
            .ok_or_else(|| invalid("form-control itemLst attribute span exceeds source"))?;
        charge_opaque_length(
            opaque_bytes,
            lexical.len(),
            limits,
            "unknown form-control itemLst attribute",
        )?;
        retained.charge(key.len(), "form-control itemLst attribute name")?;
        retained.charge(
            lexical.len(),
            "form-control itemLst attribute lexical value",
        )?;
        reserve_retained_slots(unknown, 1, retained, "form-control itemLst attributes")?;
        unknown.push(OpaqueAttribute::from_lexical(key, lexical)?);
    }
    Ok(())
}

fn parse_item_attributes(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: Decoder,
    limits: Limits,
    retained: &mut RetainedBudget,
) -> Result<(Item, bool)> {
    let mut value = None::<String>;
    let mut opaque = false;
    let mut count = 0usize;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| invalid(format!("form-control item attribute error: {error}")))?;
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid("form-control item attribute count overflow"))?;
        if count > limits.max_attributes() {
            return Err(limit(
                "item attribute count",
                count,
                limits.max_attributes(),
            ));
        }
        if attribute.key.as_namespace_binding().is_some() {
            // Namespace declarations are outside CT_FormControlPr's item
            // attribute vocabulary.  Keep the source item intact when an
            // item-list splice would otherwise canonicalize it away.
            opaque = true;
            continue;
        }
        if attribute.key.prefix().is_some() {
            match resolver.resolve_attribute(attribute.key).0 {
                ResolveResult::Bound(_) => {},
                ResolveResult::Unbound | ResolveResult::Unknown(_) => {
                    return Err(invalid(
                        "form-control item uses an unbound attribute prefix",
                    ));
                },
            }
            opaque = true;
            continue;
        }
        if attribute.key.as_ref() == b"val" {
            if value.is_some() {
                return Err(invalid("duplicate x14:item val attribute"));
            }
            if attribute.value.len() > limits.max_item_value_bytes() {
                return Err(limit(
                    "item value lexical bytes",
                    attribute.value.len(),
                    limits.max_item_value_bytes(),
                ));
            }
            let retained_item_value = attribute
                .value
                .len()
                .checked_mul(2)
                .ok_or_else(|| invalid("form-control item value retained length overflow"))?;
            retained.charge(retained_item_value, "form-control item value and raw value")?;
            let decoded = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| invalid(format!("invalid x14:item val: {error}")))?;
            validate_xml_text(decoded.as_ref(), "x14:item val")?;
            value = Some(decoded.into_owned());
        } else {
            opaque = true;
        }
    }
    let value = value.ok_or_else(|| invalid("x14:item requires an unqualified val attribute"))?;
    if value.len() > limits.max_item_value_bytes() {
        return Err(limit(
            "item value bytes",
            value.len(),
            limits.max_item_value_bytes(),
        ));
    }
    let mut item = Item::new(value)?;
    item.opaque_attributes = opaque;
    item.raw_value = Some(item.value().to_owned());
    Ok((item, opaque))
}

fn check_opaque_attribute_budget(element: &BytesStart<'_>, limits: Limits) -> Result<()> {
    let mut attributes_binding = element.attributes();
    let attributes = attributes_binding.with_checks(true);
    let mut count = 0usize;
    for attribute in attributes {
        attribute
            .map_err(|error| invalid(format!("form-control opaque attribute error: {error}")))?;
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid("form-control opaque attribute count overflow"))?;
        if count > limits.max_attributes() {
            return Err(limit(
                "opaque element attribute count",
                count,
                limits.max_attributes(),
            ));
        }
    }
    Ok(())
}

fn extension_list_container_diagnostic(element: &BytesStart<'_>, limits: Limits) -> Result<bool> {
    check_opaque_attribute_budget(element, limits)?;
    let mut attributes_binding = element.attributes();
    let attributes = attributes_binding.with_checks(true);
    let mut diagnostic = false;
    for attribute in attributes {
        let attribute = attribute
            .map_err(|error| invalid(format!("form-control extension attribute error: {error}")))?;
        if attribute.key.as_namespace_binding().is_none() {
            diagnostic = true;
        }
    }
    Ok(diagnostic)
}

fn validate_extension_element(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: Decoder,
    limits: Limits,
    scratch: &mut ScratchBudget,
) -> Result<bool> {
    let (namespace, local) = resolver.resolve_element(element.name());
    let mut valid = matches!(namespace, ResolveResult::Bound(Namespace(value))
        if namespace_uri_matches(value, b"http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
        && local.into_inner() == b"ext";
    let mut attributes_binding = element.attributes();
    let attributes = attributes_binding.with_checks(true);
    let mut uri_seen = false;
    for attribute in attributes {
        let attribute = attribute
            .map_err(|error| invalid(format!("form-control extension attribute error: {error}")))?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        if attribute.key.prefix().is_some() {
            valid = false;
            continue;
        }
        if attribute.key.as_ref() != b"uri" {
            valid = false;
            continue;
        }
        if uri_seen {
            valid = false;
            continue;
        }
        uri_seen = true;
        if attribute.value.len() > limits.max_opaque_bytes() {
            return Err(limit(
                "extension uri lexical bytes",
                attribute.value.len(),
                limits.max_opaque_bytes(),
            ));
        }
        // `decoded_and_normalized_value` may allocate when entities are
        // present.  Charge the lexical upper bound before asking quick-xml
        // to decode it, then release that transient scratch before returning
        // to the event loop.  This is intentionally separate from the final
        // output limit and retained model budget.
        let scratch_len = attribute.value.len();
        scratch.charge(scratch_len, "form-control extension URI decode scratch")?;
        let decoded = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| invalid(format!("invalid form-control extension uri: {error}")))?;
        validate_xml_text(decoded.as_ref(), "form-control extension uri")?;
        if decoded.len() > limits.max_opaque_bytes() {
            return Err(limit(
                "extension uri bytes",
                decoded.len(),
                limits.max_opaque_bytes(),
            ));
        }
        // `uri` is an xsd:token in the local schema.  XML attribute
        // normalization has already translated literal XML whitespace; an
        // all-whitespace value therefore collapses to the empty token and is
        // retained with a diagnostic instead of being treated as a typed
        // extension URI.
        let whitespace_only = decoded
            .chars()
            .all(|value| matches!(value, ' ' | '\t' | '\r' | '\n'));
        drop(decoded);
        scratch.release(scratch_len);
        if whitespace_only {
            valid = false;
        }
    }
    if !uri_seen {
        valid = false;
    }
    Ok(valid)
}

fn namespace_context_for_element(
    element: &BytesStart<'_>,
    decoder: Decoder,
    base: &Arc<[NamespaceBinding]>,
    limits: Limits,
    retained: &mut RetainedBudget,
) -> Result<Arc<[NamespaceBinding]>> {
    let mut declaration_count = 0usize;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| invalid(format!("form-control namespace attribute error: {error}")))?;
        if attribute.key.as_namespace_binding().is_some() {
            declaration_count = declaration_count
                .checked_add(1)
                .ok_or_else(|| invalid("form-control namespace binding count overflow"))?;
        }
    }
    let mut attributes_binding = element.attributes();
    let attributes = attributes_binding.with_checks(true);
    let mut additions = Vec::<NamespaceBinding>::new();
    let _addition_storage = if declaration_count == 0 {
        0
    } else {
        reserve_temporary_slots(
            &mut additions,
            declaration_count,
            retained,
            "form-control namespace binding storage",
        )?
    };
    for attribute in attributes {
        let attribute = attribute
            .map_err(|error| invalid(format!("form-control namespace attribute error: {error}")))?;
        let Some(binding) = attribute.key.as_namespace_binding() else {
            continue;
        };
        let prefix = match binding {
            quick_xml::name::PrefixDeclaration::Default => &[][..],
            quick_xml::name::PrefixDeclaration::Named(value) => value,
        };
        retained.charge(prefix.len(), "form-control namespace prefix")?;
        retained.charge(attribute.value.len(), "form-control namespace URI")?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| invalid(format!("invalid form-control namespace URI: {error}")))?;
        validate_xml_text(value.as_ref(), "form-control namespace URI")?;
        if value.len() > limits.max_opaque_bytes() {
            return Err(limit(
                "namespace URI bytes",
                value.len(),
                limits.max_opaque_bytes(),
            ));
        }
        additions.push(NamespaceBinding::new(
            String::from_utf8(prefix.to_vec())
                .map_err(|error| invalid(format!("invalid namespace prefix: {error}")))?,
            value.into_owned(),
        ));
    }
    if additions.is_empty() {
        return Ok(Arc::clone(base));
    }
    let capacity = base
        .len()
        .checked_add(additions.len())
        .ok_or_else(|| invalid("form-control namespace binding count overflow"))?;
    let mut merged = Vec::new();
    let _merged_peak_storage = reserve_temporary_slots(
        &mut merged,
        capacity,
        retained,
        "form-control namespace binding storage",
    )?;
    merged.extend(base.iter().cloned());
    merged.extend(additions);
    let final_storage = capacity
        .checked_mul(size_of::<NamespaceBinding>())
        .ok_or_else(|| invalid("form-control namespace binding storage overflow"))?;
    let final_storage = final_storage
        .checked_add(ARC_HEADER_BYTES)
        .ok_or_else(|| invalid("form-control namespace Arc storage overflow"))?;
    // Keep both the temporary Vec capacity and the final Arc slice charge
    // live until Arc publication has completed; the source Vec is still live
    // during that conversion on allocators that cannot reuse its buffer.
    retained.charge(final_storage, "form-control namespace binding storage")?;
    let context = Arc::from(merged);
    retained.release_temporary();
    Ok(context)
}

fn assign_parsed_scalar(
    properties: &mut Properties,
    field: ScalarField,
    value: &str,
    retained: &mut RetainedBudget,
) -> Result<()> {
    let value = match field {
        ScalarField::ObjectType => {
            let collapsed = collapse_token(value)?;
            match ObjectType::parse_token(collapsed.as_ref()) {
                Some(value) => ScalarValue::ObjectType(value),
                None => {
                    replace_shared_option(
                        &mut properties.object_type,
                        Some(KnownOrUnknown::Unknown(collapsed.into_owned())),
                        retained,
                        "form-control object-type backing",
                    )?;
                    return Ok(());
                },
            }
        },
        ScalarField::Checked => {
            let collapsed = collapse_token(value)?;
            match Checked::parse_token(collapsed.as_ref()) {
                Some(value) => ScalarValue::Checked(value),
                None => {
                    replace_shared_option(
                        &mut properties.checked,
                        Some(KnownOrUnknown::Unknown(collapsed.into_owned())),
                        retained,
                        "form-control checked backing",
                    )?;
                    return Ok(());
                },
            }
        },
        ScalarField::DropStyle => {
            let collapsed = collapse_token(value)?;
            match DropStyle::parse_token(collapsed.as_ref()) {
                Some(value) => ScalarValue::DropStyle(value),
                None => {
                    replace_shared_option(
                        &mut properties.drop_style,
                        Some(KnownOrUnknown::Unknown(collapsed.into_owned())),
                        retained,
                        "form-control drop-style backing",
                    )?;
                    return Ok(());
                },
            }
        },
        ScalarField::SelType => {
            let collapsed = collapse_token(value)?;
            match SelectionType::parse_token(collapsed.as_ref()) {
                Some(value) => ScalarValue::SelectionType(value),
                None => {
                    replace_shared_option(
                        &mut properties.seltype,
                        Some(KnownOrUnknown::Unknown(collapsed.into_owned())),
                        retained,
                        "form-control selection-type backing",
                    )?;
                    return Ok(());
                },
            }
        },
        ScalarField::EditVal => {
            let collapsed = collapse_token(value)?;
            match EditValidation::parse_token(collapsed.as_ref()) {
                Some(value) => ScalarValue::EditValidation(value),
                None => {
                    replace_shared_option(
                        &mut properties.edit_val,
                        Some(KnownOrUnknown::Unknown(collapsed.into_owned())),
                        retained,
                        "form-control edit-validation backing",
                    )?;
                    return Ok(());
                },
            }
        },
        ScalarField::TextHAlign => match TextHAlign::parse_token(value) {
            Some(value) => ScalarValue::TextHAlign(value),
            None => {
                replace_shared_option(
                    &mut properties.text_h_align,
                    Some(KnownOrUnknown::Unknown(value.to_owned())),
                    retained,
                    "form-control horizontal-alignment backing",
                )?;
                return Ok(());
            },
        },
        ScalarField::TextVAlign => match TextVAlign::parse_token(value) {
            Some(value) => ScalarValue::TextVAlign(value),
            None => {
                replace_shared_option(
                    &mut properties.text_v_align,
                    Some(KnownOrUnknown::Unknown(value.to_owned())),
                    retained,
                    "form-control vertical-alignment backing",
                )?;
                return Ok(());
            },
        },
        ScalarField::FmlaGroup
        | ScalarField::FmlaLink
        | ScalarField::FmlaRange
        | ScalarField::FmlaTxbx => {
            let mut formula = FormControlFormula::from_source(value.to_owned())?;
            if super::model::validate_formula_for_field(field, value).is_err() {
                formula.mark_source_only();
            }
            match field {
                ScalarField::FmlaGroup => replace_shared_option(
                    &mut properties.fmla_group,
                    Some(formula),
                    retained,
                    "form-control group-formula backing",
                )?,
                ScalarField::FmlaLink => replace_shared_option(
                    &mut properties.fmla_link,
                    Some(formula),
                    retained,
                    "form-control link-formula backing",
                )?,
                ScalarField::FmlaRange => replace_shared_option(
                    &mut properties.fmla_range,
                    Some(formula),
                    retained,
                    "form-control range-formula backing",
                )?,
                ScalarField::FmlaTxbx => replace_shared_option(
                    &mut properties.fmla_txbx,
                    Some(formula),
                    retained,
                    "form-control textbox-formula backing",
                )?,
                _ => unreachable!("formula field branch is exhaustive"),
            }
            return Ok(());
        },
        ScalarField::MultiSel => ScalarValue::String(value.to_owned()),
        ScalarField::Colored
        | ScalarField::FirstButton
        | ScalarField::Horiz
        | ScalarField::JustLastX
        | ScalarField::LockText
        | ScalarField::NoThreeD
        | ScalarField::NoThreeD2
        | ScalarField::MultiLine
        | ScalarField::VerticalBar
        | ScalarField::PasswordEdit => ScalarValue::Boolean(parse_boolean(value)?),
        ScalarField::DropLines
        | ScalarField::Dx
        | ScalarField::Inc
        | ScalarField::Max
        | ScalarField::Min
        | ScalarField::Page
        | ScalarField::Sel
        | ScalarField::Val
        | ScalarField::WidthMin => ScalarValue::Unsigned(parse_unsigned(field.wire_name(), value)?),
    };
    set_model_scalar(properties, field, Some(value), retained)
}

fn scalar_retains_string(field: ScalarField) -> bool {
    matches!(
        field,
        ScalarField::ObjectType
            | ScalarField::Checked
            | ScalarField::DropStyle
            | ScalarField::FmlaGroup
            | ScalarField::FmlaLink
            | ScalarField::FmlaRange
            | ScalarField::FmlaTxbx
            | ScalarField::MultiSel
            | ScalarField::SelType
            | ScalarField::TextHAlign
            | ScalarField::TextVAlign
            | ScalarField::EditVal
    )
}

fn collapse_token<'a>(value: &'a str) -> Result<std::borrow::Cow<'a, str>> {
    if !value
        .bytes()
        .any(|character| matches!(character, b' ' | b'\t' | b'\r' | b'\n'))
    {
        return Ok(std::borrow::Cow::Borrowed(value));
    }
    let mut collapsed = String::new();
    collapsed
        .try_reserve_exact(value.len())
        .map_err(|source| allocation("form-control collapsed token", source))?;
    let mut first = true;
    for part in value
        .split([' ', '\t', '\r', '\n'])
        .filter(|part| !part.is_empty())
    {
        if !first {
            collapsed.push(' ');
        }
        first = false;
        collapsed.push_str(part);
    }
    Ok(std::borrow::Cow::Owned(collapsed))
}

fn parse_boolean(value: &str) -> Result<bool> {
    if collapsed_token_eq(value, "0") || collapsed_token_eq(value, "false") {
        Ok(false)
    } else if collapsed_token_eq(value, "1") || collapsed_token_eq(value, "true") {
        Ok(true)
    } else {
        Err(invalid(format!("invalid XML Schema boolean '{value}'")))
    }
}

fn parse_unsigned(name: &str, value: &str) -> Result<u32> {
    parse_ascii_unsigned(value.as_bytes()).ok_or_else(|| {
        invalid(format!(
            "invalid {name} unsignedInt '{value}': invalid lexical value"
        ))
    })
}

fn find_attribute_span(
    source: &[u8],
    event_start: usize,
    event_end: usize,
    expected: &[u8],
) -> Option<(Range<usize>, Range<usize>)> {
    if event_start >= event_end || event_end > source.len() {
        return None;
    }
    let bytes = &source[event_start..event_end];
    let mut cursor = 1usize;
    while cursor < bytes.len() && !is_xml_space(bytes[cursor]) && bytes[cursor] != b'>' {
        cursor += 1;
    }
    while cursor < bytes.len() {
        while cursor < bytes.len() && is_xml_space(bytes[cursor]) {
            cursor += 1;
        }
        if cursor >= bytes.len() || bytes[cursor] == b'>' || bytes[cursor] == b'/' {
            break;
        }
        let name_start = cursor;
        while cursor < bytes.len()
            && !is_xml_space(bytes[cursor])
            && !matches!(bytes[cursor], b'=' | b'>' | b'/')
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < bytes.len() && is_xml_space(bytes[cursor]) {
            cursor += 1;
        }
        if cursor >= bytes.len() || bytes[cursor] != b'=' {
            return None;
        }
        cursor += 1;
        while cursor < bytes.len() && is_xml_space(bytes[cursor]) {
            cursor += 1;
        }
        if cursor >= bytes.len() || !matches!(bytes[cursor], b'\'' | b'"') {
            return None;
        }
        let quote = bytes[cursor];
        cursor += 1;
        let value_start = cursor;
        while cursor < bytes.len() && bytes[cursor] != quote {
            cursor += 1;
        }
        if cursor >= bytes.len() {
            return None;
        }
        let value_end = cursor;
        cursor += 1;
        if &bytes[name_start..name_end] == expected {
            let token_start = event_start + name_start;
            let token_end = event_start + cursor;
            return Some((
                token_start..token_end,
                event_start + value_start..event_start + value_end,
            ));
        }
    }
    None
}

fn root_close_insertion(source: &[u8], start: usize, end: usize) -> Result<usize> {
    let bytes = source
        .get(start..end)
        .ok_or_else(|| invalid("form-control root opening span lies outside source"))?;
    if bytes.ends_with(b"/>") {
        return end
            .checked_sub(2)
            .ok_or_else(|| invalid("form-control self-closing root span underflow"));
    }
    end.checked_sub(1)
        .ok_or_else(|| invalid("form-control root opening span underflow"))
}

fn is_xml_space(value: u8) -> bool {
    matches!(value, b' ' | b'\t' | b'\r' | b'\n')
}

fn is_xml_whitespace(value: &[u8]) -> bool {
    value.iter().all(|value| is_xml_space(*value))
}

fn set_model_scalar(
    properties: &mut Properties,
    field: ScalarField,
    value: Option<ScalarValue>,
    retained: &mut RetainedBudget,
) -> Result<()> {
    let value_ref = value.as_ref();
    validate_scalar_value(field, value_ref)?;
    properties.mark_changed();
    match field {
        ScalarField::ObjectType => replace_shared_option(
            &mut properties.object_type,
            value.and_then(|value| match value {
                ScalarValue::ObjectType(value) => Some(KnownOrUnknown::Known(value)),
                _ => None,
            }),
            retained,
            "form-control object-type backing",
        )?,
        ScalarField::Checked => replace_shared_option(
            &mut properties.checked,
            value.and_then(|value| match value {
                ScalarValue::Checked(value) => Some(KnownOrUnknown::Known(value)),
                _ => None,
            }),
            retained,
            "form-control checked backing",
        )?,
        ScalarField::Colored => properties.colored = value.and_then(bool_value),
        ScalarField::DropLines => properties.drop_lines = value.and_then(unsigned_value),
        ScalarField::DropStyle => replace_shared_option(
            &mut properties.drop_style,
            value.and_then(|value| match value {
                ScalarValue::DropStyle(value) => Some(KnownOrUnknown::Known(value)),
                _ => None,
            }),
            retained,
            "form-control drop-style backing",
        )?,
        ScalarField::Dx => properties.dx = value.and_then(unsigned_value),
        ScalarField::FirstButton => properties.first_button = value.and_then(bool_value),
        ScalarField::FmlaGroup => replace_shared_option(
            &mut properties.fmla_group,
            value.and_then(formula_value),
            retained,
            "form-control group-formula backing",
        )?,
        ScalarField::FmlaLink => replace_shared_option(
            &mut properties.fmla_link,
            value.and_then(formula_value),
            retained,
            "form-control link-formula backing",
        )?,
        ScalarField::FmlaRange => replace_shared_option(
            &mut properties.fmla_range,
            value.and_then(formula_value),
            retained,
            "form-control range-formula backing",
        )?,
        ScalarField::FmlaTxbx => replace_shared_option(
            &mut properties.fmla_txbx,
            value.and_then(formula_value),
            retained,
            "form-control textbox-formula backing",
        )?,
        ScalarField::Horiz => properties.horiz = value.and_then(bool_value),
        ScalarField::Inc => properties.inc = value.and_then(unsigned_value),
        ScalarField::JustLastX => properties.just_last_x = value.and_then(bool_value),
        ScalarField::LockText => properties.lock_text = value.and_then(bool_value),
        ScalarField::Max => properties.max = value.and_then(unsigned_value),
        ScalarField::Min => properties.min = value.and_then(unsigned_value),
        ScalarField::MultiSel => replace_shared_option(
            &mut properties.multi_sel,
            value.and_then(string_value),
            retained,
            "form-control multi-selection backing",
        )?,
        ScalarField::NoThreeD => properties.no_three_d = value.and_then(bool_value),
        ScalarField::NoThreeD2 => properties.no_three_d2 = value.and_then(bool_value),
        ScalarField::Page => properties.page = value.and_then(unsigned_value),
        ScalarField::Sel => properties.sel = value.and_then(unsigned_value),
        ScalarField::SelType => replace_shared_option(
            &mut properties.seltype,
            value.and_then(|value| match value {
                ScalarValue::SelectionType(value) => Some(KnownOrUnknown::Known(value)),
                _ => None,
            }),
            retained,
            "form-control selection-type backing",
        )?,
        ScalarField::TextHAlign => replace_shared_option(
            &mut properties.text_h_align,
            value.and_then(|value| match value {
                ScalarValue::TextHAlign(value) => Some(KnownOrUnknown::Known(value)),
                _ => None,
            }),
            retained,
            "form-control horizontal-alignment backing",
        )?,
        ScalarField::TextVAlign => replace_shared_option(
            &mut properties.text_v_align,
            value.and_then(|value| match value {
                ScalarValue::TextVAlign(value) => Some(KnownOrUnknown::Known(value)),
                _ => None,
            }),
            retained,
            "form-control vertical-alignment backing",
        )?,
        ScalarField::Val => properties.val = value.and_then(unsigned_value),
        ScalarField::WidthMin => properties.width_min = value.and_then(unsigned_value),
        ScalarField::EditVal => replace_shared_option(
            &mut properties.edit_val,
            value.and_then(|value| match value {
                ScalarValue::EditValidation(value) => Some(KnownOrUnknown::Known(value)),
                _ => None,
            }),
            retained,
            "form-control edit-validation backing",
        )?,
        ScalarField::MultiLine => properties.multi_line = value.and_then(bool_value),
        ScalarField::VerticalBar => properties.vertical_bar = value.and_then(bool_value),
        ScalarField::PasswordEdit => properties.password_edit = value.and_then(bool_value),
    }
    Ok(())
}

fn validate_scalar_value(field: ScalarField, value: Option<&ScalarValue>) -> Result<()> {
    let Some(value) = value else {
        return Ok(());
    };
    let valid = matches!(
        (field, value),
        (ScalarField::ObjectType, ScalarValue::ObjectType(_))
            | (ScalarField::Checked, ScalarValue::Checked(_))
            | (ScalarField::DropStyle, ScalarValue::DropStyle(_))
            | (ScalarField::SelType, ScalarValue::SelectionType(_))
            | (ScalarField::TextHAlign, ScalarValue::TextHAlign(_))
            | (ScalarField::TextVAlign, ScalarValue::TextVAlign(_))
            | (ScalarField::EditVal, ScalarValue::EditValidation(_))
            | (ScalarField::Colored, ScalarValue::Boolean(_))
            | (ScalarField::FirstButton, ScalarValue::Boolean(_))
            | (ScalarField::Horiz, ScalarValue::Boolean(_))
            | (ScalarField::JustLastX, ScalarValue::Boolean(_))
            | (ScalarField::LockText, ScalarValue::Boolean(_))
            | (ScalarField::NoThreeD, ScalarValue::Boolean(_))
            | (ScalarField::NoThreeD2, ScalarValue::Boolean(_))
            | (ScalarField::MultiLine, ScalarValue::Boolean(_))
            | (ScalarField::VerticalBar, ScalarValue::Boolean(_))
            | (ScalarField::PasswordEdit, ScalarValue::Boolean(_))
            | (ScalarField::DropLines, ScalarValue::Unsigned(_))
            | (ScalarField::Dx, ScalarValue::Unsigned(_))
            | (ScalarField::Inc, ScalarValue::Unsigned(_))
            | (ScalarField::Max, ScalarValue::Unsigned(_))
            | (ScalarField::Min, ScalarValue::Unsigned(_))
            | (ScalarField::Page, ScalarValue::Unsigned(_))
            | (ScalarField::Sel, ScalarValue::Unsigned(_))
            | (ScalarField::Val, ScalarValue::Unsigned(_))
            | (ScalarField::WidthMin, ScalarValue::Unsigned(_))
            | (ScalarField::FmlaGroup, ScalarValue::Formula(_))
            | (ScalarField::FmlaLink, ScalarValue::Formula(_))
            | (ScalarField::FmlaRange, ScalarValue::Formula(_))
            | (ScalarField::FmlaTxbx, ScalarValue::Formula(_))
            | (ScalarField::MultiSel, ScalarValue::String(_))
    );
    if !valid {
        return Err(invalid(format!("scalar value does not match {field:?}")));
    }
    if let ScalarValue::Formula(value) = value {
        if value.source_only()
            || super::model::validate_formula_for_field(field, value.as_str()).is_err()
        {
            return Err(invalid("form-control formula is not valid for this field"));
        }
    }
    if let ScalarValue::Unsigned(value) = value {
        if matches!(
            field,
            ScalarField::DropLines
                | ScalarField::Inc
                | ScalarField::Max
                | ScalarField::Min
                | ScalarField::Page
        ) && *value > 30_000
        {
            return Err(limit(field.wire_name(), *value as usize, 30_000));
        }
    }
    if let ScalarValue::String(value) = value {
        validate_xml_text(value, "form-control scalar value")?;
        if field == ScalarField::MultiSel {
            if value.len() > super::MAX_OPAQUE_BYTES {
                return Err(limit(
                    "multiSel bytes",
                    value.len(),
                    super::MAX_OPAQUE_BYTES,
                ));
            }
            validate_multi_selection(value)?;
        }
    }
    Ok(())
}

impl ScalarField {
    fn wire_name(self) -> &'static str {
        match self {
            Self::ObjectType => "objectType",
            Self::Checked => "checked",
            Self::Colored => "colored",
            Self::DropLines => "dropLines",
            Self::DropStyle => "dropStyle",
            Self::Dx => "dx",
            Self::FirstButton => "firstButton",
            Self::FmlaGroup => "fmlaGroup",
            Self::FmlaLink => "fmlaLink",
            Self::FmlaRange => "fmlaRange",
            Self::FmlaTxbx => "fmlaTxbx",
            Self::Horiz => "horiz",
            Self::Inc => "inc",
            Self::JustLastX => "justLastX",
            Self::LockText => "lockText",
            Self::Max => "max",
            Self::Min => "min",
            Self::MultiSel => "multiSel",
            Self::NoThreeD => "noThreeD",
            Self::NoThreeD2 => "noThreeD2",
            Self::Page => "page",
            Self::Sel => "sel",
            Self::SelType => "seltype",
            Self::TextHAlign => "textHAlign",
            Self::TextVAlign => "textVAlign",
            Self::Val => "val",
            Self::WidthMin => "widthMin",
            Self::EditVal => "editVal",
            Self::MultiLine => "multiLine",
            Self::VerticalBar => "verticalBar",
            Self::PasswordEdit => "passwordEdit",
        }
    }

    fn from_wire_name(name: &[u8]) -> Option<Self> {
        Some(match name {
            b"objectType" => Self::ObjectType,
            b"checked" => Self::Checked,
            b"colored" => Self::Colored,
            b"dropLines" => Self::DropLines,
            b"dropStyle" => Self::DropStyle,
            b"dx" => Self::Dx,
            b"firstButton" => Self::FirstButton,
            b"fmlaGroup" => Self::FmlaGroup,
            b"fmlaLink" => Self::FmlaLink,
            b"fmlaRange" => Self::FmlaRange,
            b"fmlaTxbx" => Self::FmlaTxbx,
            b"horiz" => Self::Horiz,
            b"inc" => Self::Inc,
            b"justLastX" => Self::JustLastX,
            b"lockText" => Self::LockText,
            b"max" => Self::Max,
            b"min" => Self::Min,
            b"multiSel" => Self::MultiSel,
            b"noThreeD" => Self::NoThreeD,
            b"noThreeD2" => Self::NoThreeD2,
            b"page" => Self::Page,
            b"sel" => Self::Sel,
            b"seltype" => Self::SelType,
            b"textHAlign" => Self::TextHAlign,
            b"textVAlign" => Self::TextVAlign,
            b"val" => Self::Val,
            b"widthMin" => Self::WidthMin,
            b"editVal" => Self::EditVal,
            b"multiLine" => Self::MultiLine,
            b"verticalBar" => Self::VerticalBar,
            b"passwordEdit" => Self::PasswordEdit,
            _ => return None,
        })
    }
}

fn bool_value(value: ScalarValue) -> Option<bool> {
    match value {
        ScalarValue::Boolean(value) => Some(value),
        _ => None,
    }
}

fn unsigned_value(value: ScalarValue) -> Option<u32> {
    match value {
        ScalarValue::Unsigned(value) => Some(value),
        _ => None,
    }
}

fn formula_value(value: ScalarValue) -> Option<FormControlFormula> {
    match value {
        ScalarValue::Formula(value) => Some(value),
        _ => None,
    }
}

fn string_value(value: ScalarValue) -> Option<String> {
    match value {
        ScalarValue::String(value) => Some(value),
        _ => None,
    }
}

fn scalar_xml_value(field: ScalarField, value: &ScalarValue, limits: Limits) -> Result<Vec<u8>> {
    let matches_field = matches!(
        (field, value),
        (ScalarField::ObjectType, ScalarValue::ObjectType(_))
            | (ScalarField::Checked, ScalarValue::Checked(_))
            | (ScalarField::DropStyle, ScalarValue::DropStyle(_))
            | (ScalarField::SelType, ScalarValue::SelectionType(_))
            | (ScalarField::TextHAlign, ScalarValue::TextHAlign(_))
            | (ScalarField::TextVAlign, ScalarValue::TextVAlign(_))
            | (ScalarField::EditVal, ScalarValue::EditValidation(_))
            | (ScalarField::Colored, ScalarValue::Boolean(_))
            | (ScalarField::FirstButton, ScalarValue::Boolean(_))
            | (ScalarField::Horiz, ScalarValue::Boolean(_))
            | (ScalarField::JustLastX, ScalarValue::Boolean(_))
            | (ScalarField::LockText, ScalarValue::Boolean(_))
            | (ScalarField::NoThreeD, ScalarValue::Boolean(_))
            | (ScalarField::NoThreeD2, ScalarValue::Boolean(_))
            | (ScalarField::MultiLine, ScalarValue::Boolean(_))
            | (ScalarField::VerticalBar, ScalarValue::Boolean(_))
            | (ScalarField::PasswordEdit, ScalarValue::Boolean(_))
            | (ScalarField::DropLines, ScalarValue::Unsigned(_))
            | (ScalarField::Dx, ScalarValue::Unsigned(_))
            | (ScalarField::Inc, ScalarValue::Unsigned(_))
            | (ScalarField::Max, ScalarValue::Unsigned(_))
            | (ScalarField::Min, ScalarValue::Unsigned(_))
            | (ScalarField::Page, ScalarValue::Unsigned(_))
            | (ScalarField::Sel, ScalarValue::Unsigned(_))
            | (ScalarField::Val, ScalarValue::Unsigned(_))
            | (ScalarField::WidthMin, ScalarValue::Unsigned(_))
            | (ScalarField::FmlaGroup, ScalarValue::Formula(_))
            | (ScalarField::FmlaLink, ScalarValue::Formula(_))
            | (ScalarField::FmlaRange, ScalarValue::Formula(_))
            | (ScalarField::FmlaTxbx, ScalarValue::Formula(_))
            | (ScalarField::MultiSel, ScalarValue::String(_))
    );
    if !matches_field {
        return Err(invalid(format!("scalar value does not match {field:?}")));
    }
    escape_bounded(
        scalar_lexical_value(value).as_ref(),
        limits,
        "form-control scalar value",
    )
}

fn scalar_xml_value_len(field: ScalarField, value: &ScalarValue, limits: Limits) -> Result<usize> {
    let matches_field = matches!(
        (field, value),
        (ScalarField::ObjectType, ScalarValue::ObjectType(_))
            | (ScalarField::Checked, ScalarValue::Checked(_))
            | (ScalarField::DropStyle, ScalarValue::DropStyle(_))
            | (ScalarField::SelType, ScalarValue::SelectionType(_))
            | (ScalarField::TextHAlign, ScalarValue::TextHAlign(_))
            | (ScalarField::TextVAlign, ScalarValue::TextVAlign(_))
            | (ScalarField::EditVal, ScalarValue::EditValidation(_))
            | (ScalarField::Colored, ScalarValue::Boolean(_))
            | (ScalarField::FirstButton, ScalarValue::Boolean(_))
            | (ScalarField::Horiz, ScalarValue::Boolean(_))
            | (ScalarField::JustLastX, ScalarValue::Boolean(_))
            | (ScalarField::LockText, ScalarValue::Boolean(_))
            | (ScalarField::NoThreeD, ScalarValue::Boolean(_))
            | (ScalarField::NoThreeD2, ScalarValue::Boolean(_))
            | (ScalarField::MultiLine, ScalarValue::Boolean(_))
            | (ScalarField::VerticalBar, ScalarValue::Boolean(_))
            | (ScalarField::PasswordEdit, ScalarValue::Boolean(_))
            | (ScalarField::DropLines, ScalarValue::Unsigned(_))
            | (ScalarField::Dx, ScalarValue::Unsigned(_))
            | (ScalarField::Inc, ScalarValue::Unsigned(_))
            | (ScalarField::Max, ScalarValue::Unsigned(_))
            | (ScalarField::Min, ScalarValue::Unsigned(_))
            | (ScalarField::Page, ScalarValue::Unsigned(_))
            | (ScalarField::Sel, ScalarValue::Unsigned(_))
            | (ScalarField::Val, ScalarValue::Unsigned(_))
            | (ScalarField::WidthMin, ScalarValue::Unsigned(_))
            | (ScalarField::FmlaGroup, ScalarValue::Formula(_))
            | (ScalarField::FmlaLink, ScalarValue::Formula(_))
            | (ScalarField::FmlaRange, ScalarValue::Formula(_))
            | (ScalarField::FmlaTxbx, ScalarValue::Formula(_))
            | (ScalarField::MultiSel, ScalarValue::String(_))
    );
    if !matches_field {
        return Err(invalid(format!("scalar value does not match {field:?}")));
    }
    escaped_length(
        scalar_lexical_value(value).as_ref(),
        limits,
        "form-control scalar value",
    )
}

fn scalar_lexical_value(value: &ScalarValue) -> std::borrow::Cow<'_, str> {
    match value {
        ScalarValue::ObjectType(value) => std::borrow::Cow::Borrowed(value.wire()),
        ScalarValue::Checked(value) => std::borrow::Cow::Borrowed(value.wire()),
        ScalarValue::DropStyle(value) => std::borrow::Cow::Borrowed(value.wire()),
        ScalarValue::SelectionType(value) => std::borrow::Cow::Borrowed(value.wire()),
        ScalarValue::TextHAlign(value) => std::borrow::Cow::Borrowed(value.wire()),
        ScalarValue::TextVAlign(value) => std::borrow::Cow::Borrowed(value.wire()),
        ScalarValue::EditValidation(value) => std::borrow::Cow::Borrowed(value.wire()),
        ScalarValue::Boolean(value) => std::borrow::Cow::Borrowed(if *value { "1" } else { "0" }),
        ScalarValue::Unsigned(value) => std::borrow::Cow::Owned(value.to_string()),
        ScalarValue::Formula(value) => std::borrow::Cow::Borrowed(value.as_str()),
        ScalarValue::String(value) => std::borrow::Cow::Borrowed(value.as_str()),
    }
}

fn copy_bounded(source: &[u8], maximum: usize, resource: &'static str) -> Result<Vec<u8>> {
    if source.len() > maximum {
        return Err(limit(resource, source.len(), maximum));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|source| allocation(resource, source))?;
    output.extend_from_slice(source);
    Ok(output)
}

fn arc_bytes(bytes: &[u8], resource: &'static str) -> Result<Arc<[u8]>> {
    Ok(Arc::from(owned_bytes(bytes, resource)?))
}

fn owned_bytes(bytes: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(bytes.len())
        .map_err(|source| allocation(resource, source))?;
    owned.extend_from_slice(bytes);
    Ok(owned)
}

fn splice_ranges(
    source: &[u8],
    edits: &[(Range<usize>, Vec<u8>)],
    limits: Limits,
) -> Result<Vec<u8>> {
    let mut ordered: Vec<(Range<usize>, &[u8])> = Vec::new();
    ordered
        .try_reserve_exact(edits.len())
        .map_err(|source| allocation("form-control source edits", source))?;
    for (range, replacement) in edits {
        if range.start > range.end || range.end > source.len() {
            return Err(invalid("form-control source edit lies outside source"));
        }
        ordered.push((range.clone(), replacement.as_slice()));
    }
    ordered.sort_by_key(|(range, _)| (range.start, range.end));
    let mut output_len = source.len();
    for (range, replacement) in &ordered {
        if range.end > range.start {
            output_len = output_len
                .checked_sub(range.end - range.start)
                .ok_or_else(|| invalid("form-control source edit length underflow"))?;
        }
        output_len = output_len
            .checked_add(replacement.len())
            .ok_or_else(|| invalid("form-control source edit length overflow"))?;
    }
    if output_len > limits.max_output_bytes() {
        return Err(limit(
            "generated output bytes",
            output_len,
            limits.max_output_bytes(),
        ));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("form-control source output", source))?;
    let mut cursor = 0usize;
    for (range, replacement) in ordered {
        if range.start < cursor {
            return Err(invalid("overlapping form-control source edits"));
        }
        output.extend_from_slice(&source[cursor..range.start]);
        output.extend_from_slice(replacement);
        cursor = range.end;
    }
    output.extend_from_slice(&source[cursor..]);
    Ok(output)
}

fn qname_len(prefix: Option<&[u8]>, local: &[u8]) -> Result<usize> {
    let prefix_len = prefix.map_or(Ok(0usize), |value| {
        value
            .len()
            .checked_add(1)
            .ok_or_else(|| invalid("form-control qualified name length overflow"))
    })?;
    prefix_len
        .checked_add(local.len())
        .ok_or_else(|| invalid("form-control qualified name length overflow"))
}

fn qname(prefix: Option<&[u8]>, local: &[u8]) -> Result<Vec<u8>> {
    let length = qname_len(prefix, local)?;
    let mut name = Vec::new();
    name.try_reserve_exact(length)
        .map_err(|source| allocation("form-control qualified name", source))?;
    if let Some(prefix) = prefix {
        name.extend_from_slice(prefix);
        name.push(b':');
    }
    name.extend_from_slice(local);
    Ok(name)
}

fn validate_item_for_splice(item: &Item, limits: Limits) -> Result<()> {
    if item.value().len() > limits.max_item_value_bytes() {
        return Err(limit(
            "item value bytes",
            item.value().len(),
            limits.max_item_value_bytes(),
        ));
    }
    if let Some(raw) = item.raw.as_ref() {
        if raw.byte_len() > limits.max_opaque_bytes() {
            return Err(limit(
                "opaque item bytes",
                raw.byte_len(),
                limits.max_opaque_bytes(),
            ));
        }
    }
    Ok(())
}

fn validate_item_replacement(
    properties: &Properties,
    values: &[Item],
    limits: Limits,
) -> Result<()> {
    validate_item_edit(
        properties,
        properties.item_list.is_some() || !values.is_empty(),
        values.len(),
        limits,
    )
}

fn validate_item_edit(
    properties: &Properties,
    list_present: bool,
    item_count: usize,
    limits: Limits,
) -> Result<()> {
    let Some(object_type) = properties
        .object_type
        .as_ref()
        .and_then(KnownOrUnknown::known)
    else {
        return Ok(());
    };
    if list_present && !matches!(object_type, ObjectType::List | ObjectType::Drop) {
        return Err(invalid("itemLst applies only to List and Drop controls"));
    }
    if let Some(sel) = properties.sel {
        if list_present && sel != 0 && sel as usize > item_count {
            return Err(invalid("sel exceeds the authored item count"));
        }
    }
    if item_count > limits.max_items() {
        return Err(limit("item count", item_count, limits.max_items()));
    }
    Ok(())
}

fn items_equal_by_value(left: &[Item], right: &[Item]) -> bool {
    left.len() == right.len()
        && left
            .iter()
            .zip(right)
            .all(|(left, right)| left.value() == right.value())
}

fn write_items(prefix: Option<&[u8]>, values: &[Item], limits: Limits) -> Result<Vec<u8>> {
    if values.len() > limits.max_items() {
        return Err(limit("item count", values.len(), limits.max_items()));
    }
    let name_len = qname_len(prefix, ITEM)?;
    if name_len > limits.max_output_bytes() {
        return Err(limit(
            "generated item name",
            name_len,
            limits.max_output_bytes(),
        ));
    }
    let output_len = items_output_len_with_name_len(name_len, values, limits)?;
    // The item name is allocated only after the complete candidate size has
    // passed its output preflight.  A large source prefix therefore cannot
    // consume memory before an output-cap refusal.
    let name = qname(prefix, ITEM)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("generated item bytes", source))?;
    append_items_direct(&mut output, &name, values, limits)?;
    Ok(output)
}

fn append_items_direct(
    output: &mut Vec<u8>,
    name: &[u8],
    values: &[Item],
    limits: Limits,
) -> Result<()> {
    if values.len() > limits.max_items() {
        return Err(limit("item count", values.len(), limits.max_items()));
    }
    for item in values {
        validate_item_for_splice(item, limits)?;
        append_generated_item(output, name, item, limits)?;
    }
    Ok(())
}

fn write_items_with_source_edit(
    prefix: Option<&[u8]>,
    current: &[Item],
    layout: &ItemListLayout,
    source: &[u8],
    limits: Limits,
    insertion: Option<(usize, &Item)>,
    removal: Option<usize>,
) -> Result<Vec<u8>> {
    if current.len() != layout.items.len() {
        return Err(invalid("form-control item layout count mismatch"));
    }
    if insertion.is_some() && removal.is_some() {
        return Err(invalid(
            "form-control item edit combines insertion and removal",
        ));
    }
    if let Some((index, item)) = insertion {
        if index > current.len() {
            return Err(invalid(
                "form-control item insertion index is out of bounds",
            ));
        }
        validate_item_for_splice(item, limits)?;
    }
    if let Some(index) = removal {
        if index >= current.len() {
            return Err(invalid("form-control item removal index is out of bounds"));
        }
    }
    let output_count = current
        .len()
        .checked_add(usize::from(insertion.is_some()))
        .and_then(|value| value.checked_sub(usize::from(removal.is_some())))
        .ok_or_else(|| invalid("form-control item edit count underflow"))?;
    if output_count > limits.max_items() {
        return Err(limit("item count", output_count, limits.max_items()));
    }
    let name_len = qname_len(prefix, ITEM)?;
    if name_len > limits.max_output_bytes() {
        return Err(limit(
            "generated item name",
            name_len,
            limits.max_output_bytes(),
        ));
    }

    let mut output_len = 0usize;
    for output_index in 0..output_count {
        let (item, source_span) = if let Some((insert_index, item)) = insertion {
            if output_index == insert_index {
                (item, None)
            } else {
                let old_index = if output_index < insert_index {
                    output_index
                } else {
                    output_index - 1
                };
                let span = &layout.items[old_index].span;
                (
                    current.get(old_index).ok_or_else(|| {
                        invalid("form-control item source index is out of bounds")
                    })?,
                    Some(span),
                )
            }
        } else {
            let old_index = if let Some(remove_index) = removal {
                if output_index < remove_index {
                    output_index
                } else {
                    output_index + 1
                }
            } else {
                output_index
            };
            let span = &layout.items[old_index].span;
            (
                current
                    .get(old_index)
                    .ok_or_else(|| invalid("form-control item source index is out of bounds"))?,
                Some(span),
            )
        };
        let item_len = if let Some(span) = source_span {
            let length = span
                .end
                .checked_sub(span.start)
                .ok_or_else(|| invalid("form-control item source span underflow"))?;
            if span.end > source.len() {
                return Err(invalid("form-control item source span exceeds source"));
            }
            length
        } else {
            generated_item_len_for_item_with_name_len(name_len, item, limits)?
        };
        output_len = output_len
            .checked_add(item_len)
            .ok_or_else(|| invalid("generated item bytes length overflow"))?;
        if output_len > limits.max_output_bytes() {
            return Err(limit(
                "generated item bytes",
                output_len,
                limits.max_output_bytes(),
            ));
        }
    }

    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("generated item bytes", source))?;
    let name = qname(prefix, ITEM)?;
    for output_index in 0..output_count {
        let (item, source_span) = if let Some((insert_index, item)) = insertion {
            if output_index == insert_index {
                (item, None)
            } else {
                let old_index = if output_index < insert_index {
                    output_index
                } else {
                    output_index - 1
                };
                (
                    current.get(old_index).ok_or_else(|| {
                        invalid("form-control item source index is out of bounds")
                    })?,
                    Some(&layout.items[old_index].span),
                )
            }
        } else {
            let old_index = if let Some(remove_index) = removal {
                if output_index < remove_index {
                    output_index
                } else {
                    output_index + 1
                }
            } else {
                output_index
            };
            (
                current
                    .get(old_index)
                    .ok_or_else(|| invalid("form-control item source index is out of bounds"))?,
                Some(&layout.items[old_index].span),
            )
        };
        if let Some(span) = source_span {
            output.extend_from_slice(&source[span.clone()]);
        } else {
            append_generated_item(&mut output, &name, item, limits)?;
        }
    }
    Ok(output)
}

fn generated_item_len_for_item_with_name_len(
    name_len: usize,
    item: &Item,
    limits: Limits,
) -> Result<usize> {
    validate_item_for_splice(item, limits)?;
    ensure_detached_item_markup(item)?;
    if item.opaque_attributes && item.raw_value.as_deref() != Some(item.value()) {
        return Err(invalid(
            "cannot rewrite an item while discarding opaque item markup",
        ));
    }
    if item.raw_value.as_deref() == Some(item.value()) {
        if let Some(raw) = item.raw.as_ref() {
            return Ok(raw.byte_len());
        }
    }
    generated_item_len_with_name_len(name_len, item.value(), limits)
}

fn append_generated_item(
    output: &mut Vec<u8>,
    name: &[u8],
    item: &Item,
    limits: Limits,
) -> Result<()> {
    ensure_detached_item_markup(item)?;
    if item.opaque_attributes && item.raw_value.as_deref() != Some(item.value()) {
        return Err(invalid(
            "cannot rewrite an item while discarding opaque item markup",
        ));
    }
    if item.raw_value.as_deref() == Some(item.value()) {
        if let Some(raw) = item.raw.as_ref() {
            return append_checked(output, raw.xml(), limits, "generated item bytes");
        }
    }
    append_checked(output, b"<", limits, "generated item bytes")?;
    append_checked(output, name, limits, "generated item bytes")?;
    append_checked(output, b" val=\"", limits, "generated item bytes")?;
    append_escaped_prechecked(output, item.value(), limits, "generated item bytes")?;
    append_checked(output, b"\"/>", limits, "generated item bytes")
}

fn items_output_len_with_name_len(
    name_len: usize,
    values: &[Item],
    limits: Limits,
) -> Result<usize> {
    let mut length = 0usize;
    for item in values {
        validate_item_for_splice(item, limits)?;
        ensure_detached_item_markup(item)?;
        if item.opaque_attributes && item.raw_value.as_deref() != Some(item.value()) {
            return Err(invalid(
                "cannot rewrite an item while discarding opaque item markup",
            ));
        }
        let item_len = if item.raw_value.as_deref() == Some(item.value()) {
            item.raw.as_ref().map_or_else(
                || generated_item_len_with_name_len(name_len, item.value(), limits),
                |raw| Ok(raw.byte_len()),
            )?
        } else {
            generated_item_len_with_name_len(name_len, item.value(), limits)?
        };
        length = length
            .checked_add(item_len)
            .ok_or_else(|| invalid("generated item bytes length overflow"))?;
        if length > limits.max_output_bytes() {
            return Err(limit(
                "generated item bytes",
                length,
                limits.max_output_bytes(),
            ));
        }
    }
    Ok(length)
}

fn ensure_detached_item_markup(item: &Item) -> Result<()> {
    if item.opaque_attributes && item.raw.is_none() {
        return Err(invalid(
            "opaque item markup has no retained source range for detached writing",
        ));
    }
    Ok(())
}

fn generated_item_len_with_name_len(name_len: usize, value: &str, limits: Limits) -> Result<usize> {
    let escaped = escaped_length(value, limits, "generated item bytes")?;
    name_len
        .checked_add(escaped)
        .and_then(|value| value.checked_add(10))
        .ok_or_else(|| invalid("generated item bytes length overflow"))
}

fn append_escaped_prechecked(
    output: &mut Vec<u8>,
    value: &str,
    limits: Limits,
    resource: &'static str,
) -> Result<()> {
    let encoded_len = escaped_length(value, limits, resource)?;
    let new_len = output
        .len()
        .checked_add(encoded_len)
        .ok_or_else(|| invalid("form-control escaped output length overflow"))?;
    if new_len > limits.max_output_bytes() {
        return Err(limit(resource, new_len, limits.max_output_bytes()));
    }
    output
        .try_reserve_exact(encoded_len)
        .map_err(|source| allocation(resource, source))?;
    for byte in value.bytes() {
        match byte {
            b'<' => output.extend_from_slice(b"&lt;"),
            b'>' => output.extend_from_slice(b"&gt;"),
            b'&' => output.extend_from_slice(b"&amp;"),
            b'\'' => output.extend_from_slice(b"&apos;"),
            b'"' => output.extend_from_slice(b"&quot;"),
            b'\t' => output.extend_from_slice(b"&#x9;"),
            b'\n' => output.extend_from_slice(b"&#xA;"),
            b'\r' => output.extend_from_slice(b"&#xD;"),
            byte => output.push(byte),
        }
    }
    Ok(())
}

fn append_checked(
    output: &mut Vec<u8>,
    bytes: &[u8],
    limits: Limits,
    resource: &'static str,
) -> Result<()> {
    let new_len = output
        .len()
        .checked_add(bytes.len())
        .ok_or_else(|| invalid("form-control generated output length overflow"))?;
    if new_len > limits.max_output_bytes() {
        return Err(limit(resource, new_len, limits.max_output_bytes()));
    }
    output
        .try_reserve_exact(bytes.len())
        .map_err(|source| allocation(resource, source))?;
    output.extend_from_slice(bytes);
    Ok(())
}

fn effective_default_binding(bindings: &[NamespaceBinding]) -> Option<&NamespaceBinding> {
    bindings
        .iter()
        .rev()
        .find(|binding| binding.prefix().is_empty())
}

fn root_default_uri(properties: &Properties) -> &str {
    if let Some(binding) = effective_default_binding(properties.namespaces()) {
        return binding.uri();
    }
    // A prefixed source root can legally have no default namespace.  Keep
    // that empty default when an opaque descendant contains an unprefixed
    // element which relies on the outer scope; otherwise retain the stable
    // detached writer default used for a new model and for source fragments
    // whose opaque payloads are fully qualified.
    if properties.source_bytes().is_some() && source_opaque_needs_empty_default(properties) {
        ""
    } else {
        FORM_CONTROL_NAMESPACE
    }
}

fn source_opaque_needs_empty_default(properties: &Properties) -> bool {
    let root_needs_empty = properties
        .root_extension_list()
        .is_some_and(opaque_contains_unbound_default_element);
    root_needs_empty
        || properties
            .item_list()
            .and_then(ItemList::extension_list)
            .is_some_and(opaque_contains_unbound_default_element)
}

fn opaque_contains_unbound_default_element(value: &OpaqueXml) -> bool {
    let mut reader = NsReader::from_reader(value.xml());
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    loop {
        let event = match reader.read_event() {
            Ok(event) => event,
            Err(_) => return false,
        };
        let is_unbound_default = match &event {
            Event::Start(element) | Event::Empty(element) if element.name().prefix().is_none() => {
                matches!(
                    reader.resolver().resolve_element(element.name()).0,
                    ResolveResult::Unbound
                )
            },
            _ => false,
        };
        if is_unbound_default {
            return true;
        }
        if matches!(event, Event::Eof) {
            return false;
        }
    }
}

fn item_list_default_binding<'a>(
    properties: &'a Properties,
    item_list: &'a ItemList,
) -> Option<&'a NamespaceBinding> {
    let root_uri = root_default_uri(properties);
    effective_default_binding(item_list.namespaces()).filter(|binding| binding.uri() != root_uri)
}

fn item_list_needs_explicit_prefix(properties: &Properties, item_list: &ItemList) -> bool {
    effective_default_binding(item_list.namespaces())
        .is_some_and(|binding| binding.uri() != FORM_CONTROL_NAMESPACE)
        || root_default_uri(properties) != FORM_CONTROL_NAMESPACE
}

fn detached_control_prefix(properties: &Properties) -> Result<Vec<u8>> {
    let mut candidate = Vec::new();
    candidate
        .try_reserve_exact(3)
        .map_err(|source| allocation("form-control detached prefix", source))?;
    candidate.extend_from_slice(b"x14");
    let mut suffix = 1usize;
    loop {
        let conflict = properties.namespaces().iter().any(|binding| {
            binding.prefix().as_bytes() == candidate.as_slice()
                && binding.uri() != FORM_CONTROL_NAMESPACE
        }) || properties.item_list().is_some_and(|item_list| {
            item_list.namespaces().iter().any(|binding| {
                binding.prefix().as_bytes() == candidate.as_slice()
                    && binding.uri() != FORM_CONTROL_NAMESPACE
            })
        });
        if !conflict {
            return Ok(candidate);
        }
        let suffix_text = suffix.to_string();
        candidate.clear();
        candidate
            .try_reserve_exact(4usize.saturating_add(suffix_text.len()))
            .map_err(|source| allocation("form-control detached prefix", source))?;
        candidate.extend_from_slice(b"x14_");
        candidate.extend_from_slice(suffix_text.as_bytes());
        suffix = suffix
            .checked_add(1)
            .ok_or_else(|| invalid("form-control detached prefix count overflow"))?;
    }
}

fn append_namespace_declaration(
    output: &mut Vec<u8>,
    prefix: &[u8],
    uri: &str,
    limits: Limits,
    resource: &'static str,
) -> Result<()> {
    append_checked(output, b" xmlns:", limits, resource)?;
    append_checked(output, prefix, limits, resource)?;
    append_checked(output, b"=\"", limits, resource)?;
    append_escaped_prechecked(output, uri, limits, resource)?;
    append_checked(output, b"\"", limits, resource)
}

fn detached_output_len(
    properties: &Properties,
    limits: Limits,
    prefix: Option<&[u8]>,
    root_name_len: usize,
) -> Result<usize> {
    let mut length = 0usize;
    add_output_len(&mut length, 1, limits, "form-control output")?;
    add_output_len(&mut length, root_name_len, limits, "form-control output")?;
    add_output_len(
        &mut length,
        b" xmlns=\"".len(),
        limits,
        "form-control output",
    )?;
    add_output_len(
        &mut length,
        escaped_length(
            root_default_uri(properties),
            limits,
            "form-control namespace output",
        )?,
        limits,
        "form-control output",
    )?;
    add_output_len(&mut length, 1, limits, "form-control output")?;
    if let Some(prefix) = prefix
        && !properties.namespaces().iter().any(|binding| {
            binding.prefix().as_bytes() == prefix && binding.uri() == FORM_CONTROL_NAMESPACE
        })
    {
        add_namespace_output_len(&mut length, prefix, FORM_CONTROL_NAMESPACE, limits)?;
    }
    for binding in properties.namespaces() {
        if binding.prefix().is_empty() {
            continue;
        }
        add_namespace_output_len(
            &mut length,
            binding.prefix().as_bytes(),
            binding.uri(),
            limits,
        )?;
    }
    for &field in scalar_fields_in_schema_order() {
        if let Some(value) = model_scalar_text(properties, field) {
            add_output_len(&mut length, 1, limits, "form-control output")?;
            add_output_len(
                &mut length,
                field.wire_name().len(),
                limits,
                "form-control output",
            )?;
            add_output_len(&mut length, 2, limits, "form-control output")?;
            add_output_len(
                &mut length,
                detached_scalar_output_len(properties, field, value.as_str(), limits)?,
                limits,
                "form-control output",
            )?;
            add_output_len(&mut length, 1, limits, "form-control output")?;
        }
    }
    for attribute in properties.unknown_attributes() {
        add_output_len(&mut length, 1, limits, "form-control output")?;
        add_output_len(
            &mut length,
            attribute.lexical().len(),
            limits,
            "form-control output",
        )?;
    }
    let Some(item_list) = properties.item_list() else {
        if properties.root_extension_list().is_none() {
            add_output_len(&mut length, 2, limits, "form-control output")?;
            return Ok(length);
        }
        add_output_len(&mut length, 1, limits, "form-control output")?;
        if let Some(extension) = properties.root_extension_list() {
            add_output_len(
                &mut length,
                extension.byte_len(),
                limits,
                "form-control opaque extension",
            )?;
        }
        add_output_len(
            &mut length,
            2 + root_name_len + 1,
            limits,
            "form-control output",
        )?;
        return Ok(length);
    };
    add_output_len(&mut length, 1, limits, "form-control output")?;
    add_output_len(
        &mut length,
        1 + qname_len(prefix, ITEM_LIST)?,
        limits,
        "form-control itemLst output",
    )?;
    if let Some(binding) = item_list_default_binding(properties, item_list) {
        add_output_len(
            &mut length,
            b" xmlns=\"".len() + 1,
            limits,
            "form-control itemLst namespace output",
        )?;
        add_output_len(
            &mut length,
            escaped_length(
                binding.uri(),
                limits,
                "form-control itemLst namespace output",
            )?,
            limits,
            "form-control itemLst namespace output",
        )?;
    }
    for binding in item_list.namespaces() {
        if binding.prefix().is_empty()
            || properties
                .namespaces()
                .iter()
                .any(|root| root.prefix() == binding.prefix() && root.uri() == binding.uri())
        {
            continue;
        }
        add_namespace_output_len(
            &mut length,
            binding.prefix().as_bytes(),
            binding.uri(),
            limits,
        )?;
    }
    for attribute in item_list.unknown_attributes() {
        add_output_len(&mut length, 1, limits, "form-control itemLst output")?;
        add_output_len(
            &mut length,
            attribute.lexical().len(),
            limits,
            "form-control itemLst output",
        )?;
    }
    add_output_len(&mut length, 1, limits, "form-control itemLst output")?;
    let item_name_len = qname_len(prefix, ITEM)?;
    add_output_len(
        &mut length,
        items_output_len_with_name_len(item_name_len, item_list.items(), limits)?,
        limits,
        "form-control itemLst output",
    )?;
    if let Some(extension) = item_list.extension_list() {
        add_output_len(
            &mut length,
            extension.byte_len(),
            limits,
            "form-control opaque extension",
        )?;
    }
    add_output_len(
        &mut length,
        2 + qname_len(prefix, ITEM_LIST)? + 1,
        limits,
        "form-control itemLst output",
    )?;
    if let Some(extension) = properties.root_extension_list() {
        add_output_len(
            &mut length,
            extension.byte_len(),
            limits,
            "form-control opaque extension",
        )?;
    }
    add_output_len(
        &mut length,
        2 + root_name_len + 1,
        limits,
        "form-control output",
    )?;
    Ok(length)
}

fn add_namespace_output_len(
    total: &mut usize,
    prefix: &[u8],
    uri: &str,
    limits: Limits,
) -> Result<()> {
    add_output_len(
        total,
        b" xmlns:".len(),
        limits,
        "form-control namespace output",
    )?;
    add_output_len(total, prefix.len(), limits, "form-control namespace output")?;
    add_output_len(total, 2, limits, "form-control namespace output")?;
    add_output_len(
        total,
        escaped_length(uri, limits, "form-control namespace output")?,
        limits,
        "form-control namespace output",
    )?;
    add_output_len(total, 1, limits, "form-control namespace output")
}

fn add_output_len(
    total: &mut usize,
    amount: usize,
    limits: Limits,
    resource: &'static str,
) -> Result<()> {
    *total = total
        .checked_add(amount)
        .ok_or_else(|| invalid("form-control output length overflow"))?;
    if *total > limits.max_output_bytes() {
        return Err(limit(resource, *total, limits.max_output_bytes()));
    }
    Ok(())
}

fn write_detached(properties: &Properties, limits: Limits) -> Result<Vec<u8>> {
    let needs_root_prefix = root_default_uri(properties) != FORM_CONTROL_NAMESPACE;
    let needs_item_list_prefix = properties
        .item_list()
        .is_some_and(|item_list| item_list_needs_explicit_prefix(properties, item_list));
    let detached_prefix = (needs_root_prefix || needs_item_list_prefix)
        .then(|| detached_control_prefix(properties))
        .transpose()?;
    let root_name_len = qname_len(detached_prefix.as_deref(), ROOT)?;
    let output_len = detached_output_len(
        properties,
        limits,
        detached_prefix.as_deref(),
        root_name_len,
    )?;
    // QName buffers are intentionally created only after the complete
    // candidate output has passed its exact size check.
    let root_name = qname(detached_prefix.as_deref(), ROOT)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("form-control output", source))?;
    append_checked(&mut output, b"<", limits, "form-control output")?;
    append_checked(&mut output, &root_name, limits, "form-control output")?;
    append_checked(&mut output, b" xmlns=\"", limits, "form-control output")?;
    append_escaped_prechecked(
        &mut output,
        root_default_uri(properties),
        limits,
        "form-control namespace output",
    )?;
    append_checked(&mut output, b"\"", limits, "form-control output")?;
    if let Some(prefix) = detached_prefix.as_deref()
        && !properties.namespaces().iter().any(|binding| {
            binding.prefix().as_bytes() == prefix && binding.uri() == FORM_CONTROL_NAMESPACE
        })
    {
        append_namespace_declaration(
            &mut output,
            prefix,
            FORM_CONTROL_NAMESPACE,
            limits,
            "form-control namespace output",
        )?;
    }
    for binding in properties.namespaces() {
        if binding.prefix().is_empty() || binding.uri() == FORM_CONTROL_NAMESPACE {
            if binding.prefix().is_empty() {
                continue;
            }
        }
        append_checked(
            &mut output,
            b" xmlns:",
            limits,
            "form-control namespace output",
        )?;
        append_checked(
            &mut output,
            binding.prefix().as_bytes(),
            limits,
            "form-control namespace output",
        )?;
        append_checked(&mut output, b"=\"", limits, "form-control namespace output")?;
        append_escaped_prechecked(
            &mut output,
            binding.uri(),
            limits,
            "form-control namespace output",
        )?;
        append_checked(&mut output, b"\"", limits, "form-control namespace output")?;
    }
    for &field in scalar_fields_in_schema_order() {
        if let Some(value) = model_scalar_text(properties, field) {
            append_checked(&mut output, b" ", limits, "form-control output")?;
            append_checked(
                &mut output,
                field.wire_name().as_bytes(),
                limits,
                "form-control output",
            )?;
            append_checked(&mut output, b"=\"", limits, "form-control output")?;
            append_detached_scalar(&mut output, properties, field, value.as_str(), limits)?;
            append_checked(&mut output, b"\"", limits, "form-control output")?;
        }
    }
    for attribute in properties.unknown_attributes() {
        append_checked(&mut output, b" ", limits, "form-control output")?;
        append_checked(
            &mut output,
            attribute.lexical(),
            limits,
            "form-control output",
        )?;
    }
    if properties.item_list().is_none() && properties.root_extension_list().is_none() {
        append_checked(&mut output, b"/>", limits, "form-control output")?;
        return Ok(output);
    }
    append_checked(&mut output, b">", limits, "form-control output")?;
    if let Some(item_list) = properties.item_list() {
        let list_name = qname(detached_prefix.as_deref(), ITEM_LIST)?;
        let item_name = qname(detached_prefix.as_deref(), ITEM)?;
        append_checked(&mut output, b"<", limits, "form-control itemLst output")?;
        append_checked(
            &mut output,
            &list_name,
            limits,
            "form-control itemLst output",
        )?;
        if let Some(binding) = item_list_default_binding(properties, item_list) {
            append_checked(
                &mut output,
                b" xmlns=\"",
                limits,
                "form-control itemLst namespace output",
            )?;
            append_escaped_prechecked(
                &mut output,
                binding.uri(),
                limits,
                "form-control itemLst namespace output",
            )?;
            append_checked(
                &mut output,
                b"\"",
                limits,
                "form-control itemLst namespace output",
            )?;
        }
        for binding in item_list.namespaces() {
            if binding.prefix().is_empty()
                || properties
                    .namespaces()
                    .iter()
                    .any(|root| root.prefix() == binding.prefix() && root.uri() == binding.uri())
            {
                continue;
            }
            append_namespace_declaration(
                &mut output,
                binding.prefix().as_bytes(),
                binding.uri(),
                limits,
                "form-control itemLst namespace output",
            )?;
        }
        for attribute in item_list.unknown_attributes() {
            append_checked(&mut output, b" ", limits, "form-control itemLst output")?;
            append_checked(
                &mut output,
                attribute.lexical(),
                limits,
                "form-control itemLst output",
            )?;
        }
        append_checked(&mut output, b">", limits, "form-control itemLst output")?;
        append_items_direct(&mut output, &item_name, item_list.items(), limits)?;
        if let Some(extension) = item_list.extension_list() {
            append_opaque(&mut output, extension, properties, limits)?;
        }
        append_checked(&mut output, b"</", limits, "form-control itemLst output")?;
        append_checked(
            &mut output,
            &list_name,
            limits,
            "form-control itemLst output",
        )?;
        append_checked(&mut output, b">", limits, "form-control itemLst output")?;
    }
    if let Some(extension) = properties.root_extension_list() {
        append_opaque(&mut output, extension, properties, limits)?;
    }
    append_checked(&mut output, b"</", limits, "form-control output")?;
    append_checked(&mut output, &root_name, limits, "form-control output")?;
    append_checked(&mut output, b">", limits, "form-control output")?;
    Ok(output)
}

fn append_opaque(
    output: &mut Vec<u8>,
    value: &OpaqueXml,
    properties: &Properties,
    limits: Limits,
) -> Result<()> {
    if value.byte_len() > limits.max_opaque_bytes() {
        return Err(limit(
            "opaque extension bytes",
            value.byte_len(),
            limits.max_opaque_bytes(),
        ));
    }
    // Source-bound fragments carry the root namespace context.  Detached
    // writing emits those bindings on the root, so the fragment can remain
    // byte-for-byte opaque here.
    let _ = properties;
    append_checked(output, value.xml(), limits, "form-control opaque extension")
}

fn scalar_fields_in_schema_order() -> &'static [ScalarField] {
    &[
        ScalarField::ObjectType,
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
        ScalarField::JustLastX,
        ScalarField::LockText,
        ScalarField::Max,
        ScalarField::Min,
        ScalarField::MultiSel,
        ScalarField::NoThreeD,
        ScalarField::NoThreeD2,
        ScalarField::Page,
        ScalarField::Sel,
        ScalarField::SelType,
        ScalarField::TextHAlign,
        ScalarField::TextVAlign,
        ScalarField::Val,
        ScalarField::WidthMin,
        ScalarField::EditVal,
        ScalarField::MultiLine,
        ScalarField::VerticalBar,
        ScalarField::PasswordEdit,
    ]
}

enum ScalarText<'a> {
    Borrowed(&'a str),
    Owned(String),
}

impl ScalarText<'_> {
    fn as_str(&self) -> &str {
        match self {
            Self::Borrowed(value) => value,
            Self::Owned(value) => value,
        }
    }
}

fn model_scalar_text<'a>(properties: &'a Properties, field: ScalarField) -> Option<ScalarText<'a>> {
    match field {
        ScalarField::ObjectType => {
            known_or_unknown_text(properties.object_type.as_ref(), |v| v.wire())
        },
        ScalarField::Checked => known_or_unknown_text(properties.checked.as_ref(), |v| v.wire()),
        ScalarField::Colored => properties
            .colored
            .map(|value| ScalarText::Borrowed(if value { "1" } else { "0" })),
        ScalarField::DropLines => properties
            .drop_lines
            .map(|value| ScalarText::Owned(value.to_string())),
        ScalarField::DropStyle => {
            known_or_unknown_text(properties.drop_style.as_ref(), |v| v.wire())
        },
        ScalarField::Dx => properties
            .dx
            .map(|value| ScalarText::Owned(value.to_string())),
        ScalarField::FirstButton => properties
            .first_button
            .map(|value| ScalarText::Borrowed(if value { "1" } else { "0" })),
        ScalarField::FmlaGroup => properties
            .fmla_group
            .as_ref()
            .map(|value| ScalarText::Borrowed(value.as_str())),
        ScalarField::FmlaLink => properties
            .fmla_link
            .as_ref()
            .map(|value| ScalarText::Borrowed(value.as_str())),
        ScalarField::FmlaRange => properties
            .fmla_range
            .as_ref()
            .map(|value| ScalarText::Borrowed(value.as_str())),
        ScalarField::FmlaTxbx => properties
            .fmla_txbx
            .as_ref()
            .map(|value| ScalarText::Borrowed(value.as_str())),
        ScalarField::Horiz => properties
            .horiz
            .map(|value| ScalarText::Borrowed(if value { "1" } else { "0" })),
        ScalarField::Inc => properties
            .inc
            .map(|value| ScalarText::Owned(value.to_string())),
        ScalarField::JustLastX => properties
            .just_last_x
            .map(|value| ScalarText::Borrowed(if value { "1" } else { "0" })),
        ScalarField::LockText => properties
            .lock_text
            .map(|value| ScalarText::Borrowed(if value { "1" } else { "0" })),
        ScalarField::Max => properties
            .max
            .map(|value| ScalarText::Owned(value.to_string())),
        ScalarField::Min => properties
            .min
            .map(|value| ScalarText::Owned(value.to_string())),
        ScalarField::MultiSel => properties.multi_sel.as_deref().map(ScalarText::Borrowed),
        ScalarField::NoThreeD => properties
            .no_three_d
            .map(|value| ScalarText::Borrowed(if value { "1" } else { "0" })),
        ScalarField::NoThreeD2 => properties
            .no_three_d2
            .map(|value| ScalarText::Borrowed(if value { "1" } else { "0" })),
        ScalarField::Page => properties
            .page
            .map(|value| ScalarText::Owned(value.to_string())),
        ScalarField::Sel => properties
            .sel
            .map(|value| ScalarText::Owned(value.to_string())),
        ScalarField::SelType => known_or_unknown_text(properties.seltype.as_ref(), |v| v.wire()),
        ScalarField::TextHAlign => {
            known_or_unknown_text(properties.text_h_align.as_ref(), |v| v.wire())
        },
        ScalarField::TextVAlign => {
            known_or_unknown_text(properties.text_v_align.as_ref(), |v| v.wire())
        },
        ScalarField::Val => properties
            .val
            .map(|value| ScalarText::Owned(value.to_string())),
        ScalarField::WidthMin => properties
            .width_min
            .map(|value| ScalarText::Owned(value.to_string())),
        ScalarField::EditVal => known_or_unknown_text(properties.edit_val.as_ref(), |v| v.wire()),
        ScalarField::MultiLine => properties
            .multi_line
            .map(|value| ScalarText::Borrowed(if value { "1" } else { "0" })),
        ScalarField::VerticalBar => properties
            .vertical_bar
            .map(|value| ScalarText::Borrowed(if value { "1" } else { "0" })),
        ScalarField::PasswordEdit => properties
            .password_edit
            .map(|value| ScalarText::Borrowed(if value { "1" } else { "0" })),
    }
}

fn known_or_unknown_text<'a, T>(
    value: Option<&'a KnownOrUnknown<T>>,
    known: impl Fn(&T) -> &str,
) -> Option<ScalarText<'a>> {
    value.map(|value| match value {
        KnownOrUnknown::Known(value) => ScalarText::Borrowed(known(value)),
        KnownOrUnknown::Unknown(value) => ScalarText::Borrowed(value),
    })
}

fn append_detached_scalar(
    output: &mut Vec<u8>,
    properties: &Properties,
    field: ScalarField,
    canonical: &str,
    limits: Limits,
) -> Result<()> {
    let _ = detached_scalar_output_len(properties, field, canonical, limits)?;
    let name = field.wire_name().as_bytes();
    if let Some(attribute) = properties
        .lexical_attributes
        .iter()
        .find(|attribute| attribute.name.as_ref() == name)
    {
        let matches = match field {
            ScalarField::ObjectType
            | ScalarField::Checked
            | ScalarField::DropStyle
            | ScalarField::SelType
            | ScalarField::EditVal => collapsed_token_eq(&attribute.decoded_value, canonical),
            ScalarField::Colored
            | ScalarField::FirstButton
            | ScalarField::Horiz
            | ScalarField::JustLastX
            | ScalarField::LockText
            | ScalarField::NoThreeD
            | ScalarField::NoThreeD2
            | ScalarField::MultiLine
            | ScalarField::VerticalBar
            | ScalarField::PasswordEdit => {
                boolean_canonical_eq(&attribute.decoded_value, canonical)
            },
            ScalarField::DropLines
            | ScalarField::Dx
            | ScalarField::Inc
            | ScalarField::Max
            | ScalarField::Min
            | ScalarField::Page
            | ScalarField::Sel
            | ScalarField::Val
            | ScalarField::WidthMin => unsigned_canonical_eq(&attribute.decoded_value, canonical),
            ScalarField::TextHAlign
            | ScalarField::TextVAlign
            | ScalarField::FmlaGroup
            | ScalarField::FmlaLink
            | ScalarField::FmlaRange
            | ScalarField::FmlaTxbx
            | ScalarField::MultiSel => attribute.decoded_value == canonical,
        };
        if matches {
            return append_checked(
                output,
                &attribute.raw_value,
                limits,
                "form-control scalar lexical value",
            );
        }
    }
    append_escaped_prechecked(output, canonical, limits, "form-control scalar value")
}

fn detached_scalar_output_len(
    properties: &Properties,
    field: ScalarField,
    canonical: &str,
    limits: Limits,
) -> Result<usize> {
    let name = field.wire_name().as_bytes();
    if let Some(attribute) = properties
        .lexical_attributes
        .iter()
        .find(|attribute| attribute.name.as_ref() == name)
    {
        let matches = match field {
            ScalarField::ObjectType
            | ScalarField::Checked
            | ScalarField::DropStyle
            | ScalarField::SelType
            | ScalarField::EditVal => collapsed_token_eq(&attribute.decoded_value, canonical),
            ScalarField::Colored
            | ScalarField::FirstButton
            | ScalarField::Horiz
            | ScalarField::JustLastX
            | ScalarField::LockText
            | ScalarField::NoThreeD
            | ScalarField::NoThreeD2
            | ScalarField::MultiLine
            | ScalarField::VerticalBar
            | ScalarField::PasswordEdit => {
                boolean_canonical_eq(&attribute.decoded_value, canonical)
            },
            ScalarField::DropLines
            | ScalarField::Dx
            | ScalarField::Inc
            | ScalarField::Max
            | ScalarField::Min
            | ScalarField::Page
            | ScalarField::Sel
            | ScalarField::Val
            | ScalarField::WidthMin => unsigned_canonical_eq(&attribute.decoded_value, canonical),
            ScalarField::TextHAlign
            | ScalarField::TextVAlign
            | ScalarField::FmlaGroup
            | ScalarField::FmlaLink
            | ScalarField::FmlaRange
            | ScalarField::FmlaTxbx
            | ScalarField::MultiSel => attribute.decoded_value == canonical,
        };
        if matches {
            if attribute.raw_value.len() > limits.max_output_bytes() {
                return Err(limit(
                    "form-control scalar lexical value",
                    attribute.raw_value.len(),
                    limits.max_output_bytes(),
                ));
            }
            return Ok(attribute.raw_value.len());
        }
    }
    escaped_length(canonical, limits, "form-control scalar value")
}

fn collapsed_token_eq(value: &str, expected: &str) -> bool {
    let mut parts = value
        .split([' ', '\t', '\r', '\n'])
        .filter(|part| !part.is_empty());
    parts.next() == Some(expected) && parts.next().is_none()
}

fn boolean_canonical_eq(value: &str, canonical: &str) -> bool {
    match canonical {
        "0" => collapsed_token_eq(value, "0") || collapsed_token_eq(value, "false"),
        "1" => collapsed_token_eq(value, "1") || collapsed_token_eq(value, "true"),
        _ => false,
    }
}

fn unsigned_canonical_eq(value: &str, canonical: &str) -> bool {
    let Some(expected) = parse_ascii_unsigned(canonical.as_bytes()) else {
        return false;
    };
    parse_ascii_unsigned(value.as_bytes()) == Some(expected)
}

fn parse_ascii_unsigned(value: &[u8]) -> Option<u32> {
    let mut index = 0usize;
    while value.get(index).is_some_and(|byte| is_xml_space(*byte)) {
        index += 1;
    }
    let mut parsed = 0u32;
    let mut digits = 0usize;
    while let Some(byte) = value.get(index) {
        if is_xml_space(*byte) {
            break;
        }
        let digit = byte.checked_sub(b'0')?;
        if digit > 9 {
            return None;
        }
        parsed = parsed.checked_mul(10)?.checked_add(u32::from(digit))?;
        digits += 1;
        index += 1;
    }
    while value.get(index).is_some_and(|byte| is_xml_space(*byte)) {
        index += 1;
    }
    (digits != 0 && index == value.len()).then_some(parsed)
}

fn escape_bounded(value: &str, limits: Limits, resource: &'static str) -> Result<Vec<u8>> {
    let encoded_len = escaped_length(value, limits, resource)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(encoded_len)
        .map_err(|source| allocation(resource, source))?;
    for byte in value.bytes() {
        match byte {
            b'<' => output.extend_from_slice(b"&lt;"),
            b'>' => output.extend_from_slice(b"&gt;"),
            b'&' => output.extend_from_slice(b"&amp;"),
            b'\'' => output.extend_from_slice(b"&apos;"),
            b'"' => output.extend_from_slice(b"&quot;"),
            b'\t' => output.extend_from_slice(b"&#x9;"),
            b'\n' => output.extend_from_slice(b"&#xA;"),
            b'\r' => output.extend_from_slice(b"&#xD;"),
            byte => output.push(byte),
        }
    }
    Ok(output)
}

fn escaped_length(value: &str, limits: Limits, resource: &'static str) -> Result<usize> {
    let mut encoded_len = 0usize;
    for byte in value.bytes() {
        let additional = match byte {
            b'<' | b'>' => 3,
            b'&' => 4,
            b'\'' | b'"' => 5,
            b'\t' | b'\n' | b'\r' => 4,
            _ => 0,
        };
        encoded_len = encoded_len
            .checked_add(1 + additional)
            .ok_or_else(|| invalid("form-control escaped output length overflow"))?;
    }
    if encoded_len > limits.max_output_bytes() {
        return Err(limit(resource, encoded_len, limits.max_output_bytes()));
    }
    Ok(encoded_len)
}

#[cfg(test)]
mod tests {
    use std::num::{NonZeroU64, NonZeroUsize};
    use std::path::PathBuf;
    use std::sync::Arc;

    use litchi_core::{
        Budget, CancellationSource, ExecutionContext, ExecutionLimits, FileSource,
        Limits as CoreLimits,
    };
    use litchi_opc::{OpcError, PackURI, ReadLimits, SourceBackedPackage};

    use super::*;
    use crate::form_control::{
        Checked, FormControlFormula, KnownOrUnknown, Limits, ObjectType, Properties, ScalarField,
        ScalarValue,
    };

    const NS: &str = FORM_CONTROL_NAMESPACE;

    fn managed_test_context() -> ExecutionContext {
        let budget = Budget::root(
            "form-control-source-payload-test",
            CoreLimits::new(
                64 * 1024 * 1024,
                u64::MAX,
                u64::MAX,
                u64::MAX,
                u64::MAX,
                u64::MAX,
            ),
        );
        let (_cancel_source, cancellation) = CancellationSource::pair();
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::new(1).expect("one worker"),
            NonZeroUsize::new(1).expect("one operation"),
            NonZeroU64::new(64 * 1024 * 1024).expect("in-flight cap"),
            0,
        )
        .expect("execution limits");
        ExecutionContext::new(budget, cancellation, execution_limits)
    }

    fn semantic_test_context(memory: u64) -> (Budget, CancellationSource, ExecutionContext) {
        let budget = Budget::root(
            "form-control-semantic-budget-test",
            CoreLimits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let (cancel_source, cancellation) = CancellationSource::pair();
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::new(1).expect("one worker"),
            NonZeroUsize::new(1).expect("one operation"),
            NonZeroU64::new(memory.max(1)).expect("in-flight cap"),
            0,
        )
        .expect("execution limits");
        let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
        (budget, cancel_source, context)
    }

    #[test]
    fn source_parse_context_charges_one_clone_shared_semantic_lease() {
        let xml =
            format!("<formControlPr xmlns=\"{NS}\" objectType=\"Button\" fmlaLink=\"#REF!\"/>")
                .into_bytes();
        let source = SourcePayload::Owned(Arc::new(xml.clone()));
        let (budget, _cancel_source, context) = semantic_test_context(64 * 1024);
        let properties =
            parse_source_with_limits_and_context(source, &Limits::default(), Some(&context))
                .expect("budgeted source parse");
        let charged = budget.used(Resource::Memory);
        assert!(charged > 0, "semantic parser allocations were not charged");
        assert!(properties.retained_lease.is_some());
        assert!(
            charged
                >= u64::try_from(PROPERTIES_STORAGE_BYTES + RETAINED_LEASE_STORAGE_BYTES)
                    .expect("test storage bound")
        );

        let cloned = properties.clone();
        drop(properties);
        assert_eq!(budget.used(Resource::Memory), charged);
        drop(cloned);
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(
            write(parse(&xml).expect("standalone parse")).expect("standalone write"),
            xml
        );
    }

    #[test]
    fn source_parse_context_refuses_metadata_before_model_construction() {
        let xml =
            format!("<formControlPr xmlns=\"{NS}\" objectType=\"Button\" fmlaLink=\"#REF!\"/>");
        let source = SourcePayload::Owned(Arc::new(xml.into_bytes()));
        let (_budget, _cancel_source, context) = semantic_test_context(1);
        let error =
            parse_source_with_limits_and_context(source, &Limits::default(), Some(&context))
                .expect_err("one-byte semantic budget must refuse parser metadata");
        assert!(matches!(
            error,
            super::super::FormControlError::Execution(
                ExecutionError::ResourceLimit(limit)
            ) if limit.resource == Resource::Memory
        ));
    }

    #[test]
    fn source_parse_context_checks_cancellation_at_event_boundaries() {
        let xml = format!("<formControlPr xmlns=\"{NS}\"/>").into_bytes();
        let source = SourcePayload::Owned(Arc::new(xml));
        let (_budget, cancel_source, context) = semantic_test_context(64 * 1024);
        cancel_source.cancel();
        let error =
            parse_source_with_limits_and_context(source, &Limits::default(), Some(&context))
                .expect_err("cancelled source parse");
        assert!(matches!(
            error,
            super::super::FormControlError::Execution(ExecutionError::Cancelled)
        ));
    }

    #[test]
    fn managed_source_payload_keeps_reservation_and_exact_bytes_after_owner_drop() {
        let fixture = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/form_control_properties/tdf134769.xlsx");
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            Arc::new(FileSource::open(&fixture).expect("managed fixture source")),
            ReadLimits::default(),
            managed_test_context(),
        )
        .expect("managed fixture package");
        let context = package
            .execution_context()
            .expect("managed package execution context");
        let (properties, expected) = {
            let part = package
                .part(&PackURI::new("/xl/ctrlProps/ctrlProp1.xml").expect("part URI"))
                .expect("control properties part");
            let data = part.data().expect("managed part data");
            let expected = data.as_bytes().to_vec();
            assert!(matches!(
                data.into_arc(),
                Err(OpcError::ManagedPartDataArcEscape)
            ));

            let properties = parse_source_with_limits_and_context(
                SourcePayload::Managed(data.clone()),
                &Limits::default(),
                Some(&context),
            )
            .expect("managed source parse");
            assert!(std::ptr::eq::<[u8]>(
                properties.source_bytes().expect("retained source"),
                data.as_bytes()
            ));
            (properties, expected)
        };
        drop(context);
        drop(package);
        assert_eq!(properties.source_bytes(), Some(expected.as_slice()));
        assert_eq!(
            write(&properties).expect("exact managed source write"),
            expected
        );
    }

    #[test]
    fn native_fragment_is_exact_noop_and_exposes_defaults() {
        let xml = format!(
            "<?xml version=\"1.0\"?><formControlPr xmlns=\"{NS}\" objectType=\"CheckBox\" fmlaLink=\"#REF!\"/>"
        );
        let properties = parse(xml.as_bytes()).expect("parse native form-control fragment");
        assert_eq!(
            properties.object_type(),
            Some(&KnownOrUnknown::Known(ObjectType::CheckBox))
        );
        assert!(!properties.effective_colored());
        assert_eq!(properties.effective_drop_lines(), 8);
        assert_eq!(
            properties.fmla_link().map(FormControlFormula::as_str),
            Some("#REF!")
        );
        assert_eq!(
            write(&properties).expect("exact no-op write"),
            xml.as_bytes()
        );
    }

    #[test]
    fn source_payload_parse_shares_owned_allocation_and_survives_clone() {
        let xml = format!(
            "<x14:formControlPr xmlns:x14=\"{NS}\" xmlns:future=\"urn:future\"><x14:extLst><future:ext/></x14:extLst></x14:formControlPr>"
        )
        .into_bytes();
        let source = SourcePayload::Owned(Arc::new(xml.clone()));
        let properties =
            parse_source_with_limits(source.clone(), &Limits::default()).expect("source parse");

        let retained = properties.source_bytes().expect("retained source bytes");
        assert!(std::ptr::eq::<[u8]>(retained, source.as_bytes()));
        let extension = properties
            .root_extension_list()
            .expect("retained root extension");
        let extension_start = xml
            .windows(extension.xml().len())
            .position(|window| window == extension.xml())
            .expect("extension source range");
        let retained_extension =
            &retained[extension_start..extension_start + extension.xml().len()];
        assert!(std::ptr::eq::<[u8]>(retained_extension, extension.xml()));
        assert_eq!(write(&properties).expect("exact source write"), xml);

        let cloned = properties.clone();
        drop(properties);
        assert_eq!(cloned.source_bytes(), Some(xml.as_slice()));
        assert_eq!(write(&cloned).expect("clone source write"), xml);
    }

    #[test]
    fn retained_vector_growth_accounts_actual_capacity_and_reallocation_peak() {
        let mut values = Vec::<u8>::new();
        let mut retained = RetainedBudget::new(64, None);
        reserve_retained_slots(&mut values, 2, &mut retained, "test retained slots")
            .expect("initial retained allocation");
        values.resize(values.capacity(), 0);
        let old_capacity = values.capacity();
        let target = old_capacity
            .max(1)
            .checked_mul(2)
            .expect("test capacity multiplication");
        let old_used = retained.used;

        retained.maximum = old_used + target - 1;
        assert!(matches!(
            reserve_retained_slots(&mut values, 1, &mut retained, "test retained slots"),
            Err(super::super::FormControlError::Limit { .. })
        ));
        assert_eq!(values.capacity(), old_capacity);
        assert_eq!(retained.used, old_used);

        retained.maximum = usize::MAX;
        reserve_retained_slots(&mut values, 1, &mut retained, "test retained slots")
            .expect("reallocation after peak allowance");
        assert!(values.capacity() > old_capacity);
        assert_eq!(retained.used, values.capacity());
    }

    #[test]
    fn standalone_parse_still_detaches_input_source() {
        let xml = format!("<formControlPr xmlns=\"{NS}\" objectType=\"Button\"/>").into_bytes();
        let properties = parse(&xml).expect("standalone parse");
        let retained = properties.source_bytes().expect("retained source bytes");
        assert!(!std::ptr::eq::<[u8]>(retained, xml.as_slice()));
        assert_eq!(
            write(&properties).expect("exact detached source write"),
            xml
        );
    }

    #[test]
    fn xml10_character_boundaries_are_checked_exactly() {
        assert!(Item::new("ok\u{7f}").is_ok());
        let opaque = b"<opaque>\x7f</opaque>".to_vec();
        assert!(OpaqueXml::new(opaque).is_ok());
        assert!(Item::new("bad\u{fffe}").is_err());
        assert!(Item::new("bad\u{ffff}").is_err());
    }

    #[test]
    fn generated_xsd_string_attributes_preserve_xml_whitespace_references() {
        let value = "a\tb\nc\rd";
        let item = Item::new(value).expect("item value");
        let list = ItemList::new([item]).expect("item list");
        let mut properties = Properties::new();
        properties.set_object_type(Some(ObjectType::List));
        properties.set_item_list(Some(list)).expect("set list");
        let output = write(&properties).expect("detached item output");
        let text = String::from_utf8(output.clone()).expect("UTF-8 XML");
        assert!(text.contains("a&#x9;b&#xA;c&#xD;d"));
        assert_eq!(
            parse(&output)
                .expect("reparse item output")
                .item_list()
                .expect("item list")
                .items()[0]
                .value(),
            value
        );
    }

    #[test]
    fn scalar_splice_preserves_lexical_context_and_invalid_formula() {
        let xml = format!(
            "<formControlPr xmlns=\"{NS}\" objectType='Drop' fmlaLink=\"#REF!\" colored = 'false'/>"
        );
        let view = inspect(xml.as_bytes()).expect("inspect source");
        let edited = view
            .replace_scalar(ScalarField::Colored, Some(ScalarValue::Boolean(true)))
            .expect("replace colored");
        assert!(String::from_utf8_lossy(&edited).contains("colored = '1'"));
        assert!(String::from_utf8_lossy(&edited).contains("fmlaLink=\"#REF!\""));
        let reparsed = parse(&edited).expect("parse edited source");
        assert_eq!(reparsed.colored(), Some(true));
        assert_eq!(
            reparsed.fmla_link().map(FormControlFormula::as_str),
            Some("#REF!")
        );
    }

    #[test]
    fn source_formula_shape_diagnostic_is_preserved_but_untrusted_formula_is_not_authored() {
        let xml = format!(
            "<formControlPr xmlns=\"{NS}\" objectType=\"Drop\" fmlaLink=\"A1:B2\" colored=\"0\"/>"
        );
        let view = inspect_with_limits(xml.as_bytes(), Limits::default())
            .expect("read source-only field-shape diagnostic");
        assert_eq!(
            view.properties()
                .fmla_link()
                .map(FormControlFormula::as_str),
            Some("A1:B2")
        );
        let edited = view
            .replace_scalar(ScalarField::Colored, Some(ScalarValue::Boolean(true)))
            .expect("unrelated scalar edit");
        let edited_text = String::from_utf8(edited).expect("edited UTF-8 XML");
        assert!(edited_text.contains("fmlaLink=\"A1:B2\""));
        assert!(edited_text.contains("colored=\"1\""));

        let formula = FormControlFormula::new("A1:B2").expect("generic formula wrapper");
        assert!(
            replace_scalar(
                xml.as_bytes(),
                ScalarField::FmlaLink,
                Some(ScalarValue::Formula(formula)),
            )
            .is_err()
        );
    }

    #[test]
    fn item_list_order_and_source_splice_are_bounded() {
        let xml = format!(
            "<formControlPr xmlns=\"{NS}\" objectType=\"Drop\"><itemLst>\n  <item val=\"A &amp; B\"/>\n  <item val='C'/>\n</itemLst></formControlPr>"
        );
        let properties = parse(xml.as_bytes()).expect("parse item list");
        let items = properties.item_list().expect("item list").items();
        assert_eq!(
            items.iter().map(Item::value).collect::<Vec<_>>(),
            ["A & B", "C"]
        );
        let replacement = vec![Item::new("X").expect("item"), Item::new("Y").expect("item")];
        let edited = replace_items(xml.as_bytes(), &replacement).expect("replace items");
        assert!(String::from_utf8_lossy(&edited).contains("<item val=\"X\"/>"));
        assert_eq!(
            parse(&edited)
                .expect("parse replacement")
                .item_list()
                .unwrap()
                .len(),
            2
        );
    }

    #[test]
    fn item_replacement_reuses_unchanged_lexical_item_ranges() {
        let xml = format!(
            "<formControlPr xmlns=\"{NS}\" objectType=\"Drop\"><itemLst>\n  <item val='A &amp; B'/>\n  <item val='C'/>\n</itemLst></formControlPr>"
        );
        let replacement = vec![
            Item::new("changed").expect("item"),
            Item::new("C").expect("item"),
        ];
        let edited = replace_items(xml.as_bytes(), &replacement).expect("replace one item");
        let text = String::from_utf8(edited).expect("UTF-8 XML");
        assert!(text.contains("<item val=\"changed\"/>"));
        assert!(text.contains("\n  <item val='C'/>"));
        assert!(!text.contains("<item val=\"C\"/>"));
    }

    #[test]
    fn empty_item_list_is_expanded_without_broken_self_closing_markup() {
        let xml =
            format!("<formControlPr xmlns=\"{NS}\" objectType=\"Drop\"><itemLst/></formControlPr>");
        let edited = replace_items(xml.as_bytes(), &[Item::new("one").expect("item")])
            .expect("expand item list");
        let text = String::from_utf8_lossy(&edited);
        assert!(text.contains("<itemLst><item val=\"one\"/></itemLst>"));
        assert!(!text.contains("<itemLst/>"));
        assert_eq!(
            parse(&edited)
                .expect("parse expanded list")
                .item_list()
                .unwrap()
                .len(),
            1
        );
    }

    #[test]
    fn empty_item_list_splice_preserves_opening_attributes_and_scope() {
        let xml = format!(
            "<formControlPr xmlns=\"{NS}\" objectType=\"Drop\"><itemLst extra='keep' xmlns:future=\"urn:future\"/></formControlPr>"
        );
        let edited = replace_items(xml.as_bytes(), &[Item::new("one").expect("item")])
            .expect("expand attributed item list");
        let text = String::from_utf8(edited.clone()).expect("UTF-8 XML");
        assert!(
            text.contains("<itemLst extra='keep' xmlns:future=\"urn:future\">")
                && text.contains("<item val=\"one\"/></itemLst>")
        );
        assert_eq!(
            parse(&edited)
                .expect("parse expanded attributed list")
                .item_list()
                .unwrap()
                .len(),
            1
        );
    }

    #[test]
    fn opaque_namespaces_and_extensions_survive_detached_scalar_edit() {
        let xml = format!(
            "<x14:formControlPr xmlns:x14=\"{NS}\" xmlns:foo=\"urn:future\" foo:state='A&amp;B' fmlaLink=\"#REF!\"><x14:extLst><foo:future/></x14:extLst></x14:formControlPr>"
        );
        let mut properties = parse(xml.as_bytes()).expect("parse extension source");
        properties.set_colored(Some(true));
        let edited = write(&properties).expect("detached edit");
        let text = String::from_utf8_lossy(&edited);
        assert!(text.contains("foo:state='A&amp;B'"));
        assert!(text.contains("<x14:extLst>"));
        assert!(text.contains("<foo:future/>"));
        assert!(text.contains("colored=\"1\""));
        parse(&edited).expect("parse detached extension edit");
    }

    #[test]
    fn detached_opaque_write_keeps_an_absent_outer_default_namespace_empty() {
        let xml = format!(
            "<x14:formControlPr xmlns:x14=\"{NS}\" xmlns:c=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" objectType=\"Button\"><x14:extLst><c:ext uri=\"urn:future\"><opaque><child/></opaque></c:ext></x14:extLst></x14:formControlPr>"
        );
        let mut properties = parse(xml.as_bytes()).expect("parse prefixed source");
        properties.set_lock_text(Some(true));
        let output = write(&properties).expect("detached source edit");
        let text = String::from_utf8(output.clone()).expect("UTF-8 XML");
        assert!(text.contains("xmlns=\"\""));
        assert!(text.contains("<opaque><child/></opaque>"));
        parse(&output).expect("reparse detached source edit");
    }

    #[test]
    fn borrowed_opaque_item_cannot_be_detached_without_its_source_range() {
        let source = format!(
            "<x14:formControlPr xmlns:x14=\"{NS}\"><x14:itemLst><x14:item val=\"a\" foreign=\"b\"/></x14:itemLst></x14:formControlPr>"
        );
        let view = inspect(source.as_bytes()).expect("borrowed source view");
        let source_properties = view.properties();
        assert_eq!(
            write(&source_properties).expect("source no-op"),
            source.as_bytes()
        );

        let detached = (*source_properties).clone();
        assert!(write(&detached).is_err());
    }

    #[test]
    fn detached_prefix_avoids_item_list_local_prefix_collisions() {
        let source = format!(
            "<x:formControlPr xmlns:x=\"{NS}\" xmlns=\"urn:outer\"><x:itemLst xmlns:x14=\"urn:other\"><x:item val=\"a\"/></x:itemLst></x:formControlPr>"
        );
        let mut properties = parse(source.as_bytes()).expect("source parse");
        properties.set_colored(Some(true));
        let output = write(&properties).expect("detached write");
        assert!(
            output
                .windows(b"<x14_1:formControlPr".len())
                .any(|window| window == b"<x14_1:formControlPr")
        );
        assert!(
            output
                .windows(b"<x14_1:itemLst".len())
                .any(|window| window == b"<x14_1:itemLst")
        );
        assert!(parse(&output).is_ok());
    }

    #[test]
    fn namespace_and_bounds_fail_closed() {
        let wrong = b"<formControlPr xmlns=\"urn:not-x14\"/>";
        assert!(parse(wrong).is_err());
        let xml = format!("<formControlPr xmlns=\"{NS}\"/>");
        let limits = Limits::new().with_max_part_bytes(xml.len() - 1);
        assert!(parse_with_limits(xml.as_bytes(), &limits).is_err());
        let mut properties = Properties::new();
        properties.set_object_type(Some(ObjectType::CheckBox));
        assert!(FormControlFormula::new("#REF!").is_err());
        properties.set_fmla_link(Some(FormControlFormula::new("A1").expect("formula")));
        assert!(write(&properties).is_ok());
        properties.set_checked(Some(Checked::Checked));
    }

    #[test]
    fn source_formula_and_detached_opaque_caps_fail_closed() {
        let formula = "A".repeat(super::super::MAX_FORMULA_BYTES + 1);
        let xml = format!("<formControlPr xmlns=\"{NS}\" fmlaLink=\"{formula}\"/>");
        assert!(parse(xml.as_bytes()).is_err());
        let opaque = vec![b'x'; super::super::MAX_OPAQUE_BYTES + 1];
        assert!(OpaqueXml::new(opaque).is_err());
        assert!(OpaqueXml::new(b"<a><b/></a>".to_vec()).is_ok());
        assert!(OpaqueXml::new(b"<a><b></a>".to_vec()).is_err());
        assert!(OpaqueXml::new(b"<future:item/>".to_vec()).is_err());
    }

    #[test]
    fn aggregate_opaque_cap_is_enforced_during_owned_and_borrowed_scans() {
        let item = r#"<item val="one" extra="012345678901234567890123456789012345678901234567890123456789012345678901234567890"/>"#;
        let root_attribute = r#"foreign="012345678901234567890123456789012345678901234567890123456789012345678901234567890""#;
        let xml = format!(
            r#"<formControlPr xmlns="{NS}" objectType="Drop" {root_attribute}><itemLst>{item}</itemLst><extLst/></formControlPr>"#
        );
        let owned = parse(xml.as_bytes()).expect("parse opaque source");
        assert!(owned.item_list().unwrap().items()[0].has_opaque_markup());
        let item_bytes = item.len();
        let root_attribute_bytes = root_attribute.len();
        let extension_bytes = b"<extLst/>".len();
        let aggregate = item_bytes + root_attribute_bytes + extension_bytes;
        assert!(aggregate > item_bytes + extension_bytes);
        let exact = Limits::new().with_max_opaque_bytes(aggregate);
        parse_with_limits(xml.as_bytes(), &exact).expect("exact opaque aggregate");
        inspect_with_limits(xml.as_bytes(), exact).expect("borrowed exact opaque aggregate");
        let one_under = Limits::new().with_max_opaque_bytes(aggregate - 1);
        assert!(parse_with_limits(xml.as_bytes(), &one_under).is_err());
        assert!(inspect_with_limits(xml.as_bytes(), one_under).is_err());
    }

    #[test]
    fn item_list_unknown_attributes_are_bounded_and_source_preserved() {
        let value = "x".repeat(80);
        let xml = format!(
            r#"<formControlPr xmlns="{NS}" objectType="Drop"><itemLst extra="{value}"><item val="one"/></itemLst></formControlPr>"#
        );
        let properties = parse(xml.as_bytes()).expect("parse itemLst attribute");
        let attributes = properties.item_list().unwrap().unknown_attributes();
        assert_eq!(attributes.len(), 1);
        assert_eq!(
            attributes[0].lexical(),
            format!("extra=\"{value}\"").as_bytes()
        );
        assert_eq!(
            write(&properties).expect("exact itemLst write"),
            xml.as_bytes()
        );
        let mut changed = properties.clone();
        changed.set_no_three_d(Some(true));
        let changed = write(&changed).expect("detached itemLst write");
        assert!(String::from_utf8_lossy(&changed).contains(&format!("extra=\"{value}\"")));
        let opaque = attributes[0].lexical().len();
        let exact = Limits::new().with_max_opaque_bytes(opaque);
        parse_with_limits(xml.as_bytes(), &exact).expect("exact itemLst opaque cap");
        inspect_with_limits(xml.as_bytes(), exact).expect("borrowed exact itemLst cap");
        assert!(
            parse_with_limits(
                xml.as_bytes(),
                &Limits::new().with_max_opaque_bytes(opaque - 1),
            )
            .is_err()
        );
    }

    #[test]
    fn inherited_namespace_scope_is_shared_across_many_borrowed_items() {
        let inherited = "n".repeat(4096);
        let mut xml = format!(
            r#"<formControlPr xmlns="{NS}" xmlns:inherited="urn:{inherited}" objectType="Drop"><itemLst>"#
        );
        for index in 0..64 {
            xml.push_str(&format!(
                r#"<item xmlns:local{index}="urn:local{index}" val="{index}"/>"#
            ));
        }
        xml.push_str("</itemLst></formControlPr>");
        let view = inspect(xml.as_bytes()).expect("borrowed many-item source");
        assert!(view.properties().source_bytes().is_none());
        assert_eq!(view.properties().item_list().unwrap().len(), 64);
        assert_eq!(
            write(view.properties()).expect("borrowed exact write"),
            xml.as_bytes()
        );
    }

    #[test]
    fn namespace_context_capacity_is_exact_at_parse_and_write_boundaries() {
        let xml = format!(
            r#"<x14:formControlPr xmlns:x14="{NS}" xmlns:root="urn:root" objectType="Drop"><x14:itemLst xmlns:item="urn:item"><x14:item val="one"/></x14:itemLst></x14:formControlPr>"#
        );
        let mut parse_low = 0usize;
        let mut parse_high = Limits::default().max_retained_bytes();
        while parse_low < parse_high {
            let middle = parse_low + (parse_high - parse_low) / 2;
            if parse_with_limits(
                xml.as_bytes(),
                &Limits::new().with_max_retained_bytes(middle),
            )
            .is_ok()
            {
                parse_high = middle;
            } else {
                parse_low = middle.saturating_add(1);
            }
        }
        let parse_required = parse_low;
        parse_with_limits(
            xml.as_bytes(),
            &Limits::new().with_max_retained_bytes(parse_required),
        )
        .expect("parse exact namespace retained boundary");
        assert!(
            parse_required == 0
                || parse_with_limits(
                    xml.as_bytes(),
                    &Limits::new().with_max_retained_bytes(parse_required - 1),
                )
                .is_err()
        );

        let mut dirty = parse(xml.as_bytes()).expect("parse namespace context");
        dirty.set_just_last_x(Some(true));
        let mut write_low = 0usize;
        let mut write_high = Limits::default().max_retained_bytes();
        while write_low < write_high {
            let middle = write_low + (write_high - write_low) / 2;
            if write_with_limits(&dirty, &Limits::new().with_max_retained_bytes(middle)).is_ok() {
                write_high = middle;
            } else {
                write_low = middle.saturating_add(1);
            }
        }
        let write_required = write_low;
        write_with_limits(
            &dirty,
            &Limits::new().with_max_retained_bytes(write_required),
        )
        .expect("write exact namespace retained boundary");
        assert!(
            write_required == 0
                || write_with_limits(
                    &dirty,
                    &Limits::new().with_max_retained_bytes(write_required - 1),
                )
                .is_err()
        );
        // Parsing briefly owns both the inherited-context additions and the
        // merged context.  The writer sees only the final Arc-backed model,
        // so its retained boundary can be lower while each boundary remains
        // exact and independently enforced.
        assert!(parse_required >= write_required);
    }

    #[test]
    fn distinct_namespace_context_charges_one_arc_header_at_boundary() {
        let root: Arc<[NamespaceBinding]> = Arc::from(vec![NamespaceBinding::new(
            "root".to_owned(),
            "urn:root".to_owned(),
        )]);
        let child: Arc<[NamespaceBinding]> = Arc::from(root.iter().cloned().collect::<Vec<_>>());

        let mut root_context = Properties::new();
        root_context.namespaces = Arc::clone(&root);
        root_context
            .item_list
            .replace(Some(ItemList::from_parts_with_namespace(
                Vec::new(),
                None,
                Arc::clone(&root),
                Vec::new(),
            )));
        let mut child_context = Properties::new();
        child_context.namespaces = Arc::clone(&root);
        child_context
            .item_list
            .replace(Some(ItemList::from_parts_with_namespace(
                Vec::new(),
                None,
                Arc::clone(&child),
                Vec::new(),
            )));

        fn required(properties: &Properties) -> usize {
            let mut low = 0usize;
            let mut high = Limits::default().max_retained_bytes();
            while low < high {
                let middle = low + (high - low) / 2;
                if properties
                    .validate_retained_with_limits(Limits::new().with_max_retained_bytes(middle))
                    .is_ok()
                {
                    high = middle;
                } else {
                    low = middle.saturating_add(1);
                }
            }
            low
        }

        let root_required = required(&root_context);
        let child_required = required(&child_context);
        assert_eq!(
            child_required - root_required,
            size_of::<NamespaceBinding>() + ARC_HEADER_BYTES,
            "a distinct child context owns its binding payload and one Arc header"
        );
        child_context
            .validate_retained_with_limits(Limits::new().with_max_retained_bytes(child_required))
            .expect("child namespace context at exact retained boundary");
        assert!(child_required > 0);
        assert!(
            child_context
                .validate_retained_with_limits(
                    Limits::new().with_max_retained_bytes(child_required - 1),
                )
                .is_err()
        );
    }

    #[test]
    fn namespace_arc_publication_counts_source_vec_peak() {
        let mut measured = RetainedBudget::new(usize::MAX, None);
        let mut measured_namespaces = Vec::new();
        reserve_namespace_slots(
            &mut measured_namespaces,
            2,
            &mut measured,
            "test namespace storage",
        )
        .expect("reserve namespace source vector");
        measured_namespaces.push(NamespaceBinding::new(String::new(), String::new()));
        measured_namespaces.push(NamespaceBinding::new(String::new(), String::new()));
        measured_namespaces.pop();
        let temporary_storage = measured.namespace_temporary_used;
        let final_storage = size_of::<NamespaceBinding>();
        assert!(temporary_storage > final_storage);
        let peak = measured.used + temporary_storage + final_storage;
        drop(measured_namespaces);
        drop(measured);

        for (maximum, succeeds) in [(peak - 1, false), (peak, true)] {
            let mut retained = RetainedBudget::new(maximum, None);
            let mut namespaces = Vec::new();
            reserve_namespace_slots(&mut namespaces, 2, &mut retained, "test namespace storage")
                .expect("reserve bounded namespace source vector");
            namespaces.push(NamespaceBinding::new(String::new(), String::new()));
            namespaces.push(NamespaceBinding::new(String::new(), String::new()));
            namespaces.pop();
            let result = finish_root_namespace_context(&mut namespaces, &mut retained);
            assert_eq!(result.is_ok(), succeeds);
            if succeeds {
                assert_eq!(retained.used, final_storage);
                assert_eq!(retained.namespace_temporary_used, 0);
            }
        }
    }

    #[test]
    fn detached_output_is_rejected_before_large_formula_serialization() {
        let mut properties = Properties::new();
        properties.set_object_type(Some(ObjectType::Button));
        properties.set_fmla_link(Some(
            FormControlFormula::new("Sheet1!$A$1").expect("valid formula"),
        ));
        assert!(write_with_limits(&properties, &Limits::new().with_max_output_bytes(1)).is_err());
    }

    #[test]
    fn item_splice_output_cap_is_checked_before_candidate_buffer() {
        let xml = format!(
            r#"<formControlPr xmlns="{NS}" objectType="Drop"><itemLst><item val="one"/></itemLst></formControlPr>"#
        );
        let value = Item::new("inserted").expect("item");
        let expected = insert_item(xml.as_bytes(), 1, value.clone()).expect("default item splice");
        let under = Limits::new().with_max_output_bytes(expected.len() - 1);
        let view = inspect_with_limits(xml.as_bytes(), under).expect("inspect under output cap");
        assert!(view.insert_item(1, value.clone()).is_err());
        let exact = Limits::new().with_max_output_bytes(expected.len());
        let view = inspect_with_limits(xml.as_bytes(), exact).expect("inspect exact output cap");
        assert_eq!(
            view.insert_item(1, value).expect("exact item splice"),
            expected
        );
    }
}
