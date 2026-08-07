//! Private-type Buffa projection for `GroupNode` category labels.
//!
//! The generated projection contains only UUID and scalar wrapper messages.
//! Recursive children and `CellValue` branches are routed by streaming over
//! the preflighted source bytes. This avoids generated repeated-field storage
//! and keeps traversal memory proportional to nesting depth.

use std::fmt;

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_group_node_category_generated::LitchiIwaProjection as projection;

const GROUP_UID_FIELD: u32 = 1;
const GROUP_CHILD_FIELD: u32 = 3;
const GROUP_CELL_VALUE_FIELD: u32 = 7;
const BOOLEAN_VALUE_FIELD: u32 = 2;
const DATE_VALUE_FIELD: u32 = 3;
const NUMBER_VALUE_FIELD: u32 = 4;
const STRING_VALUE_FIELD: u32 = 5;

/// Finite limits already established by the `GroupNode` wire adapter.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    recursion_limit: u32,
}

impl DecodeOptions {
    /// Build an explicit finite profile for one preflighted `GroupNode`
    /// payload.
    #[must_use]
    pub const fn new(max_message_bytes: usize, recursion_limit: u32) -> Self {
        Self {
            max_message_bytes,
            recursion_limit,
        }
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }

    fn descend(self) -> Result<Self, DecodeError> {
        let recursion_limit = self
            .recursion_limit
            .checked_sub(1)
            .ok_or(DecodeError(buffa::DecodeError::RecursionLimitExceeded))?;
        Ok(Self {
            recursion_limit,
            ..self
        })
    }
}

/// Failure from the private Buffa `GroupNode` projection decoder.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError(buffa::DecodeError);

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.0.fmt(formatter)
    }
}

impl std::error::Error for DecodeError {}

impl From<buffa::DecodeError> for DecodeError {
    fn from(error: buffa::DecodeError) -> Self {
        Self(error)
    }
}

/// The two scalar halves of a `GroupNode` `UUID`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct GroupUid {
    lower: u64,
    upper: u64,
}

impl GroupUid {
    /// Lower `UUID` half.
    #[must_use]
    pub const fn lower(self) -> u64 {
        self.lower
    }

    /// Upper `UUID` half.
    #[must_use]
    pub const fn upper(self) -> u64 {
        self.upper
    }
}

/// Borrowed `GroupNode` projection with no generated type in its public
/// surface.
#[derive(Clone, Copy, Debug)]
pub struct GroupNodeView<'source> {
    source: &'source [u8],
    options: DecodeOptions,
}

impl<'source> GroupNodeView<'source> {
    /// Merge and read the group UUID, returning `None` when field 1 is absent.
    ///
    /// Each UUID fragment is decoded independently through Buffa, then merged
    /// in source order. This preserves Protocol Buffers last-value semantics
    /// without retaining an input-sized fragment vector.
    pub fn group_uid(&self) -> Result<Option<GroupUid>, DecodeError> {
        let mut found = false;
        let mut lower = 0;
        let mut upper = 0;
        let fragment_options = self.options.descend()?;
        for fragment in self.length_delimited_fields(GROUP_UID_FIELD) {
            found = true;
            let view: projection::UuidLazyView<'source> =
                fragment_options.buffa().decode_lazy_view(fragment?)?;
            if view.has_lower() {
                lower = view.lower;
            }
            if view.has_upper() {
                upper = view.upper;
            }
        }
        Ok(found.then_some(GroupUid { lower, upper }))
    }

    /// Number of field-1 UUID fragments in source order.
    pub fn group_uid_occurrences(&self) -> Result<usize, DecodeError> {
        self.length_delimited_fields(GROUP_UID_FIELD)
            .try_fold(0usize, |count, fragment| {
                let _ = fragment?;
                Ok(count.saturating_add(1))
            })
    }

    /// Number of child-node fragments in source order.
    pub fn child_count(&self) -> Result<usize, DecodeError> {
        self.length_delimited_fields(GROUP_CHILD_FIELD)
            .try_fold(0usize, |count, fragment| {
                let _ = fragment?;
                Ok(count.saturating_add(1))
            })
    }

    /// Stream children in source order without retaining a width-sized index.
    pub fn children(&self) -> ChildNodes<'source> {
        ChildNodes {
            fields: self.length_delimited_fields(GROUP_CHILD_FIELD),
            options: self.options,
        }
    }

    /// Merge and read the optional category value, if present.
    ///
    /// Every selected scalar wrapper is decoded during this call. A malformed
    /// projected sibling therefore invalidates the whole category value
    /// instead of being hidden by label-precedence short-circuiting.
    pub fn category_value(&self) -> Result<Option<CategoryValueView<'source>>, DecodeError> {
        let mut found = false;
        let mut value = CategoryValueView::default();
        let fragment_options = self.options.descend()?;
        for fragment in self.length_delimited_fields(GROUP_CELL_VALUE_FIELD) {
            found = true;
            value.merge(fragment?, fragment_options)?;
        }
        Ok(found.then_some(value))
    }

    /// Number of field-7 category-value fragments in source order.
    pub fn category_value_occurrences(&self) -> Result<usize, DecodeError> {
        self.length_delimited_fields(GROUP_CELL_VALUE_FIELD)
            .try_fold(0usize, |count, fragment| {
                let _ = fragment?;
                Ok(count.saturating_add(1))
            })
    }

    fn length_delimited_fields(&self, field_number: u32) -> LengthDelimitedFields<'source> {
        LengthDelimitedFields {
            fields: MessageFields::new(self.source, self.options.recursion_limit),
            field_number,
        }
    }
}

/// Streaming source-order iterator over `GroupNode` children.
#[derive(Debug)]
#[must_use]
pub struct ChildNodes<'source> {
    fields: LengthDelimitedFields<'source>,
    options: DecodeOptions,
}

impl<'source> Iterator for ChildNodes<'source> {
    type Item = Result<GroupNodeView<'source>, DecodeError>;

    fn next(&mut self) -> Option<Self::Item> {
        self.fields
            .next()
            .map(|fragment| decode_group_node_with_options(fragment?, self.options.descend()?))
    }
}

impl std::iter::FusedIterator for ChildNodes<'_> {}

/// Borrowed, merged scalar category branches with no generated type in its
/// public surface.
#[derive(Debug, Default)]
pub struct CategoryValueView<'source> {
    boolean_present: bool,
    boolean: Option<bool>,
    date_present: bool,
    date: Option<f64>,
    number: Option<f64>,
    string_present: bool,
    string: Option<&'source str>,
}

impl<'source> CategoryValueView<'source> {
    /// Read the merged `Boolean` branch, if present.
    pub const fn boolean(&self) -> Result<Option<bool>, DecodeError> {
        Ok(if self.boolean_present {
            Some(match self.boolean {
                Some(value) => value,
                None => false,
            })
        } else {
            None
        })
    }

    /// Read the merged `Date` branch, if present.
    pub fn date(&self) -> Result<Option<f64>, DecodeError> {
        Ok(if self.date_present {
            Some(self.date.unwrap_or(0.0))
        } else {
            None
        })
    }

    /// Read the merged optional `Number` scalar, if present.
    pub const fn number(&self) -> Result<Option<f64>, DecodeError> {
        Ok(self.number)
    }

    /// Read the merged `String` branch, if present.
    pub const fn string(&self) -> Result<Option<&'source str>, DecodeError> {
        Ok(if self.string_present {
            Some(match self.string {
                Some(value) => value,
                None => "",
            })
        } else {
            None
        })
    }

    fn merge(&mut self, source: &'source [u8], options: DecodeOptions) -> Result<(), DecodeError> {
        let wrapper_options = options.descend()?;
        for field_result in MessageFields::new(source, options.recursion_limit) {
            let field = field_result?;
            match field.tag.field_number() {
                BOOLEAN_VALUE_FIELD => {
                    let payload = field.length_delimited()?;
                    let view: projection::BooleanCellValueLazyView<'source> =
                        wrapper_options.buffa().decode_lazy_view(payload)?;
                    self.boolean_present = true;
                    if view.has_value() {
                        self.boolean = Some(view.value);
                    }
                },
                DATE_VALUE_FIELD => {
                    let payload = field.length_delimited()?;
                    let view: projection::DateCellValueLazyView<'source> =
                        wrapper_options.buffa().decode_lazy_view(payload)?;
                    self.date_present = true;
                    if view.has_value() {
                        self.date = Some(view.value);
                    }
                },
                NUMBER_VALUE_FIELD => {
                    let payload = field.length_delimited()?;
                    let view: projection::NumberCellValueLazyView<'source> =
                        wrapper_options.buffa().decode_lazy_view(payload)?;
                    if let Some(number) = view.value {
                        self.number = Some(number);
                    }
                },
                STRING_VALUE_FIELD => {
                    let payload = field.length_delimited()?;
                    let view: projection::StringCellValueLazyView<'source> =
                        wrapper_options.buffa().decode_lazy_view(payload)?;
                    self.string_present = true;
                    if view.has_value() {
                        self.string = Some(view.value);
                    }
                },
                _ => {},
            }
        }
        Ok(())
    }
}

#[derive(Clone, Copy, Debug)]
struct Field<'source> {
    tag: buffa::encoding::Tag,
    payload: Option<&'source [u8]>,
}

impl<'source> Field<'source> {
    fn length_delimited(self) -> Result<&'source [u8], DecodeError> {
        buffa::encoding::check_wire_type(self.tag, buffa::encoding::WireType::LengthDelimited)?;
        self.payload.ok_or_else(|| {
            DecodeError(buffa::DecodeError::WireTypeMismatch {
                field_number: self.tag.field_number(),
                expected: buffa::encoding::WireType::LengthDelimited as u8,
                actual: self.tag.wire_type() as u8,
            })
        })
    }
}

#[derive(Clone, Debug)]
struct MessageFields<'source> {
    remaining: &'source [u8],
    recursion_limit: u32,
}

impl<'source> MessageFields<'source> {
    const fn new(source: &'source [u8], recursion_limit: u32) -> Self {
        Self {
            remaining: source,
            recursion_limit,
        }
    }

    fn next_field(&mut self) -> Result<Option<Field<'source>>, DecodeError> {
        if self.remaining.is_empty() {
            return Ok(None);
        }
        let tag = buffa::encoding::Tag::decode(&mut self.remaining)?;
        let payload = if tag.wire_type() == buffa::encoding::WireType::LengthDelimited {
            let encoded_length = buffa::encoding::decode_varint(&mut self.remaining)?;
            let length = usize::try_from(encoded_length)
                .map_err(|_error| buffa::DecodeError::MessageTooLarge)?;
            if self.remaining.len() < length {
                return Err(buffa::DecodeError::UnexpectedEof.into());
            }
            let (payload, rest) = self.remaining.split_at(length);
            self.remaining = rest;
            Some(payload)
        } else {
            buffa::encoding::skip_field_depth(tag, &mut self.remaining, self.recursion_limit)?;
            None
        };
        Ok(Some(Field { tag, payload }))
    }
}

impl<'source> Iterator for MessageFields<'source> {
    type Item = Result<Field<'source>, DecodeError>;

    fn next(&mut self) -> Option<Self::Item> {
        match self.next_field() {
            Ok(Some(field)) => Some(Ok(field)),
            Ok(None) => None,
            Err(error) => {
                self.remaining = &[];
                Some(Err(error))
            },
        }
    }
}

impl std::iter::FusedIterator for MessageFields<'_> {}

#[derive(Debug)]
struct LengthDelimitedFields<'source> {
    fields: MessageFields<'source>,
    field_number: u32,
}

impl<'source> Iterator for LengthDelimitedFields<'source> {
    type Item = Result<&'source [u8], DecodeError>;

    fn next(&mut self) -> Option<Self::Item> {
        loop {
            let field = match self.fields.next()? {
                Ok(field) => field,
                Err(error) => return Some(Err(error)),
            };
            if field.tag.field_number() == self.field_number {
                return Some(field.length_delimited());
            }
        }
    }
}

impl std::iter::FusedIterator for LengthDelimitedFields<'_> {}

/// Decode one already-preflighted `GroupNode` category envelope.
///
/// Buffa validates the node envelope with an empty generated shell. UUID and
/// scalar fragments are decoded lazily on access, while recursive children are
/// streamed directly from source bytes. Generated Buffa types remain private.
pub fn decode_group_node(
    source: &[u8],
    options: DecodeOptions,
) -> Result<GroupNodeView<'_>, DecodeError> {
    decode_group_node_with_options(source, options)
}

fn decode_group_node_with_options(
    source: &[u8],
    options: DecodeOptions,
) -> Result<GroupNodeView<'_>, DecodeError> {
    let _view: projection::GroupNodeCategoryLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    Ok(GroupNodeView { source, options })
}

#[cfg(test)]
mod tests {
    use prost::Message as _;

    use super::*;
    use crate::tst;

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::new(source.len().max(1), 32)
    }

    #[allow(
        clippy::cast_possible_truncation,
        reason = "Each emitted varint byte intentionally contains only the low seven bits."
    )]
    fn push_varint(mut value: u64, output: &mut Vec<u8>) {
        while value >= 0x80 {
            output.push((value as u8 & 0x7f) | 0x80);
            value >>= 7;
        }
        output.push(value as u8);
    }

    fn varint_field(field_number: u32, value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(u64::from(field_number) << 3, &mut output);
        push_varint(value, &mut output);
        output
    }

    fn fixed64_field(field_number: u32, value: f64) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint((u64::from(field_number) << 3) | 1, &mut output);
        output.extend_from_slice(&value.to_le_bytes());
        output
    }

    fn length_delimited_field(field_number: u32, payload: &[u8]) -> Vec<u8> {
        let mut output = Vec::with_capacity(payload.len() + 8);
        push_varint((u64::from(field_number) << 3) | 2, &mut output);
        push_varint(payload.len() as u64, &mut output);
        output.extend_from_slice(payload);
        output
    }

    fn uuid(lower: u64, upper: u64) -> Vec<u8> {
        let mut output = varint_field(1, lower);
        output.extend(varint_field(2, upper));
        output
    }

    fn group_node(lower: u64, upper: u64) -> Vec<u8> {
        length_delimited_field(1, &uuid(lower, upper))
    }

    fn present<T>(value: Option<T>, message: &'static str) -> Result<T, std::io::Error> {
        value.ok_or_else(|| std::io::Error::other(message))
    }

    #[test]
    fn preserves_child_order_and_matches_prost_last_wins() -> Result<(), Box<dyn std::error::Error>>
    {
        let mut first_cell_value = varint_field(1, 2);
        first_cell_value.extend(length_delimited_field(2, &varint_field(1, 1)));
        first_cell_value.extend(length_delimited_field(3, &fixed64_field(1, 1.0)));
        first_cell_value.extend(length_delimited_field(4, &fixed64_field(1, 3.0)));
        first_cell_value.extend(length_delimited_field(
            5,
            &length_delimited_field(1, b"first"),
        ));
        let mut last_cell_value = varint_field(1, 2);
        last_cell_value.extend(length_delimited_field(2, &[]));
        last_cell_value.extend(length_delimited_field(3, &fixed64_field(1, 2.0)));
        last_cell_value.extend(length_delimited_field(4, &fixed64_field(1, 4.0)));
        last_cell_value.extend(length_delimited_field(
            5,
            &length_delimited_field(1, b"second"),
        ));

        let mut source = length_delimited_field(1, &uuid(1, 2));
        source.extend(length_delimited_field(1, &varint_field(1, 7)));
        source.extend(length_delimited_field(3, &group_node(11, 12)));
        source.extend(length_delimited_field(3, &group_node(21, 22)));
        source.extend(length_delimited_field(7, &first_cell_value));
        source.extend(length_delimited_field(7, &last_cell_value));

        let native = tst::group_by_archive::GroupNodeArchive::decode(source.as_slice())?;
        assert_eq!((native.group_uid.lower, native.group_uid.upper), (7, 2));
        assert_eq!(
            native
                .child
                .iter()
                .map(|child| (child.group_uid.lower, child.group_uid.upper))
                .collect::<Vec<_>>(),
            [(11, 12), (21, 22)]
        );
        let native_value = present(native.group_cell_value, "cell value is present")?;
        assert_eq!(
            native_value.boolean_value.map(|value| value.value),
            Some(true)
        );
        assert_eq!(native_value.date_value.map(|value| value.value), Some(2.0));
        assert_eq!(
            native_value.number_value.and_then(|value| value.value),
            Some(4.0)
        );
        assert_eq!(
            native_value.string_value.map(|value| value.value),
            Some(String::from("second"))
        );

        let view = decode_group_node(&source, options(&source))?;
        assert_eq!(view.group_uid_occurrences()?, 2);
        assert_eq!(view.child_count()?, 2);
        let mut children = view.children();
        let first = present(children.next(), "first child")??;
        assert_eq!(
            first.group_uid()?,
            Some(GroupUid {
                lower: 11,
                upper: 12
            })
        );
        let second = present(children.next(), "second child")??;
        assert_eq!(
            second.group_uid()?,
            Some(GroupUid {
                lower: 21,
                upper: 22
            })
        );
        assert!(children.next().is_none());

        assert_eq!(view.group_uid()?, Some(GroupUid { lower: 7, upper: 2 }));
        assert_eq!(view.category_value_occurrences()?, 2);
        let value = present(view.category_value()?, "category value is present")?;
        assert_eq!(value.boolean()?, Some(true));
        assert_eq!(value.date()?, Some(2.0));
        assert_eq!(value.number()?, Some(4.0));
        assert_eq!(value.string()?, Some("second"));
        Ok(())
    }

    #[test]
    fn child_decode_is_deferred_until_iteration() -> Result<(), Box<dyn std::error::Error>> {
        let source = length_delimited_field(3, &[0x0a]);
        let view = decode_group_node(&source, options(&source))?;

        assert_eq!(view.child_count()?, 1);
        let child = present(view.children().next(), "child is present")?;
        assert!(child.is_err());
        Ok(())
    }

    #[test]
    fn malformed_projected_sibling_invalidates_category_value()
    -> Result<(), Box<dyn std::error::Error>> {
        let mut cell_value =
            length_delimited_field(5, &length_delimited_field(1, b"otherwise valid"));
        cell_value.extend(varint_field(4, 1));
        let source = length_delimited_field(7, &cell_value);
        let view = decode_group_node(&source, options(&source))?;

        assert!(view.category_value().is_err());
        Ok(())
    }

    #[test]
    fn omitted_groupnode_fields_remain_opaque() -> Result<(), Box<dyn std::error::Error>> {
        let mut source = length_delimited_field(1, &uuid(4, 5));
        source.extend(length_delimited_field(4, &[0x0a]));
        source.extend(length_delimited_field(5, &[0x0a]));
        source.extend(length_delimited_field(6, &[0x0a]));
        let mut cell_value = varint_field(1, 6);
        cell_value.extend(length_delimited_field(6, &[0x0a]));
        source.extend(length_delimited_field(7, &cell_value));

        let view = decode_group_node(&source, options(&source))?;
        assert_eq!(view.group_uid()?, Some(GroupUid { lower: 4, upper: 5 }));
        assert_eq!(view.child_count()?, 0);
        let value = present(view.category_value()?, "category value is present")?;
        assert_eq!(value.boolean()?, None);
        assert_eq!(value.date()?, None);
        assert_eq!(value.number()?, None);
        assert_eq!(value.string()?, None);
        Ok(())
    }

    #[test]
    fn wide_children_stream_in_source_order() -> Result<(), Box<dyn std::error::Error>> {
        const CHILDREN: u64 = 4_096;
        let mut source = Vec::new();
        for lower in 0..CHILDREN {
            source.extend(length_delimited_field(3, &group_node(lower, 0)));
        }
        let view = decode_group_node(&source, options(&source))?;
        assert_eq!(view.child_count()?, usize::try_from(CHILDREN)?);
        for (expected, child) in (0..CHILDREN).zip(view.children()) {
            assert_eq!(
                child?.group_uid()?,
                Some(GroupUid {
                    lower: expected,
                    upper: 0
                })
            );
        }
        Ok(())
    }

    #[test]
    fn source_byte_limit_is_enforced() {
        let source = length_delimited_field(1, &uuid(1, 2));
        let options = DecodeOptions::new(source.len() - 1, 1);

        assert!(decode_group_node(&source, options).is_err());
    }
}
