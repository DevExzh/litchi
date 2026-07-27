//! Strongly typed native predicate identifiers shared by encoders and decoders.

use super::*;
use crate::table_cell_conditional_highlight::{
    TableCellConditionalHighlightCondition, TableCellConditionalHighlightNumber,
    TableCellConditionalHighlightText,
};

pub(super) const PREDICATE_QUALIFIER_NONE: i32 = 0;
pub(super) const PREDICATE_CELL_ARGUMENT_INDEX: i32 = 0;
pub(super) const PREDICATE_NUMBER_ARGUMENT_INDEX: i32 = 1;
pub(super) const PREDICATE_UNUSED_ARGUMENT_INDEX: i32 = -1;
pub(super) const PREDICATE_RANGE_LOWER_ARGUMENT_INDEX: i32 = 0;
pub(super) const PREDICATE_RANGE_UPPER_ARGUMENT_INDEX: i32 = 1;
pub(super) const PREDICATE_RANGE_CELL_ARGUMENT_INDEX: i32 = 3;
pub(super) const PREDICATE_TEXT_ARGUMENT_INDEX: i32 = 0;
pub(super) const PREDICATE_TEXT_CELL_ARGUMENT_INDEX: i32 = 1;
pub(super) const PREDICATE_ARGUMENT_NONE: i32 = 0;
pub(super) const PREDICATE_ARGUMENT_NUMBER: i32 = 1;
pub(super) const PREDICATE_ARGUMENT_STRING: i32 = 3;
pub(super) const PREDICATE_ARGUMENT_RELATIVE_CELL: i32 = 4;
pub(super) const LOGICAL_AND_FUNCTION_INDEX: u32 = 7;
pub(super) const CONDITIONAL_FUNCTION_INDEX: u32 = 62;
pub(super) const LOGICAL_OR_FUNCTION_INDEX: u32 = 102;
pub(super) const BINARY_FUNCTION_ARGUMENT_COUNT: u32 = 2;
pub(super) const CONDITIONAL_FUNCTION_ARGUMENT_COUNT: u32 = 3;
pub(super) const TEXT_SEARCH_FUNCTION_INDEX: u32 = 296;
pub(super) const IS_ERROR_FUNCTION_INDEX: u32 = 70;
pub(super) const LOGICAL_NOT_FUNCTION_INDEX: u32 = 96;
pub(super) const UNARY_FUNCTION_ARGUMENT_COUNT: u32 = 1;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(i32)]
pub(super) enum NumericPredicateKind {
    EqualTo = 5,
    NotEqualTo = 6,
    GreaterThan = 7,
    GreaterThanOrEqualTo = 8,
    LessThan = 9,
    LessThanOrEqualTo = 10,
    Between = 13,
    NotBetween = 32,
}

impl NumericPredicateKind {
    pub(super) const fn from_condition(
        condition: &TableCellConditionalHighlightCondition,
    ) -> Option<Self> {
        match condition {
            TableCellConditionalHighlightCondition::EqualTo(_) => Some(Self::EqualTo),
            TableCellConditionalHighlightCondition::NotEqualTo(_) => Some(Self::NotEqualTo),
            TableCellConditionalHighlightCondition::GreaterThan(_) => Some(Self::GreaterThan),
            TableCellConditionalHighlightCondition::GreaterThanOrEqualTo(_) => {
                Some(Self::GreaterThanOrEqualTo)
            },
            TableCellConditionalHighlightCondition::LessThan(_) => Some(Self::LessThan),
            TableCellConditionalHighlightCondition::LessThanOrEqualTo(_) => {
                Some(Self::LessThanOrEqualTo)
            },
            TableCellConditionalHighlightCondition::Between(_) => Some(Self::Between),
            TableCellConditionalHighlightCondition::NotBetween(_) => Some(Self::NotBetween),
            TableCellConditionalHighlightCondition::TextContains(_) => None,
        }
    }

    pub(super) const fn native_value(self) -> i32 {
        self as i32
    }

    pub(super) const fn single_condition(
        self,
        number: TableCellConditionalHighlightNumber,
    ) -> Option<TableCellConditionalHighlightCondition> {
        match self {
            Self::EqualTo => Some(TableCellConditionalHighlightCondition::EqualTo(number)),
            Self::NotEqualTo => Some(TableCellConditionalHighlightCondition::NotEqualTo(number)),
            Self::GreaterThan => Some(TableCellConditionalHighlightCondition::GreaterThan(number)),
            Self::GreaterThanOrEqualTo => {
                Some(TableCellConditionalHighlightCondition::GreaterThanOrEqualTo(number))
            },
            Self::LessThan => Some(TableCellConditionalHighlightCondition::LessThan(number)),
            Self::LessThanOrEqualTo => Some(
                TableCellConditionalHighlightCondition::LessThanOrEqualTo(number),
            ),
            Self::Between | Self::NotBetween => None,
        }
    }

    pub(super) const fn single_ast_node_type(
        self,
    ) -> Option<tsce::ast_node_array_archive::AstNodeType> {
        use tsce::ast_node_array_archive::AstNodeType;

        match self {
            Self::EqualTo => Some(AstNodeType::EqualToNode),
            Self::NotEqualTo => Some(AstNodeType::NotEqualToNode),
            Self::GreaterThan => Some(AstNodeType::GreaterThanNode),
            Self::GreaterThanOrEqualTo => Some(AstNodeType::GreaterThanOrEqualToNode),
            Self::LessThan => Some(AstNodeType::LessThanNode),
            Self::LessThanOrEqualTo => Some(AstNodeType::LessThanOrEqualToNode),
            Self::Between | Self::NotBetween => None,
        }
    }

    pub(super) const fn is_range(self) -> bool {
        matches!(self, Self::Between | Self::NotBetween)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(i32)]
pub(super) enum TextPredicateKind {
    Contains = 3,
}

impl TextPredicateKind {
    pub(super) const fn native_value(self) -> i32 {
        self as i32
    }

    pub(super) fn condition(
        self,
        text: TableCellConditionalHighlightText,
    ) -> TableCellConditionalHighlightCondition {
        match self {
            Self::Contains => TableCellConditionalHighlightCondition::TextContains(text),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum NativePredicateKind {
    Numeric(NumericPredicateKind),
    Text(TextPredicateKind),
}

impl NativePredicateKind {
    pub(super) fn from_condition(condition: &TableCellConditionalHighlightCondition) -> Self {
        NumericPredicateKind::from_condition(condition).map_or_else(
            || match condition {
                TableCellConditionalHighlightCondition::TextContains(_) => {
                    Self::Text(TextPredicateKind::Contains)
                },
                _ => unreachable!("every public predicate has a native kind"),
            },
            Self::Numeric,
        )
    }

    pub(super) const fn native_value(self) -> i32 {
        match self {
            Self::Numeric(kind) => kind.native_value(),
            Self::Text(kind) => kind.native_value(),
        }
    }
}

impl TryFrom<i32> for NativePredicateKind {
    type Error = Error;

    fn try_from(value: i32) -> Result<Self> {
        if value == TextPredicateKind::Contains.native_value() {
            return Ok(Self::Text(TextPredicateKind::Contains));
        }
        NumericPredicateKind::try_from(value).map(Self::Numeric)
    }
}

impl TryFrom<i32> for NumericPredicateKind {
    type Error = Error;

    fn try_from(value: i32) -> Result<Self> {
        match value {
            value if value == Self::EqualTo.native_value() => Ok(Self::EqualTo),
            value if value == Self::NotEqualTo.native_value() => Ok(Self::NotEqualTo),
            value if value == Self::GreaterThan.native_value() => Ok(Self::GreaterThan),
            value if value == Self::GreaterThanOrEqualTo.native_value() => {
                Ok(Self::GreaterThanOrEqualTo)
            },
            value if value == Self::LessThan.native_value() => Ok(Self::LessThan),
            value if value == Self::LessThanOrEqualTo.native_value() => Ok(Self::LessThanOrEqualTo),
            value if value == Self::Between.native_value() => Ok(Self::Between),
            value if value == Self::NotBetween.native_value() => Ok(Self::NotBetween),
            _ => Err(Error::InvalidFormat(
                "iWork conditional-highlight rule uses an unsupported predicate".to_owned(),
            )),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn native_numeric_predicate_values_are_unique_and_reversible() {
        let kinds = [
            NumericPredicateKind::EqualTo,
            NumericPredicateKind::NotEqualTo,
            NumericPredicateKind::GreaterThan,
            NumericPredicateKind::GreaterThanOrEqualTo,
            NumericPredicateKind::LessThan,
            NumericPredicateKind::LessThanOrEqualTo,
            NumericPredicateKind::Between,
            NumericPredicateKind::NotBetween,
        ];
        for kind in kinds {
            assert_eq!(
                NumericPredicateKind::try_from(kind.native_value()).unwrap(),
                kind
            );
        }
        assert!(NumericPredicateKind::try_from(i32::MIN).is_err());
    }

    #[test]
    fn native_text_predicate_values_are_reversible() {
        let kind = NativePredicateKind::Text(TextPredicateKind::Contains);
        assert_eq!(
            NativePredicateKind::try_from(kind.native_value()).unwrap(),
            kind
        );
    }
}
