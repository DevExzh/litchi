//! Strongly typed native predicate identifiers shared by encoders and decoders.

use super::*;
use crate::table_cell_conditional_highlight::{
    TableCellConditionalHighlightCondition, TableCellConditionalHighlightNumber,
};

pub(super) const PREDICATE_QUALIFIER_NONE: i32 = 0;
pub(super) const PREDICATE_CELL_ARGUMENT_INDEX: i32 = 0;
pub(super) const PREDICATE_NUMBER_ARGUMENT_INDEX: i32 = 1;
pub(super) const PREDICATE_UNUSED_ARGUMENT_INDEX: i32 = -1;
pub(super) const PREDICATE_ARGUMENT_NONE: i32 = 0;
pub(super) const PREDICATE_ARGUMENT_NUMBER: i32 = 1;
pub(super) const PREDICATE_ARGUMENT_RELATIVE_CELL: i32 = 4;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(i32)]
pub(super) enum NumericPredicateKind {
    EqualTo = 5,
    NotEqualTo = 6,
    GreaterThan = 7,
    GreaterThanOrEqualTo = 8,
    LessThan = 9,
    LessThanOrEqualTo = 10,
}

impl NumericPredicateKind {
    pub(super) const fn from_condition(condition: TableCellConditionalHighlightCondition) -> Self {
        match condition {
            TableCellConditionalHighlightCondition::EqualTo(_) => Self::EqualTo,
            TableCellConditionalHighlightCondition::NotEqualTo(_) => Self::NotEqualTo,
            TableCellConditionalHighlightCondition::GreaterThan(_) => Self::GreaterThan,
            TableCellConditionalHighlightCondition::GreaterThanOrEqualTo(_) => {
                Self::GreaterThanOrEqualTo
            },
            TableCellConditionalHighlightCondition::LessThan(_) => Self::LessThan,
            TableCellConditionalHighlightCondition::LessThanOrEqualTo(_) => Self::LessThanOrEqualTo,
        }
    }

    pub(super) const fn native_value(self) -> i32 {
        self as i32
    }

    pub(super) const fn condition(
        self,
        number: TableCellConditionalHighlightNumber,
    ) -> TableCellConditionalHighlightCondition {
        match self {
            Self::EqualTo => TableCellConditionalHighlightCondition::EqualTo(number),
            Self::NotEqualTo => TableCellConditionalHighlightCondition::NotEqualTo(number),
            Self::GreaterThan => TableCellConditionalHighlightCondition::GreaterThan(number),
            Self::GreaterThanOrEqualTo => {
                TableCellConditionalHighlightCondition::GreaterThanOrEqualTo(number)
            },
            Self::LessThan => TableCellConditionalHighlightCondition::LessThan(number),
            Self::LessThanOrEqualTo => {
                TableCellConditionalHighlightCondition::LessThanOrEqualTo(number)
            },
        }
    }

    pub(super) const fn ast_node_type(self) -> tsce::ast_node_array_archive::AstNodeType {
        use tsce::ast_node_array_archive::AstNodeType;

        match self {
            Self::EqualTo => AstNodeType::EqualToNode,
            Self::NotEqualTo => AstNodeType::NotEqualToNode,
            Self::GreaterThan => AstNodeType::GreaterThanNode,
            Self::GreaterThanOrEqualTo => AstNodeType::GreaterThanOrEqualToNode,
            Self::LessThan => AstNodeType::LessThanNode,
            Self::LessThanOrEqualTo => AstNodeType::LessThanOrEqualToNode,
        }
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
        ];
        for kind in kinds {
            assert_eq!(
                NumericPredicateKind::try_from(kind.native_value()).unwrap(),
                kind
            );
        }
        assert!(NumericPredicateKind::try_from(i32::MIN).is_err());
    }
}
