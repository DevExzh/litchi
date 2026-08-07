//! Strongly typed native predicate identifiers shared by encoders and decoders.

use super::*;
use litchi_iwa_common::table::cell::conditional_highlight::{
    Condition, Date, DateRange, Number, Offset, Period, Text,
};

pub(super) const PREDICATE_QUALIFIER_NONE: i32 = 0;
pub(super) const PREDICATE_CELL_ARGUMENT_INDEX: i32 = 0;
pub(super) const PREDICATE_NUMBER_ARGUMENT_INDEX: i32 = 1;
pub(super) const PREDICATE_DATE_ARGUMENT_INDEX: i32 = 1;
pub(super) const PREDICATE_UNUSED_ARGUMENT_INDEX: i32 = -1;
pub(super) const PREDICATE_RANGE_LOWER_ARGUMENT_INDEX: i32 = 0;
pub(super) const PREDICATE_RANGE_UPPER_ARGUMENT_INDEX: i32 = 1;
pub(super) const PREDICATE_RANGE_CELL_ARGUMENT_INDEX: i32 = 3;
pub(super) const PREDICATE_TEXT_ARGUMENT_INDEX: i32 = 0;
pub(super) const PREDICATE_TEXT_CELL_ARGUMENT_INDEX: i32 = 1;
pub(super) const PREDICATE_TEXT_EQUALITY_CELL_ARGUMENT_INDEX: i32 = 2;
pub(super) const PREDICATE_DATE_EQUALITY_CELL_ARGUMENT_INDEX: i32 = 2;
pub(super) const PREDICATE_DATE_EQUALITY_ARGUMENT_INDEX: i32 = 0;
pub(super) const PREDICATE_ARGUMENT_NONE: i32 = 0;
pub(super) const PREDICATE_ARGUMENT_NUMBER: i32 = 1;
pub(super) const PREDICATE_ARGUMENT_DATE: i32 = 2;
pub(super) const PREDICATE_ARGUMENT_STRING: i32 = 3;
pub(super) const PREDICATE_ARGUMENT_RELATIVE_CELL: i32 = 4;
pub(super) const LOGICAL_AND_FUNCTION_INDEX: u32 = 7;
pub(super) const CONDITIONAL_FUNCTION_INDEX: u32 = 62;
pub(super) const LOGICAL_OR_FUNCTION_INDEX: u32 = 102;
pub(super) const BINARY_FUNCTION_ARGUMENT_COUNT: u32 = 2;
pub(super) const TERNARY_FUNCTION_ARGUMENT_COUNT: u32 = 3;
pub(super) const CONDITIONAL_FUNCTION_ARGUMENT_COUNT: u32 = 3;
pub(super) const TEXT_SEARCH_FUNCTION_INDEX: u32 = 296;
pub(super) const TEXT_LENGTH_FUNCTION_INDEX: u32 = 77;
pub(super) const TEXT_RIGHT_FUNCTION_INDEX: u32 = 124;
pub(super) const IS_ERROR_FUNCTION_INDEX: u32 = 70;
pub(super) const LOGICAL_NOT_FUNCTION_INDEX: u32 = 96;
pub(super) const IF_ERROR_FUNCTION_INDEX: u32 = 235;
pub(super) const IS_BLANK_FUNCTION_INDEX: u32 = 69;
pub(super) const IS_NUMBER_FUNCTION_INDEX: u32 = 304;
pub(super) const CELL_DATA_FORMAT_FUNCTION_INDEX: u32 = 326;
pub(super) const VALUE_TYPE_FUNCTION_INDEX: u32 = 327;
pub(super) const CHECKBOX_DATA_FORMAT_CODE: f64 = 8.0;
pub(super) const BOOLEAN_VALUE_TYPE_CODE: f64 = 6.0;
pub(super) const UNARY_FUNCTION_ARGUMENT_COUNT: u32 = 1;
pub(super) const ZERO_FUNCTION_ARGUMENT_COUNT: u32 = 0;
pub(super) const TODAY_FUNCTION_INDEX: u32 = 154;
pub(super) const DATE_DAY_FUNCTION_INDEX: u32 = 41;
pub(super) const DATE_DIFFERENCE_FUNCTION_INDEX: u32 = 40;
pub(super) const DATE_ADD_MONTHS_FUNCTION_INDEX: u32 = 47;
pub(super) const DATE_MONTH_FUNCTION_INDEX: u32 = 94;
pub(super) const DATE_YEAR_FUNCTION_INDEX: u32 = 167;
pub(super) const DATE_DURATION_FROM_WEEKS_DAYS_FUNCTION_INDEX: u32 = 212;

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
    pub(super) const fn from_condition(condition: &Condition) -> Option<Self> {
        match condition {
            Condition::CellIsBlank
            | Condition::CellIsNotBlank
            | Condition::CheckboxIsChecked
            | Condition::CheckboxIsNotChecked
            | Condition::BooleanIsTrue
            | Condition::BooleanIsFalse
            | Condition::NumberIsPositive
            | Condition::NumberIsNegative
            | Condition::DateIsToday
            | Condition::DateIsYesterday
            | Condition::DateIsTomorrow
            | Condition::DateIs(_)
            | Condition::DateIsBefore(_)
            | Condition::DateIsAfter(_)
            | Condition::DateIsBetween(_)
            | Condition::DateIsInNext(_)
            | Condition::DateIsInLast(_)
            | Condition::DateIsOffsetFromToday(_) => None,
            Condition::EqualTo(_) => Some(Self::EqualTo),
            Condition::NotEqualTo(_) => Some(Self::NotEqualTo),
            Condition::GreaterThan(_) => Some(Self::GreaterThan),
            Condition::GreaterThanOrEqualTo(_) => Some(Self::GreaterThanOrEqualTo),
            Condition::LessThan(_) => Some(Self::LessThan),
            Condition::LessThanOrEqualTo(_) => Some(Self::LessThanOrEqualTo),
            Condition::Between(_) => Some(Self::Between),
            Condition::NotBetween(_) => Some(Self::NotBetween),
            Condition::TextEqualTo(_)
            | Condition::TextNotEqualTo(_)
            | Condition::TextStartsWith(_)
            | Condition::TextDoesNotStartWith(_)
            | Condition::TextEndsWith(_)
            | Condition::TextDoesNotEndWith(_)
            | Condition::TextContains(_)
            | Condition::TextDoesNotContain(_) => None,
        }
    }

    pub(super) const fn native_value(self) -> i32 {
        self as i32
    }

    pub(super) const fn single_condition(self, number: Number) -> Option<Condition> {
        match self {
            Self::EqualTo => Some(Condition::EqualTo(number)),
            Self::NotEqualTo => Some(Condition::NotEqualTo(number)),
            Self::GreaterThan => Some(Condition::GreaterThan(number)),
            Self::GreaterThanOrEqualTo => Some(Condition::GreaterThanOrEqualTo(number)),
            Self::LessThan => Some(Condition::LessThan(number)),
            Self::LessThanOrEqualTo => Some(Condition::LessThanOrEqualTo(number)),
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
pub(super) enum CellPredicateKind {
    IsBlank = 34,
    IsNotBlank = 35,
}

impl CellPredicateKind {
    pub(super) const fn native_value(self) -> i32 {
        self as i32
    }

    pub(super) const fn condition(self) -> Condition {
        match self {
            Self::IsBlank => Condition::CellIsBlank,
            Self::IsNotBlank => Condition::CellIsNotBlank,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(i32)]
pub(super) enum NumericSignPredicateKind {
    IsPositive = 57,
    IsNegative = 58,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(i32)]
pub(super) enum BooleanPredicateKind {
    IsTrue = 59,
    IsFalse = 60,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(i32)]
pub(super) enum CheckboxPredicateKind {
    IsChecked = 55,
    IsNotChecked = 56,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(i32)]
pub(super) enum RelativeDatePredicateKind {
    Today = 17,
    Yesterday = 18,
    Tomorrow = 19,
}

impl RelativeDatePredicateKind {
    pub(super) const fn native_value(self) -> i32 {
        self as i32
    }

    pub(super) const fn condition(self) -> Condition {
        match self {
            Self::Today => Condition::DateIsToday,
            Self::Yesterday => Condition::DateIsYesterday,
            Self::Tomorrow => Condition::DateIsTomorrow,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(i32)]
pub(super) enum FixedDatePredicateKind {
    Equal = 20,
    Before = 21,
    After = 22,
    Between = 23,
}

impl FixedDatePredicateKind {
    pub(super) const fn native_value(self) -> i32 {
        self as i32
    }

    pub(super) const fn condition(self, date: Date) -> Option<Condition> {
        match self {
            Self::Equal => Some(Condition::DateIs(date)),
            Self::Before => Some(Condition::DateIsBefore(date)),
            Self::After => Some(Condition::DateIsAfter(date)),
            Self::Between => None,
        }
    }

    pub(super) const fn range_condition(self, range: DateRange) -> Option<Condition> {
        match self {
            Self::Between => Some(Condition::DateIsBetween(range)),
            Self::Equal | Self::Before | Self::After => None,
        }
    }

    pub(super) const fn is_range(self) -> bool {
        matches!(self, Self::Between)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(i32)]
pub(super) enum DatePeriodPredicateKind {
    InNext = 24,
    InLast = 25,
    OffsetFromToday = 26,
}

impl DatePeriodPredicateKind {
    pub(super) const fn native_value(self) -> i32 {
        self as i32
    }

    pub(super) const fn period_condition(self, period: Period) -> Option<Condition> {
        match self {
            Self::InNext => Some(Condition::DateIsInNext(period)),
            Self::InLast => Some(Condition::DateIsInLast(period)),
            Self::OffsetFromToday => None,
        }
    }

    pub(super) const fn offset_condition(self, offset: Offset) -> Option<Condition> {
        match self {
            Self::OffsetFromToday => Some(Condition::DateIsOffsetFromToday(offset)),
            Self::InNext | Self::InLast => None,
        }
    }
}

impl CheckboxPredicateKind {
    pub(super) const fn native_value(self) -> i32 {
        self as i32
    }

    pub(super) const fn is_checked(self) -> bool {
        matches!(self, Self::IsChecked)
    }

    pub(super) const fn condition(self) -> Condition {
        match self {
            Self::IsChecked => Condition::CheckboxIsChecked,
            Self::IsNotChecked => Condition::CheckboxIsNotChecked,
        }
    }

    pub(super) const fn prepivot_kind(self) -> NumericPredicateKind {
        NumericPredicateKind::EqualTo
    }
}

impl BooleanPredicateKind {
    pub(super) const fn native_value(self) -> i32 {
        self as i32
    }

    pub(super) const fn value(self) -> bool {
        matches!(self, Self::IsTrue)
    }

    pub(super) const fn condition(self) -> Condition {
        match self {
            Self::IsTrue => Condition::BooleanIsTrue,
            Self::IsFalse => Condition::BooleanIsFalse,
        }
    }

    pub(super) const fn prepivot_kind(self) -> NumericPredicateKind {
        NumericPredicateKind::EqualTo
    }
}

impl NumericSignPredicateKind {
    pub(super) const fn native_value(self) -> i32 {
        self as i32
    }

    pub(super) const fn condition(self) -> Condition {
        match self {
            Self::IsPositive => Condition::NumberIsPositive,
            Self::IsNegative => Condition::NumberIsNegative,
        }
    }

    pub(super) const fn comparison(self) -> tsce::ast_node_array_archive::AstNodeType {
        use tsce::ast_node_array_archive::AstNodeType;

        match self {
            Self::IsPositive => AstNodeType::GreaterThanNode,
            Self::IsNegative => AstNodeType::LessThanNode,
        }
    }

    pub(super) const fn prepivot_kind(self) -> NumericPredicateKind {
        match self {
            Self::IsPositive => NumericPredicateKind::GreaterThan,
            Self::IsNegative => NumericPredicateKind::LessThan,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(i32)]
pub(super) enum TextPredicateKind {
    StartsWith = 1,
    EndsWith = 2,
    Contains = 3,
    DoesNotContain = 4,
    EqualTo = 36,
    NotEqualTo = 37,
    DoesNotStartWith = 61,
    DoesNotEndWith = 62,
}

impl TextPredicateKind {
    pub(super) const fn native_value(self) -> i32 {
        self as i32
    }

    pub(super) fn condition(self, text: Text) -> Condition {
        match self {
            Self::EqualTo => Condition::TextEqualTo(text),
            Self::NotEqualTo => Condition::TextNotEqualTo(text),
            Self::StartsWith => Condition::TextStartsWith(text),
            Self::DoesNotStartWith => Condition::TextDoesNotStartWith(text),
            Self::EndsWith => Condition::TextEndsWith(text),
            Self::DoesNotEndWith => Condition::TextDoesNotEndWith(text),
            Self::Contains => Condition::TextContains(text),
            Self::DoesNotContain => Condition::TextDoesNotContain(text),
        }
    }

    pub(super) const fn cell_argument_index(self) -> i32 {
        match self {
            Self::EqualTo | Self::NotEqualTo => PREDICATE_TEXT_EQUALITY_CELL_ARGUMENT_INDEX,
            Self::StartsWith
            | Self::DoesNotStartWith
            | Self::EndsWith
            | Self::DoesNotEndWith
            | Self::Contains
            | Self::DoesNotContain => PREDICATE_TEXT_CELL_ARGUMENT_INDEX,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum NativePredicateKind {
    Cell(CellPredicateKind),
    Checkbox(CheckboxPredicateKind),
    Boolean(BooleanPredicateKind),
    DatePeriod(DatePeriodPredicateKind),
    FixedDate(FixedDatePredicateKind),
    Numeric(NumericPredicateKind),
    NumericSign(NumericSignPredicateKind),
    RelativeDate(RelativeDatePredicateKind),
    Text(TextPredicateKind),
}

impl NativePredicateKind {
    pub(super) fn from_condition(condition: &Condition) -> Self {
        match condition {
            Condition::CellIsBlank => {
                return Self::Cell(CellPredicateKind::IsBlank);
            },
            Condition::CellIsNotBlank => {
                return Self::Cell(CellPredicateKind::IsNotBlank);
            },
            Condition::CheckboxIsChecked => {
                return Self::Checkbox(CheckboxPredicateKind::IsChecked);
            },
            Condition::CheckboxIsNotChecked => {
                return Self::Checkbox(CheckboxPredicateKind::IsNotChecked);
            },
            Condition::BooleanIsTrue => {
                return Self::Boolean(BooleanPredicateKind::IsTrue);
            },
            Condition::BooleanIsFalse => {
                return Self::Boolean(BooleanPredicateKind::IsFalse);
            },
            Condition::NumberIsPositive => {
                return Self::NumericSign(NumericSignPredicateKind::IsPositive);
            },
            Condition::NumberIsNegative => {
                return Self::NumericSign(NumericSignPredicateKind::IsNegative);
            },
            Condition::DateIsToday => {
                return Self::RelativeDate(RelativeDatePredicateKind::Today);
            },
            Condition::DateIsYesterday => {
                return Self::RelativeDate(RelativeDatePredicateKind::Yesterday);
            },
            Condition::DateIsTomorrow => {
                return Self::RelativeDate(RelativeDatePredicateKind::Tomorrow);
            },
            Condition::DateIs(_) => {
                return Self::FixedDate(FixedDatePredicateKind::Equal);
            },
            Condition::DateIsBefore(_) => {
                return Self::FixedDate(FixedDatePredicateKind::Before);
            },
            Condition::DateIsAfter(_) => {
                return Self::FixedDate(FixedDatePredicateKind::After);
            },
            Condition::DateIsBetween(_) => {
                return Self::FixedDate(FixedDatePredicateKind::Between);
            },
            Condition::DateIsInNext(_) => {
                return Self::DatePeriod(DatePeriodPredicateKind::InNext);
            },
            Condition::DateIsInLast(_) => {
                return Self::DatePeriod(DatePeriodPredicateKind::InLast);
            },
            Condition::DateIsOffsetFromToday(_) => {
                return Self::DatePeriod(DatePeriodPredicateKind::OffsetFromToday);
            },
            _ => {},
        }
        NumericPredicateKind::from_condition(condition).map_or_else(
            || match condition {
                Condition::TextEqualTo(_) => Self::Text(TextPredicateKind::EqualTo),
                Condition::TextNotEqualTo(_) => Self::Text(TextPredicateKind::NotEqualTo),
                Condition::TextStartsWith(_) => Self::Text(TextPredicateKind::StartsWith),
                Condition::TextDoesNotStartWith(_) => {
                    Self::Text(TextPredicateKind::DoesNotStartWith)
                },
                Condition::TextEndsWith(_) => Self::Text(TextPredicateKind::EndsWith),
                Condition::TextDoesNotEndWith(_) => Self::Text(TextPredicateKind::DoesNotEndWith),
                Condition::TextContains(_) => Self::Text(TextPredicateKind::Contains),
                Condition::TextDoesNotContain(_) => Self::Text(TextPredicateKind::DoesNotContain),
                _ => unreachable!("every public predicate has a native kind"),
            },
            Self::Numeric,
        )
    }

    pub(super) const fn native_value(self) -> i32 {
        match self {
            Self::Cell(kind) => kind.native_value(),
            Self::Checkbox(kind) => kind.native_value(),
            Self::Boolean(kind) => kind.native_value(),
            Self::DatePeriod(kind) => kind.native_value(),
            Self::FixedDate(kind) => kind.native_value(),
            Self::Numeric(kind) => kind.native_value(),
            Self::NumericSign(kind) => kind.native_value(),
            Self::RelativeDate(kind) => kind.native_value(),
            Self::Text(kind) => kind.native_value(),
        }
    }

    pub(super) const fn prepivot_native_value(self) -> i32 {
        match self {
            Self::NumericSign(kind) => kind.prepivot_kind().native_value(),
            Self::Checkbox(kind) => kind.prepivot_kind().native_value(),
            Self::Boolean(kind) => kind.prepivot_kind().native_value(),
            _ => self.native_value(),
        }
    }
}

impl TryFrom<i32> for NativePredicateKind {
    type Error = Error;

    fn try_from(value: i32) -> Result<Self> {
        match value {
            value if value == DatePeriodPredicateKind::InNext.native_value() => {
                return Ok(Self::DatePeriod(DatePeriodPredicateKind::InNext));
            },
            value if value == DatePeriodPredicateKind::InLast.native_value() => {
                return Ok(Self::DatePeriod(DatePeriodPredicateKind::InLast));
            },
            value if value == DatePeriodPredicateKind::OffsetFromToday.native_value() => {
                return Ok(Self::DatePeriod(DatePeriodPredicateKind::OffsetFromToday));
            },
            _ => {},
        }
        match value {
            value if value == FixedDatePredicateKind::Equal.native_value() => {
                return Ok(Self::FixedDate(FixedDatePredicateKind::Equal));
            },
            value if value == FixedDatePredicateKind::Before.native_value() => {
                return Ok(Self::FixedDate(FixedDatePredicateKind::Before));
            },
            value if value == FixedDatePredicateKind::After.native_value() => {
                return Ok(Self::FixedDate(FixedDatePredicateKind::After));
            },
            value if value == FixedDatePredicateKind::Between.native_value() => {
                return Ok(Self::FixedDate(FixedDatePredicateKind::Between));
            },
            _ => {},
        }
        match value {
            value if value == RelativeDatePredicateKind::Today.native_value() => {
                return Ok(Self::RelativeDate(RelativeDatePredicateKind::Today));
            },
            value if value == RelativeDatePredicateKind::Yesterday.native_value() => {
                return Ok(Self::RelativeDate(RelativeDatePredicateKind::Yesterday));
            },
            value if value == RelativeDatePredicateKind::Tomorrow.native_value() => {
                return Ok(Self::RelativeDate(RelativeDatePredicateKind::Tomorrow));
            },
            _ => {},
        }
        match value {
            value if value == CellPredicateKind::IsBlank.native_value() => {
                return Ok(Self::Cell(CellPredicateKind::IsBlank));
            },
            value if value == CellPredicateKind::IsNotBlank.native_value() => {
                return Ok(Self::Cell(CellPredicateKind::IsNotBlank));
            },
            _ => {},
        }
        match value {
            value if value == CheckboxPredicateKind::IsChecked.native_value() => {
                return Ok(Self::Checkbox(CheckboxPredicateKind::IsChecked));
            },
            value if value == CheckboxPredicateKind::IsNotChecked.native_value() => {
                return Ok(Self::Checkbox(CheckboxPredicateKind::IsNotChecked));
            },
            _ => {},
        }
        match value {
            value if value == BooleanPredicateKind::IsTrue.native_value() => {
                return Ok(Self::Boolean(BooleanPredicateKind::IsTrue));
            },
            value if value == BooleanPredicateKind::IsFalse.native_value() => {
                return Ok(Self::Boolean(BooleanPredicateKind::IsFalse));
            },
            _ => {},
        }
        match value {
            value if value == NumericSignPredicateKind::IsPositive.native_value() => {
                return Ok(Self::NumericSign(NumericSignPredicateKind::IsPositive));
            },
            value if value == NumericSignPredicateKind::IsNegative.native_value() => {
                return Ok(Self::NumericSign(NumericSignPredicateKind::IsNegative));
            },
            _ => {},
        }
        let text = match value {
            value if value == TextPredicateKind::StartsWith.native_value() => {
                Some(TextPredicateKind::StartsWith)
            },
            value if value == TextPredicateKind::EndsWith.native_value() => {
                Some(TextPredicateKind::EndsWith)
            },
            value if value == TextPredicateKind::DoesNotStartWith.native_value() => {
                Some(TextPredicateKind::DoesNotStartWith)
            },
            value if value == TextPredicateKind::DoesNotEndWith.native_value() => {
                Some(TextPredicateKind::DoesNotEndWith)
            },
            value if value == TextPredicateKind::Contains.native_value() => {
                Some(TextPredicateKind::Contains)
            },
            value if value == TextPredicateKind::DoesNotContain.native_value() => {
                Some(TextPredicateKind::DoesNotContain)
            },
            value if value == TextPredicateKind::EqualTo.native_value() => {
                Some(TextPredicateKind::EqualTo)
            },
            value if value == TextPredicateKind::NotEqualTo.native_value() => {
                Some(TextPredicateKind::NotEqualTo)
            },
            _ => None,
        };
        if let Some(kind) = text {
            return Ok(Self::Text(kind));
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
        for text in [
            TextPredicateKind::EqualTo,
            TextPredicateKind::NotEqualTo,
            TextPredicateKind::StartsWith,
            TextPredicateKind::DoesNotStartWith,
            TextPredicateKind::EndsWith,
            TextPredicateKind::DoesNotEndWith,
            TextPredicateKind::Contains,
            TextPredicateKind::DoesNotContain,
        ] {
            let kind = NativePredicateKind::Text(text);
            assert_eq!(
                NativePredicateKind::try_from(kind.native_value()).unwrap(),
                kind
            );
        }
    }

    #[test]
    fn native_cell_predicate_values_are_reversible() {
        for cell in [CellPredicateKind::IsBlank, CellPredicateKind::IsNotBlank] {
            let kind = NativePredicateKind::Cell(cell);
            assert_eq!(
                NativePredicateKind::try_from(kind.native_value()).unwrap(),
                kind
            );
        }
    }

    #[test]
    fn native_numeric_sign_predicate_values_are_reversible() {
        for sign in [
            NumericSignPredicateKind::IsPositive,
            NumericSignPredicateKind::IsNegative,
        ] {
            let kind = NativePredicateKind::NumericSign(sign);
            assert_eq!(
                NativePredicateKind::try_from(kind.native_value()).unwrap(),
                kind
            );
        }
    }

    #[test]
    fn native_boolean_predicate_values_are_reversible() {
        for boolean in [BooleanPredicateKind::IsTrue, BooleanPredicateKind::IsFalse] {
            let kind = NativePredicateKind::Boolean(boolean);
            assert_eq!(
                NativePredicateKind::try_from(kind.native_value()).unwrap(),
                kind
            );
        }
    }

    #[test]
    fn native_checkbox_predicate_values_are_reversible() {
        for checkbox in [
            CheckboxPredicateKind::IsChecked,
            CheckboxPredicateKind::IsNotChecked,
        ] {
            let kind = NativePredicateKind::Checkbox(checkbox);
            assert_eq!(
                NativePredicateKind::try_from(kind.native_value()).unwrap(),
                kind
            );
        }
    }

    #[test]
    fn native_relative_date_predicate_values_are_reversible() {
        for date in [
            RelativeDatePredicateKind::Today,
            RelativeDatePredicateKind::Yesterday,
            RelativeDatePredicateKind::Tomorrow,
        ] {
            let kind = NativePredicateKind::RelativeDate(date);
            assert_eq!(
                NativePredicateKind::try_from(kind.native_value()).unwrap(),
                kind
            );
        }
    }

    #[test]
    fn native_fixed_date_predicate_values_are_reversible() {
        for date in [
            FixedDatePredicateKind::Equal,
            FixedDatePredicateKind::Before,
            FixedDatePredicateKind::After,
            FixedDatePredicateKind::Between,
        ] {
            let kind = NativePredicateKind::FixedDate(date);
            assert_eq!(
                NativePredicateKind::try_from(kind.native_value()).unwrap(),
                kind
            );
        }
    }

    #[test]
    fn native_date_period_predicate_values_are_reversible() {
        for period in [
            DatePeriodPredicateKind::InNext,
            DatePeriodPredicateKind::InLast,
            DatePeriodPredicateKind::OffsetFromToday,
        ] {
            let kind = NativePredicateKind::DatePeriod(period);
            assert_eq!(
                NativePredicateKind::try_from(kind.native_value()).unwrap(),
                kind
            );
        }
    }
}
