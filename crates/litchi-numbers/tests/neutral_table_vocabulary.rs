//! Record 0368 (decision D4 of record 0363): the neutral table vocabulary
//! moved to `litchi_iwa_common::table`, and every value stays reachable at
//! its established `litchi_numbers` path as the same nominal type — a
//! re-export, never a second definition or alias.

use std::any::TypeId;

use litchi_iwa_common::table::{cell::data_format as common, topology as common_topology};
use litchi_numbers::cell::data_format as numbers;
use litchi_numbers::table::topology as numbers_topology;

fn same<A: 'static, B: 'static>() -> bool {
    TypeId::of::<A>() == TypeId::of::<B>()
}

macro_rules! assert_same {
    ($($numbers:path => $common:path),+ $(,)?) => {
        $(assert!(
            same::<$numbers, $common>(),
            concat!(stringify!($numbers), " is not ", stringify!($common))
        );)+
    };
}

#[test]
fn data_formats_are_the_common_types_at_their_numbers_paths() {
    assert_same!(
        numbers::DataFormat => common::DataFormat,
        numbers::Text => common::Text,
        numbers::Checkbox => common::Checkbox,
        numbers::StarRating => common::StarRating,
        numbers::CellControl => common::CellControl,
        numbers::control::DisplayFormat => common::control::DisplayFormat,
        numbers::control::Range => common::control::Range,
        numbers::control::Error => common::control::Error,
        numbers::control::NotCellControl => common::control::NotCellControl,
        numbers::Slider => common::Slider,
        numbers::Stepper => common::Stepper,
        numbers::currency::Currency => common::currency::Currency,
        numbers::currency::CurrencyCode => common::currency::CurrencyCode,
        numbers::currency::CurrencyStyle => common::currency::CurrencyStyle,
        numbers::custom::Custom => common::custom::Custom,
        numbers::custom::Condition => common::custom::Condition,
        numbers::custom::Error => common::custom::Error,
        numbers::date_time::DateTime => common::date_time::DateTime,
        numbers::date_time::Error => common::date_time::Error,
        numbers::duration::Duration => common::duration::Duration,
        numbers::duration::Units => common::duration::Units,
        numbers::fraction::Fraction => common::fraction::Fraction,
        numbers::fraction::FractionAccuracy => common::fraction::FractionAccuracy,
        numbers::number::Number => common::number::Number,
        numbers::number::Percentage => common::number::Percentage,
        numbers::number::DecimalPlaces => common::number::DecimalPlaces,
        numbers::number::FixedDecimalPlaces => common::number::FixedDecimalPlaces,
        numbers::number::NegativeStyle => common::number::NegativeStyle,
        numbers::number::ThousandsSeparator => common::number::ThousandsSeparator,
        numbers::number::Error => common::number::Error,
        numbers::numeral_system::NumeralSystem => common::numeral_system::NumeralSystem,
        numbers::numeral_system::Base => common::numeral_system::Base,
        numbers::percentage::Percentage => common::percentage::Percentage,
        numbers::pop_up_menu::PopUpMenu => common::pop_up_menu::PopUpMenu,
        numbers::pop_up_menu::InitialSelection => common::pop_up_menu::InitialSelection,
        numbers::pop_up_menu::Item => common::pop_up_menu::Item,
        numbers::scientific::Scientific => common::scientific::Scientific,
        numbers::text::Text => common::text::Text,
        litchi_numbers::PopUpMenu => common::PopUpMenu,
    );
    assert_eq!(
        numbers::number::MAX_DECIMAL_PLACES,
        common::number::MAX_DECIMAL_PLACES
    );
}

#[test]
fn topology_coordinates_are_the_common_types() {
    assert_same!(
        numbers_topology::RowInsertion => common_topology::RowInsertion,
        numbers_topology::ColumnInsertion => common_topology::ColumnInsertion,
        numbers_topology::RowDeletion => common_topology::RowDeletion,
        numbers_topology::ColumnDeletion => common_topology::ColumnDeletion,
        litchi_numbers::RowDeletion => common_topology::RowDeletion,
        litchi_numbers::ColumnDeletion => common_topology::ColumnDeletion,
    );
}

#[test]
fn a_common_value_drives_the_numbers_transaction_vocabulary() {
    // A value built from the common crate is accepted where the Numbers API
    // names its own path, because both paths denote one type.
    let format: numbers::DataFormat = common::number::Number::default().into();
    assert_eq!(
        format,
        numbers::DataFormat::Number(numbers::number::Number::default())
    );
    let insertion: numbers_topology::RowInsertion = common_topology::RowInsertion::body(2);
    assert_eq!(insertion, numbers_topology::RowInsertion::Body { index: 2 });
}
