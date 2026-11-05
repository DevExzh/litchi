//! Record 0372 (wave W1-C; the gap record 0368 left open): Keynote's
//! slide-table Number display format and cell coordinate are the shared
//! `litchi_iwa_common::table` values, re-exported at their established
//! `litchi_keynote` paths as the same nominal types — never a second
//! definition.

use std::any::TypeId;

use litchi_iwa_common::table::cell::data_format::number as common;
use litchi_iwa_common::table::coordinate;
use litchi_keynote::slide::table::number_format as keynote;

fn same<A: 'static, B: 'static>() -> bool {
    TypeId::of::<A>() == TypeId::of::<B>()
}

macro_rules! assert_same {
    ($($keynote:path => $common:path),+ $(,)?) => {
        $(assert!(
            same::<$keynote, $common>(),
            concat!(stringify!($keynote), " is not ", stringify!($common))
        );)+
    };
}

#[test]
fn number_format_values_are_the_common_types() {
    assert_same!(
        keynote::Number => common::Number,
        keynote::DecimalPlaces => common::DecimalPlaces,
        keynote::FixedDecimalPlaces => common::FixedDecimalPlaces,
        keynote::NegativeStyle => common::NegativeStyle,
        keynote::ThousandsSeparator => common::ThousandsSeparator,
        keynote::NumberError => common::Error,
        litchi_keynote::Number => common::Number,
        litchi_keynote::NumberError => common::Error,
    );
    assert_eq!(keynote::MAX_DECIMAL_PLACES, common::MAX_DECIMAL_PLACES);
}

#[test]
fn coordinates_are_the_common_types() {
    assert_same!(
        keynote::At => coordinate::At,
        keynote::AtError => coordinate::Error,
        litchi_keynote::At => coordinate::At,
        litchi_keynote::AtError => coordinate::Error,
        litchi_keynote::slide::table::cells::At => coordinate::At,
    );
}
