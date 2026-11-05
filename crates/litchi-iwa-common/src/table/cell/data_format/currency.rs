//! Currency display values.
//!
//! Currency is a distinct semantic family even though Numbers stores its
//! decimal settings beside the plain Number and Percentage formats.  The
//! nominal types are re-exported from the compatibility decimal module while
//! this module owns the focused Currency transaction entry point.

pub use super::number::{
    Currency, CurrencyCode, CurrencyStyle, DecimalPlaces, NegativeStyle, ThousandsSeparator,
};

#[cfg(test)]
mod tests {
    use super::{Currency, CurrencyCode, CurrencyStyle, DecimalPlaces};

    #[test]
    fn currency_namespace_keeps_a_nominal_format_family() {
        let currency = Currency::new(
            CurrencyCode::EUR,
            DecimalPlaces::Automatic,
            super::NegativeStyle::MinusSign,
            super::ThousandsSeparator::Hidden,
            CurrencyStyle::Accounting,
        );

        assert_eq!(currency.code(), CurrencyCode::EUR);
        assert_eq!(currency.style(), CurrencyStyle::Accounting);
    }
}
