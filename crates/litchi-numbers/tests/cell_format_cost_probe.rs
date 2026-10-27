//! Callgrind probe of the Numbers cell-format routes (record 0379).
//!
//! Each route opens the deterministic data-format fixture of its family,
//! sets a new value on the first cell, serializes, clears the format on that
//! output and serializes again. (The control and Pop-Up Menu owners refuse
//! the data-format fixture's graph, so they are not part of this probe.) `cell_format_cost_probe`
//! is ignored by default; it exists so the routes can be measured with the
//! same input before and after the shared `litchi-iwa-table` engine move:
//!
//! ```sh
//! CARGO_PROFILE_RELEASE_DEBUG=1 cargo test --release -p litchi-numbers \
//!   --test cell_format_cost_probe --no-run
//! taskset -c 3 valgrind --tool=callgrind '--toggle-collect=*cell_format_cost_probe::run_routes*' \
//!   target/release/deps/cell_format_cost_probe-<hash> --ignored --exact \
//!   cell_format_cost_probe --test-threads=1
//! ```

use litchi_numbers::cell::data_format::currency::{Currency, CurrencyCode, CurrencyStyle};
use litchi_numbers::cell::data_format::duration::{Duration, Style, UnitRange, Units};
use litchi_numbers::cell::data_format::number::{
    DecimalPlaces, FixedDecimalPlaces, NegativeStyle, Number, Percentage, Scientific,
    ThousandsSeparator,
};
use litchi_numbers::cell::data_format::{DateTime, Fraction, FractionAccuracy, Text};
use litchi_numbers::{At, Package};

#[path = "support/table_cell_data_format_fixture.rs"]
mod fixture;

use fixture::{FormatFamily, FormatSharing};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn at() -> At {
    At::new(fixture::FIRST_CELL.0 as u32, fixture::FIRST_CELL.1 as u32)
}

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut out = Vec::new();
    package.write_to(&mut out)?;
    Ok(out)
}

macro_rules! set_then_clear {
    ($package:expr, $edit:ident, $value:expr) => {{
        let set = $package.$edit(0usize, 0usize, at())?.set($value).commit()?;
        let first = bytes(set.package())?;
        let cleared = set
            .package()
            .$edit(0usize, 0usize, at())?
            .clear()
            .commit()?;
        let second = bytes(cleared.package())?;
        (cleared.into_package(), first.len() + second.len())
    }};
}

fn number_route(source: &[u8]) -> TestResult<usize> {
    let package = Package::from_bytes(source)?;
    let value = Number::new(
        DecimalPlaces::fixed(4)?,
        NegativeStyle::Red,
        ThousandsSeparator::Hidden,
    );
    let (_, total) = set_then_clear!(package, edit_table_cell_number_format, value);
    Ok(total)
}

fn currency_route(source: &[u8]) -> TestResult<usize> {
    let value = Currency::new(
        CurrencyCode::new("EUR")?,
        DecimalPlaces::fixed(3)?,
        NegativeStyle::Parentheses,
        ThousandsSeparator::Hidden,
        CurrencyStyle::Accounting,
    );
    let (_, total) = set_then_clear!(
        Package::from_bytes(source)?,
        edit_table_cell_currency_format,
        value
    );
    Ok(total)
}

fn percentage_route(source: &[u8]) -> TestResult<usize> {
    let value = Percentage::new(
        DecimalPlaces::fixed(1)?,
        NegativeStyle::RedParentheses,
        ThousandsSeparator::Hidden,
    );
    let (_, total) = set_then_clear!(
        Package::from_bytes(source)?,
        edit_table_cell_percentage_format,
        value
    );
    Ok(total)
}

fn scientific_route(source: &[u8]) -> TestResult<usize> {
    let value = Scientific::new(FixedDecimalPlaces::new(5)?);
    let (_, total) = set_then_clear!(
        Package::from_bytes(source)?,
        edit_table_cell_scientific_format,
        value
    );
    Ok(total)
}

fn fraction_route(source: &[u8]) -> TestResult<usize> {
    let value = Fraction::new(FractionAccuracy::Sixteenths);
    let (_, total) = set_then_clear!(
        Package::from_bytes(source)?,
        edit_table_cell_fraction_format,
        value
    );
    Ok(total)
}

fn duration_route(source: &[u8]) -> TestResult<usize> {
    let value = Duration::new(Style::FullNames, Units::Custom(UnitRange::all()));
    let (_, total) = set_then_clear!(
        Package::from_bytes(source)?,
        edit_table_cell_duration_format,
        value
    );
    Ok(total)
}

fn date_time_route(source: &[u8]) -> TestResult<usize> {
    let (_, total) = set_then_clear!(
        Package::from_bytes(source)?,
        edit_table_cell_date_time_format,
        DateTime::new("yyyy-MM-dd HH:mm")?
    );
    Ok(total)
}

fn text_route(source: &[u8]) -> TestResult<usize> {
    let package = Package::from_bytes(source)?;
    let cleared = package
        .edit_table_cell_text_format(0usize, 0usize, at())?
        .clear()
        .commit()?;
    let first = bytes(cleared.package())?;
    let set = cleared
        .package()
        .edit_table_cell_text_format(0usize, 0usize, at())?
        .set(Text)
        .commit()?;
    Ok(first.len() + bytes(set.package())?.len())
}

type Route = fn(&[u8]) -> TestResult<usize>;

const ROUTES: [(&str, FormatFamily, Route); 8] = [
    ("number", FormatFamily::Number, number_route),
    ("currency", FormatFamily::Currency, currency_route),
    ("percentage", FormatFamily::Percentage, percentage_route),
    ("scientific", FormatFamily::Scientific, scientific_route),
    ("fraction", FormatFamily::Fraction, fraction_route),
    ("duration", FormatFamily::Duration, duration_route),
    ("date_time", FormatFamily::DateTime, date_time_route),
    ("text", FormatFamily::Text, text_route),
];

fn inputs() -> TestResult<Vec<Vec<u8>>> {
    ROUTES
        .iter()
        .map(|(_, family, _)| fixture::synthetic_package_for(*family, FormatSharing::Shared))
        .collect()
}

#[inline(never)]
fn run_routes(inputs: &[Vec<u8>]) -> TestResult<usize> {
    let mut total = 0usize;
    for ((_, _, route), input) in ROUTES.iter().zip(inputs) {
        total += route(input)?;
    }
    Ok(total)
}

/// Every route must succeed; otherwise the probe measures a refusal.
#[test]
fn cell_format_routes_succeed_on_probe_fixtures() -> TestResult {
    for ((name, _, route), input) in ROUTES.iter().zip(inputs()?) {
        route(&input).map_err(|error| format!("{name}: {error}"))?;
    }
    Ok(())
}

#[test]
#[ignore = "callgrind probe; run explicitly"]
fn cell_format_cost_probe() -> TestResult {
    let inputs = inputs()?;
    assert!(run_routes(&inputs)? > 0);
    Ok(())
}
