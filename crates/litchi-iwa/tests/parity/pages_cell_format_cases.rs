//! Pages body-table cell formats of wave W1-G (record 0379): the host's
//! `set_table_cell_<family>_format` / `reset_table_cell_<family>_format` /
//! `table_cell_<family>_format` against the leaf's single
//! `edit_body_table_cell_format` family over the shared
//! `litchi_iwa_table::cell_format` engine.

use litchi_iwa::pages;
use litchi_iwa_common::table::cell::data_format::currency::{
    Currency, CurrencyCode, CurrencyStyle,
};
use litchi_iwa_common::table::cell::data_format::duration::{Duration, Style, UnitRange, Units};
use litchi_iwa_common::table::cell::data_format::number::{
    DecimalPlaces, FixedDecimalPlaces, NegativeStyle, Number, Percentage, Scientific,
    ThousandsSeparator,
};
use litchi_iwa_common::table::cell::data_format::{
    DataFormat, DateTime, Fraction, FractionAccuracy, Text,
};
use litchi_iwa_common::table::coordinate::At;
use litchi_pages::Package;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

/// Native body table whose stored cells both routes read (the catalog
/// fixture's cells are outside the shared engine's tile shape, and the host
/// refuses most cells of the other fixtures as inconsistent metadata).
const CATALOG: &str = "iwork/pages/body-table-cells-native.pages";
const TABLE: usize = 0;

fn host_model(editor: &pages::Draft) -> TestResult<u64> {
    Ok(editor
        .tables()?
        .get(TABLE)
        .ok_or("body table missing")?
        .model_object_id)
}

/// One sample value per display family the shared engine writes.
pub(crate) fn samples() -> TestResult<Vec<(&'static str, DataFormat)>> {
    Ok(vec![
        (
            "number",
            Number::new(
                DecimalPlaces::fixed(2)?,
                NegativeStyle::Parentheses,
                ThousandsSeparator::Shown,
            )
            .into(),
        ),
        (
            "currency",
            Currency::new(
                CurrencyCode::new("EUR")?,
                DecimalPlaces::fixed(2)?,
                NegativeStyle::MinusSign,
                ThousandsSeparator::Shown,
                CurrencyStyle::Accounting,
            )
            .into(),
        ),
        (
            "percentage",
            Percentage::new(
                DecimalPlaces::fixed(1)?,
                NegativeStyle::Red,
                ThousandsSeparator::Hidden,
            )
            .into(),
        ),
        (
            "scientific",
            Scientific::new(FixedDecimalPlaces::new(3)?).into(),
        ),
        ("fraction", Fraction::new(FractionAccuracy::Quarters).into()),
        ("date_time", DateTime::new("yyyy-MM-dd")?.into()),
        (
            "duration",
            Duration::new(Style::Abbreviated, Units::Custom(UnitRange::all())).into(),
        ),
        ("text", Text.into()),
    ])
}

/// The first stored cell both routes read as Automatic.
fn automatic_cell(input: &[u8]) -> TestResult<At> {
    let package = Package::from_bytes(input)?;
    for row in 0..16 {
        for column in 0..8 {
            let at = At::new(row, column);
            if package.body_table_cell_format(TABLE, at).ok() == Some(DataFormat::Automatic)
                && host_read(input, at).ok() == Some(DataFormat::Automatic)
            {
                return Ok(at);
            }
        }
    }
    Err("no automatic stored cell".into())
}

fn host_read(bytes: &[u8], at: At) -> TestResult<DataFormat> {
    let editor = pages::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    Ok(editor.table_cell_data_format(model, at.row() as usize, at.column() as usize)?)
}

fn leaf_read(bytes: &[u8], at: At) -> TestResult<DataFormat> {
    Ok(Package::from_bytes(bytes)?.body_table_cell_format(TABLE, at)?)
}

fn host_set(source: &[u8], at: At, format: &DataFormat) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    let model = host_model(&editor)?;
    editor.set_table_cell_data_format(
        model,
        at.row() as usize,
        at.column() as usize,
        format.clone(),
    )?;
    Ok(editor.to_bytes()?)
}

/// The host's family-specific setter for `format`.
fn host_family_set(source: &[u8], at: At, format: &DataFormat) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    let model = host_model(&editor)?;
    let (row, column) = (at.row() as usize, at.column() as usize);
    match format {
        DataFormat::Number(value) => {
            editor.set_table_cell_number_format(model, row, column, *value)?
        },
        DataFormat::Currency(value) => {
            editor.set_table_cell_currency_format(model, row, column, *value)?
        },
        DataFormat::Percentage(value) => {
            editor.set_table_cell_percentage_format(model, row, column, *value)?
        },
        DataFormat::Scientific(value) => {
            editor.set_table_cell_scientific_format(model, row, column, *value)?
        },
        DataFormat::Fraction(value) => {
            editor.set_table_cell_fraction_format(model, row, column, *value)?
        },
        _ => return Err("no family-specific host setter".into()),
    }
    Ok(editor.to_bytes()?)
}

/// The host's family-specific read of the family of `format`.
fn host_family_read(bytes: &[u8], at: At, format: &DataFormat) -> TestResult<String> {
    let editor = pages::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    let (row, column) = (at.row() as usize, at.column() as usize);
    Ok(match format {
        DataFormat::Number(_) => {
            format!("{:?}", editor.table_cell_number_format(model, row, column)?)
        },
        DataFormat::Currency(_) => format!(
            "{:?}",
            editor.table_cell_currency_format(model, row, column)?
        ),
        DataFormat::Percentage(_) => format!(
            "{:?}",
            editor.table_cell_percentage_format(model, row, column)?
        ),
        DataFormat::Scientific(_) => format!(
            "{:?}",
            editor.table_cell_scientific_format(model, row, column)?
        ),
        DataFormat::Fraction(_) => format!(
            "{:?}",
            editor.table_cell_fraction_format(model, row, column)?
        ),
        _ => return Err("no family-specific host read".into()),
    })
}

fn leaf_set(source: &[u8], at: At, format: &DataFormat) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let commit = package
        .edit_body_table_cell_format(TABLE, at)?
        .set(format.clone())
        .commit()?;
    let restored = commit
        .package()
        .apply_body_table_cell_format(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn set_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    at: At,
    format: &DataFormat,
) -> TestResult {
    Case::new(name, Format::Pages, tier, evidence, input)
        .host(|source| host_set(source, at, format))
        .leaf(|source| leaf_set(source, at, format))
        .read("host table_cell_data_format", move |bytes| {
            host_read(bytes, at)
        })
        .read("leaf body_table_cell_format", move |bytes| {
            leaf_read(bytes, at)
        })
        .run()?;
    Ok(())
}

/// A second value of each family the engine admits on the host's output.
fn replacements() -> TestResult<Vec<DataFormat>> {
    Ok(vec![
        Number::new(
            DecimalPlaces::fixed(4)?,
            NegativeStyle::Red,
            ThousandsSeparator::Hidden,
        )
        .into(),
        Currency::new(
            CurrencyCode::new("USD")?,
            DecimalPlaces::fixed(0)?,
            NegativeStyle::Parentheses,
            ThousandsSeparator::Hidden,
            CurrencyStyle::Standard,
        )
        .into(),
        Percentage::new(
            DecimalPlaces::fixed(3)?,
            NegativeStyle::MinusSign,
            ThousandsSeparator::Shown,
        )
        .into(),
        Scientific::new(FixedDecimalPlaces::new(1)?).into(),
        Fraction::new(FractionAccuracy::Eighths).into(),
    ])
}

/// The families whose host-written cells the shared engine admits: Number,
/// Currency, Percentage, Scientific and Fraction.
const ADMITTED: usize = 5;

fn host_formatted(input: &Fixture, at: At, format: &DataFormat) -> TestResult<Fixture> {
    Ok(Fixture {
        path: input.path,
        bytes: host_set(&input.bytes, at, format)?,
    })
}

/// Automatic to each decimal family on a stored cell through the host's
/// generic setter. The native cell names an implicit Number entry: Number
/// rewrites it in place, Percentage, Scientific and Fraction intern a fresh
/// entry (record 0412 closes that part of the record 0379 §5 gap).
#[test]
fn pages_body_table_cell_data_format_set_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    // Currency changes the cell's format kind as well; the engine does not
    // convert it yet (pinned by the divergence case).
    for (_, format) in samples()?
        .iter()
        .take(ADMITTED)
        .filter(|(family, _)| *family != "currency")
    {
        set_case(
            "pages_body_table_cell_data_format_set_parity",
            Tier::B,
            Evidence::Queued,
            &input,
            at,
            format,
        )?;
    }
    Ok(())
}

/// A replacement inside each admitted family on the host's formatted
/// output, through the host's family-specific setters.
#[test]
fn pages_body_table_cell_format_set_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    let samples = samples()?;
    for ((_, format), replacement) in samples.iter().take(ADMITTED).zip(replacements()?) {
        let formatted = host_formatted(&input, at, format)?;
        Case::new(
            "pages_body_table_cell_format_set_parity",
            Format::Pages,
            Tier::B,
            Evidence::Queued,
            &formatted,
        )
        .host(|source| host_family_set(source, at, &replacement))
        .leaf(|source| leaf_set(source, at, &replacement))
        .read("host family read", |bytes| {
            host_family_read(bytes, at, &replacement)
        })
        .read("leaf cell format", move |bytes| leaf_read(bytes, at))
        .run()?;
    }
    Ok(())
}

/// Each admitted family reset to Automatic on the host's formatted output,
/// through the host's family-specific reset.
#[test]
fn pages_body_table_cell_format_reset_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    for (family, format) in samples()?.into_iter().take(ADMITTED) {
        let formatted = host_formatted(&input, at, &format)?;
        Case::new(
            "pages_body_table_cell_format_reset_parity",
            Format::Pages,
            Tier::B,
            Evidence::Queued,
            &formatted,
        )
        .host(|source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            let (row, column) = (at.row() as usize, at.column() as usize);
            let changed = match family {
                "number" => editor.reset_table_cell_number_format(model, row, column)?,
                "currency" => editor.reset_table_cell_currency_format(model, row, column)?,
                "percentage" => editor.reset_table_cell_percentage_format(model, row, column)?,
                "scientific" => editor.reset_table_cell_scientific_format(model, row, column)?,
                "fraction" => editor.reset_table_cell_fraction_format(model, row, column)?,
                _ => return Err("unmapped family".into()),
            };
            if !changed {
                return Err("host reset reported no change".into());
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| leaf_set(source, at, &DataFormat::Automatic))
        .read("host table_cell_data_format", move |bytes| {
            host_read(bytes, at)
        })
        .read("leaf body_table_cell_format", move |bytes| {
            leaf_read(bytes, at)
        })
        .run()?;
    }
    Ok(())
}

/// The host's reads and the leaf's single read agree on the host's output
/// of every admitted family.
#[test]
fn pages_body_table_cell_format_read_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    for (_, format) in samples()?.into_iter().take(ADMITTED) {
        let formatted = host_formatted(&input, at, &format)?;
        Case::new(
            "pages_body_table_cell_format_read_parity",
            Format::Pages,
            Tier::A,
            Evidence::Queued,
            &formatted,
        )
        .read_only()
        .read("host family read", |bytes| {
            host_family_read(bytes, at, &format)
        })
        .host(|source| Ok(pages::Draft::from_bytes(source)?.to_bytes()?))
        .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
        .read("host table_cell_data_format", move |bytes| {
            host_read(bytes, at)
        })
        .read("leaf body_table_cell_format", move |bytes| {
            leaf_read(bytes, at)
        })
        .run()?;
        assert_eq!(host_read(&formatted.bytes, at)?, format);
        assert_eq!(leaf_read(&formatted.bytes, at)?, format);
    }
    Ok(())
}

/// Where the routes differ (record 0379 §5; record 0412 closed Automatic to
/// Percentage, Scientific and Fraction): the engine refuses Automatic to
/// Currency, Date and Time and Duration on a number cell, and a cell the host formatted as Date and
/// Time or Duration (the host leaves the stored number under a date or
/// duration kind, which the engine reads as an invalid cell). The host
/// refuses Text on a number cell; the leaf refuses it too.
#[test]
fn pages_body_table_cell_format_divergence_is_reported() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    let samples = samples()?;
    for (family, format) in samples
        .iter()
        .filter(|(family, _)| *family == "currency")
        .chain(&samples[ADMITTED..])
    {
        assert!(host_set(&input.bytes, at, format).is_ok() || *family == "text");
        assert!(leaf_set(&input.bytes, at, format).is_err(), "{family}");
    }
    for (family, format) in &samples[ADMITTED..] {
        let Ok(host) = host_set(&input.bytes, at, format) else {
            continue;
        };
        assert!(leaf_read(&host, at).is_err(), "{family}");
    }
    Ok(())
}

/// Callgrind probe (record 0379): a Currency replacement on the host-written
/// EUR Currency cell, host family setter against the leaf family, ABBA.
/// `--toggle-collect=*probe_host*` / `*probe_leaf*` on
/// `pages_cell_format_cases::pages_cell_format_cost_abba`.
#[test]
#[ignore = "callgrind probe; run explicitly"]
fn pages_cell_format_cost_abba() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    let samples = samples()?;
    let format = samples
        .get(1)
        .map(|(_, format)| format.clone())
        .ok_or("sample")?;
    let formatted = host_formatted(&input, at, &format)?;
    let replacement = replacements()?.get(1).cloned().ok_or("replacement")?;
    crate::pages_cost_probes::abba(
        &formatted,
        |source| host_family_set(source, at, &replacement),
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_body_table_cell_format(TABLE, at)?
                .set(replacement.clone())
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

/// Numeral systems (record 0370k): the leaf reads every numeral system the
/// host writes.
#[test]
fn pages_body_table_cell_numeral_system_format_read_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    for numeral in crate::numbers_cell_data_format_cases::numeral_samples()? {
        let format = DataFormat::from(numeral);
        let formatted = host_formatted(&input, at, &format)?;
        Case::new(
            "pages_body_table_cell_numeral_system_format_read_parity",
            Format::Pages,
            Tier::A,
            Evidence::Queued,
            &formatted,
        )
        .read_only()
        .host(|source| Ok(pages::Draft::from_bytes(source)?.to_bytes()?))
        .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
        .read("host table_cell_data_format", move |bytes| {
            host_read(bytes, at)
        })
        .read("leaf cell format", move |bytes| leaf_read(bytes, at))
        .run()?;
        assert_eq!(host_read(&formatted.bytes, at)?, format);
        assert_eq!(leaf_read(&formatted.bytes, at)?, format);
    }
    Ok(())
}

/// One sample of each scalar control (record 0412).
pub(crate) fn control_samples() -> TestResult<Vec<(&'static str, DataFormat)>> {
    use litchi_iwa_common::table::cell::data_format::control::{
        DisplayFormat, Range, Slider, Stepper,
    };
    use litchi_iwa_common::table::cell::data_format::{Checkbox, StarRating};
    let number = Number::new(
        DecimalPlaces::fixed(2)?,
        NegativeStyle::MinusSign,
        ThousandsSeparator::Hidden,
    );
    Ok(vec![
        ("checkbox", Checkbox.into()),
        ("star_rating", StarRating.into()),
        (
            "slider",
            Slider::new(Range::new(0.0, 10.0, 1.0)?, DisplayFormat::Number(number)).into(),
        ),
        (
            "stepper",
            Stepper::new(Range::new(0.0, 10.0, 2.0)?, DisplayFormat::Number(number)).into(),
        ),
    ])
}

/// The host's control family setter for `format`.
fn host_control_set(source: &[u8], at: At, format: &DataFormat) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    let model = host_model(&editor)?;
    let (row, column) = (at.row() as usize, at.column() as usize);
    match format {
        DataFormat::Checkbox(value) => {
            editor.set_table_cell_checkbox_format(model, row, column, *value)?
        },
        DataFormat::StarRating(value) => {
            editor.set_table_cell_star_rating_format(model, row, column, *value)?
        },
        DataFormat::Slider(value) => {
            editor.set_table_cell_slider_format(model, row, column, value.clone())?
        },
        DataFormat::Stepper(value) => {
            editor.set_table_cell_stepper_format(model, row, column, value.clone())?
        },
        _ => return Err("not a scalar control".into()),
    }
    Ok(editor.to_bytes()?)
}

/// The host's control family reset for `format`'s family.
fn host_control_reset(source: &[u8], at: At, format: &DataFormat) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    let model = host_model(&editor)?;
    let (row, column) = (at.row() as usize, at.column() as usize);
    let changed = match format {
        DataFormat::Checkbox(_) => editor.reset_table_cell_checkbox_format(model, row, column)?,
        DataFormat::StarRating(_) => {
            editor.reset_table_cell_star_rating_format(model, row, column)?
        },
        DataFormat::Slider(_) => editor.reset_table_cell_slider_format(model, row, column)?,
        DataFormat::Stepper(_) => editor.reset_table_cell_stepper_format(model, row, column)?,
        _ => return Err("not a scalar control".into()),
    };
    if !changed {
        return Err("host reset reported no change".into());
    }
    Ok(editor.to_bytes()?)
}

/// The host's control family read of `format`'s family.
fn host_control_read(bytes: &[u8], at: At, format: &DataFormat) -> TestResult<String> {
    let editor = pages::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    let (row, column) = (at.row() as usize, at.column() as usize);
    Ok(match format {
        DataFormat::Checkbox(_) => {
            format!(
                "{:?}",
                editor
                    .table_cell_checkbox_format(model, row, column)?
                    .map(DataFormat::from)
            )
        },
        DataFormat::StarRating(_) => format!(
            "{:?}",
            editor
                .table_cell_star_rating_format(model, row, column)?
                .map(DataFormat::from)
        ),
        DataFormat::Slider(_) => {
            format!(
                "{:?}",
                editor
                    .table_cell_slider_format(model, row, column)?
                    .map(DataFormat::from)
            )
        },
        DataFormat::Stepper(_) => {
            format!(
                "{:?}",
                editor
                    .table_cell_stepper_format(model, row, column)?
                    .map(DataFormat::from)
            )
        },
        _ => return Err("not a scalar control".into()),
    })
}

/// Automatic to each scalar control through the host's family setters
/// (record 0412).
#[test]
fn pages_body_table_cell_control_format_set_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    for (_, format) in control_samples()? {
        Case::new(
            "pages_body_table_cell_control_format_set_parity",
            Format::Pages,
            Tier::B,
            Evidence::Queued,
            &input,
        )
        .host(|source| host_control_set(source, at, &format))
        .leaf(|source| leaf_set(source, at, &format))
        .read("host family read", |bytes| {
            host_control_read(bytes, at, &format)
        })
        .read("leaf cell format", move |bytes| {
            leaf_read(bytes, at).map(|value| format!("{:?}", Some(value)))
        })
        .run()?;
    }
    Ok(())
}

/// Each scalar control reset to Automatic on the host's output, through the
/// host's family resets (record 0412).
#[test]
fn pages_body_table_cell_control_format_reset_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    for (_, format) in control_samples()? {
        let formatted = Fixture {
            path: input.path,
            bytes: host_control_set(&input.bytes, at, &format)?,
        };
        Case::new(
            "pages_body_table_cell_control_format_reset_parity",
            Format::Pages,
            Tier::B,
            Evidence::Queued,
            &formatted,
        )
        .host(|source| host_control_reset(source, at, &format))
        .leaf(|source| leaf_set(source, at, &DataFormat::Automatic))
        .read("host table_cell_data_format", move |bytes| {
            host_read(bytes, at)
        })
        .read("leaf cell format", move |bytes| leaf_read(bytes, at))
        .run()?;
    }
    Ok(())
}

/// The host's control family reads and the leaf's single read agree on the
/// host's output of every scalar control (record 0412).
#[test]
fn pages_body_table_cell_control_format_read_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    for (_, format) in control_samples()? {
        let formatted = Fixture {
            path: input.path,
            bytes: host_control_set(&input.bytes, at, &format)?,
        };
        Case::new(
            "pages_body_table_cell_control_format_read_parity",
            Format::Pages,
            Tier::A,
            Evidence::Queued,
            &formatted,
        )
        .read_only()
        .host(|source| Ok(pages::Draft::from_bytes(source)?.to_bytes()?))
        .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
        .read("host family read", |bytes| {
            host_control_read(bytes, at, &format)
        })
        .read("leaf cell format", move |bytes| {
            leaf_read(bytes, at).map(|value| format!("{:?}", Some(value)))
        })
        .run()?;
        assert_eq!(leaf_read(&formatted.bytes, at)?, format);
    }
    Ok(())
}

/// Automatic to each numeral system through the host's family setter: the
/// native cell names an implicit Number entry, which both routes leave for
/// a freshly interned numeral entry (record 0412, the 0370k gap).
#[test]
fn pages_body_table_cell_numeral_system_format_set_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    for numeral in crate::numbers_cell_data_format_cases::numeral_samples()? {
        let format = DataFormat::from(numeral);
        Case::new(
            "pages_body_table_cell_numeral_system_format_set_parity",
            Format::Pages,
            Tier::B,
            Evidence::Queued,
            &input,
        )
        .host(|source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            let (row, column) = (at.row() as usize, at.column() as usize);
            editor.set_table_cell_numeral_system_format(model, row, column, numeral)?;
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| leaf_set(source, at, &format))
        .read("host table_cell_data_format", move |bytes| {
            host_read(bytes, at)
        })
        .read("leaf cell format", move |bytes| leaf_read(bytes, at))
        .run()?;
    }
    Ok(())
}

/// Each numeral system reset to Automatic on the host's output, through the
/// host's family reset (record 0412).
#[test]
fn pages_body_table_cell_numeral_system_format_reset_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    for numeral in crate::numbers_cell_data_format_cases::numeral_samples()? {
        let formatted = host_formatted(&input, at, &DataFormat::from(numeral))?;
        Case::new(
            "pages_body_table_cell_numeral_system_format_reset_parity",
            Format::Pages,
            Tier::B,
            Evidence::Queued,
            &formatted,
        )
        .host(|source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            let (row, column) = (at.row() as usize, at.column() as usize);
            if !editor.reset_table_cell_numeral_system_format(model, row, column)? {
                return Err("host reset reported no change".into());
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| leaf_set(source, at, &DataFormat::Automatic))
        .read("host table_cell_data_format", move |bytes| {
            host_read(bytes, at)
        })
        .read("leaf cell format", move |bytes| leaf_read(bytes, at))
        .run()?;
    }
    Ok(())
}

/// The automatic cell rewritten by the leaf's cell writer to hold text,
/// which the host formats as Text (it refuses Text on a number cell,
/// record 0412).
fn text_input(input: &Fixture) -> TestResult<(Fixture, At)> {
    let at = automatic_cell(&input.bytes)?;
    let package = Package::from_bytes(&input.bytes)?;
    let mut edit = package.edit_body_table_values(TABLE)?;
    edit.set(
        at.row(),
        at.column(),
        litchi_pages::table::values::Value::Text("abc".into()),
    )?;
    let written = edit.commit()?.into_package();
    Ok((
        Fixture {
            path: input.path,
            bytes: leaf_bytes!(written),
        },
        at,
    ))
}

/// The host's Text read and the leaf's read agree on the host's output
/// (record 0412).
#[test]
fn pages_body_table_cell_text_format_read_parity() -> TestResult {
    let (input, at) = text_input(&fixture(CATALOG)?)?;
    let formatted = host_formatted(&input, at, &DataFormat::Text(Text))?;
    Case::new(
        "pages_body_table_cell_text_format_read_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        &formatted,
    )
    .read_only()
    .host(|source| Ok(pages::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
    .read("host family read", |bytes| {
        let editor = pages::Draft::from_bytes(bytes)?;
        let model = host_model(&editor)?;
        let (row, column) = (at.row() as usize, at.column() as usize);
        Ok(format!(
            "{:?}",
            editor
                .table_cell_text_format(model, row, column)?
                .map(DataFormat::from)
        ))
    })
    .read("leaf cell format", move |bytes| {
        leaf_read(bytes, at).map(|value| format!("{:?}", Some(value)))
    })
    .run()?;
    assert_eq!(leaf_read(&formatted.bytes, at)?, DataFormat::Text(Text));
    Ok(())
}

/// Callgrind probe (record 0412): Automatic to a control format on the first
/// automatic stored cell, host family setter against the leaf, ABBA.
#[test]
#[ignore = "callgrind probe; run explicitly"]
fn pages_body_table_cell_control_format_cost_abba() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    let format = control_samples()?[0].1.clone();
    crate::pages_cost_probes::abba(
        &input,
        |source| host_control_set(source, at, &format),
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_body_table_cell_format(TABLE, at)?
                .set(format.clone())
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

/// Callgrind probe (record 0412): Automatic to a numeral format on the first
/// automatic stored cell, host family setter against the leaf, ABBA.
#[test]
#[ignore = "callgrind probe; run explicitly"]
fn pages_body_table_cell_numeral_format_cost_abba() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    let format = DataFormat::from(crate::numbers_cell_data_format_cases::numeral_samples()?[0]);
    crate::pages_cost_probes::abba(
        &input,
        |source| host_set(source, at, &format),
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_body_table_cell_format(TABLE, at)?
                .set(format.clone())
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}
