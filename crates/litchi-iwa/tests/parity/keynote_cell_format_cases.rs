//! Keynote slide-table cell formats of wave W1-G (record 0379): the host's
//! `set_slide_table_cell_<family>_format` /
//! `reset_slide_table_cell_<family>_format` /
//! `slide_table_cell_<family>_format` against the leaf's single
//! `edit_slide_table_cell_format` family over the shared
//! `litchi_iwa_table::cell_format` engine.

use litchi_iwa::keynote;
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
use litchi_keynote::Package;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

/// Native slide tables whose stored cells both routes read.
const CATALOG: &str = "iwork/keynote/slide-table-read-native.key";
const SLIDE: usize = 0;
const TABLE: usize = 0;

fn host_model(editor: &keynote::Draft) -> TestResult<u64> {
    Ok(editor
        .slide_tables(SLIDE)?
        .first()
        .ok_or("no slide table")?
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
            if package.slide_table_cell_format(SLIDE, TABLE, at).ok() == Some(DataFormat::Automatic)
                && host_read(input, at).ok() == Some(DataFormat::Automatic)
            {
                return Ok(at);
            }
        }
    }
    Err("no automatic stored cell".into())
}

fn host_read(bytes: &[u8], at: At) -> TestResult<DataFormat> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    Ok(editor.slide_table_cell_data_format(
        SLIDE,
        model,
        at.row() as usize,
        at.column() as usize,
    )?)
}

fn leaf_read(bytes: &[u8], at: At) -> TestResult<DataFormat> {
    Ok(Package::from_bytes(bytes)?.slide_table_cell_format(SLIDE, TABLE, at)?)
}

fn host_set(source: &[u8], at: At, format: &DataFormat) -> TestResult<Vec<u8>> {
    let mut editor = keynote::Draft::from_bytes(source)?;
    let model = host_model(&editor)?;
    editor.set_slide_table_cell_data_format(
        SLIDE,
        model,
        at.row() as usize,
        at.column() as usize,
        format.clone(),
    )?;
    Ok(editor.to_bytes()?)
}

/// The host's family-specific setter for `format`.
fn host_family_set(source: &[u8], at: At, format: &DataFormat) -> TestResult<Vec<u8>> {
    let mut editor = keynote::Draft::from_bytes(source)?;
    let model = host_model(&editor)?;
    let (row, column) = (at.row() as usize, at.column() as usize);
    match format {
        DataFormat::Currency(value) => {
            editor.set_slide_table_cell_currency_format(SLIDE, model, row, column, *value)?
        },
        DataFormat::Percentage(value) => {
            editor.set_slide_table_cell_percentage_format(SLIDE, model, row, column, *value)?
        },
        DataFormat::Scientific(value) => {
            editor.set_slide_table_cell_scientific_format(SLIDE, model, row, column, *value)?
        },
        DataFormat::Fraction(value) => {
            editor.set_slide_table_cell_fraction_format(SLIDE, model, row, column, *value)?
        },
        _ => return Err("no family-specific host setter".into()),
    }
    Ok(editor.to_bytes()?)
}

/// The host's family-specific read of the family of `format`.
fn host_family_read(bytes: &[u8], at: At, format: &DataFormat) -> TestResult<String> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    let (row, column) = (at.row() as usize, at.column() as usize);
    Ok(match format {
        DataFormat::Currency(_) => format!(
            "{:?}",
            editor.slide_table_cell_currency_format(SLIDE, model, row, column)?
        ),
        DataFormat::Percentage(_) => format!(
            "{:?}",
            editor.slide_table_cell_percentage_format(SLIDE, model, row, column)?
        ),
        DataFormat::Scientific(_) => format!(
            "{:?}",
            editor.slide_table_cell_scientific_format(SLIDE, model, row, column)?
        ),
        DataFormat::Fraction(_) => format!(
            "{:?}",
            editor.slide_table_cell_fraction_format(SLIDE, model, row, column)?
        ),
        _ => return Err("no family-specific host read".into()),
    })
}

fn leaf_set(source: &[u8], at: At, format: &DataFormat) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let commit = package
        .edit_slide_table_cell_format(SLIDE, TABLE, at)?
        .set(format.clone())
        .commit()?;
    let restored = commit
        .package()
        .apply_slide_table_cell_format(&commit.patch().inverse())?;
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
    Case::new(name, Format::Keynote, tier, evidence, input)
        .host(|source| host_set(source, at, format))
        .leaf(|source| leaf_set(source, at, format))
        .read("host slide_table_cell_data_format", move |bytes| {
            host_read(bytes, at)
        })
        .read("leaf slide_table_cell_format", move |bytes| {
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
fn keynote_slide_table_cell_data_format_set_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    // Currency, Date and Time and Duration convert the cell's kind and
    // scalar as the host's compatibility writer does (record 0412a).
    for (_, format) in samples()?.iter().filter(|(family, _)| *family != "text") {
        set_case(
            "keynote_slide_table_cell_data_format_set_parity",
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
fn keynote_slide_table_cell_format_set_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    let samples = samples()?;
    for ((_, format), replacement) in samples.iter().take(ADMITTED).zip(replacements()?).skip(1) {
        let formatted = host_formatted(&input, at, format)?;
        Case::new(
            "keynote_slide_table_cell_format_set_parity",
            Format::Keynote,
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
fn keynote_slide_table_cell_format_reset_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    for (family, format) in samples()?.into_iter().take(ADMITTED) {
        let formatted = host_formatted(&input, at, &format)?;
        Case::new(
            "keynote_slide_table_cell_format_reset_parity",
            Format::Keynote,
            Tier::B,
            Evidence::Queued,
            &formatted,
        )
        .host(|source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            let (row, column) = (at.row() as usize, at.column() as usize);
            let changed =
                match family {
                    "number" => {
                        // The host has no Number-only reset on slide tables.
                        editor.set_slide_table_cell_data_format(
                            SLIDE,
                            model,
                            row,
                            column,
                            DataFormat::Automatic,
                        )?;
                        true
                    },
                    "currency" => {
                        editor.reset_slide_table_cell_currency_format(SLIDE, model, row, column)?
                    },
                    "percentage" => editor
                        .reset_slide_table_cell_percentage_format(SLIDE, model, row, column)?,
                    "scientific" => editor
                        .reset_slide_table_cell_scientific_format(SLIDE, model, row, column)?,
                    "fraction" => {
                        editor.reset_slide_table_cell_fraction_format(SLIDE, model, row, column)?
                    },
                    _ => return Err("unmapped family".into()),
                };
            if !changed {
                return Err("host reset reported no change".into());
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| leaf_set(source, at, &DataFormat::Automatic))
        .read("host slide_table_cell_data_format", move |bytes| {
            host_read(bytes, at)
        })
        .read("leaf slide_table_cell_format", move |bytes| {
            leaf_read(bytes, at)
        })
        .run()?;
    }
    Ok(())
}

/// The host's reads and the leaf's single read agree on the host's output
/// of every admitted family.
#[test]
fn keynote_slide_table_cell_format_read_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    // The host reads Number through the generic read only (its Number
    // family is already the leaf's dedicated `slide_table_cell_number_format`).
    for (_, format) in samples()?.into_iter().take(ADMITTED).skip(1) {
        let formatted = host_formatted(&input, at, &format)?;
        Case::new(
            "keynote_slide_table_cell_format_read_parity",
            Format::Keynote,
            Tier::A,
            Evidence::Queued,
            &formatted,
        )
        .read_only()
        .read("host family read", |bytes| {
            host_family_read(bytes, at, &format)
        })
        .host(|source| Ok(keynote::Draft::from_bytes(source)?.to_bytes()?))
        .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
        .read("host slide_table_cell_data_format", move |bytes| {
            host_read(bytes, at)
        })
        .read("leaf slide_table_cell_format", move |bytes| {
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
fn keynote_slide_table_cell_format_divergence_is_reported() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    let samples = samples()?;
    for (family, format) in samples.iter().filter(|(family, _)| *family == "text") {
        assert!(host_set(&input.bytes, at, format).is_err(), "{family}");
        assert!(leaf_set(&input.bytes, at, format).is_err(), "{family}");
    }
    Ok(())
}

/// Callgrind probe (record 0379): a Currency replacement on the host-written
/// EUR Currency cell, host family setter against the leaf family, ABBA.
/// `--toggle-collect=*probe_host*` / `*probe_leaf*` on
/// `keynote_cell_format_cases::keynote_cell_format_cost_abba`.
#[test]
#[ignore = "callgrind probe; run explicitly"]
fn keynote_cell_format_cost_abba() -> TestResult {
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
                .edit_slide_table_cell_format(SLIDE, TABLE, at)?
                .set(replacement.clone())
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

/// Numeral systems (record 0370k): the leaf reads every numeral system the
/// host writes.
#[test]
fn keynote_slide_table_cell_numeral_system_format_read_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    for numeral in crate::numbers_cell_data_format_cases::numeral_samples()? {
        let format = DataFormat::from(numeral);
        let formatted = host_formatted(&input, at, &format)?;
        Case::new(
            "keynote_slide_table_cell_numeral_system_format_read_parity",
            Format::Keynote,
            Tier::A,
            Evidence::Queued,
            &formatted,
        )
        .read_only()
        .host(|source| Ok(keynote::Draft::from_bytes(source)?.to_bytes()?))
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
    let mut editor = keynote::Draft::from_bytes(source)?;
    let model = host_model(&editor)?;
    let (row, column) = (at.row() as usize, at.column() as usize);
    match format {
        DataFormat::Checkbox(value) => {
            editor.set_slide_table_cell_checkbox_format(SLIDE, model, row, column, *value)?
        },
        DataFormat::StarRating(value) => {
            editor.set_slide_table_cell_star_rating_format(SLIDE, model, row, column, *value)?
        },
        DataFormat::Slider(value) => {
            editor.set_slide_table_cell_slider_format(SLIDE, model, row, column, value.clone())?
        },
        DataFormat::Stepper(value) => {
            editor.set_slide_table_cell_stepper_format(SLIDE, model, row, column, value.clone())?
        },
        _ => return Err("not a scalar control".into()),
    }
    Ok(editor.to_bytes()?)
}

/// The host's control family reset for `format`'s family.
fn host_control_reset(source: &[u8], at: At, format: &DataFormat) -> TestResult<Vec<u8>> {
    let mut editor = keynote::Draft::from_bytes(source)?;
    let model = host_model(&editor)?;
    let (row, column) = (at.row() as usize, at.column() as usize);
    let changed = match format {
        DataFormat::Checkbox(_) => {
            editor.reset_slide_table_cell_checkbox_format(SLIDE, model, row, column)?
        },
        DataFormat::StarRating(_) => {
            editor.reset_slide_table_cell_star_rating_format(SLIDE, model, row, column)?
        },
        DataFormat::Slider(_) => {
            editor.reset_slide_table_cell_slider_format(SLIDE, model, row, column)?
        },
        DataFormat::Stepper(_) => {
            editor.reset_slide_table_cell_stepper_format(SLIDE, model, row, column)?
        },
        _ => return Err("not a scalar control".into()),
    };
    if !changed {
        return Err("host reset reported no change".into());
    }
    Ok(editor.to_bytes()?)
}

/// The host's control family read of `format`'s family.
fn host_control_read(bytes: &[u8], at: At, format: &DataFormat) -> TestResult<String> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    let (row, column) = (at.row() as usize, at.column() as usize);
    Ok(match format {
        DataFormat::Checkbox(_) => {
            format!(
                "{:?}",
                editor
                    .slide_table_cell_checkbox_format(SLIDE, model, row, column)?
                    .map(DataFormat::from)
            )
        },
        DataFormat::StarRating(_) => format!(
            "{:?}",
            editor
                .slide_table_cell_star_rating_format(SLIDE, model, row, column)?
                .map(DataFormat::from)
        ),
        DataFormat::Slider(_) => {
            format!(
                "{:?}",
                editor
                    .slide_table_cell_slider_format(SLIDE, model, row, column)?
                    .map(DataFormat::from)
            )
        },
        DataFormat::Stepper(_) => {
            format!(
                "{:?}",
                editor
                    .slide_table_cell_stepper_format(SLIDE, model, row, column)?
                    .map(DataFormat::from)
            )
        },
        _ => return Err("not a scalar control".into()),
    })
}

/// Automatic to each scalar control through the host's family setters
/// (record 0412).
#[test]
fn keynote_slide_table_cell_control_format_set_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    for (_, format) in control_samples()? {
        Case::new(
            "keynote_slide_table_cell_control_format_set_parity",
            Format::Keynote,
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
fn keynote_slide_table_cell_control_format_reset_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    for (_, format) in control_samples()? {
        let formatted = Fixture {
            path: input.path,
            bytes: host_control_set(&input.bytes, at, &format)?,
        };
        Case::new(
            "keynote_slide_table_cell_control_format_reset_parity",
            Format::Keynote,
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
fn keynote_slide_table_cell_control_format_read_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    for (_, format) in control_samples()? {
        let formatted = Fixture {
            path: input.path,
            bytes: host_control_set(&input.bytes, at, &format)?,
        };
        Case::new(
            "keynote_slide_table_cell_control_format_read_parity",
            Format::Keynote,
            Tier::A,
            Evidence::Queued,
            &formatted,
        )
        .read_only()
        .host(|source| Ok(keynote::Draft::from_bytes(source)?.to_bytes()?))
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
fn keynote_slide_table_cell_numeral_system_format_set_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    for numeral in crate::numbers_cell_data_format_cases::numeral_samples()? {
        let format = DataFormat::from(numeral);
        Case::new(
            "keynote_slide_table_cell_numeral_system_format_set_parity",
            Format::Keynote,
            Tier::B,
            Evidence::Queued,
            &input,
        )
        .host(|source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            let (row, column) = (at.row() as usize, at.column() as usize);
            editor
                .set_slide_table_cell_numeral_system_format(SLIDE, model, row, column, numeral)?;
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
fn keynote_slide_table_cell_numeral_system_format_reset_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    for numeral in crate::numbers_cell_data_format_cases::numeral_samples()? {
        let formatted = host_formatted(&input, at, &DataFormat::from(numeral))?;
        Case::new(
            "keynote_slide_table_cell_numeral_system_format_reset_parity",
            Format::Keynote,
            Tier::B,
            Evidence::Queued,
            &formatted,
        )
        .host(|source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            let (row, column) = (at.row() as usize, at.column() as usize);
            if !editor.reset_slide_table_cell_numeral_system_format(SLIDE, model, row, column)? {
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

/// Callgrind probe (record 0412): Automatic to a control format on the first
/// automatic stored cell, host family setter against the leaf, ABBA.
#[test]
#[ignore = "callgrind probe; run explicitly"]
fn keynote_slide_table_cell_control_format_cost_abba() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    let format = control_samples()?[0].1.clone();
    crate::pages_cost_probes::abba(
        &input,
        |source| host_control_set(source, at, &format),
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_slide_table_cell_format(SLIDE, TABLE, at)?
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
fn keynote_slide_table_cell_numeral_format_cost_abba() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    let format = DataFormat::from(crate::numbers_cell_data_format_cases::numeral_samples()?[0]);
    crate::pages_cost_probes::abba(
        &input,
        |source| host_set(source, at, &format),
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_slide_table_cell_format(SLIDE, TABLE, at)?
                .set(format.clone())
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

/// The host's Date and Time or Duration family setter, reset and read
/// (record 0412a).
fn host_date_set(source: &[u8], at: At, format: &DataFormat) -> TestResult<Vec<u8>> {
    let mut editor = keynote::Draft::from_bytes(source)?;
    let model = host_model(&editor)?;
    let (row, column) = (at.row() as usize, at.column() as usize);
    match format {
        DataFormat::DateTime(value) => editor.set_slide_table_cell_date_time_format(
            SLIDE,
            model,
            row,
            column,
            value.clone(),
        )?,
        DataFormat::Duration(value) => {
            editor.set_slide_table_cell_duration_format(SLIDE, model, row, column, *value)?
        },
        _ => return Err("not a date or duration".into()),
    }
    Ok(editor.to_bytes()?)
}

fn host_date_reset(source: &[u8], at: At, format: &DataFormat) -> TestResult<Vec<u8>> {
    let mut editor = keynote::Draft::from_bytes(source)?;
    let model = host_model(&editor)?;
    let (row, column) = (at.row() as usize, at.column() as usize);
    let changed = match format {
        DataFormat::DateTime(_) => {
            editor.reset_slide_table_cell_date_time_format(SLIDE, model, row, column)?
        },
        DataFormat::Duration(_) => {
            editor.reset_slide_table_cell_duration_format(SLIDE, model, row, column)?
        },
        _ => return Err("not a date or duration".into()),
    };
    if !changed {
        return Err("host reset reported no change".into());
    }
    Ok(editor.to_bytes()?)
}

fn host_date_read(bytes: &[u8], at: At, format: &DataFormat) -> TestResult<String> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    let (row, column) = (at.row() as usize, at.column() as usize);
    Ok(match format {
        DataFormat::DateTime(_) => format!(
            "{:?}",
            editor
                .slide_table_cell_date_time_format(SLIDE, model, row, column)?
                .map(DataFormat::from)
        ),
        DataFormat::Duration(_) => format!(
            "{:?}",
            editor
                .slide_table_cell_duration_format(SLIDE, model, row, column)?
                .map(DataFormat::from)
        ),
        _ => return Err("not a date or duration".into()),
    })
}

fn date_samples() -> TestResult<Vec<DataFormat>> {
    Ok(samples()?
        .into_iter()
        .filter(|(family, _)| *family == "date_time" || *family == "duration")
        .map(|(_, format)| format)
        .collect())
}

/// Automatic to Date and Time and Duration on a native number cell through
/// the host's family setters: both routes convert the cell (record 0412a).
#[test]
fn keynote_slide_table_cell_date_duration_format_set_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    for format in date_samples()? {
        Case::new(
            "keynote_slide_table_cell_date_duration_format_set_parity",
            Format::Keynote,
            Tier::B,
            Evidence::Queued,
            &input,
        )
        .host(|source| host_date_set(source, at, &format))
        .leaf(|source| leaf_set(source, at, &format))
        .read("host family read", |bytes| {
            host_date_read(bytes, at, &format)
        })
        .read("leaf cell format", move |bytes| {
            leaf_read(bytes, at).map(|value| format!("{:?}", Some(value)))
        })
        .run()?;
    }
    Ok(())
}

/// Each converted cell reset to Automatic through the host's family resets
/// (record 0412a).
#[test]
fn keynote_slide_table_cell_date_duration_format_reset_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    for format in date_samples()? {
        let formatted = Fixture {
            path: input.path,
            bytes: host_date_set(&input.bytes, at, &format)?,
        };
        Case::new(
            "keynote_slide_table_cell_date_duration_format_reset_parity",
            Format::Keynote,
            Tier::B,
            Evidence::Queued,
            &formatted,
        )
        .host(|source| host_date_reset(source, at, &format))
        .leaf(|source| leaf_set(source, at, &DataFormat::Automatic))
        .read("host table_cell_data_format", move |bytes| {
            host_read(bytes, at)
        })
        .read("leaf cell format", move |bytes| leaf_read(bytes, at))
        .run()?;
    }
    Ok(())
}

/// The host's Date and Time and Duration reads and the leaf's read agree on
/// the host's converted cells (record 0412a; a leaf control-reader defect
/// refused them before).
#[test]
fn keynote_slide_table_cell_date_duration_format_read_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    for format in date_samples()? {
        let formatted = Fixture {
            path: input.path,
            bytes: host_date_set(&input.bytes, at, &format)?,
        };
        Case::new(
            "keynote_slide_table_cell_date_duration_format_read_parity",
            Format::Keynote,
            Tier::A,
            Evidence::Queued,
            &formatted,
        )
        .read_only()
        .host(|source| Ok(keynote::Draft::from_bytes(source)?.to_bytes()?))
        .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
        .read("host family read", |bytes| {
            host_date_read(bytes, at, &format)
        })
        .read("leaf cell format", move |bytes| {
            leaf_read(bytes, at).map(|value| format!("{:?}", Some(value)))
        })
        .run()?;
        assert_eq!(leaf_read(&formatted.bytes, at)?, format);
    }
    Ok(())
}

/// B2 rewritten to hold text by the host's cell writer on
/// `table-discovery.key`: the read fixture's table holds formulas whose
/// cached values both cell writers refuse to refresh (record 0412a).
fn text_input(input: &Fixture) -> TestResult<(Fixture, At)> {
    let at = At::new(1, 1);
    let mut editor = keynote::Draft::from_bytes(&input.bytes)?;
    let model = host_model(&editor)?;
    editor.set_slide_table_cell(
        SLIDE,
        model,
        at.row() as usize,
        at.column() as usize,
        litchi_iwa::keynote::TableCellValue::Text("abc".into()),
    )?;
    Ok((
        Fixture {
            path: input.path,
            bytes: editor.to_bytes()?,
        },
        at,
    ))
}

/// A Keynote deck whose slide table holds no formulas.
const TEXT_FIXTURE: &str = "iwork/keynote/table-discovery.key";

/// Text on a text cell naming an implicit Number entry through the host's
/// family setter (the cell is converted to an explicit Text cell), then its
/// reset on the host's output (record 0412a).
#[test]
fn keynote_slide_table_cell_text_format_parity() -> TestResult {
    let (input, at) = text_input(&fixture(TEXT_FIXTURE)?)?;
    let format = DataFormat::Text(Text);
    Case::new(
        "keynote_slide_table_cell_text_format_parity",
        Format::Keynote,
        Tier::B,
        Evidence::Queued,
        &input,
    )
    .host(|source| {
        let mut editor = keynote::Draft::from_bytes(source)?;
        let model = host_model(&editor)?;
        let (row, column) = (at.row() as usize, at.column() as usize);
        editor.set_slide_table_cell_text_format(SLIDE, model, row, column)?;
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| leaf_set(source, at, &format))
    .read("host table_cell_data_format", move |bytes| {
        host_read(bytes, at)
    })
    .read("leaf cell format", move |bytes| leaf_read(bytes, at))
    .run()?;
    let formatted = host_formatted(&input, at, &format)?;
    Case::new(
        "keynote_slide_table_cell_text_format_parity",
        Format::Keynote,
        Tier::B,
        Evidence::Queued,
        &formatted,
    )
    .host(|source| {
        let mut editor = keynote::Draft::from_bytes(source)?;
        let model = host_model(&editor)?;
        let (row, column) = (at.row() as usize, at.column() as usize);
        if !editor.reset_slide_table_cell_text_format(SLIDE, model, row, column)? {
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
    Ok(())
}

/// The host's Text read and the leaf's read agree on the host's output
/// (record 0412a).
#[test]
fn keynote_slide_table_cell_text_format_read_parity() -> TestResult {
    let (input, at) = text_input(&fixture(TEXT_FIXTURE)?)?;
    let formatted = host_formatted(&input, at, &DataFormat::Text(Text))?;
    Case::new(
        "keynote_slide_table_cell_text_format_read_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Queued,
        &formatted,
    )
    .read_only()
    .host(|source| Ok(keynote::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
    .read("host family read", |bytes| {
        let editor = keynote::Draft::from_bytes(bytes)?;
        let model = host_model(&editor)?;
        let (row, column) = (at.row() as usize, at.column() as usize);
        Ok(format!(
            "{:?}",
            editor
                .slide_table_cell_text_format(SLIDE, model, row, column)?
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

/// One custom Number format (record 0412a).
fn custom_sample() -> TestResult<DataFormat> {
    use litchi_iwa_common::table::cell::data_format::Custom;
    use litchi_iwa_common::table::cell::data_format::custom::{
        Name, Number as CustomNumber, NumberPattern,
    };
    Ok(DataFormat::Custom(Custom::from(CustomNumber::new(
        Name::new("Grouped")?,
        NumberPattern::new("#,##0.00")?,
    ))))
}

fn host_custom_set(source: &[u8], at: At, format: &DataFormat) -> TestResult<Vec<u8>> {
    let DataFormat::Custom(custom) = format else {
        return Err("not a custom format".into());
    };
    let mut editor = keynote::Draft::from_bytes(source)?;
    let model = host_model(&editor)?;
    let (row, column) = (at.row() as usize, at.column() as usize);
    editor.set_slide_table_cell_custom_format(SLIDE, model, row, column, custom.clone())?;
    Ok(editor.to_bytes()?)
}

fn host_custom_read(bytes: &[u8], at: At) -> TestResult<String> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    let (row, column) = (at.row() as usize, at.column() as usize);
    Ok(format!(
        "{:?}",
        editor
            .slide_table_cell_custom_format(SLIDE, model, row, column)?
            .map(DataFormat::Custom)
    ))
}

/// Automatic to a custom Number format through the host's family setter,
/// then its reset (the registry is the document's `TSA.DocumentArchive`
/// custom-format list, record 0412a). The host mints a random registry
/// UUID, so the outputs are compared as a read differential: both reopen,
/// every host and leaf read of both outputs agrees, the leaf inverse
/// restores the source, and the reset of each route's output restores
/// Automatic on both (tier B with a queued native check).
#[test]
fn keynote_slide_table_cell_custom_format_parity() -> TestResult {
    let (_tier, _evidence) = (Tier::B, Evidence::Queued);
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    let format = custom_sample()?;
    let host = host_custom_set(&input.bytes, at, &format)?;
    let leaf = leaf_set(&input.bytes, at, &format)?;
    assert_eq!(leaf.restored.as_deref(), Some(input.bytes.as_slice()));
    for bytes in [&host, &leaf.output] {
        Package::from_bytes(bytes)?;
        assert_eq!(host_custom_read(bytes, at)?, format!("{:?}", Some(&format)));
        assert_eq!(leaf_read(bytes, at)?, format);
        assert_eq!(host_read(bytes, at)?, format);
    }
    for bytes in [&host, &leaf.output] {
        let mut editor = keynote::Draft::from_bytes(bytes)?;
        let model = host_model(&editor)?;
        let (row, column) = (at.row() as usize, at.column() as usize);
        assert!(editor.reset_slide_table_cell_custom_format(SLIDE, model, row, column)?);
        let host_reset = editor.to_bytes()?;
        let leaf_reset = leaf_set(bytes, at, &DataFormat::Automatic)?;
        for reset in [&host_reset, &leaf_reset.output] {
            assert_eq!(leaf_read(reset, at)?, DataFormat::Automatic);
            assert_eq!(host_read(reset, at)?, DataFormat::Automatic);
        }
    }
    Ok(())
}

/// The host's custom read and the leaf's read agree on the host's output
/// (record 0412a).
#[test]
fn keynote_slide_table_cell_custom_format_read_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    let format = custom_sample()?;
    let formatted = Fixture {
        path: input.path,
        bytes: host_custom_set(&input.bytes, at, &format)?,
    };
    Case::new(
        "keynote_slide_table_cell_custom_format_read_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Queued,
        &formatted,
    )
    .read_only()
    .host(|source| Ok(keynote::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
    .read("host family read", |bytes| host_custom_read(bytes, at))
    .read("leaf cell format", move |bytes| {
        leaf_read(bytes, at).map(|value| format!("{:?}", Some(value)))
    })
    .run()?;
    assert_eq!(leaf_read(&formatted.bytes, at)?, format);
    Ok(())
}

fn host_pop_up(
    source: &[u8],
    at: At,
    menu: Option<&litchi_iwa_common::table::cell::data_format::pop_up_menu::PopUpMenu>,
) -> TestResult<Vec<u8>> {
    let mut editor = keynote::Draft::from_bytes(source)?;
    let model = host_model(&editor)?;
    let (row, column) = (at.row() as usize, at.column() as usize);
    match menu {
        Some(menu) => editor.set_slide_table_cell_pop_up_menu_format(
            SLIDE,
            model,
            row,
            column,
            menu.clone(),
        )?,
        None => {
            if !editor.reset_slide_table_cell_pop_up_menu_format(SLIDE, model, row, column)? {
                return Err("host reset reported no change".into());
            }
        },
    }
    Ok(editor.to_bytes()?)
}

fn host_pop_up_read(bytes: &[u8], at: At) -> TestResult<String> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    let (row, column) = (at.row() as usize, at.column() as usize);
    Ok(format!(
        "{:?}",
        editor
            .slide_table_cell_pop_up_menu_format(SLIDE, model, row, column)?
            .map(DataFormat::from)
    ))
}

/// A Pop-Up Menu on a text cell through the host's family setter, a second
/// menu on the host's output, and its reset, against the leaf over the
/// shared Pop-Up Menu publication (record 0412a).
#[test]
fn keynote_slide_table_cell_pop_up_menu_format_parity() -> TestResult {
    use litchi_iwa_common::table::cell::data_format::pop_up_menu::PopUpMenu;
    let (input, at) = text_input(&fixture(TEXT_FIXTURE)?)?;
    let first = PopUpMenu::new(["abc", "def"])?;
    let second = PopUpMenu::new(["abc", "ghi", "jkl"])?;
    let menu = |value: &PopUpMenu| DataFormat::PopUpMenu(value.clone());
    let cases = [
        (input.bytes.clone(), Some(first.clone())),
        (
            host_pop_up(&input.bytes, at, Some(&first))?,
            Some(second.clone()),
        ),
        (host_pop_up(&input.bytes, at, Some(&first))?, None),
    ];
    for (bytes, target) in cases {
        let fixture = Fixture {
            path: input.path,
            bytes,
        };
        let format = target.as_ref().map_or(DataFormat::Automatic, menu);
        Case::new(
            "keynote_slide_table_cell_pop_up_menu_format_parity",
            Format::Keynote,
            Tier::B,
            Evidence::Queued,
            &fixture,
        )
        .leaf_structural_metadata()
        .host(|source| host_pop_up(source, at, target.as_ref()))
        .leaf(|source| leaf_set(source, at, &format))
        .read("host table_cell_data_format", move |bytes| {
            host_read(bytes, at)
        })
        .read("leaf cell format", move |bytes| leaf_read(bytes, at))
        .run()?;
    }
    Ok(())
}

/// The host's Pop-Up Menu read and the leaf's read agree on the host's
/// output (record 0412a).
#[test]
fn keynote_slide_table_cell_pop_up_menu_format_read_parity() -> TestResult {
    use litchi_iwa_common::table::cell::data_format::pop_up_menu::PopUpMenu;
    let (input, at) = text_input(&fixture(TEXT_FIXTURE)?)?;
    let menu = PopUpMenu::new(["abc", "def"])?;
    let formatted = Fixture {
        path: input.path,
        bytes: host_pop_up(&input.bytes, at, Some(&menu))?,
    };
    Case::new(
        "keynote_slide_table_cell_pop_up_menu_format_read_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Queued,
        &formatted,
    )
    .read_only()
    .host(|source| Ok(keynote::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
    .read("host family read", |bytes| host_pop_up_read(bytes, at))
    .read("leaf cell format", move |bytes| {
        leaf_read(bytes, at).map(|value| format!("{:?}", Some(value)))
    })
    .run()?;
    assert_eq!(
        leaf_read(&formatted.bytes, at)?,
        DataFormat::PopUpMenu(menu)
    );
    Ok(())
}

/// Callgrind probe (record 0412a): a Pop-Up Menu on the text cell, host
/// family setter against the leaf, ABBA.
#[test]
#[ignore = "callgrind probe; run explicitly"]
fn keynote_slide_table_cell_pop_up_menu_format_cost_abba() -> TestResult {
    use litchi_iwa_common::table::cell::data_format::pop_up_menu::PopUpMenu;
    let (input, at) = text_input(&fixture(TEXT_FIXTURE)?)?;
    let menu = PopUpMenu::new(["abc", "def"])?;
    let format = DataFormat::PopUpMenu(menu.clone());
    crate::pages_cost_probes::abba(
        &input,
        |source| host_pop_up(source, at, Some(&menu)),
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_slide_table_cell_format(SLIDE, TABLE, at)?
                .set(format.clone())
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

/// Callgrind probe (record 0412a): Automatic to Date and Time on the native
/// number cell, host family setter against the leaf, ABBA.
#[test]
#[ignore = "callgrind probe; run explicitly"]
fn keynote_slide_table_cell_date_time_format_cost_abba() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    let format = date_samples()?.into_iter().next().ok_or("no date sample")?;
    crate::pages_cost_probes::abba(
        &input,
        |source| host_date_set(source, at, &format),
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_slide_table_cell_format(SLIDE, TABLE, at)?
                .set(format.clone())
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

/// Callgrind probe (record 0412a): Automatic to a custom Number format,
/// host family setter against the leaf, ABBA.
#[test]
#[ignore = "callgrind probe; run explicitly"]
fn keynote_slide_table_cell_custom_format_cost_abba() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    let format = custom_sample()?;
    crate::pages_cost_probes::abba(
        &input,
        |source| host_custom_set(source, at, &format),
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_slide_table_cell_format(SLIDE, TABLE, at)?
                .set(format.clone())
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

/// Callgrind probe (record 0412a): Text on the text cell naming an implicit
/// Number entry, host family setter against the leaf, ABBA.
#[test]
#[ignore = "callgrind probe; run explicitly"]
fn keynote_slide_table_cell_text_format_cost_abba() -> TestResult {
    let (input, at) = text_input(&fixture(TEXT_FIXTURE)?)?;
    let format = DataFormat::Text(Text);
    crate::pages_cost_probes::abba(
        &input,
        |source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            let (row, column) = (at.row() as usize, at.column() as usize);
            editor.set_slide_table_cell_text_format(SLIDE, model, row, column)?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_slide_table_cell_format(SLIDE, TABLE, at)?
                .set(format.clone())
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}
