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

/// Automatic to Number on a stored cell through the host's generic setter.
#[test]
fn keynote_slide_table_cell_data_format_set_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    set_case(
        "keynote_slide_table_cell_data_format_set_parity",
        Tier::B,
        Evidence::Queued,
        &input,
        at,
        &samples()?[0].1,
    )
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
        assert_eq!(leaf_read(&formatted.bytes, at)?, format);
    }
    Ok(())
}

/// Where the routes differ (record 0379 §5): the engine refuses what the
/// host's compatibility writer converts. Automatic to a family other than
/// Number on a decimal cell, and a cell the host formatted as Date and
/// Time or Duration (the host leaves the stored number under a date or
/// duration kind, which the engine reads as an invalid cell). The host
/// refuses Text on a number cell; the leaf refuses it too.
#[test]
fn keynote_slide_table_cell_format_divergence_is_reported() -> TestResult {
    let input = fixture(CATALOG)?;
    let at = automatic_cell(&input.bytes)?;
    let samples = samples()?;
    for (family, format) in &samples[1..] {
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
/// host writes. Writing one is refused on these native cells, which name an
/// implicit Number format entry the engine keeps in place (the host's
/// compatibility writer interns a new entry instead); a numeral write is
/// covered on Numbers (`numbers_cell_data_format_cases`).
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
        assert_eq!(leaf_read(&formatted.bytes, at)?, format);
        assert!(leaf_set(&input.bytes, at, &format).is_err());
    }
    Ok(())
}
