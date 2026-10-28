//! Numbers cell data formats (record 0370k): the host's
//! `set_table_cell_data_format` / `table_cell_data_format` and its
//! numeral-system routes (`set_` / `reset_` / `table_cell_numeral_system_format`)
//! against the leaf's single `edit_table_cell_data_format` family.

use litchi_iwa::numbers;
use litchi_iwa_common::table::cell::data_format::currency::{
    Currency, CurrencyCode, CurrencyStyle,
};
use litchi_iwa_common::table::cell::data_format::number::{
    DecimalPlaces, FixedDecimalPlaces, NegativeStyle, Number, Percentage, Scientific,
    ThousandsSeparator,
};
use litchi_iwa_common::table::cell::data_format::numeral_system::{
    Base, NegativeStyle as NumeralNegativeStyle, NumeralSystem, Places,
};
use litchi_iwa_common::table::cell::data_format::{DataFormat, Fraction, FractionAccuracy};
use litchi_iwa_common::table::coordinate::At;
use litchi_numbers::Package;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;
use crate::numbers_cases::host_first_table_model_id;

/// The corpus fixture with a stored number cell both routes read as
/// Automatic and the shared engine admits (the other fixtures' number cells
/// are outside the engine's tile shape, or the host reads their metadata as
/// inconsistent).
const BASIC: &str = "iwork/numbers/custom-retirement-cleared.numbers";

/// Numeral systems the host writes: base 2 two's complement over eight
/// places, base 16 minimum width with a minus sign, base 36 fixed width.
pub(crate) fn numeral_samples() -> TestResult<Vec<NumeralSystem>> {
    Ok(vec![
        NumeralSystem::new(
            Base::new(2)?,
            Places::fixed(8)?,
            NumeralNegativeStyle::TwosComplement,
        )?,
        NumeralSystem::new(
            Base::new(16)?,
            Places::Minimum,
            NumeralNegativeStyle::MinusSign,
        )?,
        NumeralSystem::new(
            Base::new(36)?,
            Places::fixed(4)?,
            NumeralNegativeStyle::MinusSign,
        )?,
    ])
}

/// One value per decimal family the engine writes on a decimal cell.
fn decimal_samples() -> TestResult<Vec<DataFormat>> {
    Ok(vec![
        Number::new(
            DecimalPlaces::fixed(2)?,
            NegativeStyle::Parentheses,
            ThousandsSeparator::Shown,
        )
        .into(),
        Currency::new(
            CurrencyCode::new("EUR")?,
            DecimalPlaces::fixed(2)?,
            NegativeStyle::MinusSign,
            ThousandsSeparator::Shown,
            CurrencyStyle::Accounting,
        )
        .into(),
        Percentage::new(
            DecimalPlaces::fixed(1)?,
            NegativeStyle::Red,
            ThousandsSeparator::Hidden,
        )
        .into(),
        Scientific::new(FixedDecimalPlaces::new(3)?).into(),
        Fraction::new(FractionAccuracy::Quarters).into(),
    ])
}

fn host_read(bytes: &[u8], at: At) -> TestResult<DataFormat> {
    let editor = numbers::Draft::from_bytes(bytes)?;
    let model = host_first_table_model_id(&editor)?;
    Ok(editor.table_cell_data_format(model, at.row() as usize, at.column() as usize)?)
}

fn leaf_read(bytes: &[u8], at: At) -> TestResult<DataFormat> {
    Ok(Package::from_bytes(bytes)?.table_cell_data_format(0usize, 0usize, at)?)
}

fn host_set(source: &[u8], at: At, format: &DataFormat) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let model = host_first_table_model_id(&editor)?;
    editor.set_table_cell_data_format(
        model,
        at.row() as usize,
        at.column() as usize,
        format.clone(),
    )?;
    Ok(editor.to_bytes()?)
}

fn leaf_set(source: &[u8], at: At, format: &DataFormat) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let commit = package
        .edit_table_cell_data_format(0usize, 0usize, at)?
        .set(format.clone())
        .commit()?;
    let restored = commit
        .package()
        .apply_table_cell_data_format(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

/// The first stored number cell both routes read as Automatic.
fn decimal_cell(input: &[u8]) -> TestResult<At> {
    let package = Package::from_bytes(input)?;
    let table = package.sheets()[0].tables().next().ok_or("table")?;
    for row in 0..table.row_count() {
        for column in 0..table.column_count() {
            let at = At::new(row, column);
            let numeric = matches!(table.get(at), Some(litchi_numbers::cell::Value::Number(_)));
            if numeric
                && leaf_read(input, at).ok() == Some(DataFormat::Automatic)
                && host_read(input, at).ok() == Some(DataFormat::Automatic)
            {
                return Ok(at);
            }
        }
    }
    Err("no automatic number cell".into())
}

fn case(
    name: &'static str,
    tier: Tier,
    input: &Fixture,
    at: At,
    format: &DataFormat,
) -> TestResult {
    Case::new(name, Format::Numbers, tier, Evidence::Queued, input)
        .host(|source| host_set(source, at, format))
        .leaf(|source| leaf_set(source, at, format))
        .read("host table_cell_data_format", move |bytes| {
            host_read(bytes, at)
        })
        .read("leaf table_cell_data_format", move |bytes| {
            leaf_read(bytes, at)
        })
        .run()?;
    Ok(())
}

fn formatted(input: &Fixture, at: At, format: &DataFormat) -> TestResult<Fixture> {
    Ok(Fixture {
        path: input.path,
        bytes: host_set(&input.bytes, at, format)?,
    })
}

/// Automatic to each numeral system and to Number, through the host's
/// generic setter, and every write the host makes, replaced and reset.
#[test]
fn numbers_table_cell_data_format_set_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = fixture(BASIC)?;
    let at = decimal_cell(&input.bytes)?;
    let mut formats: Vec<DataFormat> = numeral_samples()?.into_iter().map(Into::into).collect();
    formats.push(decimal_samples()?.remove(0));
    for format in &formats {
        case(
            "numbers_table_cell_data_format_set_parity",
            Tier::B,
            &input,
            at,
            format,
        )?;
    }
    // Inside each family on the host's output, across families and back to
    // Automatic.
    let decimal = decimal_samples()?;
    let numerals = numeral_samples()?;
    let pairs: Vec<(DataFormat, DataFormat)> = vec![
        (numerals[0].into(), numerals[1].into()),
        (numerals[1].into(), numerals[2].into()),
        (numerals[0].into(), DataFormat::Automatic),
        (decimal[0].clone(), numerals[0].into()),
    ];
    for (first, second) in pairs {
        let start = formatted(&input, at, &first)?;
        case(
            "numbers_table_cell_data_format_set_parity",
            Tier::B,
            &start,
            at,
            &second,
        )
        .map_err(|error| format!("{first:?} -> {second:?}: {error}"))?;
    }
    Ok(())
}

/// The host's reads and the leaf's read agree on every host-written decimal
/// family and numeral system.
#[test]
fn numbers_table_cell_data_format_read_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    let input = fixture(BASIC)?;
    let at = decimal_cell(&input.bytes)?;
    let mut formats = decimal_samples()?;
    formats.extend(numeral_samples()?.into_iter().map(DataFormat::from));
    for format in formats {
        let start = formatted(&input, at, &format)?;
        assert_eq!(host_read(&start.bytes, at)?, format);
        if let Ok(read) = leaf_read(&start.bytes, at) {
            assert_eq!(read, format, "{format:?}");
        } else {
            // The engine reads the decimal families only on a decimal cell
            // the host leaves in the engine's tile shape (record 0379 §5).
            assert!(
                !matches!(format, DataFormat::Number(_) | DataFormat::NumeralSystem(_)),
                "{format:?}"
            );
        }
    }
    Ok(())
}

/// The host's numeral-system routes against the leaf family.
#[test]
fn numbers_table_cell_numeral_system_format_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let input = fixture(BASIC)?;
    let at = decimal_cell(&input.bytes)?;
    let (row, column) = (at.row() as usize, at.column() as usize);
    for numeral in numeral_samples()? {
        Case::new(
            "numbers_table_cell_numeral_system_format_parity",
            Format::Numbers,
            Tier::B,
            Evidence::Queued,
            &input,
        )
        .host(|source| {
            let mut editor = numbers::Draft::from_bytes(source)?;
            let model = host_first_table_model_id(&editor)?;
            editor.set_table_cell_numeral_system_format(model, row, column, numeral)?;
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| leaf_set(source, at, &numeral.into()))
        .read("host table_cell_numeral_system_format", move |bytes| {
            let editor = numbers::Draft::from_bytes(bytes)?;
            let model = host_first_table_model_id(&editor)?;
            Ok(editor.table_cell_numeral_system_format(model, row, column)?)
        })
        .run()?;
        let start = formatted(&input, at, &numeral.into())?;
        Case::new(
            "numbers_table_cell_numeral_system_format_parity",
            Format::Numbers,
            Tier::B,
            Evidence::Queued,
            &start,
        )
        .host(|source| {
            let mut editor = numbers::Draft::from_bytes(source)?;
            let model = host_first_table_model_id(&editor)?;
            if !editor.reset_table_cell_numeral_system_format(model, row, column)? {
                return Err("host reset reported no change".into());
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| leaf_set(source, at, &DataFormat::Automatic))
        .read("host table_cell_data_format", move |bytes| {
            host_read(bytes, at)
        })
        .run()?;
    }
    Ok(())
}

/// A numeral system replaced by another family: the host's compatibility
/// writer interns the new format at the next list key before it releases
/// the old one, the leaf runs two engine steps (the first to Automatic,
/// which hands the released last key back). Both read the target format;
/// the list keys differ (record 0370k, a documented difference).
#[test]
fn numbers_table_cell_data_format_cross_family_difference() -> TestResult {
    let input = fixture(BASIC)?;
    let at = decimal_cell(&input.bytes)?;
    let numerals = numeral_samples()?;
    let target = decimal_samples()?.remove(0);
    let start = formatted(&input, at, &numerals[2].into())?;
    let host = host_set(&start.bytes, at, &target)?;
    let leaf = leaf_set(&start.bytes, at, &target)?;
    assert_eq!(host_read(&host, at)?, target);
    assert_eq!(leaf_read(&leaf.output, at)?, target);
    assert_eq!(host_read(&leaf.output, at)?, target);
    Ok(())
}
