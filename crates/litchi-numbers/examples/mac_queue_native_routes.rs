//! Writes the route packages of the macOS verification queue that start from a
//! workbook Numbers itself saved (`docs/unified-ops/0363-mac-queue.md`; records
//! 0432, 0433).
//!
//! ```text
//! cargo run -p litchi-numbers --example mac_queue_native_routes -- <output-directory>
//! ```
//!
//! The parity harness compares the leaf with the host, and on these inputs the
//! host reads less or fails, so these packages are written by the leaf alone. The
//! output is deterministic. One JSON line per package goes to
//! `<output-directory>/native-routes.jsonl`, which `tools/mac_queue/generate.py`
//! folds into the queue manifest.

use std::io::Write as _;
use std::path::{Path, PathBuf};

use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_numbers::Package;
use litchi_numbers::cell::conditional_highlight::{Condition, Rule, Style};
use litchi_numbers::formula::{BinaryOperator, CachedValue, Expression};
use litchi_numbers::table::At;
use litchi_numbers::table::cells::Input;
use litchi_numbers::table::pivot::{Aggregate, Category};

type Result<T = ()> = std::result::Result<T, Box<dyn std::error::Error>>;

const PIVOT: &str = "pivot-all-summary-functions-native.numbers";
const VOLATILE: &str = "volatile-highlight-native.numbers";
const BASIC: &str = "basic.numbers";
const SOURCE_SHEET: &str = "Sheet 1";
const SOURCE_TABLE: &str = "Table 1";
const PIVOT_SHEET: &str = "Table 1 Pivot";
const PIVOT_TABLE: &str = "Table 1 Pivot";
/// 2026-11-05 at midnight, in Apple seconds, as Numbers stores a typed date.
const LADDER_START: f64 = 815_529_600.0;
const SECONDS_PER_DAY: f64 = 86_400.0;
/// Rows 2 to 22 of columns C and D: six weeks of days, one "is today" rule each.
const LADDER_ROWS: std::ops::RangeInclusive<u32> = 1..=21;
const LADDER_COLUMNS: [u32; 2] = [2, 3];

fn fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers")
        .join(name)
}

fn open(name: &str) -> Result<Package> {
    Ok(Package::open(fixture(name))?)
}

fn red_rule(condition: Condition) -> Result<Rule> {
    Ok(Rule::new(
        condition,
        Style::with_fill(Rgba::new(1.0, 0.0, 0.0, 1.0, RgbColorSpace::Srgb)?),
    ))
}

fn category<'a>(
    categories: &'a [Category],
    label: &str,
    level: i32,
    aggregate: Aggregate,
) -> Result<&'a Category> {
    categories
        .iter()
        .find(|category| {
            category.label() == Some(label)
                && category.group_level() == level
                && category.aggregate() == aggregate
        })
        .ok_or_else(|| format!("no category {label} at level {level}").into())
}

/// Grand Total Average / 2, the formula of the source-table packages.
fn half_grand_average(categories: &[Category]) -> Result<Expression> {
    Ok(Expression::binary(
        BinaryOperator::Divide,
        Expression::pivot_category(category(categories, "Grand Total", 0, Aggregate::Average)?),
        Expression::number(2.0)?,
    )?)
}

struct Writer {
    directory: PathBuf,
    lines: Vec<String>,
}

impl Writer {
    fn emit(&mut self, stem: &str, source: &str, package: &Package) -> Result {
        let mut bytes = Vec::new();
        package.write_to(&mut bytes)?;
        std::fs::write(self.directory.join(format!("{stem}.file.numbers")), &bytes)?;
        self.lines.push(format!(
            "{{\"stem\": \"{stem}\", \"format\": \"numbers\", \"extension\": \"numbers\", \
             \"operation\": \"{stem}\", \"fixture\": \"test-data/iwork/numbers/{source}\", \
             \"selector\": \"\", \"property\": \"\", \"read_back\": null}}"
        ));
        Ok(())
    }
}

fn pivot_routes(writer: &mut Writer) -> Result {
    let package = open(PIVOT)?;
    writer.emit("numbers-native-pivot-control", PIVOT, &package)?;
    let categories = package.pivot_categories()?;

    let edited = package
        .edit_table_cells(SOURCE_SHEET, SOURCE_TABLE)?
        .set(At::new(1, 2), Input::number(55.0)?)?
        .commit()?;
    writer.emit("numbers-native-pivot-source-cell", PIVOT, edited.package())?;

    for (stem, cached) in [
        ("numbers-native-pivot-formula-source-cached", 32.5),
        // A cached value only a recalculation can correct.
        ("numbers-native-pivot-formula-source-stale-cache", 1.0),
    ] {
        let commit = package
            .edit_table_cells(SOURCE_SHEET, SOURCE_TABLE)?
            .set_formula_cached(
                At::new(1, 2),
                half_grand_average(&categories)?,
                CachedValue::number(cached)?,
            )?
            .commit()?;
        writer.emit(stem, PIVOT, commit.package())?;
    }
    let commit = package
        .edit_table_cells(SOURCE_SHEET, SOURCE_TABLE)?
        .set_formula(At::new(1, 2), half_grand_average(&categories)?)?
        .commit()?;
    writer.emit(
        "numbers-native-pivot-formula-source-no-cache",
        PIVOT,
        commit.package(),
    )?;

    let commit = package
        .edit_table_cells(PIVOT_SHEET, PIVOT_TABLE)?
        .set_formula_cached(
            At::new(4, 1),
            Expression::pivot_category(category(&categories, "West", 2, Aggregate::Maximum)?),
            CachedValue::number(1.0)?,
        )?
        .commit()?;
    writer.emit("numbers-native-pivot-formula-body", PIVOT, commit.package())?;
    Ok(())
}

fn volatile_routes(writer: &mut Writer) -> Result {
    let today = red_rule(Condition::DateIsToday)?;

    // The first rule on a table Numbers saved: over text, and on an empty cell.
    let basic = open(BASIC)?;
    let commit = basic
        .edit_table_cell_conditional_highlights(0usize, 0usize)?
        .set(At::new(1, 1), [today.clone()])?
        .commit()?;
    writer.emit(
        "numbers-native-first-rule-text-cell",
        BASIC,
        commit.package(),
    )?;
    let commit = basic
        .edit_table_cell_conditional_highlights(0usize, 0usize)?
        .set(At::new(3, 1), [today.clone()])?
        .commit()?;
    writer.emit(
        "numbers-native-first-rule-empty-cell",
        BASIC,
        commit.package(),
    )?;

    // Six weeks of dates, each with the rule: Numbers evaluates volatile rules
    // when it opens a file and stores the applied rule in the cell, so the
    // re-saved bytes say which cell it took for today, with nobody looking.
    let positions: Vec<At> = LADDER_ROWS
        .flat_map(|row| LADDER_COLUMNS.map(|column| At::new(row, column)))
        .collect();
    let mut cells = basic.edit_table_cells(0usize, 0usize)?;
    let mut day = LADDER_START;
    for position in &positions {
        cells = cells.set(*position, Input::date(day)?)?;
        day += SECONDS_PER_DAY;
    }
    let dated = cells.commit()?;
    let mut rules = dated
        .package()
        .edit_table_cell_conditional_highlights(0usize, 0usize)?;
    for position in &positions {
        rules = rules.set(*position, [today.clone()])?;
    }
    let commit = rules.commit()?;
    writer.emit(
        "numbers-native-first-rule-date-ladder",
        BASIC,
        commit.package(),
    )?;

    // Numbers' own family.
    let native = open(VOLATILE)?;
    let commit = native
        .edit_table_cell_conditional_highlights(0usize, 0usize)?
        .set(At::new(3, 0), [red_rule(Condition::DateIsTomorrow)?])?
        .commit()?;
    writer.emit(
        "numbers-native-volatile-add-tomorrow",
        VOLATILE,
        commit.package(),
    )?;
    let commit = native
        .edit_table_cell_conditional_highlights(0usize, 0usize)?
        .clear(At::new(2, 0))?
        .commit()?;
    writer.emit(
        "numbers-native-volatile-clear-middle",
        VOLATILE,
        commit.package(),
    )?;
    Ok(())
}

fn main() -> Result {
    let directory = std::env::args_os()
        .nth(1)
        .map(PathBuf::from)
        .ok_or("usage: mac_queue_native_routes <output-directory>")?;
    std::fs::create_dir_all(&directory)?;
    let mut writer = Writer {
        directory,
        lines: Vec::new(),
    };
    pivot_routes(&mut writer)?;
    volatile_routes(&mut writer)?;
    let mut out = std::fs::File::create(writer.directory.join("native-routes.jsonl"))?;
    for line in &writer.lines {
        writeln!(out, "{line}")?;
    }
    Ok(())
}
