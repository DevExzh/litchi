//! Remove one Numbers table by workbook-wide table index.
//!
//! The focused lifecycle transaction removes the table's private graph,
//! saves a distinct output, reopens it, and checks the inverse against the
//! exact input package.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::env;
use std::error::Error;
use std::io;

use litchi_numbers::{Package, SheetSelector, TableSelector};

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or_else(|| invalid("missing input path"))?;
    let output = arguments
        .next()
        .ok_or_else(|| invalid("missing output path"))?;
    let table_index: usize = arguments
        .next()
        .ok_or_else(|| invalid("missing table index"))?
        .parse()
        .map_err(|_| invalid("table index must be a non-negative integer"))?;
    if arguments.next().is_some() {
        return Err(invalid("unexpected trailing arguments"));
    }
    if input == output {
        return Err(invalid("input and output paths must differ"));
    }

    let package = Package::open(&input)?;
    let (sheet_index, local_index) = locate_table(&package, table_index)?;
    let before_count = package.sheets()[sheet_index].tables().count();
    let original = exact_bytes(&package)?;
    let commit = package
        .edit_table_lifecycle(SheetSelector::index(sheet_index))?
        .remove(TableSelector::index(local_index))?
        .commit()?;
    let committed = commit.package();
    assert_count(committed, sheet_index, before_count - 1, "committed")?;

    let restored = committed.apply_table_lifecycle(&commit.patch().inverse())?;
    if exact_bytes(restored.package())? != original {
        return Err(invalid(
            "inverse patch did not restore the exact input package",
        ));
    }

    committed.save(&output)?;
    let reopened = Package::open(&output)?;
    assert_count(&reopened, sheet_index, before_count - 1, "reopened")?;

    println!(
        "removed workbook table {} from sheet {}: {} table(s) remain",
        table_index,
        sheet_index,
        before_count - 1,
    );
    Ok(())
}

fn locate_table(package: &Package, global_index: usize) -> Result<(usize, usize), Box<dyn Error>> {
    let mut offset = 0usize;
    for (sheet_index, sheet) in package.sheets().iter().enumerate() {
        let count = sheet.tables().count();
        if global_index < offset + count {
            return Ok((sheet_index, global_index - offset));
        }
        offset += count;
    }
    Err(invalid("table index was not found"))
}

fn assert_count(
    package: &Package,
    sheet_index: usize,
    expected: usize,
    stage: &str,
) -> Result<(), Box<dyn Error>> {
    let count = package
        .sheets()
        .get(sheet_index)
        .ok_or_else(|| invalid(format!("{stage} package lost the selected sheet")))?
        .tables()
        .count();
    if count != expected {
        return Err(invalid(format!(
            "{stage} package has {count} table(s), expected {expected}"
        )));
    }
    Ok(())
}

fn exact_bytes(package: &Package) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn invalid(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(io::Error::new(io::ErrorKind::InvalidInput, message.into()))
}
