use std::env;
use std::path::Path;

use litchi_iwa::keynote::KeynoteEditor;
use litchi_iwa::numbers::{
    FormulaCachedValue, FormulaCellReference, FormulaExpression, NumbersEditor,
};
use litchi_iwa::pages::PagesEditor;
use litchi_iwa::raw::package::IWorkPackage;

fn parse_cell(address: &str) -> Result<FormulaCellReference, String> {
    let letter_count = address.bytes().take_while(u8::is_ascii_alphabetic).count();
    let (letters, row) = address.split_at(letter_count);
    if letters.is_empty() || row.is_empty() {
        return Err(format!("invalid cell address {address:?}"));
    }
    let row = row
        .parse::<usize>()
        .map_err(|_| format!("invalid cell address {address:?}"))?
        .checked_sub(1)
        .ok_or_else(|| format!("cell rows are one-based in {address:?}"))?;
    let column = letters.bytes().try_fold(0usize, |column, byte| {
        let digit = usize::from(byte.to_ascii_uppercase() - b'A' + 1);
        column
            .checked_mul(26)
            .and_then(|column| column.checked_add(digit))
            .ok_or_else(|| format!("cell column overflows in {address:?}"))
    })?;
    Ok(FormulaCellReference::relative(row, column - 1))
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let arguments = env::args().skip(1).collect::<Vec<_>>();
    if !(7..=8).contains(&arguments.len()) {
        return Err("usage: edit_iwork_table_formula <input> <output> <table-id> <row> <column> <range-start> <range-end> [keynote-slide-index]".into());
    }
    let input = &arguments[0];
    let output = &arguments[1];
    let table_id = arguments[2].parse::<u64>()?;
    let row = arguments[3].parse::<usize>()?;
    let column = arguments[4].parse::<usize>()?;
    let expression = FormulaExpression::function(
        "SUM",
        [FormulaExpression::range(
            parse_cell(&arguments[5])?,
            parse_cell(&arguments[6])?,
        )],
    );
    match Path::new(input)
        .extension()
        .and_then(|extension| extension.to_str())
    {
        Some("numbers") => {
            let mut editor = NumbersEditor::open(input)?;
            editor.set_formula(table_id, row, column, expression)?;
            editor.save(output)?;
        },
        Some("pages") => {
            let mut editor = PagesEditor::open(input)?;
            editor.set_table_formula(
                table_id,
                row,
                column,
                expression,
                FormulaCachedValue::Number(0.0),
            )?;
            editor.save(output)?;
        },
        Some("key") => {
            let slide_index = arguments
                .get(7)
                .map(|value| value.parse::<usize>())
                .transpose()?
                .unwrap_or(0);
            let mut editor = KeynoteEditor::open(input)?;
            editor.set_slide_table_formula(
                slide_index,
                table_id,
                row,
                column,
                expression,
                FormulaCachedValue::Number(0.0),
            )?;
            editor.save(output)?;
        },
        extension => return Err(format!("unsupported iWork extension {extension:?}").into()),
    }

    let package = IWorkPackage::open(output)?;
    println!(
        "saved {output} with calculation engine {:?}",
        package.calculation_engine_entry_name()?
    );
    Ok(())
}
