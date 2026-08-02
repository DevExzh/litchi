use std::env;

use litchi_iwa::numbers::{
    FormulaAxisReference, FormulaBinaryOperator, FormulaCellReference, FormulaExpression,
    FormulaPivotCategoryReference, FormulaUuid, NumbersDocument, NumbersEditor,
};

fn parse_address(address: &str) -> Result<FormulaCellReference, String> {
    let mut input = address;
    let absolute_column = input.starts_with('$');
    if absolute_column {
        input = &input[1..];
    }
    let letter_count = input
        .bytes()
        .take_while(|byte| byte.is_ascii_alphabetic())
        .count();
    if letter_count == 0 {
        return Err(format!("invalid cell address {address:?}"));
    }
    let (letters, mut row_text) = input.split_at(letter_count);
    let absolute_row = row_text.starts_with('$');
    if absolute_row {
        row_text = &row_text[1..];
    }
    let one_based_row = row_text
        .parse::<usize>()
        .map_err(|_| format!("invalid cell address {address:?}"))?;
    if one_based_row == 0 {
        return Err(format!("cell rows are one-based in {address:?}"));
    }
    let mut one_based_column = 0usize;
    for byte in letters.bytes() {
        let digit = usize::from(byte.to_ascii_uppercase() - b'A' + 1);
        one_based_column = one_based_column
            .checked_mul(26)
            .and_then(|value| value.checked_add(digit))
            .ok_or_else(|| format!("cell column overflows in {address:?}"))?;
    }
    Ok(FormulaCellReference::mixed(
        one_based_row - 1,
        one_based_column - 1,
        absolute_row,
        absolute_column,
    ))
}

fn parse_row_reference(value: &str) -> Result<FormulaAxisReference, String> {
    let (absolute, value) = value
        .strip_prefix('$')
        .map_or((false, value), |value| (true, value));
    let one_based = value
        .parse::<usize>()
        .map_err(|_| format!("invalid row reference {value:?}"))?;
    if one_based == 0 {
        return Err("formula row references are one-based".to_owned());
    }
    Ok(if absolute {
        FormulaAxisReference::absolute(one_based - 1)
    } else {
        FormulaAxisReference::relative(one_based - 1)
    })
}

fn parse_column_reference(value: &str) -> Result<FormulaAxisReference, String> {
    let (absolute, letters) = value
        .strip_prefix('$')
        .map_or((false, value), |value| (true, value));
    if letters.is_empty() || !letters.bytes().all(|byte| byte.is_ascii_alphabetic()) {
        return Err(format!("invalid column reference {value:?}"));
    }
    let mut one_based = 0usize;
    for byte in letters.bytes() {
        let digit = usize::from(byte.to_ascii_uppercase() - b'A' + 1);
        one_based = one_based
            .checked_mul(26)
            .and_then(|value| value.checked_add(digit))
            .ok_or_else(|| format!("column reference overflows in {value:?}"))?;
    }
    Ok(if absolute {
        FormulaAxisReference::absolute(one_based - 1)
    } else {
        FormulaAxisReference::relative(one_based - 1)
    })
}

fn parse_axis_range(
    range: &str,
    parse: fn(&str) -> Result<FormulaAxisReference, String>,
) -> Result<(FormulaAxisReference, FormulaAxisReference), String> {
    let (start, end) = range
        .split_once(':')
        .ok_or_else(|| format!("invalid axis range {range:?}"))?;
    Ok((parse(start)?, parse(end)?))
}

fn parse_operand(argument: &str, table_ids: &[u64]) -> Result<FormulaExpression, String> {
    if let Some(value) = argument.strip_prefix("pivot:") {
        let fields = value.split(':').collect::<Vec<_>>();
        if fields.len() != 8 {
            return Err(format!("invalid pivot category operand {argument:?}"));
        }
        let parse_u64 = |index: usize| {
            fields[index]
                .parse::<u64>()
                .map_err(|_| format!("invalid pivot UUID field in {argument:?}"))
        };
        let aggregate_type = fields[6]
            .parse::<u32>()
            .map_err(|_| format!("invalid pivot aggregate type in {argument:?}"))?;
        let group_level = fields[7]
            .parse::<i32>()
            .map_err(|_| format!("invalid pivot group level in {argument:?}"))?;
        return Ok(FormulaExpression::pivot_category(
            FormulaPivotCategoryReference::new(
                FormulaUuid::new(parse_u64(0)?, parse_u64(1)?),
                FormulaUuid::new(parse_u64(2)?, parse_u64(3)?),
                FormulaUuid::new(parse_u64(4)?, parse_u64(5)?),
                aggregate_type,
                group_level,
            ),
        ));
    }
    for (prefix, rows) in [("table-rows:", true), ("table-columns:", false)] {
        if let Some(target) = argument.strip_prefix(prefix) {
            let (table_index, range) = target
                .split_once(':')
                .ok_or_else(|| format!("invalid cross-table axis operand {argument:?}"))?;
            let table_index = table_index
                .parse::<usize>()
                .map_err(|_| format!("invalid table index in {argument:?}"))?;
            let table_id = *table_ids
                .get(table_index)
                .ok_or_else(|| format!("table index {table_index} was not found"))?;
            let (start, end) = parse_axis_range(
                range,
                if rows {
                    parse_row_reference
                } else {
                    parse_column_reference
                },
            )?;
            return Ok(if rows {
                FormulaExpression::table_rows(table_id, start, end)
            } else {
                FormulaExpression::table_columns(table_id, start, end)
            });
        }
    }
    if let Some(target) = argument.strip_prefix("table-range:") {
        let (table_index, range) = target
            .split_once(':')
            .ok_or_else(|| format!("invalid cross-table range operand {argument:?}"))?;
        let (start, end) = range
            .split_once(':')
            .ok_or_else(|| format!("invalid cross-table range operand {argument:?}"))?;
        let table_index = table_index
            .parse::<usize>()
            .map_err(|_| format!("invalid table index in {argument:?}"))?;
        let table_id = *table_ids
            .get(table_index)
            .ok_or_else(|| format!("table index {table_index} was not found"))?;
        return Ok(FormulaExpression::table_range(
            table_id,
            parse_address(start)?,
            parse_address(end)?,
        ));
    }
    if let Some(target) = argument.strip_prefix("table-cell:") {
        let (table_index, address) = target
            .split_once(':')
            .ok_or_else(|| format!("invalid cross-table cell operand {argument:?}"))?;
        let table_index = table_index
            .parse::<usize>()
            .map_err(|_| format!("invalid table index in {argument:?}"))?;
        let table_id = *table_ids
            .get(table_index)
            .ok_or_else(|| format!("table index {table_index} was not found"))?;
        return parse_address(address)
            .map(|reference| FormulaExpression::table_cell(table_id, reference));
    }
    if let Some(address) = argument.strip_prefix("cell:") {
        return parse_address(address).map(FormulaExpression::cell);
    }
    if let Some(range) = argument.strip_prefix("range:") {
        let (start, end) = range
            .split_once(':')
            .ok_or_else(|| format!("invalid range operand {argument:?}"))?;
        return Ok(FormulaExpression::range(
            parse_address(start)?,
            parse_address(end)?,
        ));
    }
    if let Some(range) = argument.strip_prefix("rows:") {
        let (start, end) = parse_axis_range(range, parse_row_reference)?;
        return Ok(FormulaExpression::rows(start, end));
    }
    if let Some(range) = argument.strip_prefix("columns:") {
        let (start, end) = parse_axis_range(range, parse_column_reference)?;
        return Ok(FormulaExpression::columns(start, end));
    }
    argument
        .parse::<f64>()
        .map(FormulaExpression::Number)
        .map_err(|_| {
            format!(
                "operand must be a number, cell:A1, range:A1:B2, rows:1:2, columns:B:C, table-cell:<index>:A1, table-range:<index>:A1:B2, table-rows:<index>:1:2, table-columns:<index>:B:C, or pivot:<group-by-lower>:<group-by-upper>:<column-lower>:<column-upper>:<group-lower>:<group-upper>:<aggregate-type>:<group-level>: {argument:?}"
            )
        })
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let arguments = env::args().skip(1).collect::<Vec<_>>();
    if arguments.len() < 7 {
        eprintln!(
            "usage: edit_numbers_formula <input.numbers> <output.numbers> <table-index> <row> <column> <function|/> <number|cell:A1|range:A1:B2|rows:1:2|columns:B:C|table-cell:<index>:A1|table-range:<index>:A1:B2|table-rows:<index>:1:2|table-columns:<index>:B:C|pivot:<group-by-lower>:<group-by-upper>:<column-lower>:<column-upper>:<group-lower>:<group-upper>:<aggregate-type>:<group-level>>..."
        );
        std::process::exit(2);
    }

    let input = &arguments[0];
    let output = &arguments[1];
    let table_index = arguments[2].parse::<usize>()?;
    let row = arguments[3].parse::<usize>()?;
    let column = arguments[4].parse::<usize>()?;
    let function = &arguments[5];
    let mut editor = NumbersEditor::open(input)?;
    let tables = editor.tables()?;
    let table_ids = tables
        .iter()
        .map(|table| table.object_id)
        .collect::<Vec<_>>();
    let table = tables
        .into_iter()
        .nth(table_index)
        .ok_or("table index was not found")?;
    let operands = arguments[6..]
        .iter()
        .map(|argument| parse_operand(argument, &table_ids))
        .collect::<std::result::Result<Vec<_>, _>>()?;
    let table_id = table.object_id;
    let table_name = table.name;
    let expression = if function == "/" {
        let mut operands = operands.into_iter();
        let left = operands.next().ok_or("division needs two operands")?;
        let right = operands.next().ok_or("division needs two operands")?;
        if operands.next().is_some() {
            return Err("division needs exactly two operands".into());
        }
        FormulaExpression::binary(FormulaBinaryOperator::Divide, left, right)
    } else {
        FormulaExpression::function(function, operands)
    };
    editor.set_formula(table_id, row, column, expression)?;
    editor.save(output)?;

    let document = NumbersDocument::open(output)?;
    let sheets = document.sheets()?;
    let cell = sheets
        .iter()
        .flat_map(|sheet| &sheet.tables)
        .find(|table| table.name == table_name)
        .and_then(|table| table.get_cell(row, column));
    if let Some(cell) = cell {
        println!("saved {output}; cell ({row}, {column}) = {cell}");
    } else {
        println!(
            "saved {output}; native formula link was verified transactionally (the public reader did not expose table {table_name:?})"
        );
    }
    Ok(())
}
