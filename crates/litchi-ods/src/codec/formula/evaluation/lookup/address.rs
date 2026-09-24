//! Bounded ADDRESS formatting shared by scalar and value lookup paths.

use super::super::{EvaluationFailure, EvaluationResult, Evaluator, TextValue};
use litchi_core::Resource;
use std::fmt::Write as _;

/// Format one ADDRESS result. `row` and `column` are one-based and `abs` is
/// the validated OpenFormula mode 1 through 4. The formatter is pure with
/// respect to workbook state: R1C1 relative notation uses the supplied row
/// and column as the bracketed values and does not consult a current position.
pub(crate) fn format<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    row: usize,
    column: usize,
    abs: u8,
    a1: bool,
    sheet: Option<&str>,
) -> EvaluationResult<TextValue<'a>> {
    if row == 0 || column == 0 || !(1..=4).contains(&abs) {
        return Err(EvaluationFailure::InvalidExpression(
            "ADDRESS received an unvalidated coordinate or absolute mode",
        ));
    }

    let sheet = sheet.filter(|value| !value.is_empty());
    if let Some(sheet) = sheet {
        // The sizing pass below examines the complete sheet name, so charge
        // and fence that scan before inspecting its quoting form.
        evaluator.charge_bytes(sheet.len())?;
    }
    let (sheet_len, sheet_quoted) = match sheet {
        Some(sheet) => sheet_output_len(sheet, evaluator)?,
        None => (0, false),
    };
    let separator_len = usize::from(sheet.is_some());
    let coordinate_len = if a1 {
        a1_len(row, column, abs)?
    } else {
        r1c1_len(row, column, abs)?
    };
    let total = sheet_len
        .checked_add(separator_len)
        .and_then(|value| value.checked_add(coordinate_len))
        .ok_or(EvaluationFailure::InvalidExpression(
            "ADDRESS output length overflows",
        ))?;
    evaluator.charge_bytes(total)?;
    if total > evaluator.limits.max_text_bytes() {
        return Err(super::super::local_limit(
            Resource::Memory,
            u64::try_from(total).unwrap_or(u64::MAX),
            u64::try_from(evaluator.limits.max_text_bytes()).unwrap_or(u64::MAX),
        ));
    }

    // Declare the lease before the output so a failure drops the String before
    // its accounting token, matching the evaluator's storage drop invariant.
    let reservation = evaluator.reserve_storage(total, "formula ADDRESS output")?;
    let mut output = String::new();
    if total != 0 {
        output
            .try_reserve_exact(total)
            .map_err(|source| EvaluationFailure::Allocation {
                resource: "formula ADDRESS output",
                source,
            })?;
    }
    if let Some(sheet) = sheet {
        evaluator.charge_work(0)?;
        write_sheet(&mut output, sheet, sheet_quoted, evaluator)?;
        evaluator.charge_work(0)?;
        output.push(if a1 { '.' } else { '!' });
    }
    if a1 {
        write_a1(&mut output, row, column, abs)?;
    } else {
        write_r1c1(&mut output, row, column, abs)?;
    }
    if output.len() != total {
        return Err(EvaluationFailure::InvalidExpression(
            "ADDRESS output length disagrees with its sizing pass",
        ));
    }
    Ok(TextValue::owned(output, reservation))
}

fn a1_len(row: usize, column: usize, abs: u8) -> EvaluationResult<usize> {
    let column_len = column_label_len(column)?;
    let row_len = decimal_len(row);
    let column_marker = usize::from(abs == 1 || abs == 3);
    let row_marker = usize::from(abs == 1 || abs == 2);
    column_len
        .checked_add(row_len)
        .and_then(|value| value.checked_add(column_marker))
        .and_then(|value| value.checked_add(row_marker))
        .ok_or(EvaluationFailure::InvalidExpression(
            "ADDRESS A1 length overflows",
        ))
}

fn r1c1_len(row: usize, column: usize, abs: u8) -> EvaluationResult<usize> {
    let row_len = decimal_len(row);
    let column_len = decimal_len(column);
    let result = match abs {
        1 => 1usize
            .checked_add(row_len)
            .and_then(|value| value.checked_add(1))
            .and_then(|value| value.checked_add(column_len)),
        2 => 1usize
            .checked_add(row_len)
            .and_then(|value| value.checked_add(3))
            .and_then(|value| value.checked_add(column_len)),
        3 => 2usize
            .checked_add(row_len)
            .and_then(|value| value.checked_add(2))
            .and_then(|value| value.checked_add(column_len)),
        4 => 2usize
            .checked_add(row_len)
            .and_then(|value| value.checked_add(4))
            .and_then(|value| value.checked_add(column_len)),
        _ => None,
    };
    result.ok_or(EvaluationFailure::InvalidExpression(
        "ADDRESS R1C1 length overflows",
    ))
}

fn decimal_len(mut value: usize) -> usize {
    let mut length = 1;
    while value >= 10 {
        value /= 10;
        length += 1;
    }
    length
}

fn column_label_len(mut column: usize) -> EvaluationResult<usize> {
    if column == 0 {
        return Err(EvaluationFailure::InvalidExpression(
            "ADDRESS column is zero",
        ));
    }
    let mut length = 0usize;
    while column != 0 {
        column = (column - 1) / 26;
        length = length
            .checked_add(1)
            .ok_or(EvaluationFailure::InvalidExpression(
                "ADDRESS column label length overflows",
            ))?;
    }
    Ok(length)
}

fn sheet_output_len(
    sheet: &str,
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<(usize, bool)> {
    let mut quoted = false;
    let mut apostrophes = 0usize;
    let mut polls_remaining = 0usize;
    for character in sheet.chars() {
        if polls_remaining == 0 {
            evaluator.charge_work(0)?;
            polls_remaining = 4096;
        }
        polls_remaining -= 1;
        if !character.is_ascii_alphanumeric() && character != '_' {
            quoted = true;
        }
        if character == '\'' {
            apostrophes =
                apostrophes
                    .checked_add(1)
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "ADDRESS apostrophe count overflows",
                    ))?;
        }
    }
    evaluator.charge_work(0)?;
    if quoted {
        let length = sheet
            .len()
            .checked_add(apostrophes)
            .and_then(|value| value.checked_add(2))
            .ok_or(EvaluationFailure::InvalidExpression(
                "ADDRESS sheet output length overflows",
            ))?;
        Ok((length, true))
    } else {
        Ok((sheet.len(), false))
    }
}

fn write_sheet(
    output: &mut String,
    sheet: &str,
    quoted: bool,
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<()> {
    if quoted {
        output.push('\'');
        let mut polls_remaining = 0usize;
        for character in sheet.chars() {
            if polls_remaining == 0 {
                evaluator.charge_work(0)?;
                polls_remaining = 4096;
            }
            polls_remaining -= 1;
            output.push(character);
            if character == '\'' {
                output.push('\'');
            }
        }
        output.push('\'');
    } else {
        evaluator.charge_work(0)?;
        output.push_str(sheet);
    }
    evaluator.charge_work(0)?;
    Ok(())
}

fn write_a1(output: &mut String, row: usize, column: usize, abs: u8) -> EvaluationResult<()> {
    if abs == 1 || abs == 3 {
        output.push('$');
    }
    write_column(output, column)?;
    if abs == 1 || abs == 2 {
        output.push('$');
    }
    write!(output, "{row}")
        .map_err(|_| EvaluationFailure::InvalidExpression("ADDRESS row formatting failed"))
}

fn write_r1c1(output: &mut String, row: usize, column: usize, abs: u8) -> EvaluationResult<()> {
    match abs {
        1 => {
            write!(output, "R{row}C{column}")
        },
        2 => {
            write!(output, "R{row}C[{column}]")
        },
        3 => {
            write!(output, "R[{row}]C{column}")
        },
        4 => {
            write!(output, "R[{row}]C[{column}]")
        },
        _ => {
            return Err(EvaluationFailure::InvalidExpression(
                "ADDRESS absolute mode is outside 1..=4",
            ));
        },
    }
    .map_err(|_| EvaluationFailure::InvalidExpression("ADDRESS R1C1 formatting failed"))
}

fn write_column(output: &mut String, mut column: usize) -> EvaluationResult<()> {
    if column == 0 {
        return Err(EvaluationFailure::InvalidExpression(
            "ADDRESS column is zero",
        ));
    }
    let mut letters = [0_u8; 32];
    let mut length = 0usize;
    while column != 0 {
        if length == letters.len() {
            return Err(EvaluationFailure::InvalidExpression(
                "ADDRESS column label exceeds fixed width",
            ));
        }
        column -= 1;
        letters[length] = b'A'
            + u8::try_from(column % 26).map_err(|_| {
                EvaluationFailure::InvalidExpression("ADDRESS column label digit overflows")
            })?;
        length += 1;
        column /= 26;
    }
    for letter in letters[..length].iter().rev() {
        output.push(char::from(*letter));
    }
    Ok(())
}
