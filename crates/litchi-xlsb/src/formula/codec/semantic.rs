//! Formula token-to-text compiler and host-resolution seam.
use super::super::model::column_index_to_name;
use super::super::model::*;
use super::super::{Error, Result};
use super::validation::builtin_function_by_index;

/// Context-dependent name and relationship resolution supplied by a host.
///
/// The owner codec never discovers package parts or workbook relationships;
/// it only asks the host for already-resolved formula names and prefixes.
pub trait Resolution {
    fn sheet_prefix(&self, index: u16) -> Result<String>;
    fn defined_name(&self, index: u32) -> Result<String>;
    fn external_name(&self, sheet_index: u16, name_index: u32) -> Result<String>;
    fn table_reference(&self, reference: &TableReference) -> Result<String>;
    fn pivot_name(&self, index: u32) -> Result<String>;
}

pub struct Compiler;

impl Compiler {
    /// Convert formula tokens to string representation
    ///
    /// Uses RPN to infix conversion with proper operator precedence.
    pub fn tokens_to_string(tokens: &[Token]) -> String {
        Self::try_tokens_to_string(tokens).unwrap_or_default()
    }

    /// Convert tokens to text, rejecting token streams that cannot be
    /// represented faithfully by this converter.
    pub fn try_tokens_to_string(tokens: &[Token]) -> Result<String> {
        Self::try_tokens_to_string_with_optional_context(tokens, None)
    }

    /// Convert formula tokens using workbook extern-sheet and name metadata.
    pub fn try_tokens_to_string_with_resolution(
        tokens: &[Token],
        context: &dyn Resolution,
    ) -> Result<String> {
        Self::try_tokens_to_string_with_optional_context(tokens, Some(context))
    }

    fn try_tokens_to_string_with_optional_context(
        tokens: &[Token],
        context: Option<&dyn Resolution>,
    ) -> Result<String> {
        let mut stack: Vec<String> = Vec::new();

        for token in tokens {
            match token {
                Token::Number(n) => stack.push(format!("{}", n)),
                Token::Int(i) => stack.push(format!("{}", i)),
                Token::MissingArg => stack.push(String::new()),
                Token::Parenthesis => {
                    let Some(expression) = stack.pop() else {
                        return Err(Error::InvalidFormula(
                            "PtgParen has no preceding expression".to_string(),
                        ));
                    };
                    stack.push(format!("({expression})"));
                },
                Token::Attribute(_) => {},
                Token::Array { rows, cols, values } => {
                    let expected = usize::try_from(u64::from(*rows) * u64::from(*cols))
                        .map_err(|_| Error::InvalidFormula("array is too large".to_string()))?;
                    if values.len() != expected {
                        return Err(Error::InvalidFormula(format!(
                            "array dimensions require {expected} values, found {}",
                            values.len()
                        )));
                    }
                    let mut text = String::from("{");
                    for row in 0..*rows {
                        if row != 0 {
                            text.push(';');
                        }
                        for col in 0..*cols {
                            if col != 0 {
                                text.push(',');
                            }
                            let index =
                                usize::try_from(u64::from(row) * u64::from(*cols) + u64::from(col))
                                    .map_err(|_| {
                                        Error::InvalidFormula("array index overflow".to_string())
                                    })?;
                            match &values[index] {
                                ArrayValue::Number(value) => {
                                    text.push_str(&value.to_string());
                                },
                                ArrayValue::String(value) => {
                                    text.push('"');
                                    text.push_str(&value.replace('"', "\"\""));
                                    text.push('"');
                                },
                                ArrayValue::Bool(value) => {
                                    text.push_str(if *value { "TRUE" } else { "FALSE" });
                                },
                                ArrayValue::Error(error) => {
                                    text.push_str(&Self::error_to_string(*error));
                                },
                            }
                        }
                    }
                    text.push('}');
                    stack.push(text);
                },
                Token::Memory { .. } => {},
                Token::String(s) => stack.push(format!("\"{}\"", s.replace('"', "\"\""))),
                Token::Bool(b) => stack.push(if *b {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }),
                Token::Error(e) => stack.push(Self::error_to_string(*e)),
                Token::CellRef {
                    row,
                    col,
                    row_relative,
                    col_relative,
                } => {
                    let col_str = column_index_to_name(*col + 1);
                    let row_str = row + 1;
                    let col_prefix = if *col_relative { "" } else { "$" };
                    let row_prefix = if *row_relative { "" } else { "$" };
                    stack.push(format!(
                        "{}{}{}{}",
                        col_prefix, col_str, row_prefix, row_str
                    ));
                },
                Token::AreaRef {
                    row_first,
                    col_first,
                    row_last,
                    col_last,
                    row_first_relative,
                    row_last_relative,
                    col_first_relative,
                    col_last_relative,
                } => {
                    let first = Self::format_reference(
                        *row_first,
                        *col_first,
                        *row_first_relative,
                        *col_first_relative,
                    );
                    let last = Self::format_reference(
                        *row_last,
                        *col_last,
                        *row_last_relative,
                        *col_last_relative,
                    );
                    stack.push(format!("{}:{}", first, last));
                },
                Token::CellRef3d {
                    sheet_index,
                    row,
                    col,
                    row_relative,
                    col_relative,
                } => {
                    let context = context.ok_or_else(|| {
                        Error::UnsupportedFeature(
                            "PtgRef3d requires workbook extern-sheet resolution".to_string(),
                        )
                    })?;
                    let prefix = context.sheet_prefix(*sheet_index)?;
                    let reference =
                        Self::format_reference(*row, *col, *row_relative, *col_relative);
                    stack.push(format!("{prefix}!{reference}"));
                },
                Token::AreaRef3d {
                    sheet_index,
                    row_first,
                    row_last,
                    col_first,
                    col_last,
                    row_first_relative,
                    row_last_relative,
                    col_first_relative,
                    col_last_relative,
                } => {
                    let context = context.ok_or_else(|| {
                        Error::UnsupportedFeature(
                            "PtgArea3d requires workbook extern-sheet resolution".to_string(),
                        )
                    })?;
                    let prefix = context.sheet_prefix(*sheet_index)?;
                    let first = Self::format_reference(
                        *row_first,
                        *col_first,
                        *row_first_relative,
                        *col_first_relative,
                    );
                    let last = Self::format_reference(
                        *row_last,
                        *col_last,
                        *row_last_relative,
                        *col_last_relative,
                    );
                    stack.push(format!("{prefix}!{first}:{last}"));
                },
                Token::ReferenceError { .. } => stack.push("#REF!".to_string()),
                Token::BinaryOp(op) => {
                    if stack.len() < 2 {
                        return Err(Error::InvalidFormula(
                            "binary operator has fewer than two operands".to_string(),
                        ));
                    }
                    let right = stack.pop().expect("length checked");
                    let left = stack.pop().expect("length checked");
                    let op_str = Self::binary_op_to_string(*op);
                    stack.push(format!("({}{}{})", left, op_str, right));
                },
                Token::UnaryOp(op) => {
                    let Some(operand) = stack.pop() else {
                        return Err(Error::InvalidFormula(
                            "unary operator has no operand".to_string(),
                        ));
                    };
                    match op {
                        UnaryOperator::Plus => stack.push(format!("+({})", operand)),
                        UnaryOperator::Minus => stack.push(format!("-({})", operand)),
                        UnaryOperator::Percent => stack.push(format!("({}%)", operand)),
                    }
                },
                Token::Function {
                    index,
                    arg_count,
                    is_command,
                } => {
                    if *is_command {
                        return Err(Error::UnsupportedFeature(format!(
                            "XLSB command function index {index}"
                        )));
                    }
                    let Some(function) = builtin_function_by_index(*index) else {
                        return Err(Error::UnsupportedFeature(format!(
                            "XLSB built-in function index {index}"
                        )));
                    };
                    let func_name = function.name;
                    if stack.len() < usize::from(*arg_count) {
                        return Err(Error::InvalidFormula(format!(
                            "function {func_name} requires {arg_count} stack operands"
                        )));
                    }
                    let mut args = Vec::new();
                    for _ in 0..*arg_count {
                        if let Some(arg) = stack.pop() {
                            args.insert(0, arg);
                        }
                    }
                    stack.push(format!("{}({})", func_name, args.join(",")));
                },
                Token::Name(idx) => {
                    let context = context.ok_or_else(|| {
                        Error::UnsupportedFeature(format!(
                            "XLSB defined name index {idx} requires workbook name resolution"
                        ))
                    })?;
                    stack.push(context.defined_name(*idx)?);
                },
                Token::ExternalName {
                    sheet_index,
                    name_index,
                } => {
                    let context = context.ok_or_else(|| {
                        Error::UnsupportedFeature(
                            "PtgNameX requires workbook external-link resolution".to_string(),
                        )
                    })?;
                    stack.push(context.external_name(*sheet_index, *name_index)?);
                },
                Token::TableReference(reference) if reference.invalid => {
                    stack.push("#REF!".to_string())
                },
                Token::TableReference(reference) => {
                    let context = context.ok_or_else(|| {
                        Error::UnsupportedFeature(format!(
                            "structured table reference on Xti {} requires table-definition resolution",
                            reference.sheet_index
                        ))
                    })?;
                    stack.push(context.table_reference(reference)?);
                },
                Token::PivotName(index) => {
                    let context = context.ok_or_else(|| {
                        Error::InvalidFormula(
                            "PtgSxName requires pivot-cache calculated-name metadata".to_string(),
                        )
                    })?;
                    stack.push(context.pivot_name(*index)?);
                },
                Token::Unknown(t) => {
                    return Err(Error::UnsupportedFeature(format!(
                        "XLSB formula token 0x{t:02X}"
                    )));
                },
            }
        }

        if stack.len() != 1 {
            return Err(Error::InvalidFormula(format!(
                "formula leaves {} values on the evaluation stack",
                stack.len()
            )));
        }
        Ok(stack.pop().expect("length checked"))
    }

    fn format_reference(row: u32, col: u32, row_relative: bool, col_relative: bool) -> String {
        let col_str = column_index_to_name(col + 1);
        format!(
            "{}{}{}{}",
            if col_relative { "" } else { "$" },
            col_str,
            if row_relative { "" } else { "$" },
            row + 1
        )
    }

    /// Convert binary operator to string
    fn binary_op_to_string(op: BinaryOperator) -> &'static str {
        match op {
            BinaryOperator::Add => "+",
            BinaryOperator::Subtract => "-",
            BinaryOperator::Multiply => "*",
            BinaryOperator::Divide => "/",
            BinaryOperator::Power => "^",
            BinaryOperator::Concat => "&",
            BinaryOperator::LessThan => "<",
            BinaryOperator::LessEqual => "<=",
            BinaryOperator::Equal => "=",
            BinaryOperator::GreaterEqual => ">=",
            BinaryOperator::GreaterThan => ">",
            BinaryOperator::NotEqual => "<>",
            BinaryOperator::Intersection => " ",
            BinaryOperator::Union => ",",
            BinaryOperator::Range => ":",
        }
    }

    /// Convert error code to string
    fn error_to_string(code: u8) -> String {
        match code {
            0x00 => "#NULL!".to_string(),
            0x07 => "#DIV/0!".to_string(),
            0x0F => "#VALUE!".to_string(),
            0x17 => "#REF!".to_string(),
            0x1D => "#NAME?".to_string(),
            0x24 => "#NUM!".to_string(),
            0x2A => "#N/A".to_string(),
            0x2B => "#GETTING_DATA".to_string(),
            _ => format!("#ERR{:02X}!", code),
        }
    }
}
