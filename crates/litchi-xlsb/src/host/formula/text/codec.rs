#![allow(
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    clippy::map_err_ignore,
    reason = "legacy module confines validated BIFF12 field narrowing or exact signed-bit reinterpretation, normalization into the module's stable typed public error to this codec boundary"
)]

//! XLSB formula-text token codec.

use super::super::{ArrayValue, BinaryOperator, Error, Result, UnaryOperator, ptg_types};
use super::ast::CompileExpr;
use super::compiler::Compiler;
use super::model::FormulaEncoding;
use super::references::{
    emit_reference, emit_shared_reference, encode_shared_reference, reference_column_bits,
};
use super::validation::validate_xnum;

impl<'a> Compiler<'a> {
    pub(super) fn emit(
        expression: &CompileExpr,
        output: &mut Vec<u8>,
        extra: &mut Vec<u8>,
        encoding: FormulaEncoding,
    ) -> Result<()> {
        match expression {
            CompileExpr::Number(value) => {
                validate_xnum(*value, "compiled number")?;
                if value.fract() == 0.0 && *value >= 0.0 && *value <= f64::from(u16::MAX) {
                    output.push(ptg_types::PTG_INT);
                    output.extend_from_slice(&(*value as u16).to_le_bytes());
                } else {
                    output.push(ptg_types::PTG_NUM);
                    output.extend_from_slice(&value.to_le_bytes());
                }
            },
            CompileExpr::String(value) => {
                let utf16: Vec<u16> = value.encode_utf16().collect();
                output.push(ptg_types::PTG_STR);
                output.extend_from_slice(&(utf16.len() as u16).to_le_bytes());
                for unit in utf16 {
                    output.extend_from_slice(&unit.to_le_bytes());
                }
            },
            CompileExpr::Bool(value) => {
                output.push(ptg_types::PTG_BOOL);
                output.push(u8::from(*value));
            },
            CompileExpr::Error(error) => {
                output.push(ptg_types::PTG_ERR);
                output.push(*error);
            },
            CompileExpr::MissingArg => output.push(ptg_types::PTG_MISSING_ARG),
            CompileExpr::Parenthesized(expression) => {
                Self::emit(expression, output, extra, encoding)?;
                output.push(ptg_types::PTG_PAREN);
            },
            CompileExpr::Array { rows, cols, values } => {
                if matches!(encoding, FormulaEncoding::Shared { .. }) {
                    return Err(Error::InvalidFormula(
                        "shared formulas cannot contain PtgArray".to_string(),
                    ));
                }
                output.push(0x40); // PtgArray, VALUE class
                output.extend_from_slice(&[0; 14]);
                extra.extend_from_slice(&rows.to_le_bytes());
                extra.extend_from_slice(&cols.to_le_bytes());
                for value in values {
                    match value {
                        ArrayValue::Number(value) => {
                            extra.push(0x00);
                            extra.extend_from_slice(&value.to_le_bytes());
                        },
                        ArrayValue::String(value) => {
                            let utf16: Vec<u16> = value.encode_utf16().collect();
                            extra.push(0x01);
                            extra.extend_from_slice(&(utf16.len() as u16).to_le_bytes());
                            for unit in utf16 {
                                extra.extend_from_slice(&unit.to_le_bytes());
                            }
                        },
                        ArrayValue::Bool(value) => {
                            extra.extend_from_slice(&[0x02, u8::from(*value)]);
                        },
                        ArrayValue::Error(error) => {
                            extra.extend_from_slice(&[0x04, *error, 0, 0, 0]);
                        },
                    }
                }
            },
            CompileExpr::Ref(reference) => match encoding {
                FormulaEncoding::Cell => emit_reference(output, 0x44, *reference),
                FormulaEncoding::Shared { base_row, base_col } => {
                    emit_shared_reference(output, 0x4C, *reference, base_row, base_col)?
                },
            },
            CompileExpr::Area(first, last) => {
                match encoding {
                    FormulaEncoding::Cell => {
                        output.push(0x25); // PtgArea, REFERENCE class
                        output.extend_from_slice(&first.row.to_le_bytes());
                        output.extend_from_slice(&last.row.to_le_bytes());
                        output.extend_from_slice(&reference_column_bits(*first).to_le_bytes());
                        output.extend_from_slice(&reference_column_bits(*last).to_le_bytes());
                    },
                    FormulaEncoding::Shared { base_row, base_col } => {
                        output.push(0x2D); // PtgAreaN, REFERENCE class
                        let (first_row, first_col) =
                            encode_shared_reference(*first, base_row, base_col)?;
                        let (last_row, last_col) =
                            encode_shared_reference(*last, base_row, base_col)?;
                        output.extend_from_slice(&first_row.to_le_bytes());
                        output.extend_from_slice(&last_row.to_le_bytes());
                        output.extend_from_slice(&first_col.to_le_bytes());
                        output.extend_from_slice(&last_col.to_le_bytes());
                    },
                }
            },
            CompileExpr::Ref3d(sheet_index, reference) => {
                output.push(0x5A); // PtgRef3d, VALUE class
                output.extend_from_slice(&sheet_index.to_le_bytes());
                output.extend_from_slice(&reference.row.to_le_bytes());
                output.extend_from_slice(&reference_column_bits(*reference).to_le_bytes());
            },
            CompileExpr::Area3d(sheet_index, first, last) => {
                output.push(0x5B); // PtgArea3d, VALUE class
                output.extend_from_slice(&sheet_index.to_le_bytes());
                output.extend_from_slice(&first.row.to_le_bytes());
                output.extend_from_slice(&last.row.to_le_bytes());
                output.extend_from_slice(&reference_column_bits(*first).to_le_bytes());
                output.extend_from_slice(&reference_column_bits(*last).to_le_bytes());
            },
            CompileExpr::Name(index) => {
                output.push(0x43); // PtgName, VALUE class
                output.extend_from_slice(&index.to_le_bytes());
            },
            CompileExpr::TableReference(reference) => {
                let (token, payload) = reference.to_extended_binary()?;
                output.extend_from_slice(&token);
                extra.extend_from_slice(&payload);
            },
            CompileExpr::Unary(operator, operand) => {
                Self::emit(operand, output, extra, encoding)?;
                output.push(match operator {
                    UnaryOperator::Plus => ptg_types::PTG_UPLUS,
                    UnaryOperator::Minus => ptg_types::PTG_UMINUS,
                    UnaryOperator::Percent => ptg_types::PTG_PERCENT,
                });
            },
            CompileExpr::Binary(operator, left, right) => {
                Self::emit(left, output, extra, encoding)?;
                Self::emit(right, output, extra, encoding)?;
                output.push(match operator {
                    BinaryOperator::Add => ptg_types::PTG_ADD,
                    BinaryOperator::Subtract => ptg_types::PTG_SUB,
                    BinaryOperator::Multiply => ptg_types::PTG_MUL,
                    BinaryOperator::Divide => ptg_types::PTG_DIV,
                    BinaryOperator::Power => ptg_types::PTG_POWER,
                    BinaryOperator::Concat => ptg_types::PTG_CONCAT,
                    BinaryOperator::LessThan => ptg_types::PTG_LT,
                    BinaryOperator::LessEqual => ptg_types::PTG_LE,
                    BinaryOperator::Equal => ptg_types::PTG_EQ,
                    BinaryOperator::GreaterEqual => ptg_types::PTG_GE,
                    BinaryOperator::GreaterThan => ptg_types::PTG_GT,
                    BinaryOperator::NotEqual => ptg_types::PTG_NE,
                    BinaryOperator::Intersection => ptg_types::PTG_ISECT,
                    BinaryOperator::Union => ptg_types::PTG_UNION,
                    BinaryOperator::Range => ptg_types::PTG_RANGE,
                });
            },
            CompileExpr::Function(function, arguments) => {
                if function.index == 1 {
                    return Self::emit_if(arguments, output, extra, encoding);
                }
                if function.index == 100 {
                    return Self::emit_choose(arguments, output, extra, encoding);
                }
                if function.index == 480 {
                    return Self::emit_iferror(arguments, output, extra, encoding);
                }
                for argument in arguments {
                    Self::emit(argument, output, extra, encoding)?;
                }
                if function.min_args == function.max_args {
                    output.push(0x41); // PtgFunc, VALUE class
                    output.extend_from_slice(&function.index.to_le_bytes());
                } else {
                    output.push(0x42); // PtgFuncVar, VALUE class
                    output.push(arguments.len() as u8);
                    output.extend_from_slice(&function.index.to_le_bytes());
                }
            },
        }
        Ok(())
    }

    fn emit_if(
        arguments: &[CompileExpr],
        output: &mut Vec<u8>,
        extra: &mut Vec<u8>,
        encoding: FormulaEncoding,
    ) -> Result<()> {
        debug_assert!(matches!(arguments.len(), 2 | 3));
        Self::emit(&arguments[0], output, extra, encoding)?;
        let attr_if = append_attribute(output, 0x02, 0);
        Self::emit(&arguments[1], output, extra, encoding)?;
        let goto_true = append_attribute(output, 0x08, 0);
        let goto_false = if arguments.len() == 3 {
            Self::emit(&arguments[2], output, extra, encoding)?;
            Some(append_attribute(output, 0x08, 0))
        } else {
            None
        };
        output.extend_from_slice(&[0x42, arguments.len() as u8, 0x01, 0x00]);

        patch_attribute_offset(output, attr_if, goto_true + 4 - (attr_if + 4))?;
        patch_skip_to_end(output, goto_true)?;
        if let Some(position) = goto_false {
            patch_skip_to_end(output, position)?;
        }
        Ok(())
    }

    fn emit_iferror(
        arguments: &[CompileExpr],
        output: &mut Vec<u8>,
        extra: &mut Vec<u8>,
        encoding: FormulaEncoding,
    ) -> Result<()> {
        debug_assert_eq!(arguments.len(), 2);
        Self::emit(&arguments[0], output, extra, encoding)?;
        let attr_if_error = append_attribute(output, 0x80, 0);
        Self::emit(&arguments[1], output, extra, encoding)?;
        let goto = append_attribute(output, 0x08, 0);
        output.extend_from_slice(&[0x41, 0xE0, 0x01]);

        patch_attribute_offset(output, attr_if_error, goto + 4 - (attr_if_error + 4))?;
        patch_skip_to_end(output, goto)?;
        Ok(())
    }

    fn emit_choose(
        arguments: &[CompileExpr],
        output: &mut Vec<u8>,
        extra: &mut Vec<u8>,
        encoding: FormulaEncoding,
    ) -> Result<()> {
        debug_assert!((2..=255).contains(&arguments.len()));
        Self::emit(&arguments[0], output, extra, encoding)?;
        let choice_count = arguments.len() - 1;
        let attr_choose = output.len();
        output.extend_from_slice(&[ptg_types::PTG_ATTR, 0x04]);
        output.extend_from_slice(&(choice_count as u16).to_le_bytes());
        output.resize(output.len() + (choice_count + 1) * 2, 0);
        let attr_size = output.len() - attr_choose;
        patch_u16(
            output,
            attr_choose + 4,
            attr_size - 4,
            "PtgAttrChoose first offset",
        )?;

        let mut gotos = Vec::with_capacity(choice_count);
        for (index, argument) in arguments[1..].iter().enumerate() {
            Self::emit(argument, output, extra, encoding)?;
            gotos.push(append_attribute(output, 0x08, 0));
            let cumulative = output.len() - (attr_choose + attr_size);
            patch_u16(
                output,
                attr_choose + 6 + index * 2,
                cumulative,
                "PtgAttrChoose branch offset",
            )?;
        }
        output.extend_from_slice(&[0x42, arguments.len() as u8, 0x64, 0x00]);
        for goto in gotos {
            patch_skip_to_end(output, goto)?;
        }
        Ok(())
    }
}

fn append_attribute(output: &mut Vec<u8>, selector: u8, offset: u16) -> usize {
    let position = output.len();
    output.extend_from_slice(&[ptg_types::PTG_ATTR, selector]);
    output.extend_from_slice(&offset.to_le_bytes());
    position
}

fn patch_attribute_offset(output: &mut [u8], position: usize, offset: usize) -> Result<()> {
    patch_u16(output, position + 2, offset, "PtgAttr offset")
}

fn patch_skip_to_end(output: &mut [u8], position: usize) -> Result<()> {
    let remaining = output
        .len()
        .checked_sub(position + 4)
        .ok_or_else(|| Error::InvalidFormula("PtgAttrGoTo position exceeds formula".to_string()))?;
    let offset = remaining
        .checked_sub(1)
        .ok_or_else(|| Error::InvalidFormula("PtgAttrGoTo has no following token".to_string()))?;
    patch_attribute_offset(output, position, offset)
}

fn patch_u16(output: &mut [u8], position: usize, value: usize, context: &str) -> Result<()> {
    let value = u16::try_from(value)
        .map_err(|_| Error::InvalidFormula(format!("{context} exceeds 65,535 bytes")))?;
    let target = output
        .get_mut(position..position + 2)
        .ok_or_else(|| Error::InvalidFormula(format!("{context} position is outside formula")))?;
    target.copy_from_slice(&value.to_le_bytes());
    Ok(())
}
