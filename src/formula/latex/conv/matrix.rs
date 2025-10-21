// Matrix conversion logic for LaTeX conversion
//
// This module contains specialized matrix conversion functionality.

use super::converter::LatexConverter;
use super::error::LatexError;
use super::utils::estimate_matrix_capacity;
use crate::formula::ast::{Alignment, MatrixFence, MatrixProperties};
use crate::formula::latex::matrix::matrix_fence_to_env;
use std::fmt::Write;

/// Convert matrix with optimized performance (no temporary converters)
pub fn convert_matrix_optimized_internal(
    converter: &mut LatexConverter,
    rows: &[Vec<Vec<crate::formula::ast::MathNode>>],
    fence_type: MatrixFence,
    properties: Option<&MatrixProperties>,
) -> Result<(), LatexError> {
    if rows.is_empty() {
        return Ok(());
    }

    let use_array_env = properties.as_ref().and_then(|p| p.base_alignment).is_some();
    let env = if use_array_env {
        "array"
    } else {
        matrix_fence_to_env(fence_type)
    };

    let mut estimated_capacity = estimate_matrix_capacity(rows);
    if use_array_env {
        estimated_capacity += 20;
    }
    converter.buffer.reserve(estimated_capacity);

    if use_array_env {
        if let Some(props) = properties {
            if let Some(alignment) = props.base_alignment {
                write!(converter.buffer, "\\begin{{{}}}", env)
                    .map_err(|e| LatexError::FormatError(e.to_string()))?;
                converter.buffer.push('{');
                let align_char = match alignment {
                    Alignment::Left => 'l',
                    Alignment::Center => 'c',
                    Alignment::Right => 'r',
                    _ => 'c',
                };
                if let Some(num_cols) = rows.first().map(|r| r.len()) {
                    for _ in 0..num_cols {
                        converter.buffer.push(align_char);
                    }
                }
                converter.buffer.push('}');

                match fence_type {
                    MatrixFence::Paren => converter.buffer.push_str("\\left("),
                    MatrixFence::Bracket => converter.buffer.push_str("\\left["),
                    MatrixFence::Brace => converter.buffer.push_str("\\left\\{"),
                    MatrixFence::Pipe => converter.buffer.push_str("\\left|"),
                    MatrixFence::DoublePipe => converter.buffer.push_str("\\left\\|"),
                    MatrixFence::None => {},
                }
            } else {
                let num_cols = rows.first().map(|r| r.len()).unwrap_or(1);
                write!(converter.buffer, "\\begin{{{}}}", env)
                    .map_err(|e| LatexError::FormatError(e.to_string()))?;
                converter.buffer.push('{');
                for _ in 0..num_cols {
                    converter.buffer.push('c');
                }
                converter.buffer.push('}');
            }
        }
    } else {
        write!(converter.buffer, "\\begin{{{}}}", env)
            .map_err(|e| LatexError::FormatError(e.to_string()))?;
    }

    for (i, row) in rows.iter().enumerate() {
        if i > 0 {
            converter.buffer.push_str(" \\\\ ");
        }
        for (j, cell) in row.iter().enumerate() {
            if j > 0 {
                converter.buffer.push_str(" & ");
            }
            for node in cell {
                converter.convert_node(node)?;
            }
        }
    }

    if use_array_env {
        match fence_type {
            MatrixFence::Paren => converter.buffer.push_str("\\right)"),
            MatrixFence::Bracket => converter.buffer.push_str("\\right]"),
            MatrixFence::Brace => converter.buffer.push_str("\\right\\}"),
            MatrixFence::Pipe => converter.buffer.push_str("\\right|"),
            MatrixFence::DoublePipe => converter.buffer.push_str("\\right\\|"),
            MatrixFence::None => {},
        }
    }

    write!(converter.buffer, "\\end{{{}}}", env)
        .map_err(|e| LatexError::FormatError(e.to_string()))?;

    Ok(())
}
