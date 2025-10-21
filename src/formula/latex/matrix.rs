// Matrix conversion to LaTeX
//
// This module handles conversion of matrix nodes to LaTeX format with
// performance optimizations and proper fence handling.

use super::LatexError;
use crate::formula::ast::{MathNode, MatrixFence};

/// Convert matrix fence type to LaTeX environment name
#[inline]
pub fn matrix_fence_to_env(fence_type: MatrixFence) -> &'static str {
    match fence_type {
        MatrixFence::None => "matrix",
        MatrixFence::Paren => "pmatrix",
        MatrixFence::Bracket => "bmatrix",
        MatrixFence::Brace => "Bmatrix",
        MatrixFence::Pipe => "vmatrix",
        MatrixFence::DoublePipe => "Vmatrix",
    }
}

/// Convert matrix to LaTeX with optimized string building
///
/// Uses pre-allocated capacity and efficient string operations for performance.
/// This function provides a public API for matrix conversion with a custom node converter.
///
/// # Arguments
/// * `buffer` - The output buffer to write LaTeX to
/// * `rows` - Matrix rows containing cells with MathNode vectors
/// * `fence_type` - The type of fence to use (parentheses, brackets, etc.)
/// * `node_converter` - Function to convert individual MathNodes to LaTeX
///
/// # Performance
/// Pre-allocates buffer capacity and uses efficient string operations.
#[allow(dead_code)]
pub fn convert_matrix(
    buffer: &mut String,
    rows: &[Vec<Vec<MathNode>>],
    fence_type: MatrixFence,
    node_converter: &dyn Fn(&mut String, &MathNode) -> Result<(), LatexError>,
) -> Result<(), LatexError> {
    use std::fmt::Write;

    if rows.is_empty() {
        return Ok(());
    }

    let env = matrix_fence_to_env(fence_type);

    // Pre-calculate approximate capacity needed for better performance
    let estimated_capacity = estimate_matrix_capacity(rows);
    buffer.reserve(estimated_capacity);

    // Begin environment
    write!(buffer, "\\begin{{{}}}", env).map_err(|e| LatexError::FormatError(e.to_string()))?;

    // Convert each row
    for (i, row) in rows.iter().enumerate() {
        if i > 0 {
            buffer.push_str(" \\\\ ");
        }

        // Convert each cell in the row
        for (j, cell) in row.iter().enumerate() {
            if j > 0 {
                buffer.push_str(" & ");
            }

            // Convert all nodes in this cell using the provided converter
            for node in cell {
                node_converter(buffer, node)?;
            }
        }
    }

    // End environment
    write!(buffer, "\\end{{{}}}", env).map_err(|e| LatexError::FormatError(e.to_string()))?;

    Ok(())
}

/// Estimate capacity needed for matrix conversion to avoid reallocations
#[allow(dead_code)]
pub fn estimate_matrix_capacity(rows: &[Vec<Vec<MathNode>>]) -> usize {
    if rows.is_empty() {
        return 0;
    }

    let num_rows = rows.len();
    let num_cols = rows[0].len();

    // Estimate: environment markers + row separators + column separators + content
    let env_overhead = 20; // \begin{matrix}\end{matrix}
    let row_separators = (num_rows.saturating_sub(1)) * 4; // " \\\\ "
    let col_separators = num_rows * (num_cols.saturating_sub(1)) * 3; // " & "

    // Rough estimate for content (average 5 chars per node)
    let content_estimate = rows.iter().flatten().flatten().count() * 5;

    env_overhead + row_separators + col_separators + content_estimate
}

/// Convert matrix with alignment specification for columns
///
/// For matrices that require specific column alignment (left, center, right).
/// Uses LaTeX array environment instead of standard matrix environments when alignment is specified.
///
/// # Arguments
/// * `buffer` - The output buffer to write LaTeX to
/// * `rows` - Matrix rows containing cells with MathNode vectors
/// * `fence_type` - The type of fence to use (parentheses, brackets, etc.)
/// * `alignments` - Optional column alignment specifications ('l', 'c', 'r')
/// * `node_converter` - Function to convert individual MathNodes to LaTeX
///
/// # Performance
/// Pre-allocates buffer capacity and uses efficient string operations.
/// When alignments are specified, uses array environment for precise control.
#[allow(dead_code)]
pub fn convert_matrix_with_alignment(
    buffer: &mut String,
    rows: &[Vec<Vec<MathNode>>],
    fence_type: MatrixFence,
    alignments: Option<&[char]>,
    node_converter: &dyn Fn(&mut String, &MathNode) -> Result<(), LatexError>,
) -> Result<(), LatexError> {
    use std::fmt::Write;

    if rows.is_empty() {
        return Ok(());
    }

    // Determine if we need alignment-specific environment
    let use_array_env = alignments.is_some();

    let env = if use_array_env {
        // Use array environment for alignment control
        "array"
    } else {
        matrix_fence_to_env(fence_type)
    };

    // Pre-calculate capacity with extra space for alignment specifications
    let mut estimated_capacity = estimate_matrix_capacity(rows);
    if use_array_env {
        estimated_capacity += 20; // Extra for alignment spec
    }
    buffer.reserve(estimated_capacity);

    // Begin environment
    if use_array_env {
        // For array environment, we need alignment specification
        if let Some(aligns) = alignments {
            write!(buffer, "\\begin{{{}}}", env)
                .map_err(|e| LatexError::FormatError(e.to_string()))?;
            buffer.push('{');
            for &align in aligns {
                buffer.push(align);
            }
            buffer.push('}');
        } else {
            // Default to centered alignment if array but no alignments specified
            let num_cols = rows[0].len();
            write!(buffer, "\\begin{{{}}}", env)
                .map_err(|e| LatexError::FormatError(e.to_string()))?;
            buffer.push('{');
            for _ in 0..num_cols {
                buffer.push('c');
            }
            buffer.push('}');
        }

        // Add fence manually for array environment
        match fence_type {
            MatrixFence::Paren => buffer.push_str("\\left("),
            MatrixFence::Bracket => buffer.push_str("\\left["),
            MatrixFence::Brace => buffer.push_str("\\left\\{"),
            MatrixFence::Pipe => buffer.push_str("\\left|"),
            MatrixFence::DoublePipe => buffer.push_str("\\left\\|"),
            MatrixFence::None => {}, // No fence
        }
    } else {
        write!(buffer, "\\begin{{{}}}", env).map_err(|e| LatexError::FormatError(e.to_string()))?;
    }

    // Convert each row
    for (i, row) in rows.iter().enumerate() {
        if i > 0 {
            buffer.push_str(" \\\\ ");
        }

        // Convert each cell in the row
        for (j, cell) in row.iter().enumerate() {
            if j > 0 {
                buffer.push_str(" & ");
            }

            // Convert all nodes in this cell using the provided converter
            for node in cell {
                node_converter(buffer, node)?;
            }
        }
    }

    // Close fence for array environment
    if use_array_env {
        match fence_type {
            MatrixFence::Paren => buffer.push_str("\\right)"),
            MatrixFence::Bracket => buffer.push_str("\\right]"),
            MatrixFence::Brace => buffer.push_str("\\right\\}"),
            MatrixFence::Pipe => buffer.push_str("\\right|"),
            MatrixFence::DoublePipe => buffer.push_str("\\right\\|"),
            MatrixFence::None => {}, // No fence
        }
    }

    // End environment
    write!(buffer, "\\end{{{}}}", env).map_err(|e| LatexError::FormatError(e.to_string()))?;

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::formula::ast::{MathNode, Operator};

    fn dummy_converter(buffer: &mut String, node: &MathNode) -> Result<(), LatexError> {
        match node {
            MathNode::Number(n) => buffer.push_str(n),
            MathNode::Operator(op) => buffer.push_str(match op {
                Operator::Plus => "+",
                Operator::Minus => "-",
                _ => "?",
            }),
            _ => buffer.push('?'),
        }
        Ok(())
    }

    #[test]
    fn test_matrix_fence_to_env() {
        assert_eq!(matrix_fence_to_env(MatrixFence::None), "matrix");
        assert_eq!(matrix_fence_to_env(MatrixFence::Paren), "pmatrix");
        assert_eq!(matrix_fence_to_env(MatrixFence::Bracket), "bmatrix");
        assert_eq!(matrix_fence_to_env(MatrixFence::Brace), "Bmatrix");
        assert_eq!(matrix_fence_to_env(MatrixFence::Pipe), "vmatrix");
        assert_eq!(matrix_fence_to_env(MatrixFence::DoublePipe), "Vmatrix");
    }

    #[test]
    fn test_convert_simple_matrix() {
        let mut buffer = String::new();
        let rows = vec![
            vec![
                vec![MathNode::Number("1".into())],
                vec![MathNode::Number("2".into())],
            ],
            vec![
                vec![MathNode::Number("3".into())],
                vec![MathNode::Number("4".into())],
            ],
        ];

        convert_matrix(&mut buffer, &rows, MatrixFence::Bracket, &dummy_converter).unwrap();

        assert_eq!(buffer, "\\begin{bmatrix}1 & 2 \\\\ 3 & 4\\end{bmatrix}");
    }

    #[test]
    fn test_convert_empty_matrix() {
        let mut buffer = String::new();
        let rows: Vec<Vec<Vec<MathNode>>> = vec![];

        convert_matrix(&mut buffer, &rows, MatrixFence::None, &dummy_converter).unwrap();

        assert_eq!(buffer, "");
    }

    #[test]
    fn test_convert_matrix_with_alignment() {
        let mut buffer = String::new();
        let rows = vec![
            vec![
                vec![MathNode::Number("1".into())],
                vec![MathNode::Number("2".into())],
                vec![MathNode::Number("3".into())],
            ],
            vec![
                vec![MathNode::Number("4".into())],
                vec![MathNode::Number("5".into())],
                vec![MathNode::Number("6".into())],
            ],
        ];
        let alignments = vec!['l', 'c', 'r'];

        convert_matrix_with_alignment(
            &mut buffer,
            &rows,
            MatrixFence::Bracket,
            Some(&alignments),
            &dummy_converter,
        )
        .unwrap();

        // Should use array environment with alignment specification
        assert!(buffer.contains("\\begin{array}{lcr}"));
        assert!(buffer.contains("\\left["));
        assert!(buffer.contains("\\right]"));
        assert!(buffer.contains("\\end{array}"));
        assert!(buffer.contains("1 & 2 & 3"));
        assert!(buffer.contains("4 & 5 & 6"));
    }

    #[test]
    fn test_convert_matrix_with_alignment_default() {
        let mut buffer = String::new();
        let rows = vec![vec![
            vec![MathNode::Number("1".into())],
            vec![MathNode::Number("2".into())],
        ]];

        // No alignments specified, should use standard matrix environment
        convert_matrix_with_alignment(
            &mut buffer,
            &rows,
            MatrixFence::Paren,
            None,
            &dummy_converter,
        )
        .unwrap();

        assert!(buffer.contains("\\begin{pmatrix}"));
        assert!(buffer.contains("\\end{pmatrix}"));
    }

    #[test]
    fn test_matrix_fence_to_env_comprehensive() {
        assert_eq!(matrix_fence_to_env(MatrixFence::None), "matrix");
        assert_eq!(matrix_fence_to_env(MatrixFence::Paren), "pmatrix");
        assert_eq!(matrix_fence_to_env(MatrixFence::Bracket), "bmatrix");
        assert_eq!(matrix_fence_to_env(MatrixFence::Brace), "Bmatrix");
        assert_eq!(matrix_fence_to_env(MatrixFence::Pipe), "vmatrix");
        assert_eq!(matrix_fence_to_env(MatrixFence::DoublePipe), "Vmatrix");
    }
}
