// Performance utilities for LaTeX conversion
//
// This module contains optimized helper functions for high-performance
// LaTeX conversion operations using SIMD and efficient algorithms.

use crate::formula::ast::MathNode;
use memchr::memchr;

/// Fast check if a string represents a valid number using SIMD
#[inline]
#[allow(dead_code)]
pub fn is_valid_number_fast(s: &str) -> bool {
    if s.is_empty() {
        return false;
    }

    let bytes = s.as_bytes();
    let mut has_digit = false;
    let mut has_dot = false;

    // Use SIMD-friendly loop
    for &b in bytes {
        match b {
            b'0'..=b'9' => has_digit = true,
            b'.' => {
                if has_dot {
                    return false; // Multiple dots
                }
                has_dot = true;
            },
            b'-' if bytes.len() == 1 => return false, // Just a minus sign
            b'-' => {},                               // Allow negative numbers at start
            _ => return false,                        // Invalid character
        }
    }

    has_digit // Must have at least one digit
}

/// SIMD-accelerated check if a character sequence contains LaTeX special characters
#[inline]
pub fn contains_latex_special_simd(text: &str) -> bool {
    let bytes = text.as_bytes();

    // Use memchr for common special characters
    memchr(b' ', bytes).is_some()
        || memchr(b'#', bytes).is_some()
        || memchr(b'$', bytes).is_some()
        || memchr(b'%', bytes).is_some()
        || memchr(b'&', bytes).is_some()
        || memchr(b'_', bytes).is_some()
        || memchr(b'{', bytes).is_some()
        || memchr(b'}', bytes).is_some()
        || memchr(b'~', bytes).is_some()
        || memchr(b'^', bytes).is_some()
        || memchr(b'\\', bytes).is_some()
}

/// SIMD-accelerated LaTeX special character escaping
/// Returns true if escaping was needed
#[inline]
#[allow(dead_code)]
pub fn escape_latex_special_chars(text: &str, buffer: &mut String) -> bool {
    if !contains_latex_special_simd(text) {
        buffer.push_str(text);
        return false;
    }

    // Need to escape - process character by character
    for ch in text.chars() {
        match ch {
            ' ' | '#' | '$' | '%' | '&' | '_' | '{' | '}' | '~' | '^' | '\\' => {
                buffer.push('\\');
                buffer.push(ch);
            },
            _ => buffer.push(ch),
        }
    }
    true
}

/// Fast buffer extension with capacity management
#[inline]
pub fn extend_buffer_with_capacity(buffer: &mut String, text: &str, additional_capacity: usize) {
    if buffer.capacity() < buffer.len() + text.len() + additional_capacity {
        buffer.reserve(text.len() + additional_capacity);
    }
    buffer.push_str(text);
}

/// Fast check if text needs LaTeX protection (contains spaces or special chars)
#[inline]
pub fn needs_latex_protection(text: &str) -> bool {
    if text.is_empty() {
        return true;
    }

    // Quick check for spaces
    if memchr(b' ', text.as_bytes()).is_some() {
        return true;
    }

    // Check for other special characters
    contains_latex_special_simd(text)
}

/// Check if base needs grouping for scripts (subscript/superscript)
#[inline]
#[allow(dead_code)]
pub fn needs_grouping_for_scripts(nodes: &[MathNode]) -> bool {
    nodes.len() > 1
}

/// Estimate the output size of a formula for buffer pre-allocation
pub fn estimate_formula_size(nodes: &[MathNode]) -> usize {
    estimate_nodes_size(nodes) + 10 // Add space for delimiters
}

/// Estimate the output size of nodes for buffer pre-allocation
pub fn estimate_nodes_size(nodes: &[MathNode]) -> usize {
    nodes.iter().map(estimate_node_size).sum()
}

/// Estimate the output size of a single node
pub fn estimate_node_size(node: &MathNode) -> usize {
    match node {
        MathNode::Text(text) => {
            if needs_latex_protection(text) {
                text.len() + 10 // \text{} wrapper
            } else {
                text.len()
            }
        },
        MathNode::Number(num) => num.len(),
        MathNode::Operator(_) => 5, // Average operator length
        MathNode::Symbol(_) => 8,   // Average symbol length with escapes
        MathNode::Frac {
            numerator,
            denominator,
            ..
        } => {
            6 + estimate_nodes_size(numerator) + estimate_nodes_size(denominator) // \frac{}{}
        },
        MathNode::Root { base, index } => {
            (if index.is_some() { 8 } else { 7 }) + estimate_nodes_size(base)
        },
        MathNode::Power { base, exponent } => {
            2 + estimate_nodes_size(base) + estimate_nodes_size(exponent) // ^{}
        },
        MathNode::Sub { base, subscript } => {
            2 + estimate_nodes_size(base) + estimate_nodes_size(subscript) // _{}
        },
        MathNode::SubSup {
            base,
            subscript,
            superscript,
        } => {
            4 + estimate_nodes_size(base)
                + estimate_nodes_size(subscript)
                + estimate_nodes_size(superscript) // _{}^{}
        },
        MathNode::Under {
            base,
            under,
            position: _,
        } => {
            10 + estimate_nodes_size(base) + estimate_nodes_size(under) // \underset{}{}
        },
        MathNode::Over {
            base,
            over,
            position: _,
        } => {
            9 + estimate_nodes_size(base) + estimate_nodes_size(over) // \overset{}{}
        },
        MathNode::UnderOver {
            base,
            under,
            over,
            position: _,
        } => {
            20 + estimate_nodes_size(base) + estimate_nodes_size(under) + estimate_nodes_size(over) // \overset{}{\underset{}{}}
        },
        MathNode::Fenced {
            open: _,
            content,
            close: _,
            separator: _,
        } => {
            12 + estimate_nodes_size(content) // \left...\right...
        },
        MathNode::LargeOp {
            operator: _,
            lower_limit,
            upper_limit,
            integrand,
            hide_lower: _,
            hide_upper: _,
        } => {
            8 + // operator
            lower_limit.as_ref().map_or(0, |l| 2 + estimate_nodes_size(l)) +
            upper_limit.as_ref().map_or(0, |u| 2 + estimate_nodes_size(u)) +
            integrand.as_ref().map_or(0, |i| 1 + estimate_nodes_size(i))
        },
        MathNode::Function { name, argument } => {
            name.len() + 5 + estimate_nodes_size(argument) // \name{}
        },
        MathNode::Matrix { rows, .. } => {
            20 + // \begin{matrix}\end{matrix}
            rows.len() * 4 + // \\\\ between rows
            rows.iter().flatten().flatten().count() * 3 // & between cells and content
        },
        MathNode::Accent { base, .. } => {
            8 + estimate_nodes_size(base) // \accent{}
        },
        MathNode::Space(_) => 5,  // Space commands
        MathNode::LineBreak => 2, // \\\\
        MathNode::Style { content, .. } => {
            8 + estimate_nodes_size(content) // \style{}
        },
        MathNode::Row(nodes) => estimate_nodes_size(nodes),
        MathNode::Phantom(content) => {
            9 + estimate_nodes_size(content) // \phantom{}
        },
        MathNode::Error(msg) => {
            15 + msg.len() // \text{[Error: ...]}
        },
        MathNode::PredefinedSymbol(_) => 8, // Average predefined symbol length
        MathNode::PreSub {
            base,
            pre_subscript,
        } => {
            3 + estimate_nodes_size(base) + estimate_nodes_size(pre_subscript) // \presub{}{}
        },
        MathNode::PreSup {
            base,
            pre_superscript,
        } => {
            3 + estimate_nodes_size(base) + estimate_nodes_size(pre_superscript) // \presup{}{}
        },
        MathNode::PreSubSup {
            base,
            pre_subscript,
            pre_superscript,
        } => {
            5 + estimate_nodes_size(base)
                + estimate_nodes_size(pre_subscript)
                + estimate_nodes_size(pre_superscript) // \presubsup{}{}{}
        },
        MathNode::Bar { base, .. } => {
            6 + estimate_nodes_size(base) // \bar{}
        },
        MathNode::BorderBox { content, .. } => {
            12 + estimate_nodes_size(content) // \boxed{}
        },
        MathNode::GroupChar { base, .. } => {
            12 + estimate_nodes_size(base) // \overbrace or similar
        },
        MathNode::PredefinedFunction { argument, .. } => {
            8 + estimate_nodes_size(argument) // Average function length
        },
        MathNode::EqArray { rows, .. } => {
            25 + // \begin{align}\end{align}
            rows.len() * 4 + // \\\\ between rows
            rows.iter().flatten().count() * 2 // Content size
        },
        MathNode::Run { content, .. } => estimate_nodes_size(content),
        MathNode::Limit { content, .. } => estimate_nodes_size(content),
        MathNode::Degree(content) => estimate_nodes_size(content),
        MathNode::Base(content) => estimate_nodes_size(content),
        MathNode::Argument(content) => estimate_nodes_size(content),
        MathNode::Numerator(content) => estimate_nodes_size(content),
        MathNode::Denominator(content) => estimate_nodes_size(content),
        MathNode::Integrand(content) => estimate_nodes_size(content),
        MathNode::LowerLimit(content) => estimate_nodes_size(content),
        MathNode::UpperLimit(content) => estimate_nodes_size(content),
    }
}

/// Estimate capacity needed for matrix conversion to avoid reallocations
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
