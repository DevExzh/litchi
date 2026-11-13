// Matrix conversion to LaTeX
//
// This module handles conversion of matrix nodes to LaTeX format with
// performance optimizations and proper fence handling.

use crate::ast::MatrixFence;

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

#[cfg(test)]
mod tests {
    use super::*;

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
    fn test_matrix_fence_to_env_comprehensive() {
        assert_eq!(matrix_fence_to_env(MatrixFence::None), "matrix");
        assert_eq!(matrix_fence_to_env(MatrixFence::Paren), "pmatrix");
        assert_eq!(matrix_fence_to_env(MatrixFence::Bracket), "bmatrix");
        assert_eq!(matrix_fence_to_env(MatrixFence::Brace), "Bmatrix");
        assert_eq!(matrix_fence_to_env(MatrixFence::Pipe), "vmatrix");
        assert_eq!(matrix_fence_to_env(MatrixFence::DoublePipe), "Vmatrix");
    }
}
