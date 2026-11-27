//! SPRM (Single Property Modifier) operation constants and utilities.
//!
//! This module provides comprehensive SPRM definitions based on Apache POI's implementation.
//! SPRMs are used in DOC and PPT formats to modify character, paragraph, table, and section properties.
//!
//! Reference: Apache POI's hwpf/sprm package and usermodel/*Properties.java
//!
//! # SPRM Structure
//!
//! A SPRM consists of:
//! - **Opcode** (2 bytes): Encodes the operation type and size
//!   - Bits 0-8: Operation code
//!   - Bits 9: Special flag
//!   - Bits 10-12: Property type (CHP=2, PAP=1, TAP=5, etc.)
//!   - Bits 13-15: Size code (determines operand size)
//! - **Operand** (variable): The data for the operation
//!
//! # Size Codes
//!
//! - 0, 1: 1-byte operand
//! - 2, 4, 5: 2-byte operand
//! - 3: 4-byte operand
//! - 6: Variable length (size in first byte or word)
//! - 7: 3-byte operand
//!
//! The public constants live in the semantic model layer; opcode bitfield
//! interpretation lives in the codec layer.

mod codec;
mod model;

pub use codec::{get_sprm_operation, get_sprm_size_code, get_sprm_type, is_sprm_special};
pub use model::*;

#[cfg(test)]
mod tests;
