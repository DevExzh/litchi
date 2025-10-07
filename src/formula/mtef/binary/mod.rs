// MTEF Binary Parser - Modular implementation based on rtf2latex2e
//
// This module implements proper binary parsing of MTEF (MathType Equation Format)
// records as used in OLE documents, following the structure of rtf2latex2e.

pub mod charset;
pub mod headers;
pub mod objects;
pub mod parser;
pub mod converter;

pub use parser::*;
