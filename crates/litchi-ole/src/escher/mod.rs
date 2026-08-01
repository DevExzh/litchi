//! Shared OfficeArt (Escher) functionality for Office binary formats.
//!
//! Escher is Microsoft's drawing layer format used across Office applications
//! (DOC, XLS, PPT). This module provides shared, zero-copy parsing and writing
//! utilities for Escher records.
//!
//! # Architecture
//!
//! - **Zero-copy parsing**: Lifetime-based borrowing from source data
//! - **Modular design**: Separated concerns for parsing, writing, and shapes
//! - **Format-agnostic**: Core functionality shared across all formats
//! - **Performance-focused**: Minimal allocations and efficient iteration
//!
//! # Modules
//!
//! - `types`: Escher record type definitions
//! - `record`: Zero-copy record structure
//! - `container`: Container record handling with iterators
//! - `parser`: High-level parsing interface
//! - `properties`: Property system (Opt records)
//! - `shape`: Shape abstraction and utilities
//! - `text`: Text extraction from Escher records
//! - `writer`: Escher record generation utilities

pub mod container;
pub mod parser;
pub mod properties;
pub mod record;
pub mod shape;
pub mod shape_factory;
pub mod text;
pub mod types;
pub mod writer;
