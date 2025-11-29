//! High-performance ZIP archive library optimized for Office document formats.
//!
//! This crate provides efficient ZIP reading and writing specifically designed
//! for OOXML (.docx, .xlsx, .pptx), ODF (.odt, .ods, .odp), and iWork
//! (.pages, .numbers, .key) file formats.
//!
//! # Quick Start
//!
//! For most use cases, use the high-level [`office`] module:
//!
//! ```rust,no_run
//! use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};
//!
//! // Reading
//! let data = std::fs::read("document.docx")?;
//! let archive = ArchiveReader::new(&data)?;
//! let content = archive.read("word/document.xml")?;
//!
//! // Writing
//! let mut writer = StreamingArchiveWriter::new();
//! writer.write_deflated("content.xml", b"<root/>")?;
//! let bytes = writer.finish_to_bytes()?;
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
#![forbid(unsafe_code)]

mod archive;
mod crc;
mod errors;
pub mod extra_fields;
mod headers;
mod locator;
mod mode;
pub mod office;
pub mod path;
mod reader_at;
pub mod time;
mod utils;
mod writer;

pub use archive::*;
pub use crc::crc32;
pub use errors::{Error, ErrorKind};
pub use headers::Header;
pub use locator::*;
pub use mode::EntryMode;
pub use reader_at::{FileReader, RangeReader, ReaderAt};
pub use writer::*;
