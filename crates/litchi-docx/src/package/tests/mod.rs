//! Regression tests for the DOCX package owner.

#[allow(
    clippy::wildcard_imports,
    reason = "Package tests exercise the complete owner facade."
)]
use super::model::*;

use std::io::{Cursor, Seek, Write};
use std::panic::{AssertUnwindSafe, catch_unwind};
use tempfile::NamedTempFile;

struct FailingWriter;

impl Write for FailingWriter {
    fn write(&mut self, _buffer: &[u8]) -> std::io::Result<usize> {
        Err(std::io::Error::other("injected DOCX sink failure"))
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

impl Seek for FailingWriter {
    fn seek(&mut self, _position: std::io::SeekFrom) -> std::io::Result<u64> {
        Err(std::io::Error::other("injected DOCX seek failure"))
    }
}

struct PanickingWriter;

impl Write for PanickingWriter {
    fn write(&mut self, _buffer: &[u8]) -> std::io::Result<usize> {
        panic!("injected DOCX sink panic")
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

impl Seek for PanickingWriter {
    fn seek(&mut self, _position: std::io::SeekFrom) -> std::io::Result<u64> {
        panic!("injected DOCX seek panic")
    }
}

mod document;
mod graph;
mod settings;
