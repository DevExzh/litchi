//! Regression tests for the DOCX package owner.

use super::model::*;

use std::io::{Cursor, Write};
use tempfile::NamedTempFile;

mod document;
mod document_memos;
mod graph;
mod revision_actions;
mod settings;
