//! Regressions of the crate-private body-text engine (record 0359a).
//!
//! These suites drove `body_text::{Snapshot, Edit, Commit, Patch}` as
//! integration tests. Record 0359a makes the engine crate-private (the
//! public route is `Package::edit` and the package-level body routes), so
//! the suites compile here as unit tests of the same code. The crate is
//! named `litchi_doc` inside them (`extern crate self`), so their paths are
//! unchanged.

#[path = "../tests/common/mod.rs"]
mod common;

#[path = "../tests/internal/auxiliary_tables.rs"]
mod auxiliary_tables;
#[path = "../tests/internal/doc_body_transaction.rs"]
mod doc_body_transaction;
#[path = "../tests/internal/doc_render_handoff.rs"]
mod doc_render_handoff;
#[path = "../tests/internal/doc_semantic_body_paragraph.rs"]
mod doc_semantic_body_paragraph;
#[path = "../tests/internal/doc_story_reads.rs"]
mod doc_story_reads;
#[path = "../tests/internal/dofr_edits.rs"]
mod dofr_edits;
#[path = "../tests/internal/saved_selection.rs"]
mod saved_selection;
