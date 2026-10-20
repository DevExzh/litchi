//! Regressions of the crate-private part families (record 0358).
//!
//! These suites drove the part-level `Snapshot`/`Edit`/`Commit`/`Patch`
//! workflows and the `*_with_limits` parser twins as integration tests.
//! Record 0358 makes those workflows crate-private (the public routes are
//! `Package::edit` and `builder::Workbook`), so the suites compile here as
//! unit tests of the same code.

#[path = "../tests/support/mod.rs"]
mod support;

#[path = "../tests/internal/calculation_settings_leniency.rs"]
mod calculation_settings_leniency;
#[path = "../tests/internal/content_validation_inventory.rs"]
mod content_validation_inventory;
#[path = "../tests/internal/content_validation_transactions.rs"]
mod content_validation_transactions;
#[path = "../tests/internal/definition_transactions.rs"]
mod definition_transactions;
#[path = "../tests/internal/ods_dde_facades.rs"]
mod ods_dde_facades;
#[path = "../tests/internal/ods_dde_transactions.rs"]
mod ods_dde_transactions;
#[path = "../tests/internal/ods_formula_references.rs"]
mod ods_formula_references;
#[path = "../tests/internal/ods_metadata_transactions.rs"]
mod ods_metadata_transactions;
#[path = "../tests/internal/ods_sheet_metadata_transactions.rs"]
mod ods_sheet_metadata_transactions;
#[path = "../tests/internal/rdf_transactions.rs"]
mod rdf_transactions;
#[path = "../tests/internal/scenario_transactions.rs"]
mod scenario_transactions;
#[path = "../tests/internal/sheet_metadata_context.rs"]
mod sheet_metadata_context;
#[path = "../tests/internal/sheet_metadata_selector_search.rs"]
mod sheet_metadata_selector_search;
#[path = "../tests/internal/sheet_metadata_xml_conformance.rs"]
mod sheet_metadata_xml_conformance;
#[path = "../tests/internal/tracked_changes.rs"]
mod tracked_changes;
#[path = "../tests/internal/tracked_changes_limits.rs"]
mod tracked_changes_limits;
#[path = "../tests/internal/worksheet_cell_batch.rs"]
mod worksheet_cell_batch;
#[path = "../tests/internal/worksheet_hyperlinks.rs"]
mod worksheet_hyperlinks;
#[path = "../tests/internal/worksheet_transactions.rs"]
mod worksheet_transactions;
