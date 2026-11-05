//! Regressions of crate-private mutation paths (record 0358).
//!
//! These suites drove the `&mut self` mutators of the generic packaged and
//! flat owners, the standalone chart document and `MutableDocument` as
//! integration tests. Record 0358 makes those mutators crate-private (the
//! public routes are the owners' `edit()` families and `Package::edit`), so
//! the suites compile here as unit tests of the same code.

#[path = "../tests/support/mod.rs"]
mod support;

#[path = "../tests/internal/bibliography_configuration_api.rs"]
mod bibliography_configuration_api;
#[path = "../tests/internal/document_verbs.rs"]
mod document_verbs;
#[path = "../tests/internal/drawing_style_resources_api.rs"]
mod drawing_style_resources_api;
#[path = "../tests/internal/dynamic_text_fields.rs"]
mod dynamic_text_fields;
#[path = "../tests/internal/embedded_resource_batch.rs"]
mod embedded_resource_batch;
#[path = "../tests/internal/flat_text_templates.rs"]
mod flat_text_templates;
#[path = "../tests/internal/font_face_declarations_api.rs"]
mod font_face_declarations_api;
#[path = "../tests/internal/footnote_separators.rs"]
mod footnote_separators;
#[path = "../tests/internal/generic_content_publication.rs"]
mod generic_content_publication;
#[path = "../tests/internal/generic_variable_declaration_mutation.rs"]
mod generic_variable_declaration_mutation;
#[path = "../tests/internal/header_footer_properties.rs"]
mod header_footer_properties;
#[path = "../tests/internal/in_content_metadata.rs"]
mod in_content_metadata;
#[path = "../tests/internal/list_label_alignment.rs"]
mod list_label_alignment;
#[path = "../tests/internal/notes_configurations_api.rs"]
mod notes_configurations_api;
#[path = "../tests/internal/odt_form_controls.rs"]
mod odt_form_controls;
#[path = "../tests/internal/odt_page_sequence_authoring.rs"]
mod odt_page_sequence_authoring;
#[path = "../tests/internal/odt_section_mutation.rs"]
mod odt_section_mutation;
#[path = "../tests/internal/odt_tracked_change_mutation.rs"]
mod odt_tracked_change_mutation;
#[path = "../tests/internal/outline_styles_api.rs"]
mod outline_styles_api;
#[path = "../tests/internal/packaged_transactions.rs"]
mod packaged_transactions;
#[path = "../tests/internal/paragraph_drop_caps.rs"]
mod paragraph_drop_caps;
#[path = "../tests/internal/paragraph_tab_stops.rs"]
mod paragraph_tab_stops;
#[path = "../tests/internal/style_columns.rs"]
mod style_columns;
#[path = "../tests/internal/variable_declaration_mutation.rs"]
mod variable_declaration_mutation;
