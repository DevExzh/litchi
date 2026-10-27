//! Host-vs-leaf differential parity harness for the `litchi-iwa` exit
//! (record 0363 §4; usage in record 0369).
//!
//! Every case runs one operation, expressed once, through the host route
//! (`litchi_iwa::{numbers,pages,keynote}::Draft`) and through the leaf route
//! (`litchi_{numbers,pages,keynote}::Package`) on the same fixture bytes and
//! asserts the tier the case declares:
//!
//! - **Tier A**: the two outputs are byte-identical, the whole ZIP included;
//! - **Tier B**: every ZIP member and every IWA object the host left
//!   untouched is byte-identical in the leaf output, and every object the
//!   host changed is equal in the leaf output on every known protobuf field;
//!   unknown-field and ZIP-layout differences are reported, not failed.
//!
//! Both tiers also require identical reads of both outputs, a strict reopen
//! of both through the leaf `Package`, equal object identifiers, and (when
//! the leaf route returns one) an inverse that restores the input bytes.
//! Each case declares whether its host route has native Apple evidence
//! (owner decision D1); `tools/iwa_exit_ledger.py` reads that declaration
//! and requires a `docs/unified-ops/0363-mac-queue.md` row for every
//! migrated route without it.
//!
//! The harness lives in the host's tests so that W3 deletes it with the host.

mod archive_diff;
mod chart_style_cases;
mod cost_probes;
mod formula_refresh_cases;
mod harness;
mod known_fields;
mod media_cases;

mod keynote_build_cases;
mod keynote_builder_cases;
mod keynote_cases;
mod keynote_cell_format_cases;
mod keynote_chart_cases;
mod keynote_drawable_cases;
mod keynote_image_cases;
mod keynote_named_style_cases;
mod keynote_shape_style_cases;
mod keynote_slide_cases;
mod keynote_table_cases;
mod keynote_text_annotation_cases;
mod keynote_text_list_cases;
mod keynote_text_property_cases;
mod keynote_text_storage_cases;
mod keynote_w1c_cost;
mod numbers_builder_cases;
mod numbers_cases;
mod numbers_cell_list_cases;
mod numbers_cell_text_cases;
mod numbers_chart_style_cases;
mod numbers_conditional_highlight_cases;
mod numbers_drawable_cases;
mod numbers_lifecycle_cases;
mod numbers_owner_allocation;
mod numbers_text_box_cases;
mod numbers_w2a_cost;
mod pages_caption_cases;
mod pages_cases;
mod pages_cell_format_cases;
mod pages_cost_probes;
mod pages_drawable_cases;
mod pages_lifecycle_cases;
mod pages_media_lifecycle_cases;
mod pages_table_cases;
mod pages_text_cases;
mod physical_cases;
mod self_test;
mod series_graph_cases;
mod series_style_cases;
mod series_write_cases;
mod shape_cases;
mod shape_style_cases;
mod text_annotation_cases;
mod text_attachment_cases;
mod text_comment_cases;
mod text_cost_probes;
mod text_drop_cap_cases;
mod text_field_cases;
mod text_glue;
mod text_language_cases;
mod text_list_cases;
mod text_named_style_cases;
mod text_style_cases;
