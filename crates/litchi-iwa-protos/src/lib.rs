#![forbid(unsafe_code)]

//! Raw generated Protocol Buffer types used by Apple iWork IWA archives.
//!
//! This crate owns only the schema and code-generation boundary. It does not
//! decode IWA objects, traverse packages, or provide application-specific
//! semantics. Consumers that need those behaviors should depend on the
//! appropriate `litchi-pages`, `litchi-keynote`, `litchi-numbers`, or
//! `litchi-iwa-archive` crate instead.

#[cfg(test)]
#[path = "production_codec_guard.rs"]
mod production_codec_guard;

/// Generated source is kept behind one audited boundary so workspace lints
/// continue to apply to every hand-written item in this crate.
#[doc(hidden)]
mod generated {
    #![allow(
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        dead_code,
        unreachable_pub,
        reason = "prost-build output is generated from the native IWA schemas."
    )]

    include!(concat!(env!("OUT_DIR"), "/iwa_protos.rs"));
}

/// Private Buffa eager/lazy view projection for the archive-header adapter.
///
/// The projection is deliberately not an untrusted-ingress API. In Buffa
/// 0.9.1, deferred lazy message access rebuilds its decode context with
/// recursion and unknown-field limits only; it does not retain the original
/// `DecodeOptions::with_element_memory_limit` budget. Nested deferred
/// allocations can therefore escape that initial element-memory accounting.
/// Archive adapters must establish their own complete resource policy before
/// accepting untrusted payloads through this path.
#[doc(hidden)]
mod buffa_archive_header_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa sidecar is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the archive-header projection."
    )]
    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-archive-header/iwa_archive_header_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for the nested Keynote movie
/// `TSP.DataReference` envelope.
#[doc(hidden)]
mod buffa_data_reference_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa sidecar is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the DataReference projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-data-reference/iwa_data_reference_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for `TSWP.StorageArchive.text` only.
///
/// It is generated in isolation with unknown retention disabled. The source
/// IWA payload remains the byte-authoritative preservation representation.
#[doc(hidden)]
mod buffa_text_storage_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-text-storage/iwa_text_storage_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for Numbers comment storage.
///
/// The native repeated `replies` field is deliberately absent.  The strict
/// handwritten codec validates and streams each reply directly from the
/// caller-owned source bytes.
#[doc(hidden)]
mod buffa_comment_storage_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-numbers-comment-storage/iwa_comment_storage_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for the GroupNode category-label path.
///
/// It includes only an empty node envelope, UUID identity, and the scalar
/// Boolean, Date, Number, and String wrappers. The adapter streams child and
/// CellValue routing from source bytes, which remain authoritative for
/// preservation.
#[doc(hidden)]
mod buffa_group_node_category_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-group-node-category/iwa_group_node_category_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for the Keynote root show reference.
///
/// The required root base archive remains opaque. Only the nested show
/// identifier is decoded, while caller-owned IWA bytes remain authoritative.
#[doc(hidden)]
mod buffa_keynote_document_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-keynote-document/iwa_keynote_document_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for the Keynote chart-caption edge.
///
/// Only the drawable `super`, caption reference, and nested identifier are
/// generated. The chart extension closure and unrelated source fields remain
/// caller-owned for strict validation and preservation.
#[doc(hidden)]
mod buffa_keynote_chart_caption_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-keynote-chart-caption/iwa_keynote_chart_caption_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for the Keynote movie title/caption
/// edge. Only the required movie `super` envelope and the two optional
/// drawable references are generated; all other MovieArchive fields remain
/// source-owned for strict validation and raw preservation.
#[doc(hidden)]
mod buffa_keynote_movie_caption_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "Buffa 0.9.1 generated views are private implementation detail."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-keynote-movie-caption/iwa_keynote_movie_caption_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for the selected scalar playback fields
/// of one TSD MovieArchive. The complete media graph remains source-owned.
#[doc(hidden)]
mod buffa_movie_playback_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "Buffa 0.9.1 generated views are private implementation detail."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-movie-playback/iwa_movie_playback_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for the selected Keynote MovieArchive
/// geometry envelopes. Point and Size remain borrowed raw payloads; the
/// handwritten codec owns their strict fixed32 validation and preservation.
#[doc(hidden)]
mod buffa_keynote_movie_geometry_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "Buffa generated views are private implementation detail."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-keynote-movie-geometry/iwa_keynote_movie_geometry_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for the selected Keynote chart-title
/// generated extension fields.
///
/// Only fields 21 and 23 of the generated chart non-style extension are
/// generated. The outer non-style envelope, chart graph, and all unknown
/// source bytes remain caller-owned.
#[doc(hidden)]
mod buffa_keynote_chart_title_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-keynote-chart-title/iwa_keynote_chart_title_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for the selected Keynote chart-legend
/// generated-extension field.
#[doc(hidden)]
mod buffa_keynote_chart_legend_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-keynote-chart-legend/iwa_keynote_chart_legend_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for the selected Keynote chart Arrange
/// controls. The chart drawable envelope carries only the nested drawable
/// `super`; the lock booleans and every unrelated source byte remain owned by
/// the strict handwritten codec.
#[doc(hidden)]
mod buffa_keynote_chart_arrangement_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-keynote-chart-arrangement/iwa_keynote_chart_arrangement_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for the selected Keynote chart-axis
/// title generated-extension fields.
#[doc(hidden)]
mod buffa_keynote_chart_axis_title_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "Buffa 0.9.1 generated views are private implementation detail."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-keynote-chart-axis-title/iwa_keynote_chart_axis_title_buffa_protos.rs"
    ));
}

/// Private Buffa borrowed-view projection for the selected Keynote chart value
/// axis settings. The handwritten codec performs strict raw preflight before
/// accessing this view; generated values never cross this crate boundary.
#[doc(hidden)]
mod buffa_keynote_chart_axis_value_settings_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "Buffa 0.9.1 generated views are private implementation detail."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-keynote-chart-axis-value-settings/iwa_keynote_chart_axis_value_settings_buffa_protos.rs"
    ));
}

/// Private strict Buffa lazy-view projection for Keynote placeholder text.
///
/// Generated code sees only the singular inheritance chain, optional kind,
/// and optional owned-storage edge. Caller-owned bytes remain authoritative.
#[doc(hidden)]
mod buffa_keynote_placeholder_text_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-keynote-placeholder-text/iwa_keynote_placeholder_text_buffa_protos.rs"
    ));
}

/// Private strict Buffa lazy-view projection for focused Keynote slide owners.
///
/// Generated code sees only exact references, selector-facing slide scalars,
/// semantic note/title/body edges, and required slide/note envelopes. Unknown
/// bytes remain solely in caller-owned IWA and are never retained here.
#[doc(hidden)]
mod buffa_keynote_speaker_notes_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-keynote-speaker-notes/iwa_keynote_speaker_notes_buffa_protos.rs"
    ));
}

/// Private Buffa lazy views for the focused Keynote slide-number path.
#[doc(hidden)]
mod buffa_keynote_slide_number_generated {
    #![allow(
        elided_lifetimes_in_paths,
        unreachable_pub,
        clippy::all,
        clippy::allow_attributes_without_reason,
        clippy::arbitrary_source_item_ordering,
        clippy::map_err_ignore,
        clippy::module_name_repetitions,
        clippy::pedantic,
        clippy::shadow_reuse,
        clippy::shadow_same,
        non_snake_case,
        reason = "Buffa generated projection is private implementation detail."
    )]
    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-keynote-slide-number/iwa_keynote_slide_number_buffa_protos.rs"
    ));
}

/// Private Buffa lazy views for focused Keynote soundtrack settings.
#[doc(hidden)]
mod buffa_keynote_soundtrack_settings_generated {
    #![allow(
        elided_lifetimes_in_paths,
        unreachable_pub,
        clippy::all,
        clippy::allow_attributes_without_reason,
        clippy::arbitrary_source_item_ordering,
        clippy::map_err_ignore,
        clippy::module_name_repetitions,
        clippy::pedantic,
        clippy::shadow_reuse,
        clippy::shadow_same,
        non_snake_case,
        reason = "Buffa generated projection is private implementation detail."
    )]
    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-keynote-soundtrack-settings/iwa_keynote_soundtrack_settings_buffa_protos.rs"
    ));
}

#[doc(hidden)]
mod buffa_numbers_sheet_order_generated {
    #![allow(
        clippy::all,
        clippy::pedantic,
        clippy::arbitrary_source_item_ordering,
        clippy::allow_attributes_without_reason,
        clippy::module_name_repetitions,
        clippy::shadow_same,
        elided_lifetimes_in_paths,
        unreachable_pub,
        non_snake_case,
        reason = "Private Buffa generated projection."
    )]
    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-numbers-sheet-order/iwa_numbers_sheet_order_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for scalar fields of one Pages
/// drawable-order reference. The repeated TP envelope remains handwritten and
/// borrowed by `pages_drawable_order_codec`.
#[doc(hidden)]
mod buffa_pages_drawable_order_generated {
    #![allow(
        clippy::all,
        clippy::pedantic,
        clippy::arbitrary_source_item_ordering,
        clippy::allow_attributes_without_reason,
        clippy::module_name_repetitions,
        clippy::shadow_same,
        elided_lifetimes_in_paths,
        unreachable_pub,
        non_snake_case,
        reason = "Private Buffa generated projection."
    )]
    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-pages-drawable-order/iwa_pages_drawable_order_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for Numbers table-title scalars.
///
/// Title style references are routed from caller-owned bytes and forced
/// through the separately pinned scalar Numbers reference projection. This
/// module therefore contains no collection or nested message field.
#[doc(hidden)]
mod buffa_numbers_table_title_generated {
    #![allow(
        clippy::all,
        clippy::pedantic,
        clippy::arbitrary_source_item_ordering,
        clippy::allow_attributes_without_reason,
        clippy::module_name_repetitions,
        clippy::shadow_same,
        elided_lifetimes_in_paths,
        unreachable_pub,
        non_snake_case,
        reason = "Private Buffa generated projection."
    )]
    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-numbers-table-title/iwa_numbers_table_title_buffa_protos.rs"
    ));
}

/// Private lazy-view roots for Numbers table-cell storage envelopes.
#[doc(hidden)]
mod buffa_numbers_table_cell_storage_generated {
    #![allow(
        clippy::all,
        clippy::pedantic,
        clippy::arbitrary_source_item_ordering,
        clippy::allow_attributes_without_reason,
        clippy::module_name_repetitions,
        clippy::shadow_same,
        elided_lifetimes_in_paths,
        unreachable_pub,
        non_snake_case,
        reason = "Private Buffa generated projection."
    )]
    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-numbers-table-cell-storage/iwa_numbers_table_cell_storage_buffa_protos.rs"
    ));
}

/// Private lazy-view roots for the Keynote/format-neutral physical table-sort
/// storage seam. Repeated rows, headers, and UID arrays remain on the strict
/// source-preserving path and never cross this generated boundary.
#[doc(hidden)]
mod buffa_table_physical_sort_generated {
    #![allow(
        clippy::all,
        clippy::pedantic,
        clippy::arbitrary_source_item_ordering,
        clippy::allow_attributes_without_reason,
        clippy::module_name_repetitions,
        clippy::shadow_same,
        elided_lifetimes_in_paths,
        unreachable_pub,
        non_snake_case,
        reason = "Private Buffa generated projection."
    )]
    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-table-physical-sort/iwa_table_physical_sort_buffa_protos.rs"
    ));
}

/// Private lazy-view roots for singular Numbers popup-menu cell envelopes.
/// Repeated menu values and control-cell list entries stay on handwritten
/// source-preserving paths and never cross the generated boundary.
#[doc(hidden)]
mod buffa_numbers_table_cell_pop_up_menu_generated {
    #![allow(
        clippy::all,
        clippy::pedantic,
        clippy::arbitrary_source_item_ordering,
        clippy::allow_attributes_without_reason,
        clippy::module_name_repetitions,
        clippy::shadow_same,
        elided_lifetimes_in_paths,
        unreachable_pub,
        non_snake_case,
        reason = "Private Buffa generated projection."
    )]
    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-numbers-table-cell-pop-up-menu/iwa_numbers_table_cell_pop_up_menu_buffa_protos.rs"
    ));
}

/// Private lazy-view roots for the scalar Numbers currency-format envelope.
/// The complete `FormatStructArchive` and all unknown fields remain owned by
/// the caller's source bytes; this sidecar is parity-only.
#[doc(hidden)]
mod buffa_numbers_table_cell_currency_format_generated {
    #![allow(
        clippy::all,
        clippy::pedantic,
        clippy::arbitrary_source_item_ordering,
        clippy::allow_attributes_without_reason,
        clippy::module_name_repetitions,
        clippy::shadow_same,
        elided_lifetimes_in_paths,
        unreachable_pub,
        non_snake_case,
        reason = "buffa-build output is generated from the derived wire projection."
    )]
    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-numbers-table-cell-currency-format/iwa_numbers_table_cell_currency_format_buffa_protos.rs"
    ));
}

/// Private lazy-view roots for the scalar Numbers fraction-format envelope.
/// The handwritten decoder owns strict field policy and keeps every unknown
/// source span authoritative; this sidecar is parity-only.
#[doc(hidden)]
mod buffa_numbers_table_cell_fraction_format_generated {
    #![allow(
        clippy::all,
        clippy::pedantic,
        clippy::arbitrary_source_item_ordering,
        clippy::allow_attributes_without_reason,
        clippy::module_name_repetitions,
        clippy::shadow_same,
        elided_lifetimes_in_paths,
        unreachable_pub,
        non_snake_case,
        reason = "buffa-build output is generated from the derived wire projection."
    )]
    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-numbers-table-cell-fraction-format/iwa_numbers_table_cell_fraction_format_buffa_protos.rs"
    ));
}

/// Private lazy-view roots for the scalar Numbers Text-format envelope.
/// The handwritten decoder owns strict field policy and keeps every unknown
/// source span authoritative; this sidecar is parity-only.
#[doc(hidden)]
mod buffa_numbers_table_cell_text_format_generated {
    #![allow(
        clippy::all,
        clippy::pedantic,
        clippy::arbitrary_source_item_ordering,
        clippy::allow_attributes_without_reason,
        clippy::module_name_repetitions,
        clippy::shadow_same,
        elided_lifetimes_in_paths,
        unreachable_pub,
        non_snake_case,
        reason = "buffa-build output is generated from the derived wire projection."
    )]
    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-numbers-table-cell-text-format/iwa_numbers_table_cell_text_format_buffa_protos.rs"
    ));
}

/// Private lazy-view roots for the scalar Numbers date-and-time format
/// envelope. The handwritten decoder owns the complete source payload,
/// including unknown fields and groups; this sidecar is parity-only.
#[doc(hidden)]
mod buffa_numbers_table_cell_date_time_format_generated {
    #![allow(
        clippy::all,
        clippy::pedantic,
        clippy::arbitrary_source_item_ordering,
        clippy::allow_attributes_without_reason,
        clippy::module_name_repetitions,
        clippy::shadow_same,
        elided_lifetimes_in_paths,
        unreachable_pub,
        non_snake_case,
        reason = "buffa-build output is generated from the derived wire projection."
    )]
    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-numbers-table-cell-date-time-format/iwa_numbers_table_cell_date_time_format_buffa_protos.rs"
    ));
}

/// Private lazy-view roots for the Numbers custom-format registry and its
/// selected `FormatStructArchive` scalar fields.  Repeated registry entries
/// and nested condition payloads are represented as borrowed bytes; the
/// strict custom-format codec owns their complete validation and rewrites.
#[doc(hidden)]
mod buffa_numbers_table_cell_custom_format_generated {
    #![allow(
        clippy::all,
        clippy::pedantic,
        clippy::arbitrary_source_item_ordering,
        clippy::allow_attributes_without_reason,
        clippy::module_name_repetitions,
        clippy::shadow_same,
        elided_lifetimes_in_paths,
        unreachable_pub,
        non_snake_case,
        reason = "Private Buffa generated projection."
    )]
    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-numbers-table-cell-custom-format/iwa_numbers_table_cell_custom_format_buffa_protos.rs"
    ));
}

/// Private lazy-view roots for the scalar Numbers duration-format envelope.
/// The handwritten decoder owns the complete source payload, including
/// unknown fields and groups; this sidecar is parity-only.
#[doc(hidden)]
mod buffa_numbers_table_cell_duration_format_generated {
    #![allow(
        clippy::all,
        clippy::pedantic,
        clippy::arbitrary_source_item_ordering,
        clippy::allow_attributes_without_reason,
        clippy::module_name_repetitions,
        clippy::shadow_same,
        elided_lifetimes_in_paths,
        unreachable_pub,
        non_snake_case,
        reason = "buffa-build output is generated from the derived wire projection."
    )]
    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-numbers-table-cell-duration-format/iwa_numbers_table_cell_duration_format_buffa_protos.rs"
    ));
}

/// Private lazy-view roots for Numbers formula dependency envelopes.
#[doc(hidden)]
mod buffa_numbers_table_cell_dependency_generated {
    #![allow(
        clippy::all,
        clippy::pedantic,
        clippy::arbitrary_source_item_ordering,
        clippy::allow_attributes_without_reason,
        clippy::module_name_repetitions,
        clippy::shadow_same,
        elided_lifetimes_in_paths,
        unreachable_pub,
        non_snake_case,
        reason = "Private Buffa generated projection."
    )]
    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-numbers-table-cell-dependency/iwa_numbers_table_cell_dependency_buffa_protos.rs"
    ));
}

/// Private lazy-view roots for raw-preserving PackageMetadata publication.
#[doc(hidden)]
mod buffa_package_metadata_generated {
    #![allow(
        clippy::all,
        clippy::pedantic,
        clippy::arbitrary_source_item_ordering,
        clippy::allow_attributes_without_reason,
        clippy::module_name_repetitions,
        clippy::shadow_same,
        elided_lifetimes_in_paths,
        unreachable_pub,
        non_snake_case,
        reason = "Private Buffa generated projection."
    )]
    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-package-metadata/iwa_package_metadata_buffa_protos.rs"
    ));
}

/// Private Buffa scalar views for PackageMetadata media lifecycle records.
///
/// The generated projection intentionally contains no repeated fields.  The
/// handwritten media codec streams DataInfo and ComponentDataReference owner
/// records directly from the caller-owned source bytes and uses these views
/// only as schema/parity oracles.
#[doc(hidden)]
mod buffa_package_metadata_media_generated {
    #![allow(
        clippy::all,
        clippy::pedantic,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::shadow_same,
        elided_lifetimes_in_paths,
        unreachable_pub,
        non_snake_case,
        reason = "Private Buffa generated projection."
    )]
    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-package-metadata-media/iwa_package_metadata_media_buffa_protos.rs"
    ));
}

/// Private scalar lazy-view parity roots for the strict formula reader.
#[doc(hidden)]
mod buffa_formula_generated {
    #![allow(
        clippy::all,
        clippy::pedantic,
        clippy::arbitrary_source_item_ordering,
        clippy::allow_attributes_without_reason,
        clippy::module_name_repetitions,
        clippy::shadow_same,
        elided_lifetimes_in_paths,
        unreachable_pub,
        non_snake_case,
        reason = "Private Buffa generated projection."
    )]
    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-formula/iwa_formula_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for `TST.TableInfoArchive.table_model`.
///
/// The required drawable base archive and all unselected table metadata remain
/// caller-owned opaque source bytes. Generated code sees only the model
/// reference identifier after the Numbers adapter's strict raw preflight.
#[doc(hidden)]
mod buffa_table_info_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-table-info/iwa_table_info_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for Keynote show settings.
///
/// The repeated slide tree is deliberately absent from generated code. A
/// bounded handwritten router streams its references directly from the
/// caller-owned payload, while this projection validates direct references,
/// size, and scalar show settings.
#[doc(hidden)]
mod buffa_keynote_show_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-keynote-show/iwa_keynote_show_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for Pages section pagination.
///
/// Only the three scalar pagination fields are generated. Caller-owned section
/// bytes remain authoritative for preservation and rewriting.
#[doc(hidden)]
mod buffa_pages_section_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-pages-section/iwa_pages_section_buffa_protos.rs"
    ));
}

/// Private Buffa parity projection for a Pages section background solid fill.
///
/// Generated views are read-only. Handwritten strict routing retains and
/// rewrites caller-owned source bytes.
#[doc(hidden)]
mod buffa_pages_section_background_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-pages-section-background/iwa_pages_section_background_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for Pages root/body graph leaves.
///
/// It contains two root references and one singular section-boundary entry.
/// The repeated enclosing table and opaque base document stay in caller-owned
/// source bytes and never enter generated code.
#[doc(hidden)]
mod buffa_pages_body_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-pages-body/iwa_pages_body_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for the selected Pages media
/// discriminator. The complete MovieArchive graph and unknown source fields
/// remain caller-owned and outside generated storage.
#[doc(hidden)]
mod buffa_pages_media_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "Buffa generated projection is private implementation detail."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-pages-media/iwa_pages_media_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for the parent edge of a native
/// `TSD.DrawableArchive`. The complete drawable graph remains caller-owned;
/// only the optional nested `TSP.Reference` is projected.
#[doc(hidden)]
mod buffa_drawable_parent_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "Buffa generated output is private implementation detail."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-drawable-parent/iwa_drawable_parent_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for the Pages movie caption metadata
/// edge. The strict codec validates selected fields before forcing this view;
/// caller-owned source bytes remain the preservation representation.
#[doc(hidden)]
mod buffa_pages_movie_caption_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "Buffa generated projection is private implementation detail."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-pages-movie-caption/iwa_pages_movie_caption_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for the Pages footnote reference edge.
///
/// The strict footnote codec validates selected known fields before forcing
/// this view. Unknown source bytes remain caller-owned and are never retained
/// by the generated projection.
#[doc(hidden)]
mod buffa_pages_footnote_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-pages-footnote/iwa_pages_footnote_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for one Pages footnote marker.
///
/// The marker codec validates the selected scalar fields before forcing this
/// view. Unknown source bytes remain caller-owned and are never retained by
/// the generated projection.
#[doc(hidden)]
mod buffa_pages_footnote_marker_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build generated projection is private implementation detail."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-pages-footnote-marker/iwa_pages_footnote_marker_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for Numbers sheet, form, and table names.
#[doc(hidden)]
mod buffa_numbers_names_generated {
    #![allow(
        elided_lifetimes_in_paths,
        unreachable_pub,
        clippy::all,
        clippy::allow_attributes_without_reason,
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        non_snake_case,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "Buffa generated projection is private to this crate."
    )]
    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-numbers-names/iwa_numbers_names_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for Numbers table-header settings.
#[doc(hidden)]
mod buffa_numbers_table_header_settings_generated {
    #![allow(
        elided_lifetimes_in_paths,
        unreachable_pub,
        clippy::all,
        clippy::allow_attributes_without_reason,
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        non_snake_case,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "Buffa generated projection is private to this crate."
    )]
    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-numbers-table-header-settings/iwa_numbers_table_header_settings_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for the Numbers table sort-order
/// envelope. Repeated rules remain source-authoritative in the handwritten
/// strict codec.
#[doc(hidden)]
mod buffa_numbers_table_sort_order_generated {
    #![allow(
        elided_lifetimes_in_paths,
        unreachable_pub,
        clippy::all,
        clippy::allow_attributes_without_reason,
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        non_snake_case,
        reason = "Buffa generated projection is private to this crate."
    )]
    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-numbers-table-sort-order/iwa_numbers_table_sort_order_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for Keynote slide transitions.
///
/// It contains only the nested transition attributes and slide-node transition
/// flag.  The format crate owns strict ingress validation and retains the raw
/// IWA source for every write.
#[doc(hidden)]
mod buffa_keynote_slide_transition_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-keynote-slide-transition/iwa_keynote_slide_transition_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for Keynote slide backgrounds.
///
/// Only the singular color branch is generated; gradient and image payloads
/// stay borrowed bytes so a future fill extension cannot create an owned
/// generated graph at the format boundary.
#[doc(hidden)]
mod buffa_keynote_slide_background_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-keynote-slide-background/iwa_keynote_slide_background_buffa_protos.rs"
    ));
}

/// Internal archive-header codec implemented by the private Buffa sidecar.
///
/// This module exchanges only owned, schema-neutral archive-header DTOs.
/// Buffa-generated types and the full Prost schema remain implementation
/// details of their respective boundaries and cannot leak through this
/// focused format API.
#[doc(hidden)]
pub mod archive_codec;

/// Internal raw-text projection codec implemented by the private Buffa
/// sidecar. Generated types remain inaccessible to downstream crates.
#[doc(hidden)]
pub mod text_storage_codec;

/// Internal strict raw-preserving codec for complete TSWP hyperlink fields.
#[doc(hidden)]
pub mod hyperlink_codec;

/// Internal strict Numbers comment-storage projection. Generated types remain
/// private, and replies are streamed from caller-owned source bytes.
#[doc(hidden)]
pub mod comment_storage_codec;

/// Internal raw GroupNode category-label projection implemented by the
/// private Buffa sidecar. Generated types remain inaccessible to downstream
/// crates.
#[doc(hidden)]
pub mod group_node_category_codec;

/// Internal Keynote root-document projection implemented by the private
/// Buffa sidecar. Generated types remain inaccessible to downstream crates.
#[doc(hidden)]
pub mod keynote_document_codec;

/// Internal strict Keynote chart-caption projection. Generated types remain
/// private and caller-owned source bytes remain authoritative.
#[doc(hidden)]
pub mod keynote_chart_caption_codec;

/// Internal strict Keynote MovieArchive title/caption edge projection.
/// Generated types remain private and caller-owned source bytes remain the
/// preservation authority.
#[doc(hidden)]
pub mod keynote_movie_caption_codec;

/// Internal strict raw-preserving MovieArchive playback-settings codec.
/// Generated Buffa values remain private and the source payload is the
/// preservation authority.
#[doc(hidden)]
pub mod movie_playback_codec;

/// Internal strict raw-preserving Keynote MovieArchive geometry projection.
/// Generated Buffa values remain private and the source payload is the
/// preservation authority.
#[doc(hidden)]
pub mod keynote_movie_geometry_codec;

/// Format-neutral strict chart-caption edge projection and wire rewrite.
///
/// `keynote_chart_caption_codec` remains the compatibility spelling used by
/// the focused Keynote package; this alias owns the shared TSCH/TSD/TSP edge
/// reused by Pages and Numbers without exposing generated Buffa types.
#[doc(hidden)]
pub mod chart_caption_codec {
    pub use super::keynote_chart_caption_codec::*;
}

/// Internal generated-free encoder for one canonical inline Keynote chart
/// caption graph. Generated protobuf types remain test-only or private to
/// their existing owners.
#[doc(hidden)]
pub mod keynote_chart_caption_graph_codec;

/// Internal strict Keynote chart-title generated-extension projection.
/// Generated types remain private and caller-owned source bytes remain the
/// preservation authority.
#[doc(hidden)]
pub mod keynote_chart_title_codec;

/// Internal strict Keynote chart-legend visibility projection. Generated
/// types remain private and source bytes remain the preservation authority.
#[doc(hidden)]
pub mod keynote_chart_legend_codec;

/// Internal strict Keynote chart Arrange-state projection. Generated types
/// remain private and caller-owned source bytes remain the preservation
/// authority.
#[doc(hidden)]
pub mod keynote_chart_arrangement_codec;

/// Internal strict Keynote chart-axis-title generated-extension projection.
/// Generated types remain private and caller-owned source bytes remain the
/// preservation authority.
#[doc(hidden)]
pub mod keynote_chart_axis_title_codec;

/// Internal strict Keynote value-axis settings projection. Generated types
/// remain private and caller-owned source bytes remain the preservation
/// authority.
#[doc(hidden)]
pub mod keynote_chart_axis_value_settings_codec;

/// Internal strict Keynote placeholder text-owner projection. Generated types
/// remain inaccessible and source bytes stay authoritative.
#[doc(hidden)]
pub mod keynote_placeholder_text_codec;

/// Internal strict Keynote speaker-note owner projection. Generated types
/// remain inaccessible to downstream crates and source bytes stay authoritative.
#[doc(hidden)]
pub mod keynote_speaker_notes_codec;

/// Internal strict Keynote slide-number projection. Generated views remain
/// private; the attachment table is borrowed raw source bytes.
#[doc(hidden)]
pub mod keynote_slide_number_codec;

/// Internal strict Keynote soundtrack-settings projection.
#[doc(hidden)]
pub mod keynote_soundtrack_settings_codec;

/// Internal strict Keynote movie-media data-reference projection. Generated
/// values remain private and caller-owned source bytes remain authoritative.
#[doc(hidden)]
pub mod keynote_media_codec;

/// Internal Numbers TableInfo model-reference projection implemented by a
/// private strict Buffa lazy-view sidecar. Generated types remain inaccessible
/// to downstream crates.
#[doc(hidden)]
pub mod table_info_codec;

/// Internal Keynote show projection and bounded streaming slide-tree codec.
/// Generated types remain inaccessible to downstream crates.
#[doc(hidden)]
pub mod keynote_show_codec;

/// Internal Pages section-pagination projection implemented by a private
/// Buffa lazy-view sidecar. Generated types remain inaccessible downstream.
#[doc(hidden)]
pub mod pages_section_codec;

/// Internal strict raw-preserving Pages section-background codec.
#[doc(hidden)]
pub mod pages_section_background_codec;

/// Internal Pages root/body projection implemented by a private strict Buffa
/// lazy-view adapter. Generated types remain inaccessible downstream.
#[doc(hidden)]
pub mod pages_body_codec;

/// Internal strict Pages media discriminator projection. Generated types
/// remain private and caller-owned raw bytes remain authoritative.
#[doc(hidden)]
pub mod pages_media_codec;

/// Internal strict TSD drawable-parent projection. Generated types remain
/// private and the caller-owned drawable payload remains authoritative.
#[doc(hidden)]
pub mod drawable_parent_codec;

/// Internal strict Pages movie-caption metadata projection. Generated types
/// remain private and caller-owned raw bytes remain authoritative.
#[doc(hidden)]
pub mod pages_movie_caption_codec;

/// Internal strict Pages footnote-reference projection. Generated types
/// remain private and caller-owned raw bytes remain authoritative.
#[doc(hidden)]
pub mod pages_footnote_codec;

/// Internal strict raw-preserving Pages footnote-marker projection. Generated
/// types remain private and caller-owned source bytes remain authoritative.
#[doc(hidden)]
pub mod pages_footnote_marker_codec;

/// Internal generated-free Pages body-footnote graph creation and table
/// rewrite codec.  It composes the focused body, footnote-reference, marker,
/// and text-storage projections without exposing generated protobuf values.
#[doc(hidden)]
pub mod pages_footnote_graph_codec;

/// Internal strict Pages document page-layout projection. Generated types stay
/// private and caller-owned raw bytes remain authoritative.
#[doc(hidden)]
pub mod pages_page_layout_codec;

/// Internal strict Pages document-settings projection. Generated types remain
/// private and caller-owned raw bytes stay authoritative.
#[doc(hidden)]
pub mod pages_document_settings_codec;

/// Internal strict Numbers names projection. Generated types remain private
/// and all decoded names borrow caller-owned source bytes.
#[doc(hidden)]
pub mod numbers_names_codec;

/// Format-neutral strict borrowed table-model discovery projection.
///
/// The snapshot contains only required identity/name/dimension facts; the
/// complete generated TableModel and all storage/style fields remain outside
/// this boundary.
#[doc(hidden)]
pub mod table_model_discovery_codec;

#[doc(hidden)]
pub mod numbers_sheet_order_codec;

/// Internal strict Pages drawable-order projection. Generated scalar views
/// remain private and the repeated TP archive stays source-authoritative.
#[doc(hidden)]
pub mod pages_drawable_order_codec;
#[doc(hidden)]
pub mod pages_header_footer_codec;

/// Internal strict Numbers table-header settings projection. Generated types
/// remain private and scalar facts borrow only caller-owned source bytes.
#[doc(hidden)]
pub mod numbers_table_header_settings_codec;

/// Format-neutral strict table-header settings projection and wire rewrite.
///
/// `numbers_table_header_settings_codec` remains the compatibility spelling
/// used by the Numbers package; this alias is the seam for other iWork
/// adapters and exposes no generated Buffa types.
#[doc(hidden)]
pub mod table_header_settings_codec {
    pub use super::numbers_table_header_settings_codec::*;
}

/// Internal strict Numbers table-title projection. Generated types remain
/// private and caller-owned source bytes remain the rewrite authority.
#[doc(hidden)]
pub mod numbers_table_title_codec;

/// Strict generated-free Numbers table-sort-order projection and rewrite.
#[doc(hidden)]
pub mod numbers_table_sort_order_codec;

/// Neutral spelling for the persisted Numbers table-sort-order seam.
#[doc(hidden)]
pub mod table_sort_order_codec {
    pub use super::numbers_table_sort_order_codec::*;
}

/// Strict generated-free Numbers table-cell storage projection.
#[doc(hidden)]
pub mod numbers_table_cell_storage_codec;

/// Strict generated-free Numbers control-cell popup-menu projection.
#[doc(hidden)]
pub mod numbers_table_cell_pop_up_menu_codec;

/// Neutral strict Numbers checkbox/star-rating/slider/stepper CellSpec and
/// FormatStruct projection.  The implementation aliases the audited popup
/// parser/writer so control packages cannot accidentally create a second wire
/// policy or bypass prepared resource limits.
#[doc(hidden)]
pub mod numbers_table_cell_control_codec;

/// Neutral strict plain-number `FormatStructArchive` projection and measured
/// source-preserving rewrite. Generated Buffa values remain private to the
/// shared table-cell implementation.
#[doc(hidden)]
pub mod numbers_table_cell_number_format_codec;

/// Neutral strict plain-percentage `FormatStructArchive` projection and
/// measured source-preserving rewrite. The wire implementation reuses the
/// shared table-cell decimal-format core and the existing scalar Buffa view.
#[doc(hidden)]
pub mod numbers_table_cell_percentage_format_codec;

/// Neutral strict native-currency `FormatStructArchive` projection and
/// measured source-preserving rewrite. Generated Buffa values remain private
/// to the shared table-cell implementation.
#[doc(hidden)]
pub mod numbers_table_cell_currency_format_codec;

/// Neutral strict native-fraction `FormatStructArchive` projection and
/// measured source-preserving rewrite. Generated Buffa values remain private
/// to the shared table-cell implementation.
#[doc(hidden)]
pub mod numbers_table_cell_fraction_format_codec;

/// Neutral strict native-scientific `FormatStructArchive` projection and
/// measured source-preserving rewrite. The four scalar fields use the shared
/// private Number/Percentage Buffa projection.
#[doc(hidden)]
pub mod numbers_table_cell_scientific_format_codec;

/// Neutral strict native-Text `FormatStructArchive` projection and measured
/// source-preserving rewrite. Generated Buffa values remain private to the
/// shared table-cell implementation.
#[doc(hidden)]
pub mod numbers_table_cell_text_format_codec;

/// Neutral strict native date-and-time `FormatStructArchive` projection and
/// measured source-preserving rewrite. Generated Buffa values remain private
/// to the shared table-cell implementation.
#[doc(hidden)]
pub mod numbers_table_cell_date_time_format_codec;

/// Strict source-preserving native Numbers custom-format registry codec.
/// Generated lazy views remain private to this crate; callers receive only
/// source-borrowed snapshots and measured rewrite plans.
#[doc(hidden)]
pub mod numbers_table_cell_custom_format_codec;

/// Strict source-preserving native Numbers Duration format codec.
#[doc(hidden)]
pub mod numbers_table_cell_duration_format_codec;

/// Format-neutral spelling for the Numbers interactive-cell control seam.
#[doc(hidden)]
pub mod table_cell_control_codec {
    pub use super::numbers_table_cell_control_codec::*;
}

/// Format-neutral strict table-dimension/header-bucket seam.
///
/// The implementation remains in [`numbers_table_cell_storage_codec`] so
/// Numbers' existing storage reader and attached-table adapters continue to
/// share one source-authoritative parser. This alias deliberately exposes
/// only the dimension-relevant borrowed snapshots, visitor, and plan/execute
/// APIs; generated repeated storage never crosses the protos boundary.
#[doc(hidden)]
pub mod table_dimension_codec {
    pub use super::numbers_table_cell_storage_codec::{
        DataStoreSnapshot, DecodeError, DecodeLimit, DecodeOptions, DecodeReport,
        DecodeResourceUpperBound, HeaderRecord, HeaderSizeEdit, HeaderSizeRewritePlan,
        HeaderSizeRewriteReport, HeaderSizeRewriteRequirements, HeaderSnapshot,
        HeaderStorageBucketSnapshot, HeaderStorageSnapshot, ReferenceRecord, ReferenceSnapshot,
        StorageVisitor, TableModelSnapshot, decode_data_store, decode_data_store_with_report,
        decode_data_store_with_visitor, decode_header, decode_header_storage,
        decode_header_storage_bucket, decode_header_storage_bucket_with_report,
        decode_header_storage_bucket_with_visitor, decode_header_storage_with_report,
        decode_header_storage_with_visitor, decode_header_with_report, decode_table_model,
        decode_table_model_with_report, decode_table_model_with_visitor,
        execute_header_storage_bucket_size_plan, plan_header_storage_bucket_sizes,
        rewrite_header_storage_bucket_sizes,
    };
}

/// Format-neutral strict table-appearance projection and source-preserving
/// rewrite. Generated schema types remain private to this crate.
#[doc(hidden)]
pub mod table_appearance_codec;

/// Strict generated-free Numbers table-cell dependency/cache projection.
#[doc(hidden)]
pub mod numbers_table_cell_dependency_codec;

/// Strict generated-free physical table-sort storage projection and
/// source-preserving repeated-field rewrites.
#[doc(hidden)]
pub mod numbers_table_physical_sort_codec;

/// Format-neutral spelling for the physical table-sort storage seam.
#[doc(hidden)]
pub mod table_physical_sort_codec {
    pub use super::numbers_table_physical_sort_codec::*;
}

/// Keynote spelling retained for the concrete format owner.
#[doc(hidden)]
pub mod keynote_table_physical_sort_codec {
    pub use super::numbers_table_physical_sort_codec::*;
}

/// Strict raw-preserving PackageMetadata sparse-publication codec.
#[doc(hidden)]
pub mod package_metadata_codec;

/// Strict, lazy, raw-preserving PackageMetadata DataInfo and media-owner
/// lifecycle codec. Generated Buffa values remain private to this crate.
#[doc(hidden)]
pub mod package_metadata_media_codec;

/// Strict generated-free streaming reader for table-local scalar formulas.
#[doc(hidden)]
pub mod numbers_formula_codec;

/// Internal Keynote slide-transition projection and opaque color/path
/// validator implemented around a private Buffa lazy-view sidecar. Generated
/// types remain inaccessible downstream.
#[doc(hidden)]
pub mod keynote_slide_transition_codec;

/// Internal strict raw-preserving Keynote slide-background codec. Generated
/// types remain private and all unrecognized fill payloads stay borrowed.
#[doc(hidden)]
pub mod keynote_slide_background_codec;

// Keep only schema modules with current workspace consumers at the crate root.
// The generated collaboration/change-set modules (`*_sos`) and `tsck` remain
// available to the private include above for wire-schema completeness, but do
// not form part of the supported raw-protobuf surface.
pub use generated::{kn, tn, tp, tsa, tsce, tsch, tsd, tsk, tsp, tss, tst, tswp};

#[cfg(test)]
mod tests {
    use buffa::{LazyMessageView as _, Message as _};
    use prost::Message as _;

    #[test]
    fn generated_messages_round_trip_without_runtime_names() -> Result<(), prost::DecodeError> {
        let input = super::tsp::ArchiveInfo {
            identifier: Some(42),
            message_infos: Vec::new(),
            should_merge: Some(true),
        };
        let encoded = input.encode_to_vec();
        let decoded = super::tsp::ArchiveInfo::decode(encoded.as_slice())?;
        assert_eq!(decoded, input);
        Ok(())
    }

    #[test]
    fn buffa_archive_info_matches_prost_wire_format() -> Result<(), Box<dyn std::error::Error>> {
        let input = super::tsp::ArchiveInfo {
            identifier: Some(42),
            message_infos: Vec::new(),
            should_merge: Some(true),
        };
        let prost_encoded = input.encode_to_vec();

        let buffa_decoded =
            super::buffa_archive_header_generated::LitchiIwaArchiveHeaderProjection::ArchiveInfo::decode_from_slice(&prost_encoded)?;
        assert_eq!(buffa_decoded.identifier, input.identifier);
        assert!(buffa_decoded.message_infos.is_empty());
        assert_eq!(buffa_decoded.should_merge, input.should_merge);

        let buffa_encoded = buffa_decoded.try_encode_to_vec()?;
        let prost_decoded = super::tsp::ArchiveInfo::decode(buffa_encoded.as_slice())?;
        assert_eq!(prost_decoded, input);
        Ok(())
    }

    #[test]
    fn buffa_archive_info_lazy_view_round_trips() -> Result<(), Box<dyn std::error::Error>> {
        let input = super::buffa_archive_header_generated::LitchiIwaArchiveHeaderProjection::ArchiveInfo {
            identifier: Some(42),
            message_infos: vec![super::buffa_archive_header_generated::LitchiIwaArchiveHeaderProjection::MessageInfo {
                r#type: 7,
                length: 11,
                ..Default::default()
            }],
            should_merge: Some(true),
            ..Default::default()
        };
        let encoded = input.try_encode_to_vec()?;
        let lazy: super::buffa_archive_header_generated::LitchiIwaArchiveHeaderProjection::ArchiveInfoLazyView<'_> =
            buffa::DecodeOptions::new().decode_lazy_view(&encoded)?;

        assert_eq!(lazy.message_infos.len(), 1);
        let message_info_view = lazy.message_infos.try_get(0)?;
        assert_eq!(
            message_info_view.map(|view| (view.r#type, view.length)),
            Some((7, 11))
        );
        assert_eq!(lazy.to_owned_message()?, input);
        assert_eq!(lazy.try_encode_to_vec()?, encoded);
        Ok(())
    }
}
