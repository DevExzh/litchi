#!/usr/bin/env python3
"""Reject workspace dependency edges that violate the accepted crate topology."""

from __future__ import annotations


import sys
from pathlib import Path

if not __package__:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from tools.crate_boundaries import common, keynote, numbers, pages
from tools.crate_boundaries.common import (
    PolicyError,
    audit_annotation_author_codec_source_topology,
    audit_chart_arrangement_codec_source_topology,
    audit_iwa_chart_data_source_topology,
    audit_iwa_chart_data_write_source_topology,
    audit_iwa_chart_grid_creation_source_topology,
    audit_iwa_chart_metadata_source_topology,
    audit_iwa_common_formula_render_ownership,
    audit_iwa_common_table_merge_source_topology,
    audit_iwa_common_table_read_source_topology,
    audit_iwa_drawable_container_codec_source_topology,
    audit_iwa_package_metadata_component_removal_source_topology,
    audit_iwa_protos_root_alias_source_topology,
    audit_iwa_retired_host_seams,
    audit_iwa_shared_image_adjustments_source_topology,
    audit_iwa_shared_media_playback_source_topology,
    audit_iwa_table_cell_borders_source_topology,
    audit_iwa_table_merge_formula_codec_source_topology,
    audit_iwork_atomic_publication,
    audit_iwork_example_source_topology,
    audit_litchi_semantic_facade_source_topology,
    audit_manifest_dependency_inventory,
    audit_manifest_inventory,
    audit_neutral_table_model_source_topology,
    audit_snapshot,
    audit_spreadsheet_chart_source_topology,
    audit_spreadsheet_shape_source_topology,
    audit_spreadsheet_sheet_view_source_topology,
    audit_xlsb_source_topology,
    cargo_metadata,
    debt_report,
    load_policy,
    parse_args,
    snapshot_from_metadata,
)
from tools.crate_boundaries.keynote import (
    audit_keynote_audio_position_codec_source_topology,
    audit_keynote_audio_position_completion_source_topology,
    audit_keynote_audio_position_facade_source_topology,
    audit_keynote_audio_position_resource_source_topology,
    audit_keynote_chart_arrangement_facade_source_topology,
    audit_keynote_chart_axis_title_completion_source_topology,
    audit_keynote_chart_axis_title_facade_source_topology,
    audit_keynote_chart_axis_title_legacy_calls,
    audit_keynote_chart_axis_title_resource_source_topology,
    audit_keynote_chart_axis_value_settings_completion_source_topology,
    audit_keynote_chart_axis_value_settings_facade_source_topology,
    audit_keynote_chart_axis_value_settings_legacy_calls,
    audit_keynote_chart_axis_value_settings_resource_source_topology,
    audit_keynote_chart_caption_facade_source_topology,
    audit_keynote_chart_caption_legacy_calls,
    audit_keynote_chart_legend_visibility_facade_source_topology,
    audit_keynote_chart_title_legacy_calls,
    audit_keynote_document_public_api,
    audit_keynote_media_creation_codec_source_topology,
    audit_keynote_media_properties_codec_source_topology,
    audit_keynote_media_properties_facade_source_topology,
    audit_keynote_media_properties_resource_source_topology,
    audit_keynote_movie_caption_facade_source_topology,
    audit_keynote_movie_caption_lifecycle_source_topology,
    audit_keynote_movie_geometry_completion_source_topology,
    audit_keynote_movie_geometry_facade_source_topology,
    audit_keynote_movie_geometry_resource_source_topology,
    audit_keynote_movie_playback_facade_source_topology,
    audit_keynote_movie_playback_resource_source_topology,
    audit_keynote_movie_title_codec_source_topology,
    audit_keynote_movie_title_facade_source_topology,
    audit_keynote_movie_title_legacy_calls,
    audit_keynote_movie_title_lifecycle_source_topology,
    audit_keynote_package_no_eager_prost_source_topology,
    audit_keynote_placeholder_visibility_facade_source_topology,
    audit_keynote_show_settings_facade_source_topology,
    audit_keynote_slide_audio_creation_source_topology,
    audit_keynote_slide_background_facade_source_topology,
    audit_keynote_slide_delete_facade_source_topology,
    audit_keynote_slide_drawable_comment_source_topology,
    audit_keynote_slide_media_data_codec_source_topology,
    audit_keynote_slide_media_data_core_boundary_source_topology,
    audit_keynote_slide_media_data_facade_source_topology,
    audit_keynote_slide_media_data_metadata_boundary_source_topology,
    audit_keynote_slide_media_data_resource_source_topology,
    audit_keynote_slide_media_data_transaction_source_topology,
    audit_keynote_slide_media_lifecycle_codec_source_topology,
    audit_keynote_slide_media_lifecycle_facade_source_topology,
    audit_keynote_slide_media_lifecycle_transaction_source_topology,
    audit_keynote_slide_movie_creation_source_topology,
    audit_keynote_slide_table_appearance_facade_source_topology,
    audit_keynote_slide_table_appearance_resource_source_topology,
    audit_keynote_slide_table_dimension_facade_source_topology,
    audit_keynote_slide_table_headers_facade_source_topology,
    audit_keynote_slide_table_headers_resource_source_topology,
    audit_keynote_slide_table_lock_state_facade_source_topology,
    audit_keynote_slide_table_name_facade_source_topology,
    audit_keynote_slide_table_name_resource_source_topology,
    audit_keynote_slide_table_physical_sort_facade_source_topology,
    audit_keynote_slide_table_physical_sort_resource_source_topology,
    audit_keynote_slide_table_sort_facade_source_topology,
    audit_keynote_slide_table_sort_resource_source_topology,
    audit_keynote_slide_table_title_facade_source_topology,
    audit_keynote_slide_table_title_resource_source_topology,
    audit_keynote_slide_transition_facade_source_topology,
    audit_keynote_soundtrack_items_facade_source_topology,
    audit_keynote_soundtrack_settings_facade_source_topology,
    audit_keynote_table_cells_source_topology,
    audit_keynote_table_merge_reader_source_topology,
    audit_keynote_table_merge_source_topology,
    audit_keynote_table_merge_transaction_source_topology,
)
from tools.crate_boundaries.numbers import (
    audit_iwa_numbers_table_cell_dependency_inbound_source_topology,
    audit_iwa_numbers_table_cell_dependency_removal_source_topology,
    audit_iwa_numbers_wire_cell_value_ownership,
    audit_iwa_numbers_wire_formula_envelope_ownership,
    audit_iwa_numbers_wire_formula_names_source_topology,
    audit_iwa_numbers_wire_formula_render_ownership,
    audit_iwa_numbers_wire_table_cells_source_topology,
    audit_iwa_numbers_wire_table_data_list_ownership,
    audit_iwa_numbers_wire_table_merge_writer_source_topology,
    audit_iwa_numbers_wire_table_merges_source_topology,
    audit_iwa_numbers_wire_table_sidecars_source_topology,
    audit_numbers_chart_arrangement_owner_source_topology,
    audit_numbers_comment_clear_metadata_prerequisite_source_topology,
    audit_numbers_document_public_api,
    audit_numbers_extractor_no_eager_comment_storage_source_topology,
    audit_numbers_extractor_no_eager_formula_source_topology,
    audit_numbers_extractor_no_eager_rich_text_source_topology,
    audit_numbers_extractor_no_eager_table_data_list_source_topology,
    audit_numbers_extractor_no_eager_tile_source_topology,
    audit_numbers_formula_facade_source_topology,
    audit_numbers_identity_boundary_source_topology,
    audit_numbers_names_facade_source_topology,
    audit_numbers_names_package_no_eager_prost_source_topology,
    audit_numbers_names_save_token_publication_source_topology,
    audit_numbers_package_no_eager_prost_source_topology,
    audit_numbers_sheet_order_facade_source_topology,
    audit_numbers_table_appearance_facade_source_topology,
    audit_numbers_table_cell_comment_reply_codec_rewrite_source_topology,
    audit_numbers_table_cell_comment_reply_facade_source_topology,
    audit_numbers_table_cell_comment_reply_mutation_facade_source_topology,
    audit_numbers_table_cell_control_facade_source_topology,
    audit_numbers_table_cell_control_popup_split_lifecycle_source_topology,
    audit_numbers_table_cell_currency_format_codec_source_topology,
    audit_numbers_table_cell_currency_format_facade_source_topology,
    audit_numbers_table_cell_custom_format_codec_source_topology,
    audit_numbers_table_cell_custom_format_facade_source_topology,
    audit_numbers_table_cell_date_time_format_codec_source_topology,
    audit_numbers_table_cell_date_time_format_facade_source_topology,
    audit_numbers_table_cell_display_format_shared_core_source_topology,
    audit_numbers_table_cell_duration_format_codec_source_topology,
    audit_numbers_table_cell_duration_format_facade_source_topology,
    audit_numbers_table_cell_fraction_format_codec_source_topology,
    audit_numbers_table_cell_fraction_format_facade_source_topology,
    audit_numbers_table_cell_number_format_codec_source_topology,
    audit_numbers_table_cell_number_format_facade_source_topology,
    audit_numbers_table_cell_percentage_format_codec_source_topology,
    audit_numbers_table_cell_percentage_format_facade_source_topology,
    audit_numbers_table_cell_pop_up_menu_facade_source_topology,
    audit_numbers_table_cell_scientific_format_codec_source_topology,
    audit_numbers_table_cell_scientific_format_facade_source_topology,
    audit_numbers_table_cell_text_format_codec_source_topology,
    audit_numbers_table_cell_text_format_facade_source_topology,
    audit_numbers_table_cells_facade_source_topology,
    audit_numbers_table_cells_mutation_facade_source_topology,
    audit_numbers_table_dimension_facade_source_topology,
    audit_numbers_table_header_settings_facade_source_topology,
    audit_numbers_table_lock_facade_source_topology,
    audit_numbers_table_merge_source_topology,
    audit_numbers_table_merge_transaction_source_topology,
    audit_numbers_table_move_facade_source_topology,
    audit_numbers_table_sort_facade_source_topology,
    audit_numbers_table_title_settings_facade_source_topology,
)
from tools.crate_boundaries.pages import (
    audit_pages_body_chart_arrangement_owner_source_topology,
    audit_pages_body_table_deletion_source_topology,
    audit_pages_document_public_api,
    audit_pages_document_settings_facade_source_topology,
    audit_pages_footnote_lifecycle_facade_source_topology,
    audit_pages_footnote_text_facade_source_topology,
    audit_pages_header_footer_facade_source_topology,
    audit_pages_package_no_eager_prost_source_topology,
    audit_pages_package_output_api_source_topology,
    audit_pages_page_layout_facade_source_topology,
    audit_pages_section_background_facade_source_topology,
    audit_pages_section_settings_facade_source_topology,
    audit_pages_section_text_facade_source_topology,
    audit_pages_table_appearance_facade_source_topology,
    audit_pages_table_appearance_resource_source_topology,
    audit_pages_table_cells_source_topology,
    audit_pages_table_dimension_facade_source_topology,
    audit_pages_table_headers_facade_source_topology,
    audit_pages_table_hidden_axes_facade_source_topology,
    audit_pages_table_lock_facade_source_topology,
    audit_pages_table_merge_reader_source_topology,
    audit_pages_table_merge_source_topology,
    audit_pages_table_merge_transaction_source_topology,
    audit_pages_table_name_facade_source_topology,
    audit_pages_table_sort_facade_source_topology,
    audit_pages_table_title_facade_source_topology,
    audit_pages_transaction_family_source_topology,
)

def __getattr__(name: str):
    """Expose the checker rules to callers of the original command module."""
    for module in (common, keynote, numbers, pages):
        if hasattr(module, name):
            return getattr(module, name)
    raise AttributeError(name)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    try:
        policy = load_policy(args.policy)
    except PolicyError as error:
        print(f"crate-boundary policy error: {error}", file=sys.stderr)
        return 2

    snapshot = snapshot_from_metadata(cargo_metadata())
    violations = (
        audit_manifest_inventory(snapshot)
        + audit_manifest_dependency_inventory(policy=policy)
        + audit_snapshot(snapshot, policy)
        + audit_iwa_protos_root_alias_source_topology()
        + audit_litchi_semantic_facade_source_topology()
        + audit_pages_transaction_family_source_topology()
        + audit_iwork_example_source_topology()
        + audit_keynote_show_settings_facade_source_topology()
        + audit_keynote_soundtrack_settings_facade_source_topology()
        + audit_keynote_soundtrack_items_facade_source_topology()
        + audit_keynote_slide_transition_facade_source_topology()
        + audit_keynote_slide_background_facade_source_topology()
        + audit_keynote_slide_delete_facade_source_topology()
        + audit_keynote_placeholder_visibility_facade_source_topology()
        + audit_keynote_package_no_eager_prost_source_topology()
        + audit_keynote_chart_title_legacy_calls()
        + audit_keynote_chart_legend_visibility_facade_source_topology()
        + audit_keynote_chart_arrangement_facade_source_topology()
        + audit_chart_arrangement_codec_source_topology()
        + audit_numbers_chart_arrangement_owner_source_topology()
        + audit_pages_body_chart_arrangement_owner_source_topology()
        + audit_iwa_chart_metadata_source_topology()
        + audit_iwa_chart_data_source_topology()
        + audit_iwa_chart_data_write_source_topology()
        + audit_iwa_chart_grid_creation_source_topology()
        + audit_keynote_chart_axis_title_legacy_calls()
        + audit_keynote_chart_axis_title_facade_source_topology()
        + audit_keynote_chart_axis_title_resource_source_topology()
        + audit_keynote_chart_axis_title_completion_source_topology()
        + audit_keynote_chart_axis_value_settings_legacy_calls()
        + audit_keynote_chart_axis_value_settings_facade_source_topology()
        + audit_keynote_chart_axis_value_settings_resource_source_topology()
        + audit_keynote_chart_axis_value_settings_completion_source_topology()
        + audit_keynote_movie_title_legacy_calls()
        + audit_keynote_movie_title_facade_source_topology()
        + audit_keynote_movie_title_codec_source_topology()
        + audit_keynote_movie_title_lifecycle_source_topology()
        + audit_keynote_chart_caption_legacy_calls()
        + audit_keynote_chart_caption_facade_source_topology()
        + audit_keynote_movie_caption_facade_source_topology()
        + audit_keynote_movie_caption_lifecycle_source_topology()
        + audit_keynote_movie_playback_facade_source_topology()
        + audit_keynote_movie_playback_resource_source_topology()
        + audit_keynote_movie_geometry_facade_source_topology()
        + audit_keynote_movie_geometry_resource_source_topology()
        + audit_keynote_movie_geometry_completion_source_topology()
        + audit_keynote_audio_position_facade_source_topology()
        + audit_keynote_audio_position_codec_source_topology()
        + audit_keynote_audio_position_completion_source_topology()
        + audit_keynote_audio_position_resource_source_topology()
        + audit_keynote_media_properties_facade_source_topology()
        + audit_keynote_media_properties_codec_source_topology()
        + audit_keynote_media_properties_resource_source_topology()
        + audit_keynote_slide_media_data_facade_source_topology()
        + audit_keynote_slide_media_data_codec_source_topology()
        + audit_keynote_slide_media_data_core_boundary_source_topology()
        + audit_keynote_slide_media_data_metadata_boundary_source_topology()
        + audit_keynote_slide_media_data_resource_source_topology()
        + audit_keynote_slide_media_data_transaction_source_topology()
        + audit_keynote_slide_media_lifecycle_facade_source_topology()
        + audit_keynote_slide_media_lifecycle_codec_source_topology()
        + audit_keynote_slide_media_lifecycle_transaction_source_topology()
        + audit_annotation_author_codec_source_topology()
        + audit_keynote_slide_drawable_comment_source_topology()
        + audit_keynote_slide_table_title_facade_source_topology()
        + audit_keynote_slide_table_title_resource_source_topology()
        + audit_keynote_slide_table_sort_facade_source_topology()
        + audit_keynote_slide_table_sort_resource_source_topology()
        + audit_keynote_slide_table_physical_sort_facade_source_topology()
        + audit_keynote_slide_table_physical_sort_resource_source_topology()
        + audit_keynote_slide_table_dimension_facade_source_topology()
        + audit_keynote_slide_table_name_facade_source_topology()
        + audit_keynote_slide_table_name_resource_source_topology()
        + audit_keynote_slide_table_lock_state_facade_source_topology()
        + audit_keynote_slide_table_headers_facade_source_topology()
        + audit_keynote_slide_table_headers_resource_source_topology()
        + audit_keynote_slide_table_appearance_facade_source_topology()
        + audit_keynote_slide_table_appearance_resource_source_topology()
        + audit_keynote_document_public_api()
        + audit_numbers_identity_boundary_source_topology()
        + audit_numbers_package_no_eager_prost_source_topology()
        + audit_numbers_extractor_no_eager_rich_text_source_topology()
        + audit_numbers_extractor_no_eager_tile_source_topology()
        + audit_iwa_numbers_wire_formula_render_ownership()
        + audit_iwa_numbers_wire_cell_value_ownership()
        + audit_iwa_numbers_wire_table_data_list_ownership()
        + audit_iwa_numbers_wire_table_cells_source_topology()
        + audit_iwa_numbers_wire_table_sidecars_source_topology()
        + audit_iwa_numbers_wire_formula_names_source_topology()
        + audit_iwa_numbers_wire_formula_envelope_ownership()
        + audit_iwa_common_formula_render_ownership()
        + audit_iwa_common_table_read_source_topology()
        + audit_iwa_shared_media_playback_source_topology()
        + audit_iwa_shared_image_adjustments_source_topology()
        + audit_iwa_retired_host_seams()
        + audit_numbers_extractor_no_eager_table_data_list_source_topology()
        + audit_numbers_comment_clear_metadata_prerequisite_source_topology()
        + audit_numbers_table_cell_comment_reply_facade_source_topology()
        + audit_numbers_table_cell_comment_reply_codec_rewrite_source_topology()
        + audit_numbers_table_cell_comment_reply_mutation_facade_source_topology()
        + audit_numbers_extractor_no_eager_comment_storage_source_topology()
        + audit_numbers_extractor_no_eager_formula_source_topology()
        + audit_numbers_names_package_no_eager_prost_source_topology()
        + audit_numbers_names_save_token_publication_source_topology()
        + audit_numbers_names_facade_source_topology()
        + audit_numbers_sheet_order_facade_source_topology()
        + audit_numbers_table_sort_facade_source_topology()
        + audit_numbers_table_move_facade_source_topology()
        + audit_numbers_table_header_settings_facade_source_topology()
        + audit_numbers_table_title_settings_facade_source_topology()
        + audit_numbers_table_appearance_facade_source_topology()
        + audit_numbers_table_cell_control_facade_source_topology()
        + audit_numbers_table_cell_control_popup_split_lifecycle_source_topology()
        + audit_numbers_table_cell_pop_up_menu_facade_source_topology()
        + audit_numbers_table_cell_number_format_codec_source_topology()
        + audit_numbers_table_cell_number_format_facade_source_topology()
        + audit_numbers_table_cell_text_format_codec_source_topology()
        + audit_numbers_table_cell_text_format_facade_source_topology()
        + audit_numbers_table_cell_date_time_format_codec_source_topology()
        + audit_numbers_table_cell_date_time_format_facade_source_topology()
        + audit_numbers_table_cell_duration_format_codec_source_topology()
        + audit_numbers_table_cell_duration_format_facade_source_topology()
        + audit_numbers_table_cell_custom_format_codec_source_topology()
        + audit_numbers_table_cell_custom_format_facade_source_topology()
        + audit_numbers_table_cell_percentage_format_codec_source_topology()
        + audit_numbers_table_cell_percentage_format_facade_source_topology()
        + audit_numbers_table_cell_currency_format_codec_source_topology()
        + audit_numbers_table_cell_currency_format_facade_source_topology()
        + audit_numbers_table_cell_scientific_format_codec_source_topology()
        + audit_numbers_table_cell_scientific_format_facade_source_topology()
        + audit_numbers_table_cell_fraction_format_codec_source_topology()
        + audit_numbers_table_cell_fraction_format_facade_source_topology()
        + audit_numbers_table_cell_display_format_shared_core_source_topology()
        + audit_numbers_table_dimension_facade_source_topology()
        + audit_numbers_formula_facade_source_topology()
        + audit_numbers_table_cells_facade_source_topology()
        + audit_numbers_table_cells_mutation_facade_source_topology()
        + audit_numbers_table_lock_facade_source_topology()
        + audit_numbers_document_public_api()
        + audit_iwa_drawable_container_codec_source_topology()
        + audit_keynote_media_creation_codec_source_topology()
        + audit_keynote_slide_audio_creation_source_topology()
        + audit_keynote_slide_movie_creation_source_topology()
        + audit_pages_document_public_api()
        + audit_pages_package_output_api_source_topology()
        + audit_iwork_atomic_publication()
        + audit_pages_package_no_eager_prost_source_topology()
        + audit_pages_footnote_text_facade_source_topology()
        + audit_pages_footnote_lifecycle_facade_source_topology()
        + audit_pages_page_layout_facade_source_topology()
        + audit_pages_table_lock_facade_source_topology()
        + audit_pages_table_title_facade_source_topology()
        + audit_pages_table_name_facade_source_topology()
        + audit_pages_table_headers_facade_source_topology()
        + audit_pages_table_dimension_facade_source_topology()
        + audit_pages_table_hidden_axes_facade_source_topology()
        + audit_pages_table_sort_facade_source_topology()
        + audit_pages_table_appearance_facade_source_topology()
        + audit_pages_table_appearance_resource_source_topology()
        + audit_pages_header_footer_facade_source_topology()
        + audit_pages_section_text_facade_source_topology()
        + audit_pages_document_settings_facade_source_topology()
        + audit_pages_section_settings_facade_source_topology()
        + audit_pages_section_background_facade_source_topology()
        + audit_iwa_table_cell_borders_source_topology()
        + audit_neutral_table_model_source_topology()
        + audit_iwa_common_table_merge_source_topology()
        + audit_iwa_numbers_wire_table_merges_source_topology()
        + audit_pages_table_merge_source_topology()
        + audit_pages_table_merge_reader_source_topology()
        + audit_keynote_table_merge_source_topology()
        + audit_keynote_table_merge_reader_source_topology()
        + audit_numbers_table_merge_source_topology()
        + audit_iwa_numbers_wire_table_merge_writer_source_topology()
        + audit_numbers_table_merge_transaction_source_topology()
        + audit_pages_table_merge_transaction_source_topology()
        + audit_keynote_table_merge_transaction_source_topology()
        + audit_iwa_table_merge_formula_codec_source_topology()
        + audit_pages_body_table_deletion_source_topology()
        + audit_iwa_package_metadata_component_removal_source_topology()
        + audit_iwa_numbers_table_cell_dependency_removal_source_topology()
        + audit_iwa_numbers_table_cell_dependency_inbound_source_topology()
        + audit_pages_table_cells_source_topology()
        + audit_keynote_table_cells_source_topology()
        + audit_xlsb_source_topology()
        + audit_spreadsheet_sheet_view_source_topology()
        + audit_spreadsheet_chart_source_topology()
        + audit_spreadsheet_shape_source_topology()
    )
    if violations:
        for violation in sorted(set(violations)):
            print(f"crate-boundary error: {violation}", file=sys.stderr)
        return 1

    declaration_count = sum(len(items) for items in snapshot.edges.values())
    debt_count = (
        len(policy.migration_debt)
        + len(policy.core_dependency_debt)
        + len(policy.core_feature_debt)
    )
    print(
        f"crate boundaries valid for {len(snapshot.packages)} workspace packages and "
        f"{declaration_count} internal dependency declarations ({debt_count} explicit debt items)"
    )
    for line in debt_report(policy, args.explain):
        print(line)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
