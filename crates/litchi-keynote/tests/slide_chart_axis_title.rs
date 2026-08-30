//! Strict, source-preserving Keynote chart-axis-title integration coverage.
//!
//! The fixture below follows the producer graph used by source-built iWork
//! charts: a chart drawable owns a title stand-in, chart non-style, primary
//! and secondary axis style/non-style objects, and records the same private
//! graph in its IWA message metadata.  The tests deliberately mutate only
//! this test-local fixture; no production source is involved.

use std::io;

use litchi_iwa_archive::Limits;
use litchi_iwa_common::wire::WireView;
use litchi_iwa_protos::{tsd, tsp};
use litchi_keynote::{
    Axis, ChartAxisTitleError, ChartAxisTitleLimitKind, ChartSelector, Package, Position,
    ReadOptions, SemanticLimits, SlideSelector,
};
use prost::Message as _;

#[path = "support/chart_axis_fixture.rs"]
mod chart_fixture;
use chart_fixture::*;

#[test]
fn primary_category_and_value_titles_use_semantic_selectors() -> TestResult<()> {
    let source = synthetic_metadata_package(false)?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.slide_chart_axis_title("Charts", "Revenue", Axis::Category)?,
        Some("Month".to_owned())
    );
    assert_eq!(
        package.slide_chart_axis_title(
            SlideSelector::index(0),
            ChartSelector::index(0),
            Axis::Value
        )?,
        Some("Revenue".to_owned())
    );

    let noop = package
        .edit_slide_chart_axis_title(0usize, 0usize, Axis::Category)?
        .set("Month")?
        .commit()?;
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(exact_bytes(noop.package())?, source);
    let reapplied_noop = package.apply_slide_chart_axis_title(noop.patch())?;
    assert!(reapplied_noop.patch().is_noop());
    assert!(!reapplied_noop.diagnostics().changed());
    assert_eq!(exact_bytes(reapplied_noop.package())?, source);

    let category = package.edit_slide_chart_axis_title("Charts", "Revenue", Axis::Category)?;
    assert_eq!(category.slide_position(), Position::new(0));
    assert_eq!(category.chart_position(), Position::new(0));
    assert_eq!(category.before(), Some("Month"));
    assert_eq!(category.after(), Some("Month"));
    let category = category.set("Month name")?.commit()?;
    assert_eq!(
        category
            .package()
            .slide_chart_axis_title(0usize, 0usize, Axis::Category)?,
        Some("Month name".to_owned())
    );
    assert_eq!(
        category
            .package()
            .slide_chart_axis_title(0usize, 0usize, Axis::Value)?,
        Some("Revenue".to_owned())
    );
    assert_locality(&source, &exact_bytes(category.package())?)?;
    assert!(category.diagnostics().changed());
    assert!(category.diagnostics().full_reparse_performed());
    assert!(category.diagnostics().touched_components() >= 1);
    assert_eq!(category.diagnostics().deleted_previews(), PREVIEWS.len());

    let value = package
        .edit_slide_chart_axis_title("Charts", "Revenue", Axis::Value)?
        .set("Gross revenue")?
        .commit()?;
    assert_eq!(
        value
            .package()
            .slide_chart_axis_title(0usize, 0usize, Axis::Value)?,
        Some("Gross revenue".to_owned())
    );
    assert_eq!(
        value
            .package()
            .slide_chart_axis_title(0usize, 0usize, Axis::Category)?,
        Some("Month".to_owned())
    );
    Ok(())
}

#[test]
fn public_axis_title_handles_are_semantic_and_redact_native_identity() -> TestResult<()> {
    let package = Package::from_bytes(&synthetic_metadata_package(false)?)?;
    let edit = package.edit_slide_chart_axis_title(
        SlideSelector::name("Charts"),
        ChartSelector::name("Revenue"),
        Axis::Category,
    )?;
    assert_eq!(edit.slide_position(), Position::new(0));
    assert_eq!(edit.chart_position(), Position::new(0));

    // The public transaction handle intentionally exposes only semantic
    // positions. Native chart, stand-in, and axis object identifiers must not
    // become an accidental debugging or logging surface.
    let debug = format!("{edit:?}");
    for identifier in [
        CHARTS[0],
        TITLES[0],
        CHART_NON_STYLES[0],
        CATEGORY_STYLES[0],
        CATEGORY_NON_STYLES[0],
        VALUE_STYLES[0],
        VALUE_NON_STYLES[0],
    ] {
        assert!(
            !debug.contains(&identifier.to_string()),
            "native identifier leaked from semantic edit debug output: {debug}"
        );
    }

    let commit = edit.set("typed selector title")?.commit()?;
    let patch_debug = format!("{:?}", commit.patch());
    for identifier in [
        CHARTS[0],
        TITLES[0],
        CHART_NON_STYLES[0],
        CATEGORY_STYLES[0],
        CATEGORY_NON_STYLES[0],
    ] {
        assert!(
            !patch_debug.contains(&identifier.to_string()),
            "native identifier leaked from semantic patch debug output: {patch_debug}"
        );
    }
    Ok(())
}

#[test]
fn visible_without_text_is_empty_and_hidden_stale_text_is_ignored() -> TestResult<()> {
    let package = Package::from_bytes(&synthetic_package()?)?;
    assert_eq!(
        package.slide_chart_axis_title(0usize, 1usize, Axis::Category)?,
        Some(String::new())
    );
    assert_eq!(
        package.slide_chart_axis_title(0usize, 1usize, Axis::Value)?,
        None
    );
    assert_eq!(
        package.slide_chart_axis_title(0usize, 1usize, Axis::Value)?,
        None
    );
    Ok(())
}

#[test]
fn absent_clear_is_exact_noop_and_set_clear_inverse_reopens_exactly() -> TestResult<()> {
    let source = synthetic_package_without_category_extension()?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.slide_chart_axis_title(0usize, 0usize, Axis::Category)?,
        None
    );
    let absent_clear = package
        .edit_slide_chart_axis_title(0usize, 0usize, Axis::Category)?
        .clear()?
        .commit()?;
    assert!(absent_clear.patch().is_noop());
    assert!(!absent_clear.diagnostics().changed());
    assert_eq!(exact_bytes(absent_clear.package())?, source);

    let created = package
        .edit_slide_chart_axis_title(0usize, 0usize, Axis::Category)?
        .set("Created category")?
        .commit()?;
    let target = exact_bytes(created.package())?;
    let reopened = Package::from_bytes(&target)?;
    assert_eq!(
        reopened.slide_chart_axis_title(0usize, 0usize, Axis::Category)?,
        Some("Created category".to_owned())
    );
    let cleared = reopened
        .edit_slide_chart_axis_title(0usize, 0usize, Axis::Category)?
        .clear()?
        .commit()?;
    assert_eq!(
        cleared
            .package()
            .slide_chart_axis_title(0usize, 0usize, Axis::Category)?,
        None
    );
    let restored = cleared
        .package()
        .apply_slide_chart_axis_title(&cleared.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, target);
    let restored_source = created
        .package()
        .apply_slide_chart_axis_title(&created.patch().inverse())?;
    assert_eq!(exact_bytes(restored_source.package())?, source);
    Ok(())
}

#[test]
fn empty_and_unicode_titles_round_trip_with_explicit_presence() -> TestResult<()> {
    let source = synthetic_package_without_category_extension()?;
    let package = Package::from_bytes(&source)?;

    let empty = package
        .edit_slide_chart_axis_title(
            SlideSelector::position(Position::new(0)),
            ChartSelector::position(Position::new(0)),
            Axis::Category,
        )?
        .set("")?;
    assert_eq!(empty.before(), None);
    assert_eq!(empty.after(), Some(""));
    let empty = empty.commit()?;
    assert_eq!(
        empty.package().slide_chart_axis_title(
            Position::new(0),
            Position::new(0),
            Axis::Category
        )?,
        Some(String::new())
    );

    let empty_bytes = exact_bytes(empty.package())?;
    let reopened = Package::from_bytes(&empty_bytes)?;
    let unicode = reopened
        .edit_slide_chart_axis_title(0usize, 0usize, Axis::Category)?
        .set("売上 📈")?
        .commit()?;
    let _unicode_bytes = exact_bytes(unicode.package())?;
    assert_eq!(
        unicode
            .package()
            .slide_chart_axis_title(0usize, 0usize, Axis::Category)?,
        Some("売上 📈".to_owned())
    );

    let restored_empty = unicode
        .package()
        .apply_slide_chart_axis_title(&unicode.patch().inverse())?;
    assert_eq!(exact_bytes(restored_empty.package())?, empty_bytes);

    let cleared = reopened
        .edit_slide_chart_axis_title(0usize, 0usize, Axis::Category)?
        .clear()?
        .commit()?;
    assert_eq!(
        cleared
            .package()
            .slide_chart_axis_title(0usize, 0usize, Axis::Category)?,
        None
    );
    let restored_source = cleared
        .package()
        .apply_slide_chart_axis_title(&cleared.patch().inverse())?;
    assert_eq!(exact_bytes(restored_source.package())?, empty_bytes);
    Ok(())
}

#[test]
fn secondary_axis_and_unknown_wire_spans_are_untouched() -> TestResult<()> {
    let source = synthetic_metadata_package(false)?;
    let selected_before =
        message_payload(&source, VALUE_NON_STYLES[0], AXIS_NON_STYLE_MESSAGE_TYPE)?;
    let selected_style_before = message_payload(&source, VALUE_STYLES[0], AXIS_STYLE_MESSAGE_TYPE)?;
    let secondary_before = message_payload(
        &source,
        SECONDARY_VALUE_NON_STYLES[0],
        AXIS_NON_STYLE_MESSAGE_TYPE,
    )?;
    let opposite_before =
        message_payload(&source, CATEGORY_NON_STYLES[0], AXIS_NON_STYLE_MESSAGE_TYPE)?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_slide_chart_axis_title(0usize, 0usize, Axis::Value)?
        .set("Primary changed")?
        .commit()?;
    let target = exact_bytes(commit.package())?;
    let selected_after =
        message_payload(&target, VALUE_NON_STYLES[0], AXIS_NON_STYLE_MESSAGE_TYPE)?;
    let selected_style_after = message_payload(&target, VALUE_STYLES[0], AXIS_STYLE_MESSAGE_TYPE)?;
    let secondary_after = message_payload(
        &target,
        SECONDARY_VALUE_NON_STYLES[0],
        AXIS_NON_STYLE_MESSAGE_TYPE,
    )?;
    let opposite_after =
        message_payload(&target, CATEGORY_NON_STYLES[0], AXIS_NON_STYLE_MESSAGE_TYPE)?;
    assert_ne!(selected_before, selected_after);
    assert_eq!(selected_style_before, selected_style_after);
    assert_eq!(secondary_before, secondary_after);
    assert_eq!(opposite_before, opposite_after);
    assert!(
        selected_after
            .windows(b"opaque axis bytes".len())
            .any(|window| window == b"opaque axis bytes")
    );
    assert!(
        selected_style_after
            .windows(b"opaque axis style".len())
            .any(|window| window == b"opaque axis style")
    );
    assert_eq!(
        nested_field_raw(
            &selected_before,
            GENERATED_EXTENSION_FIELD,
            UNKNOWN_GENERATED_FIELD
        )?,
        nested_field_raw(
            &selected_after,
            GENERATED_EXTENSION_FIELD,
            UNKNOWN_GENERATED_FIELD
        )?,
    );

    let category_before =
        message_payload(&source, CATEGORY_NON_STYLES[0], AXIS_NON_STYLE_MESSAGE_TYPE)?;
    let cleared = package
        .edit_slide_chart_axis_title(0usize, 0usize, Axis::Category)?
        .clear()?
        .commit()?;
    let category_after = message_payload(
        &exact_bytes(cleared.package())?,
        CATEGORY_NON_STYLES[0],
        AXIS_NON_STYLE_MESSAGE_TYPE,
    )?;
    for field in [
        UNKNOWN_OUTER_FIELD,
        SUPPORTS_PRIMARY_FEATURE_FIELD,
        SUPPORTS_SECONDARY_FEATURE_FIELD,
    ] {
        assert_eq!(
            raw_fields(&category_before, field)?,
            raw_fields(&category_after, field)?,
            "unknown or capability field changed while clearing the category title: {field}"
        );
    }
    assert_eq!(
        nested_field_raw(
            &category_before,
            GENERATED_EXTENSION_FIELD,
            UNKNOWN_GENERATED_FIELD,
        )?,
        nested_field_raw(
            &category_after,
            GENERATED_EXTENSION_FIELD,
            UNKNOWN_GENERATED_FIELD,
        )?,
    );
    assert_eq!(
        metadata_stream(&source)?,
        metadata_stream(&exact_bytes(cleared.package())?)?
    );
    Ok(())
}

#[test]
fn stale_and_foreign_patches_conflict_without_source_changes() -> TestResult<()> {
    let source = synthetic_metadata_package(false)?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_slide_chart_axis_title(0usize, 0usize, Axis::Category)?
        .set("new category")?
        .commit()?;
    let applied = package.apply_slide_chart_axis_title(commit.patch())?;
    assert_eq!(
        exact_bytes(applied.package())?,
        exact_bytes(commit.package())?
    );
    assert!(matches!(
        applied
            .package()
            .apply_slide_chart_axis_title(commit.patch()),
        Err(ChartAxisTitleError::PatchConflict)
    ));
    assert_eq!(exact_bytes(&package)?, source);

    let stale_bytes = with_axis_payload(
        &source,
        CATEGORY_NON_STYLES[0],
        axis_payload(
            AxisState {
                visible: Some(true),
                title: Some(b"stale source"),
            },
            Axis::Category,
            true,
        )?,
    )?;
    let stale = Package::from_bytes(&stale_bytes)?;
    assert!(matches!(
        stale.apply_slide_chart_axis_title(commit.patch()),
        Err(ChartAxisTitleError::PatchConflict)
    ));
    assert_eq!(exact_bytes(&stale)?, stale_bytes);

    let foreign_source = with_axis_payload(
        &source,
        CATEGORY_NON_STYLES[0],
        axis_payload(
            AxisState {
                visible: Some(true),
                title: Some(b"foreign source"),
            },
            Axis::Category,
            true,
        )?,
    )?;
    let foreign = Package::from_bytes(&foreign_source)?;
    assert!(matches!(
        foreign.apply_slide_chart_axis_title(commit.patch()),
        Err(ChartAxisTitleError::PatchConflict)
    ));
    assert_eq!(exact_bytes(&foreign)?, foreign_source);
    Ok(())
}

#[test]
fn selectors_report_missing_and_ambiguous_chart_names() -> TestResult<()> {
    let source = synthetic_metadata_package(false)?;
    let package = Package::from_bytes(&source)?;
    assert!(matches!(
        package.slide_chart_axis_title(0usize, ChartSelector::name("missing"), Axis::Category),
        Err(ChartAxisTitleError::ChartNameNotFound)
    ));
    assert!(matches!(
        package.slide_chart_axis_title(SlideSelector::name("missing"), 0usize, Axis::Category),
        Err(ChartAxisTitleError::SlideNameNotFound)
    ));
    assert!(matches!(
        package.slide_chart_axis_title(Position::new(9), 0usize, Axis::Category),
        Err(ChartAxisTitleError::SlidePositionNotFound { .. })
    ));
    assert!(matches!(
        package.slide_chart_axis_title(SlideSelector::name(""), 0usize, Axis::Category),
        Err(ChartAxisTitleError::EmptySlideName)
    ));
    assert!(matches!(
        package.slide_chart_axis_title(0usize, ChartSelector::name(""), Axis::Category),
        Err(ChartAxisTitleError::EmptyChartName)
    ));
    assert!(matches!(
        package.edit_slide_chart_axis_title(0usize, 4usize, Axis::Category),
        Err(ChartAxisTitleError::ChartPositionNotFound { position })
            if position == Position::new(4)
    ));

    let duplicate_chart_title = with_document_message_payload(
        &source,
        CHART_NON_STYLES[1],
        CHART_NON_STYLE_MESSAGE_TYPE,
        chart_non_style_payload("Revenue")?,
    )?;
    let duplicate = Package::from_bytes(&duplicate_chart_title)?;
    assert!(matches!(
        duplicate.slide_chart_axis_title(0usize, ChartSelector::name("Revenue"), Axis::Value),
        Err(ChartAxisTitleError::AmbiguousSelector)
    ));
    Ok(())
}

#[test]
fn locked_or_unsupported_chart_graphs_are_rejected_for_mutation() -> TestResult<()> {
    let source = synthetic_metadata_package(false)?;
    let chart = message_payload(&source, CHARTS[0], CHART_MESSAGE_TYPE)?;
    let drawable = WireView::parse(&chart)?
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("missing chart drawable"))?;
    let mut drawable_value = tsd::DrawableArchive::decode(drawable.payload())?;
    drawable_value.locked = Some(true);
    let drawable_replacement = length_delimited_field_with_key_width(
        1,
        &drawable_value.encode_to_vec(),
        drawable.key().len(),
    );
    let locked_chart = replace_first_field(&chart, 1, &drawable_replacement)?;
    let locked =
        with_document_message_payload(&source, CHARTS[0], CHART_MESSAGE_TYPE, locked_chart)?;
    assert_axis_rejected_atomically(&locked, Axis::Category)?;

    let missing_axis = with_chart_axis_variant(&source, ChartAxisVariant::MissingCategory)?;
    assert_axis_rejected_atomically(&missing_axis, Axis::Category)?;
    Ok(())
}

#[test]
fn malformed_axis_extensions_fail_closed_and_preserve_source_bytes() -> TestResult<()> {
    let source = synthetic_metadata_package(false)?;
    for variant in [
        AxisWireVariant::DuplicateOuter,
        AxisWireVariant::WrongOuterWire,
        AxisWireVariant::NonCanonicalOuter,
        AxisWireVariant::DuplicateTitle,
        AxisWireVariant::InvalidUtf8,
    ] {
        let hostile = with_axis_extension_variant(&source, variant)?;
        let package = Package::from_bytes(&hostile)?;
        assert!(matches!(
            package.slide_chart_axis_title(0usize, 0usize, Axis::Category),
            Err(ChartAxisTitleError::InvalidSource)
        ));
        assert_eq!(exact_bytes(&package)?, hostile);
        assert_axis_rejected_atomically(&hostile, Axis::Category)?;
    }
    Ok(())
}

#[test]
fn missing_duplicate_and_role_aliased_axis_refs_fail_closed() -> TestResult<()> {
    let source = synthetic_metadata_package(false)?;
    for variant in [
        ChartAxisVariant::MissingCategory,
        ChartAxisVariant::MissingValue,
        ChartAxisVariant::DuplicateCategory,
        ChartAxisVariant::AliasedRoles,
    ] {
        let hostile = with_chart_axis_variant(&source, variant)?;
        let axis = if matches!(variant, ChartAxisVariant::MissingValue) {
            Axis::Value
        } else {
            Axis::Category
        };
        let package = Package::from_bytes(&hostile)?;
        assert!(matches!(
            package.slide_chart_axis_title(0usize, 0usize, axis),
            Err(ChartAxisTitleError::InvalidSource)
        ));
        assert_eq!(exact_bytes(&package)?, hostile);
        assert_axis_rejected_atomically(&hostile, axis)?;
    }
    assert_axis_rejected_atomically(&with_axis_role_alias(&source)?, Axis::Category)?;
    Ok(())
}

#[test]
fn foreign_object_and_colliding_data_inbound_edges_are_rejected() -> TestResult<()> {
    let source = synthetic_metadata_package(true)?;
    for mode in [ForeignInbound::Object, ForeignInbound::Field] {
        let hostile = with_foreign_inbound(&source, CATEGORY_NON_STYLES[0], mode)?;
        assert_axis_rejected_atomically(&hostile, Axis::Category)?;
    }
    for mode in [ForeignInbound::Data, ForeignInbound::FieldData] {
        let hostile = with_foreign_inbound(&source, CATEGORY_NON_STYLES[0], mode)?;
        assert_axis_rejected_atomically(&hostile, Axis::Category)?;
    }
    Ok(())
}

#[test]
fn unrelated_data_edges_and_stylesheet_registration_remain_editable() -> TestResult<()> {
    let source = synthetic_metadata_package(true)?;
    for mode in [ForeignInbound::Data, ForeignInbound::FieldData] {
        let unrelated = with_foreign_inbound(&source, 2_002, mode)?;
        let commit = Package::from_bytes(&unrelated)?
            .edit_slide_chart_axis_title(0usize, 0usize, Axis::Category)?
            .set("Registered axis")?
            .commit()?;
        assert_eq!(
            commit
                .package()
                .slide_chart_axis_title(0usize, 0usize, Axis::Category)?,
            Some("Registered axis".to_owned())
        );
    }

    let registered =
        with_stylesheet_registration(&synthetic_metadata_package(false)?, VALUE_NON_STYLES[0])?;
    let commit = Package::from_bytes(&registered)?
        .edit_slide_chart_axis_title(0usize, 0usize, Axis::Value)?
        .set("Native registry")?
        .commit()?;
    assert_eq!(
        commit
            .package()
            .slide_chart_axis_title(0usize, 0usize, Axis::Value)?,
        Some("Native registry".to_owned())
    );
    Ok(())
}

#[test]
fn native_stylesheet_owned_chart_graph_remains_selector_first_and_reversible() -> TestResult<()> {
    let source = with_native_stylesheet_layout(&synthetic_metadata_package(false)?)?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.slide_chart_title(0usize, 0usize)?,
        Some("Revenue".to_owned())
    );
    assert_eq!(
        package.slide_chart_axis_title(0usize, 0usize, Axis::Value)?,
        Some("Revenue".to_owned())
    );

    let commit = package
        .edit_slide_chart_axis_title(0usize, ChartSelector::name("Revenue"), Axis::Value)?
        .set("Native registry")?
        .commit()?;
    assert_eq!(
        commit
            .package()
            .slide_chart_axis_title(0usize, 0usize, Axis::Value)?,
        Some("Native registry".to_owned())
    );
    let inverse = commit
        .package()
        .apply_slide_chart_axis_title(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(inverse.package())?, source);

    let title_commit = package
        .edit_slide_chart_title(0usize, 0usize)?
        .set("Native chart title")?
        .commit()?;
    assert_eq!(
        title_commit.package().slide_chart_title(0usize, 0usize)?,
        Some("Native chart title".to_owned())
    );
    Ok(())
}

#[test]
fn metadata_identity_and_authority_namespaces_are_rejected() -> TestResult<()> {
    let source = synthetic_metadata_package(false)?;
    let variants = [
        with_metadata(&source, |metadata| {
            let document = metadata
                .components
                .iter_mut()
                .find(|component| component.identifier == DOCUMENT_COMPONENT)
                .expect("document component");
            document
                .object_uuid_map_entries
                .retain(|entry| entry.identifier != CATEGORY_NON_STYLES[0]);
        })?,
        with_metadata(&source, |metadata| {
            let entry = metadata
                .components
                .iter()
                .find(|component| component.identifier == DOCUMENT_COMPONENT)
                .and_then(|component| {
                    component
                        .object_uuid_map_entries
                        .iter()
                        .find(|entry| entry.identifier == CATEGORY_NON_STYLES[0])
                })
                .cloned()
                .expect("category axis uuid");
            metadata
                .components
                .iter_mut()
                .find(|component| component.identifier == DOCUMENT_COMPONENT)
                .expect("document component")
                .object_uuid_map_entries
                .retain(|candidate| candidate.identifier != CATEGORY_NON_STYLES[0]);
            metadata
                .components
                .iter_mut()
                .find(|component| component.identifier == UNRELATED_COMPONENT)
                .expect("unrelated component")
                .object_uuid_map_entries
                .push(entry);
        })?,
        with_metadata(&source, |metadata| {
            metadata
                .components
                .iter_mut()
                .find(|component| component.identifier == DOCUMENT_COMPONENT)
                .expect("document component")
                .object_uuid_map_entries
                .push(metadata_uuid_entry(CATEGORY_NON_STYLES[0]));
        })?,
        with_metadata(&source, |metadata| {
            metadata.versioned_components.push(tsp::ComponentInfo {
                identifier: DOCUMENT_COMPONENT,
                preferred_locator: "Document".to_owned(),
                locator: Some("Document".to_owned()),
                object_uuid_map_entries: vec![metadata_uuid_entry(CATEGORY_NON_STYLES[0])],
                ..tsp::ComponentInfo::default()
            });
        })?,
        with_metadata(&source, |metadata| {
            metadata
                .components
                .iter_mut()
                .find(|component| component.identifier == DOCUMENT_COMPONENT)
                .expect("document component")
                .ambiguous_object_identifiers
                .push(CATEGORY_NON_STYLES[0]);
        })?,
        with_metadata(&source, |metadata| {
            metadata.data_metadata_map = Some(reference(CATEGORY_NON_STYLES[0]));
        })?,
        with_metadata(&source, |metadata| {
            metadata
                .components
                .iter_mut()
                .find(|component| component.identifier == DOCUMENT_COMPONENT)
                .expect("document component")
                .data_references
                .push(tsp::ComponentDataReference {
                    data_identifier: 2_002,
                    object_reference_list: vec![tsp::component_data_reference::ObjectReference {
                        object_identifier: CATEGORY_NON_STYLES[0],
                        count: 1,
                    }],
                });
        })?,
    ];
    for hostile in variants {
        assert_axis_rejected_atomically(&hostile, Axis::Category)?;
    }
    Ok(())
}

#[test]
fn output_limit_at_max_minus_one_rejects_atomically() -> TestResult<()> {
    let source = synthetic_metadata_package(false)?;
    let long_title = "a".repeat(16 * 1024);
    let baseline = Package::from_bytes(&source)?
        .edit_slide_chart_axis_title(0usize, 0usize, Axis::Category)?
        .set(long_title.clone())?
        .commit()?;
    let target = exact_bytes(baseline.package())?;
    assert!(target.len() > source.len());
    let defaults = Limits::default();
    let limits = Limits::new(
        u64::try_from(target.len() - 1)?,
        defaults.max_entries(),
        defaults.max_entry_bytes(),
        defaults.max_total_bytes(),
        defaults.max_iwa_stream_bytes(),
    )?;
    let package = Package::from_bytes_with_options(
        &source,
        ReadOptions::new(limits, SemanticLimits::default()),
    )?;
    let before = exact_bytes(&package)?;
    let result = package
        .edit_slide_chart_axis_title(0usize, 0usize, Axis::Category)
        .and_then(|edit| edit.set(long_title))
        .and_then(|edit| edit.commit());
    assert!(matches!(
        result,
        Err(ChartAxisTitleError::LimitExceeded {
            kind: ChartAxisTitleLimitKind::OutputBytes,
            ..
        })
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}
