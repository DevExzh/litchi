//! Strict, source-preserving integration coverage for Keynote value-axis settings.
//!
//! The source graph and package-reassembly helpers are shared with the
//! chart-axis-title fixture.  Keeping the fixture in one place is deliberate:
//! both owners must admit the same slide/chart ownership proof, including
//! metadata and normalized `DocumentStylesheet` layouts.  The tests below add
//! only value-axis wire fields to that graph and exercise the aggregate,
//! selector-first transaction.
use std::io;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::{WireView, append_length_delimited_field, append_varint_field};
use litchi_iwa_protos::{tsch, tsd, tsp};
use litchi_keynote::{
    ChartSelector, Package, Position, ReadOptions, SemanticLimits, SlideSelector,
};
use prost::Message as _;

#[path = "support/chart_axis_fixture.rs"]
mod chart_fixture;
use chart_fixture::*;

mod value_axis {
    use super::*;

    use litchi_iwa_common::chart::axis::{
        Bound, Bounds, MajorStepCount, MinorStepCount, Scale, Steps, ValueAxisSettings,
    };
    use litchi_keynote::{ChartValueAxisError, ChartValueAxisLimitKind};

    const VALUE_MAJOR_FIELD: u32 = 5;
    const VALUE_MINOR_FIELD: u32 = 6;
    const VALUE_SCALE_FIELD: u32 = 8;
    const VALUE_MAXIMUM_FIELD: u32 = 17;
    const VALUE_MINIMUM_FIELD: u32 = 18;
    const DRAWABLE_SUPER_FIELD: u32 = 1;

    fn fixed64_field(number: u32, bits: u64) -> Vec<u8> {
        let mut output = Vec::with_capacity(9);
        let key = u64::from(number) << 3 | 1;
        let mut key = key;
        while key >= 0x80 {
            output.push((key as u8 & 0x7f) | 0x80);
            key >>= 7;
        }
        output.push(key as u8);
        output.extend_from_slice(&bits.to_le_bytes());
        output
    }

    fn encoded_bound(value: f64, unknown: bool) -> Vec<u8> {
        let mut nested = fixed64_field(1, value.to_bits());
        if unknown {
            // Deliberately place an opaque field after the scalar.  A source
            // preserving rewrite must retain it and its original order.
            append_length_delimited_field(&mut nested, 4_777, b"nested bound extension")
                .expect("small test field fits");
        }
        nested
    }

    fn settings_wire(settings: ValueAxisSettings, unknown_nested: bool) -> Vec<u8> {
        let mut generated = Vec::new();
        if let Some(major) = settings.steps().major() {
            append_varint_field(&mut generated, VALUE_MAJOR_FIELD, u64::from(major.value()))
                .expect("validated major step fits");
        }
        if let Some(minor) = settings.steps().minor() {
            append_varint_field(&mut generated, VALUE_MINOR_FIELD, u64::from(minor.value()))
                .expect("validated minor step fits");
        }
        append_varint_field(
            &mut generated,
            VALUE_SCALE_FIELD,
            settings.scale().native_value() as u32 as u64,
        )
        .expect("native scale fits");
        if let Some(maximum) = settings.bounds().maximum() {
            append_length_delimited_field(
                &mut generated,
                VALUE_MAXIMUM_FIELD,
                &encoded_bound(maximum.value(), unknown_nested),
            )
            .expect("small bound fits");
        }
        if let Some(minimum) = settings.bounds().minimum() {
            append_length_delimited_field(
                &mut generated,
                VALUE_MINIMUM_FIELD,
                &encoded_bound(minimum.value(), unknown_nested),
            )
            .expect("small bound fits");
        }
        generated
    }

    fn with_value_axis_settings(
        source: &[u8],
        settings: ValueAxisSettings,
        unknown_nested: bool,
    ) -> TestResult<Vec<u8>> {
        let original = message_payload(source, VALUE_NON_STYLES[0], AXIS_NON_STYLE_MESSAGE_TYPE)?;
        let extension = WireView::parse(&original)?
            .fields()
            .find(|field| field.number() == GENERATED_EXTENSION_FIELD)
            .ok_or_else(|| io::Error::other("missing generated axis extension"))?;
        let mut generated = extension.payload().to_vec();
        generated.extend_from_slice(&settings_wire(settings, unknown_nested));
        let replacement = length_delimited_field_with_key_width(
            GENERATED_EXTENSION_FIELD,
            &generated,
            extension.key().len(),
        );
        let payload = replace_first_field(&original, GENERATED_EXTENSION_FIELD, &replacement)?;
        with_axis_payload(source, VALUE_NON_STYLES[0], payload)
    }

    fn nested_unknown_fields(
        payload: &[u8],
        outer: u32,
        nested: u32,
        unknown: u32,
    ) -> TestResult<Vec<Vec<u8>>> {
        let outer = WireView::parse(payload)?
            .fields()
            .find(|field| field.number() == outer)
            .ok_or_else(|| io::Error::other("missing synthetic outer field"))?;
        let nested = WireView::parse(outer.payload())?
            .fields()
            .find(|field| field.number() == nested)
            .ok_or_else(|| io::Error::other("missing synthetic nested field"))?;
        Ok(WireView::parse(nested.payload())?
            .fields()
            .filter(|field| field.number() == unknown)
            .map(|field| field.raw().to_vec())
            .collect())
    }

    fn without_previews(source: &[u8]) -> TestResult<Vec<u8>> {
        let catalog = Catalog::from_bytes(source)?;
        let entries = catalog
            .iter()
            .filter(|entry| !PREVIEWS.contains(&entry.name()))
            .map(|entry| (entry.name(), entry.data()))
            .collect::<Vec<_>>();
        Ok(litchi_iwa_archive::package::to_bytes(
            entries,
            Limits::default(),
        )?)
    }

    fn with_shared_primary_value_axis(source: &[u8]) -> TestResult<Vec<u8>> {
        let original = message_payload(source, CHARTS[1], CHART_MESSAGE_TYPE)?;
        let chart_extension = WireView::parse(&original)?
            .fields()
            .find(|field| field.number() == GENERATED_EXTENSION_FIELD)
            .ok_or_else(|| io::Error::other("missing synthetic chart extension"))?;
        let chart = tsch::ChartArchive::decode(chart_extension.payload())?;
        let chart = tsch::ChartArchive {
            value_axis_nonstyles: vec![
                reference(VALUE_NON_STYLES[0]),
                reference(SECONDARY_VALUE_NON_STYLES[1]),
            ],
            ..chart
        };
        let replacement = length_delimited_field_with_key_width(
            GENERATED_EXTENSION_FIELD,
            &chart.encode_to_vec(),
            chart_extension.key().len(),
        );
        let payload = replace_first_field(&original, GENERATED_EXTENSION_FIELD, &replacement)?;
        with_chart_payload(source, 1, payload)
    }

    fn settings(
        minimum: Option<f64>,
        maximum: Option<f64>,
        major: Option<u32>,
        minor: Option<u32>,
        scale: Scale,
    ) -> ValueAxisSettings {
        let minimum = minimum.map(|value| Bound::new(value).expect("finite test minimum"));
        let maximum = maximum.map(|value| Bound::new(value).expect("finite test maximum"));
        let bounds = Bounds::new(minimum, maximum).expect("ordered test bounds");
        let major = major.map(|value| MajorStepCount::new(value).expect("valid major step"));
        let minor = minor.map(|value| MinorStepCount::new(value).expect("valid minor step"));
        ValueAxisSettings::new(bounds, Steps::new(major, minor), scale)
    }

    fn value_axis(package: &Package) -> TestResult<ValueAxisSettings> {
        Ok(package.slide_chart_value_axis_settings(
            SlideSelector::name("Charts"),
            ChartSelector::name("Revenue"),
        )?)
    }

    fn commit_settings(
        package: &Package,
        target: ValueAxisSettings,
    ) -> TestResult<litchi_keynote::ChartValueAxisCommit> {
        Ok(package
            .edit_slide_chart_value_axis_settings(
                SlideSelector::name("Charts"),
                ChartSelector::name("Revenue"),
            )?
            .set(target)?
            .commit()?)
    }

    fn error_is_redacted<E: std::fmt::Display>(error: &E) {
        let text = error.to_string();
        for identifier in [
            CHARTS[0],
            TITLES[0],
            CHART_NON_STYLES[0],
            VALUE_STYLES[0],
            VALUE_NON_STYLES[0],
        ] {
            assert!(
                !text.contains(&identifier.to_string()),
                "native identifier leaked from value-axis error: {text}"
            );
        }
    }

    #[test]
    fn defaults_and_name_or_position_selectors_read_the_primary_value_axis() -> TestResult<()> {
        let source = synthetic_metadata_package(false)?;
        let package = Package::from_bytes(&source)?;
        let expected = ValueAxisSettings::automatic();
        assert_eq!(value_axis(&package)?, expected);
        assert_eq!(
            package.slide_chart_value_axis_settings(
                SlideSelector::index(0),
                ChartSelector::index(0)
            )?,
            expected
        );
        assert_eq!(
            package.slide_chart_value_axis_settings(
                SlideSelector::name("Charts"),
                ChartSelector::name("Revenue")
            )?,
            expected
        );
        Ok(())
    }

    #[test]
    fn explicit_bounds_preserve_fixed_partial_zero_and_automatic_values() -> TestResult<()> {
        let source = synthetic_metadata_package(false)?;
        let automatic = ValueAxisSettings::automatic();
        let cases = [
            settings(Some(-10.0), Some(40.0), None, None, Scale::Linear),
            settings(Some(0.0), None, None, None, Scale::Linear),
            settings(None, Some(0.0), None, None, Scale::Linear),
            settings(None, None, None, None, Scale::Linear),
        ];
        for target in cases {
            let source = with_value_axis_settings(&source, target, false)?;
            let package = Package::from_bytes(&source)?;
            assert_eq!(value_axis(&package)?, target);
            let commit = commit_settings(&package, automatic)?;
            assert_eq!(value_axis(commit.package())?, automatic);
        }
        Ok(())
    }

    #[test]
    fn steps_preserve_both_major_only_minor_zero_and_automatic_modes() -> TestResult<()> {
        let source = synthetic_metadata_package(false)?;
        let cases = [
            settings(None, None, Some(6), Some(2), Scale::Linear),
            settings(None, None, Some(4), None, Scale::Linear),
            settings(None, None, None, Some(0), Scale::Linear),
            settings(None, None, None, None, Scale::Linear),
        ];
        for target in cases {
            let source = with_value_axis_settings(&source, target, false)?;
            let package = Package::from_bytes(&source)?;
            assert_eq!(value_axis(&package)?.steps(), target.steps());
        }
        Ok(())
    }

    #[test]
    fn known_and_future_scales_round_trip_without_normalization() -> TestResult<()> {
        let source = synthetic_metadata_package(false)?;
        for scale in [Scale::Linear, Scale::Logarithmic, Scale::Unsupported(9_001)] {
            let target = ValueAxisSettings::automatic().with_scale(scale);
            let source = with_value_axis_settings(&source, target, false)?;
            let package = Package::from_bytes(&source)?;
            assert_eq!(value_axis(&package)?.scale(), scale);
            let commit = commit_settings(&package, target.with_scale(Scale::Linear))?;
            assert_eq!(value_axis(commit.package())?.scale(), Scale::Linear);
        }
        Ok(())
    }

    #[test]
    fn one_edit_atomically_changes_all_three_controls_and_invalidates_previews() -> TestResult<()> {
        let source = synthetic_metadata_package(false)?;
        let package = Package::from_bytes(&source)?;
        let target = settings(Some(1.0), Some(30.0), Some(6), Some(2), Scale::Logarithmic);
        let edit = package.edit_slide_chart_value_axis_settings(
            SlideSelector::name("Charts"),
            ChartSelector::name("Revenue"),
        )?;
        assert_eq!(edit.before(), ValueAxisSettings::automatic());
        assert_eq!(edit.after(), ValueAxisSettings::automatic());
        let commit = edit.set(target)?.commit()?;
        assert_eq!(value_axis(commit.package())?, target);
        assert!(commit.diagnostics().changed());
        assert!(commit.diagnostics().full_reparse_performed());
        assert_eq!(commit.diagnostics().deleted_previews(), PREVIEWS.len());
        assert!(commit.diagnostics().touched_components() >= 1);
        assert_locality(&source, &exact_bytes(commit.package())?)?;
        Ok(())
    }

    #[test]
    fn focused_setters_share_one_atomic_transaction() -> TestResult<()> {
        let source = synthetic_metadata_package(false)?;
        let package = Package::from_bytes(&source)?;
        let target = settings(Some(5.0), Some(40.0), Some(5), Some(0), Scale::Logarithmic);
        let edit = package.edit_slide_chart_value_axis_settings(
            SlideSelector::index(0),
            ChartSelector::index(0),
        )?;
        let edit = edit.set_bounds(target.bounds())?;
        let edit = edit.set_steps(target.steps())?;
        let edit = edit.set_scale(target.scale())?;
        let commit = edit.commit()?;
        assert_eq!(value_axis(commit.package())?, target);
        Ok(())
    }

    #[test]
    fn primary_value_edit_is_isolated_from_duplicate_chart_and_other_axis_roles() -> TestResult<()>
    {
        let source = synthetic_metadata_package(false)?;
        let package = Package::from_bytes(&source)?;
        let second_chart_before = package
            .slide_chart_value_axis_settings(SlideSelector::index(0), ChartSelector::index(1))?;
        let second_chart_message_before = message_payload(&source, CHARTS[1], CHART_MESSAGE_TYPE)?;
        let category_before =
            message_payload(&source, CATEGORY_NON_STYLES[0], AXIS_NON_STYLE_MESSAGE_TYPE)?;
        let value_style_before =
            message_payload(&source, VALUE_STYLES[0], AXIS_STYLE_MESSAGE_TYPE)?;
        let secondary_style_before =
            message_payload(&source, SECONDARY_VALUE_STYLES[0], AXIS_STYLE_MESSAGE_TYPE)?;

        let target = settings(Some(1.0), Some(40.0), Some(8), Some(2), Scale::Logarithmic);
        let commit = commit_settings(&package, target)?;
        let after = exact_bytes(commit.package())?;
        assert_eq!(
            commit.package().slide_chart_value_axis_settings(
                SlideSelector::index(0),
                ChartSelector::index(1),
            )?,
            second_chart_before
        );
        assert_eq!(
            message_payload(&after, CHARTS[1], CHART_MESSAGE_TYPE)?,
            second_chart_message_before
        );
        assert_eq!(
            message_payload(&after, CATEGORY_NON_STYLES[0], AXIS_NON_STYLE_MESSAGE_TYPE)?,
            category_before
        );
        assert_eq!(
            message_payload(&after, VALUE_STYLES[0], AXIS_STYLE_MESSAGE_TYPE)?,
            value_style_before
        );
        assert_eq!(
            message_payload(&after, SECONDARY_VALUE_STYLES[0], AXIS_STYLE_MESSAGE_TYPE)?,
            secondary_style_before
        );
        Ok(())
    }

    #[test]
    fn unknown_spans_styles_and_unrelated_zip_members_survive_value_edit() -> TestResult<()> {
        let source = with_value_axis_settings(
            &synthetic_metadata_package(true)?,
            settings(Some(0.0), Some(100.0), Some(5), Some(0), Scale::Linear),
            true,
        )?;
        let before_axis =
            message_payload(&source, VALUE_NON_STYLES[0], AXIS_NON_STYLE_MESSAGE_TYPE)?;
        let before_outer_unknown = raw_fields(&before_axis, UNKNOWN_OUTER_FIELD)?;
        let before_generated_unknown = nested_field_raw(
            &before_axis,
            GENERATED_EXTENSION_FIELD,
            UNKNOWN_GENERATED_FIELD,
        )?;
        let before_bound_unknown = nested_unknown_fields(
            &before_axis,
            GENERATED_EXTENSION_FIELD,
            VALUE_MAXIMUM_FIELD,
            4_777,
        )?;
        let before_category_style =
            message_payload(&source, CATEGORY_STYLES[0], AXIS_STYLE_MESSAGE_TYPE)?;
        let before_secondary_axis = message_payload(
            &source,
            SECONDARY_VALUE_NON_STYLES[0],
            AXIS_NON_STYLE_MESSAGE_TYPE,
        )?;

        let package = Package::from_bytes(&source)?;
        let commit = commit_settings(
            &package,
            settings(Some(10.0), Some(80.0), Some(4), Some(1), Scale::Logarithmic),
        )?;
        let after = exact_bytes(commit.package())?;
        let after_axis = message_payload(&after, VALUE_NON_STYLES[0], AXIS_NON_STYLE_MESSAGE_TYPE)?;
        assert_eq!(
            raw_fields(&after_axis, UNKNOWN_OUTER_FIELD)?,
            before_outer_unknown
        );
        assert_eq!(
            nested_field_raw(
                &after_axis,
                GENERATED_EXTENSION_FIELD,
                UNKNOWN_GENERATED_FIELD
            )?,
            before_generated_unknown
        );
        assert_eq!(
            nested_unknown_fields(
                &after_axis,
                GENERATED_EXTENSION_FIELD,
                VALUE_MAXIMUM_FIELD,
                4_777
            )?,
            before_bound_unknown
        );
        assert_eq!(
            message_payload(&after, CATEGORY_STYLES[0], AXIS_STYLE_MESSAGE_TYPE)?,
            before_category_style
        );
        assert_eq!(
            message_payload(
                &after,
                SECONDARY_VALUE_NON_STYLES[0],
                AXIS_NON_STYLE_MESSAGE_TYPE
            )?,
            before_secondary_axis
        );
        assert_locality(&source, &after)?;
        Ok(())
    }

    #[test]
    fn semantic_noop_keeps_exact_package_bytes_and_skips_publication() -> TestResult<()> {
        let source = with_value_axis_settings(
            &synthetic_metadata_package(false)?,
            settings(Some(0.0), Some(100.0), Some(5), Some(0), Scale::Linear),
            true,
        )?;
        let package = Package::from_bytes(&source)?;
        let before = value_axis(&package)?;
        let commit = commit_settings(&package, before)?;
        assert!(commit.patch().is_noop());
        assert!(!commit.diagnostics().changed());
        assert_eq!(exact_bytes(commit.package())?, source);
        assert_eq!(commit.diagnostics().deleted_previews(), 0);
        let reapplied = package.apply_slide_chart_value_axis_settings(commit.patch())?;
        assert!(reapplied.patch().is_noop());
        assert!(!reapplied.diagnostics().changed());
        assert_eq!(exact_bytes(reapplied.package())?, source);
        Ok(())
    }

    #[test]
    fn semantic_validation_rejects_inverted_nonfinite_and_invalid_log_ranges() -> TestResult<()> {
        let low = Bound::new(10.0).expect("finite lower bound");
        let high = Bound::new(1.0).expect("finite upper bound");
        assert!(Bounds::new(Some(low), Some(high)).is_err());
        assert!(Bound::new(f64::NAN).is_err());
        assert!(Bound::new(f64::INFINITY).is_err());
        assert!(MajorStepCount::new(0).is_err());
        assert!(MinorStepCount::new(u32::MAX).is_err());

        let package = Package::from_bytes(&synthetic_metadata_package(false)?)?;
        let invalid = ValueAxisSettings::automatic()
            .with_bounds(Bounds::new(Some(Bound::new(-1.0)?), Some(high))?)
            .with_scale(Scale::Logarithmic);
        let before = exact_bytes(&package)?;
        let error = package
            .edit_slide_chart_value_axis_settings(SlideSelector::index(0), ChartSelector::index(0))?
            .set(invalid)
            .expect_err("logarithmic axes must have positive manual bounds");
        assert!(matches!(error, ChartValueAxisError::InvalidSettings));
        error_is_redacted(&error);
        assert_eq!(exact_bytes(&package)?, before);
        Ok(())
    }

    #[test]
    fn changed_candidate_reopens_and_inverse_restores_exact_source() -> TestResult<()> {
        let source = synthetic_metadata_package(false)?;
        let package = Package::from_bytes(&source)?;
        let source_before = exact_bytes(&package)?;
        let target = settings(Some(1.0), Some(100.0), Some(5), Some(1), Scale::Logarithmic);
        let commit = commit_settings(&package, target)?;
        assert_eq!(exact_bytes(&package)?, source_before);
        let target_bytes = exact_bytes(commit.package())?;
        let reopened = Package::from_bytes(&target_bytes)?;
        assert_eq!(value_axis(&reopened)?, target);
        let restored = reopened.apply_slide_chart_value_axis_settings(&commit.patch().inverse())?;
        assert_eq!(
            value_axis(restored.package())?,
            ValueAxisSettings::automatic()
        );
        assert_eq!(exact_bytes(restored.package())?, source);
        Ok(())
    }

    #[test]
    fn patch_conflicts_are_rejected_without_mutating_source() -> TestResult<()> {
        let source = synthetic_metadata_package(false)?;
        let package = Package::from_bytes(&source)?;
        let source_before = exact_bytes(&package)?;
        let commit = commit_settings(
            &package,
            settings(Some(1.0), Some(20.0), Some(4), Some(1), Scale::Logarithmic),
        )?;
        assert_eq!(exact_bytes(&package)?, source_before);
        let target_bytes = exact_bytes(commit.package())?;
        let target = Package::from_bytes(&target_bytes)?;
        let stale = Package::from_bytes(&target_bytes)?;
        let stale_before = exact_bytes(&stale)?;
        let error = stale
            .apply_slide_chart_value_axis_settings(commit.patch())
            .expect_err("source patch must not apply to stale source");
        assert!(matches!(error, ChartValueAxisError::PatchConflict));
        assert_eq!(exact_bytes(&stale)?, stale_before);

        let foreign = Package::from_bytes(&synthetic_metadata_package(true)?)?;
        let error = foreign
            .apply_slide_chart_value_axis_settings(commit.patch())
            .expect_err("source patch must not cross package artifacts");
        assert!(matches!(error, ChartValueAxisError::PatchConflict));
        assert_eq!(value_axis(&target)?, commit.patch().after());
        Ok(())
    }

    #[test]
    fn semantic_handles_and_errors_redact_native_graph_identity() -> TestResult<()> {
        let package = Package::from_bytes(&synthetic_metadata_package(false)?)?;
        let edit = package.edit_slide_chart_value_axis_settings(
            SlideSelector::name("Charts"),
            ChartSelector::name("Revenue"),
        )?;
        let debug = format!("{edit:?}");
        for identifier in [
            CHARTS[0],
            TITLES[0],
            CHART_NON_STYLES[0],
            VALUE_STYLES[0],
            VALUE_NON_STYLES[0],
        ] {
            assert!(
                !debug.contains(&identifier.to_string()),
                "native ID leaked: {debug}"
            );
        }
        let commit = edit.set(ValueAxisSettings::automatic())?.commit()?;
        let patch_debug = format!("{:?}", commit.patch());
        for identifier in [
            CHARTS[0],
            TITLES[0],
            CHART_NON_STYLES[0],
            VALUE_STYLES[0],
            VALUE_NON_STYLES[0],
        ] {
            assert!(
                !patch_debug.contains(&identifier.to_string()),
                "native ID leaked from semantic patch debug output: {patch_debug}"
            );
        }
        Ok(())
    }

    #[test]
    fn selectors_and_graph_failures_are_atomic_and_redacted() -> TestResult<()> {
        let source = synthetic_metadata_package(false)?;
        let package = Package::from_bytes(&source)?;
        for result in [
            package
                .slide_chart_value_axis_settings(
                    SlideSelector::name("missing"),
                    ChartSelector::index(0),
                )
                .map(|_| ()),
            package
                .slide_chart_value_axis_settings(
                    SlideSelector::index(0),
                    ChartSelector::name("missing"),
                )
                .map(|_| ()),
            package
                .slide_chart_value_axis_settings(
                    SlideSelector::position(Position::new(9)),
                    ChartSelector::index(0),
                )
                .map(|_| ()),
        ] {
            let error = result.expect_err("invalid selector must fail");
            error_is_redacted(&error);
        }
        for variant in [
            ChartAxisVariant::MissingValue,
            ChartAxisVariant::DuplicateCategory,
            ChartAxisVariant::AliasedRoles,
        ] {
            let hostile = with_chart_axis_variant(&source, variant)?;
            let package = Package::from_bytes(&hostile)?;
            let error = package
                .edit_slide_chart_value_axis_settings(
                    SlideSelector::index(0),
                    ChartSelector::index(0),
                )
                .expect_err("invalid graph must fail closed");
            error_is_redacted(&error);
            assert_eq!(exact_bytes(&package)?, hostile);
        }
        Ok(())
    }

    #[test]
    fn locked_shared_foreign_and_ambiguous_graphs_fail_closed() -> TestResult<()> {
        let source = synthetic_metadata_package(true)?;

        let chart = message_payload(&source, CHARTS[0], CHART_MESSAGE_TYPE)?;
        let drawable = WireView::parse(&chart)?
            .fields()
            .find(|field| field.number() == DRAWABLE_SUPER_FIELD)
            .ok_or_else(|| io::Error::other("missing chart drawable"))?;
        let mut drawable_value = tsd::DrawableArchive::decode(drawable.payload())?;
        drawable_value.locked = Some(true);
        let drawable_replacement = length_delimited_field_with_key_width(
            DRAWABLE_SUPER_FIELD,
            &drawable_value.encode_to_vec(),
            drawable.key().len(),
        );
        let locked_chart =
            replace_first_field(&chart, DRAWABLE_SUPER_FIELD, &drawable_replacement)?;
        let locked = with_chart_payload(&source, 0, locked_chart)?;
        let package = Package::from_bytes(&locked)?;
        let before = exact_bytes(&package)?;
        let error = package
            .edit_slide_chart_value_axis_settings(SlideSelector::index(0), ChartSelector::index(0))
            .expect_err("locked chart must be immutable");
        error_is_redacted(&error);
        assert_eq!(exact_bytes(&package)?, before);

        let shared = with_shared_primary_value_axis(&source)?;
        let package = Package::from_bytes(&shared)?;
        let before = exact_bytes(&package)?;
        let error = package
            .edit_slide_chart_value_axis_settings(SlideSelector::index(0), ChartSelector::index(0))
            .expect_err("shared primary axis must be rejected");
        error_is_redacted(&error);
        assert_eq!(exact_bytes(&package)?, before);

        for mode in [
            ForeignInbound::Object,
            ForeignInbound::Field,
            ForeignInbound::Data,
            ForeignInbound::FieldData,
        ] {
            let foreign = with_foreign_inbound(&source, VALUE_NON_STYLES[0], mode)?;
            let package = Package::from_bytes(&foreign)?;
            let before = exact_bytes(&package)?;
            let error = package
                .edit_slide_chart_value_axis_settings(
                    SlideSelector::index(0),
                    ChartSelector::index(0),
                )
                .expect_err("foreign inbound edge must be rejected");
            error_is_redacted(&error);
            assert_eq!(exact_bytes(&package)?, before);
        }

        let duplicate_chart_title = with_document_message_payload(
            &source,
            CHART_NON_STYLES[1],
            CHART_NON_STYLE_MESSAGE_TYPE,
            chart_non_style_payload("Revenue")?,
        )?;
        let package = Package::from_bytes(&duplicate_chart_title)?;
        let error = package
            .slide_chart_value_axis_settings(
                SlideSelector::index(0),
                ChartSelector::name("Revenue"),
            )
            .expect_err("duplicate chart title must be ambiguous");
        error_is_redacted(&error);
        Ok(())
    }

    #[test]
    fn stylesheet_registration_and_native_normalized_layouts_round_trip() -> TestResult<()> {
        let registered =
            with_stylesheet_registration(&synthetic_metadata_package(false)?, VALUE_NON_STYLES[0])?;
        let target = settings(Some(5.0), Some(55.0), Some(5), Some(1), Scale::Logarithmic);
        let package = Package::from_bytes(&registered)?;
        let commit = commit_settings(&package, target)?;
        assert_eq!(value_axis(commit.package())?, target);
        let inverse = commit
            .package()
            .apply_slide_chart_value_axis_settings(&commit.patch().inverse())?;
        assert_eq!(exact_bytes(inverse.package())?, registered);

        let normalized = with_native_stylesheet_layout(&synthetic_metadata_package(false)?)?;
        let package = Package::from_bytes(&normalized)?;
        assert_eq!(value_axis(&package)?, ValueAxisSettings::automatic());
        let commit = commit_settings(&package, target)?;
        assert_eq!(value_axis(commit.package())?, target);
        let inverse = commit
            .package()
            .apply_slide_chart_value_axis_settings(&commit.patch().inverse())?;
        assert_eq!(exact_bytes(inverse.package())?, normalized);
        Ok(())
    }

    #[test]
    fn metadata_identity_authority_variants_remain_atomic() -> TestResult<()> {
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
                    .retain(|entry| entry.identifier != VALUE_NON_STYLES[0]);
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
                            .find(|entry| entry.identifier == VALUE_NON_STYLES[0])
                    })
                    .cloned()
                    .expect("value axis uuid");
                metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == DOCUMENT_COMPONENT)
                    .expect("document component")
                    .object_uuid_map_entries
                    .retain(|candidate| candidate.identifier != VALUE_NON_STYLES[0]);
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
                    .push(metadata_uuid_entry(VALUE_NON_STYLES[0]));
            })?,
            with_metadata(&source, |metadata| {
                metadata.versioned_components.push(tsp::ComponentInfo {
                    identifier: DOCUMENT_COMPONENT,
                    preferred_locator: "Document".to_owned(),
                    locator: Some("Document".to_owned()),
                    object_uuid_map_entries: vec![metadata_uuid_entry(VALUE_NON_STYLES[0])],
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
                    .push(VALUE_NON_STYLES[0]);
            })?,
            with_metadata(&source, |metadata| {
                metadata.data_metadata_map = Some(reference(VALUE_NON_STYLES[0]));
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
                        object_reference_list: vec![
                            tsp::component_data_reference::ObjectReference {
                                object_identifier: VALUE_NON_STYLES[0],
                                count: 1,
                            },
                        ],
                    });
            })?,
        ];
        for hostile in variants {
            let package = Package::from_bytes(&hostile)?;
            let before = exact_bytes(&package)?;
            let error = package
                .edit_slide_chart_value_axis_settings(
                    SlideSelector::index(0),
                    ChartSelector::index(0),
                )
                .expect_err("invalid metadata identity must be rejected");
            error_is_redacted(&error);
            assert_eq!(exact_bytes(&package)?, before);
        }
        Ok(())
    }

    #[test]
    fn malformed_value_axis_wire_rejects_duplicates_wrong_types_noncanonical_and_truncation()
    -> TestResult<()> {
        let source = synthetic_metadata_package(false)?;
        let baseline = ValueAxisSettings::automatic();
        let base = with_value_axis_settings(&source, baseline, false)?;
        let original = message_payload(&base, VALUE_NON_STYLES[0], AXIS_NON_STYLE_MESSAGE_TYPE)?;
        let extension = WireView::parse(&original)?
            .fields()
            .find(|field| field.number() == GENERATED_EXTENSION_FIELD)
            .ok_or_else(|| io::Error::other("missing generated axis extension"))?;

        let malformed = [
            ("duplicate-major", {
                let mut generated = extension.payload().to_vec();
                append_varint_field(&mut generated, VALUE_MAJOR_FIELD, 5)?;
                append_varint_field(&mut generated, VALUE_MAJOR_FIELD, 6)?;
                generated
            }),
            ("wrong-major-wire", {
                let mut generated = extension.payload().to_vec();
                append_length_delimited_field(&mut generated, VALUE_MAJOR_FIELD, b"wrong")?;
                generated
            }),
            ("truncated-bound", {
                let mut generated = extension.payload().to_vec();
                append_length_delimited_field(&mut generated, VALUE_MINIMUM_FIELD, &[0x09, 1, 2])?;
                generated
            }),
            ("missing-bound-scalar", {
                let mut generated = extension.payload().to_vec();
                append_length_delimited_field(&mut generated, VALUE_MAXIMUM_FIELD, &[])?;
                generated
            }),
            ("duplicate-bound-scalar", {
                let mut generated = extension.payload().to_vec();
                let mut nested = fixed64_field(1, 1.0_f64.to_bits());
                nested.extend_from_slice(&fixed64_field(1, 2.0_f64.to_bits()));
                append_length_delimited_field(&mut generated, VALUE_MAXIMUM_FIELD, &nested)?;
                generated
            }),
            ("nonfinite-bound", {
                let mut generated = extension.payload().to_vec();
                append_length_delimited_field(
                    &mut generated,
                    VALUE_MAXIMUM_FIELD,
                    &fixed64_field(1, f64::NAN.to_bits()),
                )?;
                generated
            }),
            ("inverted-bounds", {
                let mut generated = extension.payload().to_vec();
                append_length_delimited_field(
                    &mut generated,
                    VALUE_MAXIMUM_FIELD,
                    &encoded_bound(1.0, false),
                )?;
                append_length_delimited_field(
                    &mut generated,
                    VALUE_MINIMUM_FIELD,
                    &encoded_bound(2.0, false),
                )?;
                generated
            }),
            ("negative-major", {
                let mut generated = extension.payload().to_vec();
                // A canonical protobuf `int32(-1)` uses sign-extended ten-byte
                // varint framing; the semantic layer must still reject it as
                // an invalid major-step count.
                append_varint_field(&mut generated, VALUE_MAJOR_FIELD, u64::MAX)?;
                generated
            }),
            ("negative-minor", {
                let mut generated = extension.payload().to_vec();
                append_varint_field(&mut generated, VALUE_MINOR_FIELD, u64::MAX)?;
                generated
            }),
            ("noncanonical-major-varint", {
                let mut generated = extension.payload().to_vec();
                // Major field 5, value 5 encoded with a redundant continuation
                // byte.  Strict Buffa preflight rejects non-canonical varints.
                generated.extend_from_slice(&[0x28, 0x85, 0x00]);
                generated
            }),
            ("wrong-scale-wire", {
                let mut generated = extension.payload().to_vec();
                generated.extend_from_slice(&[0xC1, 0x80, 0x00, 0x01]);
                generated
            }),
        ];
        for (label, generated) in malformed {
            let replacement = length_delimited_field_with_key_width(
                GENERATED_EXTENSION_FIELD,
                &generated,
                extension.key().len(),
            );
            let payload = replace_first_field(&original, GENERATED_EXTENSION_FIELD, &replacement)?;
            let hostile = with_axis_payload(&base, VALUE_NON_STYLES[0], payload)?;
            let package = Package::from_bytes(&hostile)?;
            let result = package.edit_slide_chart_value_axis_settings(
                SlideSelector::index(0),
                ChartSelector::index(0),
            );
            let error = result.expect_err(label);
            error_is_redacted(&error);
            assert_eq!(exact_bytes(&package)?, hostile);
        }
        Ok(())
    }

    #[test]
    fn semantic_limits_fail_before_reassembly() -> TestResult<()> {
        let source = synthetic_metadata_package(false)?;
        // The package's semantic profile is checked at ingress.  A one-object
        // ceiling is intentionally below this graph and must fail before an
        // edit can clone or reassemble anything.
        let limited_semantics = SemanticLimits::new(1, 1, 1, 1, 1, 1)?;
        let limited = Package::from_bytes_with_options(
            &source,
            ReadOptions::new(Limits::default(), limited_semantics),
        );
        let error = limited.expect_err("semantic limit must reject hostile graph");
        error_is_redacted(&error);
        assert!(format!("{error:?}").len() < 512);
        Ok(())
    }

    #[test]
    fn output_limit_rejects_candidate_without_mutating_source() -> TestResult<()> {
        let source = without_previews(&synthetic_metadata_package(false)?)?;
        let target = settings(Some(1.0), Some(100.0), Some(5), Some(1), Scale::Logarithmic);
        let baseline = commit_settings(&Package::from_bytes(&source)?, target)?;
        let candidate = exact_bytes(baseline.package())?;
        assert!(candidate.len() > source.len());

        let defaults = Limits::default();
        let limits = Limits::new(
            u64::try_from(candidate.len() - 1)?,
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
        let error = package
            .edit_slide_chart_value_axis_settings(SlideSelector::index(0), ChartSelector::index(0))
            .and_then(|edit| edit.set(target).and_then(|edit| edit.commit()))
            .expect_err("candidate should exceed the package output ceiling");
        error_is_redacted(&error);
        assert!(matches!(
            error,
            ChartValueAxisError::LimitExceeded {
                kind: ChartValueAxisLimitKind::OutputBytes,
                ..
            }
        ));
        assert_eq!(exact_bytes(&package)?, before);
        Ok(())
    }
}
