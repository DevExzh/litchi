//! Sheet-drawable families on native fixtures (record 0378): the drawable
//! list, frame/property/playback edits and z-order moves, their exact
//! inverses and their refusals.

use std::path::PathBuf;
use std::time::Duration;

use litchi_core::Position;
use litchi_core::edit::RefusalKind;
use litchi_numbers::Package;
use litchi_numbers::sheet::drawable::{
    self, DrawableSelector, FlipAxis, Kind, MediaLoopMode, MediaPlaybackSettings, Point,
    Properties, properties,
};
#[path = "support/chart_arrangement_fixture.rs"]
mod chart_arrangement_fixture;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn open(name: &str) -> TestResult<Package> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers")
        .join(name);
    Ok(Package::open(path)?)
}

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    Ok(output)
}

#[test]
fn the_list_reports_kinds_in_back_to_front_order() -> TestResult {
    let package = open("image-adjustments-native.numbers")?;
    let list = package.sheet_drawable_list(0)?;
    assert!(list.iter().any(|item| item.kind() == Kind::Table));
    let image = list
        .iter()
        .find(|item| item.kind() == Kind::Image)
        .ok_or("no image")?;
    assert_eq!(image.kind_selector(), DrawableSelector::image(0));
    assert!(image.geometry().is_some());
    for (position, item) in list.iter().enumerate() {
        assert_eq!(item.z_position(), Position::new(position));
        assert_eq!(item.selector(), DrawableSelector::index(position));
    }
    let audio = open("audio-playback-native.numbers")?.sheet_drawable_list(0)?;
    assert!(audio.iter().any(|item| item.kind() == Kind::Audio));
    Ok(())
}

#[test]
fn a_frame_edit_round_trips_and_its_inverse_restores_the_source() -> TestResult {
    let package = open("image-adjustments-native.numbers")?;
    let before = package.sheet_drawable_properties(0, DrawableSelector::image(0))?;
    let mut edit = package.edit_sheet_drawable_properties(0, DrawableSelector::image(0))?;
    edit.set_position(Point { x: 12.0, y: 34.0 })?
        .flip(FlipAxis::Horizontal)?
        .set_properties(Properties {
            locked: Some(true),
            ..before.properties().clone()
        });
    let commit = edit.commit()?;
    assert!(commit.diagnostics().changed());
    let after = commit
        .package()
        .sheet_drawable_properties(0, DrawableSelector::image(0))?;
    let frame = after.geometry().ok_or("no frame")?;
    assert_eq!(frame.position, Some(Point { x: 12.0, y: 34.0 }));
    assert!(frame.flags.ok_or("no flags")?.is_reflected());
    assert_eq!(after.properties().locked, Some(true));
    let restored = commit
        .package()
        .apply_sheet_drawable_properties(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, bytes(&package)?);
    assert!(matches!(
        package.apply_sheet_drawable_properties(&commit.patch().inverse()),
        Err(properties::Error::PatchConflict)
    ));
    Ok(())
}

#[test]
fn an_unchanged_edit_is_an_exact_no_op() -> TestResult {
    let package = open("image-adjustments-native.numbers")?;
    let commit = package
        .edit_sheet_drawable_properties(0, DrawableSelector::image(0))?
        .commit()?;
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(bytes(commit.package())?, bytes(&package)?);
    Ok(())
}

#[test]
fn audio_playback_is_edited_through_the_same_family() -> TestResult {
    let package = open("audio-playback-native.numbers")?;
    let settings = MediaPlaybackSettings {
        start_time: Some(Duration::from_millis(250)),
        end_time: Duration::from_millis(1_000),
        poster_time: None,
        loop_mode: Some(MediaLoopMode::Repeat),
        volume: None,
    };
    let mut edit = package.edit_sheet_drawable_properties(0, DrawableSelector::audio(0))?;
    edit.set_playback(settings)?;
    let commit = edit.commit()?;
    let after = commit
        .package()
        .sheet_drawable_properties(0, DrawableSelector::audio(0))?;
    assert_eq!(after.playback(), Some(settings.canonicalize()?));
    Ok(())
}

#[test]
fn verbs_that_do_not_apply_are_typed_refusals() -> TestResult {
    let package = open("audio-playback-native.numbers")?;
    let mut edit = package.edit_sheet_drawable_properties(0, DrawableSelector::audio(0))?;
    let error = edit.flip(FlipAxis::Vertical).map(drop).unwrap_err();
    let properties::Error::Refused(refusal) = error else {
        return Err(format!("unexpected {error:?}").into());
    };
    assert_eq!(refusal.kind(), RefusalKind::UnsupportedOperation);
    assert!(matches!(
        package.sheet_drawable_properties(0, DrawableSelector::image(5)),
        Err(properties::Error::Refused(
            properties::Refusal::DrawableNotFound { .. }
        ))
    ));
    assert!(matches!(
        package.sheet_drawable_properties("No such sheet", 0),
        Err(properties::Error::Refused(
            properties::Refusal::SheetNotFound
        ))
    ));
    Ok(())
}

#[test]
fn z_order_moves_reorder_the_list_and_tables_are_not_removed() -> TestResult {
    let package = open("image-adjustments-native.numbers")?;
    let list = package.sheet_drawable_list(0)?;
    let last = list.len() - 1;
    let mut edit = package.edit_sheet_drawables(0)?;
    edit.move_to(DrawableSelector::index(last), Position::new(0))?;
    let commit = edit.commit()?;
    let moved = commit.package().sheet_drawable_list(0)?;
    assert_eq!(moved[0].kind(), list[last].kind());
    let restored = commit
        .package()
        .apply_sheet_drawables(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, bytes(&package)?);

    let mut edit = package.edit_sheet_drawables(0)?;
    let table = list
        .iter()
        .find(|item| item.kind() == Kind::Table)
        .ok_or("no table")?;
    assert!(matches!(
        edit.remove(table.selector()).map(drop),
        Err(drawable::Error::Refused(
            drawable::Refusal::UnsupportedOperation { .. }
        ))
    ));
    assert!(matches!(
        edit.move_to(DrawableSelector::index(0), Position::new(list.len()))
            .map(drop),
        Err(drawable::Error::Refused(
            drawable::Refusal::ZPositionOutOfRange { .. }
        ))
    ));
    assert!(matches!(
        package.edit_sheet_drawables(0)?.commit(),
        Err(drawable::Error::Refused(
            drawable::Refusal::NoStagedOperation
        ))
    ));
    Ok(())
}

#[test]
fn a_created_text_box_takes_text_properties_and_can_be_removed() -> TestResult {
    use litchi_numbers::sheet::drawable::text::font::Font;
    use litchi_numbers::sheet::drawable::text::{Property, PropertyKind};
    use litchi_numbers::sheet::drawable::{Geometry, Size};

    let package = open("basic.numbers")?;
    let before = package.sheet_drawable_list(0)?.len();
    let mut edit = package.edit_sheet_drawables(0)?;
    edit.add_text_box(
        Geometry::new(
            Point { x: 40.0, y: 60.0 },
            Size {
                width: 200.0,
                height: 50.0,
            },
        ),
        "Leaf text box",
    )?;
    let created = edit.commit()?.into_package();
    let list = created.sheet_drawable_list(0)?;
    assert_eq!(list.len(), before + 1);
    assert_eq!(list[before].kind(), Kind::TextBox);

    let alignment =
        created.sheet_text_box_text(0, DrawableSelector::text_box(0), PropertyKind::Alignment)?;
    let mut edit = created.edit_sheet_text_box_text(0, DrawableSelector::text_box(0))?;
    edit.set(Property::TextFont(Font::named("Courier")?))?;
    let styled = edit.commit()?;
    assert_eq!(
        styled.package().sheet_text_box_text(
            0,
            DrawableSelector::text_box(0),
            PropertyKind::Alignment
        )?,
        alignment
    );
    let restored = styled
        .package()
        .apply_sheet_text_box_text(&styled.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, bytes(&created)?);

    let mut edit = created.edit_sheet_drawables(0)?;
    edit.remove(DrawableSelector::text_box(0))?;
    let removed = edit.commit()?.into_package();
    assert_eq!(removed.sheet_drawable_list(0)?.len(), before);
    Ok(())
}

#[test]
fn a_text_box_named_style_is_created_applied_renamed_and_deleted() -> TestResult {
    use litchi_numbers::sheet::drawable::named_style::{ParagraphStyleName, Refusal};
    use litchi_numbers::sheet::drawable::{Geometry, Size, named_style};

    let package = open("basic.numbers")?;
    let mut edit = package.edit_sheet_drawables(0)?;
    edit.add_text_box(
        Geometry::new(
            Point { x: 40.0, y: 60.0 },
            Size {
                width: 200.0,
                height: 50.0,
            },
        ),
        "Named",
    )?;
    let created = edit.commit()?.into_package();
    let text_box = DrawableSelector::text_box(0);
    let styles = created.sheet_text_box_named_styles(0, text_box)?;
    let (applied, overridden) = created.sheet_text_box_applied_named_style(0, text_box)?;
    assert!(!overridden);
    assert!(styles.contains(&applied));

    let mut edit = created.edit_sheet_text_box_named_styles(0, text_box)?;
    edit.create(applied.id(), ParagraphStyleName::new("Leaf Copy")?)?;
    let copied = edit.commit()?;
    let copy = copied
        .package()
        .sheet_text_box_named_styles(0, text_box)?
        .into_iter()
        .find(|style| style.name().as_str() == "Leaf Copy")
        .ok_or("no copy")?;
    let restored = copied
        .package()
        .apply_sheet_text_box_named_styles(&copied.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, bytes(&created)?);

    let mut edit = copied
        .package()
        .edit_sheet_text_box_named_styles(0, text_box)?;
    edit.apply(copy.id())?
        .rename_style(copy.id(), ParagraphStyleName::new("Leaf Renamed")?)?;
    let renamed = edit.commit()?;
    let (now, _) = renamed
        .package()
        .sheet_text_box_applied_named_style(0, text_box)?;
    assert_eq!(now.id(), copy.id());
    assert_eq!(now.name().as_str(), "Leaf Renamed");

    let mut edit = renamed
        .package()
        .edit_sheet_text_box_named_styles(0, text_box)?;
    edit.delete(copy.id())?;
    let Err(named_style::Error::Refused(refusal)) = edit.commit() else {
        return Err("deleting an applied style must be refused".into());
    };
    assert_eq!(refusal, Refusal::StyleInUse);
    assert_eq!(refusal.kind(), RefusalKind::UnsupportedContent);

    let mut edit = renamed
        .package()
        .edit_sheet_text_box_named_styles(0, text_box)?;
    edit.delete_applied(copy.id(), applied.id())?;
    let deleted = edit.commit()?.into_package();
    assert_eq!(deleted.sheet_text_box_named_styles(0, text_box)?, styles);
    assert_eq!(
        deleted.sheet_text_box_applied_named_style(0, text_box)?,
        (applied, false)
    );

    let mut edit = deleted.edit_sheet_text_box_named_styles(0, text_box)?;
    edit.redefine()?;
    let Err(named_style::Error::Refused(Refusal::NoOverrides)) = edit.commit() else {
        return Err("redefining without overrides must be refused".into());
    };
    Ok(())
}

#[test]
fn a_drawable_comment_thread_is_written_and_cleared() -> TestResult {
    use litchi_numbers::sheet::drawable::comment::{self, Refusal};
    use litchi_numbers::sheet::drawable::{Geometry, Size};

    let package = open("basic.numbers")?;
    let mut edit = package.edit_sheet_drawables(0)?;
    edit.add_text_box(
        Geometry::new(
            Point { x: 40.0, y: 60.0 },
            Size {
                width: 200.0,
                height: 50.0,
            },
        ),
        "Commented",
    )?;
    let created = edit.commit()?.into_package();
    let text_box = DrawableSelector::text_box(0);
    assert_eq!(created.sheet_drawable_comment(0, text_box)?, None);

    let mut edit = created.edit_sheet_drawable_comment(0, text_box)?;
    edit.add_reply("orphan")?;
    let Err(comment::Error::Refused(refusal)) = edit.commit() else {
        return Err("a reply without a comment must be refused".into());
    };
    assert_eq!(refusal, Refusal::NoComment);
    assert_eq!(refusal.kind(), RefusalKind::NotFound);

    let mut edit = created.edit_sheet_drawable_comment(0, text_box)?;
    edit.set("Root")?
        .add_reply("First")?
        .add_reply("Second")?
        .set_reply(0, "Edited")?
        .remove_reply(1)?;
    let threaded = edit.commit()?;
    let comment = threaded
        .package()
        .sheet_drawable_comment(0, text_box)?
        .ok_or("no comment")?;
    assert_eq!(comment.body(), "Root");
    assert_eq!(comment.replies(), ["Edited".to_owned()]);
    let restored = threaded
        .package()
        .apply_sheet_drawable_comment(&threaded.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, bytes(&created)?);

    let mut edit = threaded
        .package()
        .edit_sheet_drawable_comment(0, text_box)?;
    edit.clear()?;
    let cleared = edit.commit()?.into_package();
    assert_eq!(cleared.sheet_drawable_comment(0, text_box)?, None);

    let Err(comment::Error::Refused(Refusal::NotCommentable { .. })) =
        created.edit_sheet_drawable_comment(0, DrawableSelector::of(Kind::Table, 0))
    else {
        return Err("a table takes no direct comment".into());
    };
    Ok(())
}

#[test]
fn a_sheet_image_takes_new_data_and_its_inverse_restores_the_asset() -> TestResult {
    use litchi_numbers::sheet::drawable::media::{self, Refusal};

    let package = open("image-adjustments-native.numbers")?;
    let assets = package.media_assets()?;
    assert!(!assets.is_empty());
    assert!(
        package
            .sheet_media_assets(0)?
            .iter()
            .all(|asset| assets.contains(asset))
    );
    let mut edit = package.edit_media();
    edit.replace_data(0, DrawableSelector::image(0), b"litchi replacement")?;
    let replaced = edit.commit()?;
    let changed: Vec<_> = replaced
        .package()
        .media_assets()?
        .into_iter()
        .filter(|asset| !assets.contains(asset))
        .collect();
    let [asset] = changed.as_slice() else {
        return Err("exactly one asset changes".into());
    };
    assert_eq!(
        replaced.package().extract_media(asset.data())?,
        b"litchi replacement"
    );
    let restored = replaced
        .package()
        .apply_media(&replaced.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, bytes(&package)?);

    let mut edit = package.edit_media();
    let Err(media::Error::Refused(Refusal::EmptyData)) =
        edit.replace_data(0, DrawableSelector::image(0), b"")
    else {
        return Err("empty data must be refused".into());
    };
    let created = open("basic.numbers")?;
    let mut edit = created.edit_media();
    edit.replace_poster(0, DrawableSelector::of(Kind::Table, 0), b"x")?;
    let Err(media::Error::Refused(Refusal::NoMedia { .. })) = edit.commit() else {
        return Err("a table has no poster".into());
    };
    Ok(())
}

#[test]
fn a_duplicate_shares_its_media_and_only_text_boxes_take_new_text() -> TestResult {
    let package = open("image-adjustments-native.numbers")?;
    let before = package.sheet_drawable_list(0)?.len();
    let assets = package.media_assets()?;
    let mut edit = package.edit_sheet_drawables(0)?;
    edit.duplicate(DrawableSelector::image(0))?;
    let duplicated = edit.commit()?;
    let list = duplicated.package().sheet_drawable_list(0)?;
    assert_eq!(list.len(), before + 1);
    assert_eq!(list[before].kind(), Kind::Image);
    let shared = |assets: &[drawable::media::Embedded]| {
        assets.iter().map(|asset| asset.data()).collect::<Vec<_>>()
    };
    assert_eq!(
        shared(&duplicated.package().media_assets()?),
        shared(&assets)
    );
    let restored = duplicated
        .package()
        .apply_sheet_drawables(&duplicated.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, bytes(&package)?);

    let mut edit = package.edit_sheet_drawables(0)?;
    let Err(drawable::Error::Refused(drawable::Refusal::UnsupportedOperation { .. })) =
        edit.duplicate_text_box(DrawableSelector::image(0), "x")
    else {
        return Err("an image is not a text box".into());
    };
    Ok(())
}

#[test]
fn a_new_image_embeds_its_data_and_bad_media_is_refused() -> TestResult {
    use litchi_numbers::sheet::drawable::{ImageOptions, Size};

    let package = open("basic.numbers")?;
    let png = b"\x89PNG\r\n\x1a\nlitchi".as_slice();
    let options = ImageOptions::new(
        Point { x: 10.0, y: 20.0 },
        Size {
            width: 30.0,
            height: 40.0,
        },
    );
    let mut edit = package.edit_sheet_drawables(0)?;
    edit.add_image("dot.png", png, options)?;
    let created = edit.commit()?.into_package();
    let list = created.sheet_drawable_list(0)?;
    assert_eq!(list.last().map(drawable::Drawable::kind), Some(Kind::Image));
    let asset = created
        .media_assets()?
        .into_iter()
        .find(|asset| asset.preferred_file_name() == "dot.png")
        .ok_or("no asset")?;
    assert_eq!(created.extract_media(asset.data())?, png);

    let mut edit = package.edit_sheet_drawables(0)?;
    for (name, data) in [
        ("dot.mov", png),
        ("dot.png", b"".as_slice()),
        ("../dot.png", png),
    ] {
        let Err(drawable::Error::Refused(drawable::Refusal::InvalidMedia)) =
            edit.add_image(name, data, options)
        else {
            return Err(format!("{name} must be refused").into());
        };
    }
    Ok(())
}

#[test]
fn chart_series_values_round_trip_and_name_every_series() -> TestResult {
    use litchi_numbers::sheet::chart::series::{Property, PropertyKind, Refusal, error_bar};

    let package = open("chart-data-native.numbers")?;
    let Property::ErrorBars(bars) = package.sheet_chart_series(0, 0, PropertyKind::ErrorBars)?
    else {
        return Err("error bars".into());
    };
    let mut edit = package.edit_sheet_chart_series(0, 0)?;
    edit.set(Property::ErrorBars(bars.clone()))?;
    assert!(!edit.commit()?.diagnostics().changed());

    let mut wrong = bars.clone();
    wrong.push(error_bar::Series::None);
    let mut edit = package.edit_sheet_chart_series(0, 0)?;
    edit.set(Property::ErrorBars(wrong))?;
    let Err(litchi_numbers::sheet::chart::series::Error::Refused(Refusal::SeriesCount {
        expected,
        given,
    })) = edit.commit()
    else {
        return Err("a value for another series count must be refused".into());
    };
    assert_eq!((expected, given), (bars.len(), bars.len() + 1));

    let mut changed = bars.clone();
    changed[0] = error_bar::Series::fixed_value(error_bar::Direction::PositiveAndNegative);
    let source = bytes(&package)?;
    let mut edit = package.edit_sheet_chart_series(0, 0)?;
    edit.set(Property::ErrorBars(changed))?;
    let commit = edit.commit()?;
    assert!(commit.diagnostics().changed());
    let Property::ErrorBars(after) =
        commit
            .package()
            .sheet_chart_series(0, 0, PropertyKind::ErrorBars)?
    else {
        return Err("error bars changed kind".into());
    };
    assert!(matches!(
        after.first(),
        Some(error_bar::Series::FixedValue { .. })
    ));
    let restored = commit
        .package()
        .apply_sheet_chart_series(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);

    Ok(())
}

#[test]
fn chart_series_reference_ceiling_is_typed_before_graph_publication() -> TestResult {
    use litchi_numbers::sheet::chart::series::{Error as SeriesError, PropertyKind};
    use litchi_numbers::{OpenOptions, SemanticLimits};

    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/chart-data-native.numbers");
    let bytes = std::fs::read(path)?;
    let semantic = SemanticLimits::builder().max_references(1)?.build()?;
    let package = match Package::from_bytes_with_options(
        &bytes,
        OpenOptions::new().semantic_limits(semantic),
    ) {
        Ok(package) => package,
        Err(litchi_numbers::Error::SemanticLimit { .. }) => return Ok(()),
        Err(error) => return Err(error.into()),
    };
    let result = package.sheet_chart_series(0, 0, PropertyKind::ErrorBars);
    assert!(matches!(result, Err(SeriesError::Limit { .. })));
    Ok(())
}

#[test]
fn chart_series_symbols_scope_reopen_inverse_and_noop() -> TestResult {
    use litchi_numbers::sheet::chart::series::{
        Error as SeriesError, Property, PropertyKind, SymbolShape,
    };

    let mut selected = None;
    let mut fixture_source = None;
    for name in [
        "chart-grid-authored-native.numbers",
        "chart-data-native.numbers",
        "chart-data-edited-native.numbers",
        "chart-grid-replaced-native.numbers",
        "chart-arrangement-native.numbers",
    ] {
        let package = open(name)?;
        match package.sheet_chart_series(0, 0, PropertyKind::Symbols) {
            Ok(Property::Symbols(values)) => {
                selected = Some((package, values));
                break;
            },
            Ok(_) => return Err("symbols read returned another property kind".into()),
            Err(SeriesError::Refused(_)) => continue,
            Err(error) => return Err(error.into()),
        }
    }
    if selected.is_none() {
        let source = chart_arrangement_fixture::fixture()?;
        let payload = chart_arrangement_fixture::chart_payload_from_source(&source, 0)?;
        let archive_field = litchi_iwa_common::wire::WireView::parse(
            &payload,
            litchi_iwa_common::WireLimits::DEFAULT,
        )?
        .fields()
        .find(|field| field.number() == chart_arrangement_fixture::CHART_EXTENSION_FIELD)
        .ok_or("missing chart archive extension")?;
        let line_type = litchi_iwa_protos::tsch::ChartType::LineChartType2D as u64;
        let mut rewritten_archive = Vec::new();
        for field in litchi_iwa_common::wire::WireView::parse(
            archive_field.payload(),
            litchi_iwa_common::WireLimits::DEFAULT,
        )?
        .fields()
        {
            if field.number() == 1 {
                litchi_iwa_common::wire::append_varint_field(&mut rewritten_archive, 1, line_type)?;
            } else {
                rewritten_archive.extend_from_slice(field.raw());
            }
        }
        let mut rewritten = Vec::new();
        for field in litchi_iwa_common::wire::WireView::parse(
            &payload,
            litchi_iwa_common::WireLimits::DEFAULT,
        )?
        .fields()
        {
            if field.number() == chart_arrangement_fixture::CHART_EXTENSION_FIELD {
                litchi_iwa_common::wire::append_length_delimited_field(
                    &mut rewritten,
                    chart_arrangement_fixture::CHART_EXTENSION_FIELD,
                    &rewritten_archive,
                    litchi_iwa_common::WireLimits::DEFAULT,
                )?;
            } else {
                rewritten.extend_from_slice(field.raw());
            }
        }
        let source = chart_arrangement_fixture::with_chart_payload(&source, 0, rewritten)?;
        let source_package = Package::from_bytes(&source)?;
        fixture_source = Some(source.clone());
        selected = Some((
            source_package,
            match Package::from_bytes(&source)?.sheet_chart_series(0, 0, PropertyKind::Symbols) {
                Ok(Property::Symbols(values)) => values,
                Ok(_) => return Err("source-built symbols kind mismatch".into()),
                Err(error) => return Err(format!("source-built symbols failed: {error:?}").into()),
            },
        ));
    }
    let Some((package, values)) = selected else {
        return Err("no Numbers fixture exposes chart symbols".into());
    };
    let source = bytes(&package)?;
    let mut edit = package
        .edit_sheet_chart_series(0, 0)
        .map_err(|error| format!("edit noop: {error:?}"))?;
    edit.set(Property::Symbols(values.clone()))
        .map_err(|error| format!("stage noop: {error:?}"))?;
    let noop = edit
        .commit()
        .map_err(|error| format!("commit noop: {error:?}"))?;
    assert!(!noop.diagnostics().changed());
    assert_eq!(bytes(noop.package())?, source);

    let mut changed = values;
    if let Some(first) = changed.first_mut() {
        *first = Some(litchi_iwa_common::chart::series::Symbol::automatic(
            SymbolShape::Triangle,
        ));
    }
    let mut edit = package
        .edit_sheet_chart_series(0, 0)
        .map_err(|error| format!("edit changed: {error:?}"))?;
    edit.set(Property::Symbols(changed))
        .map_err(|error| format!("stage changed: {error:?}"))?;
    let commit = edit
        .commit()
        .map_err(|error| format!("commit changed: {error:?}"))?;
    assert!(commit.diagnostics().changed());
    if let Some(source) = fixture_source {
        let candidate = bytes(commit.package())?;
        assert_eq!(
            chart_arrangement_fixture::component_stream(
                &source,
                chart_arrangement_fixture::UNRELATED_MEMBER,
            )?,
            chart_arrangement_fixture::component_stream(
                &candidate,
                chart_arrangement_fixture::UNRELATED_MEMBER,
            )?
        );
        let source_central = chart_arrangement_fixture::central_record(
            &source,
            chart_arrangement_fixture::UNRELATED_MEMBER,
        )?;
        let candidate_central = chart_arrangement_fixture::central_record(
            &candidate,
            chart_arrangement_fixture::UNRELATED_MEMBER,
        )?;
        assert_eq!(&source_central[..42], &candidate_central[..42]);
        assert_eq!(&source_central[46..], &candidate_central[46..]);
        assert_eq!(
            chart_arrangement_fixture::raw_drawable_fields(
                &source,
                0,
                chart_arrangement_fixture::UNKNOWN_DRAWABLE_BYTES_FIELD,
            )?,
            chart_arrangement_fixture::raw_drawable_fields(
                &candidate,
                0,
                chart_arrangement_fixture::UNKNOWN_DRAWABLE_BYTES_FIELD,
            )?
        );
        let malformed = chart_arrangement_fixture::with_parent_mismatch(&source)?;
        let malformed_package = Package::from_bytes(&malformed)?;
        assert!(
            malformed_package
                .sheet_chart_series(0, 0, PropertyKind::Symbols)
                .is_err()
        );
        let foreign_style = chart_arrangement_fixture::with_foreign_chart_series_style(&source)?;
        let foreign_style_package = Package::from_bytes(&foreign_style)?;
        let mut foreign_edit = foreign_style_package.edit_sheet_chart_series(0, 0)?;
        let mut foreign_values =
            foreign_style_package.sheet_chart_series(0, 0, PropertyKind::Symbols)?;
        let Property::Symbols(values) = &mut foreign_values else {
            return Err("foreign style symbols kind mismatch".into());
        };
        if let Some(first) = values.first_mut() {
            *first = Some(litchi_iwa_common::chart::series::Symbol::automatic(
                SymbolShape::Triangle,
            ));
        }
        foreign_edit.set(foreign_values)?;
        let foreign_commit = foreign_edit
            .commit()
            .map_err(|error| format!("shared style commit: {error:?}"))?;
        let foreign_candidate = bytes(foreign_commit.package())?;
        assert_eq!(
            chart_arrangement_fixture::object_message(
                &source,
                chart_arrangement_fixture::CALCULATION_MEMBER,
                chart_arrangement_fixture::ids(1).drawable,
                chart_arrangement_fixture::CHART_MESSAGE_TYPE,
            )?,
            chart_arrangement_fixture::object_message(
                &foreign_candidate,
                chart_arrangement_fixture::CALCULATION_MEMBER,
                chart_arrangement_fixture::ids(1).drawable,
                chart_arrangement_fixture::CHART_MESSAGE_TYPE,
            )?
        );

        let source_limits = litchi_numbers::SourceLimits::builder()
            .max_input_bytes(source.len() as u64)?
            .build()?;
        let limited = Package::from_bytes_with_options(
            &source,
            litchi_numbers::OpenOptions::new().source_limits(source_limits),
        )?;
        match limited.sheet_chart_series(0, 0, PropertyKind::Symbols) {
            Err(SeriesError::Limit { .. }) => {
                assert_eq!(bytes(&limited)?, source);
            },
            Ok(Property::Symbols(mut limited_values)) => {
                if let Some(first) = limited_values.first_mut() {
                    *first = Some(litchi_iwa_common::chart::series::Symbol::automatic(
                        SymbolShape::Triangle,
                    ));
                }
                let mut limited_edit = limited.edit_sheet_chart_series(0, 0)?;
                limited_edit.set(Property::Symbols(limited_values))?;
                assert!(matches!(
                    limited_edit.commit(),
                    Err(SeriesError::Limit { .. })
                ));
                assert_eq!(bytes(&limited)?, source);
            },
            Ok(_) => return Err("limited symbols kind mismatch".into()),
            Err(error) => return Err(format!("unexpected low-limit error: {error:?}").into()),
        }
    }
    let restored = commit
        .package()
        .apply_sheet_chart_series(&commit.patch().inverse())
        .map_err(|error| format!("apply inverse: {error:?}"))?;
    assert_eq!(bytes(restored.package())?, source);
    Ok(())
}
