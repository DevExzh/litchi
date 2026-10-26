//! Slide creation: duplicate slides and add slides from theme layouts
//! (record 0298).

use std::path::{Path, PathBuf};

use litchi_keynote::slide::creation::{Error, Layout, LayoutSelector, Unsupported};
use litchi_keynote::slide::placeholder::Kind;
use litchi_keynote::{Package, Position, SlideSelector};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;
type Texts = (Option<String>, Option<String>, Option<String>);

const MULTI_SLIDE_FIXTURE: &str = "multi-slide.key";

fn keynote_dir() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/keynote")
}

fn open(name: &str) -> TestResult<Package> {
    Ok(Package::open(keynote_dir().join(name))?)
}

fn basic() -> TestResult<Package> {
    open("basic.key")
}

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    Ok(output)
}

fn texts(package: &Package, index: usize) -> TestResult<Texts> {
    let selector = SlideSelector::index(index);
    Ok((
        package.slide_title(selector)?,
        package.slide_body(selector)?,
        package.slide_notes(selector)?,
    ))
}

fn is_empty(text: &Option<String>) -> bool {
    text.as_deref().is_none_or(str::is_empty)
}

fn duplicate(package: &Package, source: usize, destination: usize) -> TestResult<Package> {
    let mut edit = package.edit_slide_creation();
    edit.duplicate_slide(SlideSelector::index(source), Position::new(destination))?;
    Ok(edit.commit()?.into_package())
}

#[test]
fn duplicate_creates_an_equal_second_slide() -> TestResult {
    let package = basic()?;
    let source = texts(&package, 0)?;
    assert!(source.0.as_deref().is_some_and(|title| !title.is_empty()));
    let mut edit = package.edit_slide_creation();
    edit.duplicate_slide(SlideSelector::index(0), Position::new(1))?;
    let commit = edit.commit()?;
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().created_slides(), 1);
    assert!(commit.diagnostics().full_reparse_performed());
    assert!(commit.diagnostics().touched_components() >= 3);
    assert!(!commit.patch().is_noop());
    assert_eq!(commit.patch().source_slide_count(), 1);
    assert_eq!(commit.patch().target_slide_count(), 2);
    let created = commit.package();
    assert_eq!(created.slides()?.len(), 2);
    assert_eq!(texts(created, 0)?, source);
    assert_eq!(texts(created, 1)?, source);
    assert_eq!(
        created.slides()?[1].text_content(),
        package.slides()?[0].text_content()
    );
    Ok(())
}

#[test]
fn duplicate_to_the_front_and_to_the_end() -> TestResult {
    let package = basic()?;
    let first = duplicate(&package, 0, 0)?;
    assert_eq!(first.slides()?.len(), 2);
    assert_eq!(texts(&first, 0)?, texts(&package, 0)?);
    let changed = {
        let mut edit = first.edit_slide_text(SlideSelector::index(1), Kind::Title)?;
        edit.set("Original")?;
        edit.commit()?.into_package()
    };
    let front = duplicate(&changed, 1, 0)?;
    assert_eq!(front.slides()?.len(), 3);
    assert_eq!(
        front.slide_title(SlideSelector::index(0))?.as_deref(),
        Some("Original")
    );
    assert_eq!(
        front.slide_title(SlideSelector::index(2))?.as_deref(),
        Some("Original")
    );
    let end = duplicate(&changed, 1, 2)?;
    assert_eq!(
        end.slide_title(SlideSelector::index(2))?.as_deref(),
        Some("Original")
    );
    assert_ne!(
        end.slide_title(SlideSelector::index(0))?.as_deref(),
        Some("Original")
    );
    Ok(())
}

#[test]
fn several_operations_publish_one_commit() -> TestResult {
    let package = basic()?;
    let source = texts(&package, 0)?;
    let mut edit = package.edit_slide_creation();
    edit.duplicate_slide(SlideSelector::index(0), Position::new(1))?
        .add_slide("Title & Bullets", Position::new(0))?
        // The selector still names the base slide 0, now at position 1.
        .duplicate_slide(SlideSelector::index(0), Position::new(3))?
        .add_slide(LayoutSelector::Position(Position::new(0)), Position::new(4))?;
    let commit = edit.commit()?;
    assert_eq!(commit.diagnostics().created_slides(), 4);
    let created = commit.package();
    assert_eq!(created.slides()?.len(), 5);
    for index in [1, 2, 3] {
        assert_eq!(texts(created, index)?, source, "slide {index}");
    }
    for index in [0, 4] {
        let (title, body, notes) = texts(created, index)?;
        assert!(
            is_empty(&title) && is_empty(&body) && is_empty(&notes),
            "slide {index}"
        );
    }
    let reopened = Package::from_bytes(&bytes(created)?)?;
    reopened.validate()?;
    assert_eq!(reopened.slides()?.len(), 5);
    Ok(())
}

#[test]
fn duplicate_then_delete_the_original() -> TestResult {
    let package = basic()?;
    let source = texts(&package, 0)?;
    let two = duplicate(&package, 0, 1)?;
    let mut deletion = two.edit_slide_deletion();
    deletion.remove_slide(SlideSelector::index(0))?;
    let one = deletion.commit()?.into_package();
    assert_eq!(one.slides()?.len(), 1);
    assert_eq!(texts(&one, 0)?, source);
    let reopened = Package::from_bytes(&bytes(&one)?)?;
    reopened.validate()?;
    // The copy can be removed as well.
    let mut deletion = two.edit_slide_deletion();
    deletion.remove_slide(SlideSelector::index(1))?;
    assert_eq!(deletion.commit()?.package().slides()?.len(), 1);
    Ok(())
}

#[test]
fn every_basic_layout_adds_an_empty_slide() -> TestResult {
    let package = basic()?;
    let layouts = package.slide_layouts()?;
    assert_eq!(layouts.len(), 17);
    let default = package.default_slide_layout()?.ok_or("no default layout")?;
    assert!(layouts.contains(&default));
    for layout in layouts.iter() {
        assert_eq!(layouts[layout.position().get()], *layout);
        let mut edit = package.edit_slide_creation();
        edit.add_slide(layout, Position::new(1))?;
        let commit = edit.commit()?;
        let created = commit.package();
        assert_eq!(created.slides()?.len(), 2, "{:?}", layout.name());
        let (title, body, notes) = texts(created, 1)?;
        assert!(is_empty(&title), "{:?}", layout.name());
        assert!(is_empty(&body), "{:?}", layout.name());
        assert!(is_empty(&notes), "{:?}", layout.name());
        assert_eq!(texts(created, 0)?, texts(&package, 0)?);
        // Every added slide can be deleted again.
        let mut deletion = created.edit_slide_deletion();
        deletion.remove_slide(SlideSelector::index(1))?;
        assert_eq!(deletion.commit()?.package().slides()?.len(), 1);
    }
    Ok(())
}

#[test]
fn added_slides_take_text_through_the_text_transactions() -> TestResult {
    let package = basic()?;
    let mut edit = package.edit_slide_creation();
    edit.add_slide("Title & Bullets", Position::new(1))?;
    let created = edit.commit()?.into_package();
    let mut title = created.edit_slide_text(SlideSelector::index(1), Kind::Title)?;
    title.set("Fresh title")?;
    let created = title.commit()?.into_package();
    let mut notes = created.edit_slide_notes(SlideSelector::index(1))?;
    notes.set("Fresh notes")?;
    let created = notes.commit()?.into_package();
    assert_eq!(
        texts(&created, 1)?,
        (
            Some("Fresh title".to_owned()),
            Some(String::new()),
            Some("Fresh notes".to_owned())
        )
    );
    Ok(())
}

#[test]
fn commit_round_trips_through_bytes() -> TestResult {
    let package = basic()?;
    let mut edit = package.edit_slide_creation();
    edit.duplicate_slide(SlideSelector::index(0), Position::new(1))?
        .add_slide("Photo", Position::new(2))?;
    let commit = edit.commit()?;
    let output = bytes(commit.package())?;
    let reopened = Package::from_bytes(&output)?;
    reopened.validate()?;
    assert_eq!(reopened.slides()?.len(), 3);
    for index in 0..3 {
        assert_eq!(texts(&reopened, index)?, texts(commit.package(), index)?);
    }
    Ok(())
}

#[test]
fn inverse_patch_restores_the_exact_source() -> TestResult {
    let package = basic()?;
    let source = bytes(&package)?;
    let mut edit = package.edit_slide_creation();
    edit.duplicate_slide(SlideSelector::index(0), Position::new(1))?
        .add_slide("Title & Bullets", Position::new(0))?;
    let (created, patch) = edit.commit()?.into_parts();
    let restored = created.apply_slide_creation(&patch.inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    assert_eq!(restored.package().slides()?.len(), 1);
    assert_eq!(restored.diagnostics().created_slides(), 2);
    assert_eq!(patch.inverse().inverse(), patch);
    let replayed = package.apply_slide_creation(&patch)?;
    assert_eq!(bytes(replayed.package())?, bytes(&created)?);
    Ok(())
}

#[test]
fn stale_patches_conflict() -> TestResult {
    let package = basic()?;
    let mut edit = package.edit_slide_creation();
    edit.duplicate_slide(0, Position::new(1))?;
    let (created, patch) = edit.commit()?.into_parts();
    assert!(matches!(
        created.apply_slide_creation(&patch),
        Err(Error::PatchConflict)
    ));
    assert!(matches!(
        package.apply_slide_creation(&patch.inverse()),
        Err(Error::PatchConflict)
    ));
    let other = open("image-adjustments-native.key")?;
    assert!(matches!(
        other.apply_slide_creation(&patch),
        Err(Error::PatchConflict)
    ));
    Ok(())
}

#[test]
fn empty_edit_commits_a_noop() -> TestResult {
    let package = basic()?;
    let commit = package.edit_slide_creation().commit()?;
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(bytes(commit.package())?, bytes(&package)?);
    Ok(())
}

#[test]
fn selectors_and_destinations_are_checked() -> TestResult {
    let package = basic()?;
    let mut edit = package.edit_slide_creation();
    assert!(matches!(
        edit.duplicate_slide(SlideSelector::index(1), Position::new(0)),
        Err(Error::Refused(
            litchi_keynote::slide::creation::Refusal::SlidePositionNotFound { .. }
        ))
    ));
    assert!(matches!(
        edit.duplicate_slide(SlideSelector::name("missing"), Position::new(0)),
        Err(Error::Refused(
            litchi_keynote::slide::creation::Refusal::SlideNameNotFound
        ))
    ));
    assert!(matches!(
        edit.duplicate_slide(SlideSelector::index(0), Position::new(2)),
        Err(Error::Refused(
            litchi_keynote::slide::creation::Refusal::PositionOutOfRange { count: 1, .. }
        ))
    ));
    assert!(matches!(
        edit.add_slide("No such layout", Position::new(0)),
        Err(Error::Refused(
            litchi_keynote::slide::creation::Refusal::LayoutNameNotFound
        ))
    ));
    assert!(matches!(
        edit.add_slide(
            LayoutSelector::Position(Position::new(17)),
            Position::new(0)
        ),
        Err(Error::Refused(
            litchi_keynote::slide::creation::Refusal::LayoutPositionNotFound { .. }
        ))
    ));
    // A staged operation extends the destination range of the next one.
    edit.duplicate_slide(SlideSelector::index(0), Position::new(1))?;
    edit.add_slide("Blank", Position::new(2))?;
    assert!(matches!(
        edit.add_slide("Blank", Position::new(4)),
        Err(Error::Refused(
            litchi_keynote::slide::creation::Refusal::PositionOutOfRange { count: 3, .. }
        ))
    ));
    assert_eq!(edit.commit()?.package().slides()?.len(), 3);
    Ok(())
}

#[test]
fn layouts_resolve_by_name_and_position() -> TestResult {
    let package = basic()?;
    let layouts = package.slide_layouts()?;
    let bullets = layouts
        .iter()
        .find(|layout| layout.name() == Some("Title & Bullets"))
        .ok_or("missing layout")?;
    let by_name = {
        let mut edit = package.edit_slide_creation();
        edit.add_slide("Title & Bullets", Position::new(1))?;
        bytes(edit.commit()?.package())?
    };
    let by_layout = {
        let mut edit = package.edit_slide_creation();
        edit.add_slide(bullets, Position::new(1))?;
        bytes(edit.commit()?.package())?
    };
    assert_eq!(by_name, by_layout, "creation is deterministic");
    let _: &Layout = bullets;
    Ok(())
}

#[test]
fn table_and_chart_slides_are_typed_refusals() -> TestResult {
    for name in [
        "table-discovery.key",
        "chart-data-native.key",
        "chart-grid-authored-native.key",
    ] {
        let package = open(name)?;
        let mut edit = package.edit_slide_creation();
        edit.duplicate_slide(SlideSelector::index(0), Position::new(1))?;
        let error = edit.commit().err().ok_or("expected a refusal")?;
        assert!(
            matches!(
                error,
                Error::Refused(
                    litchi_keynote::slide::creation::Refusal::UnsupportedContent {
                        reason: Unsupported::ExternalContent | Unsupported::UnmodelledContent
                    }
                )
            ),
            "{name}: {error:?}"
        );
        // Adding a fresh slide from a layout still works on these decks.
        let layout = package
            .default_slide_layout()?
            .ok_or("theme without a default layout")?;
        let mut edit = package.edit_slide_creation();
        edit.add_slide(&layout, Position::new(1))?;
        assert_eq!(edit.commit()?.package().slides()?.len(), 2, "{name}");
    }
    Ok(())
}

#[test]
fn corpus_duplication_succeeds_or_refuses_typed() -> TestResult {
    let mut entries = std::fs::read_dir(keynote_dir())?.collect::<Result<Vec<_>, _>>()?;
    entries.sort_by_key(std::fs::DirEntry::path);
    let mut succeeded = 0usize;
    let mut refused = 0usize;
    for entry in entries {
        let path = entry.path();
        if path.extension().and_then(|extension| extension.to_str()) != Some("key")
            || !path.is_file()
        {
            continue;
        }
        let Ok(package) = Package::open(&path) else {
            continue;
        };
        let count = package.slides()?.len();
        let mut edit = package.edit_slide_creation();
        edit.duplicate_slide(SlideSelector::index(0), Position::new(1))?;
        match edit.commit() {
            Ok(commit) => {
                let reopened = Package::from_bytes(&bytes(commit.package())?)?;
                reopened.validate()?;
                assert_eq!(reopened.slides()?.len(), count + 1, "{path:?}");
                assert_eq!(texts(&reopened, 1)?, texts(&package, 0)?, "{path:?}");
                succeeded += 1;
            },
            Err(
                Error::Refused(litchi_keynote::slide::creation::Refusal::UnsupportedContent {
                    ..
                })
                | Error::Refused(litchi_keynote::slide::creation::Refusal::UnsupportedTopology),
            ) => refused += 1,
            Err(error) => return Err(format!("{path:?}: untyped refusal {error:?}").into()),
        }
    }
    assert!(succeeded >= 10, "succeeded {succeeded}");
    assert!(refused >= 10, "refused {refused}");
    Ok(())
}

/// Build the committed multi-slide fixture from `basic.key` with the public
/// transactions only.
fn build_multi_slide() -> TestResult<Package> {
    let package = basic()?;
    let mut creation = package.edit_slide_creation();
    creation
        .add_slide("Title & Bullets", Position::new(1))?
        .add_slide("Section", Position::new(2))?
        .add_slide("Title Only", Position::new(3))?;
    let mut package = creation.commit()?.into_package();
    for (index, role, text) in [
        (0, Kind::Title, "Litchi multi-slide fixture"),
        (1, Kind::Title, "Agenda"),
        (1, Kind::Body, "First topic\nSecond topic\nThird topic"),
        (2, Kind::Title, "Hidden section"),
        (3, Kind::Title, "Closing"),
    ] {
        let mut edit = package.edit_slide_text(SlideSelector::index(index), role)?;
        edit.set(text)?;
        package = edit.commit()?.into_package();
    }
    let mut creation = package.edit_slide_creation();
    creation.duplicate_slide(SlideSelector::index(1), Position::new(4))?;
    package = creation.commit()?.into_package();
    for (index, role, text) in [
        (4, Kind::Title, "Agenda (continued)"),
        (4, Kind::Body, "Fourth topic\nFifth topic"),
    ] {
        let mut edit = package.edit_slide_text(SlideSelector::index(index), role)?;
        edit.set(text)?;
        package = edit.commit()?.into_package();
    }
    for (index, text) in [(0, "Welcome everyone."), (1, "Walk through the agenda.")] {
        let mut edit = package.edit_slide_notes(SlideSelector::index(index))?;
        edit.set(text)?;
        package = edit.commit()?.into_package();
    }
    let mut skip = package.edit();
    skip.skip_slide(SlideSelector::index(2))?;
    Ok(skip.commit()?.into_package())
}

type SlideSummary = (Texts, bool);

fn summary(package: &Package) -> TestResult<Vec<SlideSummary>> {
    let slides = package.slides()?;
    (0..slides.len())
        .map(|index| Ok((texts(package, index)?, slides[index].is_skipped())))
        .collect()
}

#[test]
fn multi_slide_fixture_is_reproducible() -> TestResult {
    let rebuilt = build_multi_slide()?;
    let path = keynote_dir().join(MULTI_SLIDE_FIXTURE);
    if std::env::var_os("LITCHI_BLESS").is_some() {
        std::fs::write(&path, bytes(&rebuilt)?)?;
    }
    let committed = Package::open(&path)?;
    committed.validate()?;
    let expected = summary(&rebuilt)?;
    assert_eq!(summary(&committed)?, expected);
    assert_eq!(expected.len(), 5);
    let titles = expected
        .iter()
        .map(|((title, _, _), _)| title.clone().unwrap_or_default())
        .collect::<Vec<_>>();
    let mut distinct = titles.clone();
    distinct.sort();
    distinct.dedup();
    assert_eq!(distinct.len(), 5, "distinct titles: {titles:?}");
    assert_eq!(
        expected.iter().filter(|(_, skipped)| *skipped).count(),
        1,
        "one hidden slide"
    );
    assert!(
        expected
            .iter()
            .filter(|((title, body, _), _)| !is_empty(title) && !is_empty(body))
            .count()
            >= 2
    );
    assert!(expected.iter().any(|((_, _, notes), _)| !is_empty(notes)));
    assert!(std::fs::metadata(&path)?.len() < 1024 * 1024);
    Ok(())
}

#[test]
fn slide_layout_list_names_each_slides_layout() -> TestResult {
    let package = basic()?;
    let layouts = package.slide_layouts()?;
    let assigned = package.slide_layout_list()?;
    assert_eq!(assigned.len(), 1);
    let first = assigned[0].as_ref().ok_or("basic.key slide has a layout")?;
    assert!(layouts.contains(first));
    // A slide added from a layout reports that layout.
    let last = layouts.last().ok_or("layouts")?;
    let mut edit = package.edit_slide_creation();
    edit.add_slide(last, Position::new(1))?;
    let added = edit.commit()?.into_package();
    let assigned = added.slide_layout_list()?;
    assert_eq!(assigned.len(), 2);
    assert_eq!(assigned[1].as_ref(), Some(last));
    // A duplicate keeps its source's layout.
    let duplicated = duplicate(&package, 0, 1)?;
    let assigned = duplicated.slide_layout_list()?;
    assert_eq!(assigned[0], assigned[1]);
    Ok(())
}

#[test]
fn set_layout_reassigns_a_slide_and_keeps_its_text() -> TestResult {
    let package = basic()?;
    let before = texts(&package, 0)?;
    for layout in package.slide_layouts()?.iter() {
        let name = layout.name().ok_or("unnamed layout")?;
        let mut edit = package.edit_slide_creation();
        edit.set_layout(SlideSelector::index(0), name)?;
        let commit = edit.commit()?;
        let edited = commit.package();
        let assigned = edited.slide_layout_list()?;
        assert_eq!(assigned[0].as_ref().and_then(Layout::name), Some(name));
        assert_eq!(texts(edited, 0)?, before, "{name}: slide text changed");
        assert_eq!(edited.slides()?.len(), package.slides()?.len());
        let restored = edited.apply_slide_creation(&commit.patch().inverse())?;
        assert_eq!(bytes(restored.package())?, bytes(&package)?);
    }
    // Reassigning a slide to its own layout stages nothing.
    let current = package.slide_layout_list()?[0]
        .clone()
        .ok_or("slide 0 has a layout")?;
    let mut edit = package.edit_slide_creation();
    edit.set_layout(
        SlideSelector::index(0),
        LayoutSelector::Position(current.position()),
    )?;
    let commit = edit.commit()?;
    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    Ok(())
}
