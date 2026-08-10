#![cfg(feature = "keynote")]

use std::io;
use std::path::PathBuf;

use litchi::keynote::show::{
    Commit, Diagnostics, Edit, Error, LimitKind, Mode, Patch, Settings, Size,
};
use litchi::keynote::slide::placeholder::Kind;
use litchi::keynote::{Package, Position, SlideNotesError, SlideSelector, TextSpan};

trait ExactPackageBytes {
    fn exact_bytes(&self) -> &'static [u8];
}

impl ExactPackageBytes for Package {
    fn exact_bytes(&self) -> &'static [u8] {
        let mut bytes = Vec::new();
        self.write_to(&mut bytes)
            .expect("an in-memory Vec accepts every package byte");
        Box::leak(bytes.into_boxed_slice())
    }
}

fn assert_send_sync<T: Send + Sync>() {}

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/keynote/basic.key")
}

#[test]
fn slide_skip_state_is_available_through_the_root_facade() -> Result<(), Box<dyn std::error::Error>>
{
    let package = Package::open(fixture_path())?;
    let is_skipped = package
        .slides()?
        .first()
        .ok_or_else(|| io::Error::other("native Keynote file has no slide"))?
        .is_skipped();

    let mut edit = package.edit();
    edit.set_slide_skipped(SlideSelector::index(0), is_skipped)?;
    let commit = edit.commit()?;

    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.package().exact_bytes(), package.exact_bytes());
    Ok(())
}

#[test]
fn slide_order_transaction_is_available_through_the_root_facade()
-> Result<(), Box<dyn std::error::Error>> {
    let package = Package::open(fixture_path())?;
    let source_snapshot = package.exact_bytes();
    let mut edit = package.edit_slide_order();
    edit.move_slide(SlideSelector::index(0), Position::new(0))?;
    let commit = edit.commit()?;

    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.package().exact_bytes(), source_snapshot);

    let reapplied = package.apply_slide_order(commit.patch())?;
    assert!(reapplied.patch().is_noop());
    assert_eq!(reapplied.package().exact_bytes(), source_snapshot);
    Ok(())
}

#[test]
fn show_settings_transaction_is_available_through_the_root_facade()
-> Result<(), Box<dyn std::error::Error>> {
    assert_send_sync::<Settings>();
    assert_send_sync::<Mode>();
    assert_send_sync::<Size>();
    assert_send_sync::<Edit<'static>>();
    assert_send_sync::<Patch>();
    assert_send_sync::<Commit>();
    assert_send_sync::<Diagnostics>();
    assert_send_sync::<Error>();
    assert_send_sync::<LimitKind>();

    let package = Package::open(fixture_path())?;
    let source_snapshot = package.exact_bytes();
    let before = package.show_settings()?;
    assert_eq!(before, *package.show()?.settings());

    let noop = package.edit_show_settings()?.set(before).commit()?;
    assert_eq!(noop.patch().before(), before);
    assert_eq!(noop.patch().after(), before);
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert!(!noop.diagnostics().full_reparse_performed());
    assert_eq!(noop.package().exact_bytes(), source_snapshot);

    let mut after = before;
    after.set_loop_presentation(Some(!before.loop_presentation().unwrap_or(false)));
    let changed = package.edit_show_settings()?.set(after).commit()?;
    assert_eq!(changed.patch().before(), before);
    assert_eq!(changed.patch().after(), after);
    assert_eq!(changed.package().show_settings()?, after);
    assert!(changed.diagnostics().changed());
    assert!(changed.diagnostics().touched_components() >= 1);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_eq!(package.exact_bytes(), source_snapshot);
    assert_ne!(changed.package().exact_bytes(), source_snapshot);

    let restored = changed
        .package()
        .apply_show_settings(&changed.patch().inverse())?;
    assert_eq!(restored.package().exact_bytes(), source_snapshot);
    assert_eq!(restored.package().show_settings()?, before);
    Ok(())
}

#[test]
fn slide_transition_transaction_is_available_through_the_root_facade()
-> Result<(), Box<dyn std::error::Error>> {
    use litchi::keynote::transition::{
        Commit, Diagnostics, Edit, Effect, Error, LimitKind, Patch, Settings,
    };

    assert_send_sync::<Settings>();
    assert_send_sync::<Edit<'static>>();
    assert_send_sync::<Patch>();
    assert_send_sync::<Commit>();
    assert_send_sync::<Diagnostics>();
    assert_send_sync::<Error>();
    assert_send_sync::<LimitKind>();

    let package = Package::open(fixture_path())?;
    let source_snapshot = package.exact_bytes();
    let selector = SlideSelector::index(0);
    let before = package
        .slide_transition(selector)?
        .ok_or_else(|| io::Error::other("native Keynote file has no editable transition"))?;

    let noop = package
        .edit_slide_transition(selector)?
        .set(before.clone())?
        .commit()?;
    assert_eq!(noop.patch().before(), Some(&before));
    assert_eq!(noop.patch().after(), Some(&before));
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert!(!noop.diagnostics().full_reparse_performed());
    assert_eq!(noop.package().exact_bytes(), source_snapshot);

    let mut after = before.clone();
    let replacement = if after.effect() == Some(&Effect::Dissolve) {
        Effect::None
    } else {
        Effect::Dissolve
    };
    after.set_effect(Some(replacement))?;
    let changed = package
        .edit_slide_transition(selector)?
        .set(after.clone())?
        .commit()?;
    assert_eq!(changed.patch().before(), Some(&before));
    assert_eq!(changed.patch().after(), Some(&after));
    assert!(!changed.patch().is_noop());
    assert!(changed.diagnostics().changed());
    assert!(changed.diagnostics().touched_components() >= 1);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_eq!(changed.package().slide_transition(selector)?, Some(after));
    assert_eq!(package.exact_bytes(), source_snapshot);
    assert_ne!(changed.package().exact_bytes(), source_snapshot);

    let reapplied = package.apply_slide_transition(changed.patch())?;
    assert!(reapplied.diagnostics().changed());
    assert_eq!(
        reapplied.package().exact_bytes(),
        changed.package().exact_bytes()
    );

    let inverse = changed.patch().inverse();
    let restored = changed.package().apply_slide_transition(&inverse)?;
    assert!(restored.diagnostics().changed());
    assert_eq!(restored.package().slide_transition(selector)?, Some(before));
    assert_eq!(restored.package().exact_bytes(), source_snapshot);
    Ok(())
}

#[test]
fn slide_placeholder_visibility_transaction_reaches_the_root_facade()
-> Result<(), Box<dyn std::error::Error>> {
    use litchi::keynote::slide::placeholder::{
        Commit, Diagnostics, Edit, Error, Kind, LimitKind, Patch, State,
    };

    assert_send_sync::<Kind>();
    assert_send_sync::<State>();
    assert_send_sync::<Edit<'static>>();
    assert_send_sync::<Patch>();
    assert_send_sync::<Commit>();
    assert_send_sync::<Diagnostics>();
    assert_send_sync::<Error>();
    assert_send_sync::<LimitKind>();

    let package = Package::open(fixture_path())?;
    let source = package.exact_bytes();
    let selector = SlideSelector::index(0);
    let kind = Kind::Title;
    let before = package
        .slide_placeholder_visibility(selector, kind)?
        .ok_or_else(|| io::Error::other("native Keynote file has no title placeholder"))?;
    assert_eq!(before, State::Visible);

    let edit = package.edit_slide_placeholder_visibility(selector, kind)?;
    assert_eq!(edit.position(), Position::new(0));
    assert_eq!(edit.kind(), kind);
    assert_eq!(edit.state(), before);
    let edit_debug = format!("{edit:?}");
    assert!(!edit_debug.contains("identifier"));
    assert!(!edit_debug.contains(".iwa"));

    let noop = edit.set(before).commit()?;
    assert_eq!(noop.patch().position(), Position::new(0));
    assert_eq!(noop.patch().kind(), kind);
    assert_eq!(noop.patch().before(), before);
    assert_eq!(noop.patch().after(), before);
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert_eq!(noop.diagnostics().deleted_previews(), 0);
    assert!(!noop.diagnostics().full_reparse_performed());
    assert_eq!(noop.package().exact_bytes(), source);

    let noop_applied = package.apply_slide_placeholder_visibility(noop.patch())?;
    assert!(!noop_applied.diagnostics().changed());
    assert_eq!(noop_applied.package().exact_bytes(), source);

    let changed = package
        .edit_slide_placeholder_visibility(selector, kind)?
        .hide()
        .commit()?;
    assert_eq!(changed.patch().position(), Position::new(0));
    assert_eq!(changed.patch().kind(), kind);
    assert_eq!(changed.patch().before(), State::Visible);
    assert_eq!(changed.patch().after(), State::Hidden);
    assert!(!changed.patch().is_noop());
    assert!(changed.diagnostics().changed());
    assert!(changed.diagnostics().touched_components() >= 1);
    assert!(changed.diagnostics().deleted_previews() >= 1);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_eq!(
        changed
            .package()
            .slide_placeholder_visibility(selector, kind)?,
        Some(State::Hidden)
    );
    assert_eq!(package.exact_bytes(), source);
    assert_ne!(changed.package().exact_bytes(), source);
    let patch_debug = format!("{:?}", changed.patch());
    assert!(!patch_debug.contains("identifier"));
    assert!(!patch_debug.contains(".iwa"));

    let applied = package.apply_slide_placeholder_visibility(changed.patch())?;
    assert!(applied.diagnostics().changed());
    assert_eq!(
        applied
            .package()
            .slide_placeholder_visibility(selector, kind)?,
        Some(State::Hidden)
    );
    assert_eq!(
        applied.package().exact_bytes(),
        changed.package().exact_bytes()
    );

    let inverse = changed.patch().inverse();
    let restored = changed
        .package()
        .apply_slide_placeholder_visibility(&inverse)?;
    assert!(restored.diagnostics().changed());
    assert_eq!(
        restored
            .package()
            .slide_placeholder_visibility(selector, kind)?,
        Some(State::Visible)
    );
    assert_eq!(restored.package().exact_bytes(), source);
    Ok(())
}

#[test]
fn slide_notes_transaction_is_available_through_the_root_facade()
-> Result<(), Box<dyn std::error::Error>> {
    let package = Package::open(fixture_path())?;
    let selector = SlideSelector::index(0);
    let source_snapshot = package.exact_bytes();

    if let Some(notes) = package.slide_notes(selector)? {
        let mut edit = package.edit_slide_notes(selector)?;
        edit.set(&notes)?;
        let commit = edit.commit()?;

        assert!(commit.patch().is_noop());
        assert!(!commit.diagnostics().changed());
        assert_eq!(commit.package().exact_bytes(), source_snapshot);

        let reapplied = package.apply_slide_notes(commit.patch())?;
        assert!(reapplied.patch().is_noop());
        assert_eq!(reapplied.package().exact_bytes(), source_snapshot);
    } else {
        assert!(matches!(
            package.edit_slide_notes(selector),
            Err(SlideNotesError::NotesStorageNotFound)
        ));
        assert_eq!(package.exact_bytes(), source_snapshot);
    }
    Ok(())
}

#[test]
fn slide_text_transaction_is_available_through_the_root_facade()
-> Result<(), Box<dyn std::error::Error>> {
    let package = Package::open(fixture_path())?;
    let source_snapshot = package.exact_bytes();
    let source_bytes = package.exact_bytes();

    let title = package
        .slide_text(SlideSelector::index(0), Kind::Title)?
        .ok_or_else(|| io::Error::other("native Keynote file has no title placeholder"))?;
    let body = package
        .slide_text(SlideSelector::index(0), Kind::Body)?
        .ok_or_else(|| io::Error::other("native Keynote file has no body placeholder"))?;
    assert_eq!(title, "Litchi native Keynote fixture");
    assert_eq!(body, "Buffa lazy-view migration verification");
    assert!(!body.contains("2026-08-07"));
    assert_eq!(
        package.slide_title(SlideSelector::index(0))?,
        Some(title.clone())
    );
    assert_eq!(
        package.slide_body(SlideSelector::index(0))?,
        Some(body.clone())
    );

    let mut no_op = package.edit_slide_title(SlideSelector::index(0))?;
    no_op.set(&title)?;
    let no_op = no_op.commit()?;
    assert!(no_op.patch().is_noop());
    assert!(!no_op.diagnostics().changed());
    assert_eq!(no_op.package().exact_bytes(), source_snapshot);

    let replacement = "Litchi native Keynote fixture — root facade";
    let mut edit = package.edit_slide_text(SlideSelector::index(0), Kind::Title)?;
    edit.set(replacement)?;
    let changed = edit.commit()?;
    assert_eq!(changed.patch().role(), Kind::Title);
    assert_eq!(changed.patch().before(), title);
    assert_eq!(changed.patch().after(), replacement);
    assert!(changed.diagnostics().changed());
    assert_eq!(
        changed.package().slide_title(SlideSelector::index(0))?,
        Some(replacement.to_owned())
    );
    assert_eq!(
        changed.package().slide_body(SlideSelector::index(0))?,
        Some(body.clone())
    );

    let inverse = changed.patch().inverse();
    let restored = changed.package().apply_slide_text(&inverse)?;
    assert_eq!(restored.package().exact_bytes(), source_bytes);
    assert_eq!(
        restored.package().slide_title(SlideSelector::index(0))?,
        Some(title)
    );
    assert_eq!(
        restored.package().slide_body(SlideSelector::index(0))?,
        Some(body)
    );
    Ok(())
}

#[test]
fn slide_body_span_edit_uses_only_public_semantic_facade_types()
-> Result<(), Box<dyn std::error::Error>> {
    let package = Package::open(fixture_path())?;
    let source_snapshot = package.exact_bytes();
    let source_bytes = package.exact_bytes();
    let slides = package.slides()?;
    let slide = slides
        .first()
        .ok_or_else(|| io::Error::other("native Keynote file has no slide"))?;
    let selector = slide.position_selector();

    let title = package
        .slide_title(selector)?
        .ok_or_else(|| io::Error::other("native Keynote file has no title placeholder"))?;
    let body = package
        .slide_body(selector)?
        .ok_or_else(|| io::Error::other("native Keynote file has no body placeholder"))?;
    let span = TextSpan::from_utf16_indexes(6, 15)?;
    let mut edit = package.edit_slide_body(selector)?;
    assert_eq!(edit.position(), Position::new(slide.index()));
    assert_eq!(edit.role(), Kind::Body);
    assert_eq!(edit.text(), body);
    assert_eq!(edit.span(), None);
    edit.replace(span, "selector-first")?;
    assert_eq!(edit.span(), Some(span));
    let edit_debug = format!("{edit:?}");
    assert!(!edit_debug.contains("identifier"));
    assert!(!edit_debug.contains(".iwa"));

    let changed = edit.commit()?;
    assert_eq!(changed.patch().position(), Position::new(slide.index()));
    assert_eq!(changed.patch().role(), Kind::Body);
    assert_eq!(changed.patch().span(), span);
    assert_eq!(changed.patch().before(), body);
    assert_eq!(
        changed.patch().after(),
        "Buffa selector-first migration verification"
    );
    assert!(changed.diagnostics().changed());
    assert_eq!(changed.diagnostics().touched_components(), 2);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_eq!(package.exact_bytes(), source_snapshot);
    assert_eq!(package.exact_bytes(), source_bytes);
    assert_eq!(changed.package().slide_title(selector)?, Some(title));
    assert_eq!(
        changed.package().slide_body(selector)?,
        Some("Buffa selector-first migration verification".to_owned())
    );
    let patch_debug = format!("{:?}", changed.patch());
    assert!(!patch_debug.contains("identifier"));
    assert!(!patch_debug.contains(".iwa"));

    let restored = changed
        .package()
        .apply_slide_text(&changed.patch().inverse())?;
    assert_eq!(restored.package().exact_bytes(), source_bytes);
    assert_eq!(restored.package().slide_body(selector)?, Some(body));
    Ok(())
}

#[cfg(feature = "iwork")]
#[test]
fn focused_slide_text_edit_does_not_mutate_the_read_only_iwork_snapshot()
-> Result<(), Box<dyn std::error::Error>> {
    let package = Package::open(fixture_path())?;
    let document = litchi::iwork::Document::from_bytes(package.exact_bytes())?;
    let snapshot = document.snapshot();
    let original_slide = snapshot
        .slide(0)
        .ok_or_else(|| io::Error::other("root iWork snapshot has no slide"))?;
    let original_title = original_slide
        .title()
        .ok_or_else(|| io::Error::other("root iWork snapshot has no slide title"))?
        .to_owned();
    assert_eq!(original_title, "Litchi native Keynote fixture");

    let replacement = "Focused Keynote edit, immutable iWork view";
    let mut edit = package.edit_slide_title(SlideSelector::index(0))?;
    edit.set(replacement)?;
    let changed = edit.commit()?;

    let unchanged_slide = snapshot
        .slide(0)
        .ok_or_else(|| io::Error::other("retained iWork snapshot lost its slide"))?;
    assert_eq!(unchanged_slide.title(), Some(original_title.as_str()));
    let fresh_original_snapshot = document.snapshot();
    let fresh_original_slide = fresh_original_snapshot
        .slide(0)
        .ok_or_else(|| io::Error::other("fresh iWork snapshot lost its slide"))?;
    assert_eq!(fresh_original_slide.title(), Some(original_title.as_str()));

    let candidate = litchi::iwork::Document::from_bytes(changed.package().exact_bytes())?;
    let candidate_snapshot = candidate.snapshot();
    let candidate_slide = candidate_snapshot
        .slide(0)
        .ok_or_else(|| io::Error::other("edited iWork snapshot has no slide"))?;
    assert_eq!(candidate_slide.title(), Some(replacement));
    Ok(())
}
