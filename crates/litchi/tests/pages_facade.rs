#![cfg(feature = "pages")]

use std::io;
use std::path::PathBuf;

use litchi::pages::document_settings::{
    Commit, Diagnostics, Edit, Error, LimitKind, Patch, Settings,
};
use litchi::pages::{
    Package, PageLayoutCommit, PageLayoutDiagnostics, PageLayoutEdit, PageLayoutError,
    PageLayoutLimitKind, PageLayoutPatch, SectionSelector, SectionTextCommit,
    SectionTextDiagnostics, SectionTextEdit, SectionTextError, SectionTextLimitKind,
    SectionTextPatch, TextPosition, TextSpan,
    page_layout::{Layout, Orientation},
    section::{PageNumber, PageNumbering, Start},
};

fn assert_send_sync<T: Send + Sync>() {}

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/pages/basic.pages")
}

fn package_bytes(package: &Package) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

/// Pages body text that the native reader projects into a rich-text storage
/// reaches the facade reader as one paragraph.
///
/// M0a step S8c: Pages has no `litchi::File` variant, so the facade reader is
/// `litchi::iwork::Document` (the facade's `Document` read the same
/// storage); `litchi::open` refuses the file with its format and a
/// suggestion to use that reader.
#[test]
fn body_storage_reaches_facade_paragraphs() -> Result<(), Box<dyn std::error::Error>> {
    let path = fixture_path();
    let native = Package::open(&path)?;
    let section = native
        .sections()
        .first()
        .ok_or_else(|| io::Error::other("native Pages file has no body section"))?;

    // This fixture reproduces the facade bug: Pages projects native body text
    // into a rich-text storage rather than the legacy paragraph collection.
    assert!(section.heading().is_none());
    assert!(section.paragraphs().is_empty());
    assert_eq!(section.text_storages().len(), 1);
    let expected = section
        .text_storages()
        .first()
        .ok_or_else(|| io::Error::other("native Pages body has no text storage"))?
        .text();
    assert!(!expected.is_empty());

    #[cfg(feature = "iwork")]
    {
        use litchi::iwork::{SectionKind, TextRole};

        let snapshot = litchi::iwork::Document::open(&path)?.snapshot();
        // The document's text is exactly the storage's text, as one
        // paragraph-level item (the facade's `text()`, `paragraph_count()`,
        // `paragraphs()` and `elements()` all reported that one paragraph).
        let text = snapshot.iter_text().collect::<Vec<_>>();
        assert_eq!(text.len(), 1);
        let paragraph = text
            .first()
            .ok_or_else(|| io::Error::other("iWork reader returned no Pages text"))?;
        assert_eq!(paragraph.value(), expected);
        assert_eq!(paragraph.role(), TextRole::SectionAdditional);
        assert_eq!(snapshot.all_text(), [expected.to_owned()]);

        assert_eq!(snapshot.section_count(), 1);
        let body = snapshot
            .section(0)
            .ok_or_else(|| io::Error::other("iWork reader returned no Pages section"))?;
        assert_eq!(body.kind(), SectionKind::Body);
        assert_eq!(body.paragraphs().len(), 0);
        assert_eq!(body.additional_text().collect::<Vec<_>>(), [expected]);
    }

    // `litchi::open` reads the same body through the unified `File`
    // (record 0296): one paragraph per line of the section's text.
    let file = litchi::open(&path)?;
    assert_eq!(file.format(), Some(litchi::FileFormat::Pages));
    let first_line = expected.split('\n').next().unwrap_or_default();
    let paragraph = file.get("/body/p[1]")?;
    assert_eq!(paragraph.view().text()?.as_deref(), Some(first_line));

    Ok(())
}

#[test]
fn section_name_transaction_reaches_pages_facade() -> Result<(), Box<dyn std::error::Error>> {
    let package = Package::open(fixture_path())?;
    let source = package_bytes(&package)?;

    let mut noop = package.edit_section_name(SectionSelector::index(0))?;
    noop.set_name(Some("Blank"))?;
    let noop = noop.commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(package_bytes(noop.package())?, source);

    let mut edit = package.edit_section_name(SectionSelector::name("Blank"))?;
    edit.set_name(Some("Facade Section"))?;
    let commit = edit.commit()?;
    assert_eq!(
        commit.package().sections()[0].name(),
        Some("Facade Section")
    );
    let restored = commit
        .package()
        .apply_section_name(&commit.patch().inverse())?;
    assert_eq!(package_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn section_pagination_transaction_reaches_pages_facade() -> Result<(), Box<dyn std::error::Error>> {
    let package = Package::open(fixture_path())?;
    let source = package_bytes(&package)?;

    let mut noop = package.edit_section_pagination(SectionSelector::index(0))?;
    noop.set_start(Some(Start::NextPage))?;
    noop.set_page_numbering(Some(PageNumbering::ContinueFromPrevious))?;
    noop.set_starting_page_number(Some(PageNumber::new(1)?));
    let noop = noop.commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(package_bytes(noop.package())?, source);

    let mut edit = package.edit_section_pagination(SectionSelector::name("Blank"))?;
    edit.set_start(Some(Start::LeftPage))?;
    edit.set_page_numbering(Some(PageNumbering::Restart))?;
    edit.set_starting_page_number(Some(PageNumber::new(17)?));
    let commit = edit.commit()?;
    let changed = commit
        .package()
        .section_pagination(SectionSelector::index(0))?;
    assert_eq!(changed.start(), Some(Start::LeftPage));
    assert_eq!(changed.page_numbering(), Some(PageNumbering::Restart));
    assert_eq!(
        changed.starting_page_number().map(PageNumber::get),
        Some(17)
    );

    let restored = commit
        .package()
        .apply_section_pagination(&commit.patch().inverse())?;
    assert_eq!(package_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn section_settings_transaction_reaches_pages_facade() -> Result<(), Box<dyn std::error::Error>> {
    use litchi::pages::section::{
        Settings,
        settings::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path},
    };

    assert_send_sync::<Settings>();
    assert_send_sync::<Edit<'static>>();
    assert_send_sync::<Patch>();
    assert_send_sync::<Commit>();
    assert_send_sync::<Diagnostics>();
    assert_send_sync::<Error>();
    assert_send_sync::<LimitKind>();
    assert_send_sync::<Path>();

    let package = Package::open(fixture_path())?;
    let source = package_bytes(&package)?;
    let selector = SectionSelector::index(0);
    let before = package.section_settings(selector)?;

    let noop_edit = package.edit_section_settings(selector)?;
    assert_eq!(noop_edit.settings(), &before);
    assert!(noop_edit.path().position().is_some());
    let edit_debug = format!("{noop_edit:?}");
    assert!(!edit_debug.contains("Blank"));
    assert!(!edit_debug.contains("Index/"));
    assert!(!edit_debug.contains(".iwa"));
    assert!(!edit_debug.contains("identifier"));

    let noop = noop_edit.set(before.clone())?.commit()?;
    assert_eq!(noop.patch().before(), &before);
    assert_eq!(noop.patch().after(), &before);
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert_eq!(noop.diagnostics().deleted_previews(), 0);
    assert!(!noop.diagnostics().full_reparse_performed());
    assert_eq!(package_bytes(noop.package())?, source);

    let replayed_noop = package.apply_section_settings(noop.patch())?;
    assert!(replayed_noop.patch().is_noop());
    assert_eq!(package_bytes(replayed_noop.package())?, source);

    let mut after = before.clone();
    after.set_first_page_hides_header_footer(Some(
        !before.first_page_hides_header_footer().unwrap_or(false),
    ));
    assert_ne!(after, before);
    let changed = package
        .edit_section_settings(selector)?
        .set(after.clone())?
        .commit()?;
    assert_eq!(changed.patch().before(), &before);
    assert_eq!(changed.patch().after(), &after);
    assert!(!changed.patch().is_noop());
    assert!(changed.diagnostics().changed());
    assert!(changed.diagnostics().touched_components() >= 1);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_eq!(package_bytes(&package)?, source);
    assert_ne!(package_bytes(changed.package())?, source);
    assert_eq!(changed.package().section_settings(selector)?, after);

    let patch_debug = format!("{:?}", changed.patch());
    assert!(!patch_debug.contains("Blank"));
    assert!(!patch_debug.contains("Index/"));
    assert!(!patch_debug.contains(".iwa"));
    assert!(!patch_debug.contains("identifier"));
    assert!(!patch_debug.contains("fingerprint"));

    let applied = package.apply_section_settings(changed.patch())?;
    assert_eq!(
        package_bytes(applied.package())?,
        package_bytes(changed.package())?
    );
    assert_eq!(applied.package().section_settings(selector)?, after);

    let conflict = changed
        .package()
        .apply_section_settings(changed.patch())
        .expect_err("a section-settings patch must authorize its exact source");
    assert!(matches!(&conflict, Error::PatchConflict));
    assert_eq!(format!("{conflict:?}"), "PatchConflict");

    let inverse = changed.patch().inverse();
    let inverse_conflict = package
        .apply_section_settings(&inverse)
        .expect_err("an inverse section-settings patch must authorize its target");
    assert!(matches!(inverse_conflict, Error::PatchConflict));

    let restored = changed.package().apply_section_settings(&inverse)?;
    assert_eq!(package_bytes(restored.package())?, source);
    assert_eq!(restored.package().section_settings(selector)?, before);
    Ok(())
}

#[test]
fn section_text_transaction_reaches_pages_facade() -> Result<(), Box<dyn std::error::Error>> {
    assert_send_sync::<SectionTextCommit>();
    assert_send_sync::<SectionTextDiagnostics>();
    assert_send_sync::<SectionTextEdit<'static>>();
    assert_send_sync::<SectionTextError>();
    assert_send_sync::<SectionTextLimitKind>();
    assert_send_sync::<SectionTextPatch>();
    assert_send_sync::<TextPosition>();

    let package = Package::open(fixture_path())?;
    let selector = SectionSelector::index(0);
    let original = package.section_text(selector)?.to_owned();
    let source = package_bytes(&package)?;

    let mut noop = package.edit_section_text(selector)?;
    noop.set(&original)?;
    let noop = noop.commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(package_bytes(noop.package())?, source);

    let prefix = "Facade section: ";
    let mut edit = package.edit_section_text(SectionSelector::name("Blank"))?;
    edit.replace(TextSpan::from_utf16_indexes(0, 0)?, prefix)?;
    let commit = edit.commit()?;
    assert!(commit.diagnostics().changed());
    assert_eq!(
        commit.package().section_text(SectionSelector::index(0))?,
        format!("{prefix}{original}")
    );

    let inverse = commit.patch().inverse();
    let restored = commit.package().apply_section_text(&inverse)?;
    assert_eq!(package_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn page_layout_transaction_reaches_pages_facade() -> Result<(), Box<dyn std::error::Error>> {
    assert_send_sync::<Layout>();
    assert_send_sync::<PageLayoutCommit>();
    assert_send_sync::<PageLayoutDiagnostics>();
    assert_send_sync::<PageLayoutEdit<'static>>();
    assert_send_sync::<PageLayoutError>();
    assert_send_sync::<PageLayoutLimitKind>();
    assert_send_sync::<PageLayoutPatch>();

    let package = Package::open(fixture_path())?;
    let source = package_bytes(&package)?;
    let before = package.page_layout()?;
    assert!(before.page_width().is_some_and(|width| width > 0.0));
    assert!(before.page_height().is_some_and(|height| height > 0.0));

    let replacement_orientation = if before.orientation() == Some(Orientation::Landscape) {
        Orientation::Portrait
    } else {
        Orientation::Landscape
    };
    let mut after = before;
    after.set_orientation(Some(replacement_orientation))?;
    assert_ne!(after, before);

    let mut edit = package.edit_page_layout()?;
    assert_eq!(edit.layout(), before);
    edit.set_layout(after)?;
    let edit_debug = format!("{edit:?}");
    assert!(edit_debug.contains("PageLayoutEdit"));
    assert!(!edit_debug.contains("Index/"));
    assert!(!edit_debug.contains(".iwa"));
    assert!(!edit_debug.contains("identifier"));

    let changed = edit.commit()?;
    assert_eq!(changed.patch().before(), before);
    assert_eq!(changed.patch().after(), after);
    assert_eq!(changed.package().page_layout()?, after);
    assert!(changed.diagnostics().changed());
    assert!(changed.diagnostics().touched_components() >= 1);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_eq!(package_bytes(&package)?, source);
    assert_ne!(package_bytes(changed.package())?, source);

    let patch_debug = format!("{:?}", changed.patch());
    assert!(patch_debug.contains("PageLayoutPatch"));
    assert!(!patch_debug.contains("Index/"));
    assert!(!patch_debug.contains(".iwa"));
    assert!(!patch_debug.contains("identifier"));
    assert!(!patch_debug.contains("fingerprint"));

    let restored = changed
        .package()
        .apply_page_layout(&changed.patch().inverse())?;
    assert_eq!(package_bytes(restored.package())?, source);
    assert_eq!(restored.package().page_layout()?, before);
    Ok(())
}

#[test]
fn document_settings_transaction_reaches_pages_facade() -> Result<(), Box<dyn std::error::Error>> {
    assert_send_sync::<Settings>();
    assert_send_sync::<Edit<'static>>();
    assert_send_sync::<Patch>();
    assert_send_sync::<Commit>();
    assert_send_sync::<Diagnostics>();
    assert_send_sync::<Error>();
    assert_send_sync::<LimitKind>();

    let package = Package::open(fixture_path())?;
    let source = package_bytes(&package)?;
    let before = package.document_settings()?;

    let noop = package.edit_document_settings()?.set(before).commit()?;
    assert_eq!(noop.patch().before(), before);
    assert_eq!(noop.patch().after(), before);
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert!(!noop.diagnostics().full_reparse_performed());
    assert_eq!(package_bytes(noop.package())?, source);

    let mut options = before.options();
    options.set_automatic_hyphenation(Some(!options.uses_automatic_hyphenation()));
    let mut after = before;
    after.set_options(options);
    assert_ne!(after, before);

    let changed = package.edit_document_settings()?.set(after).commit()?;
    assert_eq!(changed.patch().before(), before);
    assert_eq!(changed.patch().after(), after);
    assert_eq!(changed.package().document_settings()?, after);
    assert!(changed.diagnostics().changed());
    assert!(changed.diagnostics().touched_components() >= 1);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_eq!(package_bytes(&package)?, source);
    assert_ne!(package_bytes(changed.package())?, source);

    let restored = changed
        .package()
        .apply_document_settings(&changed.patch().inverse())?;
    assert_eq!(package_bytes(restored.package())?, source);
    assert_eq!(restored.package().document_settings()?, before);
    Ok(())
}
