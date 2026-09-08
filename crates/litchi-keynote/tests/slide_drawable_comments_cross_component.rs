//! Cross-component direct-drawable comment acceptance and transaction tests.
//!
//! The source fixture is native Keynote data.  The test-only fixture builder
//! relocates selected archive objects into existing native components and
//! repairs PackageMetadata edges, so these tests exercise the same ownership
//! witnesses as a package produced by Keynote rather than a raw-ID shortcut.

#[path = "support/drawable_comment_fixtures.rs"]
mod fixtures;

use std::{env, fs, io, path::Path};

use litchi_iwa_archive::{
    Limits,
    iwa::{Archive, RawMessage, SnappyStream},
    package::Catalog,
};
use litchi_iwa_protos::{tsd, tsp};
use litchi_keynote::{DrawableSelector, Package, ReplySelector, SlideSelector};
use prost::Message as _;

use fixtures::{RelocatedFixture, TestResult};

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn export_candidate(bytes: &[u8], name: &str) -> TestResult<()> {
    let Some(directory) =
        env::var_os("LITCHI_KEYNOTE_DRAWABLE_COMMENTS_CROSS_COMPONENT_OUTPUT_DIR")
    else {
        return Ok(());
    };
    let directory = Path::new(&directory);
    fs::create_dir_all(directory)?;
    let path = directory.join(name);
    fs::write(&path, bytes)?;
    eprintln!(
        "exported Keynote cross-component candidate to {}",
        path.display()
    );
    Ok(())
}

fn metadata_payload(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == "Index/Metadata.iwa")
        .ok_or_else(|| io::Error::other("missing metadata component"))?;
    let stream = SnappyStream::decompress(entry.data())?.into_bytes();
    let archive = Archive::parse(&stream)?;
    archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .find(|message| message.type_ == 11_006)
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("missing PackageMetadata payload").into())
}

fn replace_metadata_payload(source: &[u8], metadata: &tsp::PackageMetadata) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == "Index/Metadata.iwa")
        .ok_or_else(|| io::Error::other("missing metadata component"))?;
    let stream = SnappyStream::decompress(entry.data())?.into_bytes();
    let mut archive = Archive::parse(&stream)?;
    let object = archive
        .objects
        .iter_mut()
        .find(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == 11_006)
        })
        .ok_or_else(|| io::Error::other("missing metadata root"))?;
    let message = object
        .messages
        .iter_mut()
        .find(|message| message.type_ == 11_006)
        .ok_or_else(|| io::Error::other("missing metadata message"))?;
    message.data = metadata.encode_to_vec();
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    let entries = catalog
        .iter()
        .map(|entry| {
            if entry.name() == "Index/Metadata.iwa" {
                (entry.name(), compressed.as_slice())
            } else {
                (entry.name(), entry.data())
            }
        })
        .collect::<Vec<_>>();
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
}

fn without_reply_payload_uuid(source: &[u8], reply_identifier: u64) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut replacement = None;
    for entry in catalog.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let stream = SnappyStream::decompress(entry.data())?.into_bytes();
        let mut archive = Archive::parse(&stream)?;
        let Some(object) = archive.object_mut(reply_identifier) else {
            continue;
        };
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == 3_056)
            .ok_or_else(|| io::Error::other("reply object has no comment payload"))?;
        let mut comment =
            tsd::CommentStorageArchive::decode(object.messages[message_index].data.as_slice())?;
        if comment.storage_uuid.is_none() {
            return Err(io::Error::other("reply payload UUID already absent").into());
        }
        comment.storage_uuid = None;
        object.replace_message(
            message_index,
            RawMessage {
                type_: 3_056,
                data: comment.encode_to_vec(),
            },
        )?;
        replacement = Some((
            entry.name().to_owned(),
            SnappyStream::compress(&archive.to_bytes()?)?,
        ));
        break;
    }
    let Some((component_name, compressed)) = replacement else {
        return Err(io::Error::other("reply component disappeared").into());
    };
    let entries = catalog
        .iter()
        .map(|entry| {
            if entry.name() == component_name {
                (entry.name(), compressed.as_slice())
            } else {
                (entry.name(), entry.data())
            }
        })
        .collect::<Vec<_>>();
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
}

fn physical_locator(name: &str) -> &str {
    name.strip_prefix("Index/")
        .unwrap_or(name)
        .strip_suffix(".iwa")
        .unwrap_or(name.strip_prefix("Index/").unwrap_or(name))
}

fn inject_reply_data_owner(
    source: &[u8],
    reply_identifier: u64,
    strip_uuid_registration: bool,
    strip_payload_uuid: bool,
) -> TestResult<Vec<u8>> {
    let source = if strip_payload_uuid {
        without_reply_payload_uuid(source, reply_identifier)?
    } else {
        source.to_vec()
    };
    let mut metadata = tsp::PackageMetadata::decode(metadata_payload(&source)?.as_slice())?;
    let registration_count = metadata
        .components
        .iter()
        .flat_map(|component| &component.object_uuid_map_entries)
        .filter(|entry| entry.identifier == reply_identifier)
        .count();
    if strip_uuid_registration {
        for component in &mut metadata.components {
            component
                .object_uuid_map_entries
                .retain(|entry| entry.identifier != reply_identifier);
        }
        if registration_count != 1 {
            return Err(io::Error::other("reply UUID registration is not unique").into());
        }
    } else if registration_count != 1 {
        return Err(io::Error::other("reply UUID registration is missing").into());
    }
    let component_index = metadata
        .components
        .iter()
        .position(|component| {
            let locator = component
                .locator
                .as_deref()
                .unwrap_or(&component.preferred_locator);
            physical_locator(locator) == physical_locator("Index/AnnotationAuthorStorage.iwa")
        })
        .ok_or_else(|| io::Error::other("author-storage metadata component disappeared"))?;
    let data_identifier = metadata.components[component_index]
        .data_references
        .iter()
        .find(|reference| {
            metadata
                .datas
                .iter()
                .any(|data| data.identifier == reference.data_identifier)
        })
        .map(|reference| reference.data_identifier)
        .or_else(|| metadata.datas.first().map(|data| data.identifier))
        .ok_or_else(|| io::Error::other("native fixture has no valid data record"))?;
    let component = &mut metadata.components[component_index];
    if let Some(reference) = component
        .data_references
        .iter_mut()
        .find(|reference| reference.data_identifier == data_identifier)
    {
        if reference
            .object_reference_list
            .iter()
            .any(|owner| owner.object_identifier == reply_identifier)
        {
            return Err(io::Error::other("reply data owner already present").into());
        }
        reference
            .object_reference_list
            .push(tsp::component_data_reference::ObjectReference {
                object_identifier: reply_identifier,
                count: 1,
            });
    } else {
        component.data_references.push(tsp::ComponentDataReference {
            data_identifier,
            object_reference_list: vec![tsp::component_data_reference::ObjectReference {
                object_identifier: reply_identifier,
                count: 1,
            }],
        });
    }
    replace_metadata_payload(&source, &metadata)
}

fn selected_comment(package: &Package, fixture: &RelocatedFixture) -> TestResult<String> {
    package
        .slide_drawable_comment(
            SlideSelector::index(0),
            DrawableSelector::index(fixture.target.position),
        )?
        .map(|comment| comment.text().to_owned())
        .ok_or_else(|| io::Error::other("candidate has no selected root comment").into())
}

fn assert_sibling_unchanged(
    before: &Package,
    after: &Package,
    fixture: &RelocatedFixture,
) -> TestResult<()> {
    let Some(sibling) = fixture.sibling.as_ref() else {
        return Ok(());
    };
    let slide = SlideSelector::index(0);
    let selector = DrawableSelector::index(sibling.position);
    assert_eq!(
        before.slide_drawable_comment(slide, selector)?,
        after.slide_drawable_comment(slide, selector)?,
        "an unselected sibling changed during a cross-component edit",
    );
    let before_replies = before.slide_drawable_comment_replies(slide, selector)?;
    let after_replies = after.slide_drawable_comment_replies(slide, selector)?;
    assert_eq!(
        before_replies.as_ref(),
        after_replies.as_ref(),
        "an unselected sibling reply thread changed during a cross-component edit",
    );
    Ok(())
}

fn assert_metadata_shape(fixture: &RelocatedFixture, moved: u64) -> TestResult<()> {
    fixtures::assert_metadata_relocated(&fixture.metadata_source, &fixture.bytes, moved)
}

fn assert_noop_and_inverse(fixture: &RelocatedFixture) -> TestResult<()> {
    let source = Package::from_bytes(&fixture.bytes)?;
    let slide = SlideSelector::index(0);
    let target = DrawableSelector::index(fixture.target.position);
    let text = selected_comment(&source, fixture)?;
    let noop = source
        .edit_slide_drawable_comment(slide, target)?
        .set(&text)?
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(exact_bytes(noop.package())?, fixture.bytes);

    let changed = source
        .edit_slide_drawable_comment(slide, target)?
        .set("foreign root transaction")?
        .commit()?;
    assert_eq!(
        selected_comment(changed.package(), fixture)?,
        "foreign root transaction"
    );
    let restored = changed
        .package()
        .apply_slide_drawable_comment(&changed.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, fixture.bytes);

    // A patch is source-bound: applying it to an unrelated candidate must
    // report a conflict and leave that candidate's bytes untouched.
    let other = source
        .edit_slide_drawable_comment(slide, target)?
        .set("different source")?
        .commit()?;
    let other_bytes = exact_bytes(other.package())?;
    assert!(matches!(
        other
            .package()
            .apply_slide_drawable_comment(changed.patch()),
        Err(litchi_keynote::SlideDrawableCommentError::PatchConflict)
    ));
    assert_eq!(exact_bytes(other.package())?, other_bytes);
    Ok(())
}

#[test]
fn foreign_root_supports_read_set_clear_noop_inverse_and_stale_patch() -> TestResult {
    let fixture = fixtures::root_foreign_fixture()?;
    export_candidate(&fixture.bytes, "root-foreign-source.key")?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let before = Package::from_bytes(&fixture.bytes)?;
    let original = selected_comment(&package, &fixture)?;
    assert_metadata_shape(
        &fixture,
        fixture.root_identifier.expect("root fixture root"),
    )?;

    let changed = package
        .edit_slide_drawable_comment(
            SlideSelector::index(0),
            DrawableSelector::index(fixture.target.position),
        )?
        .set("foreign root update")?
        .commit()?;
    assert_eq!(
        selected_comment(changed.package(), &fixture)?,
        "foreign root update"
    );
    assert_sibling_unchanged(&before, changed.package(), &fixture)?;
    let cleared = changed
        .package()
        .edit_slide_drawable_comment(
            SlideSelector::index(0),
            DrawableSelector::index(fixture.target.position),
        )?
        .clear()?
        .commit()?;
    assert!(
        cleared
            .package()
            .slide_drawable_comment(
                SlideSelector::index(0),
                DrawableSelector::index(fixture.target.position),
            )?
            .is_none()
    );
    let restored = cleared
        .package()
        .apply_slide_drawable_comment(&cleared.patch().inverse())?;
    assert_eq!(
        selected_comment(restored.package(), &fixture)?,
        "foreign root update"
    );
    let original_bytes = restored
        .package()
        .apply_slide_drawable_comment(&changed.patch().inverse())?;
    assert_eq!(
        selected_comment(original_bytes.package(), &fixture)?,
        original
    );
    assert_eq!(exact_bytes(original_bytes.package())?, fixture.bytes);
    assert_noop_and_inverse(&fixture)?;
    Ok(())
}

#[test]
fn foreign_reply_supports_add_set_remove_and_inverse() -> TestResult<()> {
    let fixture = fixtures::reply_foreign_fixture()?;
    export_candidate(&fixture.bytes, "reply-foreign-source.key")?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let sibling = Package::from_bytes(&fixture.bytes)?;
    let slide = SlideSelector::index(0);
    let target = DrawableSelector::index(fixture.target.position);
    let before = package.slide_drawable_comment_replies(slide, target)?;
    assert!(
        !before.is_empty(),
        "reply fixture must contain a foreign reply"
    );
    assert_metadata_shape(
        &fixture,
        fixture.reply_identifier.expect("reply fixture reply"),
    )?;

    // Exercise the reply that was registered in PackageMetadata before the
    // relocation.  The later add flow intentionally creates an unregistered
    // reply; using ordinal zero here keeps the registered path covered.
    let registered_set = Package::from_bytes(&fixture.bytes)?
        .edit_slide_drawable_comment(slide, target)?
        .set_reply(ReplySelector::index(0), "registered reply updated")?
        .commit()?;
    let after_registered_set = registered_set
        .package()
        .slide_drawable_comment_replies(slide, target)?;
    assert_eq!(after_registered_set.len(), before.len());
    assert_eq!(
        after_registered_set.first().map(|reply| reply.text()),
        Some("registered reply updated")
    );
    let restored_registered_set = registered_set
        .package()
        .apply_slide_drawable_comment(&registered_set.patch().inverse())?;
    assert_eq!(
        exact_bytes(restored_registered_set.package())?,
        fixture.bytes
    );

    let registered_remove = Package::from_bytes(&fixture.bytes)?
        .edit_slide_drawable_comment(slide, target)?
        .remove_reply(ReplySelector::index(0))?
        .commit()?;
    let after_registered_remove = registered_remove
        .package()
        .slide_drawable_comment_replies(slide, target)?;
    assert_eq!(after_registered_remove.len(), before.len() - 1);
    let restored_registered_remove = registered_remove
        .package()
        .apply_slide_drawable_comment(&registered_remove.patch().inverse())?;
    assert_eq!(
        exact_bytes(restored_registered_remove.package())?,
        fixture.bytes
    );

    let added = package
        .edit_slide_drawable_comment(slide, target)?
        .add_reply("reply added after relocation")?
        .commit()?;
    let after_add = added
        .package()
        .slide_drawable_comment_replies(slide, target)?;
    assert_eq!(after_add.len(), before.len() + 1);
    assert_eq!(
        after_add.last().map(|reply| reply.text()),
        Some("reply added after relocation")
    );
    assert_sibling_unchanged(&sibling, added.package(), &fixture)?;

    let selected = ReplySelector::index(after_add.len() - 1);
    let updated = added
        .package()
        .edit_slide_drawable_comment(slide, target)?
        .set_reply(selected, "reply updated after relocation")?
        .commit()?;
    assert_eq!(
        updated
            .package()
            .slide_drawable_comment_replies(slide, target)?
            .last()
            .map(|reply| reply.text()),
        Some("reply updated after relocation")
    );
    let removed = updated
        .package()
        .edit_slide_drawable_comment(slide, target)?
        .remove_reply(selected)?
        .commit()?;
    let removed_replies = removed
        .package()
        .slide_drawable_comment_replies(slide, target)?;
    assert_eq!(removed_replies.as_ref(), before.as_ref());
    let restored = removed
        .package()
        .apply_slide_drawable_comment(&removed.patch().inverse())?;
    assert_eq!(
        restored
            .package()
            .slide_drawable_comment_replies(slide, target)?
            .len(),
        before.len() + 1
    );
    Ok(())
}

#[test]
fn data_reference_owner_rejects_reply_removal_atomically() -> TestResult<()> {
    let fixture = fixtures::reply_foreign_fixture()?;
    let reply = fixture.reply_identifier.expect("reply fixture reply");
    let slide = SlideSelector::index(0);
    let target = DrawableSelector::index(fixture.target.position);

    // Cover the full cross-product: a valid current UUID map, an unregistered
    // payload UUID, and a payload with no UUID at all.  The data owner must
    // reject physical removal before either identity path can run.
    for (strip_registration, strip_payload_uuid) in
        [(false, false), (true, false), (false, true), (true, true)]
    {
        let hostile = inject_reply_data_owner(
            &fixture.bytes,
            reply,
            strip_registration,
            strip_payload_uuid,
        )?;
        let package = Package::from_bytes(&hostile)?;
        assert_eq!(
            package.slide_drawable_comment_replies(slide, target)?.len(),
            1,
            "hostile source must retain the selected reply before editing"
        );
        let before = exact_bytes(&package)?;
        let result = package
            .edit_slide_drawable_comment(slide, target)?
            .remove_reply(ReplySelector::index(0))
            .and_then(|edit| edit.commit());
        assert!(matches!(
            result,
            Err(litchi_keynote::SlideDrawableCommentError::InvalidSource)
        ));
        assert_eq!(exact_bytes(&package)?, before);
    }
    Ok(())
}

#[test]
fn shared_foreign_root_cow_keeps_unselected_drawable_unchanged() -> TestResult<()> {
    let fixture = fixtures::shared_foreign_root_fixture()?;
    export_candidate(&fixture.bytes, "shared-foreign-root-source.key")?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let sibling = Package::from_bytes(&fixture.bytes)?;
    let changed = package
        .edit_slide_drawable_comment(
            SlideSelector::index(0),
            DrawableSelector::index(fixture.target.position),
        )?
        .set("selected shared root only")?
        .commit()?;
    assert_eq!(
        selected_comment(changed.package(), &fixture)?,
        "selected shared root only"
    );
    assert_sibling_unchanged(&sibling, changed.package(), &fixture)?;
    assert_metadata_shape(&fixture, fixture.root_identifier.expect("shared root"))?;
    Ok(())
}

#[test]
fn shared_foreign_reply_clear_keeps_sibling_thread() -> TestResult<()> {
    let fixture = fixtures::shared_foreign_reply_fixture()?;
    export_candidate(&fixture.bytes, "shared-foreign-reply-source.key")?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let sibling = Package::from_bytes(&fixture.bytes)?;
    let slide = SlideSelector::index(0);
    let target = DrawableSelector::index(fixture.target.position);
    let sibling_selector = DrawableSelector::index(
        fixture
            .sibling
            .as_ref()
            .expect("shared reply sibling")
            .position,
    );
    let sibling_before = sibling.slide_drawable_comment_replies(slide, sibling_selector)?;
    assert!(!sibling_before.is_empty());
    let cleared = package
        .edit_slide_drawable_comment(slide, target)?
        .clear()?
        .commit()?;
    assert!(
        cleared
            .package()
            .slide_drawable_comment(slide, target)?
            .is_none()
    );
    let sibling_after = cleared
        .package()
        .slide_drawable_comment_replies(slide, sibling_selector)?;
    assert_eq!(sibling_after.as_ref(), sibling_before.as_ref());
    assert_metadata_shape(&fixture, fixture.reply_identifier.expect("shared reply"))?;
    Ok(())
}

#[test]
fn foreign_slide_owned_drawable_keeps_source_order_and_supports_root_creation() -> TestResult<()> {
    let fixture = fixtures::foreign_drawable_fixture()?;
    export_candidate(&fixture.bytes, "foreign-drawable-source.key")?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let sibling = Package::from_bytes(&fixture.bytes)?;
    let slide = SlideSelector::index(0);
    let target = DrawableSelector::index(fixture.target.position);
    assert!(package.slide_drawable_comment(slide, target)?.is_none());
    let created = package
        .edit_slide_drawable_comment(slide, target)?
        .set("foreign drawable root")?
        .commit()?;
    assert_eq!(
        selected_comment(created.package(), &fixture)?,
        "foreign drawable root"
    );
    assert_sibling_unchanged(&sibling, created.package(), &fixture)?;
    let restored = created
        .package()
        .apply_slide_drawable_comment(&created.patch().inverse())?;
    assert!(
        restored
            .package()
            .slide_drawable_comment(slide, target)?
            .is_none()
    );
    assert_metadata_shape(&fixture, fixture.target.identifier)?;
    Ok(())
}

/// Materialize every valid cross-component source before any CRUD assertions.
/// The parent task opens these files in Keynote and saves accepted candidates
/// through the native GUI; this focused test only proves that each candidate
/// has a coherent package archive, source-order inventory, and metadata shape.
#[test]
fn export_cross_component_native_sources() -> TestResult<()> {
    let candidates: [(&str, fn() -> TestResult<RelocatedFixture>); 5] = [
        ("root-foreign-source.key", fixtures::root_foreign_fixture),
        ("reply-foreign-source.key", fixtures::reply_foreign_fixture),
        (
            "shared-foreign-root-source.key",
            fixtures::shared_foreign_root_fixture,
        ),
        (
            "shared-foreign-reply-source.key",
            fixtures::shared_foreign_reply_fixture,
        ),
        (
            "foreign-drawable-source.key",
            fixtures::foreign_drawable_fixture,
        ),
    ];
    for (name, build) in candidates {
        let fixture = build()?;
        let package = Package::from_bytes(&fixture.bytes)?;
        let inventory = package.slide_drawables(SlideSelector::index(0))?;
        assert!(fixture.target.position < inventory.len());
        let selected = package.slide_drawable_comment(
            SlideSelector::index(0),
            DrawableSelector::index(fixture.target.position),
        )?;
        if fixture.root_identifier.is_some() {
            assert!(selected.is_some(), "{name} must retain its root comment");
        } else {
            assert!(selected.is_none(), "{name} starts without a root comment");
        }
        if fixture.reply_identifier.is_some() {
            assert!(
                !package
                    .slide_drawable_comment_replies(
                        SlideSelector::index(0),
                        DrawableSelector::index(fixture.target.position),
                    )?
                    .is_empty(),
                "{name} must retain its relocated reply"
            );
        }
        if let Some(identifier) = fixture.root_identifier.or(fixture.reply_identifier) {
            fixtures::assert_metadata_relocated(
                &fixture.metadata_source,
                &fixture.bytes,
                identifier,
            )?;
        } else {
            fixtures::assert_metadata_relocated(
                &fixture.metadata_source,
                &fixture.bytes,
                fixture.target.identifier,
            )?;
        }
        export_candidate(&fixture.bytes, name)?;
    }
    Ok(())
}
