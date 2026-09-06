//! Native Keynote coverage for layout reassignment ownership and publication.

use std::{
    collections::{BTreeMap, BTreeSet},
    io,
};

use litchi_iwa::keynote::{KeynoteEditor, KeynoteSlideImageInfo};
use litchi_iwa_archive::{SourceCatalog, iwa};
use litchi_iwa_protos::{
    kn,
    package_metadata_codec::{
        ComponentDescriptor, ExternalReferenceDescriptor, PackageMetadataVisitor, RewriteError,
        RewriteOptions, inspect_package_metadata_with_visitor,
    },
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const SOURCE: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/keynote/image-adjustments-native.key"
);
const RESAVED_SOURCE: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/keynote/slide-layout-native-resaved.key"
);
const SLIDE_MEMBER: &str = "Index/Slide-2652150.iwa";
const SLIDE_COMPONENT_LOCATOR: &str = "Slide-2652150";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const SLIDE_ID: u64 = 2_652_150;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const PLACEHOLDER_MESSAGE_TYPE: u32 = 7;
const METADATA_MESSAGE_TYPE: u32 = 11_006;

const SOURCE_STYLE: u64 = 2_651_718;
const SOURCE_TEMPLATE: u64 = 2_651_716;
const SOURCE_TITLE_STYLE: u64 = 2_651_719;
const TARGET_STYLE: u64 = 2_651_956;
const TARGET_TEMPLATE: u64 = 2_651_954;
const TARGET_TITLE_STYLE: u64 = 2_651_724;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct SlideReferences {
    style: u64,
    template: u64,
    title_style: u64,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
struct Dependency {
    source: u64,
    target_component: u64,
    object: Option<u64>,
    weak: Option<bool>,
}

#[derive(Debug, Default)]
struct MetadataFacts {
    current_components: BTreeMap<String, u64>,
    dependencies: BTreeSet<Dependency>,
}

impl MetadataFacts {
    fn dependencies_for(&self, source: u64) -> BTreeSet<Dependency> {
        self.dependencies
            .iter()
            .filter(|dependency| dependency.source == source)
            .copied()
            .collect()
    }
}

impl PackageMetadataVisitor for MetadataFacts {
    fn visit_component(&mut self, component: ComponentDescriptor<'_>) -> Result<(), RewriteError> {
        if component.is_current() {
            self.current_components.insert(
                component.effective_locator().to_owned(),
                component.identifier(),
            );
        }
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: ExternalReferenceDescriptor<'_>,
    ) -> Result<(), RewriteError> {
        if reference.source().is_current() && !reference.is_versioned() {
            self.dependencies.insert(Dependency {
                source: reference.source().identifier(),
                target_component: reference.target_component_identifier(),
                object: reference.object_identifier(),
                weak: reference.is_weak(),
            });
        }
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq)]
struct ImageSnapshot {
    info: KeynoteSlideImageInfo,
    media: Vec<u8>,
    raw: RawImageArchive,
}

#[derive(Debug, Clone, PartialEq)]
struct RawImageArchive {
    archive_info: iwa::ArchiveInfo,
    messages: Vec<iwa::RawMessage>,
}

#[derive(Debug, Clone, PartialEq)]
struct SlideContentSnapshot {
    name: Option<String>,
    title: Option<String>,
    body: Option<String>,
    notes: Option<String>,
    images: Vec<ImageSnapshot>,
}

#[derive(Debug, Clone, PartialEq)]
struct SlideSnapshot {
    layout_name: Option<String>,
    content: SlideContentSnapshot,
}

fn object_message<'a>(
    catalog: &'a SourceCatalog,
    member: &str,
    object_id: u64,
    message_type: u32,
) -> TestResult<&'a [u8]> {
    let component = catalog.components().get(member).ok_or_else(|| {
        io::Error::other(format!("component {member} is missing from the package"))
    })?;
    let object = component
        .archive()
        .object(object_id)
        .ok_or_else(|| io::Error::other(format!("object {object_id} is missing from {member}")))?;
    object
        .messages
        .iter()
        .find(|message| message.type_ == message_type)
        .map(|message| message.data.as_slice())
        .ok_or_else(|| {
            io::Error::other(format!(
                "message type {message_type} is missing from {member} object {object_id}"
            ))
            .into()
        })
}

fn slide_references(bytes: &[u8]) -> TestResult<SlideReferences> {
    let catalog = SourceCatalog::from_bytes(bytes)?;
    let slide = kn::SlideArchive::decode(object_message(
        &catalog,
        SLIDE_MEMBER,
        SLIDE_ID,
        SLIDE_MESSAGE_TYPE,
    )?)?;
    let title_placeholder = slide
        .title_placeholder
        .as_ref()
        .ok_or_else(|| io::Error::other("native slide has no title placeholder"))?;
    let placeholder = kn::PlaceholderArchive::decode(object_message(
        &catalog,
        SLIDE_MEMBER,
        title_placeholder.identifier,
        PLACEHOLDER_MESSAGE_TYPE,
    )?)?;

    Ok(SlideReferences {
        style: slide.style.identifier,
        template: slide
            .template_slide
            .as_ref()
            .ok_or_else(|| io::Error::other("native slide has no template reference"))?
            .identifier,
        title_style: placeholder
            .super_
            .super_
            .style
            .ok_or_else(|| io::Error::other("native title placeholder has no style"))?
            .identifier,
    })
}

fn metadata_payload(bytes: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = SourceCatalog::from_bytes(bytes)?;
    let component = catalog.components().get(METADATA_MEMBER).ok_or_else(|| {
        io::Error::other(format!(
            "component {METADATA_MEMBER} is missing from the package"
        ))
    })?;
    let mut payload = None;
    for object in &component.archive().objects {
        for message in &object.messages {
            if message.type_ == METADATA_MESSAGE_TYPE {
                if payload.replace(message.data.clone()).is_some() {
                    return Err(io::Error::other(
                        "package contains multiple PackageMetadata payloads",
                    )
                    .into());
                }
            }
        }
    }
    payload.ok_or_else(|| io::Error::other("PackageMetadata payload is missing").into())
}

fn metadata_facts(bytes: &[u8]) -> TestResult<MetadataFacts> {
    let payload = metadata_payload(bytes)?;
    let source_limit = payload.len().max(1);
    let options = RewriteOptions::new(
        source_limit,
        source_limit,
        source_limit.saturating_mul(2).max(1),
        payload.len().saturating_mul(64).max(1),
        64,
        source_limit,
        source_limit,
        0,
    );
    let mut facts = MetadataFacts::default();
    inspect_package_metadata_with_visitor(&payload, options, &mut facts)
        .map_err(|error| io::Error::other(format!("PackageMetadata inspection failed: {error}")))?;
    Ok(facts)
}

fn snapshot(bytes: &[u8], editor: &KeynoteEditor) -> TestResult<SlideSnapshot> {
    let catalog = SourceCatalog::from_bytes(bytes)?;
    let slide = editor
        .slides()?
        .into_iter()
        .next()
        .ok_or_else(|| io::Error::other("native presentation has no slides"))?;
    let mut images = Vec::new();
    for image in editor.slide_images(0)? {
        let raw = catalog
            .components()
            .iter()
            .find_map(|component| component.archive().object(image.drawable_object_id))
            .map(|object| RawImageArchive {
                archive_info: object.archive_info.clone(),
                messages: object.messages.clone(),
            })
            .ok_or_else(|| {
                io::Error::other(format!(
                    "image object {} is missing from the package",
                    image.drawable_object_id
                ))
            })?;
        images.push(ImageSnapshot {
            media: editor.extract_media(image.image_data_identifier)?,
            info: image,
            raw,
        });
    }

    Ok(SlideSnapshot {
        layout_name: slide.layout.map(|layout| layout.name),
        content: SlideContentSnapshot {
            name: slide.name,
            title: slide.title,
            body: slide.body,
            notes: slide.notes,
            images,
        },
    })
}

fn assert_roots_present(
    bytes: &[u8],
    references: SlideReferences,
    snapshot: &SlideSnapshot,
) -> TestResult {
    let catalog = SourceCatalog::from_bytes(bytes)?;
    let object_ids = catalog
        .components()
        .iter()
        .flat_map(|component| component.archive().objects.iter())
        .filter_map(|object| object.archive_info.identifier)
        .collect::<BTreeSet<_>>();
    let roots = [
        SLIDE_ID,
        references.style,
        references.template,
        references.title_style,
    ];
    for root in roots.into_iter().chain(
        snapshot
            .content
            .images
            .iter()
            .map(|image| image.info.drawable_object_id),
    ) {
        if !object_ids.contains(&root) {
            return Err(io::Error::other(format!(
                "rooted native object {root} is missing from the package"
            ))
            .into());
        }
    }
    Ok(())
}

fn object_ids(edges: &BTreeSet<Dependency>) -> BTreeSet<u64> {
    edges.iter().filter_map(|edge| edge.object).collect()
}

fn current_component(facts: &MetadataFacts, locator: &str) -> TestResult<u64> {
    facts
        .current_components
        .get(locator)
        .copied()
        .ok_or_else(|| io::Error::other(format!("metadata has no current component {locator}")))
        .map_err(Into::into)
}

fn component_ids_without_object(edges: &BTreeSet<Dependency>) -> BTreeSet<u64> {
    edges
        .iter()
        .filter(|edge| edge.object.is_none())
        .map(|edge| edge.target_component)
        .collect()
}

fn first_image_media(editor: &KeynoteEditor) -> TestResult<Vec<u8>> {
    let image = editor
        .slide_images(0)?
        .into_iter()
        .next()
        .ok_or_else(|| io::Error::other("native presentation has no slide image"))?;
    Ok(editor.extract_media(image.image_data_identifier)?)
}

#[test]
fn native_layout_reassignment_updates_metadata_and_preserves_slide_content() -> TestResult {
    let source = std::fs::read(SOURCE)?;
    let source_editor = KeynoteEditor::from_bytes(&source)?;
    let source_snapshot = snapshot(&source, &source_editor)?;
    let source_references = slide_references(&source)?;
    assert_eq!(
        source_references,
        SlideReferences {
            style: SOURCE_STYLE,
            template: SOURCE_TEMPLATE,
            title_style: SOURCE_TITLE_STYLE,
        }
    );
    assert_roots_present(&source, source_references, &source_snapshot)?;

    let mut editor = KeynoteEditor::from_bytes(&source)?;
    let target_layout = editor
        .slide_layouts()?
        .into_iter()
        .find(|layout| layout.name == "Title Only")
        .ok_or_else(|| io::Error::other("native theme has no Title Only layout"))?;
    assert_ne!(
        source_snapshot.layout_name.as_deref(),
        Some(target_layout.name.as_str())
    );

    editor.set_slide_layout(0, target_layout.id)?;
    let candidate = editor.to_bytes()?;
    editor.set_slide_layout(0, target_layout.id)?;
    assert_eq!(
        editor.to_bytes()?,
        candidate,
        "reassigning the active layout must be an exact no-op"
    );

    let candidate_editor = KeynoteEditor::from_bytes(&candidate)?;
    let candidate_snapshot = snapshot(&candidate, &candidate_editor)?;
    assert_eq!(
        candidate_snapshot.layout_name.as_deref(),
        Some(target_layout.name.as_str())
    );
    assert_eq!(
        source_snapshot.content, candidate_snapshot.content,
        "layout reassignment changed slide text, notes, image metadata, media, or native image fields"
    );
    assert_eq!(
        candidate_snapshot,
        snapshot(&candidate, &KeynoteEditor::from_bytes(&candidate)?)?,
        "serialized candidate does not reopen to the same typed slide projection"
    );
    let candidate_references = slide_references(&candidate)?;
    assert_eq!(
        candidate_references,
        SlideReferences {
            style: TARGET_STYLE,
            template: TARGET_TEMPLATE,
            title_style: TARGET_TITLE_STYLE,
        }
    );
    assert_roots_present(&candidate, candidate_references, &candidate_snapshot)?;

    let source_metadata = metadata_facts(&source)?;
    let candidate_metadata = metadata_facts(&candidate)?;
    let source_component = current_component(&source_metadata, SLIDE_COMPONENT_LOCATOR)?;
    let candidate_component = current_component(&candidate_metadata, SLIDE_COMPONENT_LOCATOR)?;
    assert_eq!(source_component, candidate_component);

    let source_edges = source_metadata.dependencies_for(source_component);
    let candidate_edges = candidate_metadata.dependencies_for(candidate_component);
    let removed = source_edges
        .difference(&candidate_edges)
        .copied()
        .collect::<BTreeSet<_>>();
    let added = candidate_edges
        .difference(&source_edges)
        .copied()
        .collect::<BTreeSet<_>>();
    assert_eq!(
        object_ids(&removed),
        BTreeSet::from([SOURCE_STYLE, SOURCE_TITLE_STYLE]),
        "layout reassignment must remove every stale current-slide object dependency"
    );
    assert_eq!(
        object_ids(&added),
        BTreeSet::from([TARGET_STYLE]),
        "layout reassignment must publish every current-slide style/template dependency"
    );
    let source_template_locator = format!("TemplateSlide-{SOURCE_TEMPLATE}");
    let target_template_locator = format!("TemplateSlide-{TARGET_TEMPLATE}");
    let source_template_component = current_component(&source_metadata, &source_template_locator)?;
    let target_template_component =
        current_component(&candidate_metadata, &target_template_locator)?;
    assert_eq!(
        component_ids_without_object(&removed),
        BTreeSet::from([source_template_component]),
        "layout reassignment must remove the stale slide-template component dependency"
    );
    assert_eq!(
        component_ids_without_object(&added),
        BTreeSet::from([target_template_component]),
        "layout reassignment must publish the selected slide-template component dependency"
    );
    assert!(
        candidate_edges
            .iter()
            .any(|edge| edge.object == Some(TARGET_TITLE_STYLE)),
        "the selected title placeholder style must remain registered"
    );
    Ok(())
}

#[test]
fn native_resaved_layout_fixture_preserves_typed_slide_and_media() -> TestResult {
    const TITLE: &str = "Saved layout Native image adjustment marker";

    let bytes = std::fs::read(RESAVED_SOURCE)?;
    let mut editor = KeynoteEditor::from_bytes(&bytes)?;
    let slide = editor
        .slides()?
        .into_iter()
        .next()
        .ok_or_else(|| io::Error::other("native resaved presentation has no slides"))?;
    let layout = slide
        .layout
        .as_ref()
        .ok_or_else(|| io::Error::other("native resaved slide has no layout"))?;
    assert_eq!(layout.name, "Title Only");
    assert_eq!(slide.is_title_visible, Some(true));
    assert_eq!(slide.is_body_visible, Some(false));
    assert_eq!(slide.title.as_deref(), Some(TITLE));

    let source_editor = KeynoteEditor::from_bytes(&std::fs::read(SOURCE)?)?;
    assert_eq!(
        first_image_media(&editor)?,
        first_image_media(&source_editor)?
    );

    let before = editor.to_bytes()?;
    editor.set_slide_layout(0, layout.id)?;
    assert_eq!(
        editor.to_bytes()?,
        before,
        "setting the active native layout must be an exact no-op"
    );

    let reopened = KeynoteEditor::from_bytes(&before)?;
    let reopened_slide = reopened
        .slides()?
        .into_iter()
        .next()
        .ok_or_else(|| io::Error::other("reopened native presentation has no slides"))?;
    assert_eq!(
        reopened_slide
            .layout
            .as_ref()
            .map(|layout| layout.name.as_str()),
        Some("Title Only")
    );
    assert_eq!(reopened_slide.is_title_visible, Some(true));
    assert_eq!(reopened_slide.is_body_visible, Some(false));
    assert_eq!(reopened_slide.title.as_deref(), Some(TITLE));
    assert_eq!(
        first_image_media(&reopened)?,
        first_image_media(&source_editor)?
    );
    Ok(())
}
