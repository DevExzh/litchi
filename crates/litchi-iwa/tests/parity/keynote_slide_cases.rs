//! Keynote slide-lifecycle parity cases (wave W1-C, record 0372): slide
//! duplication and creation from layouts through `slide::creation`, and the
//! layout catalog reads.

use litchi_iwa::keynote;
use litchi_keynote::slide::creation::LayoutSelector;
use litchi_keynote::{Package, Position, SlideSelector};

use litchi_iwa_protos::{kn, tsp};
use prost::Message as _;

use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
use crate::known_fields::{Decoded, decoded};
use crate::leaf_bytes;

/// Documented divergence: the host writes `slide_specific_hyperlink_count:
/// 0` on a created slide node, the leaf keeps the source node's value (the
/// native default, absent). Equal semantics; compared as equal.
#[allow(deprecated, reason = "the host writes the deprecated native field")]
fn slide_node_hyperlink_count_default(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut node = kn::SlideNodeArchive::decode(data)?;
    if node.slide_specific_hyperlink_count == Some(0) {
        node.slide_specific_hyperlink_count = None;
    }
    Ok(decoded(node, data))
}

thread_local! {
    /// Component identifiers of the case's input package.
    static INPUT_COMPONENTS: std::cell::RefCell<std::collections::BTreeSet<u64>> =
        std::cell::RefCell::default();
}

/// Documented divergences of the package metadata of a created slide
/// component:
///
/// - the leaf registers it as native Keynote does (preferred locator
///   `Slide`, locator `Slide-<id>`; record 0298 §2.1) where the host writes
///   preferred locator `Slide-<id>` and no locator;
/// - the leaf derives the new component's external references from the
///   references its objects hold (header and payload walk, which also finds
///   the storages' stylesheet reference), where the host copies the
///   layout's list (two unreferenced styles and a weak link included).
///   The external references of components absent from the input are
///   therefore not compared here; [`cross_references_registered`] checks
///   instead that each output registers every cross-component reference
///   its new components make.
fn metadata_native_slide_components(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut metadata = tsp::PackageMetadata::decode(data)?;
    INPUT_COMPONENTS.with(|input| {
        let input = input.borrow();
        for component in &mut metadata.components {
            if component.preferred_locator == "Slide"
                && let Some(locator) = component.locator.take()
            {
                component.preferred_locator = locator;
            }
            if !input.contains(&component.identifier) {
                component.external_references.clear();
            }
        }
    });
    Ok(decoded(metadata, data))
}

fn component_of_member(metadata: &tsp::PackageMetadata, member: &str) -> Option<u64> {
    metadata
        .components
        .iter()
        .find(|component| {
            let locator = component
                .locator
                .as_deref()
                .unwrap_or(&component.preferred_locator);
            format!("Index/{locator}.iwa") == member
        })
        .map(|component| component.identifier)
}

fn package_metadata(bytes: &[u8]) -> TestResult<tsp::PackageMetadata> {
    for member in crate::archive_diff::members(bytes)? {
        if member.name != "Index/Metadata.iwa" {
            continue;
        }
        for object in crate::archive_diff::objects(&member.data)? {
            for message in &object.messages {
                if message.type_ == 11_006 {
                    return Ok(tsp::PackageMetadata::decode(message.data.as_slice())?);
                }
            }
        }
    }
    Err("no package metadata".into())
}

/// Read: every object reference a component absent from the input makes
/// (archive headers) into another component is registered in that
/// component's external references; returns the checked components.
fn cross_references_registered(bytes: &[u8]) -> TestResult<Vec<String>> {
    let metadata = package_metadata(bytes)?;
    let input = INPUT_COMPONENTS.with(|input| input.borrow().clone());
    let members = crate::archive_diff::members(bytes)?;
    let mut located = std::collections::BTreeMap::new();
    let mut objects = std::collections::BTreeMap::new();
    for member in &members {
        if member.name.ends_with(".iwa") {
            let parsed = crate::archive_diff::objects(&member.data)?;
            for object in &parsed {
                located.insert(object.identifier, member.name.clone());
            }
            objects.insert(member.name.clone(), parsed);
        }
    }
    let mut checked = Vec::new();
    for component in metadata
        .components
        .iter()
        .filter(|c| !input.contains(&c.identifier))
    {
        let locator = component
            .locator
            .clone()
            .unwrap_or_else(|| component.preferred_locator.clone());
        let member = format!("Index/{locator}.iwa");
        let Some(parsed) = objects.get(&member) else {
            continue;
        };
        for object in parsed {
            for info in &object.info.message_infos {
                for target in info.object_references.iter().chain(
                    info.field_infos
                        .iter()
                        .flat_map(|f| f.object_references.iter()),
                ) {
                    let Some(target_member) = located.get(target) else {
                        continue;
                    };
                    if *target_member == member {
                        continue;
                    }
                    // An object reference, or a component link (no object)
                    // to the target's component.
                    let target_component = component_of_member(&metadata, target_member);
                    if !component.external_references.iter().any(|reference| {
                        reference.object_identifier == Some(*target)
                            || (reference.object_identifier.is_none()
                                && Some(reference.component_identifier) == target_component)
                    }) {
                        return Err(format!(
                            "{member}: reference to {target} in {target_member} is not registered"
                        )
                        .into());
                    }
                }
            }
        }
        checked.push(member);
    }
    Ok(checked)
}

const BASIC: &str = "iwork/keynote/basic.key";
const MULTI: &str = "iwork/keynote/multi-slide.key";

fn leaf_run(
    source: &[u8],
    stage: impl FnOnce(&mut litchi_keynote::slide::creation::Edit<'_>) -> TestResult,
) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_slide_creation();
    stage(&mut edit)?;
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_slide_creation(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

/// Leaf read: slide names, titles and layouts in show order.
fn leaf_slides(bytes: &[u8]) -> TestResult<Vec<(Option<String>, Option<String>)>> {
    let package = Package::from_bytes(bytes)?;
    let count = package.slides()?.len();
    let mut slides = Vec::with_capacity(count);
    for position in 0..count {
        let selector = SlideSelector::index(position);
        slides.push((
            package.slide_title(selector)?,
            package.slide_body(selector)?,
        ));
    }
    Ok(slides)
}

/// Host read: slide count, layout names and skip state in show order.
fn host_slides(bytes: &[u8]) -> TestResult<Vec<(Option<String>, bool)>> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    Ok(editor
        .slides()?
        .into_iter()
        .map(|slide| (slide.layout.map(|layout| layout.name), slide.is_skipped))
        .collect())
}

fn host_layout(editor: &keynote::Draft, name: &str) -> TestResult<keynote::SlideLayoutId> {
    Ok(editor
        .slide_layouts()?
        .into_iter()
        .find(|layout| layout.name == name)
        .ok_or_else(|| format!("host has no layout {name}"))?
        .id)
}

fn run(
    name: &'static str,
    path: &'static str,
    tier: Tier,
    evidence: Evidence,
    host: impl FnOnce(&mut keynote::Draft) -> TestResult,
    leaf: impl FnOnce(&[u8]) -> TestResult<LeafRun>,
) -> TestResult {
    let input = fixture(path)?;
    let components = package_metadata(&input.bytes)?
        .components
        .iter()
        .map(|component| component.identifier)
        .collect();
    INPUT_COMPONENTS.with(|input| *input.borrow_mut() = components);
    Case::new(name, Format::Keynote, tier, evidence, &input)
        .host(|source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            host(&mut editor)?;
            Ok(editor.to_bytes()?)
        })
        .leaf(leaf)
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .known_fields(|known| {
            known
                .with(4, "KN.SlideNodeArchive", slide_node_hyperlink_count_default)
                .with(
                    11_006,
                    "TSP.PackageMetadata",
                    metadata_native_slide_components,
                )
        })
        .read("leaf slide text", leaf_slides)
        .read("host slides", host_slides)
        .read(
            "cross-component references registered",
            cross_references_registered,
        )
        .run()?;
    Ok(())
}

#[test]
fn keynote_slide_duplicate_parity() -> TestResult {
    run(
        "keynote_slide_duplicate_parity",
        BASIC,
        Tier::B,
        Evidence::Queued,
        |editor| {
            editor.duplicate_slide(0)?;
            Ok(())
        },
        |source| {
            leaf_run(source, |edit| {
                edit.duplicate_slide(SlideSelector::index(0), Position::new(1))?;
                Ok(())
            })
        },
    )
}

#[test]
fn keynote_slide_add_parity() -> TestResult {
    run(
        "keynote_slide_add_parity",
        BASIC,
        Tier::B,
        Evidence::Queued,
        |editor| {
            let layout = editor.default_slide_layout()?;
            editor.add_slide(layout)?;
            Ok(())
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let layout = package
                .default_slide_layout()?
                .ok_or("leaf has no default layout")?;
            let count = package.slides()?.len();
            let name = layout
                .name()
                .ok_or("default layout has no name")?
                .to_owned();
            leaf_run(source, |edit| {
                edit.add_slide(LayoutSelector::Name(&name), Position::new(count))?;
                Ok(())
            })
        },
    )
}

#[test]
fn keynote_slide_insert_parity() -> TestResult {
    run(
        "keynote_slide_insert_parity",
        MULTI,
        Tier::B,
        Evidence::Queued,
        |editor| {
            let layout = host_layout(editor, "Title & Bullets")?;
            editor.insert_slide(1, layout)?;
            Ok(())
        },
        |source| {
            leaf_run(source, |edit| {
                edit.add_slide(LayoutSelector::Name("Title & Bullets"), Position::new(1))?;
                Ok(())
            })
        },
    )
}

/// Host and leaf layout catalogs agree (names in theme order and the
/// default), or an error naming the difference.
fn layouts_agree(bytes: &[u8]) -> TestResult<Vec<String>> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let host: Vec<String> = editor
        .slide_layouts()?
        .into_iter()
        .map(|layout| layout.name)
        .collect();
    let default = editor.default_slide_layout()?;
    let host_default = editor
        .slide_layouts()?
        .into_iter()
        .find(|layout| layout.id == default)
        .map(|layout| layout.name);
    let package = Package::from_bytes(bytes)?;
    let leaf: Vec<String> = package
        .slide_layouts()?
        .iter()
        .map(|layout| layout.name().unwrap_or_default().to_owned())
        .collect();
    let leaf_default = package
        .default_slide_layout()?
        .and_then(|layout| layout.name().map(str::to_owned));
    if host != leaf || host_default != leaf_default {
        return Err(format!(
            "layouts: host {host:?} (default {host_default:?}), leaf {leaf:?} (default {leaf_default:?})"
        )
        .into());
    }
    Ok(leaf)
}

#[test]
fn keynote_slide_layouts_read_parity() -> TestResult {
    for path in [BASIC, MULTI] {
        let input = fixture(path)?;
        Case::new(
            "keynote_slide_layouts_read_parity",
            Format::Keynote,
            Tier::A,
            Evidence::Queued,
            &input,
        )
        .read_only()
        .host(|source| Ok(keynote::Draft::from_bytes(source)?.to_bytes()?))
        .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
        .read("host and leaf layouts agree", layouts_agree)
        .run()?;
    }
    Ok(())
}
