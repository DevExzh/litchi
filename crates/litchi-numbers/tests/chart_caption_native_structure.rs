//! Native structure of a Numbers chart caption (record 0431).
//!
//! A native Numbers chart lives in the calculation engine's component, and its
//! caption's text storage names a paragraph style in the stylesheet's
//! component. Native files list every such cross-component reference in the
//! component's `external_references`; only a component's root object goes
//! unlisted. Numbers 14.4 refused the four packages that broke this rule.

use std::collections::{HashMap, HashSet};

use litchi_iwa_archive::Limits;
use litchi_iwa_archive::iwa::{Archive, SnappyStream};
use litchi_iwa_archive::package::Catalog;
use litchi_iwa_protos::tsp;
use litchi_numbers::Package;
use litchi_numbers::sheet::drawable::caption::Label;
use litchi_numbers::sheet::drawable::{DrawableSelector, Kind};
use prost::Message;

const NATIVE_CHART: &[u8] =
    include_bytes!("../../../test-data/iwork/numbers/chart-data-native.numbers");
const NATIVE_BASIC: &[u8] = include_bytes!("../../../test-data/iwork/numbers/basic.numbers");
const PACKAGE_METADATA: u32 = 11_006;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

/// `(owner, target)` of every reference to a non-root object of another
/// component that the owner's component does not list.
fn unlisted_references(bytes: &[u8]) -> TestResult<Vec<(u64, u64)>> {
    let catalog = Catalog::from_bytes(bytes, Limits::default())?;
    let mut archives = Vec::new();
    for entry in catalog.iter() {
        let name = entry.name().to_owned();
        if name.ends_with(".iwa") {
            let stream = SnappyStream::decompress(entry.data(), Default::default())?;
            archives.push((name, Archive::parse(stream.as_bytes(), Default::default())?));
        }
    }
    let metadata = archives
        .iter()
        .flat_map(|(_, archive)| &archive.objects)
        .flat_map(|object| &object.messages)
        .find(|message| message.type_ == PACKAGE_METADATA)
        .ok_or("no package metadata")?;
    let metadata = tsp::PackageMetadata::decode(metadata.data.as_slice())?;
    let mut components = HashMap::new();
    let mut roots = HashSet::new();
    for component in &metadata.components {
        roots.insert(component.identifier);
        let stem = component
            .locator
            .clone()
            .unwrap_or_else(|| component.preferred_locator.clone());
        components.insert(format!("Index/{stem}.iwa"), component);
        let identified = format!(
            "Index/{}-{}.iwa",
            component.preferred_locator, component.identifier
        );
        components.entry(identified).or_insert(component);
    }
    let mut homes = HashMap::new();
    for (name, archive) in &archives {
        if let Some(component) = components.get(name) {
            for object in &archive.objects {
                homes.insert(object.archive_info.identifier, component.identifier);
            }
        }
    }
    let mut unlisted = Vec::new();
    for (name, archive) in &archives {
        let Some(component) = components.get(name) else {
            continue;
        };
        let listed = component
            .external_references
            .iter()
            .filter_map(|reference| {
                reference
                    .object_identifier
                    .map(|object| (reference.component_identifier, object))
            })
            .collect::<HashSet<_>>();
        for object in &archive.objects {
            let references = object
                .archive_info
                .message_infos
                .iter()
                .flat_map(|info| info.object_references.iter().copied());
            for target in references {
                let Some(&home) = homes.get(&Some(target)) else {
                    continue;
                };
                if home != component.identifier
                    && !roots.contains(&target)
                    && !listed.contains(&(home, target))
                {
                    unlisted.push((object.archive_info.identifier.unwrap_or(0), target));
                }
            }
        }
    }
    Ok(unlisted)
}

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    Ok(output)
}

#[test]
fn native_files_list_every_cross_component_reference() -> TestResult {
    assert_eq!(unlisted_references(NATIVE_BASIC)?, Vec::new());
    assert_eq!(unlisted_references(NATIVE_CHART)?, Vec::new());
    Ok(())
}

#[test]
fn chart_caption_lists_its_paragraph_style_like_a_native_file() -> TestResult {
    let selector = DrawableSelector::of(Kind::Chart, 0);
    let package = Package::from_bytes(NATIVE_CHART)?;
    let mut edit = package.edit_sheet_drawable_caption(0, selector, Label::Caption)?;
    edit.set("Revenue report")?;
    let captioned = edit.commit()?;
    assert_eq!(
        unlisted_references(&bytes(captioned.package())?)?,
        Vec::new()
    );
    assert_eq!(
        captioned
            .package()
            .sheet_drawable_caption(0, selector, Label::Caption)?
            .as_deref(),
        Some("Revenue report")
    );

    // Replacing and clearing keep the (undo-retained) graph's references listed.
    let mut edit = captioned
        .package()
        .edit_sheet_drawable_caption(0, selector, Label::Caption)?;
    edit.set("Updated")?;
    let replaced = edit.commit()?;
    assert_eq!(
        unlisted_references(&bytes(replaced.package())?)?,
        Vec::new()
    );
    let mut edit = replaced
        .package()
        .edit_sheet_drawable_caption(0, selector, Label::Caption)?;
    edit.clear()?;
    let cleared = edit.commit()?;
    assert_eq!(unlisted_references(&bytes(cleared.package())?)?, Vec::new());
    Ok(())
}
