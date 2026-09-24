//! Source-bound drawing-layer declaration edits.
//!
//! Layer declarations are intentionally handled as a small XML splice plan.
//! The package is never rendered and no shape tree is regenerated: a rename
//! changes the declaration and all references proven to resolve to that owner,
//! while an add or remove changes only the checked declaration span.

use crate::core::OwnedPackage;
use crate::model::layer::{
    self, DRAW_NAMESPACE, Layer, LayerInventory, LayerLocation, LayerOwner, LayerSetLocation,
    ParentLocation, ParsedLayers,
};
use litchi_core::{Error, Result, xml::escape_xml};
use litchi_odf_common::core::{
    AuthoredXmlFragment, XmlSourcePart, XmlSplicePublication, rebuild_package_with_xml_splices,
};
use std::collections::BTreeMap;
use std::ops::Range;

const MAX_PACKAGE_BYTES: usize = 128 * 1024 * 1024;
const MAX_EDITS: usize = 65_536;

/// A transaction-local declaration operation.
#[derive(Clone, Debug)]
pub(crate) enum Operation {
    Add {
        owner: LayerOwner,
        layer: Layer,
    },
    Rename {
        owner: LayerOwner,
        from: String,
        to: String,
    },
    Remove {
        owner: LayerOwner,
        name: String,
        replacement: Option<String>,
    },
}

#[derive(Clone)]
struct PartState {
    path: &'static str,
    xml: String,
    parsed: ParsedLayers,
}

#[derive(Clone)]
enum Fragment {
    Markup(Vec<u8>),
    Text(Vec<u8>),
    Delete,
}

/// Apply one checked declaration operation to an immutable package.
pub(crate) fn apply(source: &OwnedPackage, operation: &Operation) -> Result<Vec<u8>> {
    let parts = load_parts(source)?;
    let mut edits = BTreeMap::<&'static str, Vec<(Range<usize>, Fragment)>>::new();
    match operation {
        Operation::Add { owner, layer } => {
            layer::validate_name(layer.name(), "ODP layer name")?;
            let part = part_for_owner(&parts, owner)?;
            let existing = find_set(&part.parsed, owner)?;
            if let Some(set) = existing {
                if set
                    .layers
                    .iter()
                    .any(|candidate| candidate.layer.name() == layer.name())
                {
                    return invalid("ODP layer already exists in the selected layer set");
                }
                if set.empty {
                    let replacement = expand_empty_set(part, set, layer)?;
                    push_edit(
                        &mut edits,
                        part.path,
                        set.start..set.end,
                        Fragment::Markup(replacement),
                    )?;
                } else {
                    let end = set
                        .end_tag_start
                        .ok_or_else(|| invalid_error("ODP layer-set closing span is missing"))?;
                    let prefix = generated_prefix(set);
                    let markup = layer.to_xml(&prefix, true)?.into_bytes();
                    push_edit(&mut edits, part.path, end..end, Fragment::Markup(markup))?;
                }
            } else {
                let parent = find_parent(&part.parsed, owner)?.ok_or_else(|| {
                    invalid_error("selected ODP layer owner has no parent element")
                })?;
                let replacement = add_missing_set(part, parent, layer)?;
                let range = if parent.empty {
                    parent.start..parent.end
                } else {
                    parent.layer_insert_at..parent.layer_insert_at
                };
                push_edit(&mut edits, part.path, range, Fragment::Markup(replacement))?;
            }
        },
        Operation::Rename { owner, from, to } => {
            layer::validate_name(from, "ODP source layer name")?;
            layer::validate_name(to, "ODP destination layer name")?;
            if from == to {
                return Ok(source.as_bytes().to_vec());
            }
            let selected = select_layer(&parts, owner, from)?;
            if selected
                .set
                .layers
                .iter()
                .any(|candidate| candidate.layer.name() == to)
            {
                return invalid("ODP destination layer name already exists");
            }
            let escaped = escape_xml(to).into_bytes();
            push_edit(
                &mut edits,
                selected.path,
                selected.location.name_value.0..selected.location.name_value.1,
                Fragment::Text(escaped.clone()),
            )?;
            for part in &parts {
                for reference in &part.parsed.references {
                    if reference.owner == *owner && reference.value == *from {
                        push_edit(
                            &mut edits,
                            part.path,
                            reference.range.0..reference.range.1,
                            Fragment::Text(escaped.clone()),
                        )?;
                    }
                }
            }
        },
        Operation::Remove {
            owner,
            name,
            replacement,
        } => {
            layer::validate_name(name, "ODP layer name")?;
            let selected = select_layer(&parts, owner, name)?;
            let references = parts
                .iter()
                .flat_map(|part| {
                    part.parsed
                        .references
                        .iter()
                        .filter(|reference| reference.owner == *owner && reference.value == *name)
                })
                .collect::<Vec<_>>();
            if !references.is_empty() {
                let Some(replacement) = replacement else {
                    return unsupported(
                        "cannot remove an ODP layer while draw:layer references remain",
                    );
                };
                layer::validate_name(replacement, "ODP replacement layer name")?;
                if replacement == name
                    || !selected
                        .set
                        .layers
                        .iter()
                        .any(|candidate| candidate.layer.name() == replacement)
                {
                    return invalid("ODP layer replacement is not declared in the selected set");
                }
                let escaped = escape_xml(replacement).into_bytes();
                for reference in references {
                    let path = parts
                        .iter()
                        .find(|part| {
                            part.parsed
                                .references
                                .iter()
                                .any(|candidate| std::ptr::eq(candidate, reference))
                        })
                        .map(|part| part.path)
                        .ok_or_else(|| {
                            invalid_error("ODP layer reference source part disappeared")
                        })?;
                    push_edit(
                        &mut edits,
                        path,
                        reference.range.0..reference.range.1,
                        Fragment::Text(escaped.clone()),
                    )?;
                }
            }
            push_edit(
                &mut edits,
                selected.path,
                selected.location.start..selected.location.end,
                Fragment::Delete,
            )?;
        },
    }
    publish(source, edits)
}

/// Materialize the layer inventory from a package after any prior operation.
pub(crate) fn inventory(source: &OwnedPackage) -> Result<LayerInventory> {
    let parts = load_parts(source)?;
    let content = &parts
        .iter()
        .find(|part| part.path == "content.xml")
        .ok_or_else(|| invalid_error("ODP content.xml is missing"))?
        .parsed;
    let styles = parts
        .iter()
        .find(|part| part.path == "styles.xml")
        .map(|part| &part.parsed);
    let sets = content
        .sets
        .iter()
        .chain(styles.into_iter().flat_map(|parsed| parsed.sets.iter()))
        .map(|set| layer::LayerSet {
            owner: set.owner.clone(),
            layers: set.layers.iter().map(|layer| layer.layer.clone()).collect(),
        })
        .collect::<Vec<_>>();
    Ok(LayerInventory { sets })
}

struct SelectedLayer<'a> {
    path: &'static str,
    set: &'a LayerSetLocation,
    location: &'a LayerLocation,
}

/// Layer operations splice only these two existing XML parts. Comparing their
/// complete bytes also checks unknown markup and references, so a net-zero
/// sequence can discard its rebuilt container and retain the original package.
pub(crate) fn same_parts(left: &OwnedPackage, right: &OwnedPackage) -> Result<bool> {
    for path in ["content.xml", "styles.xml"] {
        let present = left.has_file(path)?;
        if present != right.has_file(path)? {
            return Ok(false);
        }
        if present && left.get_file(path)? != right.get_file(path)? {
            return Ok(false);
        }
    }
    Ok(true)
}

fn load_parts(source: &OwnedPackage) -> Result<Vec<PartState>> {
    let mut parts = Vec::with_capacity(2);
    let content = read_xml(source, "content.xml")?;
    parts.push(PartState {
        path: "content.xml",
        parsed: layer::scan(&content, "content.xml")?,
        xml: content,
    });
    if source.has_file("styles.xml")? {
        let styles = read_xml(source, "styles.xml")?;
        parts.push(PartState {
            path: "styles.xml",
            parsed: layer::scan(&styles, "styles.xml")?,
            xml: styles,
        });
    }
    Ok(parts)
}

fn read_xml(source: &OwnedPackage, path: &str) -> Result<String> {
    String::from_utf8(source.get_file(path)?)
        .map_err(|error| invalid_error(format!("ODP {path} is not UTF-8: {error}")))
}

fn part_for_owner<'a>(parts: &'a [PartState], owner: &LayerOwner) -> Result<&'a PartState> {
    match owner {
        LayerOwner::Page(_) => parts
            .iter()
            .find(|part| part.path == "content.xml")
            .ok_or_else(|| invalid_error("ODP content.xml is missing")),
        LayerOwner::MasterStyles | LayerOwner::MasterPage(_) => parts
            .iter()
            .find(|part| part.path == "styles.xml")
            .ok_or_else(|| invalid_error("selected ODP style owner requires styles.xml")),
    }
}

fn find_set<'a>(
    parsed: &'a ParsedLayers,
    owner: &LayerOwner,
) -> Result<Option<&'a LayerSetLocation>> {
    let mut matches = parsed.sets.iter().filter(|set| set.owner == *owner);
    let first = matches.next();
    if matches.next().is_some() {
        return invalid("selected ODP layer owner has duplicate layer sets");
    }
    Ok(first)
}

fn find_parent<'a>(
    parsed: &'a ParsedLayers,
    owner: &LayerOwner,
) -> Result<Option<&'a ParentLocation>> {
    let mut matches = parsed
        .parents
        .iter()
        .filter(|(candidate, _)| candidate == owner)
        .map(|(_, parent)| parent);
    let first = matches.next();
    if matches.next().is_some() {
        return invalid("selected ODP layer owner has duplicate parent elements");
    }
    Ok(first)
}

fn select_layer<'a>(
    parts: &'a [PartState],
    owner: &LayerOwner,
    name: &str,
) -> Result<SelectedLayer<'a>> {
    let part = part_for_owner(parts, owner)?;
    let set = find_set(&part.parsed, owner)?
        .ok_or_else(|| invalid_error("selected ODP layer owner has no layer-set"))?;
    let mut matches = set.layers.iter().filter(|layer| layer.layer.name() == name);
    let location = matches
        .next()
        .ok_or_else(|| invalid_error("selected ODP layer does not exist"))?;
    if matches.next().is_some() {
        return invalid("selected ODP layer name is ambiguous");
    }
    Ok(SelectedLayer {
        path: part.path,
        set,
        location,
    })
}

fn generated_prefix(set: &LayerSetLocation) -> String {
    if !set.name_prefix.is_empty() {
        set.name_prefix.clone()
    } else if let Some(prefix) = set
        .draw_prefix
        .as_deref()
        .filter(|prefix| !prefix.is_empty())
    {
        prefix.to_owned()
    } else {
        "draw".to_owned()
    }
}

fn expand_empty_set(
    part: &PartState,
    set: &LayerSetLocation,
    layer_value: &Layer,
) -> Result<Vec<u8>> {
    let source = part
        .xml
        .as_bytes()
        .get(set.start..set.end)
        .ok_or_else(|| invalid_error("ODP layer-set source span is invalid"))?;
    let opening_end = source
        .len()
        .checked_sub(2)
        .filter(|end| source.get(*end..).is_some_and(|tail| tail == b"/>"))
        .ok_or_else(|| invalid_error("ODP empty layer-set source tag is invalid"))?;
    let prefix = generated_prefix(set);
    let mut output = Vec::with_capacity(source.len() + 160);
    output.extend_from_slice(&source[..opening_end]);
    output.push(b'>');
    output.extend_from_slice(
        layer_value
            .to_xml(&prefix, set.name_prefix.is_empty())?
            .as_bytes(),
    );
    output.extend_from_slice(b"</");
    output.extend_from_slice(
        &source[1..source
            .iter()
            .position(|byte| *byte == b' ' || *byte == b'>' || *byte == b'/')
            .ok_or_else(|| invalid_error("ODP layer-set element name is missing"))?],
    );
    output.push(b'>');
    Ok(output)
}

fn add_missing_set(
    part: &PartState,
    parent: &ParentLocation,
    layer_value: &Layer,
) -> Result<Vec<u8>> {
    let prefix = parent
        .draw_prefix
        .as_deref()
        .filter(|prefix| !prefix.is_empty())
        .unwrap_or("draw");
    let layer_set_name = qualified(prefix, "layer-set");
    let layer = layer_value.to_xml(prefix, false)?;
    let set = format!(
        "<{layer_set_name} xmlns:{prefix}=\"{DRAW_NAMESPACE}\">{layer}</{layer_set_name}>",
    );
    if !parent.empty {
        return Ok(set.into_bytes());
    }
    let source = part
        .xml
        .as_bytes()
        .get(parent.start..parent.end)
        .ok_or_else(|| invalid_error("ODP owner source span is invalid"))?;
    let opening_end = source
        .len()
        .checked_sub(2)
        .filter(|end| source.get(*end..).is_some_and(|tail| tail == b"/>"))
        .ok_or_else(|| invalid_error("ODP empty owner source tag is invalid"))?;
    let mut output = Vec::with_capacity(source.len() + set.len() + parent.qualified_name.len() + 4);
    output.extend_from_slice(&source[..opening_end]);
    output.push(b'>');
    output.extend_from_slice(set.as_bytes());
    output.extend_from_slice(b"</");
    output.extend_from_slice(&parent.qualified_name);
    output.push(b'>');
    Ok(output)
}

fn qualified(prefix: &str, local: &str) -> String {
    if prefix.is_empty() {
        local.to_owned()
    } else {
        format!("{prefix}:{local}")
    }
}

fn push_edit(
    edits: &mut BTreeMap<&'static str, Vec<(Range<usize>, Fragment)>>,
    path: &'static str,
    range: Range<usize>,
    fragment: Fragment,
) -> Result<()> {
    if range.start > range.end {
        return invalid("ODP layer edit range is reversed");
    }
    let list = edits.entry(path).or_default();
    if list.len() >= MAX_EDITS {
        return invalid("ODP layer transaction exceeds the edit limit");
    }
    if list.iter().any(|(other, _)| {
        range.start < other.end && other.start < range.end
            || (range.start == range.end && other.start == other.end && range.start == other.start)
    }) {
        return invalid("ODP layer transaction contains overlapping edits");
    }
    list.push((range, fragment));
    Ok(())
}

fn publish(
    source: &OwnedPackage,
    edits: BTreeMap<&'static str, Vec<(Range<usize>, Fragment)>>,
) -> Result<Vec<u8>> {
    if edits.is_empty() {
        return Ok(source.as_bytes().to_vec());
    }
    let mut publications = Vec::with_capacity(edits.len());
    for (path, mut changes) in edits {
        let part = XmlSourcePart::load(source, path)?;
        changes.sort_by_key(|(range, _)| range.start);
        let mut publication = XmlSplicePublication::new(part.clone());
        for (range, fragment) in changes {
            let expected = part
                .bytes()
                .get(range.clone())
                .ok_or_else(|| invalid_error("ODP layer edit range is out of bounds"))?;
            let proof = part.checked_range(range, expected)?;
            let authored = match fragment {
                Fragment::Markup(bytes) => AuthoredXmlFragment::markup(bytes)?,
                Fragment::Text(bytes) => AuthoredXmlFragment::text(bytes)?,
                Fragment::Delete => AuthoredXmlFragment::deletion(),
            };
            publication.replace(proof, authored)?;
        }
        publications.push(publication);
    }
    rebuild_package_with_xml_splices(source, publications, MAX_PACKAGE_BYTES)
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn unsupported<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Unsupported(message.into()))
}
