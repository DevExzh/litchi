//! Parts a removal leaves unreachable.
//!
//! Removing a picture, a chart or a header must not leave its content in the
//! saved file. A main relationship the main document no longer references
//! is removed with the parts only it reaches, and a removed part takes the
//! parts only it reaches with it. A part another relationship still reaches,
//! or one this edit's earlier graph changes touch, stays.

use std::collections::{BTreeSet, HashSet};
use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI};

use super::super::transaction::GraphChange;
use super::super::{Error, Result};
use super::Edit;
use crate::Package;
use crate::document::{TransferPart, TransferRelationship};

fn xml_error() -> Error {
    Error::Document(crate::Error::Xml("invalid Word XML".into()))
}

impl Edit<Package> {
    /// The identifiers of the main part's relationships that an attribute of
    /// the projected main document names.
    pub(super) fn referenced_main_relationships(&mut self) -> Result<BTreeSet<String>> {
        let ids: BTreeSet<String> = self
            .package
            .opc_package()
            .main_document_part()
            .map_err(crate::Error::from)?
            .rels()
            .iter()
            .map(|relationship| relationship.r_id().to_owned())
            .collect();
        let mut found = BTreeSet::new();
        if ids.is_empty() {
            return Ok(found);
        }
        let mut reader = quick_xml::Reader::from_reader(self.inner_mut()?.main_xml());
        loop {
            match reader.read_event().map_err(|_| xml_error())? {
                quick_xml::events::Event::Start(element)
                | quick_xml::events::Event::Empty(element) => {
                    use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
                    for attribute in element.unchecked_attributes() {
                        let attribute = attribute.map_err(|_| xml_error())?;
                        if let Ok(value) = std::str::from_utf8(attribute.value.as_ref())
                            && ids.contains(value)
                        {
                            found.insert(value.to_owned());
                        }
                    }
                },
                quick_xml::events::Event::Eof => return Ok(found),
                _ => {},
            }
        }
    }

    /// Stage the removal of the main relationships in `before` that the
    /// projected main document no longer references, with the parts only
    /// they reach.
    pub(super) fn drop_released_relationships(&mut self, before: &BTreeSet<String>) -> Result<()> {
        let after = self.referenced_main_relationships()?;
        let released: BTreeSet<String> = before.difference(&after).cloned().collect();
        if released.is_empty() {
            return Ok(());
        }
        let main = self
            .package
            .opc_package()
            .main_document_part()
            .map_err(crate::Error::from)?;
        let mut relationships = Vec::new();
        let mut roots = Vec::new();
        for relationship in main.rels().iter() {
            if !released.contains(relationship.r_id()) {
                continue;
            }
            relationships.push(TransferRelationship {
                id: relationship.r_id().to_owned(),
                relationship_type: relationship.reltype().to_owned(),
                target: relationship.target_ref().to_owned(),
                external: relationship.is_external(),
            });
            if !relationship.is_external() {
                roots.push(relationship.target_partname().map_err(crate::Error::from)?);
            }
        }
        // A relationship this edit staged earlier also changes the graph;
        // leave such a relationship and its parts alone.
        let staged = self.staged_names()?;
        relationships.retain(|relationship| !staged.relationships.contains(&relationship.id));
        if relationships.is_empty() {
            return Ok(());
        }
        let parts = self.unreachable_parts(&roots, &released, false)?;
        self.inner_mut()?.stage_graph_change(
            Arc::new(GraphChange {
                relationships: relationships.into(),
                parts: parts.into(),
                replaced: Arc::new([]),
            }),
            false,
        )?;
        Ok(())
    }

    /// The parts to remove with `roots` once the main relationships named
    /// by `dropped` are gone: every part a root reaches that nothing else
    /// still reaches, roots first. With `forced`, the roots themselves are
    /// always removed.
    ///
    /// Reachability is taken after every removal this edit has already
    /// staged: a part an earlier removal of the same edit takes, and a main
    /// relationship it drops, no longer keep anything. A part or relationship
    /// target an earlier staged addition or replacement names stays.
    pub(super) fn unreachable_parts(
        &self,
        roots: &[PackURI],
        dropped: &BTreeSet<String>,
        forced: bool,
    ) -> Result<Vec<TransferPart>> {
        let package = self.package.opc_package();
        let main_name = package
            .main_document_part()
            .map_err(crate::Error::from)?
            .partname()
            .clone();
        let staged = self.staged_names()?;
        let mut dropped = dropped.clone();
        dropped.extend(staged.dropped_relationships.iter().cloned());
        let forced_roots: HashSet<&str> = if forced {
            roots.iter().map(PackURI::as_str).collect()
        } else {
            HashSet::new()
        };
        let gone = |name: &PackURI| {
            forced_roots.contains(name.as_str()) || staged.removed.contains(name.as_str())
        };
        let internal_targets = |relationships: &litchi_opc::Relationships, skip_dropped: bool| {
            relationships
                .iter()
                .filter(|relationship| !relationship.is_external())
                .filter(|relationship| !(skip_dropped && dropped.contains(relationship.r_id())))
                .filter_map(|relationship| relationship.target_partname().ok())
                .collect::<Vec<_>>()
        };
        // One pass: everything still reachable from the package root, or
        // named by a staged addition or replacement.
        let mut reachable: HashSet<PackURI> = HashSet::new();
        let mut pending = internal_targets(package.rels(), false);
        for name in &staged.kept {
            if let Ok(name) = PackURI::new(name.clone()) {
                pending.push(name);
            }
        }
        while let Some(name) = pending.pop() {
            if gone(&name) || !reachable.insert(name.clone()) {
                continue;
            }
            if let Ok(part) = package.get_part(&name) {
                pending.extend(internal_targets(part.rels(), name == main_name));
            }
        }
        // Everything a root reaches that is not reachable otherwise.
        let mut removed: HashSet<PackURI> = HashSet::new();
        let mut ordered = Vec::new();
        let mut pending: Vec<PackURI> = roots.iter().rev().cloned().collect();
        while let Some(name) = pending.pop() {
            let take = forced_roots.contains(name.as_str())
                || (!reachable.contains(&name) && !staged.removed.contains(name.as_str()));
            if !take || removed.contains(&name) || package.get_part(&name).is_err() {
                continue;
            }
            removed.insert(name.clone());
            ordered.push(name.clone());
            if let Ok(part) = package.get_part(&name) {
                let mut targets = internal_targets(part.rels(), false);
                targets.sort_by(|left, right| right.as_str().cmp(left.as_str()));
                pending.extend(targets);
            }
        }
        ordered
            .iter()
            .map(|name| transfer_part(package, name))
            .collect()
    }

    /// What this edit's staged graph changes do: the parts and main
    /// relationships they remove, and the parts and relationship targets
    /// they add or replace.
    fn staged_names(&self) -> Result<Staged> {
        let base = self
            .package
            .opc_package()
            .main_document_part()
            .map_err(crate::Error::from)?
            .partname()
            .base_uri()
            .to_owned();
        self.inner_ref(|inner| {
            let mut staged = Staged::default();
            for (change, forward) in inner.staged_graph_changes() {
                for part in change.parts.iter() {
                    if forward {
                        staged.kept.insert(part.name.clone());
                    } else {
                        staged.removed.insert(part.name.clone());
                    }
                }
                for part in change.replaced.iter() {
                    staged.kept.insert(part.name.clone());
                }
                for relationship in change.relationships.iter() {
                    staged.relationships.insert(relationship.id.clone());
                    if !forward {
                        staged.dropped_relationships.insert(relationship.id.clone());
                    } else if !relationship.external
                        && let Ok(target) = PackURI::from_rel_ref(&base, &relationship.target)
                    {
                        staged.kept.insert(target.as_str().to_owned());
                    }
                }
            }
            staged
        })
    }
}

#[derive(Default)]
struct Staged {
    /// Parts a staged change removes.
    removed: BTreeSet<String>,
    /// Parts a staged change adds or replaces, and targets of main
    /// relationships it adds.
    kept: BTreeSet<String>,
    /// Main relationship identifiers a staged change adds or removes.
    relationships: BTreeSet<String>,
    /// Main relationship identifiers a staged change removes.
    dropped_relationships: BTreeSet<String>,
}

/// The exact content of the part `name`, as a removal records it.
fn transfer_part(package: &OpcPackage, name: &PackURI) -> Result<TransferPart> {
    let part = package.get_part(name).map_err(crate::Error::from)?;
    Ok(TransferPart {
        name: name.as_str().to_owned(),
        content_type: part.content_type().to_owned(),
        blob: part.blob_arc(),
        relationships: part
            .rels()
            .iter()
            .map(|relationship| TransferRelationship {
                id: relationship.r_id().to_owned(),
                relationship_type: relationship.reltype().to_owned(),
                target: relationship.target_ref().to_owned(),
                external: relationship.is_external(),
            })
            .collect(),
    })
}
