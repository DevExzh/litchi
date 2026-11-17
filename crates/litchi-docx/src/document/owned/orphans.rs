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
        let mut ids: BTreeSet<String> = self
            .package
            .opc_package()
            .main_document_part()
            .map_err(crate::Error::from)?
            .rels()
            .iter()
            .map(|relationship| relationship.r_id().to_owned())
            .collect();
        // Relationships this edit added are the main part's too.
        ids.extend(
            self.staged_relationships(None)?
                .into_iter()
                .map(|relationship| relationship.id),
        );
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
        // A relationship of the package this edit staged a change of
        // earlier is left alone with its parts.
        let staged = self.staged_names()?;
        relationships.retain(|relationship| !staged.relationships.contains(&relationship.id));
        let mut parts = if relationships.is_empty() {
            Vec::new()
        } else {
            self.unreachable_parts(&roots, &released, false)?
        };
        // A relationship this edit added and no longer needs goes with the
        // part it added, so nothing it staged survives.
        let added: Vec<TransferRelationship> = self
            .staged_relationships(None)?
            .into_iter()
            .filter(|relationship| released.contains(&relationship.id))
            .collect();
        let dropped: Vec<(Option<String>, String)> = added
            .iter()
            .map(|relationship| (None, relationship.id.clone()))
            .collect();
        parts.extend(self.released_staged_parts(None, &added, &dropped)?);
        relationships.extend(added);
        if relationships.is_empty() {
            return Ok(());
        }
        self.inner_mut()?.stage_graph_change(
            Arc::new(GraphChange {
                relationships: relationships.into(),
                parts: parts.into(),
                replaced: Arc::new([]),
                part_relationships: Arc::new([]),
            }),
            false,
        )?;
        Ok(())
    }

    /// The relationships of `part` (the main part for `None`) this edit's
    /// staged changes add and have not removed again, in staging order.
    pub(super) fn staged_relationships(
        &self,
        part: Option<&str>,
    ) -> Result<Vec<TransferRelationship>> {
        self.inner_ref(|inner| {
            let mut present: Vec<TransferRelationship> = Vec::new();
            for (change, forward) in inner.staged_graph_changes() {
                let relationships: Vec<&TransferRelationship> = match part {
                    None => change.relationships.iter().collect(),
                    Some(part) => change
                        .part_relationships
                        .iter()
                        .filter(|owned| owned.part == part)
                        .map(|owned| &owned.relationship)
                        .collect(),
                };
                for relationship in relationships {
                    present.retain(|existing| existing.id != relationship.id);
                    if forward {
                        present.push(relationship.clone());
                    }
                }
            }
            present
        })
    }

    /// The parts this edit's staged changes add and have not removed again,
    /// as they were added: name, content type and relationships.
    fn staged_parts(&self) -> Result<Vec<TransferPart>> {
        self.inner_ref(|inner| {
            let mut present: Vec<TransferPart> = Vec::new();
            for (change, forward) in inner.staged_graph_changes() {
                for part in change.parts.iter() {
                    present.retain(|existing| existing.name != part.name);
                    if forward {
                        present.push(part.clone());
                    }
                }
            }
            present
        })
    }

    /// The parts this edit added that the relationships `released` (of
    /// `owner`, the main part for `None`) reach and nothing else names once
    /// the relationships `dropped` (a part, `None` for the main part, and an
    /// identifier) are gone, as a removal records them.
    pub(super) fn released_staged_parts(
        &self,
        owner: Option<&str>,
        released: &[TransferRelationship],
        dropped: &[(Option<String>, String)],
    ) -> Result<Vec<TransferPart>> {
        let package = self.package.opc_package();
        let main = package.main_document_part().map_err(crate::Error::from)?;
        let main_name = main.partname().as_str().to_owned();
        let base_of = |part: Option<&str>| -> String {
            let name = part.unwrap_or(&main_name);
            name[..=name.rfind('/').unwrap_or(0)].to_owned()
        };
        let resolve = |part: Option<&str>, target: &str| -> Option<String> {
            PackURI::from_rel_ref(&base_of(part), target)
                .ok()
                .map(|name| name.as_str().to_owned())
        };
        let staged_parts = self.staged_parts()?;
        // Every target something still names: the package's relationships,
        // and the relationships this edit added that stay.
        let mut named: BTreeSet<String> = BTreeSet::new();
        for part in package.iter_parts() {
            let owner_name = part.partname().as_str().to_owned();
            for relationship in part.rels().iter() {
                let key = (
                    (owner_name != main_name).then(|| owner_name.clone()),
                    relationship.r_id().to_owned(),
                );
                if relationship.is_external() || dropped.contains(&key) {
                    continue;
                }
                if let Ok(target) = relationship.target_partname() {
                    named.insert(target.as_str().to_owned());
                }
            }
        }
        for relationship in package.rels().iter() {
            if let Ok(target) = relationship.target_partname() {
                named.insert(target.as_str().to_owned());
            }
        }
        let mut owners: Vec<Option<String>> = vec![None];
        owners.extend(
            self.inner_ref(|inner| {
                inner
                    .staged_graph_changes()
                    .flat_map(|(change, _)| {
                        change
                            .part_relationships
                            .iter()
                            .map(|owned| owned.part.clone())
                            .collect::<Vec<_>>()
                    })
                    .collect::<BTreeSet<_>>()
            })?
            .into_iter()
            .map(Some),
        );
        for part in &owners {
            for relationship in self.staged_relationships(part.as_deref())? {
                if relationship.external
                    || dropped.contains(&(part.clone(), relationship.id.clone()))
                {
                    continue;
                }
                if let Some(target) = resolve(part.as_deref(), &relationship.target) {
                    named.insert(target);
                }
            }
        }
        // The internal targets of an added part's own relationships.
        let targets_of = |part: &TransferPart| -> Vec<String> {
            part.relationships
                .iter()
                .filter(|relationship| !relationship.external)
                .filter_map(|relationship| resolve(Some(&part.name), &relationship.target))
                .collect()
        };
        // Whether an added part that is not being removed names `target`.
        let kept_names = |target: &str, removing: &[String]| {
            staged_parts
                .iter()
                .filter(|part| !removing.contains(&part.name))
                .any(|part| targets_of(part).iter().any(|named| named == target))
        };
        // The roots, then every added part only removed parts reach, until
        // nothing more goes: an added chart takes its embedded workbook and
        // its style and colour parts with it.
        let mut removing: Vec<String> = Vec::new();
        let mut pending: Vec<String> = released
            .iter()
            .filter(|relationship| !relationship.external)
            .filter_map(|relationship| resolve(owner, &relationship.target))
            .collect();
        while let Some(target) = pending.first().cloned() {
            pending.remove(0);
            if removing.contains(&target)
                || named.contains(&target)
                || !staged_parts.iter().any(|part| part.name == target)
            {
                continue;
            }
            let mut trial = removing.clone();
            trial.push(target.clone());
            if kept_names(&target, &trial) {
                continue;
            }
            removing.push(target.clone());
            if let Some(part) = staged_parts.iter().find(|part| part.name == target) {
                pending.extend(targets_of(part));
            }
        }
        let mut removed = Vec::with_capacity(removing.len());
        for name in removing {
            let Some(part) = staged_parts.iter().find(|part| part.name == name) else {
                continue;
            };
            let Some(blob) = self.part_payload(&name)? else {
                continue;
            };
            removed.push(TransferPart {
                name,
                content_type: part.content_type.clone(),
                blob,
                relationships: Arc::clone(&part.relationships),
            });
        }
        Ok(removed)
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
        self.unreachable_parts_without(roots, dropped, &BTreeSet::new(), forced)
    }

    /// As [`Self::unreachable_parts`], once the relationships `owned` names
    /// (a part and a relationship identifier of it) are gone too.
    pub(super) fn unreachable_parts_without(
        &self,
        roots: &[PackURI],
        dropped: &BTreeSet<String>,
        owned: &BTreeSet<(String, String)>,
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
        let mut owned = owned.clone();
        owned.extend(staged.dropped_part_relationships.iter().cloned());
        let forced_roots: HashSet<&str> = if forced {
            roots.iter().map(PackURI::as_str).collect()
        } else {
            HashSet::new()
        };
        let gone = |name: &PackURI| {
            forced_roots.contains(name.as_str()) || staged.removed.contains(name.as_str())
        };
        let internal_targets = |relationships: &litchi_opc::Relationships,
                                skip_dropped: bool,
                                part: Option<&PackURI>| {
            relationships
                .iter()
                .filter(|relationship| !relationship.is_external())
                .filter(|relationship| !(skip_dropped && dropped.contains(relationship.r_id())))
                .filter(|relationship| {
                    part.is_none_or(|part| {
                        !owned.contains(&(part.as_str().to_owned(), relationship.r_id().to_owned()))
                    })
                })
                .filter_map(|relationship| relationship.target_partname().ok())
                .collect::<Vec<_>>()
        };
        // One pass: everything still reachable from the package root, or
        // named by a staged addition or replacement.
        let mut reachable: HashSet<PackURI> = HashSet::new();
        let mut pending = internal_targets(package.rels(), false, None);
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
                pending.extend(internal_targets(
                    part.rels(),
                    name == main_name,
                    Some(&name),
                ));
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
                let mut targets = internal_targets(part.rels(), false, Some(&name));
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
                for owned in change.part_relationships.iter() {
                    let relationship = &owned.relationship;
                    if !forward {
                        staged
                            .dropped_part_relationships
                            .insert((owned.part.clone(), relationship.id.clone()));
                    } else if !relationship.external
                        && let Ok(owner) = PackURI::new(owned.part.clone())
                        && let Ok(target) =
                            PackURI::from_rel_ref(owner.base_uri(), &relationship.target)
                    {
                        staged.kept.insert(target.as_str().to_owned());
                    }
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
    /// Relationships of other parts a staged change removes: the part and
    /// the identifier.
    dropped_part_relationships: BTreeSet<(String, String)>,
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
