//! Detached composition of ordinary presentation-domain edits.

use std::collections::{BTreeSet, HashMap, HashSet, VecDeque};

use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};

use super::model::{Slide, Snapshot, capture, invalid, package_fingerprint};
use super::patch::Patch;
use crate::{Error, Result};

/// One failure-atomic edit rooted in an immutable opened-package snapshot.
#[derive(Clone)]
pub struct Transaction {
    source: Snapshot,
    working: OpcPackage,
    slides: Vec<Slide>,
}

impl std::fmt::Debug for Transaction {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Transaction")
            .field("source", &self.source)
            .field("slides", &self.slides)
            .finish_non_exhaustive()
    }
}

impl Transaction {
    pub(crate) fn new(source: Snapshot) -> Self {
        Self {
            working: source.package.as_ref().clone(),
            slides: source.slides.clone(),
            source,
        }
    }

    /// Immutable source root for this transaction.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Slides in the transaction's currently staged order.
    #[must_use]
    pub fn slides(&self) -> &[Slide] {
        &self.slides
    }

    /// Whether any managed resource differs from the source root.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        package_fingerprint(&self.working).is_ok_and(|revision| revision != self.source.revision)
    }

    /// Move one slide between checked zero-based positions.
    ///
    /// # Errors
    ///
    /// Returns an error for an out-of-range position or an unsupported
    /// slide-list dependency in the raw presentation XML.
    pub fn move_slide(&mut self, from: usize, to: usize) -> Result<bool> {
        let length = self.slides.len();
        if from >= length {
            return Err(Error::SlideIndexOutOfBounds {
                index: from,
                len: length,
            });
        }
        if to >= length {
            return Err(Error::SlideIndexOutOfBounds {
                index: to,
                len: length,
            });
        }
        if from == to {
            return Ok(false);
        }
        let mut order: Vec<_> = self.slides.iter().map(Slide::id).collect();
        let id = order.remove(from);
        order.insert(to, id);
        self.reorder_slides(&order)
    }

    /// Replace the complete slide order by stable slide IDs.
    ///
    /// # Errors
    ///
    /// Returns an error unless the IDs are an exact permutation of the
    /// currently staged slide identities.
    pub fn reorder_slides(&mut self, ordered_ids: &[u32]) -> Result<bool> {
        if ordered_ids == self.slides.iter().map(Slide::id).collect::<Vec<_>>() {
            return Ok(false);
        }
        let main = self.working.get_part(&self.source.presentation_name)?;
        let xml = super::xml::reorder_slides(main.blob(), &self.slides, ordered_ids)?;
        let mut reordered = Vec::new();
        reordered
            .try_reserve_exact(self.slides.len())
            .map_err(|source| Error::Allocation {
                resource: "opened-presentation reordered identities",
                source,
            })?;
        for id in ordered_ids {
            reordered.push(
                self.slides
                    .iter()
                    .find(|slide| slide.id == *id)
                    .cloned()
                    .ok_or_else(|| invalid("opened-presentation slide order lost an identity"))?,
            );
        }
        self.working
            .get_part_mut(&self.source.presentation_name)?
            .set_blob(xml);
        self.slides = reordered;
        Ok(true)
    }

    /// Remove one slide and any now-unreferenced acyclic dependency parts.
    ///
    /// Shared layouts, masters, themes, media, charts, and notes remain when
    /// any package edge still references them. At least one slide is retained.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing selector, the final remaining slide, or
    /// malformed slide-list and relationship topology.
    pub fn remove_slide<'s>(&mut self, slide: impl Into<crate::slide::Key<'s>>) -> Result<Slide> {
        if self.slides.len() == 1 {
            return Err(invalid("opened-presentation cannot remove the final slide"));
        }
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        let presentation = candidate.get_part(&self.source.presentation_name)?;
        let xml = super::xml::remove_slide(presentation.blob(), &self.slides, selected.id)?;
        let dependency_roots: Vec<_> = self
            .working
            .get_part(&selected.part_name)?
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
            .map(|relationship| relationship.target_partname().map_err(Error::Opc))
            .collect::<Result<_>>()?;
        let presentation = candidate.get_part_mut(&self.source.presentation_name)?;
        if presentation
            .rels_mut()
            .remove(&selected.relationship_id)
            .is_none()
        {
            return Err(invalid(
                "opened-presentation slide relationship disappeared during removal",
            ));
        }
        presentation.set_blob(xml);
        if !candidate.remove_part(&selected.part_name) {
            return Err(invalid(
                "opened-presentation slide part disappeared during removal",
            ));
        }
        remove_unreferenced_dependencies(&mut candidate, dependency_roots)?;
        self.working = candidate;
        self.slides.retain(|candidate| candidate.id != selected.id);
        Ok(selected)
    }

    /// Replace all visible text runs in one existing semantic shape.
    ///
    /// The first existing `a:t` run receives the escaped replacement and
    /// later runs become empty. Shape structure, formatting, unknown XML, and
    /// relationships remain exact.
    ///
    /// # Errors
    ///
    /// Returns an error for missing or ambiguous selectors, shapes without a
    /// text body, malformed raw XML, invalid characters, or exceeded bounds.
    pub fn set_shape_text<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        text: impl AsRef<str>,
    ) -> Result<bool> {
        let selected_slide = self.resolve_slide(slide.into())?;
        let key = shape.into();
        let text = text.as_ref();
        let owner = self.working.get_part(&selected_slide.part_name)?;
        crate::parts::validate_content_type(owner, litchi_opc::constants::content_type::PML_SLIDE)?;
        let scene = crate::shape::Scene::read(owner.blob())?;
        let selected_shape = scene.shape(key)?;
        if selected_shape.common().text() == Some(text) {
            return Ok(false);
        }
        if selected_shape.common().text().is_none() {
            return Err(invalid(
                "opened-presentation selected shape has no text body",
            ));
        }
        let span = crate::tag::shape::selected_raw_span(owner.blob(), key)?;
        let xml = super::xml::rewrite_shape_text(
            owner.blob(),
            span,
            text,
            self.source.limits.max_text_bytes(),
        )?;
        let staged = crate::shape::Scene::read(&xml)?;
        let staged_shape = staged.shape(key)?;
        if staged_shape.common().text() != Some(text) || staged.len() != scene.len() {
            return Err(invalid(
                "opened-presentation shape text did not round-trip semantically",
            ));
        }
        self.working
            .get_part_mut(&selected_slide.part_name)?
            .set_blob(xml);
        Ok(true)
    }

    /// Add a plain `DrawingML` text box without rebuilding the surrounding slide.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid text, geometry, malformed slide XML, or an
    /// exhausted non-visual shape-ID space.
    pub fn add_text_box<'s>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        text: impl AsRef<str>,
        bounds: (i64, i64, i64, i64),
    ) -> Result<u32> {
        let selected = self.resolve_slide(slide.into())?;
        let owner = self.working.get_part(&selected.part_name)?;
        let scene = crate::shape::Scene::read(owner.blob())?;
        let shape_id = next_shape_id(&scene, "text box")?;
        let conformance = crate::media_parts::document_conformance(owner.blob())?;
        let fragment = text_box_fragment(
            text.as_ref(),
            bounds,
            shape_id,
            conformance,
            self.source.limits,
        )?;
        self.append_checked_shape(&selected, shape_id, fragment.as_bytes())?;
        Ok(shape_id)
    }

    /// Add a rectangular preset shape with an optional six-digit sRGB fill.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid geometry/fill, malformed slide XML, or an
    /// exhausted non-visual shape-ID space.
    pub fn add_rectangle<'s>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        bounds: (i64, i64, i64, i64),
        fill_color: Option<&str>,
    ) -> Result<u32> {
        self.add_preset_shape(slide.into(), "rect", "Rectangle", bounds, fill_color)
    }

    /// Add an elliptical preset shape with an optional six-digit sRGB fill.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid geometry/fill, malformed slide XML, or an
    /// exhausted non-visual shape-ID space.
    pub fn add_ellipse<'s>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        bounds: (i64, i64, i64, i64),
        fill_color: Option<&str>,
    ) -> Result<u32> {
        self.add_preset_shape(slide.into(), "ellipse", "Ellipse", bounds, fill_color)
    }

    /// Add one visible picture and its collision-safe inert image part.
    ///
    /// The resource must use an `image/*` content type and a `/ppt/media/`
    /// preferred part name. Its bytes are retained inertly and never decoded.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid resource metadata, empty/oversized data,
    /// invalid geometry, malformed slide XML, or an exhausted shape-ID space.
    pub fn add_picture<'s>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        name: impl AsRef<str>,
        resource: &crate::media_parts::Resource,
        bounds: (i64, i64, i64, i64),
    ) -> Result<u32> {
        let selected = self.resolve_slide(slide.into())?;
        validate_bounds(bounds, "picture")?;
        validate_shape_text(name.as_ref(), self.source.limits, "picture name")?;
        if resource.data.is_empty() {
            return Err(invalid(
                "opened-presentation picture resource cannot be empty",
            ));
        }
        if resource.data.len() > self.source.limits.max_patch_bytes() {
            return Err(Error::Limit {
                resource: "opened-presentation picture bytes",
                limit: self.source.limits.max_patch_bytes(),
            });
        }
        if !resource.content_type.starts_with("image/") {
            return Err(invalid(
                "opened-presentation picture resource must use an image content type",
            ));
        }
        let preferred = PackURI::new(&resource.part_name).map_err(Error::Invalid)?;
        if !preferred.as_str().starts_with("/ppt/media/") {
            return Err(invalid(
                "opened-presentation picture resource must use a /ppt/media/ part name",
            ));
        }
        let owner = self.working.get_part(&selected.part_name)?;
        let scene = crate::shape::Scene::read(owner.blob())?;
        let shape_id = next_shape_id(&scene, "picture")?;
        let conformance = crate::media_parts::document_conformance(owner.blob())?;
        let reserved: HashSet<_> = self
            .working
            .iter_parts()
            .map(|part| part.partname().clone())
            .collect();
        let part_name = available_transfer_name(&preferred, &reserved)?;
        let mut candidate = self.working.clone();
        candidate.try_add_part(Box::new(BlobPart::new(
            part_name.clone(),
            resource.content_type.clone(),
            resource.data.as_slice().to_vec(),
        )))?;
        let target = part_name.relative_ref(selected.part_name.base_uri());
        let relationship_id = candidate
            .get_part_mut(&selected.part_name)?
            .relate_to(&target, conformance_image_relationship(conformance));
        let fragment = picture_fragment(
            name.as_ref(),
            bounds,
            shape_id,
            &relationship_id,
            conformance,
        );
        let xml = super::xml::append_shape(
            candidate.get_part(&selected.part_name)?.blob(),
            fragment.as_bytes(),
        )?;
        ensure_shape_id(&xml, shape_id, "picture")?;
        candidate.get_part_mut(&selected.part_name)?.set_blob(xml);
        self.working = candidate;
        Ok(shape_id)
    }

    /// Remove one top-level shape and any relationship closure made unreachable.
    ///
    /// Relationships still referenced elsewhere on the slide and parts shared
    /// from any package owner are retained.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/ambiguous selector or malformed shape and
    /// relationship topology.
    pub fn remove_shape<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        let key = shape.into();
        let owner = self.working.get_part(&selected.part_name)?;
        let scene = crate::shape::Scene::read(owner.blob())?;
        let _shape = scene.shape(key)?;
        let span = crate::tag::shape::selected_raw_span(owner.blob(), key)?;
        let removed_relationships = shape_relationship_ids(&owner.blob()[span.clone()])?;
        let mut xml = Vec::with_capacity(owner.blob().len().saturating_sub(span.len()));
        xml.extend_from_slice(&owner.blob()[..span.start]);
        xml.extend_from_slice(&owner.blob()[span.end..]);
        let retained_relationships = shape_relationship_ids(&xml)?;
        let mut candidate = self.working.clone();
        let mut dependency_roots = Vec::new();
        {
            let owner = candidate.get_part_mut(&selected.part_name)?;
            owner.set_blob(xml);
            for relationship_id in removed_relationships {
                if retained_relationships.contains(&relationship_id) {
                    continue;
                }
                let Some(relationship) = owner.rels_mut().remove(&relationship_id) else {
                    continue;
                };
                if !relationship.is_external() {
                    dependency_roots.push(relationship.target_partname()?);
                }
            }
        }
        remove_unreferenced_dependencies(&mut candidate, dependency_roots)?;
        self.working = candidate;
        Ok(true)
    }

    /// Copy one common top-level shape and every internal relationship target
    /// reachable from it into a destination slide.
    ///
    /// Part collisions, relationship IDs, every grouped shape identity, and
    /// connector endpoint references are remapped deterministically. Attached
    /// connectors pull in the complete top-level owner of each endpoint,
    /// including grouped children and their dependency relationships. Unknown
    /// shape kinds and unresolved endpoint identities remain explicit refusal
    /// boundaries.
    ///
    /// # Errors
    ///
    /// Returns an error for unsupported shapes, malformed XML/relationships,
    /// missing dependencies, or exceeded transaction bounds.
    pub fn transfer_shape<'ss, 'sk, 'ds>(
        &mut self,
        source: &Snapshot,
        source_slide: impl Into<crate::slide::Key<'ss>>,
        source_shape: impl Into<crate::shape::Key<'sk>>,
        destination_slide: impl Into<crate::slide::Key<'ds>>,
    ) -> Result<u32> {
        let source_slide = resolve_snapshot_slide(source, source_slide.into())?;
        let destination = self.resolve_slide(destination_slide.into())?;
        let source_owner = source.package.get_part(&source_slide.part_name)?;
        let source_scene = crate::shape::Scene::read(source_owner.blob())?;
        let key = source_shape.into();
        let shape = source_scene.shape(key)?;
        let source_roots: Vec<_> = source_scene.roots().collect();
        let selected_root = source_roots
            .iter()
            .position(|root| root.common().index() == shape.common().index());
        let Some(selected_root) = selected_root else {
            return Err(shape_transfer_refusal(
                crate::ShapeTransferRefusal::NestedShape,
                "transfer requires the selected shape's top-level owner",
            ));
        };
        let mut root_by_shape_id = HashMap::new();
        for (root_index, root) in source_roots.iter().copied().enumerate() {
            let mut root_ids = BTreeSet::new();
            collect_routable_shape_ids(root, &mut root_ids)?;
            for shape_id in root_ids {
                if root_by_shape_id.insert(shape_id, root_index).is_some() {
                    return Err(invalid(format!(
                        "opened-presentation source slide repeats shape identity {shape_id}"
                    )));
                }
            }
        }
        let mut pending = VecDeque::from([selected_root]);
        let mut included_roots = BTreeSet::new();
        let mut planned_roots = Vec::new();
        let mut source_shape_ids = BTreeSet::new();
        let mut relationship_ids = BTreeSet::new();
        while let Some(root_index) = pending.pop_front() {
            if !included_roots.insert(root_index) {
                continue;
            }
            if included_roots.len() > self.source.limits.max_parts() {
                return Err(Error::Limit {
                    resource: "opened-presentation transferred shape roots",
                    limit: self.source.limits.max_parts(),
                });
            }
            let root = source_roots
                .get(root_index)
                .copied()
                .ok_or_else(|| invalid("opened-presentation transferred shape root disappeared"))?;
            collect_transfer_shape_ids(root, &mut source_shape_ids)?;
            let span = crate::tag::shape::selected_raw_span(
                source_owner.blob(),
                crate::shape::Key::Index(root.common().index()),
            )?;
            let fragment = &source_owner.blob()[span.clone()];
            relationship_ids.extend(shape_relationship_ids(fragment)?);
            for connected_id in super::xml::connector_connection_ids(fragment)? {
                let endpoint_root = root_by_shape_id.get(&connected_id).copied().ok_or_else(|| {
                    shape_transfer_refusal(
                        crate::ShapeTransferRefusal::UnresolvedConnectorEndpoint,
                        format!(
                            "connector endpoint {connected_id} does not identify a transferable source shape"
                        ),
                    )
                })?;
                pending.push_back(endpoint_root);
            }
            planned_roots.push((root_index, root, span));
        }
        planned_roots.sort_unstable_by_key(|(root_index, _, _)| *root_index);
        let source_shape_id = shape.id().ok_or_else(|| {
            shape_transfer_refusal(
                crate::ShapeTransferRefusal::MissingIdentity,
                "selected shape has no non-visual identity",
            )
        })?;
        if !source_shape_ids.contains(&source_shape_id) {
            return Err(invalid(
                "opened-presentation transferred root identity is absent from its XML",
            ));
        }
        let mut roots = Vec::new();
        let mut relationships = Vec::new();
        for relationship_id in relationship_ids {
            let relationship = source_owner.rels().get(&relationship_id).ok_or_else(|| {
                invalid(format!(
                    "opened-presentation transferred shape relationship {relationship_id} is missing"
                ))
            })?;
            if !relationship.is_external() {
                roots.push(relationship.target_partname()?);
            }
            relationships.push(relationship.clone());
        }
        let mapping = plan_transfer_roots(
            source.package.as_ref(),
            &self.working,
            roots,
            self.source.limits.max_parts(),
        )?;
        let mut candidate = self.working.clone();
        publish_transfer(source.package.as_ref(), &mut candidate, &mapping)?;
        let destination_scene =
            crate::shape::Scene::read(candidate.get_part(&destination.part_name)?.blob())?;
        let shape_id_mapping = plan_shape_identity_transfer(&destination_scene, &source_shape_ids)?;
        let shape_id = *shape_id_mapping.get(&source_shape_id).ok_or_else(|| {
            invalid("opened-presentation transferred root identity mapping disappeared")
        })?;
        let mut relationship_mapping = HashMap::new();
        for relationship in relationships {
            let target = if relationship.is_external() {
                relationship.target_ref().to_owned()
            } else {
                let source_target = relationship.target_partname()?;
                let copied_target = mapping.get(&source_target).ok_or_else(|| {
                    invalid("opened-presentation transferred shape dependency is missing")
                })?;
                copied_target.relative_ref(destination.part_name.base_uri())
            };
            let new_id = candidate
                .get_part_mut(&destination.part_name)?
                .relate_to(&target, relationship.reltype());
            relationship_mapping.insert(relationship.r_id().to_owned(), new_id);
        }
        let mut xml = candidate.get_part(&destination.part_name)?.blob().to_vec();
        for (_root_index, root, span) in planned_roots {
            let root_source_id = root.id().ok_or_else(|| {
                invalid("opened-presentation transferred root has no non-visual identity")
            })?;
            let root_destination_id = shape_id_mapping.get(&root_source_id).ok_or_else(|| {
                invalid("opened-presentation transferred root identity mapping disappeared")
            })?;
            let name = format!(
                "{} Copy {root_destination_id}",
                root.name().unwrap_or("Shape")
            );
            let fragment = super::xml::remap_shape_fragment(
                &source_owner.blob()[span],
                root_source_id,
                &name,
                &shape_id_mapping,
                &relationship_mapping,
            )?;
            xml = super::xml::append_shape(&xml, &fragment)?;
        }
        ensure_shape_id(&xml, shape_id, "transferred shape")?;
        candidate
            .get_part_mut(&destination.part_name)?
            .set_blob(xml);
        self.working = candidate;
        Ok(shape_id)
    }

    /// Replace the existing notes text owned by one checked slide.
    ///
    /// # Errors
    ///
    /// Returns an error when the slide has no notes, the notes graph has an
    /// unsupported dependency, or the replacement is malformed or unbounded.
    pub fn set_notes_text<'s>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        text: impl AsRef<str>,
    ) -> Result<bool> {
        let selected_slide = self.resolve_slide(slide.into())?;
        let source = crate::notes::load_snapshot(&self.working, &self.source.presentation_name)?
            .ok_or_else(|| invalid("opened-presentation notes graph is absent"))?;
        let note = source
            .slides()
            .iter()
            .find(|note| note.owner() == selected_slide.part_name.as_str())
            .ok_or_else(|| invalid("opened-presentation selected slide has no notes"))?;
        let text = text.as_ref();
        if note.text()?.as_deref() == (!text.is_empty()).then_some(text) {
            return Ok(false);
        }
        if text.len() > self.source.limits.max_text_bytes() {
            return Err(Error::Limit {
                resource: "opened-presentation notes text bytes",
                limit: self.source.limits.max_text_bytes(),
            });
        }
        let part_name = PackURI::new(note.part()).map_err(Error::Invalid)?;
        let xml = crate::notes::rewrite_text(note.xml(), text)?;
        self.working.get_part_mut(&part_name)?.set_blob(xml);
        let staged = crate::notes::load_snapshot(&self.working, &self.source.presentation_name)?
            .ok_or_else(|| invalid("opened-presentation staged notes graph disappeared"))?;
        let staged_note = staged
            .slides()
            .iter()
            .find(|note| note.owner() == selected_slide.part_name.as_str())
            .ok_or_else(|| invalid("opened-presentation staged notes owner disappeared"))?;
        let actual = staged_note.text()?;
        if actual.as_deref() != (!text.is_empty()).then_some(text) {
            return Err(invalid(format!(
                "opened-presentation notes text did not round-trip semantically: requested {text:?}, read {actual:?}"
            )));
        }
        Ok(true)
    }

    /// Create or replace the validated table-style catalog in this transaction.
    ///
    /// # Errors
    ///
    /// Returns an error when the catalog or its attachment graph is invalid.
    pub fn put_table_styles(&mut self, styles: crate::table::style::List) -> Result<bool> {
        crate::table::style::put(&mut self.working, styles)
    }

    /// Remove the optional table-style catalog and its relationship.
    ///
    /// # Errors
    ///
    /// Returns an error when the existing attachment graph is unsafe.
    pub fn remove_table_styles(&mut self) -> Result<Option<crate::table::style::List>> {
        crate::table::style::remove(&mut self.working)
    }

    /// Add a visible rectangular table with uniformly sized rows and columns.
    ///
    /// Cell text is XML-escaped, row widths must be rectangular, and geometry
    /// must use positive extents. The surrounding producer XML is preserved.
    ///
    /// # Errors
    ///
    /// Returns an error for empty/ragged input, invalid XML characters,
    /// exceeded text bounds, malformed slide XML, or exhausted shape IDs.
    pub fn add_table<'s>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        cells: &[Vec<String>],
        bounds: (i64, i64, i64, i64),
    ) -> Result<u32> {
        let selected = self.resolve_slide(slide.into())?;
        let owner = self.working.get_part(&selected.part_name)?;
        let scene = crate::shape::Scene::read(owner.blob())?;
        let shape_id = next_shape_id(&scene, "table")?;
        let conformance = crate::media_parts::document_conformance(owner.blob())?;
        let fragment = table_fragment(cells, bounds, shape_id, conformance, self.source.limits)?;
        let _table = crate::table::Table::from_graphic_frame(fragment.as_bytes())?;
        let xml = super::xml::append_shape(owner.blob(), fragment.as_bytes())?;
        let staged = crate::shape::Scene::read(&xml)?;
        if !staged.iter().any(|shape| shape.id() == Some(shape_id)) {
            return Err(invalid(
                "opened-presentation table did not round-trip semantically",
            ));
        }
        self.working
            .get_part_mut(&selected.part_name)?
            .set_blob(xml);
        Ok(shape_id)
    }

    /// Add an ordinary chart part, relationship, and visible graphic frame.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing slide, malformed shape tree, invalid
    /// chart, or exhausted non-visual shape IDs.
    pub fn add_chart<'s>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        chart: &crate::chart::Chart,
    ) -> Result<String> {
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        let scene = crate::shape::Scene::read(candidate.get_part(&selected.part_name)?.blob())?;
        let shape_id = next_shape_id(&scene, "chart")?;
        let relationship_id =
            crate::chart::add(&mut candidate, selected.part_name.as_str(), chart)?;
        let fragment = crate::chart::write_graphic_frame(shape_id, &relationship_id, chart);
        let owner = candidate.get_part(&selected.part_name)?;
        let xml = super::xml::append_shape(owner.blob(), fragment.as_bytes())?;
        let scene = crate::shape::Scene::read(&xml)?;
        if !scene.iter().any(|shape| shape.id() == Some(shape_id)) {
            return Err(invalid(
                "opened-presentation chart frame did not round-trip semantically",
            ));
        }
        candidate.get_part_mut(&selected.part_name)?.set_blob(xml);
        self.working = candidate;
        Ok(relationship_id)
    }

    /// Store bounded slide media, its visible picture frames, and inert resources.
    ///
    /// # Errors
    ///
    /// Returns an error when the slide already contains media or the typed
    /// media graph is invalid.
    pub fn store_media<'s>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        media: &crate::media_parts::List,
        conformance: crate::media_parts::Conformance,
    ) -> Result<()> {
        let selected = self.resolve_slide(slide.into())?;
        crate::media_parts::store(&mut self.working, &selected.part_name, media, conformance)
    }

    /// Add a slide master and all required theme relationships.
    ///
    /// # Errors
    ///
    /// Returns an error when the master/layout graph cannot be extended safely.
    pub fn add_slide_master(&mut self) -> Result<crate::master_layout::AuthoredSlideMaster> {
        crate::master_layout::validate_master_layout_graph(&self.working)?;
        crate::master_layout::add_slide_master(&mut self.working)
    }

    /// Add a typed slide layout to an existing master.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid master, name, placeholder set, or graph.
    pub fn add_slide_layout(
        &mut self,
        master: &PackURI,
        kind: crate::master_layout::SlideLayoutKind,
        name: &str,
        placeholders: &[crate::master_layout::PlaceholderSpec],
    ) -> Result<crate::master_layout::AuthoredSlideLayout> {
        crate::master_layout::validate_master_layout_graph(&self.working)?;
        crate::master_layout::add_slide_layout(&mut self.working, master, kind, name, placeholders)
    }

    /// Remove an unreferenced slide layout and its exact graph edges.
    ///
    /// # Errors
    ///
    /// Returns an error when the layout is still used or its graph is malformed.
    pub fn remove_slide_layout(&mut self, layout: &PackURI) -> Result<()> {
        crate::master_layout::validate_master_layout_graph(&self.working)?;
        crate::master_layout::remove_slide_layout(&mut self.working, layout)
    }

    /// Add a legacy comment author to the presentation comment graph.
    ///
    /// # Errors
    ///
    /// Returns an error when the author or existing comment graph is invalid.
    pub fn add_comment_author(
        &mut self,
        author: crate::comments::Author,
        conformance: crate::comments::Conformance,
    ) -> Result<()> {
        crate::comments::add_presentation_comment_author(&mut self.working, author, conformance)
    }

    /// Add a legacy comment owned by one selected slide.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing author, duplicate identity, or invalid graph.
    pub fn add_comment<'s>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        comment: crate::comments::Comment,
        conformance: crate::comments::Conformance,
    ) -> Result<()> {
        let selected = self.resolve_slide(slide.into())?;
        crate::comments::add_presentation_comment(
            &mut self.working,
            selected.part_name.as_str(),
            comment,
            conformance,
        )
    }

    /// Remove one legacy comment by stable author/index identity.
    ///
    /// # Errors
    ///
    /// Returns an error when the comment graph is malformed.
    pub fn remove_comment<'s>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        author_id: u32,
        index: u32,
        conformance: crate::comments::Conformance,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        crate::comments::remove_presentation_comment(
            &mut self.working,
            selected.part_name.as_str(),
            author_id,
            index,
            conformance,
        )
    }

    /// Add a typed `PowerPoint` 2018 modern-comment author.
    ///
    /// A missing author part and presentation relationship are allocated
    /// collision-safely by the modern-comment owner.
    ///
    /// # Errors
    ///
    /// Returns an error for duplicate/invalid authors or a malformed graph.
    pub fn add_modern_comment_author(
        &mut self,
        author: crate::modern_comments::Author,
    ) -> Result<crate::modern_comments::AuthorPart> {
        let mut candidate = self.working.clone();
        let result = crate::modern_comments::add_modern_comment_author(&mut candidate, author)?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Replace a modern-comment author without changing its stable GUID.
    ///
    /// # Errors
    ///
    /// Returns an error for an identity mismatch, unresolved references, or a
    /// malformed modern-comment graph.
    pub fn replace_modern_comment_author(
        &mut self,
        author_id: &str,
        replacement: crate::modern_comments::Author,
    ) -> Result<bool> {
        let mut candidate = self.working.clone();
        let result = crate::modern_comments::replace_modern_comment_author(
            &mut candidate,
            author_id,
            replacement,
        )?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Remove one unreferenced modern-comment author.
    ///
    /// # Errors
    ///
    /// Returns an error while any modeled comment/reply still references the
    /// author or when the graph is malformed.
    pub fn remove_modern_comment_author(&mut self, author_id: &str) -> Result<bool> {
        let mut candidate = self.working.clone();
        let result =
            crate::modern_comments::remove_modern_comment_author(&mut candidate, author_id)?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Reorder all modern-comment authors by a complete stable-GUID list.
    ///
    /// # Errors
    ///
    /// Returns an error for unknown, duplicate, or incomplete identities.
    pub fn reorder_modern_comment_authors(
        &mut self,
        ordered_author_ids: &[String],
    ) -> Result<Vec<crate::modern_comments::Author>> {
        let mut candidate = self.working.clone();
        let result = crate::modern_comments::reorder_modern_comment_authors(
            &mut candidate,
            ordered_author_ids,
        )?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Add a typed `PowerPoint` 2018 modern comment to one selected slide.
    ///
    /// # Errors
    ///
    /// Returns an error for unresolved authors, duplicate identities, invalid
    /// extension payloads, or a malformed package graph.
    pub fn add_modern_comment<'s>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        comment: crate::modern_comments::Comment,
    ) -> Result<crate::modern_comments::Part> {
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        let result = crate::modern_comments::add_modern_comment(
            &mut candidate,
            &selected.part_name,
            comment,
        )?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Replace one modern comment without changing its stable GUID.
    ///
    /// # Errors
    ///
    /// Returns an error for an identity mismatch, invalid author/extension
    /// references, or a malformed package graph.
    pub fn replace_modern_comment<'s>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        comment_id: &str,
        replacement: crate::modern_comments::Comment,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        let result = crate::modern_comments::replace_modern_comment(
            &mut candidate,
            &selected.part_name,
            comment_id,
            replacement,
        )?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Reorder all modern comments on one slide by a complete stable-GUID list.
    ///
    /// # Errors
    ///
    /// Returns an error for unknown, duplicate, or incomplete identities.
    pub fn reorder_modern_comments<'s>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        ordered_comment_ids: &[String],
    ) -> Result<Vec<crate::modern_comments::Comment>> {
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        let result = crate::modern_comments::reorder_modern_comments(
            &mut candidate,
            &selected.part_name,
            ordered_comment_ids,
        )?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Add a stable-ID reply to one modern-comment thread.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing thread/author, duplicate reply identity,
    /// invalid extensions, or a malformed package graph.
    pub fn add_modern_comment_reply<'s>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        comment_id: &str,
        reply: crate::modern_comments::Reply,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        let result = crate::modern_comments::add_modern_comment_reply(
            &mut candidate,
            &selected.part_name,
            comment_id,
            reply,
        )?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Replace one modern-comment reply without changing its stable GUID.
    ///
    /// # Errors
    ///
    /// Returns an error for an identity mismatch, unresolved author, invalid
    /// extensions, or a malformed graph.
    pub fn replace_modern_comment_reply<'s>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        comment_id: &str,
        reply_id: &str,
        replacement: crate::modern_comments::Reply,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        let result = crate::modern_comments::replace_modern_comment_reply(
            &mut candidate,
            &selected.part_name,
            comment_id,
            reply_id,
            replacement,
        )?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Update typed task/reaction extensions on one modern comment while
    /// retaining opaque extension entries.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/malformed thread or invalid extension.
    pub fn update_modern_comment_extensions<'s, F>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        comment_id: &str,
        update: F,
    ) -> Result<bool>
    where
        F: FnOnce(&mut crate::modern_comments::semantic::extensions::List),
    {
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        let result = crate::modern_comments::update_modern_comment_extensions(
            &mut candidate,
            &selected.part_name,
            comment_id,
            update,
        )?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Update typed reaction extensions on one modern-comment reply while
    /// retaining opaque extension entries.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/malformed reply or invalid extension.
    pub fn update_modern_comment_reply_extensions<'s, F>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        comment_id: &str,
        reply_id: &str,
        update: F,
    ) -> Result<bool>
    where
        F: FnOnce(&mut crate::modern_comments::semantic::extensions::List),
    {
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        let result = crate::modern_comments::update_modern_comment_reply_extensions(
            &mut candidate,
            &selected.part_name,
            comment_id,
            reply_id,
            update,
        )?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Remove one modern comment and its empty per-slide part when unshared.
    ///
    /// # Errors
    ///
    /// Returns an error when the modern-comment graph is malformed.
    pub fn remove_modern_comment<'s>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        comment_id: &str,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        let result = crate::modern_comments::remove_modern_comment(
            &mut candidate,
            &selected.part_name,
            comment_id,
        )?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Remove one modern-comment reply by stable identity.
    ///
    /// # Errors
    ///
    /// Returns an error when the modern-comment graph is malformed.
    pub fn remove_modern_comment_reply<'s>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        comment_id: &str,
        reply_id: &str,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        let result = crate::modern_comments::remove_modern_comment_reply(
            &mut candidate,
            &selected.part_name,
            comment_id,
            reply_id,
        )?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Copy one internal relationship target and its complete dependency closure.
    ///
    /// Part-name collisions are deterministically remapped, relationship IDs
    /// inside copied parts remain stable, and every internal edge is retargeted
    /// to the copied dependency. The returned relationship ID is attached to
    /// `destination_slide`; callers that transfer a visually referenced object
    /// must also use an object-specific API such as [`Self::add_chart`] to add
    /// its slide XML anchor.
    ///
    /// # Errors
    ///
    /// Returns an error for external roots, missing dependencies, fragments or
    /// queries on internal targets, or exceeded transaction bounds.
    pub fn transfer_relationship_closure<'s>(
        &mut self,
        source: &Snapshot,
        source_owner: &PackURI,
        relationship_id: &str,
        destination_slide: impl Into<crate::slide::Key<'s>>,
    ) -> Result<String> {
        let destination = self.resolve_slide(destination_slide.into())?;
        let owner = source.package.get_part(source_owner)?;
        let relationship = owner
            .rels()
            .get(relationship_id)
            .ok_or_else(|| invalid("opened-presentation transfer relationship is missing"))?;
        if relationship.is_external() {
            return Err(invalid(
                "opened-presentation transfer root relationship is external",
            ));
        }
        let root = relationship.target_partname()?;
        let mapping = plan_transfer(
            source.package.as_ref(),
            &self.working,
            &root,
            self.source.limits.max_parts(),
        )?;
        let mut candidate = self.working.clone();
        publish_transfer(source.package.as_ref(), &mut candidate, &mapping)?;
        let copied_root = mapping
            .get(&root)
            .ok_or_else(|| invalid("opened-presentation transfer root disappeared"))?;
        let target = copied_root.relative_ref(destination.part_name.base_uri());
        let id = candidate
            .get_part_mut(&destination.part_name)?
            .relate_to(&target, relationship.reltype());
        self.working = candidate;
        Ok(id)
    }

    /// Validate and consume all staged edits into one atomic commit.
    ///
    /// # Errors
    ///
    /// Returns an error if a dependency changed, the patch exceeds bounds, or
    /// the complete staged package cannot be captured as a coherent snapshot.
    pub fn commit(self) -> Result<Commit> {
        let mut working = self.working;
        compact_changed_slides(&mut working, self.source.package.as_ref(), &self.slides)?;
        if package_fingerprint(&working)? != self.source.revision {
            working.unsign();
        }
        let patch = Patch::capture(
            self.source.package.as_ref(),
            &working,
            self.source.presentation_name.clone(),
            self.source.limits,
        )?;
        if patch.is_empty() {
            return Ok(Commit {
                snapshot: self.source,
                patch,
            });
        }
        let snapshot = capture(&working, self.source.limits)?;
        Ok(Commit { snapshot, patch })
    }

    /// Discard all staged edits and recover the immutable source root.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    fn add_preset_shape(
        &mut self,
        slide: crate::slide::Key<'_>,
        preset: &str,
        label: &str,
        bounds: (i64, i64, i64, i64),
        fill_color: Option<&str>,
    ) -> Result<u32> {
        let selected = self.resolve_slide(slide)?;
        let owner = self.working.get_part(&selected.part_name)?;
        let scene = crate::shape::Scene::read(owner.blob())?;
        let shape_id = next_shape_id(&scene, label)?;
        let conformance = crate::media_parts::document_conformance(owner.blob())?;
        let fragment =
            preset_shape_fragment(preset, label, bounds, fill_color, shape_id, conformance)?;
        self.append_checked_shape(&selected, shape_id, fragment.as_bytes())?;
        Ok(shape_id)
    }

    fn append_checked_shape(
        &mut self,
        slide: &Slide,
        shape_id: u32,
        fragment: &[u8],
    ) -> Result<()> {
        let xml =
            super::xml::append_shape(self.working.get_part(&slide.part_name)?.blob(), fragment)?;
        ensure_shape_id(&xml, shape_id, "shape")?;
        self.working.get_part_mut(&slide.part_name)?.set_blob(xml);
        Ok(())
    }

    fn resolve_slide(&self, key: crate::slide::Key<'_>) -> Result<Slide> {
        match key {
            crate::slide::Key::Index(index) => {
                self.slides
                    .get(index)
                    .cloned()
                    .ok_or(Error::SlideIndexOutOfBounds {
                        index,
                        len: self.slides.len(),
                    })
            },
            crate::slide::Key::Name(name) => {
                let mut matches = self.slides.iter().filter(|slide| slide.name == name);
                let selected = matches.next().cloned();
                if matches.next().is_some() {
                    return Err(Error::AmbiguousSlideName {
                        name: name.to_owned(),
                        matches: self
                            .slides
                            .iter()
                            .filter(|slide| slide.name == name)
                            .count(),
                    });
                }
                selected.ok_or_else(|| Error::SlideNameNotFound(name.to_owned()))
            },
        }
    }
}

fn compact_changed_slides(
    working: &mut OpcPackage,
    source: &OpcPackage,
    slides: &[Slide],
) -> Result<()> {
    for slide in slides {
        let Ok(before) = source.get_part(&slide.part_name) else {
            continue;
        };
        let after = working.get_part(&slide.part_name)?;
        if before.blob() == after.blob() {
            continue;
        }
        let expected = shape_semantics(after.blob())?;
        let compact = super::xml::compact_changed_slide_xml(after.blob())?;
        if shape_semantics(&compact)? != expected {
            return Err(invalid(
                "opened-presentation slide compaction changed shape semantics",
            ));
        }
        working.get_part_mut(&slide.part_name)?.set_blob(compact);
    }
    Ok(())
}

fn shape_semantics(xml: &[u8]) -> Result<Vec<(Option<u32>, Option<String>, Option<String>)>> {
    let scene = crate::shape::Scene::read(xml)?;
    Ok(scene
        .iter()
        .map(|shape| {
            (
                shape.id(),
                shape.name().map(str::to_owned),
                shape.common().text().map(str::to_owned),
            )
        })
        .collect())
}

fn remove_unreferenced_dependencies(package: &mut OpcPackage, roots: Vec<PackURI>) -> Result<()> {
    let mut queue = VecDeque::from(roots);
    let mut checked = HashSet::new();
    while let Some(name) = queue.pop_front() {
        if !checked.insert(name.clone()) || package.get_part(&name).is_err() {
            continue;
        }
        if package
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
            .any(|relationship| {
                relationship
                    .target_partname()
                    .is_ok_and(|target| target == name)
            })
            || package.iter_parts().any(|part| {
                part.rels()
                    .iter()
                    .filter(|relationship| !relationship.is_external())
                    .any(|relationship| {
                        relationship
                            .target_partname()
                            .is_ok_and(|target| target == name)
                    })
            })
        {
            continue;
        }
        let dependencies: Vec<_> = package
            .get_part(&name)?
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
            .map(|relationship| relationship.target_partname().map_err(Error::Opc))
            .collect::<Result<_>>()?;
        if package.remove_part(&name) {
            queue.extend(dependencies);
        }
    }
    Ok(())
}

fn next_shape_id(scene: &crate::shape::Scene<'_>, label: &str) -> Result<u32> {
    scene
        .iter()
        .filter_map(crate::shape::Shape::id)
        .max()
        .unwrap_or(1)
        .checked_add(1)
        .ok_or_else(|| invalid(format!("opened-presentation {label} shape ID overflow")))
}

fn collect_transfer_shape_ids(
    shape: crate::shape::Shape<'_>,
    identities: &mut BTreeSet<u32>,
) -> Result<()> {
    match shape {
        crate::shape::Shape::Content(_) => Err(shape_transfer_refusal(
            crate::ShapeTransferRefusal::ContentPart,
            "p:contentPart has no common non-visual identity; use its typed content owner",
        )),
        crate::shape::Shape::Unknown(value) => Err(shape_transfer_refusal(
            crate::ShapeTransferRefusal::UnknownExtensionShape,
            format!(
                "unknown source element {} is retained but cannot be dependency-closed",
                value
                    .common()
                    .source_name()
                    .unwrap_or("without a source name")
            ),
        )),
        crate::shape::Shape::Group(group) => {
            insert_transfer_shape_id(shape, identities)?;
            for child in group.shapes() {
                collect_transfer_shape_ids(child, identities)?;
            }
            Ok(())
        },
        _ => insert_transfer_shape_id(shape, identities),
    }
}

fn collect_routable_shape_ids(
    shape: crate::shape::Shape<'_>,
    identities: &mut BTreeSet<u32>,
) -> Result<()> {
    if let Some(identity) = shape.id()
        && !identities.insert(identity)
    {
        return Err(invalid(format!(
            "opened-presentation source shape root repeats identity {identity}"
        )));
    }
    if let crate::shape::Shape::Group(group) = shape {
        for child in group.shapes() {
            collect_routable_shape_ids(child, identities)?;
        }
    }
    Ok(())
}

fn insert_transfer_shape_id(
    shape: crate::shape::Shape<'_>,
    identities: &mut BTreeSet<u32>,
) -> Result<()> {
    let identity = shape.id().ok_or_else(|| {
        shape_transfer_refusal(
            crate::ShapeTransferRefusal::MissingIdentity,
            format!(
                "shape {} has no non-visual identity",
                shape.name().unwrap_or("without a name")
            ),
        )
    })?;
    if identities.insert(identity) {
        Ok(())
    } else {
        Err(invalid(format!(
            "opened-presentation transferred shape repeats identity {identity}"
        )))
    }
}

fn shape_transfer_refusal(kind: crate::ShapeTransferRefusal, detail: impl Into<String>) -> Error {
    Error::ShapeTransfer {
        kind,
        detail: detail.into(),
    }
}

fn plan_shape_identity_transfer(
    destination: &crate::shape::Scene<'_>,
    source_ids: &BTreeSet<u32>,
) -> Result<HashMap<u32, u32>> {
    if source_ids.is_empty() {
        return Err(invalid(
            "opened-presentation transferred shape has no non-visual identities",
        ));
    }
    let mut next = destination
        .iter()
        .filter_map(crate::shape::Shape::id)
        .max()
        .unwrap_or(1);
    let mut mapping = HashMap::new();
    mapping
        .try_reserve(source_ids.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation transferred shape identities",
            source,
        })?;
    for source_id in source_ids {
        next = next.checked_add(1).ok_or_else(|| {
            invalid("opened-presentation transferred shape identity space is exhausted")
        })?;
        mapping.insert(*source_id, next);
    }
    Ok(mapping)
}

fn resolve_snapshot_slide(source: &Snapshot, key: crate::slide::Key<'_>) -> Result<Slide> {
    match key {
        crate::slide::Key::Index(index) => {
            source
                .slides
                .get(index)
                .cloned()
                .ok_or(Error::SlideIndexOutOfBounds {
                    index,
                    len: source.slides.len(),
                })
        },
        crate::slide::Key::Name(name) => {
            let matches: Vec<_> = source
                .slides
                .iter()
                .filter(|slide| slide.name == name)
                .cloned()
                .collect();
            match matches.as_slice() {
                [slide] => Ok(slide.clone()),
                [] => Err(Error::SlideNameNotFound(name.to_owned())),
                _ => Err(Error::AmbiguousSlideName {
                    name: name.to_owned(),
                    matches: matches.len(),
                }),
            }
        },
    }
}

fn ensure_shape_id(xml: &[u8], shape_id: u32, label: &str) -> Result<()> {
    let scene = crate::shape::Scene::read(xml)?;
    if scene.iter().any(|shape| shape.id() == Some(shape_id)) {
        Ok(())
    } else {
        Err(invalid(format!(
            "opened-presentation {label} did not round-trip semantically"
        )))
    }
}

fn validate_bounds((_, _, width, height): (i64, i64, i64, i64), label: &str) -> Result<()> {
    if width <= 0 || height <= 0 {
        Err(invalid(format!(
            "opened-presentation {label} extents must be positive"
        )))
    } else {
        Ok(())
    }
}

fn validate_shape_text(text: &str, limits: super::Limits, label: &str) -> Result<()> {
    if text.len() > limits.max_text_bytes() {
        return Err(Error::Limit {
            resource: "opened-presentation shape text bytes",
            limit: limits.max_text_bytes(),
        });
    }
    if !text.chars().all(is_xml_char) {
        return Err(invalid(format!(
            "opened-presentation {label} contains an invalid XML character"
        )));
    }
    Ok(())
}

fn conformance_namespaces(
    conformance: crate::media_parts::Conformance,
) -> (&'static str, &'static str, &'static str) {
    match conformance {
        crate::media_parts::Conformance::Transitional => (
            "http://schemas.openxmlformats.org/presentationml/2006/main",
            "http://schemas.openxmlformats.org/drawingml/2006/main",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ),
        crate::media_parts::Conformance::Strict => (
            "http://purl.oclc.org/ooxml/presentationml/main",
            "http://purl.oclc.org/ooxml/drawingml/main",
            "http://purl.oclc.org/ooxml/officeDocument/relationships",
        ),
    }
}

fn conformance_image_relationship(conformance: crate::media_parts::Conformance) -> &'static str {
    match conformance {
        crate::media_parts::Conformance::Transitional => {
            litchi_opc::constants::relationship_type::IMAGE
        },
        crate::media_parts::Conformance::Strict => {
            litchi_opc::constants::relationship_type::STRICT_IMAGE
        },
    }
}

fn text_box_fragment(
    text: &str,
    bounds: (i64, i64, i64, i64),
    shape_id: u32,
    conformance: crate::media_parts::Conformance,
    limits: super::Limits,
) -> Result<String> {
    validate_bounds(bounds, "text box")?;
    validate_shape_text(text, limits, "text box")?;
    let (pml, dml, _) = conformance_namespaces(conformance);
    let (x, y, width, height) = bounds;
    Ok(format!(
        "<p:sp xmlns:p=\"{pml}\" xmlns:a=\"{dml}\"><p:nvSpPr><p:cNvPr id=\"{shape_id}\" name=\"Text Box {shape_id}\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{width}\" cy=\"{height}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap=\"square\"/><a:lstStyle/><a:p><a:r><a:t xml:space=\"preserve\">{}</a:t></a:r><a:endParaRPr/></a:p></p:txBody></p:sp>",
        quick_xml::escape::escape(text)
    ))
}

fn preset_shape_fragment(
    preset: &str,
    label: &str,
    bounds: (i64, i64, i64, i64),
    fill_color: Option<&str>,
    shape_id: u32,
    conformance: crate::media_parts::Conformance,
) -> Result<String> {
    validate_bounds(bounds, label)?;
    if !matches!(preset, "rect" | "ellipse") {
        return Err(invalid("opened-presentation preset shape is unsupported"));
    }
    if fill_color.is_some_and(|color| {
        color.len() != 6 || !color.bytes().all(|byte| byte.is_ascii_hexdigit())
    }) {
        return Err(invalid(
            "opened-presentation preset fill must be six hexadecimal digits",
        ));
    }
    let (pml, dml, _) = conformance_namespaces(conformance);
    let (x, y, width, height) = bounds;
    let fill = fill_color.map_or_else(
        || "<a:noFill/>".to_owned(),
        |color| format!("<a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill>"),
    );
    Ok(format!(
        "<p:sp xmlns:p=\"{pml}\" xmlns:a=\"{dml}\"><p:nvSpPr><p:cNvPr id=\"{shape_id}\" name=\"{label} {shape_id}\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{width}\" cy=\"{height}\"/></a:xfrm><a:prstGeom prst=\"{preset}\"><a:avLst/></a:prstGeom>{fill}</p:spPr></p:sp>"
    ))
}

fn picture_fragment(
    name: &str,
    bounds: (i64, i64, i64, i64),
    shape_id: u32,
    relationship_id: &str,
    conformance: crate::media_parts::Conformance,
) -> String {
    let (pml, dml, rel) = conformance_namespaces(conformance);
    let (x, y, width, height) = bounds;
    format!(
        "<p:pic xmlns:p=\"{pml}\" xmlns:a=\"{dml}\" xmlns:r=\"{rel}\"><p:nvPicPr><p:cNvPr id=\"{shape_id}\" name=\"{}\"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed=\"{relationship_id}\"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{width}\" cy=\"{height}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr></p:pic>",
        quick_xml::escape::escape(name)
    )
}

fn shape_relationship_ids(fragment: &[u8]) -> Result<BTreeSet<String>> {
    let mut reader = quick_xml::Reader::from_reader(fragment);
    reader.config_mut().trim_text(false);
    let mut relationships = BTreeSet::new();
    loop {
        match reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            quick_xml::events::Event::Start(element) | quick_xml::events::Event::Empty(element) => {
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                    let key = attribute.key.as_ref();
                    if !key.starts_with(b"r:") || !matches!(&key[2..], b"id" | b"embed" | b"link") {
                        continue;
                    }
                    let value = attribute
                        .decoded_and_normalized_value(
                            quick_xml::XmlVersion::Implicit1_0,
                            reader.decoder(),
                        )
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    relationships.insert(value.into_owned());
                }
            },
            quick_xml::events::Event::DocType(_) => {
                return Err(invalid(
                    "opened-presentation shape contains an unsupported document type",
                ));
            },
            quick_xml::events::Event::Eof => break,
            _ => {},
        }
    }
    Ok(relationships)
}

fn table_fragment(
    cells: &[Vec<String>],
    (x, y, width, height): (i64, i64, i64, i64),
    shape_id: u32,
    conformance: crate::media_parts::Conformance,
    limits: super::Limits,
) -> Result<String> {
    let columns = cells
        .first()
        .map(Vec::len)
        .filter(|columns| *columns != 0)
        .ok_or_else(|| invalid("opened-presentation table must contain cells"))?;
    if cells.iter().any(|row| row.len() != columns) {
        return Err(invalid("opened-presentation table rows are ragged"));
    }
    if width <= 0 || height <= 0 {
        return Err(invalid(
            "opened-presentation table extents must be positive",
        ));
    }
    let text_bytes = cells
        .iter()
        .flatten()
        .try_fold(0usize, |total, cell| total.checked_add(cell.len()))
        .ok_or_else(|| invalid("opened-presentation table text size overflow"))?;
    if text_bytes > limits.max_text_bytes() {
        return Err(Error::Limit {
            resource: "opened-presentation table text bytes",
            limit: limits.max_text_bytes(),
        });
    }
    if !cells
        .iter()
        .flatten()
        .flat_map(|cell| cell.chars())
        .all(is_xml_char)
    {
        return Err(invalid(
            "opened-presentation table contains an invalid XML character",
        ));
    }
    let columns_i64 = i64::try_from(columns)
        .map_err(|_err| invalid("opened-presentation table column count exceeds i64"))?;
    let rows_i64 = i64::try_from(cells.len())
        .map_err(|_err| invalid("opened-presentation table row count exceeds i64"))?;
    let column_width = width / columns_i64;
    let row_height = height / rows_i64;
    let (presentation_namespace, drawing_namespace) = match conformance {
        crate::media_parts::Conformance::Transitional => (
            "http://schemas.openxmlformats.org/presentationml/2006/main",
            "http://schemas.openxmlformats.org/drawingml/2006/main",
        ),
        crate::media_parts::Conformance::Strict => (
            "http://purl.oclc.org/ooxml/presentationml/main",
            "http://purl.oclc.org/ooxml/drawingml/main",
        ),
    };
    let mut xml = format!(
        "<p:graphicFrame xmlns:p=\"{presentation_namespace}\" xmlns:a=\"{drawing_namespace}\"><p:nvGraphicFramePr><p:cNvPr id=\"{shape_id}\" name=\"Table {shape_id}\"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{width}\" cy=\"{height}\"/></p:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\"><a:tbl><a:tblPr firstRow=\"1\" bandRow=\"1\"/><a:tblGrid>"
    );
    for _ in 0..columns {
        xml.push_str(&format!("<a:gridCol w=\"{column_width}\"/>"));
    }
    xml.push_str("</a:tblGrid>");
    for row in cells {
        xml.push_str(&format!("<a:tr h=\"{row_height}\">"));
        for cell in row {
            xml.push_str("<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>");
            xml.push_str(&quick_xml::escape::escape(cell));
            xml.push_str("</a:t></a:r><a:endParaRPr/></a:p></a:txBody><a:tcPr/></a:tc>");
        }
        xml.push_str("</a:tr>");
    }
    xml.push_str("</a:tbl></a:graphicData></a:graphic></p:graphicFrame>");
    Ok(xml)
}

fn is_xml_char(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
}

fn plan_transfer(
    source: &OpcPackage,
    destination: &OpcPackage,
    root: &PackURI,
    limit: usize,
) -> Result<HashMap<PackURI, PackURI>> {
    plan_transfer_roots(source, destination, vec![root.clone()], limit)
}

fn plan_transfer_roots(
    source: &OpcPackage,
    destination: &OpcPackage,
    roots: Vec<PackURI>,
    limit: usize,
) -> Result<HashMap<PackURI, PackURI>> {
    let mut queue = VecDeque::from(roots);
    let mut seen = HashSet::new();
    while let Some(name) = queue.pop_front() {
        if !seen.insert(name.clone()) {
            continue;
        }
        if seen.len() > limit {
            return Err(Error::Limit {
                resource: "opened-presentation transfer dependency parts",
                limit,
            });
        }
        let part = source.get_part(&name)?;
        for relationship in part.rels().iter().filter(|item| !item.is_external()) {
            if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
                return Err(invalid(
                    "opened-presentation transfer dependency has a query or fragment",
                ));
            }
            queue.push_back(relationship.target_partname()?);
        }
    }
    let mut names: Vec<_> = seen.into_iter().collect();
    names.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
    let mut reserved: HashSet<_> = destination
        .iter_parts()
        .map(|part| part.partname().clone())
        .collect();
    let mut mapping = HashMap::new();
    for source_name in names {
        let target = available_transfer_name(&source_name, &reserved)?;
        reserved.insert(target.clone());
        mapping.insert(source_name, target);
    }
    Ok(mapping)
}

fn available_transfer_name(source: &PackURI, reserved: &HashSet<PackURI>) -> Result<PackURI> {
    if !reserved.contains(source) {
        return Ok(source.clone());
    }
    let value = source.as_str();
    let (stem, extension) = value
        .rfind('.')
        .map_or((value, ""), |position| value.split_at(position));
    for index in 1..=u32::MAX {
        let candidate =
            PackURI::new(format!("{stem}-transfer{index}{extension}")).map_err(Error::Invalid)?;
        if !reserved.contains(&candidate) {
            return Ok(candidate);
        }
    }
    Err(invalid(
        "opened-presentation transfer part-name space is exhausted",
    ))
}

fn publish_transfer(
    source: &OpcPackage,
    destination: &mut OpcPackage,
    mapping: &HashMap<PackURI, PackURI>,
) -> Result<()> {
    let mut names: Vec<_> = mapping.keys().cloned().collect();
    names.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
    let mut staged = Vec::new();
    for source_name in names {
        let source_part = source.get_part(&source_name)?;
        let target_name = mapping
            .get(&source_name)
            .ok_or_else(|| invalid("opened-presentation transfer mapping disappeared"))?;
        let mut part = BlobPart::new_shared(
            target_name.clone(),
            source_part.content_type().to_owned(),
            source_part.blob_arc(),
        );
        for relationship in source_part.rels().iter() {
            let target = if relationship.is_external() {
                relationship.target_ref().to_owned()
            } else {
                let source_target = relationship.target_partname()?;
                let copied_target = mapping.get(&source_target).ok_or_else(|| {
                    invalid("opened-presentation transfer dependency closure is incomplete")
                })?;
                copied_target.relative_ref(target_name.base_uri())
            };
            part.rels_mut().try_add_relationship(
                relationship.reltype().to_owned(),
                target,
                relationship.r_id().to_owned(),
                if relationship.is_external() {
                    TargetMode::External
                } else {
                    TargetMode::Internal
                },
            )?;
        }
        staged.push(part);
    }
    for part in staged {
        destination.try_add_part(Box::new(part))?;
    }
    Ok(())
}

/// Validated result of one opened-presentation transaction.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Candidate snapshot after atomic publication.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Durable exact-source patch for this commit.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Whether publication changes any OPC resource.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.patch.is_empty()
    }

    /// Consume the commit into its candidate snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }

    /// Consume the commit into the durable patch.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}
