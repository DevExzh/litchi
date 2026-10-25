//! Source-backed presentation-level master and layout reads.
//!
//! The presentation and master XML lists are the semantic owners of their
//! respective inventories.  Relationship collections are only used for
//! lookup after those lists have supplied the order.  Snapshot construction
//! is split into fenced acquisition and subsequent cache publication;
//! a failed final fence removes every scene that
//! this operation actually published.

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{PackURI, Part, TargetMode};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use crate::error::{Error, Result};
use crate::notes::Conformance;
use crate::parts::{PresentationPart, SlideMasterPart};
use crate::presentation::source::{SourceBackedPresentation, SourceInner, SourcePart};
use crate::shape::owned::Scene as OwnedScene;

use super::{Layout, LayoutPending, RelatedPart, SceneCandidate};

const STRICT_SLIDE_MASTER_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/slideMaster";
const STRICT_SLIDE_LAYOUT_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/slideLayout";

const MAX_REFERENCES: usize = 4_096;
const MAX_BULK_SOURCE_BYTES: usize = 64 * 1024 * 1024;
const MAX_PRESENTATION_PART_BYTES: usize = crate::notes::MAX_PRESENTATION_XML;

/// A detached source-backed slide master.
///
/// ```
/// use std::io::Cursor;
/// use litchi_core::Position;
/// use litchi_pptx::{Package, SourceBackedPresentation};
///
/// let mut package = Package::new()?;
/// let source = SourceBackedPresentation::from_reader(Cursor::new(package.to_bytes()?))?;
/// let master = source.master(Position::new(0))?.unwrap();
/// drop(source);
/// assert_eq!(master.position(), Position::new(0));
/// # Ok::<(), litchi_pptx::Error>(())
/// ```
#[derive(Clone)]
pub struct Master {
    pub(super) data: Arc<MasterData>,
}

pub(super) struct MasterData {
    pub(super) related: RelatedPart,
    position: Position,
    name: Option<String>,
    preserved: bool,
    scene: OwnedScene,
}

impl fmt::Debug for Master {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Master").finish_non_exhaustive()
    }
}

impl Master {
    /// The zero-based presentation position of this master.
    #[must_use]
    pub fn position(&self) -> Position {
        self.data.position
    }

    /// The authored master name, if present.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.data.name.as_deref()
    }

    /// Whether the master is marked as preserved.
    #[must_use]
    pub fn is_preserved(&self) -> bool {
        self.data.preserved
    }

    /// The detached shape scene retained by this snapshot.
    #[must_use]
    pub fn scene(&self) -> &OwnedScene {
        &self.data.scene
    }
}

struct MasterPending {
    related: RelatedPart,
    position: Position,
    name: Option<String>,
    preserved: bool,
    candidate: SceneCandidate,
}

impl MasterPending {
    fn into_master(self) -> Master {
        Master {
            data: Arc::new(MasterData {
                related: self.related,
                position: self.position,
                name: self.name,
                preserved: self.preserved,
                scene: self.candidate.scene,
            }),
        }
    }
}

#[derive(Clone)]
struct Target {
    relationship_id: String,
    part_uri: PackURI,
    conformance: Conformance,
}

struct BulkBudget {
    used: usize,
}

impl BulkBudget {
    const fn new() -> Self {
        Self { used: 0 }
    }

    fn remaining(&self) -> Result<usize> {
        MAX_BULK_SOURCE_BYTES.checked_sub(self.used).ok_or_else(|| {
            Error::internal(
                "bulk_source_budget_remaining",
                "used source bytes exceed the bulk source budget",
            )
        })
    }

    fn charge(
        &mut self,
        related: &RelatedPart,
        scene: &OwnedScene,
        name: Option<&str>,
    ) -> Result<()> {
        let mut charge = related.data.as_bytes().len();
        charge = charge
            .checked_add(related.partname.as_str().len())
            .and_then(|value| value.checked_add(related.relationship_id.len()))
            .and_then(|value| value.checked_add(scene.retained_bytes()))
            .and_then(|value| value.checked_add(name.map_or(0, str::len)))
            .ok_or(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::Snapshots),
                observed: usize::MAX,
                maximum: MAX_BULK_SOURCE_BYTES,
            })?;
        let next = self.used.checked_add(charge).ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::Snapshots),
            observed: usize::MAX,
            maximum: MAX_BULK_SOURCE_BYTES,
        })?;
        if next > MAX_BULK_SOURCE_BYTES {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::Snapshots),
                observed: next,
                maximum: MAX_BULK_SOURCE_BYTES,
            });
        }
        self.used = next;
        Ok(())
    }
}

fn reserve<T>(values: &mut Vec<T>, count: usize, resource: &'static str) -> Result<()> {
    values
        .try_reserve_exact(count)
        .map_err(|source| Error::Allocation { resource, source })
}

fn presentation_part(owner: &SourceInner) -> Result<SourcePart> {
    let uri = owner._presentation.source_partname().clone();
    super::load_part(
        owner,
        &uri,
        owner._presentation.content_type(),
        MAX_PRESENTATION_PART_BYTES,
    )
}

fn root_conformance_for(part: &SourcePart, root: &str) -> Result<Conformance> {
    let mut reader = NsReader::from_reader(part.blob());
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|_| Error::Xml("PresentationML part has malformed XML".into()))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if element.name().local_name().as_ref() != root.as_bytes() {
                    return Err(Error::Invalid(
                        "PresentationML part has an invalid root".into(),
                    ));
                }
                return match namespace {
                    ResolveResult::Bound(Namespace(value))
                        if value == crate::namespace::STRICT_PRESENTATIONML_NAMESPACE =>
                    {
                        Ok(Conformance::Strict)
                    },
                    ResolveResult::Bound(Namespace(value))
                        if value == crate::namespace::PRESENTATIONML_NAMESPACE =>
                    {
                        Ok(Conformance::Transitional)
                    },
                    _ => Err(Error::Invalid(
                        "PresentationML part has an invalid root namespace".into(),
                    )),
                };
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Text(value) if value.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            Event::Eof => {
                return Err(Error::Invalid(
                    "PresentationML part lacks an element root".into(),
                ));
            },
            _ => {
                return Err(Error::Invalid(
                    "PresentationML part has invalid root content".into(),
                ));
            },
        }
    }
}

fn relationship_type_for(part: &SourcePart, root: &str) -> Result<(&'static str, Conformance)> {
    let conformance = root_conformance_for(part, root)?;
    let relationship_type = match conformance {
        Conformance::Strict => {
            if root == "presentation" {
                STRICT_SLIDE_MASTER_REL
            } else {
                STRICT_SLIDE_LAYOUT_REL
            }
        },
        Conformance::Transitional => {
            if root == "presentation" {
                rt::SLIDE_MASTER
            } else {
                rt::SLIDE_LAYOUT
            }
        },
    };
    Ok((relationship_type, conformance))
}

fn require_conformance(part: &SourcePart, root: &str, expected: Conformance) -> Result<()> {
    if root_conformance_for(part, root)? != expected {
        return Err(Error::Relationship(
            "PresentationML relationship closure has mixed conformance".into(),
        ));
    }
    Ok(())
}

fn resolve_target(
    part: &SourcePart,
    relationship_id: &str,
    expected_type: &str,
    conformance: Conformance,
    missing_message: &'static str,
    external_message: &'static str,
    type_message: &'static str,
) -> Result<Target> {
    let relationship = part
        .source_relationships()
        .get(relationship_id)
        .ok_or_else(|| Error::Relationship(missing_message.into()))?;
    if relationship.target_mode() == TargetMode::External {
        return Err(Error::Relationship(external_message.into()));
    }
    if relationship.reltype() != expected_type {
        return Err(Error::Relationship(type_message.into()));
    }
    Ok(Target {
        relationship_id: relationship_id.to_owned(),
        part_uri: relationship.target_partname()?,
        conformance,
    })
}

fn master_targets(owner: &SourceInner, presentation: &SourcePart) -> Result<Vec<Target>> {
    let presentation_part = PresentationPart::from_part(presentation)?;
    let relationship_ids = presentation_part.slide_master_references()?;
    if relationship_ids.len() > MAX_REFERENCES {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::References),
            observed: relationship_ids.len(),
            maximum: MAX_REFERENCES,
        });
    }
    let (expected_type, conformance) = relationship_type_for(presentation, "presentation")?;
    let mut targets = Vec::new();
    reserve(
        &mut targets,
        relationship_ids.len(),
        "source presentation master references",
    )?;
    for relationship_id in relationship_ids {
        owner.package.check_execution()?;
        targets.push(resolve_target(
            presentation,
            &relationship_id,
            expected_type,
            conformance,
            "presentation master relationship is missing",
            "presentation master relationship must be internal",
            "presentation master relationship has an unexpected type",
        )?);
    }
    Ok(targets)
}

fn layout_targets(owner: &SourceInner, master: &SourcePart) -> Result<Vec<Target>> {
    let master_part = SlideMasterPart::from_part(master)?;
    let relationship_ids = master_part.layout_references()?;
    if relationship_ids.len() > MAX_REFERENCES {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::References),
            observed: relationship_ids.len(),
            maximum: MAX_REFERENCES,
        });
    }
    let (expected_type, conformance) = relationship_type_for(master, "sldMaster")?;
    let mut targets = Vec::new();
    reserve(
        &mut targets,
        relationship_ids.len(),
        "source master layout references",
    )?;
    for relationship_id in relationship_ids {
        owner.package.check_execution()?;
        targets.push(resolve_target(
            master,
            &relationship_id,
            expected_type,
            conformance,
            "slide master layout relationship is missing",
            "slide master layout relationship must be internal",
            "slide master layout relationship has an unexpected type",
        )?);
    }
    Ok(targets)
}

fn preflight_bulk_targets(owner: &SourceInner, targets: &[Target]) -> Result<()> {
    let mut total = 0u64;
    for target in targets {
        owner.package.check_execution()?;
        let view = owner.package.part(&target.part_uri)?;
        total = total
            .checked_add(view.declared_uncompressed_size()?)
            .ok_or(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::Snapshots),
                observed: usize::MAX,
                maximum: MAX_BULK_SOURCE_BYTES,
            })?;
        if total > MAX_BULK_SOURCE_BYTES as u64 {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::Snapshots),
                observed: usize::try_from(total).unwrap_or(usize::MAX),
                maximum: MAX_BULK_SOURCE_BYTES,
            });
        }
    }
    Ok(())
}

fn load_bulk_part(
    owner: &SourceInner,
    target: &Target,
    content_type: &str,
    budget: &BulkBudget,
) -> Result<SourcePart> {
    let remaining = budget.remaining()?;
    let semantic_maximum = remaining.min(
        crate::open_options::selected(owner.semantic_limits.as_ref())
            .shape()
            .input_bytes(),
    );
    if semantic_maximum == 0 {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::Snapshots),
            observed: 0,
            maximum: MAX_BULK_SOURCE_BYTES,
        });
    }
    super::load_part(owner, &target.part_uri, content_type, semantic_maximum)
}

fn shape_input_limit(owner: &SourceInner) -> usize {
    crate::open_options::selected(owner.semantic_limits.as_ref())
        .shape()
        .input_bytes()
}

fn prepare_master(
    owner: &SourceInner,
    part: SourcePart,
    relationship_id: &str,
    position: Position,
    conformance: Conformance,
) -> Result<MasterPending> {
    require_conformance(&part, "sldMaster", conformance)?;
    let master = SlideMasterPart::from_part(&part)?;
    let name = master.authored_name()?;
    let preserved = master.is_preserved()?;
    let candidate = super::acquire_scene(owner, &part, |source| {
        let master = SlideMasterPart::from_part(source)?;
        let scene = OwnedScene::from_source(
            master.shapes_with_limits(
                crate::open_options::selected(owner.semantic_limits.as_ref()).shape(),
            )?,
            source.source_data().clone(),
        )?;
        super::strict_scene(&scene, "master")?;
        Ok(scene)
    })?;
    Ok(MasterPending {
        related: RelatedPart::new(part, relationship_id),
        position,
        name,
        preserved,
        candidate,
    })
}

fn finish_one_master(owner: &SourceInner, mut pending: MasterPending) -> Result<Master> {
    let key = pending.candidate.key.clone();
    let built = pending.candidate.scene.clone();
    let needs_publish = pending.candidate.needs_publish;
    let scene = if needs_publish {
        owner.scenes.publish(&key, built.clone())
    } else {
        built
    };
    let published = needs_publish && OwnedScene::ptr_eq(&scene, &pending.candidate.scene);
    pending.candidate.scene = scene;
    let result = pending.into_master();
    if let Err(error) = super::fence(owner) {
        if published {
            owner.scenes.discard(&key, result.scene());
        }
        return Err(error);
    }
    Ok(result)
}

fn finish_one_layout(owner: &SourceInner, mut pending: LayoutPending) -> Result<Layout> {
    let key = pending.candidate.key.clone();
    let built = pending.candidate.scene.clone();
    let needs_publish = pending.candidate.needs_publish;
    let scene = if needs_publish {
        owner.scenes.publish(&key, built.clone())
    } else {
        built
    };
    let published = needs_publish && OwnedScene::ptr_eq(&scene, &pending.candidate.scene);
    pending.candidate.scene = scene;
    let result = pending.into_layout();
    if let Err(error) = super::fence(owner) {
        if published {
            owner.scenes.discard(&key, result.scene());
        }
        return Err(error);
    }
    Ok(result)
}

fn finish_many_masters(
    owner: &SourceInner,
    mut pending: Vec<MasterPending>,
) -> Result<Vec<Master>> {
    let mut result = Vec::new();
    reserve(&mut result, pending.len(), "source presentation masters")?;
    let mut published = Vec::new();
    reserve(
        &mut published,
        pending.len(),
        "source presentation master cache publications",
    )?;
    for mut item in pending.drain(..) {
        let key = item.candidate.key.clone();
        let built = item.candidate.scene.clone();
        let needs_publish = item.candidate.needs_publish;
        let scene = if needs_publish {
            owner.scenes.publish(&key, built.clone())
        } else {
            built
        };
        if needs_publish && OwnedScene::ptr_eq(&scene, &item.candidate.scene) {
            published.push((key, scene.clone()));
        }
        item.candidate.scene = scene;
        result.push(item.into_master());
    }
    if let Err(error) = super::fence(owner) {
        for (key, scene) in published {
            owner.scenes.discard(&key, &scene);
        }
        return Err(error);
    }
    Ok(result)
}

fn finish_many_layouts(
    owner: &SourceInner,
    mut pending: Vec<LayoutPending>,
) -> Result<Vec<Layout>> {
    let mut result = Vec::new();
    reserve(&mut result, pending.len(), "source master layouts")?;
    let mut published = Vec::new();
    reserve(
        &mut published,
        pending.len(),
        "source master layout cache publications",
    )?;
    for mut item in pending.drain(..) {
        let key = item.candidate.key.clone();
        let built = item.candidate.scene.clone();
        let needs_publish = item.candidate.needs_publish;
        let scene = if needs_publish {
            owner.scenes.publish(&key, built.clone())
        } else {
            built
        };
        if needs_publish && OwnedScene::ptr_eq(&scene, &item.candidate.scene) {
            published.push((key, scene.clone()));
        }
        item.candidate.scene = scene;
        result.push(item.into_layout());
    }
    if let Err(error) = super::fence(owner) {
        for (key, scene) in published {
            owner.scenes.discard(&key, &scene);
        }
        return Err(error);
    }
    Ok(result)
}

impl SourceBackedPresentation {
    /// Acquire all slide-master snapshots in presentation XML order.
    ///
    /// # Errors
    /// Returns typed source, execution, limit, allocation, relationship or XML errors.
    pub fn masters(&self) -> Result<Vec<Master>> {
        let pending = super::with_fences(&self.inner, || {
            let presentation = presentation_part(&self.inner)?;
            let targets = master_targets(&self.inner, &presentation)?;
            preflight_bulk_targets(&self.inner, &targets)?;
            let mut pending = Vec::new();
            reserve(&mut pending, targets.len(), "source presentation masters")?;
            let mut budget = BulkBudget::new();
            for (index, target) in targets.iter().enumerate() {
                let part = load_bulk_part(&self.inner, target, ct::PML_SLIDE_MASTER, &budget)?;
                let item = prepare_master(
                    &self.inner,
                    part,
                    &target.relationship_id,
                    Position::new(index),
                    target.conformance,
                )?;
                budget.charge(&item.related, &item.candidate.scene, item.name.as_deref())?;
                pending.push(item);
            }
            Ok(pending)
        })?;
        finish_many_masters(&self.inner, pending)
    }

    /// The presentation theme: the presentation part's own theme, else the
    /// first slide master's (record 0321). `None` when there is neither.
    /// Only the presentation part, at most one master part and the theme
    /// part are read.
    ///
    /// # Errors
    /// Returns typed source, execution, limit, relationship or XML errors.
    pub fn theme(&self) -> Result<Option<litchi_drawingml::theme::Theme>> {
        super::with_fences(&self.inner, || {
            let presentation = presentation_part(&self.inner)?;
            let target = match crate::shape::theme::package::theme_target(&presentation)? {
                Some(target) => Some(target),
                None => match master_targets(&self.inner, &presentation)?.first() {
                    Some(master) => {
                        let part = super::load_part(
                            &self.inner,
                            &master.part_uri,
                            ct::PML_SLIDE_MASTER,
                            shape_input_limit(&self.inner),
                        )?;
                        crate::shape::theme::package::theme_target(&part)?
                    },
                    None => None,
                },
            };
            let Some(target) = target else {
                return Ok(None);
            };
            let part = super::load_part(
                &self.inner,
                &target,
                ct::OFC_THEME,
                shape_input_limit(&self.inner),
            )?;
            Ok(Some(litchi_drawingml::theme::codec::read(
                part.blob(),
                litchi_drawingml::theme::codec::Limits::DEFAULT,
            )?))
        })
    }

    /// Acquire one slide-master snapshot by its zero-based presentation position.
    ///
    /// # Errors
    /// Returns typed source, execution, limit, allocation, relationship or XML errors.
    pub fn master(&self, position: Position) -> Result<Option<Master>> {
        let pending = super::with_fences(&self.inner, || {
            let presentation = presentation_part(&self.inner)?;
            let targets = master_targets(&self.inner, &presentation)?;
            let Some(target) = targets.get(position.get()) else {
                return Ok(None);
            };
            let part = super::load_part(
                &self.inner,
                &target.part_uri,
                ct::PML_SLIDE_MASTER,
                shape_input_limit(&self.inner),
            )?;
            prepare_master(
                &self.inner,
                part,
                &target.relationship_id,
                position,
                target.conformance,
            )
            .map(Some)
        })?;
        pending
            .map(|pending| finish_one_master(&self.inner, pending))
            .transpose()
    }

    /// Acquire all layouts owned by one master, in that master XML's order.
    /// Returns `None` for an absent master and `Some(empty)` for no layouts.
    ///
    /// # Errors
    /// Returns typed source, execution, limit, allocation, relationship or XML errors.
    pub fn layouts(&self, master_position: Position) -> Result<Option<Vec<Layout>>> {
        let pending = super::with_fences(&self.inner, || {
            let presentation = presentation_part(&self.inner)?;
            let masters = master_targets(&self.inner, &presentation)?;
            let Some(master_target) = masters.get(master_position.get()) else {
                return Ok(None);
            };
            let master = super::load_part(
                &self.inner,
                &master_target.part_uri,
                ct::PML_SLIDE_MASTER,
                shape_input_limit(&self.inner),
            )?;
            require_conformance(&master, "sldMaster", master_target.conformance)?;
            let targets = layout_targets(&self.inner, &master)?;
            preflight_bulk_targets(&self.inner, &targets)?;
            let mut pending = Vec::new();
            reserve(&mut pending, targets.len(), "source master layouts")?;
            let mut budget = BulkBudget::new();
            for target in &targets {
                let part = load_bulk_part(&self.inner, target, ct::PML_SLIDE_LAYOUT, &budget)?;
                require_conformance(&part, "sldLayout", target.conformance)?;
                let item = super::prepare_layout(&self.inner, part, &target.relationship_id)?;
                budget.charge(&item.related, &item.candidate.scene, item.name.as_deref())?;
                pending.push(item);
            }
            Ok(Some(pending))
        })?;
        pending
            .map(|pending| finish_many_layouts(&self.inner, pending))
            .transpose()
    }

    /// Acquire one layout owned by one master using zero-based positions.
    ///
    /// # Errors
    /// Returns typed source, execution, limit, allocation, relationship or XML errors.
    pub fn layout(
        &self,
        master_position: Position,
        layout_position: Position,
    ) -> Result<Option<Layout>> {
        let pending = super::with_fences(&self.inner, || {
            let presentation = presentation_part(&self.inner)?;
            let masters = master_targets(&self.inner, &presentation)?;
            let Some(master_target) = masters.get(master_position.get()) else {
                return Ok(None);
            };
            let master = super::load_part(
                &self.inner,
                &master_target.part_uri,
                ct::PML_SLIDE_MASTER,
                shape_input_limit(&self.inner),
            )?;
            require_conformance(&master, "sldMaster", master_target.conformance)?;
            let targets = layout_targets(&self.inner, &master)?;
            let Some(target) = targets.get(layout_position.get()) else {
                return Ok(None);
            };
            let part = super::load_part(
                &self.inner,
                &target.part_uri,
                ct::PML_SLIDE_LAYOUT,
                shape_input_limit(&self.inner),
            )?;
            require_conformance(&part, "sldLayout", target.conformance)?;
            super::prepare_layout(&self.inner, part, &target.relationship_id).map(Some)
        })?;
        pending
            .map(|pending| finish_one_layout(&self.inner, pending))
            .transpose()
    }
}
