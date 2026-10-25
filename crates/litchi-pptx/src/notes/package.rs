//! Transactional OPC graph ownership for `PresentationML` notes.

use super::codec::{
    root_conformance_with_limits, scan_xml_with_limits, validate_resource_xml_with_limits,
};
use super::model::{Conformance, Graph, Link, Master, Slide, Theme};
use super::transaction::{Change, PartState, Snapshot};
use super::{
    Limits, MAX_MASTER_XML, MAX_NOTES_SLIDES, MAX_NOTES_XML, MAX_OWNED_PARTS, MAX_PRESENTATION_XML,
    MAX_SLIDE_XML, MAX_THEME_XML, SLIDE_CT, THEME_CT, allocation, checked_add, invalid, limit,
    own_blob,
};
use crate::package::replay_index::ReplayIndex;
use crate::package::signature::{refuse_signed_change, settle_unsigned_change};
use crate::{Error, Result};
#[cfg(test)]
use litchi_opc::TargetMode;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
#[cfg(test)]
use litchi_opc::part::BlobPart;
use litchi_opc::part::Part;
use litchi_opc::{OpcError, OpcPackage, PackURI};
#[cfg(test)]
use std::collections::BTreeMap;
use std::collections::{BTreeSet, HashMap, HashSet};
use std::sync::Arc;

#[derive(Clone, Copy)]
enum Disposition {
    Retain,
    Follow,
    External,
}

#[derive(Debug)]
struct ThemeIndex {
    relationship_id: String,
    part_name: PackURI,
    content_type: String,
}

#[derive(Debug)]
struct MasterIndex {
    presentation_relationship_id: String,
    part_name: PackURI,
    content_type: String,
    theme: ThemeIndex,
}

#[derive(Debug)]
struct SlideIndex {
    slide_part_name: PackURI,
    slide_relationship_id: String,
    part_name: PackURI,
    content_type: String,
    backlink_relationship_id: String,
    notes_master_relationship_id: String,
}

#[derive(Debug)]
struct GraphIndex {
    conformance: Conformance,
    master: MasterIndex,
    slides: Vec<SlideIndex>,
}

/// Capture-local classification of one slide root. The raw slice is kept as
/// a witness for the exact source passed through MCE; processed XML is not
/// retained after the capture projection returns.
#[derive(Clone, Copy)]
pub(crate) struct SlideRootProof<'a> {
    raw: &'a [u8],
    conformance: Option<Conformance>,
    source_checked: bool,
}

impl<'a> SlideRootProof<'a> {
    #[cfg(test)]
    pub(crate) fn new(raw: &'a [u8], conformance: Option<Conformance>) -> Self {
        Self {
            raw,
            conformance,
            source_checked: false,
        }
    }

    pub(crate) fn with_source_checked(
        raw: &'a [u8],
        conformance: Option<Conformance>,
        source_checked: bool,
    ) -> Self {
        Self {
            raw,
            conformance,
            source_checked,
        }
    }

    pub(crate) fn is_valid(self) -> bool {
        self.conformance.is_some()
    }
}

/// Allocation key of one raw slide observation: its address and length.
pub(crate) type RawKey = (usize, usize);

fn raw_key(raw: &[u8]) -> RawKey {
    (raw.as_ptr() as usize, raw.len())
}

/// Resource named by the typed error a refused memo reservation raises.
const SLIDE_ROOT_MEMO: &str = "opened-presentation slide-root memo";

/// One retained classification: the allocation it classified, held strongly
/// so its address cannot be reused while the entry lives, and the result.
struct SlideRootEntry {
    key: RawKey,
    raw: Arc<Vec<u8>>,
    conformance: Conformance,
}

/// Notes-root classifications proved by one opened-presentation capture,
/// keyed by the payload allocation each one read (ADR 0032).
///
/// An entry asserts *the slide payload at address `a` of length `l`,
/// processed under the default markup-compatibility profile, classifies as
/// `c` under the complete notes-graph root scan*. That classification is a
/// deterministic, total function of the payload bytes alone: the scan's
/// profile, root name and ceilings are compile-time constants, and it reads no
/// part name, content type, relationship, other part or caller limit. Only a
/// classification the scan returned is kept; a slide the scan refuses is
/// never memoized, so its refusal is always recomputed from the bytes.
///
/// The entry retains the payload `Arc`, so the address cannot be recycled by
/// a different payload while the entry lives, and a hit on the exact raw
/// observation a later capture makes therefore proves byte identity. A miss is
/// an ordinary scan, so no value, refusal or published byte depends on which
/// entries the memo holds; a hit is re-derived by a `debug_assert!` in test and
/// debug builds.
///
/// Entries are admitted only for allocations the owning snapshot's package
/// holds, so the memo pins no payload that package does not already own. It is
/// built by one capture or projected by one rebind and never mutated after; it
/// is deliberately not `Clone`, so a memo can reach another snapshot only
/// through [`Self::project`], never by copy. Its resident cost is one 32-byte
/// entry (key, `Arc`, classification) per captured slide, bounded by the
/// capture's slide limit, reserved fallibly.
#[derive(Default)]
pub(crate) struct SlideRootMemo {
    entries: Vec<SlideRootEntry>,
}

impl SlideRootMemo {
    /// Admit every successful classification in `proofs` whose allocation
    /// `owner` proves the owning snapshot holds.
    ///
    /// Each proof borrows the very payload slice the capture classified, so
    /// the compiler proves every classified allocation is alive — and its
    /// address therefore not reusable by other bytes — while its entry is
    /// admitted; an entry is kept only when `owner` hands over that same
    /// allocation. The table is reserved once, fallibly, for one entry per
    /// proof, and a refused reservation is the typed [`Error::Allocation`] the
    /// capture reports for its other reservations, never an abort and never a
    /// silently empty memo (ADR 0032 section 3).
    pub(crate) fn from_records(
        proofs: &[SlideRootProof<'_>],
        owner: impl Fn(RawKey) -> Option<Arc<Vec<u8>>>,
    ) -> Result<Self> {
        let mut entries = Vec::new();
        reserve_entries(&mut entries, proofs.len())
            .map_err(|source| allocation(SLIDE_ROOT_MEMO, source))?;
        for proof in proofs {
            // A slide the scan refused is recomputed from its bytes every time.
            let Some(conformance) = proof.conformance else {
                continue;
            };
            let key = raw_key(proof.raw);
            let Some(raw) = owner(key).filter(|raw| std::ptr::eq(raw.as_slice(), proof.raw)) else {
                continue;
            };
            // At most one entry per proof, all reserved above: no push grows
            // the table.
            debug_assert!(entries.len() < entries.capacity());
            entries.push(SlideRootEntry {
                key,
                raw,
                conformance,
            });
        }
        Ok(Self::sorted(entries))
    }

    /// Keep only the entries whose allocation `owner` proves a rebound
    /// snapshot's own package holds, each retaining the `Arc` `owner` returns.
    ///
    /// A rebind is infallible, so a table that cannot be reserved is an empty
    /// memo, exactly as the part-digest projection beside it degrades (ADR 0032
    /// section 3); an empty memo is only a miss.
    pub(crate) fn project(&self, owner: impl Fn(RawKey) -> Option<Arc<Vec<u8>>>) -> Self {
        let mut entries = Vec::new();
        if reserve_entries(&mut entries, self.entries.len()).is_err() {
            return Self::default();
        }
        for entry in &self.entries {
            // The entry keeps its allocation alive, so an `Arc` with its
            // address and length is that very allocation.
            if let Some(raw) =
                owner(entry.key).filter(|raw| std::ptr::eq(raw.as_slice(), entry.raw.as_slice()))
            {
                entries.push(SlideRootEntry {
                    key: entry.key,
                    raw,
                    conformance: entry.conformance,
                });
            }
        }
        // `self.entries` is sorted and unique by key, and so is its subsequence.
        Self { entries }
    }

    fn sorted(mut entries: Vec<SlideRootEntry>) -> Self {
        // Neither the unstable sort nor the deduplication allocates.
        entries.sort_unstable_by_key(|entry| entry.key);
        // Two slides may share one payload allocation; its bytes, and so its
        // classification, are the same for both.
        debug_assert!(
            entries.windows(2).all(
                |pair| pair[0].key != pair[1].key || pair[0].conformance == pair[1].conformance
            ),
            "one slide payload allocation was classified two different ways"
        );
        entries.dedup_by_key(|entry| entry.key);
        Self { entries }
    }

    /// Classification proved for exactly this raw observation, if any.
    pub(crate) fn lookup(&self, raw: &[u8]) -> Option<Conformance> {
        let index = self
            .entries
            .binary_search_by_key(&raw_key(raw), |entry| entry.key)
            .ok()?;
        let entry = self.entries.get(index)?;
        std::ptr::eq(entry.raw.as_slice(), raw).then_some(entry.conformance)
    }

    /// Number of retained classifications.
    #[cfg(test)]
    pub(crate) fn len(&self) -> usize {
        self.entries.len()
    }

    /// Each entry's key, the payload allocation it retains and its
    /// classification, for the tests that prove what an entry names.
    #[cfg(test)]
    pub(crate) fn retained(
        &self,
    ) -> impl Iterator<Item = (RawKey, &Arc<Vec<u8>>, Conformance)> + '_ {
        self.entries
            .iter()
            .map(|entry| (entry.key, &entry.raw, entry.conformance))
    }

    /// Bytes the table occupies itself, excluding the payloads it names and
    /// the `Arc` header a snapshot keeps it behind.
    #[cfg(test)]
    pub(crate) fn resident_bytes(&self) -> usize {
        self.entries.capacity() * size_of::<SlideRootEntry>()
    }

    /// Size of one entry, stated by the implementing record.
    #[cfg(test)]
    pub(crate) const ENTRY_BYTES: usize = size_of::<SlideRootEntry>();
}

#[cfg(test)]
thread_local! {
    /// Whether this thread's memo reservations are refused, for the tests that
    /// prove a refused reservation is a typed error rather than an abort or a
    /// silently empty memo.
    static REFUSE_SLIDE_ROOT_RESERVATION: std::cell::Cell<bool> =
        const { std::cell::Cell::new(false) };
}

/// Reserve room for `additional` more entries, fallibly.
fn reserve_entries(
    entries: &mut Vec<SlideRootEntry>,
    additional: usize,
) -> std::result::Result<(), std::collections::TryReserveError> {
    #[cfg(test)]
    if REFUSE_SLIDE_ROOT_RESERVATION.with(std::cell::Cell::get) {
        // No allocator can satisfy this request, so the error is the genuine
        // one a refused reservation returns.
        return entries.try_reserve_exact(usize::MAX);
    }
    entries.try_reserve_exact(additional)
}

/// Run `operation` with every slide-root memo reservation on this thread
/// refused, restoring the previous setting afterwards, even on a panic.
#[cfg(test)]
pub(crate) fn with_refused_slide_root_reservation<T>(operation: impl FnOnce() -> T) -> T {
    struct Restore(bool);
    impl Drop for Restore {
        fn drop(&mut self) {
            REFUSE_SLIDE_ROOT_RESERVATION.with(|refuse| refuse.set(self.0));
        }
    }
    let _restore = Restore(REFUSE_SLIDE_ROOT_RESERVATION.with(|refuse| refuse.replace(true)));
    operation()
}

/// Load and validate the complete bounded notes graph for a presentation part.
///
/// The returned graph is lifetime-free and independently editable, so each
/// validated notes, master, and theme payload is copied exactly once. Package
/// deletion uses the metadata-only index and does not perform these copies.
#[cfg(test)]
pub(crate) fn load(package: &OpcPackage, presentation_name: &PackURI) -> Result<Option<Graph>> {
    load_with_limits(
        package,
        presentation_name,
        Limits::from_source(package.read_limits()),
    )
}

pub(crate) fn load_with_limits(
    package: &OpcPackage,
    presentation_name: &PackURI,
    limits: Limits,
) -> Result<Option<Graph>> {
    let Some(index) = load_index_with_limits(package, presentation_name, limits)? else {
        return Ok(None);
    };
    Ok(Some(materialize(package, index)?))
}

/// Validate one changed notes payload under the recipient's selected policy.
/// Patch replay calls this only for payload deltas, avoiding a second complete
/// notes-graph scan when the patch changes one XML member.
pub(crate) fn validate_replay_part(part: &dyn Part, limits: Limits) -> Result<()> {
    let (max, root, label) = match part.content_type() {
        ct::PML_NOTES_MASTER => (MAX_MASTER_XML, "notesMaster", "notes master"),
        ct::PML_NOTES_SLIDE => (MAX_NOTES_XML, "notes", "notes slide"),
        THEME_CT => (MAX_THEME_XML, "theme", "notes-master theme"),
        _ => return Ok(()),
    };
    let conformance = root_conformance_with_limits(part.blob(), max, root, limits)?;
    validate_resource_xml_with_limits(part.blob(), max, conformance, root, label, limits)
}

/// Check the recipient's notes-graph aggregate after a payload replay.
///
/// The aggregate ceiling is read from the retained ingress index, whose
/// notes-master/theme edges and slide contributions are updated by the exact
/// replay write set.
pub(crate) fn validate_replay_total(index: &ReplayIndex, limits: Limits) -> Result<()> {
    if let Some(theme) = index.missing_theme_name()? {
        return Err(Error::Opc(OpcError::PartNotFound(theme.to_string())));
    }
    let total = usize::try_from(index.notes_total_bytes())
        .map_err(|_| invalid("notes aggregate bytes overflow"))?;
    if total > limits.max_total_bytes {
        return Err(limit(
            crate::error::CodecLimitKind::XmlBytes,
            total,
            limits.max_total_bytes,
        ));
    }
    Ok(())
}

/// Capture a complete source-checked snapshot of an existing notes graph.
#[cfg(test)]
pub(crate) fn load_snapshot(
    package: &OpcPackage,
    presentation_name: &PackURI,
) -> Result<Option<Snapshot>> {
    load_snapshot_with_limits(
        package,
        presentation_name,
        Limits::from_source(package.read_limits()),
    )
}

pub(crate) fn load_snapshot_with_limits(
    package: &OpcPackage,
    presentation_name: &PackURI,
    limits: Limits,
) -> Result<Option<Snapshot>> {
    snapshot_from_index(
        package,
        presentation_name,
        load_index_with_limits(package, presentation_name, limits)?,
        limits,
    )
}

pub(crate) fn load_snapshot_with_slide_root_proofs_and_limits<'a>(
    package: &OpcPackage,
    presentation_name: &PackURI,
    expected_catalog_len: usize,
    proofs: &[SlideRootProof<'a>],
    limits: Limits,
    proof_limits: Limits,
) -> Result<Option<Snapshot>> {
    let index = load_index_with_slide_root_proofs_and_limits(
        package,
        presentation_name,
        Some((expected_catalog_len, proofs)),
        limits,
        proof_limits,
    )?;
    snapshot_from_index(package, presentation_name, index, limits)
}

fn snapshot_from_index(
    package: &OpcPackage,
    presentation_name: &PackURI,
    index: Option<GraphIndex>,
    limits: Limits,
) -> Result<Option<Snapshot>> {
    let Some(index) = index else {
        return Ok(None);
    };
    let graph = materialize(package, index)?;
    let presentation = package.get_part(presentation_name)?;
    let presentation_part = PartState::from_part(presentation);
    let mut parts = vec![presentation_part.clone()];
    parts.push(PartState::from_part(package.get_part(
        &PackURI::new(graph.master().part()).map_err(Error::Invalid)?,
    )?));
    parts.push(PartState::from_part(package.get_part(
        &PackURI::new(graph.master().theme().part()).map_err(Error::Invalid)?,
    )?));
    for slide in graph.slides() {
        parts.push(PartState::from_part(
            package.get_part(&PackURI::new(slide.owner()).map_err(Error::Invalid)?)?,
        ));
        parts.push(PartState::from_part(
            package.get_part(&PackURI::new(slide.part()).map_err(Error::Invalid)?)?,
        ));
    }
    Snapshot::from_parts(
        presentation.partname().clone(),
        presentation_part,
        parts,
        graph,
        limits,
    )
    .map(Some)
}

/// Apply a source-checked notes patch atomically.
pub(crate) fn apply_patch(package: &mut OpcPackage, patch: &Change) -> Result<Snapshot> {
    let limits = patch.before().limits;
    let current =
        load_snapshot_with_limits(package, patch.before().presentation_part_name(), limits)?
            .ok_or_else(|| invalid("notes patch source graph is absent"))?;
    if !current.same_source(patch.before()) {
        return Err(invalid("notes patch source is stale"));
    }
    if patch.is_noop() {
        return Ok(current);
    }
    refuse_signed_change(package)?;
    let mut candidate = package.clone();
    for part in &patch.after().parts {
        let target = candidate.get_part_mut(&part.name)?;
        target.set_content_type(part.content_type.clone())?;
        target.set_blob_shared(Arc::clone(&part.data));
        let ids: Vec<_> = target
            .rels()
            .iter()
            .map(|relationship| relationship.r_id().to_owned())
            .collect();
        for id in ids {
            target.rels_mut().remove(&id);
        }
        for link in &part.relationships {
            target.rels_mut().try_add_relationship(
                link.relationship_type.clone(),
                link.target_ref.clone(),
                link.id.clone(),
                link.target_mode,
            )?;
        }
    }
    let resulting = load_snapshot_with_limits(
        &candidate,
        patch.after().presentation_part_name(),
        patch.after().limits,
    )?
    .ok_or_else(|| invalid("published notes graph is absent"))?;
    if !resulting.same_source(patch.after()) {
        return Err(invalid("published notes graph differs from the commit"));
    }
    settle_unsigned_change(&mut candidate)?;
    *package = candidate;
    Ok(resulting)
}

/// Apply a committed notes edit atomically.
pub(crate) fn apply_commit(
    package: &mut OpcPackage,
    commit: super::transaction::Prepared,
) -> Result<Snapshot> {
    let patch = commit.into_patch();
    apply_patch(package, &patch)
}

fn load_index_with_limits(
    package: &OpcPackage,
    presentation_name: &PackURI,
    limits: Limits,
) -> Result<Option<GraphIndex>> {
    load_index_with_slide_root_proofs_and_limits(package, presentation_name, None, limits, limits)
}

fn load_index_with_slide_root_proofs_and_limits<'a>(
    package: &OpcPackage,
    presentation_name: &PackURI,
    slide_root_proofs: Option<(usize, &[SlideRootProof<'a>])>,
    limits: Limits,
    proof_limits: Limits,
) -> Result<Option<GraphIndex>> {
    let presentation = package.get_part(presentation_name)?;
    if !is_presentation_main_content_type(presentation.content_type()) {
        return Err(invalid("notes graph requires a PresentationML main part"));
    }
    let conformance = root_conformance_with_limits(
        presentation.blob(),
        MAX_PRESENTATION_XML,
        "presentation",
        limits,
    )?;
    let presentation_scan = scan_xml_with_limits(
        presentation.blob(),
        MAX_PRESENTATION_XML,
        conformance,
        "presentation",
        limits,
    )?;
    if presentation_scan.notes_master_ids.len() > 1 {
        return Err(invalid(
            "presentation has multiple notesMasterId references",
        ));
    }
    if presentation_scan.slide_ids.len() > MAX_NOTES_SLIDES {
        return Err(limit(
            crate::error::CodecLimitKind::Slides,
            presentation_scan.slide_ids.len(),
            MAX_NOTES_SLIDES,
        ));
    }
    let proof_inventory_matches = slide_root_proofs
        .is_some_and(|(expected_len, _)| expected_len == presentation_scan.slide_ids.len());
    let mut slide_sources = Vec::new();
    slide_sources
        .try_reserve_exact(presentation_scan.slide_ids.len())
        .map_err(|source| allocation("notes source slide inventory", source))?;
    for (slide_index, id) in presentation_scan.slide_ids.iter().enumerate() {
        validate_id(id)?;
        let relationship = presentation
            .rels()
            .get(id)
            .ok_or_else(|| invalid("presentation slide reference is missing its relationship"))?;
        if relationship.reltype() != conformance.slide_rel() || relationship.is_external() {
            return Err(invalid(
                "presentation slide relationship has wrong type or target mode",
            ));
        }
        let target = relationship_target(presentation, relationship)?;
        let slide = package.get_part(&target)?;
        if slide.content_type() != SLIDE_CT {
            return Err(invalid("slide has invalid content type"));
        }
        let actual = if proof_inventory_matches {
            let current = slide.blob();
            let proof = slide_root_proofs.and_then(|(_, proofs)| proofs.get(slide_index));
            if !source_proof_matches_current(proof, current) {
                crate::parts::validate_source_part_xml(current, package.read_limits())?;
            }
            resolve_slide_root_with_profiles(current, proof, proof_limits, limits)
        } else {
            crate::parts::validate_source_part_xml(slide.blob(), package.read_limits())?;
            root_conformance_with_limits(slide.blob(), MAX_SLIDE_XML, "sld", limits)
        }?;
        if actual != conformance {
            return Err(invalid("slide conformance differs from presentation"));
        }
        slide_sources.push((id.clone(), slide.partname().clone()));
    }
    let master_relationships: Vec<_> = presentation
        .rels()
        .iter()
        .filter(|relationship| is_notes_master_rel(relationship.reltype()))
        .collect();
    let Some(master_id) = presentation_scan.notes_master_ids.first() else {
        if !master_relationships.is_empty()
            || package.iter_parts().any(|part| {
                matches!(
                    part.content_type(),
                    ct::PML_NOTES_MASTER | ct::PML_NOTES_SLIDE
                )
            })
            || slide_sources.iter().any(|(_, name)| {
                package.get_part(name).is_ok_and(|part| {
                    part.rels()
                        .iter()
                        .any(|relationship| is_notes_slide_rel(relationship.reltype()))
                })
            })
        {
            return Err(invalid("orphan notes graph exists without notesMasterId"));
        }
        return Ok(None);
    };
    validate_id(master_id)?;
    if master_relationships.len() != 1 {
        return Err(invalid(
            "notesMasterId and presentation notes-master relationships differ",
        ));
    }
    let master_relationship = presentation
        .rels()
        .get(master_id)
        .ok_or_else(|| invalid("notesMasterId relationship is missing"))?;
    if master_relationship.reltype() != conformance.notes_master_rel()
        || master_relationship.is_external()
    {
        return Err(invalid(
            "notes-master relationship has wrong type or target mode",
        ));
    }
    let master_name = relationship_target(presentation, master_relationship)?;
    validate_leaf_path(&master_name, "/ppt/notesMasters/", "notes master")?;
    let master_part = package.get_part(&master_name)?;
    let master_part_name = master_part.partname().clone();
    require_content_type(master_part, ct::PML_NOTES_MASTER, "notes master")?;
    validate_resource_xml_with_limits(
        master_part.blob(),
        MAX_MASTER_XML,
        conformance,
        "notesMaster",
        "notes master",
        limits,
    )?;
    let theme_relationships: Vec<_> = master_part
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == conformance.theme_rel())
        .collect();
    if theme_relationships.len() != 1 {
        return Err(invalid(
            "notes master must have exactly one bounded theme relationship",
        ));
    }
    let theme_relationship = theme_relationships[0];
    validate_id(theme_relationship.r_id())?;
    if theme_relationship.reltype() != conformance.theme_rel() || theme_relationship.is_external() {
        return Err(invalid(
            "notes-master theme relationship has wrong type or target mode",
        ));
    }
    for relationship in master_part.rels().iter() {
        if relationship.r_id() != theme_relationship.r_id() {
            validate_opaque_relationship(package, master_part, relationship)?;
        }
    }
    let theme_name = relationship_target(master_part, theme_relationship)?;
    validate_leaf_path(&theme_name, "/ppt/theme/", "notes-master theme")?;
    let theme_part = package.get_part(&theme_name)?;
    let theme_part_name = theme_part.partname().clone();
    require_content_type(theme_part, THEME_CT, "notes-master theme")?;
    validate_resource_xml_with_limits(
        theme_part.blob(),
        MAX_THEME_XML,
        conformance,
        "theme",
        "notes-master theme",
        limits,
    )?;
    for relationship in theme_part.rels().iter() {
        validate_opaque_relationship(package, theme_part, relationship)?;
    }
    let mut total = checked_add(
        master_part.blob().len(),
        theme_part.blob().len(),
        "aggregate bytes",
    )?;
    if total > limits.max_total_bytes {
        return Err(limit(
            crate::error::CodecLimitKind::XmlBytes,
            total,
            limits.max_total_bytes,
        ));
    }
    let mut slides = Vec::new();
    let mut discovered = BTreeSet::new();
    for (_, slide_name) in &slide_sources {
        let slide_part = package.get_part(slide_name)?;
        let relationships: Vec<_> = slide_part
            .rels()
            .iter()
            .filter(|relationship| is_notes_slide_rel(relationship.reltype()))
            .collect();
        if relationships.len() > 1 {
            return Err(invalid("slide has multiple notes-slide relationships"));
        }
        let Some(relationship) = relationships.first() else {
            continue;
        };
        validate_id(relationship.r_id())?;
        if relationship.reltype() != conformance.notes_slide_rel() || relationship.is_external() {
            return Err(invalid(
                "notes-slide relationship has wrong type or target mode",
            ));
        }
        let notes_name = relationship_target(slide_part, relationship)?;
        validate_leaf_path(&notes_name, "/ppt/notesSlides/", "notes slide")?;
        let notes_part = package.get_part(&notes_name)?;
        let notes_part_name = notes_part.partname().clone();
        if !discovered.insert(notes_part_name.as_str().to_owned()) {
            return Err(invalid("multiple slides reference the same notes slide"));
        }
        require_content_type(notes_part, ct::PML_NOTES_SLIDE, "notes slide")?;
        validate_resource_xml_with_limits(
            notes_part.blob(),
            MAX_NOTES_XML,
            conformance,
            "notes",
            "notes slide",
            limits,
        )?;
        let known_relationships = notes_part
            .rels()
            .iter()
            .filter(|relationship| {
                relationship.reltype() == conformance.slide_rel()
                    || relationship.reltype() == conformance.notes_master_rel()
            })
            .count();
        if known_relationships != 2 {
            return Err(invalid(
                "notes slide must have exactly slide and notes-master relationships",
            ));
        }
        let mut backlink = None;
        let mut notes_master = None;
        for child in notes_part.rels().iter() {
            validate_id(child.r_id())?;
            if child.is_external() {
                return Err(invalid("notes slide has an external relationship"));
            }
            if child.reltype() == conformance.slide_rel() {
                if backlink.replace(child).is_some() {
                    return Err(invalid("notes slide has multiple slide backlinks"));
                }
            } else if child.reltype() == conformance.notes_master_rel() {
                if notes_master.replace(child).is_some() {
                    return Err(invalid(
                        "notes slide has multiple notes-master relationships",
                    ));
                }
            } else {
                validate_opaque_relationship(package, notes_part, child)?;
            }
        }
        let backlink = backlink.ok_or_else(|| invalid("notes slide lacks slide backlink"))?;
        let notes_master =
            notes_master.ok_or_else(|| invalid("notes slide lacks notes-master relationship"))?;
        let backlink_target = relationship_target(notes_part, backlink)?;
        if package.get_part(&backlink_target)?.partname() != slide_name {
            return Err(invalid("notes slide backlink targets the wrong slide"));
        }
        let notes_master_target = relationship_target(notes_part, notes_master)?;
        if package.get_part(&notes_master_target)?.partname() != &master_part_name {
            return Err(invalid("notes slide targets the wrong notes master"));
        }
        total = checked_add(total, notes_part.blob().len(), "aggregate bytes")?;
        if total > limits.max_total_bytes {
            return Err(limit(
                crate::error::CodecLimitKind::XmlBytes,
                total,
                limits.max_total_bytes,
            ));
        }
        slides.push(SlideIndex {
            slide_part_name: slide_part.partname().clone(),
            slide_relationship_id: relationship.r_id().to_owned(),
            part_name: notes_part_name,
            content_type: notes_part.content_type().to_owned(),
            backlink_relationship_id: backlink.r_id().to_owned(),
            notes_master_relationship_id: notes_master.r_id().to_owned(),
        });
    }
    if package
        .iter_parts()
        .filter(|part| part.content_type() == ct::PML_NOTES_MASTER)
        .count()
        != 1
        || package
            .iter_parts()
            .filter(|part| part.content_type() == ct::PML_NOTES_SLIDE)
            .any(|part| !discovered.contains(part.partname().as_str()))
    {
        return Err(invalid("package contains orphan notes parts"));
    }
    Ok(Some(GraphIndex {
        conformance,
        master: MasterIndex {
            presentation_relationship_id: master_id.clone(),
            part_name: master_part_name,
            content_type: master_part.content_type().to_owned(),
            theme: ThemeIndex {
                relationship_id: theme_relationship.r_id().to_owned(),
                part_name: theme_part_name,
                content_type: theme_part.content_type().to_owned(),
            },
        },
        slides,
    }))
}

fn same_raw_source(left: &[u8], right: &[u8]) -> bool {
    left.as_ptr() == right.as_ptr() && left.len() == right.len()
}

fn source_proof_matches_current(proof: Option<&SlideRootProof<'_>>, current: &[u8]) -> bool {
    proof.is_some_and(|proof| proof.source_checked && same_raw_source(proof.raw, current))
}

/// A capture proof was classified with the default notes scanner. Slide-root
/// byte and node ceilings are intentionally independent from that profile,
/// but depth and aggregate attribute bytes remain selected bounds. A proof
/// can therefore answer a recipient whose bounds are at least the defaults;
/// tighter recipients must revalidate the current source.
fn default_proof_covers(limits: Limits) -> bool {
    limits.max_depth >= Limits::DEFAULT.max_depth
        && limits.max_attribute_bytes >= Limits::DEFAULT.max_attribute_bytes
}

#[cfg(test)]
fn resolve_slide_root<'a>(
    current: &[u8],
    proof: Option<&SlideRootProof<'a>>,
    limits: Limits,
) -> Result<Conformance> {
    resolve_slide_root_with_profiles(current, proof, limits, limits)
}

fn resolve_slide_root_with_profiles<'a>(
    current: &[u8],
    proof: Option<&SlideRootProof<'a>>,
    proof_limits: Limits,
    scan_limits: Limits,
) -> Result<Conformance> {
    match proof
        .filter(|proof| same_raw_source(proof.raw, current) && default_proof_covers(proof_limits))
    {
        Some(proof) => proof
            .conformance
            .ok_or_else(|| invalid("invalid sld root or namespace")),
        None => root_conformance_with_limits(current, MAX_SLIDE_XML, "sld", scan_limits),
    }
}

fn materialize(package: &OpcPackage, index: GraphIndex) -> Result<Graph> {
    let master_data = own_blob(
        package.get_part(&index.master.part_name)?.blob(),
        "notes-master payload",
    )?;
    let theme_data = own_blob(
        package.get_part(&index.master.theme.part_name)?.blob(),
        "notes-master theme payload",
    )?;
    let mut slides = Vec::new();
    slides
        .try_reserve(index.slides.len())
        .map_err(|source| allocation("notes-slide graph", source))?;
    for slide in index.slides {
        let data = own_blob(
            package.get_part(&slide.part_name)?.blob(),
            "notes-slide payload",
        )?;
        slides.push(Slide {
            slide_part_name: slide.slide_part_name.as_str().to_owned(),
            slide_relationship_id: slide.slide_relationship_id,
            part_name: slide.part_name.as_str().to_owned(),
            content_type: slide.content_type,
            data,
            relationships: relationship_links(package.get_part(&slide.part_name)?.rels()),
            backlink_relationship_id: slide.backlink_relationship_id,
            notes_master_relationship_id: slide.notes_master_relationship_id,
        });
    }
    Ok(Graph {
        conformance: index.conformance,
        master: Master {
            presentation_relationship_id: index.master.presentation_relationship_id,
            part_name: index.master.part_name.as_str().to_owned(),
            content_type: index.master.content_type,
            data: master_data,
            relationships: relationship_links(package.get_part(&index.master.part_name)?.rels()),
            theme: Theme {
                relationship_id: index.master.theme.relationship_id,
                part_name: index.master.theme.part_name.as_str().to_owned(),
                content_type: index.master.theme.content_type,
                data: theme_data,
                relationships: relationship_links(
                    package.get_part(&index.master.theme.part_name)?.rels(),
                ),
            },
        },
        slides,
    })
}

#[derive(Debug)]
struct Removal {
    slide_part_name: PackURI,
    relationship_id: String,
    notes_part_name: PackURI,
}

impl From<&SlideIndex> for Removal {
    fn from(slide: &SlideIndex) -> Self {
        Self {
            slide_part_name: slide.slide_part_name.clone(),
            relationship_id: slide.slide_relationship_id.clone(),
            notes_part_name: slide.part_name.clone(),
        }
    }
}

/// Remove the speaker-notes resource owned by one presentation slide.
///
/// The complete notes graph and every inbound edge to the selected resource
/// are validated before mutation. Missing notes are an idempotent `Ok(false)`.
/// Shared notes-master and theme resources are retained.
///
/// # Errors
///
/// Returns an error when the notes graph, selected slide, relationships, or
/// descendant ownership is malformed or exceeds a resource limit.
#[cfg(test)]
pub(crate) fn remove(
    package: &mut OpcPackage,
    presentation_name: &PackURI,
    slide_name: &PackURI,
) -> Result<bool> {
    remove_with_limits(
        package,
        presentation_name,
        slide_name,
        Limits::from_source(package.read_limits()),
    )
}

pub(crate) fn remove_with_limits(
    package: &mut OpcPackage,
    presentation_name: &PackURI,
    slide_name: &PackURI,
    limits: Limits,
) -> Result<bool> {
    let Some(index) = load_index_with_limits(package, presentation_name, limits)? else {
        return Ok(false);
    };
    let slide_name = package.get_part(slide_name)?.partname().clone();
    let Some(slide) = index
        .slides
        .iter()
        .find(|slide| slide.slide_part_name == slide_name)
    else {
        return Ok(false);
    };
    let removals = [Removal::from(slide)];
    let descendants = validate_notes_removals(package, &removals)?;
    apply_notes_removals(package, &removals, &descendants)?;
    Ok(true)
}

#[cfg(test)]
pub(crate) fn remove_checked(
    package: &mut OpcPackage,
    source: &Snapshot,
    slide_name: &PackURI,
) -> Result<bool> {
    require_current_source(package, source)?;
    remove(package, source.presentation_part_name(), slide_name)
}

pub(crate) fn remove_checked_with_limits(
    package: &mut OpcPackage,
    source: &Snapshot,
    slide_name: &PackURI,
    limits: Limits,
) -> Result<bool> {
    require_current_source_with_limits(package, source, limits)?;
    remove_with_limits(package, source.presentation_part_name(), slide_name, limits)
}

/// Remove every speaker-notes resource from a presentation.
///
/// Returns the number of removed notes slides. The operation is idempotent,
/// validates the complete graph before mutation, and retains the shared notes
/// master and its theme so ordinary presentation layout remains unchanged.
///
/// # Errors
///
/// Returns an error when the notes graph, relationships, or descendant
/// ownership is malformed or exceeds a resource limit.
#[cfg(test)]
pub(crate) fn clear(package: &mut OpcPackage, presentation_name: &PackURI) -> Result<usize> {
    clear_with_limits(
        package,
        presentation_name,
        Limits::from_source(package.read_limits()),
    )
}

pub(crate) fn clear_with_limits(
    package: &mut OpcPackage,
    presentation_name: &PackURI,
    limits: Limits,
) -> Result<usize> {
    let Some(index) = load_index_with_limits(package, presentation_name, limits)? else {
        return Ok(0);
    };
    let mut removals = Vec::new();
    removals
        .try_reserve(index.slides.len())
        .map_err(|source| allocation("notes-removal plan", source))?;
    removals.extend(index.slides.iter().map(Removal::from));
    if removals.is_empty() {
        return Ok(0);
    }
    let descendants = validate_notes_removals(package, &removals)?;
    apply_notes_removals(package, &removals, &descendants)
}

pub(crate) fn clear_checked_with_limits(
    package: &mut OpcPackage,
    source: &Snapshot,
    limits: Limits,
) -> Result<usize> {
    require_current_source_with_limits(package, source, limits)?;
    clear_with_limits(package, source.presentation_part_name(), limits)
}

#[cfg(test)]
fn require_current_source(package: &OpcPackage, source: &Snapshot) -> Result<()> {
    require_current_source_with_limits(package, source, Limits::from_source(package.read_limits()))
}

fn require_current_source_with_limits(
    package: &OpcPackage,
    source: &Snapshot,
    limits: Limits,
) -> Result<()> {
    let current = load_snapshot_with_limits(package, source.presentation_part_name(), limits)?
        .ok_or_else(|| invalid("notes mutation source graph is absent"))?;
    if current.same_source(source) {
        Ok(())
    } else {
        Err(invalid("notes mutation source is stale"))
    }
}

fn validate_notes_removals(package: &OpcPackage, removals: &[Removal]) -> Result<Vec<PackURI>> {
    let mut by_target = HashMap::new();
    by_target
        .try_reserve(removals.len())
        .map_err(|source| allocation("notes-removal index", source))?;
    for (index, removal) in removals.iter().enumerate() {
        if by_target
            .insert(removal.notes_part_name.clone(), index)
            .is_some()
        {
            return Err(invalid("notes-removal plan contains a duplicate target"));
        }
    }
    let mut inbound_counts = Vec::new();
    inbound_counts
        .try_reserve(removals.len())
        .map_err(|source| allocation("notes-removal counters", source))?;
    inbound_counts.resize(removals.len(), 0usize);

    for relationship in package.rels().iter() {
        validate_notes_inbound(
            package,
            None,
            relationship,
            removals,
            &by_target,
            &mut inbound_counts,
        )?;
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            validate_notes_inbound(
                package,
                Some(source.partname()),
                relationship,
                removals,
                &by_target,
                &mut inbound_counts,
            )?;
        }
    }
    if inbound_counts.iter().any(|count| *count != 1) {
        return Err(invalid(
            "notes-removal target does not have exactly one owning slide relationship",
        ));
    }
    plan_owned_descendants(package, removals)
}

fn relationship_disposition(value: &str, notes_root: bool) -> Option<Disposition> {
    const TRANSITIONAL: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    const STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/";
    const OFFICE_2007: &str = "http://schemas.microsoft.com/office/2007/relationships/";
    const OFFICE_2011: &str = "http://schemas.microsoft.com/office/2011/relationships/";
    let kind = value
        .strip_prefix(TRANSITIONAL)
        .or_else(|| value.strip_prefix(STRICT))
        .or_else(|| value.strip_prefix(OFFICE_2007))
        .or_else(|| value.strip_prefix(OFFICE_2011))?;
    match kind {
        "slide" | "notesMaster" => Some(Disposition::Retain),
        "hyperlink" => Some(Disposition::External),
        "additionalCharacteristics"
        | "bibliography"
        | "customXml"
        | "themeOverride"
        | "thumbnail"
        | "audio"
        | "chart"
        | "contentPart"
        | "diagramColors"
        | "diagramData"
        | "diagramLayout"
        | "diagramQuickStyle"
        | "control"
        | "oleObject"
        | "package"
        | "image"
        | "video" => Some(Disposition::Follow),
        "chartUserShapes" | "chartStyle" | "chartColorStyle" | "ctrlProp" | "worksheet"
            if !notes_root =>
        {
            Some(Disposition::Follow)
        },
        _ => None,
    }
}

fn plan_owned_descendants(package: &OpcPackage, removals: &[Removal]) -> Result<Vec<PackURI>> {
    let mut roots = HashSet::new();
    roots
        .try_reserve(removals.len())
        .map_err(|reserve_error| allocation("notes-removal roots", reserve_error))?;
    roots.extend(
        removals
            .iter()
            .map(|removal| removal.notes_part_name.clone()),
    );

    let mut closure = HashSet::new();
    closure
        .try_reserve(removals.len())
        .map_err(|reserve_error| allocation("notes-owned part closure", reserve_error))?;
    closure.extend(roots.iter().cloned());
    let mut pending = Vec::new();
    pending
        .try_reserve(removals.len())
        .map_err(|reserve_error| allocation("notes-owned part work list", reserve_error))?;
    pending.extend(roots.iter().cloned());

    let mut cursor = 0usize;
    while let Some(source_name) = pending.get(cursor).cloned() {
        cursor = cursor
            .checked_add(1)
            .ok_or_else(|| invalid("notes-owned part work-list index overflow"))?;
        let source = package.get_part(&source_name)?;
        for relationship in source.rels().iter() {
            let disposition =
                relationship_disposition(relationship.reltype(), roots.contains(&source_name))
                    .ok_or_else(|| {
                        invalid(format!(
                            "notes deletion refuses unknown relationship type '{}' from '{}'",
                            relationship.reltype(),
                            source_name.as_str(),
                        ))
                    })?;
            match disposition {
                Disposition::Retain if relationship.is_external() => {
                    return Err(invalid(format!(
                        "retained notes relationship '{}' cannot be external",
                        relationship.r_id(),
                    )));
                },
                Disposition::Retain => continue,
                Disposition::External => {
                    if relationship.is_external() {
                        continue;
                    }
                    return Err(invalid(format!(
                        "notes hyperlink relationship '{}' must be external",
                        relationship.r_id(),
                    )));
                },
                Disposition::Follow => {
                    if relationship.is_external() {
                        continue;
                    }
                },
            }
            let target = relationship.target_partname()?;
            let stored = package.get_part(&target)?.partname().clone();
            if !closure.contains(&stored) {
                if closure.len().saturating_sub(roots.len()) >= MAX_OWNED_PARTS {
                    return Err(limit(
                        crate::error::CodecLimitKind::Parts,
                        closure.len().checked_add(1).unwrap_or(usize::MAX),
                        MAX_OWNED_PARTS,
                    ));
                }
                closure.try_reserve(1).map_err(|reserve_error| {
                    allocation("notes-owned part closure", reserve_error)
                })?;
                pending.try_reserve(1).map_err(|reserve_error| {
                    allocation("notes-owned part work list", reserve_error)
                })?;
                closure.insert(stored.clone());
                pending.push(stored);
            }
        }
    }

    for relationship in package.rels().iter() {
        validate_descendant_inbound(package, None, relationship, &roots, &closure)?;
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            validate_descendant_inbound(
                package,
                Some(source.partname()),
                relationship,
                &roots,
                &closure,
            )?;
        }
    }

    let mut descendants = Vec::new();
    descendants
        .try_reserve(closure.len().saturating_sub(roots.len()))
        .map_err(|reserve_error| allocation("notes-owned descendant plan", reserve_error))?;
    descendants.extend(closure.into_iter().filter(|name| !roots.contains(name)));
    descendants.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
    Ok(descendants)
}

fn validate_descendant_inbound(
    package: &OpcPackage,
    source: Option<&PackURI>,
    relationship: &litchi_opc::Relationship,
    roots: &HashSet<PackURI>,
    closure: &HashSet<PackURI>,
) -> Result<()> {
    if relationship.is_external() {
        return Ok(());
    }
    let Ok(target_reference) = relationship.target_partname() else {
        return Ok(());
    };
    let stored_target = match package.get_part(&target_reference) {
        Ok(part) => part.partname(),
        Err(OpcError::PartNotFound(_)) => return Ok(()),
        Err(error) => return Err(error.into()),
    };
    if roots.contains(stored_target) || !closure.contains(stored_target) {
        return Ok(());
    }
    if source.is_some_and(|source_name| closure.contains(source_name)) {
        return Ok(());
    }
    Err(invalid(format!(
        "notes-owned descendant '{}' has a shared inbound relationship '{}' from '{}'",
        stored_target.as_str(),
        relationship.r_id(),
        source.map_or("package root", PackURI::as_str),
    )))
}

fn validate_notes_inbound(
    package: &OpcPackage,
    source: Option<&PackURI>,
    relationship: &litchi_opc::Relationship,
    removals: &[Removal],
    by_target: &HashMap<PackURI, usize>,
    inbound_counts: &mut [usize],
) -> Result<()> {
    if relationship.is_external() {
        return Ok(());
    }
    let Ok(target) = relationship.target_partname() else {
        return Ok(());
    };
    let index = if let Some(index) = by_target.get(&target).copied() {
        Some(index)
    } else {
        match package.get_part(&target) {
            Ok(part) => by_target.get(part.partname()).copied(),
            Err(OpcError::PartNotFound(_)) => None,
            Err(error) => return Err(error.into()),
        }
    };
    let Some(index) = index else {
        return Ok(());
    };
    let removal = &removals[index];
    if source != Some(&removal.slide_part_name) || relationship.r_id() != removal.relationship_id {
        let source = source.map_or("package root", PackURI::as_str);
        return Err(invalid(format!(
            "notes slide '{}' has an unexpected inbound relationship '{}' from '{}'",
            removal.notes_part_name.as_str(),
            relationship.r_id(),
            source,
        )));
    }
    inbound_counts[index] = inbound_counts[index]
        .checked_add(1)
        .ok_or_else(|| invalid("notes inbound relationship count overflow"))?;
    Ok(())
}

fn apply_notes_removals(
    package: &mut OpcPackage,
    removals: &[Removal],
    descendants: &[PackURI],
) -> Result<usize> {
    // Stage cloned slide owners before the first package mutation. Built-in
    // parts retain their shared payload allocation while relationships are
    // detached on the staged clone.
    let mut staged_slides = Vec::new();
    staged_slides
        .try_reserve(removals.len())
        .map_err(|source| allocation("notes-removal staging", source))?;
    for removal in removals {
        let slide = package.get_part(&removal.slide_part_name)?;
        let relationship = slide
            .rels()
            .get(&removal.relationship_id)
            .ok_or_else(|| invalid("validated notes relationship disappeared before commit"))?;
        if relationship.is_external() {
            return Err(invalid(
                "validated notes relationship changed before commit",
            ));
        }
        let target = relationship.target_partname()?;
        if package.get_part(&target)?.partname() != &removal.notes_part_name {
            return Err(invalid(
                "validated notes relationship changed before commit",
            ));
        }
        package.get_part(&removal.notes_part_name)?;
        let mut staged = slide.clone_part();
        if staged.rels_mut().remove(&removal.relationship_id).is_none() {
            return Err(invalid("validated notes relationship was not removed"));
        }
        staged_slides.push(staged);
    }
    for descendant in descendants {
        package.get_part(descendant)?;
    }

    // Every mutation below is infallible after validation, staging and the
    // signature refusal. Exact stored part names avoid the case-insensitive
    // lookup/exact-removal trap.
    refuse_signed_change(package)?;
    for staged in staged_slides {
        package.add_part(staged);
    }
    for removal in removals {
        package.remove_part(&removal.notes_part_name);
    }
    for descendant in descendants {
        package.remove_part(descendant);
    }
    settle_unsigned_change(package)?;
    Ok(removals.len())
}

/// Deterministically replace the resources of an already coherent notes graph.
/// Validation completes before the first package mutation.
///
/// # Errors
///
/// Returns an error when graph validation, ownership validation, allocation,
/// or staged relationship publication fails.
#[cfg(test)]
pub(crate) fn put(
    package: &mut OpcPackage,
    presentation_name: &PackURI,
    graph: Graph,
) -> Result<()> {
    put_changed(package, presentation_name, graph).map(|_| ())
}

/// Replace a graph and report whether package bytes changed.
///
/// # Errors
///
/// Returns the same failures as [`put`].
#[cfg(test)]
pub(crate) fn put_changed(
    package: &mut OpcPackage,
    presentation_name: &PackURI,
    graph: Graph,
) -> Result<bool> {
    put_changed_with_limits(
        package,
        presentation_name,
        graph,
        Limits::from_source(package.read_limits()),
    )
}

#[cfg(test)]
pub(crate) fn put_changed_with_limits(
    package: &mut OpcPackage,
    presentation_name: &PackURI,
    graph: Graph,
    limits: Limits,
) -> Result<bool> {
    let current = load_index_with_limits(package, presentation_name, limits)?
        .ok_or_else(|| invalid("store requires an existing coherent notes graph"))?;
    validate_graph_with_limits(&graph, limits)?;
    if indexed_ownership(&current) != ownership(&graph) {
        return Err(invalid(
            "store cannot retarget or orphan existing notes parts",
        ));
    }
    if graph_matches(package, &current, &graph)? {
        return Ok(false);
    }
    let presentation = package.get_part(presentation_name)?;
    let presentation_scan = scan_xml_with_limits(
        presentation.blob(),
        MAX_PRESENTATION_XML,
        graph.conformance,
        "presentation",
        limits,
    )?;
    if presentation_scan.notes_master_ids.as_slice()
        != [graph.master.presentation_relationship_id.as_str()]
    {
        return Err(invalid(
            "presentation notesMasterId does not match graph metadata",
        ));
    }
    let mut slide_names = BTreeSet::new();
    for id in presentation_scan.slide_ids {
        let relationship = presentation
            .rels()
            .get(&id)
            .ok_or_else(|| invalid("presentation slide reference is missing"))?;
        let target = relationship_target(presentation, relationship)?;
        slide_names.insert(package.get_part(&target)?.partname().as_str().to_owned());
    }
    if graph
        .slides
        .iter()
        .any(|slide| !slide_names.contains(&slide.slide_part_name))
    {
        return Err(invalid(
            "notes graph references a slide outside the presentation",
        ));
    }

    let Graph {
        conformance,
        master,
        slides,
    } = graph;
    let Master {
        presentation_relationship_id,
        part_name: master_name,
        content_type: master_content_type,
        data: master_data,
        relationships: mut master_relationships,
        theme,
    } = master;
    let Theme {
        relationship_id: theme_relationship_id,
        part_name: theme_name,
        content_type: theme_content_type,
        data: theme_data,
        relationships: theme_relationships,
    } = theme;
    let theme_uri = PackURI::new(theme_name).map_err(Error::Invalid)?;
    let master_uri = PackURI::new(master_name).map_err(Error::Invalid)?;

    let mut theme_part = BlobPart::new(theme_uri.clone(), theme_content_type, theme_data);
    let mut master_part = BlobPart::new(master_uri.clone(), master_content_type, master_data);
    if master_relationships.is_empty() {
        master_relationships.push(Link::new(
            theme_relationship_id,
            conformance.theme_rel(),
            theme_uri.relative_ref(master_uri.base_uri()),
            TargetMode::Internal,
        ));
    }
    add_links(&mut master_part, &master_relationships)?;
    add_links(&mut theme_part, &theme_relationships)?;

    let mut note_parts = Vec::new();
    note_parts
        .try_reserve(slides.len())
        .map_err(|source| allocation("notes-part staging", source))?;
    let mut by_slide = BTreeMap::new();
    for slide in slides {
        let Slide {
            slide_part_name,
            slide_relationship_id,
            part_name,
            content_type,
            data,
            relationships,
            backlink_relationship_id,
            notes_master_relationship_id,
        } = slide;
        let notes_uri = PackURI::new(part_name).map_err(Error::Invalid)?;
        let slide_uri = PackURI::new(&slide_part_name).map_err(Error::Invalid)?;
        let mut notes_part = BlobPart::new(notes_uri.clone(), content_type, data);
        let mut relationships = relationships;
        if relationships.is_empty() {
            relationships.push(Link::new(
                backlink_relationship_id,
                conformance.slide_rel(),
                slide_uri.relative_ref(notes_uri.base_uri()),
                TargetMode::Internal,
            ));
            relationships.push(Link::new(
                notes_master_relationship_id,
                conformance.notes_master_rel(),
                master_uri.relative_ref(notes_uri.base_uri()),
                TargetMode::Internal,
            ));
        }
        add_links(&mut notes_part, &relationships)?;
        if by_slide
            .insert(slide_part_name, (notes_uri, slide_relationship_id))
            .is_some()
        {
            return Err(invalid("notes graph contains duplicate slide owners"));
        }
        note_parts.push(notes_part);
    }

    let mut staged_presentation = package.get_part(presentation_name)?.clone_part();
    let presentation_ids: Vec<_> = staged_presentation
        .rels()
        .iter()
        .filter(|relationship| is_notes_master_rel(relationship.reltype()))
        .map(|relationship| relationship.r_id().to_owned())
        .collect();
    for id in presentation_ids {
        staged_presentation.rels_mut().remove(&id);
    }
    staged_presentation.rels_mut().try_add_relationship(
        conformance.notes_master_rel().into(),
        master_uri.relative_ref(presentation_name.base_uri()),
        presentation_relationship_id,
        TargetMode::Internal,
    )?;

    let mut staged_slides = Vec::new();
    staged_slides
        .try_reserve(slide_names.len())
        .map_err(|source| allocation("notes slide-owner staging", source))?;
    for slide_name in slide_names {
        let uri = PackURI::new(&slide_name).map_err(Error::Invalid)?;
        let mut part = package.get_part(&uri)?.clone_part();
        let ids: Vec<_> = part
            .rels()
            .iter()
            .filter(|relationship| is_notes_slide_rel(relationship.reltype()))
            .map(|relationship| relationship.r_id().to_owned())
            .collect();
        for id in ids {
            part.rels_mut().remove(&id);
        }
        if let Some((notes_uri, relationship_id)) = by_slide.remove(&slide_name) {
            part.rels_mut().try_add_relationship(
                conformance.notes_slide_rel().into(),
                notes_uri.relative_ref(uri.base_uri()),
                relationship_id,
                TargetMode::Internal,
            )?;
        }
        staged_slides.push(part);
    }
    if !by_slide.is_empty() {
        return Err(invalid("notes graph contains an unknown slide owner"));
    }

    // Prepared is infallible after all URI, graph, allocation, relationship and
    // signature checks succeed. Owned payload buffers move into their
    // canonical parts.
    refuse_signed_change(package)?;
    package.add_part(Box::new(theme_part));
    package.add_part(Box::new(master_part));
    for part in note_parts {
        package.add_part(Box::new(part));
    }
    package.add_part(staged_presentation);
    for slide in staged_slides {
        package.add_part(slide);
    }
    settle_unsigned_change(package)?;
    Ok(true)
}

pub(crate) fn validate_graph_with_limits(graph: &Graph, limits: Limits) -> Result<()> {
    if graph.slides.len() > MAX_NOTES_SLIDES {
        return Err(limit(
            crate::error::CodecLimitKind::Slides,
            graph.slides.len(),
            MAX_NOTES_SLIDES,
        ));
    }
    validate_id(&graph.master.presentation_relationship_id)?;
    validate_id(&graph.master.theme.relationship_id)?;
    if graph.master.content_type != ct::PML_NOTES_MASTER
        || graph.master.theme.content_type != THEME_CT
    {
        return Err(invalid("notes master or theme has invalid content type"));
    }
    let master_uri = PackURI::new(&graph.master.part_name).map_err(Error::Invalid)?;
    validate_leaf_path(&master_uri, "/ppt/notesMasters/", "notes master")?;
    let theme_uri = PackURI::new(&graph.master.theme.part_name).map_err(Error::Invalid)?;
    validate_leaf_path(&theme_uri, "/ppt/theme/", "notes-master theme")?;
    validate_resource_xml_with_limits(
        &graph.master.data,
        MAX_MASTER_XML,
        graph.conformance,
        "notesMaster",
        "notes master",
        limits,
    )?;
    super::validation::validate_links(&graph.master.relationships)?;
    validate_resource_xml_with_limits(
        &graph.master.theme.data,
        MAX_THEME_XML,
        graph.conformance,
        "theme",
        "notes-master theme",
        limits,
    )?;
    super::validation::validate_links(&graph.master.theme.relationships)?;
    let mut total = checked_add(
        graph.master.data.len(),
        graph.master.theme.data.len(),
        "aggregate bytes",
    )?;
    let mut sources = HashSet::new();
    let mut parts = HashSet::new();
    parts.insert(graph.master.part_name.as_str());
    parts.insert(graph.master.theme.part_name.as_str());
    for slide in &graph.slides {
        validate_id(&slide.slide_relationship_id)?;
        validate_id(&slide.backlink_relationship_id)?;
        validate_id(&slide.notes_master_relationship_id)?;
        if slide.backlink_relationship_id == slide.notes_master_relationship_id {
            return Err(invalid("notes-slide relationship IDs collide"));
        }
        if slide.content_type != ct::PML_NOTES_SLIDE {
            return Err(invalid("notes slide has invalid content type"));
        }
        let source = PackURI::new(&slide.slide_part_name).map_err(Error::Invalid)?;
        validate_leaf_path(&source, "/ppt/slides/", "slide")?;
        let uri = PackURI::new(&slide.part_name).map_err(Error::Invalid)?;
        validate_leaf_path(&uri, "/ppt/notesSlides/", "notes slide")?;
        if !sources.insert(slide.slide_part_name.as_str())
            || !parts.insert(slide.part_name.as_str())
        {
            return Err(invalid(
                "notes graph has duplicate source or resource part names",
            ));
        }
        validate_resource_xml_with_limits(
            &slide.data,
            MAX_NOTES_XML,
            graph.conformance,
            "notes",
            "notes slide",
            limits,
        )?;
        super::validation::validate_links(&slide.relationships)?;
        total = checked_add(total, slide.data.len(), "aggregate bytes")?;
        if total > limits.max_total_bytes {
            return Err(limit(
                crate::error::CodecLimitKind::XmlBytes,
                total,
                limits.max_total_bytes,
            ));
        }
    }
    Ok(())
}
#[cfg(test)]
fn ownership(graph: &Graph) -> BTreeSet<&str> {
    std::iter::once(graph.master.part_name.as_str())
        .chain(std::iter::once(graph.master.theme.part_name.as_str()))
        .chain(graph.slides.iter().map(|slide| slide.part_name.as_str()))
        .collect()
}

#[cfg(test)]
fn indexed_ownership(graph: &GraphIndex) -> BTreeSet<&str> {
    std::iter::once(graph.master.part_name.as_str())
        .chain(std::iter::once(graph.master.theme.part_name.as_str()))
        .chain(graph.slides.iter().map(|slide| slide.part_name.as_str()))
        .collect()
}

#[cfg(test)]
fn graph_matches(package: &OpcPackage, index: &GraphIndex, graph: &Graph) -> Result<bool> {
    if index.conformance != graph.conformance
        || index.master.presentation_relationship_id != graph.master.presentation_relationship_id
        || index.master.part_name.as_str() != graph.master.part_name
        || index.master.content_type != graph.master.content_type
        || index.master.theme.relationship_id != graph.master.theme.relationship_id
        || index.master.theme.part_name.as_str() != graph.master.theme.part_name
        || index.master.theme.content_type != graph.master.theme.content_type
        || index.slides.len() != graph.slides.len()
        || package.get_part(&index.master.part_name)?.blob() != graph.master.data
        || package.get_part(&index.master.theme.part_name)?.blob() != graph.master.theme.data
    {
        return Ok(false);
    }
    for (stored, candidate) in index.slides.iter().zip(&graph.slides) {
        if stored.slide_part_name.as_str() != candidate.slide_part_name
            || stored.slide_relationship_id != candidate.slide_relationship_id
            || stored.part_name.as_str() != candidate.part_name
            || stored.content_type != candidate.content_type
            || stored.backlink_relationship_id != candidate.backlink_relationship_id
            || stored.notes_master_relationship_id != candidate.notes_master_relationship_id
            || package.get_part(&stored.part_name)?.blob() != candidate.data
        {
            return Ok(false);
        }
    }
    Ok(true)
}

fn relationship_target(
    part: &dyn Part,
    relationship: &litchi_opc::Relationship,
) -> Result<PackURI> {
    if relationship.is_external() {
        return Err(invalid("external relationship is rejected"));
    }
    PackURI::from_rel_ref(part.partname().base_uri(), relationship.target_ref())
        .map_err(Error::Invalid)
}

fn relationship_links(relationships: &litchi_opc::Relationships) -> Vec<Link> {
    let mut links: Vec<_> = relationships
        .iter()
        .map(|relationship| {
            Link::new(
                relationship.r_id(),
                relationship.reltype(),
                relationship.target_ref(),
                relationship.target_mode(),
            )
        })
        .collect();
    links.sort_unstable_by(|left, right| left.id.cmp(&right.id));
    links
}

#[cfg(test)]
fn add_links(part: &mut dyn Part, links: &[Link]) -> Result<()> {
    for link in links {
        part.rels_mut().try_add_relationship(
            link.relationship_type.clone(),
            link.target_ref.clone(),
            link.id.clone(),
            link.target_mode,
        )?;
    }
    Ok(())
}

fn validate_opaque_relationship(
    package: &OpcPackage,
    source: &dyn Part,
    relationship: &litchi_opc::Relationship,
) -> Result<()> {
    validate_id(relationship.r_id())?;
    if relationship.is_external() {
        return Ok(());
    }
    let target = relationship.target_partname()?;
    package.get_part(&target).map(|_| ()).map_err(|_err| {
        invalid(format!(
            "opaque notes relationship '{}' from '{}' targets a missing part",
            relationship.r_id(),
            source.partname().as_str()
        ))
    })
}

fn validate_leaf_path(uri: &PackURI, prefix: &str, label: &str) -> Result<()> {
    let Some(rest) = uri.as_str().strip_prefix(prefix) else {
        return Err(invalid(format!("{label} is outside {prefix}")));
    };
    if rest.is_empty()
        || rest.contains('/')
        || !rest
            .rsplit_once('.')
            .is_some_and(|(_, extension)| extension.eq_ignore_ascii_case("xml"))
    {
        return Err(invalid(format!("invalid {label} part path")));
    }
    Ok(())
}
fn require_content_type(part: &dyn Part, expected: &str, label: &str) -> Result<()> {
    if part.content_type() == expected {
        Ok(())
    } else {
        Err(invalid(format!("{label} has invalid content type")))
    }
}
pub(super) fn is_notes_slide_rel(value: &str) -> bool {
    matches!(value, rt::NOTES_SLIDE | rt::STRICT_NOTES_SLIDE)
}
pub(super) fn is_notes_master_rel(value: &str) -> bool {
    matches!(value, rt::NOTES_MASTER | rt::STRICT_NOTES_MASTER)
}
fn is_presentation_main_content_type(value: &str) -> bool {
    matches!(
        value,
        ct::PML_PRESENTATION_MAIN
            | ct::PML_SLIDESHOW_MAIN
            | ct::PML_TEMPLATE_MAIN
            | ct::PML_PRES_MACRO_MAIN
            | ct::PML_SLIDESHOW_MACRO_MAIN
            | ct::PML_TEMPLATE_MACRO_MAIN
    )
}
fn validate_id(value: &str) -> Result<()> {
    let mut bytes = value.bytes();
    let Some(first) = bytes.next() else {
        return Err(invalid("relationship ID is empty"));
    };
    if !(first.is_ascii_alphabetic() || first == b'_')
        || !bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        Err(invalid("invalid relationship ID"))
    } else {
        Ok(())
    }
}

#[cfg(test)]
mod slide_root_proof_tests {
    use super::*;

    const VALID_ROOT: &[u8] =
        br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#;
    const STRICT_ROOT: &[u8] =
        br#"<p:sld xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"/>"#;
    const VALID_ROOT_WITH_TRAILING_SPACE: &[u8] =
        br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/> "#;

    fn limit_dimensions(error: Error) -> (crate::error::CodecLimitKind, usize, usize) {
        match error {
            Error::Limit {
                kind: crate::error::LimitKind::Codec(kind),
                observed,
                maximum,
            } => (kind, observed, maximum),
            other => panic!("expected a typed codec limit, got {other:?}"),
        }
    }

    #[test]
    fn root_proof_requires_exact_source_identity_and_falls_back_safely() {
        let invalid_proof = SlideRootProof::new(VALID_ROOT, None);
        let cached_error = resolve_slide_root(VALID_ROOT, Some(&invalid_proof), Limits::DEFAULT)
            .expect_err("an exact invalid proof must retain its generic refusal");
        assert_eq!(
            cached_error.to_string(),
            "invalid PresentationML: invalid sld root or namespace"
        );

        let equal_bytes = VALID_ROOT.to_vec();
        assert_eq!(
            resolve_slide_root(&equal_bytes, Some(&invalid_proof), Limits::DEFAULT)
                .expect("equal bytes in a different allocation must use the raw fallback"),
            Conformance::Transitional
        );

        assert_eq!(
            resolve_slide_root(STRICT_ROOT, Some(&invalid_proof), Limits::DEFAULT)
                .expect("changed valid source must use the raw fallback"),
            Conformance::Strict
        );

        let shortened = &VALID_ROOT_WITH_TRAILING_SPACE[..VALID_ROOT.len()];
        let length_mismatch_proof = SlideRootProof::new(VALID_ROOT_WITH_TRAILING_SPACE, None);
        assert_eq!(
            resolve_slide_root(shortened, Some(&length_mismatch_proof), Limits::DEFAULT)
                .expect("same pointer with a different length must use the raw fallback"),
            Conformance::Transitional
        );
    }

    #[test]
    fn reused_default_proof_revalidates_tighter_depth_and_attribute_limits() {
        let depth_xml = br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:x><p:y/></p:x></p:sld>"#.to_vec();
        let depth_proof =
            SlideRootProof::new(depth_xml.as_slice(), Some(Conformance::Transitional));
        let depth_limits = Limits {
            max_depth: 1,
            ..Limits::DEFAULT
        };
        let reused_depth =
            resolve_slide_root(depth_xml.as_slice(), Some(&depth_proof), depth_limits)
                .expect_err("a tighter depth must not trust the default proof");
        let fresh_depth = resolve_slide_root(depth_xml.as_slice(), None, depth_limits)
            .expect_err("the cold scan must refuse the same depth");
        let reused_depth = limit_dimensions(reused_depth);
        let fresh_depth = limit_dimensions(fresh_depth);
        assert_eq!(reused_depth, fresh_depth);
        assert_eq!(fresh_depth, (crate::error::CodecLimitKind::XmlDepth, 2, 1));

        let attribute_xml = br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:x a="1234"/></p:sld>"#.to_vec();
        let attribute_proof =
            SlideRootProof::new(attribute_xml.as_slice(), Some(Conformance::Transitional));
        let attribute_limits = Limits {
            max_attribute_bytes: 4,
            ..Limits::DEFAULT
        };
        let reused_attribute = resolve_slide_root(
            attribute_xml.as_slice(),
            Some(&attribute_proof),
            attribute_limits,
        )
        .expect_err("a tighter attribute budget must not trust the default proof");
        let fresh_attribute = resolve_slide_root(attribute_xml.as_slice(), None, attribute_limits)
            .expect_err("the cold scan must refuse the same attribute budget");
        let reused_attribute = limit_dimensions(reused_attribute);
        let fresh_attribute = limit_dimensions(fresh_attribute);
        assert_eq!(reused_attribute, fresh_attribute);
        assert_eq!(
            fresh_attribute,
            (crate::error::CodecLimitKind::XmlAttributeBytes, 5, 4)
        );

        assert_eq!(
            resolve_slide_root(
                VALID_ROOT,
                Some(&SlideRootProof::new(
                    VALID_ROOT,
                    Some(Conformance::Transitional),
                )),
                Limits::DEFAULT
            )
            .expect("default bounds retain the memo fast path"),
            Conformance::Transitional
        );
    }

    #[test]
    fn source_admission_does_not_trust_a_foreign_blob_observation() {
        let proof =
            SlideRootProof::with_source_checked(VALID_ROOT, Some(Conformance::Transitional), true);
        assert!(source_proof_matches_current(Some(&proof), VALID_ROOT));
        let copy = VALID_ROOT.to_vec();
        assert!(!source_proof_matches_current(Some(&proof), &copy));
        let shortened = &VALID_ROOT_WITH_TRAILING_SPACE[..VALID_ROOT.len()];
        assert!(!source_proof_matches_current(Some(&proof), shortened));
        assert!(!source_proof_matches_current(None, VALID_ROOT));
    }
}
