#![expect(
    clippy::expect_used,
    reason = "the invariant is established immediately before extraction"
)]
#![expect(
    clippy::module_name_repetitions,
    reason = "public names retain established OOXML facade terminology"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
#![expect(
    clippy::similar_names,
    reason = "domain names mirror distinct OOXML roles"
)]
//! Bounded ownership inventory for reachable `WordprocessingML` stories.
//!
//! Ownership is derived only from OPC relationships. Directory names and XML
//! references do not create ownership, and no linked or embedded content is
//! activated while an inventory is built.

use super::Package;
use crate::{Error, Result};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;
use std::sync::Arc;

const TRANSITIONAL_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const STRICT_RELATIONSHIPS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/";
const TRANSITIONAL_WORD: &[u8] = b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";
// Retain the original conflict topology encoding so existing stale bindings
// remain deterministic while all package features move to this neutral owner.
const TOPOLOGY_MAGIC: &[u8] = b"litchi.docx.conflict.topology.v1\0";

const HARD_MAX_PACKAGE_PARTS: usize = 1_048_576;
const HARD_MAX_STORIES: usize = 4_096;
const HARD_MAX_STORY_BYTES: usize = 512 * 1024 * 1024;
const HARD_MAX_TOTAL_STORY_BYTES: usize = 512 * 1024 * 1024;
const HARD_MAX_RELATIONSHIPS_PER_OWNER: usize = 65_536;
const HARD_MAX_TOTAL_RELATIONSHIPS: usize = 1_000_000;
const HARD_MAX_TOPOLOGY_BYTES: usize = 128 * 1024 * 1024;
const HARD_MAX_PROLOG_EVENTS: usize = 1_000_000;

/// Relationship and XML namespace family used by a Word package.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum StoryDialect {
    /// ECMA-376 Transitional relationship and `WordprocessingML` namespaces.
    Transitional,
    /// ISO/IEC 29500 Strict relationship and `WordprocessingML` namespaces.
    Strict,
}

impl StoryDialect {
    const fn relationships(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_RELATIONSHIPS,
            Self::Strict => STRICT_RELATIONSHIPS,
        }
    }

    const fn word(self) -> &'static [u8] {
        match self {
            Self::Transitional => TRANSITIONAL_WORD,
            Self::Strict => STRICT_WORD,
        }
    }
}

/// Semantic role of one independently parsed `WordprocessingML` story.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub enum StoryKind {
    Main,
    Header,
    Footer,
    Footnotes,
    Endnotes,
    Comments,
    Glossary,
}

impl StoryKind {
    const fn content_type(self) -> Option<&'static str> {
        match self {
            Self::Main => None,
            Self::Header => Some(ct::WML_HEADER),
            Self::Footer => Some(ct::WML_FOOTER),
            Self::Footnotes => Some(ct::WML_FOOTNOTES),
            Self::Endnotes => Some(ct::WML_ENDNOTES),
            Self::Comments => Some(ct::WML_COMMENTS),
            Self::Glossary => Some(ct::WML_DOCUMENT_GLOSSARY),
        }
    }

    const fn root(self) -> &'static [u8] {
        match self {
            Self::Main => b"document",
            Self::Header => b"hdr",
            Self::Footer => b"ftr",
            Self::Footnotes => b"footnotes",
            Self::Endnotes => b"endnotes",
            Self::Comments => b"comments",
            Self::Glossary => b"glossaryDocument",
        }
    }

    const fn singleton_index(self) -> Option<usize> {
        match self {
            Self::Footnotes => Some(0),
            Self::Endnotes => Some(1),
            Self::Comments => Some(2),
            Self::Glossary => Some(3),
            Self::Main | Self::Header | Self::Footer => None,
        }
    }

    const fn code(self) -> u8 {
        match self {
            Self::Main => 0,
            Self::Header => 1,
            Self::Footer => 2,
            Self::Footnotes => 3,
            Self::Endnotes => 4,
            Self::Comments => 5,
            Self::Glossary => 6,
        }
    }
}

/// Exact incoming relationship that owns one story.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct StoryOwner {
    part: Option<PackURI>,
    relationship_id: String,
    relationship_type: String,
    target_ref: String,
}

impl StoryOwner {
    /// Owning part, or `None` when the package root owns the main story.
    #[must_use]
    pub const fn part(&self) -> Option<&PackURI> {
        self.part.as_ref()
    }

    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    #[must_use]
    pub fn relationship_type(&self) -> &str {
        &self.relationship_type
    }

    #[must_use]
    pub fn target_ref(&self) -> &str {
        &self.target_ref
    }
}

/// One reachable, uniquely owned `WordprocessingML` story part.
#[derive(Clone, Debug)]
pub struct StoryPart {
    pub(crate) part: PackURI,
    pub(crate) content_type: String,
    pub(crate) source: Arc<Vec<u8>>,
    kind: StoryKind,
    owner: StoryOwner,
}

impl StoryPart {
    #[must_use]
    pub const fn part(&self) -> &PackURI {
        &self.part
    }

    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    #[must_use]
    pub const fn kind(&self) -> StoryKind {
        self.kind
    }

    #[must_use]
    pub const fn owner(&self) -> &StoryOwner {
        &self.owner
    }

    #[must_use]
    pub fn source(&self) -> &[u8] {
        self.source.as_slice()
    }

    /// Share the immutable OPC payload without copying story bytes.
    #[must_use]
    pub fn source_arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.source)
    }
}

/// Opaque canonical token for stale package-topology detection.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct StoryTopology(Arc<[u8]>);

impl StoryTopology {
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.0.as_ref()
    }

    #[allow(
        dead_code,
        reason = "package transactions share this token as their integrations land"
    )]
    pub(crate) fn shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.0)
    }
}

/// Complete validated inventory of reachable Word stories.
#[derive(Clone, Debug)]
pub struct StoryInventory {
    pub(crate) main: PackURI,
    pub(crate) stories: Vec<StoryPart>,
    pub(crate) topology: Arc<[u8]>,
    pub(crate) total_story_bytes: usize,
    dialect: StoryDialect,
}

impl StoryInventory {
    #[must_use]
    pub const fn main(&self) -> &PackURI {
        &self.main
    }

    /// Main first, then every subsidiary story by canonical part name.
    #[must_use]
    pub fn stories(&self) -> &[StoryPart] {
        &self.stories
    }

    #[must_use]
    pub fn get(&self, part: &PackURI) -> Option<&StoryPart> {
        self.stories.iter().find(|story| story.part == *part)
    }

    #[must_use]
    pub fn topology(&self) -> StoryTopology {
        StoryTopology(Arc::clone(&self.topology))
    }

    #[must_use]
    pub const fn total_story_bytes(&self) -> usize {
        self.total_story_bytes
    }

    #[must_use]
    pub const fn dialect(&self) -> StoryDialect {
        self.dialect
    }
}

/// Resource policy for one semantic story-ownership inventory.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct StoryLimits {
    pub max_package_parts: usize,
    pub max_stories: usize,
    pub max_story_bytes: usize,
    pub max_total_story_bytes: usize,
    pub max_relationships_per_owner: usize,
    pub max_total_relationships: usize,
    pub max_topology_bytes: usize,
    pub max_xml_prolog_events: usize,
}

impl Default for StoryLimits {
    fn default() -> Self {
        Self {
            max_package_parts: 65_536,
            max_stories: 128,
            max_story_bytes: 16 * 1024 * 1024,
            max_total_story_bytes: 128 * 1024 * 1024,
            max_relationships_per_owner: 4_096,
            max_total_relationships: 32_768,
            max_topology_bytes: 16 * 1024 * 1024,
            max_xml_prolog_events: 256,
        }
    }
}

impl StoryLimits {
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn validate(self) -> Result<Self> {
        for (name, value, maximum) in [
            (
                "max_package_parts",
                self.max_package_parts,
                HARD_MAX_PACKAGE_PARTS,
            ),
            ("max_stories", self.max_stories, HARD_MAX_STORIES),
            (
                "max_story_bytes",
                self.max_story_bytes,
                HARD_MAX_STORY_BYTES,
            ),
            (
                "max_total_story_bytes",
                self.max_total_story_bytes,
                HARD_MAX_TOTAL_STORY_BYTES,
            ),
            (
                "max_relationships_per_owner",
                self.max_relationships_per_owner,
                HARD_MAX_RELATIONSHIPS_PER_OWNER,
            ),
            (
                "max_total_relationships",
                self.max_total_relationships,
                HARD_MAX_TOTAL_RELATIONSHIPS,
            ),
            (
                "max_topology_bytes",
                self.max_topology_bytes,
                HARD_MAX_TOPOLOGY_BYTES,
            ),
            (
                "max_xml_prolog_events",
                self.max_xml_prolog_events,
                HARD_MAX_PROLOG_EVENTS,
            ),
        ] {
            if value == 0 || value > maximum {
                return Err(invalid(format!(
                    "DOCX story limit {name} must be in 1..={maximum}, got {value}"
                )));
            }
        }
        Ok(self)
    }
}

impl Package {
    /// Inventory every uniquely owned reachable Word story.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn story_inventory(&self) -> Result<StoryInventory> {
        self.story_inventory_with_limits(StoryLimits::default())
    }

    /// Inventory Word stories under an explicit semantic resource policy.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn story_inventory_with_limits(&self, limits: StoryLimits) -> Result<StoryInventory> {
        self.ensure_story_opc_current("story_inventory_with_limits")?;
        capture(self.opc_package(), limits)
    }
}

/// Capture a validated story graph directly from an OPC candidate.
///
/// This crate-private hook lets failure-atomic package transactions revalidate
/// the candidate inside `edit_semantic_opc` before publication.
pub(crate) fn capture(package: &OpcPackage, limits: StoryLimits) -> Result<StoryInventory> {
    capture_with_policy(package, limits, true)
}

pub(crate) fn capture_with_policy(
    package: &OpcPackage,
    limits: StoryLimits,
    enforce_story_bytes: bool,
) -> Result<StoryInventory> {
    let limits = limits.validate()?;
    if package.part_count() > limits.max_package_parts {
        return Err(invalid(format!(
            "DOCX package part count exceeds {}",
            limits.max_package_parts
        )));
    }

    let root = root_owner(package)?;
    let dialect = relationship_dialect(&root.relationship_type)
        .ok_or_else(|| invalid("unsupported main-document relationship dialect"))?;
    let main_part = package.main_document_part()?;
    validate_main_content_type(main_part.content_type())?;
    let main = main_part.partname().clone();

    let capacity = package.part_count().min(limits.max_stories);
    let mut owned = HashSet::new();
    owned
        .try_reserve(capacity)
        .map_err(|source| Error::Allocation {
            resource: "DOCX story ownership set",
            source,
        })?;
    let mut stories = Vec::new();
    stories
        .try_reserve(capacity)
        .map_err(|source| Error::Allocation {
            resource: "DOCX story inventory",
            source,
        })?;
    let mut total_story_bytes = 0usize;
    push_story(
        &mut stories,
        &mut owned,
        main_part,
        StoryKind::Main,
        root,
        dialect,
        limits,
        enforce_story_bytes,
        &mut total_story_bytes,
    )?;

    let mut relationship_count = 0usize;
    let glossary = discover_owned(
        package,
        &main,
        true,
        dialect,
        limits,
        enforce_story_bytes,
        &mut owned,
        &mut stories,
        &mut relationship_count,
        &mut total_story_bytes,
    )?;
    if let Some(glossary) = glossary {
        discover_owned(
            package,
            &glossary,
            false,
            dialect,
            limits,
            enforce_story_bytes,
            &mut owned,
            &mut stories,
            &mut relationship_count,
            &mut total_story_bytes,
        )?;
    }

    for story in stories
        .iter()
        .filter(|story| !matches!(story.kind, StoryKind::Main | StoryKind::Glossary))
    {
        let part = package.get_part(&story.part)?;
        if part
            .rels()
            .iter()
            .any(|relationship| relationship_kind(relationship.reltype()).is_some())
        {
            return Err(invalid(format!(
                "story '{}' cannot own another Word story",
                story.part
            )));
        }
    }

    for part in package.iter_parts() {
        if kind_from_content_type(part.content_type()).is_some() && !owned.contains(part.partname())
        {
            return Err(invalid(format!(
                "Word story part '{}' is orphaned",
                part.partname()
            )));
        }
    }

    stories.sort_unstable_by(|left, right| left.part.as_str().cmp(right.part.as_str()));
    let main_index = stories
        .iter()
        .position(|story| story.part == main)
        .ok_or_else(|| invalid("resolved main document disappeared during story capture"))?;
    stories.swap(0, main_index);
    let topology = encode_topology(&stories, limits.max_topology_bytes)?;
    Ok(StoryInventory {
        main,
        stories,
        topology,
        total_story_bytes,
        dialect,
    })
}

#[allow(
    clippy::too_many_arguments,
    reason = "signature mirrors the corresponding OOXML record"
)]
fn discover_owned(
    package: &OpcPackage,
    owner: &PackURI,
    allow_glossary: bool,
    dialect: StoryDialect,
    limits: StoryLimits,
    enforce_story_bytes: bool,
    owned: &mut HashSet<PackURI>,
    stories: &mut Vec<StoryPart>,
    relationship_count: &mut usize,
    total_story_bytes: &mut usize,
) -> Result<Option<PackURI>> {
    let owner_part = package.get_part(owner)?;
    let mut relationships = Vec::new();
    for relationship in owner_part.rels().iter() {
        let Some((edge_dialect, kind)) = relationship_kind(relationship.reltype()) else {
            continue;
        };
        if edge_dialect != dialect {
            return Err(invalid(format!(
                "story '{owner}' mixes Strict and Transitional relationship dialects"
            )));
        }
        if kind == StoryKind::Glossary && !allow_glossary {
            return Err(invalid(
                "a glossary story cannot own another glossary story",
            ));
        }
        relationships
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "DOCX story ownership relationships",
                source,
            })?;
        relationships.push((kind, relationship));
        if relationships.len() > limits.max_relationships_per_owner {
            return Err(invalid(format!(
                "story '{}' relationship count exceeds {}",
                owner, limits.max_relationships_per_owner
            )));
        }
    }
    relationships.sort_unstable_by(|left, right| {
        left.1
            .r_id()
            .cmp(right.1.r_id())
            .then_with(|| left.1.reltype().cmp(right.1.reltype()))
            .then_with(|| left.1.target_ref().cmp(right.1.target_ref()))
    });

    let mut singleton = [false; 4];
    let mut glossary = None;
    for (kind, relationship) in relationships {
        if let Some(index) = kind.singleton_index()
            && std::mem::replace(&mut singleton[index], true)
        {
            return Err(invalid(format!(
                "story '{owner}' has multiple {kind:?} relationships"
            )));
        }
        *relationship_count = relationship_count
            .checked_add(1)
            .ok_or_else(|| invalid("DOCX story relationship count overflow"))?;
        if *relationship_count > limits.max_total_relationships {
            return Err(invalid(format!(
                "DOCX story relationship count exceeds {}",
                limits.max_total_relationships
            )));
        }
        if relationship.is_external() {
            return Err(invalid(format!(
                "story '{owner}' has an external {kind:?} relationship"
            )));
        }
        let requested = relationship.target_partname()?;
        let target_part = package.get_part(&requested).map_err(|_source_error| {
            invalid(format!(
                "story '{owner}' {kind:?} relationship targets missing part '{requested}'"
            ))
        })?;
        let expected = kind
            .content_type()
            .ok_or_else(|| invalid("main story cannot be owned by another story"))?;
        if target_part.content_type() != expected {
            return Err(invalid(format!(
                "story '{}' {:?} relationship targets content type '{}'",
                owner,
                kind,
                target_part.content_type()
            )));
        }
        let target = target_part.partname().clone();
        if owned.contains(&target) {
            return Err(invalid(format!(
                "Word story part '{target}' has ambiguous ownership"
            )));
        }
        let incoming = StoryOwner {
            part: Some(owner.clone()),
            relationship_id: relationship.r_id().to_owned(),
            relationship_type: relationship.reltype().to_owned(),
            target_ref: relationship.target_ref().to_owned(),
        };
        push_story(
            stories,
            owned,
            target_part,
            kind,
            incoming,
            dialect,
            limits,
            enforce_story_bytes,
            total_story_bytes,
        )?;
        if kind == StoryKind::Glossary {
            glossary = Some(target);
        }
    }
    Ok(glossary)
}

#[allow(
    clippy::too_many_arguments,
    reason = "signature mirrors the corresponding OOXML record"
)]
fn push_story(
    stories: &mut Vec<StoryPart>,
    owned: &mut HashSet<PackURI>,
    part: &dyn litchi_opc::Part,
    kind: StoryKind,
    owner: StoryOwner,
    dialect: StoryDialect,
    limits: StoryLimits,
    enforce_story_bytes: bool,
    total_story_bytes: &mut usize,
) -> Result<()> {
    if stories.len() == limits.max_stories {
        return Err(invalid(format!(
            "DOCX story count exceeds {}",
            limits.max_stories
        )));
    }
    let length = part.blob().len();
    if enforce_story_bytes && length > limits.max_story_bytes {
        return Err(invalid(format!(
            "DOCX story '{}' has {length} bytes, exceeding {}",
            part.partname(),
            limits.max_story_bytes
        )));
    }
    *total_story_bytes = total_story_bytes
        .checked_add(length)
        .ok_or_else(|| invalid("DOCX aggregate story byte count overflow"))?;
    if enforce_story_bytes && *total_story_bytes > limits.max_total_story_bytes {
        return Err(invalid(format!(
            "DOCX aggregate story bytes exceed {}",
            limits.max_total_story_bytes
        )));
    }
    // Main-only semantic readers still need the complete ownership topology
    // for stale-package detection, but must not parse unrelated story XML.
    // Strict inventories and all multi-story transactions pass
    // `enforce_story_bytes = true`, so every reachable story root remains
    // validated before it can participate in a mutation.
    if enforce_story_bytes || kind == StoryKind::Main {
        validate_root(part.blob(), kind, dialect, limits.max_xml_prolog_events)?;
    }
    owned.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "DOCX story ownership set",
        source,
    })?;
    stories.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "DOCX story inventory",
        source,
    })?;
    owned.insert(part.partname().clone());
    stories.push(StoryPart {
        part: part.partname().clone(),
        content_type: part.content_type().to_owned(),
        source: part.blob_arc(),
        kind,
        owner,
    });
    Ok(())
}

fn root_owner(package: &OpcPackage) -> Result<StoryOwner> {
    let mut relationships = package.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            rt::OFFICE_DOCUMENT | rt::STRICT_OFFICE_DOCUMENT
        )
    });
    let relationship = relationships
        .next()
        .ok_or_else(|| invalid("main-document relationship is missing"))?;
    if relationships.next().is_some() {
        return Err(invalid("package has multiple main-document relationships"));
    }
    if relationship.is_external() {
        return Err(invalid("main-document relationship cannot be external"));
    }
    Ok(StoryOwner {
        part: None,
        relationship_id: relationship.r_id().to_owned(),
        relationship_type: relationship.reltype().to_owned(),
        target_ref: relationship.target_ref().to_owned(),
    })
}

fn relationship_dialect(value: &str) -> Option<StoryDialect> {
    if value.starts_with(TRANSITIONAL_RELATIONSHIPS) {
        Some(StoryDialect::Transitional)
    } else if value.starts_with(STRICT_RELATIONSHIPS) {
        Some(StoryDialect::Strict)
    } else {
        None
    }
}

fn relationship_kind(value: &str) -> Option<(StoryDialect, StoryKind)> {
    let dialect = relationship_dialect(value)?;
    let suffix = value.strip_prefix(dialect.relationships())?;
    let kind = match suffix {
        "header" => StoryKind::Header,
        "footer" => StoryKind::Footer,
        "footnotes" => StoryKind::Footnotes,
        "endnotes" => StoryKind::Endnotes,
        "comments" => StoryKind::Comments,
        "glossaryDocument" => StoryKind::Glossary,
        _ => return None,
    };
    Some((dialect, kind))
}

fn kind_from_content_type(value: &str) -> Option<StoryKind> {
    match value {
        ct::WML_DOCUMENT_MAIN
        | ct::WML_TEMPLATE_MAIN
        | ct::WML_DOCUMENT_MACRO_MAIN
        | ct::WML_TEMPLATE_MACRO_MAIN => Some(StoryKind::Main),
        ct::WML_HEADER => Some(StoryKind::Header),
        ct::WML_FOOTER => Some(StoryKind::Footer),
        ct::WML_FOOTNOTES => Some(StoryKind::Footnotes),
        ct::WML_ENDNOTES => Some(StoryKind::Endnotes),
        ct::WML_COMMENTS => Some(StoryKind::Comments),
        ct::WML_DOCUMENT_GLOSSARY => Some(StoryKind::Glossary),
        _ => None,
    }
}

fn validate_main_content_type(content_type: &str) -> Result<()> {
    if matches!(
        content_type,
        ct::WML_DOCUMENT_MAIN
            | ct::WML_TEMPLATE_MAIN
            | ct::WML_DOCUMENT_MACRO_MAIN
            | ct::WML_TEMPLATE_MACRO_MAIN
    ) {
        Ok(())
    } else {
        Err(invalid(format!(
            "main document has unsupported content type '{content_type}'"
        )))
    }
}

fn validate_root(
    xml: &[u8],
    kind: StoryKind,
    dialect: StoryDialect,
    max_events: usize,
) -> Result<()> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_comments = true;
    let mut events = 0usize;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("Word story XML prolog event count overflow"))?;
        if events > max_events {
            return Err(invalid(format!(
                "Word story XML prolog exceeds {max_events} events"
            )));
        }
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let valid_namespace = matches!(
                    namespace,
                    ResolveResult::Bound(Namespace(value)) if value == dialect.word()
                );
                if !valid_namespace || element.local_name().as_ref() != kind.root() {
                    return Err(invalid(format!(
                        "{kind:?} story has an invalid XML root or namespace"
                    )));
                }
                return Ok(());
            },
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) => {},
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            Event::DocType(_) | Event::GeneralRef(_) => {
                return Err(invalid(
                    "DTD and entity references are forbidden in Word stories",
                ));
            },
            Event::Eof => return Err(invalid("Word story XML root is missing")),
            Event::End(_) | Event::Text(_) | Event::CData(_) => {
                return Err(invalid("invalid content before Word story XML root"));
            },
        }
    }
}

fn encode_topology(stories: &[StoryPart], limit: usize) -> Result<Arc<[u8]>> {
    let main = stories
        .first()
        .ok_or_else(|| invalid("Word story inventory is empty"))?;
    let mut size = TOPOLOGY_MAGIC.len();
    for value in [
        main.owner.relationship_id.as_bytes(),
        main.owner.relationship_type.as_bytes(),
        main.owner.target_ref.as_bytes(),
    ] {
        size = add_field_size(size, value.len())?;
    }
    size = add_size(size, 8)?;
    for story in stories {
        size = add_field_size(size, story.part.as_str().len())?;
        size = add_field_size(size, story.content_type.len())?;
    }
    size = add_size(size, 8)?;
    for story in stories.iter().skip(1) {
        let owner = story
            .owner
            .part
            .as_ref()
            .ok_or_else(|| invalid("subsidiary Word story is missing its owning part"))?;
        for length in [
            owner.as_str().len(),
            story.owner.relationship_id.len(),
            story.owner.relationship_type.len(),
            story.owner.target_ref.len(),
            story.part.as_str().len(),
        ] {
            size = add_field_size(size, length)?;
        }
        size = add_size(size, 1)?;
    }
    if size > limit {
        return Err(invalid(format!(
            "DOCX story topology exceeds {limit} bytes"
        )));
    }
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(size)
        .map_err(|source| Error::Allocation {
            resource: "DOCX story topology",
            source,
        })?;
    bytes.extend_from_slice(TOPOLOGY_MAGIC);
    put_field(&mut bytes, main.owner.relationship_id.as_bytes());
    put_field(&mut bytes, main.owner.relationship_type.as_bytes());
    put_field(&mut bytes, main.owner.target_ref.as_bytes());
    put_number(&mut bytes, stories.len());
    for story in stories {
        put_field(&mut bytes, story.part.as_str().as_bytes());
        put_field(&mut bytes, story.content_type.as_bytes());
    }
    put_number(&mut bytes, stories.len().saturating_sub(1));
    let mut edges: Vec<_> = stories.iter().skip(1).collect();
    edges.sort_unstable_by(|left, right| {
        left.owner
            .part
            .as_ref()
            .map(PackURI::as_str)
            .cmp(&right.owner.part.as_ref().map(PackURI::as_str))
            .then_with(|| left.owner.relationship_id.cmp(&right.owner.relationship_id))
            .then_with(|| {
                left.owner
                    .relationship_type
                    .cmp(&right.owner.relationship_type)
            })
            .then_with(|| left.owner.target_ref.cmp(&right.owner.target_ref))
    });
    for story in edges {
        let owner = story.owner.part.as_ref().expect("validated owner");
        put_field(&mut bytes, owner.as_str().as_bytes());
        put_field(&mut bytes, story.owner.relationship_id.as_bytes());
        put_field(&mut bytes, story.owner.relationship_type.as_bytes());
        put_field(&mut bytes, story.owner.target_ref.as_bytes());
        put_field(&mut bytes, story.part.as_str().as_bytes());
        bytes.push(story.kind.code());
    }
    debug_assert_eq!(bytes.len(), size);
    Ok(Arc::from(bytes.into_boxed_slice()))
}

fn add_field_size(size: usize, length: usize) -> Result<usize> {
    add_size(add_size(size, 8)?, length)
}

fn add_size(size: usize, added: usize) -> Result<usize> {
    size.checked_add(added)
        .ok_or_else(|| invalid("DOCX story topology size overflow"))
}

fn put_number(output: &mut Vec<u8>, value: usize) {
    output.extend_from_slice(&(value as u64).to_le_bytes());
}

fn put_field(output: &mut Vec<u8>, value: &[u8]) {
    put_number(output, value.len());
    output.extend_from_slice(value);
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::BlobPart;

    fn xml(kind: StoryKind, dialect: StoryDialect) -> Vec<u8> {
        format!(
            r#"<w:{} xmlns:w="{}"/>"#,
            std::str::from_utf8(kind.root()).unwrap(),
            std::str::from_utf8(dialect.word()).unwrap()
        )
        .into_bytes()
    }

    fn add(
        package: &mut Package,
        owner: &PackURI,
        target: &str,
        target_ref: &str,
        kind: StoryKind,
        id: &str,
        dialect: StoryDialect,
    ) {
        let uri = PackURI::new(target).unwrap();
        package.opc.add_part(Box::new(BlobPart::new(
            uri,
            kind.content_type().unwrap().to_owned(),
            xml(kind, dialect),
        )));
        package
            .opc
            .get_part_mut(owner)
            .unwrap()
            .rels_mut()
            .add_relationship(
                format!(
                    "{}{suffix}",
                    dialect.relationships(),
                    suffix = match kind {
                        StoryKind::Header => "header",
                        StoryKind::Footer => "footer",
                        StoryKind::Footnotes => "footnotes",
                        StoryKind::Endnotes => "endnotes",
                        StoryKind::Comments => "comments",
                        StoryKind::Glossary => "glossaryDocument",
                        StoryKind::Main => unreachable!(),
                    }
                ),
                target_ref.to_owned(),
                id.to_owned(),
                false,
            );
    }

    #[test]
    fn inventories_main_and_dual_owner_stories_without_copying_sources() {
        let mut package = Package::new().unwrap();
        let main = package.opc.main_document_part().unwrap().partname().clone();
        add(
            &mut package,
            &main,
            "/word/header1.xml",
            "header1.xml",
            StoryKind::Header,
            "rStory1",
            StoryDialect::Transitional,
        );
        add(
            &mut package,
            &main,
            "/word/glossary/document.xml",
            "glossary/document.xml",
            StoryKind::Glossary,
            "rStory2",
            StoryDialect::Transitional,
        );
        let glossary = PackURI::new("/word/glossary/document.xml").unwrap();
        add(
            &mut package,
            &glossary,
            "/word/glossary/comments.xml",
            "comments.xml",
            StoryKind::Comments,
            "rStory3",
            StoryDialect::Transitional,
        );

        let source = package
            .opc
            .get_part(&PackURI::new("/word/header1.xml").unwrap())
            .unwrap()
            .blob_arc();
        let inventory = capture(&package.opc, StoryLimits::default()).unwrap();
        assert_eq!(inventory.stories().len(), 4);
        assert_eq!(inventory.stories()[0].kind(), StoryKind::Main);
        let header = inventory
            .get(&PackURI::new("/word/header1.xml").unwrap())
            .unwrap();
        assert_eq!(header.owner().part(), Some(&main));
        assert!(Arc::ptr_eq(&source, &header.source_arc()));
    }

    #[test]
    fn rejects_orphan_and_ambiguous_story_ownership() {
        let mut orphan = Package::new().unwrap();
        orphan.opc.add_part(Box::new(BlobPart::new(
            PackURI::new("/word/comments-orphan.xml").unwrap(),
            ct::WML_COMMENTS.to_owned(),
            xml(StoryKind::Comments, StoryDialect::Transitional),
        )));
        assert!(capture(&orphan.opc, StoryLimits::default()).is_err());

        let mut shared = Package::new().unwrap();
        let main = shared.opc.main_document_part().unwrap().partname().clone();
        add(
            &mut shared,
            &main,
            "/word/glossary/document.xml",
            "glossary/document.xml",
            StoryKind::Glossary,
            "rStory1",
            StoryDialect::Transitional,
        );
        add(
            &mut shared,
            &main,
            "/word/comments.xml",
            "comments.xml",
            StoryKind::Comments,
            "rStory2",
            StoryDialect::Transitional,
        );
        let glossary = PackURI::new("/word/glossary/document.xml").unwrap();
        shared
            .opc
            .get_part_mut(&glossary)
            .unwrap()
            .rels_mut()
            .add_relationship(
                format!("{}comments", TRANSITIONAL_RELATIONSHIPS),
                "../comments.xml".to_owned(),
                "rStory3".to_owned(),
                false,
            );
        assert!(capture(&shared.opc, StoryLimits::default()).is_err());
    }

    #[test]
    fn rejects_mixed_dialect_and_external_story_edges() {
        let mut mixed = Package::new().unwrap();
        let main = mixed.opc.main_document_part().unwrap().partname().clone();
        add(
            &mut mixed,
            &main,
            "/word/header1.xml",
            "header1.xml",
            StoryKind::Header,
            "rStory1",
            StoryDialect::Strict,
        );
        assert!(capture(&mixed.opc, StoryLimits::default()).is_err());

        let mut external = Package::new().unwrap();
        let main = external
            .opc
            .main_document_part()
            .unwrap()
            .partname()
            .clone();
        external
            .opc
            .get_part_mut(&main)
            .unwrap()
            .rels_mut()
            .add_relationship(
                format!("{}comments", TRANSITIONAL_RELATIONSHIPS),
                "https://example.invalid/comments.xml".to_owned(),
                "rStory1".to_owned(),
                true,
            );
        assert!(capture(&external.opc, StoryLimits::default()).is_err());
    }

    #[test]
    fn enforces_semantic_story_limits() {
        let package = Package::new().unwrap();
        let limits = StoryLimits {
            max_story_bytes: 1,
            ..StoryLimits::default()
        };
        assert!(capture(&package.opc, limits).is_err());
    }
}
