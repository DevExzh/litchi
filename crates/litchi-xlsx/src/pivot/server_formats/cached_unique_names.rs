//! Source-preserving diagnostic support for the `x15:cachedUniqueNames`
//! cache-field extension.
//!
//! This module intentionally lives below the source-bound pivot extension
//! owner so it can share the bounded XML scanner, relationship index, and
//! source-part capture code.  It does not use the server-format graph: the
//! cache-field extension has a different owner and does not imply the C510
//! non-worksheet closure.

use std::collections::{HashMap, HashSet};
use std::ops::Range;
use std::sync::Arc;

use super::*;
use quick_xml::encoding::Decoder;
use quick_xml::events::BytesStart;
use quick_xml::name::QName;

const CACHED_UNIQUE_NAMES_URI: &str = "{4F2E5C28-24EA-4EB8-9CBF-B6C8F9C3D259}";
const CACHED_UNIQUE_NAMES_PAYLOAD: &[u8] = b"cachedUniqueNames";
const CACHED_UNIQUE_NAME_PAYLOAD: &[u8] = b"cachedUniqueName";
const CACHE_FIELD_PAYLOAD: &[u8] = b"cacheField";
const CACHE_FIELDS_PAYLOAD: &[u8] = b"cacheFields";
const CACHE_SOURCE_PAYLOAD: &[u8] = b"cacheSource";
const PIVOT_CACHE_PAYLOAD: &[u8] = b"pivotCache";
const PIVOT_CACHES_PAYLOAD: &[u8] = b"pivotCaches";
const WORKBOOK_PAYLOAD: &str = "workbook";
const PIVOT_CACHE_DEFINITION_PAYLOAD: &str = "pivotCacheDefinition";
const CONNECTION_ID_ATTRIBUTE: &[u8] = b"connectionId";
const TYPE_ATTRIBUTE: &[u8] = b"type";
const CACHE_ID_ATTRIBUTE: &[u8] = b"cacheId";
const NAME_ATTRIBUTE: &[u8] = b"name";
const INDEX_ATTRIBUTE: &[u8] = b"index";
const MAX_UNIQUE_NAME_UTF16: usize = 65_535;
const MAX_CACHE_FIELDS: usize = (1usize << 31) - 1;
const MAX_UNIQUE_NAMES: usize = (1usize << 31) - 1;

/// The small amount of MCE branch metadata needed by this owner.  It is
/// recorded by the shared XML scan so branch selection never performs a
/// second whole-Part pass.  A source selected through a branch remains
/// read-only for source-preserving edits because the original owner span does
/// not identify a projected MCE branch.
#[derive(Debug)]
pub(crate) enum MceBranch {
    Choice(Vec<NamespaceRef>),
    Fallback,
}

/// Parse one MCE branch while the shared namespace resolver still has the
/// complete in-scope declaration set.  The cache-field reader later compares
/// the expanded `Requires` namespaces against its narrowly supported profile.
pub(super) fn scan_mce_branch(
    namespace: &[u8],
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: Decoder,
    limits: XmlScanLimits,
    ignorable_scope: usize,
    ignorable_scopes: &[IgnorableScope],
) -> Result<Option<MceBranch>> {
    if namespace != MCE_NS {
        return Ok(None);
    }
    let local = element.name().local_name();
    match local.as_ref() {
        b"Choice" => {
            let mut requires = None;
            for attribute in element.checked_attributes() {
                let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
                let key = attribute.key.as_ref();
                if key == b"Requires" {
                    if requires.is_some() {
                        return Err(invalid("MCE Choice has duplicate Requires attributes"));
                    }
                    if attribute.value.len() > limits.attribute_bytes {
                        return Err(invalid("MCE Choice Requires exceeds its text limit"));
                    }
                    let decoded = attribute
                        .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
                        .map_err(|error| invalid(error.to_string()))?;
                    if decoded.len() > limits.attribute_bytes {
                        return Err(invalid("MCE Choice Requires exceeds its text limit"));
                    }
                    let mut owned = String::new();
                    owned
                        .try_reserve_exact(decoded.len())
                        .map_err(|source| Error::Allocation {
                            resource: "cached unique-name MCE Requires",
                            source,
                        })?;
                    owned.push_str(&decoded);
                    requires = Some(owned);
                } else if !is_namespace_binding(key)
                    && !qualified_attribute_is_ignorable(
                        attribute.key,
                        resolver,
                        ignorable_scope,
                        ignorable_scopes,
                    )?
                {
                    return Err(invalid("MCE Choice has an unsupported attribute"));
                }
            }
            let requires = requires.ok_or_else(|| invalid("MCE Choice requires Requires"))?;
            if requires.len() > limits.attribute_bytes {
                return Err(invalid("MCE Choice Requires exceeds its text limit"));
            }
            let mut namespaces: Vec<NamespaceRef> = Vec::new();
            for prefix in requires.split([' ', '\t', '\r', '\n']) {
                if prefix.is_empty() {
                    continue;
                }
                if !valid_ncname(prefix.as_bytes()) {
                    return Err(invalid("MCE Choice Requires contains an invalid prefix"));
                }
                let mut qualified = Vec::new();
                let qualified_len = prefix
                    .len()
                    .checked_add(2)
                    .ok_or_else(|| invalid("MCE Choice Requires prefix length overflows"))?;
                qualified
                    .try_reserve_exact(qualified_len)
                    .map_err(|source| Error::Allocation {
                        resource: "cached unique-name MCE requirement",
                        source,
                    })?;
                qualified.extend_from_slice(prefix.as_bytes());
                qualified.extend_from_slice(b":x");
                let resolved = resolver.resolve_element(QName(&qualified))?.0;
                let namespace = match resolved {
                    NamespaceResolution::Bound(value) => value,
                    NamespaceResolution::Unknown(_) => {
                        return Err(invalid("MCE Choice Requires contains an unbound prefix"));
                    },
                };
                if namespace.as_ref() == MCE_NS {
                    return Err(invalid(
                        "MCE Choice Requires cannot contain the MCE namespace",
                    ));
                }
                if namespaces
                    .iter()
                    .any(|known| known.as_ref() == namespace.as_ref())
                {
                    continue;
                }
                if namespaces.len() >= limits.namespace_declarations {
                    return Err(invalid("MCE Choice Requires namespace count exceeds limit"));
                }
                namespaces
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "cached unique-name MCE requirements",
                        source,
                    })?;
                namespaces.push(namespace);
            }
            if namespaces.is_empty() {
                return Err(invalid("MCE Choice Requires must contain a prefix"));
            }
            Ok(Some(MceBranch::Choice(namespaces)))
        },
        b"Fallback" => {
            for attribute in element.checked_attributes() {
                let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
                if !is_namespace_binding(attribute.key.as_ref())
                    && !qualified_attribute_is_ignorable(
                        attribute.key,
                        resolver,
                        ignorable_scope,
                        ignorable_scopes,
                    )?
                {
                    return Err(invalid("MCE Fallback has an unsupported attribute"));
                }
            }
            Ok(Some(MceBranch::Fallback))
        },
        _ => Ok(None),
    }
}

fn qualified_attribute_is_ignorable(
    key: QName<'_>,
    resolver: &NamespaceResolver,
    ignorable_scope: usize,
    ignorable_scopes: &[IgnorableScope],
) -> Result<bool> {
    let (resolved, _) = resolver.resolve_attribute(key)?;
    let NamespaceResolution::Bound(namespace) = resolved else {
        return Ok(false);
    };
    if namespace.as_ref() == XML_NAMESPACE_URI {
        return Ok(false);
    }
    Ok(namespace.as_ref() == MCE_NS
        || scope_contains(ignorable_scopes, ignorable_scope, namespace.as_ref()))
}

pub(super) fn validate_mce_alternate_content_attributes(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    ignorable_scope: usize,
    ignorable_scopes: &[IgnorableScope],
) -> Result<()> {
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        if is_namespace_binding(attribute.key.as_ref()) {
            continue;
        }
        let (resolved, _) = resolver.resolve_attribute(attribute.key)?;
        let NamespaceResolution::Bound(namespace) = resolved else {
            return Err(invalid(
                "MCE AlternateContent has an unqualified or unbound attribute",
            ));
        };
        if namespace.as_ref() == XML_NAMESPACE_URI {
            return Err(invalid(
                "MCE AlternateContent cannot carry XML namespace attributes",
            ));
        }
        if namespace.as_ref() != MCE_NS
            && !scope_contains(ignorable_scopes, ignorable_scope, namespace.as_ref())
        {
            return Err(invalid(
                "MCE AlternateContent has a non-ignorable qualified attribute",
            ));
        }
    }
    Ok(())
}

fn is_namespace_binding(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn valid_ncname(value: &[u8]) -> bool {
    std::str::from_utf8(value)
        .ok()
        .is_some_and(litchi_ooxml_common::xml_name::is_ncname)
}

/// Semantic workbook PivotCache identity.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct PivotCacheId(pub u32);

impl From<u32> for PivotCacheId {
    fn from(value: u32) -> Self {
        Self(value)
    }
}

/// Ordinary semantic selector for one workbook PivotCache.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotCacheSelector {
    /// Select the cache by its workbook-level semantic cache ID.
    Id(PivotCacheId),
    /// Select the zero-based workbook cache-reference position.
    Position(usize),
}

impl From<PivotCacheId> for PivotCacheSelector {
    fn from(value: PivotCacheId) -> Self {
        Self::Id(value)
    }
}

impl From<u32> for PivotCacheSelector {
    fn from(value: u32) -> Self {
        Self::Id(PivotCacheId(value))
    }
}

impl From<usize> for PivotCacheSelector {
    fn from(value: usize) -> Self {
        Self::Position(value)
    }
}

/// Ordinary semantic selector for one cache field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotCacheFieldSelector<'a> {
    /// Select a field by its zero-based `cacheFields/cacheField` ordinal.
    Ordinal(usize),
    /// Select a field by its decoded semantic name.
    Name(&'a str),
}

impl<'a> From<&'a str> for PivotCacheFieldSelector<'a> {
    fn from(value: &'a str) -> Self {
        Self::Name(value)
    }
}

impl From<usize> for PivotCacheFieldSelector<'static> {
    fn from(value: usize) -> Self {
        Self::Ordinal(value)
    }
}

/// The index-to-item bound is deliberately unresolved for this batch.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum IndexBoundStatus {
    /// The local schema exposes `sharedItems`, while MS-XLSX names
    /// `CT_Items@count`; no normative substitution is made.
    Unresolved,
}

/// Canonical advanced-layer cache selector name.
pub type CacheSelector = PivotCacheSelector;

/// Canonical advanced-layer field selector name.
pub type FieldSelector<'a> = PivotCacheFieldSelector<'a>;

/// Canonical advanced-layer diagnostic status name.
pub type DiagnosticStatus = IndexBoundStatus;

/// One typed `x15:cachedUniqueName` leaf.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CachedUniqueName {
    /// The unique cache-item index carried by the leaf.
    pub index: u32,
    /// The decoded SpreadsheetML `ST_Xstring` name.
    pub name: String,
}

/// Typed diagnostic view of one cache-field extension collection.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CachedUniqueNames {
    cache_id: PivotCacheId,
    field_ordinal: usize,
    field_name: String,
    names: Box<[CachedUniqueName]>,
    index_bound: IndexBoundStatus,
}

impl CachedUniqueNames {
    /// The semantic workbook cache ID.
    #[must_use]
    pub const fn cache_id(&self) -> PivotCacheId {
        self.cache_id
    }

    /// The selected cache-field ordinal.
    #[must_use]
    pub const fn field_ordinal(&self) -> usize {
        self.field_ordinal
    }

    /// The decoded semantic cache-field name.
    #[must_use]
    pub fn field_name(&self) -> &str {
        &self.field_name
    }

    /// The bounded typed list of cached unique names.
    #[must_use]
    pub fn names(&self) -> &[CachedUniqueName] {
        &self.names
    }

    /// The status of the unresolved cache-item count dependency.
    #[must_use]
    pub const fn index_bound_status(&self) -> IndexBoundStatus {
        self.index_bound
    }

    /// Canonical advanced-layer diagnostic status accessor.
    #[must_use]
    pub const fn diagnostic_status(&self) -> DiagnosticStatus {
        self.index_bound_status()
    }

    /// Canonical advanced-layer entry accessor.
    #[must_use]
    pub fn entries(&self) -> &[CachedUniqueName] {
        self.names()
    }
}

/// A workbook-facing cache handle.  Resolution is deferred until a typed
/// field view is read, so ordinary callers never supply a Part or relationship
/// identity.
#[derive(Clone, Debug)]
pub struct PivotCacheHandle {
    workbook: Workbook,
    selector: PivotCacheSelector,
}

impl PivotCacheHandle {
    /// Select one cache field for a diagnostic read.
    #[must_use]
    pub fn field<'a>(self, selector: PivotCacheFieldSelector<'a>) -> PivotCacheFieldHandle<'a> {
        PivotCacheFieldHandle {
            workbook: self.workbook,
            cache_selector: self.selector,
            field_selector: selector,
        }
    }
}

/// A workbook-facing cache-field handle.
#[derive(Clone, Debug)]
pub struct PivotCacheFieldHandle<'a> {
    workbook: Workbook,
    cache_selector: PivotCacheSelector,
    field_selector: PivotCacheFieldSelector<'a>,
}

impl<'a> PivotCacheFieldHandle<'a> {
    /// Read the typed source-backed cached-name collection.
    pub fn cached_unique_names(&self) -> Result<CachedUniqueNames> {
        Snapshot::load(
            self.workbook.pivot_package(),
            self.cache_selector,
            self.field_selector,
        )
        .map(|snapshot| snapshot.cached_unique_names().clone())
    }

    /// Read the low-level source-backed snapshot for diagnostics and patches.
    pub fn cached_unique_names_source(&self) -> Result<Snapshot> {
        Snapshot::load(
            self.workbook.pivot_package(),
            self.cache_selector,
            self.field_selector,
        )
    }
}

/// A source-bound cache-field snapshot retaining exact owner spans and the
/// relationship/readset closure needed for reversible edits.
#[derive(Clone, Debug)]
pub struct Snapshot {
    value: CachedUniqueNames,
    cache: SourcePart,
    connections: Option<SourcePart>,
    workbook: SourcePart,
    workbook_owner: Arc<Vec<u8>>,
    workbook_context: Arc<Vec<u8>>,
    owner: Range<usize>,
    entries: Box<[EntrySource]>,
    mce_ambiguous: bool,
    selection: usize,
}

impl Snapshot {
    /// Resolve and read one semantic cache field.
    pub fn load<'a>(
        package: &OpcPackage,
        cache_selector: impl Into<PivotCacheSelector>,
        field_selector: PivotCacheFieldSelector<'a>,
    ) -> Result<Self> {
        let graph = Graph::load(package)?;
        graph.snapshot(cache_selector.into(), field_selector)
    }

    /// Alias emphasizing that this state is tied to exact source bytes.
    pub fn read<'a>(
        package: &OpcPackage,
        cache_selector: impl Into<PivotCacheSelector>,
        field_selector: PivotCacheFieldSelector<'a>,
    ) -> Result<Self> {
        Self::load(package, cache_selector, field_selector)
    }

    /// The typed diagnostic value.
    #[must_use]
    pub fn cached_unique_names(&self) -> &CachedUniqueNames {
        &self.value
    }

    /// Alias for callers using the element name.
    #[must_use]
    pub fn names(&self) -> &[CachedUniqueName] {
        self.value.names()
    }

    /// The semantic cache ID.
    #[must_use]
    pub const fn cache_id(&self) -> PivotCacheId {
        self.value.cache_id()
    }

    /// The selected field ordinal.
    #[must_use]
    pub const fn field_ordinal(&self) -> usize {
        self.value.field_ordinal()
    }

    /// The unresolved index-bound status.
    #[must_use]
    pub const fn index_bound_status(&self) -> IndexBoundStatus {
        self.value.index_bound_status()
    }

    /// Canonical advanced-layer diagnostic status accessor.
    #[must_use]
    pub const fn diagnostic_status(&self) -> DiagnosticStatus {
        self.value.diagnostic_status()
    }

    /// Canonical advanced-layer entry accessor.
    #[must_use]
    pub fn entries(&self) -> &[CachedUniqueName] {
        self.names()
    }

    /// Whether the recognized payload was selected through an MCE branch.
    #[must_use]
    pub const fn has_ambiguous_mce_owner(&self) -> bool {
        self.mce_ambiguous
    }

    /// Exact source XML for the cache-definition Part.
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.cache.bytes.as_slice()
    }

    /// Share the exact cache-definition source allocation.
    #[must_use]
    pub fn source_arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.cache.bytes)
    }

    /// Low-level physical cache Part identity.
    #[must_use]
    pub fn cache_part(&self) -> &PackURI {
        &self.cache.name
    }

    /// Exact extension-owner range in the cache-definition source.
    #[must_use]
    pub fn source_owner_range(&self) -> Range<usize> {
        self.owner.clone()
    }

    fn same_source(&self, other: &Self) -> bool {
        self.selection == other.selection
            && self.value.cache_id == other.value.cache_id
            && self.value.field_ordinal == other.value.field_ordinal
            && self.cache.same_source(&other.cache)
            && same_optional_source(&self.connections, &other.connections)
            && self.workbook.same_source(&other.workbook)
    }

    fn same_readset(&self, other: &Self) -> bool {
        self.selection == other.selection
            && self.value.cache_id == other.value.cache_id
            && self.value.field_ordinal == other.value.field_ordinal
            && self.cache.same_source(&other.cache)
            && same_optional_source(&self.connections, &other.connections)
            && self.workbook.same_closure(&other.workbook)
            && self.workbook_owner.as_slice() == other.workbook_owner.as_slice()
            && self.workbook_context.as_slice() == other.workbook_context.as_slice()
    }

    fn same_closure(&self, other: &Self) -> bool {
        self.cache.same_closure(&other.cache)
            && same_optional_closure(&self.connections, &other.connections)
            && self.workbook.same_closure(&other.workbook)
    }
}

#[derive(Clone, Debug)]
struct EntrySource {
    name: AttributeSource,
}

#[derive(Clone, Debug)]
struct AttributeSource {
    value: Range<usize>,
}

/// Failure-atomic existing-name edits.  The list length and every index are
/// immutable in this batch; no structural index operation is exposed.
pub struct Transaction<'a> {
    target: &'a mut OpcPackage,
    before: Snapshot,
    staged: Vec<String>,
    selection: usize,
}

impl<'a> Transaction<'a> {
    /// Resolve a cache field and begin a source-bound transaction.
    pub fn new<'selector>(
        target: &'a mut OpcPackage,
        cache_selector: impl Into<PivotCacheSelector>,
        field_selector: PivotCacheFieldSelector<'selector>,
    ) -> Result<Self> {
        let before = Snapshot::load(target, cache_selector, field_selector)?;
        let mut staged = Vec::new();
        staged
            .try_reserve_exact(before.names().len())
            .map_err(|source| Error::Allocation {
                resource: "cached unique-name staged values",
                source,
            })?;
        for name in before.names() {
            let mut copy = String::new();
            copy.try_reserve_exact(name.name.len())
                .map_err(|source| Error::Allocation {
                    resource: "cached unique-name staged text",
                    source,
                })?;
            copy.push_str(&name.name);
            staged.push(copy);
        }
        Self::from_staged(target, before, staged)
    }

    /// Reuse an already validated source snapshot and staged values.  The
    /// ordinary Workbook facade uses this after cloning its package so the
    /// commit path does not reparse and copy the same staged names merely to
    /// overwrite them immediately afterward.
    fn from_staged(
        target: &'a mut OpcPackage,
        before: Snapshot,
        staged: Vec<String>,
    ) -> Result<Self> {
        if staged.len() != before.names().len() {
            return Err(invalid(
                "cached unique-name staged values do not match source length",
            ));
        }
        for value in &staged {
            validate_name_value(value, target.read_limits())?;
        }
        Ok(Self {
            selection: before.selection,
            target,
            before,
            staged,
        })
    }

    /// Typed state captured at transaction start.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Currently staged names in source order.
    #[must_use]
    pub fn names(&self) -> &[String] {
        &self.staged
    }

    /// Replace one existing name by its unique `index` attribute.
    pub fn set_cached_unique_name(
        &mut self,
        item_index: u32,
        value: impl AsRef<str>,
    ) -> Result<bool> {
        let value = value.as_ref();
        validate_name_value(value, self.target.read_limits())?;
        let position = self
            .before
            .names()
            .iter()
            .position(|entry| entry.index == item_index)
            .ok_or_else(|| invalid("cached unique-name index is out of range"))?;
        let slot = self
            .staged
            .get_mut(position)
            .ok_or_else(|| invalid("cached unique-name source order is incomplete"))?;
        if *slot == value {
            return Ok(false);
        }
        let mut replacement = String::new();
        replacement
            .try_reserve_exact(value.len())
            .map_err(|source| Error::Allocation {
                resource: "cached unique-name staged replacement",
                source,
            })?;
        replacement.push_str(value);
        *slot = replacement;
        Ok(true)
    }

    /// Structural list changes are outside the unresolved index-bound scope.
    pub fn insert_cached_unique_name(
        &mut self,
        _index: usize,
        _value: CachedUniqueName,
    ) -> Result<()> {
        Err(invalid(
            "cachedUniqueNames structural insertion requires the unresolved CT_Items bound",
        ))
    }

    /// Structural list changes are outside the unresolved index-bound scope.
    pub fn remove_cached_unique_name(&mut self, _item_index: u32) -> Result<CachedUniqueName> {
        Err(invalid(
            "cachedUniqueNames structural removal requires the unresolved CT_Items bound",
        ))
    }

    /// Structural list changes are outside the unresolved index-bound scope.
    pub fn set_cached_unique_name_index(&mut self, _old: u32, _new: u32) -> Result<()> {
        Err(invalid(
            "cachedUniqueNames index edits require the unresolved CT_Items bound",
        ))
    }

    /// Whether a source name differs from the staged value.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.before
            .names()
            .iter()
            .zip(&self.staged)
            .any(|(before, after)| before.name != *after)
    }

    /// Validate, rewrite, reopen, and publish the scalar edit.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, false));
        }
        ensure_editable(&self.before)?;
        if self.target.is_signed() || self.target.requires_signature_edit_policy() {
            return Err(Error::Signed);
        }
        let current = Snapshot::load(
            self.target,
            PivotCacheSelector::Position(self.selection),
            PivotCacheFieldSelector::Ordinal(self.before.field_ordinal()),
        )?;
        if !current.same_source(&self.before) {
            return Err(Error::PatchConflict {
                part: self.before.cache_part().to_string(),
            });
        }
        let current_total = package_total_bytes(self.target)?;
        let output = rewrite_cache(
            &self.before,
            &self.staged,
            caller_part_limit(self.target.read_limits()),
            self.target.read_limits().max_total_part_bytes(),
            caller_attribute_limit(self.target.read_limits()),
            current_total,
            self.target.read_limits(),
        )?;
        let mut candidate = self.target.clone();
        candidate
            .get_part_mut(self.before.cache_part())?
            .set_blob_shared(Arc::new(output));
        validate_candidate_limits(&candidate)?;
        let after = Snapshot::load(
            &candidate,
            PivotCacheSelector::Position(self.selection),
            PivotCacheFieldSelector::Ordinal(self.before.field_ordinal()),
        )?;
        if !after
            .names()
            .iter()
            .map(|entry| entry.name.as_str())
            .eq(self.staged.iter().map(String::as_str))
            || !after.same_closure(&self.before)
        {
            return Err(invalid(
                "cachedUniqueNames publication changed semantic state or source closure",
            ));
        }
        let patch = Patch::new(self.before, after.clone());
        *self.target = candidate;
        Ok(Commit::new(after, patch, true))
    }
}

fn ensure_editable(snapshot: &Snapshot) -> Result<()> {
    if snapshot.index_bound_status() != IndexBoundStatus::Unresolved {
        return Err(invalid("unsupported cachedUniqueNames index-bound state"));
    }
    if snapshot.has_ambiguous_mce_owner() {
        return Err(invalid(
            "cachedUniqueNames selected through an MCE branch is read-only",
        ));
    }
    Ok(())
}

/// An exact reversible source-bound patch.
#[derive(Clone, Debug)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Required source state.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Resulting source state.
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether the patch preserves the exact source bytes.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply after validating the complete source closure.
    pub fn apply(&self, package: &mut OpcPackage) -> Result<()> {
        let current = Snapshot::load(
            package,
            PivotCacheSelector::Position(self.before.selection),
            PivotCacheFieldSelector::Ordinal(self.before.field_ordinal()),
        )?;
        if !current.same_source(&self.before) {
            return Err(Error::PatchConflict {
                part: self.before.cache_part().to_string(),
            });
        }
        if self.is_empty() {
            return Ok(());
        }
        if package.is_signed() || package.requires_signature_edit_policy() {
            return Err(Error::Signed);
        }
        let mut candidate = package.clone();
        candidate
            .get_part_mut(self.after.cache_part())?
            .set_blob_shared(Arc::clone(&self.after.cache.bytes));
        validate_candidate_limits(&candidate)?;
        let resulting = Snapshot::load(
            &candidate,
            PivotCacheSelector::Position(self.after.selection),
            PivotCacheFieldSelector::Ordinal(self.after.field_ordinal()),
        )?;
        if !resulting.same_source(&self.after)
            || resulting.cached_unique_names() != self.after.cached_unique_names()
        {
            return Err(invalid("cachedUniqueNames patch verification failed"));
        }
        *package = candidate;
        Ok(())
    }
}

/// Committed source-bound edit.
#[derive(Clone, Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    fn new(snapshot: Snapshot, patch: Patch, changed: bool) -> Self {
        Self {
            snapshot,
            patch,
            changed,
        }
    }

    /// Resulting snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Exact inverse-capable patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Whether the name value changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }
}

/// Source-bound ordinary Workbook edit.
pub struct WorkbookTransaction {
    source: Workbook,
    before: Snapshot,
    staged: Vec<String>,
}

impl WorkbookTransaction {
    pub(crate) fn new<'selector>(
        source: &Workbook,
        cache_selector: impl Into<PivotCacheSelector>,
        field_selector: PivotCacheFieldSelector<'selector>,
    ) -> Result<Self> {
        let before = Snapshot::load(source.pivot_package(), cache_selector, field_selector)?;
        let mut staged = Vec::new();
        staged
            .try_reserve_exact(before.names().len())
            .map_err(|source| Error::Allocation {
                resource: "Workbook cached unique-name staged values",
                source,
            })?;
        for name in before.names() {
            let mut copy = String::new();
            copy.try_reserve_exact(name.name.len())
                .map_err(|source| Error::Allocation {
                    resource: "Workbook cached unique-name staged text",
                    source,
                })?;
            copy.push_str(&name.name);
            staged.push(copy);
        }
        Ok(Self {
            source: source.clone(),
            before,
            staged,
        })
    }

    /// Typed source state at transaction start.
    #[must_use]
    pub fn before(&self) -> &CachedUniqueNames {
        self.before.cached_unique_names()
    }

    /// Replace one existing unique name.
    pub fn set_cached_unique_name(
        &mut self,
        item_index: u32,
        value: impl AsRef<str>,
    ) -> Result<bool> {
        let value = value.as_ref();
        validate_name_value(value, self.source.pivot_package().read_limits())?;
        let position = self
            .before
            .names()
            .iter()
            .position(|entry| entry.index == item_index)
            .ok_or_else(|| invalid("cached unique-name index is out of range"))?;
        let slot = self
            .staged
            .get_mut(position)
            .ok_or_else(|| invalid("cached unique-name source order is incomplete"))?;
        if *slot == value {
            return Ok(false);
        }
        let mut replacement = String::new();
        replacement
            .try_reserve_exact(value.len())
            .map_err(|source| Error::Allocation {
                resource: "Workbook cached unique-name staged replacement",
                source,
            })?;
        replacement.push_str(value);
        *slot = replacement;
        Ok(true)
    }

    /// Structural operations remain unavailable until the item-count anchor
    /// is resolved.
    pub fn insert_cached_unique_name(
        &mut self,
        _index: usize,
        _value: CachedUniqueName,
    ) -> Result<()> {
        Err(invalid(
            "cachedUniqueNames structural insertion requires the unresolved CT_Items bound",
        ))
    }

    /// Structural operations remain unavailable until the item-count anchor
    /// is resolved.
    pub fn remove_cached_unique_name(&mut self, _item_index: u32) -> Result<CachedUniqueName> {
        Err(invalid(
            "cachedUniqueNames structural removal requires the unresolved CT_Items bound",
        ))
    }

    /// Structural operations remain unavailable until the item-count anchor
    /// is resolved.
    pub fn set_cached_unique_name_index(&mut self, _old: u32, _new: u32) -> Result<()> {
        Err(invalid(
            "cachedUniqueNames index edits require the unresolved CT_Items bound",
        ))
    }

    /// Whether a staged name differs from source.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.before
            .names()
            .iter()
            .zip(&self.staged)
            .any(|(before, after)| before.name != *after)
    }

    /// Commit to a new immutable Workbook snapshot.
    pub fn commit(self) -> Result<WorkbookCommit> {
        let changed = self.is_changed();
        let source = self.source;
        let before = self.before;
        let staged = self.staged;
        if !changed {
            let workbook = source.clone();
            let low_patch = Patch::new(before.clone(), before);
            let patch = WorkbookPatch::new(source, workbook.clone(), low_patch);
            return Ok(WorkbookCommit::new(workbook, patch, false));
        }
        ensure_editable(&before)?;
        let mut candidate = source.pivot_package().clone();
        let transaction = Transaction::from_staged(&mut candidate, before, staged)?;
        let committed = transaction.commit()?;
        let changed = committed.changed();
        let low_patch = committed.patch().clone();
        let workbook = source.adopt_published_package(candidate)?;
        let patch = WorkbookPatch::new(source, workbook.clone(), low_patch);
        Ok(WorkbookCommit::new(workbook, patch, changed))
    }
}

/// Exact reversible ordinary Workbook patch.
#[derive(Clone, Debug)]
pub struct WorkbookPatch {
    before: Workbook,
    after: Workbook,
    source_patch: Patch,
}

impl WorkbookPatch {
    fn new(before: Workbook, after: Workbook, source_patch: Patch) -> Self {
        Self {
            before,
            after,
            source_patch,
        }
    }

    /// Typed state before the edit.
    #[must_use]
    pub fn before(&self) -> &CachedUniqueNames {
        self.source_patch.before().cached_unique_names()
    }

    /// Typed state after the edit.
    #[must_use]
    pub fn after(&self) -> &CachedUniqueNames {
        self.source_patch.after().cached_unique_names()
    }

    /// Whether the patch is a byte-exact no-op.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.source_patch.is_empty()
    }

    /// Exact inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            source_patch: self.source_patch.inverse(),
        }
    }

    /// Apply after ordinary Workbook mutation-policy and source-readset checks.
    pub fn apply(&self, source: &Workbook) -> Result<WorkbookCommit> {
        source.ensure_mutation_allowed("apply_cached_unique_names_patch")?;
        let mut candidate = source.pivot_package().clone();
        let current = Snapshot::load(
            &candidate,
            PivotCacheSelector::Position(self.source_patch.before.selection),
            PivotCacheFieldSelector::Ordinal(self.source_patch.before.field_ordinal()),
        )?;
        if !current.same_readset(&self.source_patch.before) {
            return Err(Error::PatchConflict {
                part: self.source_patch.before.cache_part().to_string(),
            });
        }
        if self.is_empty() {
            return Ok(WorkbookCommit::new(
                source.clone(),
                Self::new(source.clone(), source.clone(), self.source_patch.clone()),
                false,
            ));
        }
        if candidate.is_signed() || candidate.requires_signature_edit_policy() {
            return Err(Error::Signed);
        }
        candidate
            .get_part_mut(self.source_patch.after.cache_part())?
            .set_blob_shared(Arc::clone(&self.source_patch.after.cache.bytes));
        validate_candidate_limits(&candidate)?;
        let resulting = Snapshot::load(
            &candidate,
            PivotCacheSelector::Position(self.source_patch.after.selection),
            PivotCacheFieldSelector::Ordinal(self.source_patch.after.field_ordinal()),
        )?;
        if !resulting.same_readset(&self.source_patch.after)
            || resulting.cached_unique_names() != self.source_patch.after().cached_unique_names()
        {
            return Err(invalid(
                "cachedUniqueNames Workbook patch verification failed",
            ));
        }
        let workbook = source.adopt_published_package(candidate)?;
        Ok(WorkbookCommit::new(
            workbook.clone(),
            Self::new(source.clone(), workbook, self.source_patch.clone()),
            true,
        ))
    }
}

/// Committed ordinary Workbook edit.
#[derive(Clone, Debug)]
pub struct WorkbookCommit {
    workbook: Workbook,
    patch: WorkbookPatch,
    changed: bool,
}

impl WorkbookCommit {
    fn new(workbook: Workbook, patch: WorkbookPatch, changed: bool) -> Self {
        Self {
            workbook,
            patch,
            changed,
        }
    }

    /// Resulting immutable Workbook.
    #[must_use]
    pub fn workbook(&self) -> &Workbook {
        &self.workbook
    }

    /// Resulting typed cache-field state.
    #[must_use]
    pub fn snapshot(&self) -> &CachedUniqueNames {
        self.patch.after()
    }

    /// Exact reversible patch.
    #[must_use]
    pub fn patch(&self) -> &WorkbookPatch {
        &self.patch
    }

    /// Whether a name changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }
}

/// Load one low-level snapshot.
pub fn load<'a>(
    package: &OpcPackage,
    cache_selector: impl Into<PivotCacheSelector>,
    field_selector: PivotCacheFieldSelector<'a>,
) -> Result<Snapshot> {
    Snapshot::load(package, cache_selector, field_selector)
}

/// Start one low-level source-bound transaction.
pub fn edit<'package, 'selector>(
    package: &'package mut OpcPackage,
    cache_selector: impl Into<PivotCacheSelector>,
    field_selector: PivotCacheFieldSelector<'selector>,
) -> Result<Transaction<'package>> {
    Transaction::new(package, cache_selector, field_selector)
}

/// Apply an exact source-bound patch.
pub fn apply_patch(package: &mut OpcPackage, patch: &Patch) -> Result<()> {
    patch.apply(package)
}

/// An ordinary semantic cache edit handle.
#[derive(Clone, Debug)]
pub struct PivotCacheEditHandle {
    workbook: Workbook,
    selector: PivotCacheSelector,
}

impl PivotCacheEditHandle {
    /// Resolve a field and begin its source-bound transaction.
    pub fn field<'a>(self, selector: PivotCacheFieldSelector<'a>) -> Result<WorkbookTransaction> {
        WorkbookTransaction::new(&self.workbook, self.selector, selector)
    }
}

/// Resolve an ordinary cache handle.
pub(crate) fn workbook_cache(
    workbook: &Workbook,
    selector: impl Into<PivotCacheSelector>,
) -> PivotCacheHandle {
    PivotCacheHandle {
        workbook: workbook.clone(),
        selector: selector.into(),
    }
}

/// Begin an ordinary cache edit handle.
pub(crate) fn edit_workbook(
    workbook: &Workbook,
    selector: impl Into<PivotCacheSelector>,
) -> PivotCacheEditHandle {
    PivotCacheEditHandle {
        workbook: workbook.clone(),
        selector: selector.into(),
    }
}

struct Graph {
    workbook: SourcePart,
    workbook_owner: Arc<Vec<u8>>,
    workbook_context: Arc<Vec<u8>>,
    caches: Vec<CacheInfo>,
    connections: Option<ConnectionInfo>,
    limits: ReadLimits,
}

struct CacheInfo {
    id: PivotCacheId,
    source: SourcePart,
}

struct ConnectionInfo {
    source: SourcePart,
    by_id: HashMap<u32, String>,
    by_name: HashMap<String, u32>,
    model_ids: HashSet<u32>,
}

struct SelectedField {
    cache_id: PivotCacheId,
    field_ordinal: usize,
    field_name: String,
    names: Box<[CachedUniqueName]>,
    entries: Box<[EntrySource]>,
    owner: Range<usize>,
    connection_id: Option<u32>,
    source_connection_name: Option<String>,
    mce_ambiguous: bool,
}

impl Graph {
    fn load(package: &OpcPackage) -> Result<Self> {
        validate_candidate_limits(package)?;
        let relationship_index = RelationshipIndex::build(package)?;
        let workbook_part = package.main_document_part()?;
        let workbook_scan = scan_xml(
            workbook_part.blob(),
            WORKBOOK_PAYLOAD,
            package.read_limits(),
        )?;
        let workbook_root =
            root_element(&workbook_scan, CORE_NS, STRICT_CORE_NS, WORKBOOK_PAYLOAD)?;
        let workbook = SourcePart::capture(&relationship_index, workbook_part)?;
        let pivot_caches = direct_children(
            &workbook_scan,
            workbook_root,
            workbook_root.ns.as_ref(),
            PIVOT_CACHES_PAYLOAD,
        )?;
        let pivot_caches = one_or_none(pivot_caches, "workbook pivotCaches")?
            .ok_or_else(|| invalid("workbook has no pivotCaches element"))?;
        let owner = source_copy(
            workbook_part.blob(),
            pivot_caches.start.start..pivot_caches.end,
            MAX_FRAGMENT_BYTES,
            "workbook pivotCaches owner",
        )?;
        let context = source_copy(
            workbook_part.blob(),
            workbook_root.start.clone(),
            MAX_FRAGMENT_BYTES,
            "workbook root context",
        )?;
        let references = direct_children(
            &workbook_scan,
            pivot_caches,
            workbook_root.ns.as_ref(),
            PIVOT_CACHE_PAYLOAD,
        )?;
        if references.is_empty() || references.len() > MAX_CACHE_FIELDS {
            return Err(invalid(
                "workbook pivotCaches has no bounded cache references",
            ));
        }
        let mut caches = Vec::new();
        caches
            .try_reserve_exact(references.len())
            .map_err(|source| Error::Allocation {
                resource: "cached unique-name cache references",
                source,
            })?;
        let mut seen_ids = HashSet::new();
        seen_ids
            .try_reserve(references.len())
            .map_err(|source| Error::Allocation {
                resource: "cached unique-name cache ID index",
                source,
            })?;
        let mut seen_targets = HashSet::new();
        seen_targets
            .try_reserve(references.len())
            .map_err(|source| Error::Allocation {
                resource: "cached unique-name cache target index",
                source,
            })?;
        for reference in references {
            let id = unique_attr(reference, CACHE_ID_ATTRIBUTE, "workbook pivotCache")?
                .ok_or_else(|| invalid("workbook pivotCache requires cacheId"))
                .and_then(|value| parse_u32(value, "workbook pivotCache cacheId"))?;
            let id = PivotCacheId(id);
            if !seen_ids.insert(id) {
                return Err(invalid("workbook pivotCaches contains duplicate cacheId"));
            }
            let relation_id = required_rel_id(reference, "workbook pivotCache")?;
            let relationship = workbook_part
                .rels()
                .get(&relation_id)
                .ok_or_else(|| invalid("workbook pivotCache relationship is missing"))?;
            if relationship.is_external()
                || relationship.target_query().is_some()
                || relationship.target_fragment().is_some()
                || !matches!(
                    relationship.reltype(),
                    rt::PIVOT_CACHE_DEFINITION | rt::STRICT_PIVOT_CACHE_DEFINITION
                )
            {
                return Err(invalid("workbook pivotCache relationship is invalid"));
            }
            let target = relationship.target_partname()?;
            let canonical = relationship_index
                .canonical_part(&target)?
                .ok_or_else(|| invalid("workbook pivotCache target is not a package Part"))?;
            if !seen_targets.insert(canonical.clone()) {
                return Err(invalid("workbook pivotCaches target one definition twice"));
            }
            let part = package.get_part(canonical)?;
            if part.content_type() != ct::SML_PIVOT_CACHE_DEFINITION {
                return Err(invalid(
                    "workbook pivotCache targets the wrong content type",
                ));
            }
            caches.push(CacheInfo {
                id,
                source: SourcePart::capture(&relationship_index, part)?,
            });
        }
        let connections =
            load_connections(package, workbook_part, &relationship_index, true, false)?.map(
                |(source, catalog)| ConnectionInfo {
                    source,
                    by_id: catalog.by_id,
                    by_name: catalog.by_name,
                    model_ids: catalog.model_ids,
                },
            );
        Ok(Self {
            workbook,
            workbook_owner: Arc::new(owner),
            workbook_context: Arc::new(context),
            caches,
            connections,
            limits: package.read_limits(),
        })
    }

    fn snapshot<'a>(
        &self,
        cache_selector: PivotCacheSelector,
        field_selector: PivotCacheFieldSelector<'a>,
    ) -> Result<Snapshot> {
        let selection = resolve_cache(&self.caches, cache_selector)?;
        let cache = self
            .caches
            .get(selection)
            .ok_or_else(|| invalid("PivotCache selector did not resolve"))?;
        let selected = parse_cache_field(&cache.source, cache.id, field_selector, self.limits)?;
        let connections = validate_connection_route(
            selected.connection_id,
            selected.source_connection_name.as_deref(),
            &self.connections,
        )?;
        let names = selected.names.clone();
        Ok(Snapshot {
            value: CachedUniqueNames {
                cache_id: selected.cache_id,
                field_ordinal: selected.field_ordinal,
                field_name: selected.field_name,
                names,
                index_bound: IndexBoundStatus::Unresolved,
            },
            cache: cache.source.clone(),
            connections: connections.map(|connection| connection.source.clone()),
            workbook: self.workbook.clone(),
            workbook_owner: Arc::clone(&self.workbook_owner),
            workbook_context: Arc::clone(&self.workbook_context),
            owner: selected.owner,
            entries: selected.entries,
            mce_ambiguous: selected.mce_ambiguous,
            selection,
        })
    }
}

fn resolve_cache(caches: &[CacheInfo], selector: PivotCacheSelector) -> Result<usize> {
    match selector {
        PivotCacheSelector::Position(position) => {
            if position >= caches.len() {
                Err(invalid("PivotCache selector position is out of range"))
            } else {
                Ok(position)
            }
        },
        PivotCacheSelector::Id(id) => {
            let mut found = None;
            for (position, cache) in caches.iter().enumerate() {
                if cache.id == id {
                    if found.is_some() {
                        return Err(invalid("PivotCache selector ID is ambiguous"));
                    }
                    found = Some(position);
                }
            }
            found.ok_or_else(|| invalid("PivotCache selector ID did not resolve"))
        },
    }
}

fn parse_cache_field<'a>(
    cache: &SourcePart,
    cache_id: PivotCacheId,
    selector: PivotCacheFieldSelector<'a>,
    limits: ReadLimits,
) -> Result<SelectedField> {
    let scan = scan_xml_with_mce(
        cache.bytes.as_slice(),
        PIVOT_CACHE_DEFINITION_PAYLOAD,
        limits,
    )?;
    let root = root_element(
        &scan,
        CORE_NS,
        STRICT_CORE_NS,
        PIVOT_CACHE_DEFINITION_PAYLOAD,
    )?;
    if root.attrs.iter().any(|attribute| {
        attribute.ns.is_empty()
            && (attribute.local == b"cacheId" || attribute.local == b"pivotCacheId")
    }) {
        return Err(invalid("core pivotCacheDefinition cannot carry a cache ID"));
    }
    validate_optional_definition_id(&scan, root, cache_id)?;
    validate_cache_id_version(&scan, root)?;
    let source = one_or_none(
        direct_children(&scan, root, root.ns.as_ref(), CACHE_SOURCE_PAYLOAD)?,
        "cacheSource",
    )?;
    let source = source.ok_or_else(|| invalid("PivotCache has no cacheSource"))?;
    let source_type = unique_attr(source, TYPE_ATTRIBUTE, "cacheSource")?
        .ok_or_else(|| invalid("cacheSource requires type"))?;
    if source_type != "external" {
        return Err(invalid(
            "cachedUniqueNames requires an external cacheSource",
        ));
    }
    let connection_id = unique_attr(source, CONNECTION_ID_ATTRIBUTE, "cacheSource")?
        .map(|value| parse_u32(value, "cacheSource connectionId"))
        .transpose()?;
    let source_connection_name = find_f057_source_connection(&scan, root, source, limits)?;
    let fields = one_or_none(
        direct_children(&scan, root, root.ns.as_ref(), CACHE_FIELDS_PAYLOAD)?,
        "cacheFields",
    )?
    .ok_or_else(|| invalid("PivotCache has no cacheFields element"))?;
    let field_elements = direct_children(&scan, fields, root.ns.as_ref(), CACHE_FIELD_PAYLOAD)?;
    if let Some(count) = unique_attr(fields, b"count", "cacheFields")? {
        if parse_u32(count, "cacheFields count")? as usize != field_elements.len() {
            return Err(invalid("cacheFields count does not equal child count"));
        }
    }
    if field_elements.is_empty() || field_elements.len() > MAX_CACHE_FIELDS {
        return Err(invalid("cacheFields has no bounded cacheField children"));
    }
    let mut field_names = Vec::new();
    field_names
        .try_reserve_exact(field_elements.len())
        .map_err(|source| Error::Allocation {
            resource: "cached unique-name field names",
            source,
        })?;
    for field in &field_elements {
        let name = unique_attr(field, NAME_ATTRIBUTE, "cacheField")?
            .ok_or_else(|| invalid("cacheField requires name"))
            .and_then(decode_spreadsheet_text)?;
        field_names.push(name);
    }
    let field_ordinal = match selector {
        PivotCacheFieldSelector::Ordinal(position) => {
            if position >= field_elements.len() {
                return Err(invalid("PivotCache field selector is out of range"));
            }
            position
        },
        PivotCacheFieldSelector::Name(name) => {
            let mut found = None;
            for (position, field_name) in field_names.iter().enumerate() {
                if field_name == name {
                    if found.is_some() {
                        return Err(invalid("PivotCache field selector name is ambiguous"));
                    }
                    found = Some(position);
                }
            }
            found.ok_or_else(|| invalid("PivotCache field selector name did not resolve"))?
        },
    };
    let field = field_elements[field_ordinal];
    let extension = direct_exts(&scan, field, CACHED_UNIQUE_NAMES_URI)?;
    for candidate in &extension {
        check_fragment_limit(candidate, "cachedUniqueNames extension")?;
    }
    if extension.len() > 1 {
        return Err(invalid(
            "cacheField has duplicate cachedUniqueNames extensions",
        ));
    }
    let extension = extension
        .first()
        .copied()
        .ok_or_else(|| invalid("cacheField has no cachedUniqueNames extension"))?;
    let mut payload = None;
    let mut mce_ambiguous = false;
    for candidate in &scan.elements {
        if candidate.ns.as_ref() != EXT_NS || candidate.local != CACHED_UNIQUE_NAMES_PAYLOAD {
            continue;
        }
        let Some(through_mce) = owned_cache_payload(&scan, candidate, extension)? else {
            continue;
        };
        check_fragment_limit(candidate, "cachedUniqueNames payload")?;
        if payload.replace(candidate).is_some() {
            return Err(invalid(
                "cacheField has duplicate cachedUniqueNames payloads",
            ));
        }
        mce_ambiguous = through_mce || candidate.mce_context;
    }
    let payload = payload.ok_or_else(|| invalid("cachedUniqueNames extension has no payload"))?;
    if payload.has_non_whitespace_text || payload.has_non_whitespace_cdata {
        return Err(invalid(
            "cachedUniqueNames cannot contain non-whitespace text",
        ));
    }
    for child in scan
        .elements
        .iter()
        .filter(|child| child.parent_index == Some(payload.index))
    {
        if child.ns.as_ref() != EXT_NS || child.local != CACHED_UNIQUE_NAME_PAYLOAD {
            return Err(invalid("cachedUniqueNames has an unsupported direct child"));
        }
    }
    let leaves = direct_children(&scan, payload, EXT_NS, CACHED_UNIQUE_NAME_PAYLOAD)?;
    if leaves.is_empty() || leaves.len() > MAX_UNIQUE_NAMES {
        return Err(invalid(
            "cachedUniqueNames requires one or more bounded leaves",
        ));
    }
    let mut names = Vec::new();
    let mut entries = Vec::new();
    names
        .try_reserve_exact(leaves.len())
        .map_err(|source| Error::Allocation {
            resource: "cached unique-name values",
            source,
        })?;
    entries
        .try_reserve_exact(leaves.len())
        .map_err(|source| Error::Allocation {
            resource: "cached unique-name source ranges",
            source,
        })?;
    let mut indices = HashSet::new();
    indices
        .try_reserve(leaves.len())
        .map_err(|source| Error::Allocation {
            resource: "cached unique-name index set",
            source,
        })?;
    for leaf in leaves {
        if leaf.has_element_child || leaf.has_text || leaf.has_cdata {
            return Err(invalid("cachedUniqueName must be an attribute-only leaf"));
        }
        reject_unknown_attributes(leaf, &[INDEX_ATTRIBUTE, NAME_ATTRIBUTE], "cachedUniqueName")?;
        let index = unique_attr(leaf, INDEX_ATTRIBUTE, "cachedUniqueName")?
            .ok_or_else(|| invalid("cachedUniqueName requires index"))
            .and_then(|value| parse_u32(value, "cachedUniqueName index"))?;
        if !indices.insert(index) {
            return Err(invalid("cachedUniqueNames contains duplicate index"));
        }
        let name_attr = unique_attribute(leaf, NAME_ATTRIBUTE, "cachedUniqueName")?
            .ok_or_else(|| invalid("cachedUniqueName requires name"))?;
        let name = decode_spreadsheet_text(&name_attr.value)?;
        validate_decoded_name(&name, limits)?;
        names.push(CachedUniqueName { index, name });
        entries.push(EntrySource {
            name: AttributeSource {
                value: name_attr.value_range.clone(),
            },
        });
    }
    Ok(SelectedField {
        cache_id,
        field_ordinal,
        field_name: field_names[field_ordinal].clone(),
        names: names.into_boxed_slice(),
        entries: entries.into_boxed_slice(),
        owner: extension.start.start..extension.end,
        connection_id,
        source_connection_name,
        mce_ambiguous,
    })
}

fn validate_cache_id_version(scan: &XmlScan, root: &XmlElement) -> Result<()> {
    let extensions = direct_exts(scan, root, PIVOT_CACHE_ID_VERSION_URI)?;
    for candidate in &extensions {
        check_fragment_limit(candidate, "pivotCacheIdVersion owner")?;
    }
    if extensions.len() > 1 {
        return Err(invalid("duplicate pivotCacheIdVersion extensions"));
    }
    let extension = extensions
        .first()
        .copied()
        .ok_or_else(|| invalid("external PivotCache is missing pivotCacheIdVersion extension"))?;
    let payloads = scan.elements.iter().filter(|element| {
        element.parent_index == Some(extension.index)
            && element.ns.as_ref() == EXT_NS
            && element.local == b"pivotCacheIdVersion"
    });
    let mut payloads = payloads;
    let payload = payloads
        .next()
        .ok_or_else(|| invalid("pivotCacheIdVersion extension has no payload"))?;
    check_fragment_limit(payload, "pivotCacheIdVersion payload")?;
    if let Some(duplicate) = payloads.next() {
        check_fragment_limit(duplicate, "pivotCacheIdVersion payload")?;
        return Err(invalid("duplicate pivotCacheIdVersion payload"));
    }
    if payload.has_element_child || payload.has_text || payload.has_cdata {
        return Err(invalid(
            "pivotCacheIdVersion payload must be an attribute-only leaf",
        ));
    }
    reject_unknown_attributes(
        payload,
        &[
            b"cacheIdSupportedVersion".as_slice(),
            b"cacheIdCreatedVersion".as_slice(),
        ],
        "pivotCacheIdVersion",
    )?;
    for name in [
        b"cacheIdSupportedVersion".as_slice(),
        b"cacheIdCreatedVersion".as_slice(),
    ] {
        let value = unique_attr(payload, name, "pivotCacheIdVersion")?
            .ok_or_else(|| invalid("pivotCacheIdVersion is missing a required attribute"))?;
        if parse_u32(value, "pivotCacheIdVersion attribute")? > u8::MAX as u32 {
            return Err(invalid(
                "pivotCacheIdVersion attribute exceeds unsignedByte",
            ));
        }
    }
    Ok(())
}

fn owned_cache_payload(
    scan: &XmlScan,
    candidate: &XmlElement,
    extension: &XmlElement,
) -> Result<Option<bool>> {
    let mut current = candidate.parent_index;
    let mut child_index = candidate.index;
    let mut through_branch = false;
    while let Some(index) = current {
        let Some(parent) = scan.elements.get(index) else {
            return Ok(None);
        };
        if parent.index == extension.index {
            return Ok(if candidate.parent_index == Some(extension.index) {
                Some(false)
            } else {
                through_branch.then_some(true)
            });
        }
        if parent.ns.as_ref() != MCE_NS {
            return Ok(None);
        }
        match parent.local.as_slice() {
            b"AlternateContent" => {
                let Some(child) = scan.elements.get(child_index) else {
                    return Ok(None);
                };
                if child.parent_index != Some(parent.index)
                    || child.ns.as_ref() != MCE_NS
                    || !matches!(child.local.as_slice(), b"Choice" | b"Fallback")
                {
                    return Ok(None);
                }
            },
            b"Choice" | b"Fallback" => {
                if !mce_branch_selected(scan, parent)? {
                    return Ok(None);
                }
                through_branch = true;
            },
            _ => return Ok(None),
        }
        child_index = parent.index;
        current = parent.parent_index;
    }
    Ok(None)
}

fn mce_branch_selected(scan: &XmlScan, branch: &XmlElement) -> Result<bool> {
    let Some(alternate_index) = branch.parent_index else {
        return Err(invalid("MCE branch has no AlternateContent parent"));
    };
    let alternate = scan
        .elements
        .get(alternate_index)
        .ok_or_else(|| invalid("MCE branch parent index is invalid"))?;
    if alternate.ns.as_ref() != MCE_NS || alternate.local.as_slice() != b"AlternateContent" {
        return Err(invalid("MCE branch is not a direct AlternateContent child"));
    }
    let mut selected = None;
    let mut fallback = None;
    let mut saw_choice = false;
    let mut saw_fallback = false;
    for candidate in scan
        .elements
        .iter()
        .filter(|candidate| candidate.parent_index == Some(alternate.index))
    {
        if candidate.ns.as_ref() != MCE_NS {
            // A direct foreign child is part of the opaque MCE surface when
            // its namespace is named by an in-scope mc:Ignorable declaration.
            // It cannot itself be selected as a branch, but it must not make
            // an otherwise valid Choice/Fallback collection malformed.
            if namespace_is_ignorable(scan, candidate, candidate.ns.as_ref()) {
                continue;
            }
            return Err(invalid("AlternateContent has a non-ignorable branch child"));
        }
        match candidate.local.as_slice() {
            b"Choice" => {
                if saw_fallback {
                    return Err(invalid("MCE Choice follows Fallback"));
                }
                saw_choice = true;
                if selected.is_none() {
                    let Some(MceBranch::Choice(requirements)) = candidate.mce_branch.as_ref()
                    else {
                        return Err(invalid("MCE Choice branch metadata is missing"));
                    };
                    if requirements
                        .iter()
                        .all(|namespace| supported_mce_namespace(namespace))
                    {
                        selected = Some(candidate.index);
                    }
                }
            },
            b"Fallback" => {
                if saw_fallback {
                    return Err(invalid("AlternateContent has duplicate Fallback branches"));
                }
                saw_fallback = true;
                fallback = Some(candidate.index);
            },
            _ => return Err(invalid("AlternateContent has an invalid branch child")),
        }
    }
    if !saw_choice {
        return Err(invalid("AlternateContent requires a Choice branch"));
    }
    Ok(selected.or(fallback) == Some(branch.index))
}

fn namespace_is_ignorable(scan: &XmlScan, element: &XmlElement, namespace: &[u8]) -> bool {
    let scope = element
        .parent_index
        .and_then(|index| scan.elements.get(index))
        .map_or(0, |parent| parent.ignorable_scope);
    scope_contains(&scan.ignorable_scopes, scope, namespace)
}

fn supported_mce_namespace(namespace: &[u8]) -> bool {
    matches!(
        namespace,
        CORE_NS | STRICT_CORE_NS | EXT_NS | X14_NS | REL_NS | STRICT_REL_NS
    )
}

fn validate_optional_definition_id(
    scan: &XmlScan,
    root: &XmlElement,
    expected: PivotCacheId,
) -> Result<()> {
    let extensions = direct_exts(scan, root, PIVOT_CACHE_DEFINITION_URI)?;
    for candidate in &extensions {
        check_fragment_limit(candidate, "pivotCacheDefinition owner")?;
    }
    if extensions.len() > 1 {
        return Err(invalid("duplicate pivotCacheDefinition extensions"));
    }
    let Some(extension) = extensions.first().copied() else {
        return Ok(());
    };
    let payloads = direct_children(scan, extension, X14_NS, b"pivotCacheDefinition")?;
    for payload in &payloads {
        check_fragment_limit(payload, "pivotCacheDefinition payload")?;
    }
    let payload = one_or_none(payloads, "pivotCacheDefinition extension")?
        .ok_or_else(|| invalid("pivotCacheDefinition extension has no payload"))?;
    if let Some(value) = unique_attr(payload, b"pivotCacheId", "pivotCacheDefinition")? {
        if parse_u32(value, "pivotCacheDefinition pivotCacheId")? != expected.0 {
            return Err(invalid(
                "pivotCacheDefinition extension ID disagrees with workbook cache ID",
            ));
        }
    }
    Ok(())
}

fn validate_connection_route<'a>(
    numeric: Option<u32>,
    f057: Option<&str>,
    connections: &'a Option<ConnectionInfo>,
) -> Result<Option<&'a ConnectionInfo>> {
    let Some(connections) = connections.as_ref() else {
        return Err(invalid(
            "cachedUniqueNames requires a workbook connections Part",
        ));
    };
    let by_numeric = numeric
        .map(|id| {
            if id == 0 {
                return Err(invalid(
                    "cacheSource connectionId zero is the default-only value",
                ));
            }
            if !connections.by_id.contains_key(&id) {
                return Err(invalid(
                    "cacheSource connectionId does not resolve to a workbook connection",
                ));
            }
            Ok(id)
        })
        .transpose()?;
    let by_name = f057
        .map(|name| {
            connections.by_name.get(name).copied().ok_or_else(|| {
                invalid(
                    "cacheSource sourceConnection name does not resolve to a workbook connection",
                )
            })
        })
        .transpose()?;
    let resolved = match (by_numeric, by_name) {
        (Some(left), Some(right)) if left != right => {
            return Err(invalid("cacheSource connection routes disagree"));
        },
        (Some(id), _) | (_, Some(id)) => id,
        (None, None) => {
            return Err(invalid(
                "cachedUniqueNames requires connectionId or F057 sourceConnection",
            ));
        },
    };
    if !connections.model_ids.contains(&resolved) {
        return Err(invalid(
            "cachedUniqueNames connection is not a DE250 model connection",
        ));
    }
    Ok(Some(connections))
}

fn find_f057_source_connection(
    scan: &XmlScan,
    root: &XmlElement,
    source: &XmlElement,
    limits: ReadLimits,
) -> Result<Option<String>> {
    let mut found = None;
    for ext_list in direct_children(scan, source, root.ns.as_ref(), b"extLst")? {
        for ext in direct_children(scan, ext_list, root.ns.as_ref(), b"ext")? {
            if unique_attr(ext, b"uri", "cacheSource ext")?
                .is_some_and(|value| xml_token_eq(value, CACHE_SOURCE_URI))
            {
                check_fragment_limit(ext, "F057 extension")?;
                if found.is_some() {
                    return Err(invalid("cacheSource has duplicate F057 extensions"));
                }
                let payloads = direct_children(scan, ext, X14_NS, b"sourceConnection")?;
                for payload in &payloads {
                    check_fragment_limit(payload, "F057 sourceConnection payload")?;
                }
                let payload = one_or_none(payloads, "sourceConnection")?
                    .ok_or_else(|| invalid("F057 extension has no sourceConnection"))?;
                if ext.mce_context
                    || payload.mce_context
                    || payload.has_element_child
                    || payload.has_text
                    || payload.has_cdata
                {
                    return Err(invalid(
                        "F057 sourceConnection has unproven MCE or leaf content",
                    ));
                }
                let name = unique_attr(payload, NAME_ATTRIBUTE, "sourceConnection")?
                    .ok_or_else(|| invalid("sourceConnection requires name"))
                    .and_then(decode_spreadsheet_text)?;
                validate_decoded_name(&name, limits)?;
                if name.encode_utf16().count() >= 65_536 {
                    return Err(invalid("sourceConnection name exceeds its text limit"));
                }
                found = Some(name);
            }
        }
    }
    Ok(found)
}

fn root_element<'a>(
    scan: &'a XmlScan,
    core: &[u8],
    strict: &[u8],
    local: &str,
) -> Result<&'a XmlElement> {
    let root = scan
        .elements
        .iter()
        .find(|element| element.parent_index.is_none())
        .ok_or_else(|| invalid("XML Part has no root"))?;
    if (root.ns.as_ref() != core && root.ns.as_ref() != strict) || root.local != local.as_bytes() {
        return Err(invalid(format!("XML Part root must be {local}")));
    }
    Ok(root)
}

fn direct_children<'a>(
    scan: &'a XmlScan,
    parent: &XmlElement,
    namespace: &[u8],
    local: &[u8],
) -> Result<Vec<&'a XmlElement>> {
    let mut children = Vec::new();
    for element in &scan.elements {
        if element.parent_index == Some(parent.index)
            && element.ns.as_ref() == namespace
            && element.local == local
        {
            children
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "cached unique-name direct child index",
                    source,
                })?;
            children.push(element);
        }
    }
    Ok(children)
}

fn one_or_none<'a>(values: Vec<&'a XmlElement>, owner: &str) -> Result<Option<&'a XmlElement>> {
    if values.len() > 1 {
        return Err(invalid(format!("{owner} has duplicate elements")));
    }
    Ok(values.into_iter().next())
}

fn unique_attribute<'a>(
    element: &'a XmlElement,
    local: &[u8],
    owner: &str,
) -> Result<Option<&'a XmlAttribute>> {
    let mut found = None;
    for attribute in &element.attrs {
        if attribute.ns.is_empty() && attribute.local == local {
            if found.is_some() {
                return Err(invalid(format!("{owner} has duplicate attribute")));
            }
            found = Some(attribute);
        }
    }
    Ok(found)
}

fn unique_attr<'a>(element: &'a XmlElement, local: &[u8], owner: &str) -> Result<Option<&'a str>> {
    unique_attribute(element, local, owner)
        .map(|attribute| attribute.map(|attribute| attribute.value.as_str()))
}

fn reject_unknown_attributes(element: &XmlElement, known: &[&[u8]], owner: &str) -> Result<()> {
    for attribute in &element.attrs {
        if attribute.ns.is_empty() && known.iter().any(|name| *name == attribute.local) {
            continue;
        }
        return Err(invalid(format!("{owner} has an unsupported attribute")));
    }
    Ok(())
}

fn source_copy(
    source: &[u8],
    range: Range<usize>,
    maximum: usize,
    resource: &'static str,
) -> Result<Vec<u8>> {
    let bytes = source
        .get(range)
        .ok_or_else(|| invalid(format!("{resource} range is invalid")))?;
    if bytes.len() > maximum {
        return Err(invalid(format!("{resource} exceeds its limit")));
    }
    let mut copy = Vec::new();
    copy.try_reserve_exact(bytes.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    copy.extend_from_slice(bytes);
    Ok(copy)
}

fn validate_name_value(value: &str, limits: ReadLimits) -> Result<()> {
    // ST_Xstring permits SpreadsheetML escape sequences that decode to XML
    // control code points.  The source serializer re-encodes those units as
    // `_xHHHH_`; rejecting them here would make a valid decoded value
    // impossible to edit.
    validate_decoded_name(value, limits)
}

fn validate_decoded_name(value: &str, limits: ReadLimits) -> Result<()> {
    if value.encode_utf16().count() > MAX_UNIQUE_NAME_UTF16 {
        return Err(invalid(
            "cached unique-name exceeds 65,535 decoded UTF-16 units",
        ));
    }
    let encoded = escaped_xstring_len(value)?;
    let caller = caller_attribute_limit(limits);
    if encoded > caller || encoded > MAX_ATTRIBUTE_TEXT_BYTES {
        return Err(invalid("cached unique-name exceeds its caller text limit"));
    }
    Ok(())
}

fn rewrite_cache(
    before: &Snapshot,
    staged: &[String],
    maximum_output_bytes: usize,
    maximum_total_bytes: u64,
    maximum_attribute_bytes: usize,
    current_total_bytes: u64,
    limits: ReadLimits,
) -> Result<Vec<u8>> {
    if staged.len() != before.entries.len() {
        return Err(invalid("cachedUniqueNames source order changed"));
    }
    let source = before.cache.bytes.as_slice();
    let owner = source
        .get(before.owner.clone())
        .ok_or_else(|| invalid("cachedUniqueNames owner range is invalid"))?;
    if owner.len() > MAX_FRAGMENT_BYTES {
        return Err(invalid("cachedUniqueNames owner fragment exceeds limit"));
    }
    let mut output_len = source.len();
    let mut owner_len = owner.len();
    let mut changed = false;
    for ((entry, previous), next) in before.entries.iter().zip(before.names()).zip(staged) {
        validate_name_value(next, limits)?;
        let encoded_len = escaped_xstring_len(next)?;
        if encoded_len > maximum_attribute_bytes || encoded_len > MAX_ATTRIBUTE_TEXT_BYTES {
            return Err(invalid(
                "cached unique-name encoded value exceeds caller limit",
            ));
        }
        let old_len = entry.name.value.end.saturating_sub(entry.name.value.start);
        if source.get(entry.name.value.clone()).is_none() {
            return Err(invalid("cached unique-name source value range is invalid"));
        }
        if previous.name != *next {
            changed = true;
            output_len = output_len
                .checked_sub(old_len)
                .and_then(|length| length.checked_add(encoded_len))
                .ok_or_else(|| invalid("cachedUniqueNames output length overflows"))?;
            if entry.name.value.start >= before.owner.start
                && entry.name.value.end <= before.owner.end
            {
                owner_len = owner_len
                    .checked_sub(old_len)
                    .and_then(|length| length.checked_add(encoded_len))
                    .ok_or_else(|| invalid("cachedUniqueNames owner length overflows"))?;
            }
        }
    }
    if !changed {
        let mut output = Vec::new();
        output
            .try_reserve_exact(source.len())
            .map_err(|source| Error::Allocation {
                resource: "cachedUniqueNames no-op source",
                source,
            })?;
        output.extend_from_slice(source);
        return Ok(output);
    }
    if owner_len > MAX_FRAGMENT_BYTES {
        return Err(invalid(
            "rewritten cachedUniqueNames owner exceeds fragment limit",
        ));
    }
    if output_len > MAX_PART_BYTES || output_len > maximum_output_bytes {
        return Err(invalid(
            "cachedUniqueNames output exceeds the caller's Part limit",
        ));
    }
    let source_len = u64::try_from(source.len()).unwrap_or(u64::MAX);
    let output_len_u64 = u64::try_from(output_len).unwrap_or(u64::MAX);
    let prospective_total = current_total_bytes
        .checked_sub(source_len)
        .and_then(|total| total.checked_add(output_len_u64))
        .ok_or_else(|| invalid("cachedUniqueNames aggregate bytes overflow"))?;
    if prospective_total > maximum_total_bytes {
        return Err(invalid(
            "cachedUniqueNames output exceeds aggregate Part limit",
        ));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "cachedUniqueNames output",
            source,
        })?;
    let mut cursor = 0usize;
    for ((entry, previous), next) in before.entries.iter().zip(before.names()).zip(staged) {
        let range = &entry.name.value;
        output.extend_from_slice(
            source
                .get(cursor..range.start)
                .ok_or_else(|| invalid("cachedUniqueNames source edit range is invalid"))?,
        );
        let old_raw = source
            .get(range.clone())
            .ok_or_else(|| invalid("cachedUniqueNames source value range is invalid"))?;
        if previous.name == *next {
            output.extend_from_slice(old_raw);
        } else {
            append_escaped_xstring(&mut output, next);
        }
        cursor = range.end;
    }
    output.extend_from_slice(
        source
            .get(cursor..)
            .ok_or_else(|| invalid("cachedUniqueNames source tail is invalid"))?,
    );
    if output.len() != output_len {
        return Err(invalid(
            "cachedUniqueNames output length preflight mismatch",
        ));
    }
    Ok(output)
}

fn same_optional_source(left: &Option<SourcePart>, right: &Option<SourcePart>) -> bool {
    match (left, right) {
        (None, None) => true,
        (Some(left), Some(right)) => left.same_source(right),
        _ => false,
    }
}

fn same_optional_closure(left: &Option<SourcePart>, right: &Option<SourcePart>) -> bool {
    match (left, right) {
        (None, None) => true,
        (Some(left), Some(right)) => left.same_closure(right),
        _ => false,
    }
}

#[cfg(test)]
mod tests;
