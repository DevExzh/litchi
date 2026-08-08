//! Source-checked edits for inert XLSB timeline metadata.

use std::collections::HashSet;
use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI};

use super::{Cache, Views, package, write_cache, write_views};
use crate::package::error::{Error, Result};
use crate::package::owner_transaction::{Caseless, Source};

/// The semantic value owned by a timeline snapshot.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Content {
    /// Workbook-scoped timeline cache definitions.
    Caches(Vec<Cache>),
    /// Timeline views attached to one worksheet.
    Views(Views),
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum Scope {
    Caches(PackURI),
    Views(PackURI),
}

/// An immutable timeline owner bound to its exact package source.
#[derive(Debug, Clone)]
pub struct Snapshot {
    content: Content,
    scope: Scope,
    source: Arc<Source>,
    package: Arc<OpcPackage>,
    canonical: bool,
}

impl Snapshot {
    /// Read workbook timeline caches and bind their dependency closure.
    ///
    /// # Errors
    ///
    /// Returns an error when the package graph, XML owner, or dependency closure is invalid.
    pub(crate) fn read_caches(package: &OpcPackage, workbook: &PackURI) -> Result<Self> {
        Self::read_at(package, Scope::Caches(workbook.clone()))
    }
    /// Read one worksheet's timeline views and bind their dependency closure.
    ///
    /// # Errors
    ///
    /// Returns an error when the package graph, XML owner, or dependency closure is invalid.
    pub(crate) fn read_views(package: &OpcPackage, worksheet: &PackURI) -> Result<Self> {
        Self::read_at(package, Scope::Views(worksheet.clone()))
    }
    /// Borrow the semantic content.
    #[must_use]
    pub const fn content(&self) -> &Content {
        &self.content
    }
    /// Borrow caches when this is a cache snapshot.
    #[must_use]
    pub fn caches(&self) -> Option<&[Cache]> {
        match &self.content {
            Content::Caches(caches) => Some(caches),
            Content::Views(_) => None,
        }
    }
    /// Borrow views when this is a view snapshot.
    #[must_use]
    pub const fn views(&self) -> Option<&Views> {
        match &self.content {
            Content::Views(views) => Some(views),
            Content::Caches(_) => None,
        }
    }
    /// Start a detached edit.
    ///
    /// # Errors
    ///
    /// Returns an error when the bounded semantic fingerprint cannot be allocated or encoded.
    pub fn edit(self) -> Result<Transaction> {
        let before_fingerprint = fingerprint(&self.content)?;
        Ok(Transaction {
            scope: self.scope,
            before_source: self.source,
            before_package: self.package,
            before_fingerprint,
            canonical: self.canonical,
            content: self.content,
        })
    }

    fn read_at(package_value: &OpcPackage, scope: Scope) -> Result<Self> {
        let (content, canonical) = inspect(package_value, &scope)?;
        validate_closure(package_value, &scope, &content)?;
        let source = Source::capture(package_value)?;
        Ok(Self {
            content,
            scope,
            source: Arc::new(source),
            package: Arc::new(package_value.clone()),
            canonical,
        })
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.content == other.content && self.scope == other.scope && self.source == other.source
    }
}
impl Eq for Snapshot {}

/// A detached, failure-atomic timeline edit.
#[derive(Debug, Clone)]
pub struct Transaction {
    scope: Scope,
    before_source: Arc<Source>,
    before_package: Arc<OpcPackage>,
    before_fingerprint: Vec<Vec<u8>>,
    canonical: bool,
    content: Content,
}

impl Transaction {
    /// Borrow staged content.
    #[must_use]
    pub const fn content(&self) -> &Content {
        &self.content
    }
    /// Replace every timeline cache.
    ///
    /// # Errors
    ///
    /// Returns an error when this transaction has the wrong scope or a cache is invalid.
    pub fn replace_caches(&mut self, caches: Vec<Cache>) -> Result<()> {
        ensure_cache_scope(&self.scope)?;
        validate_caches(&caches)?;
        self.content = Content::Caches(caches);
        Ok(())
    }
    /// Add or replace a cache by case-insensitive name.
    ///
    /// # Errors
    ///
    /// Returns an error when this transaction has the wrong scope, allocation fails, or the candidate is invalid.
    pub fn upsert_cache(&mut self, cache: Cache) -> Result<()> {
        ensure_cache_scope(&self.scope)?;
        let Content::Caches(caches) = &mut self.content else {
            return Err(scope_error("cache"));
        };
        write_cache(&cache)?;
        if let Some(index) = caches
            .iter()
            .position(|value| value.name.eq_ignore_ascii_case(&cache.name))
        {
            caches[index] = cache;
        } else {
            caches.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "timeline cache transaction",
                source,
            })?;
            caches.push(cache);
        }
        Ok(())
    }
    /// Remove a cache by case-insensitive name.
    ///
    /// # Errors
    ///
    /// Returns an error when this transaction has the wrong scope or the resulting collection is invalid.
    pub fn remove_cache(&mut self, name: &str) -> Result<Option<Cache>> {
        ensure_cache_scope(&self.scope)?;
        let Content::Caches(caches) = &mut self.content else {
            return Err(scope_error("cache"));
        };
        let Some(index) = caches
            .iter()
            .position(|value| value.name.eq_ignore_ascii_case(name))
        else {
            return Ok(None);
        };
        Ok(Some(caches.remove(index)))
    }
    /// Replace every worksheet timeline view.
    ///
    /// # Errors
    ///
    /// Returns an error when this transaction has the wrong scope or the views are invalid.
    pub fn replace_views(&mut self, views: Views) -> Result<()> {
        ensure_view_scope(&self.scope)?;
        write_views(&views)?;
        self.content = Content::Views(views);
        Ok(())
    }
    /// Add or replace a view by case-insensitive name.
    ///
    /// # Errors
    ///
    /// Returns an error when this transaction has the wrong scope, allocation fails, or the candidate is invalid.
    pub fn upsert_view(&mut self, view: super::View) -> Result<()> {
        ensure_view_scope(&self.scope)?;
        let Content::Views(views) = &mut self.content else {
            return Err(scope_error("view"));
        };
        if let Some(index) = views
            .items
            .iter()
            .position(|value| value.name.eq_ignore_ascii_case(&view.name))
        {
            let previous = std::mem::replace(&mut views.items[index], view);
            if let Err(error) = write_views(views) {
                views.items[index] = previous;
                return Err(error);
            }
        } else {
            views
                .items
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "timeline view transaction",
                    source,
                })?;
            views.items.push(view);
            if let Err(error) = write_views(views) {
                let _ = views.items.pop();
                return Err(error);
            }
        }
        Ok(())
    }
    /// Remove a view by case-insensitive name.
    ///
    /// # Errors
    ///
    /// Returns an error when this transaction has the wrong scope or the resulting collection is invalid.
    pub fn remove_view(&mut self, name: &str) -> Result<Option<super::View>> {
        ensure_view_scope(&self.scope)?;
        let Content::Views(views) = &mut self.content else {
            return Err(scope_error("view"));
        };
        let Some(index) = views
            .items
            .iter()
            .position(|value| value.name.eq_ignore_ascii_case(name))
        else {
            return Ok(None);
        };
        Ok(Some(views.items.remove(index)))
    }
    /// Validate, clone-stage, and produce a reversible source-checked patch.
    ///
    /// # Errors
    ///
    /// Returns an error when source XML is noncanonical or staged validation/readback fails.
    pub fn commit(self) -> Result<Commit> {
        validate_content(&self.content)?;
        let changed = fingerprint(&self.content)? != self.before_fingerprint;
        if changed && !self.canonical {
            return Err(Error::UnsupportedFeature("changed timeline edits require canonical source XML; lexical or unmodeled source content is preserved only by an exact no-op".to_string()));
        }
        let snapshot = if changed {
            let mut candidate = self.before_package.as_ref().clone();
            store(&mut candidate, &self.scope, &self.content)?;
            candidate.unsign();
            let snapshot = Snapshot::read_at(&candidate, self.scope.clone())?;
            if snapshot.content != self.content {
                return Err(Error::InvalidFormat(
                    "timeline transaction readback did not match staged content".to_string(),
                ));
            }
            snapshot
        } else {
            Snapshot::read_at(self.before_package.as_ref(), self.scope.clone())?
        };
        Ok(Commit {
            patch: Patch {
                scope: self.scope,
                before_source: self.before_source,
                after_source: Arc::clone(&snapshot.source),
                before_package: self.before_package,
                after_package: Arc::clone(&snapshot.package),
                changed,
            },
            snapshot,
        })
    }
}

/// A successful timeline transaction result.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}
impl Commit {
    /// Borrow the committed snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }
    /// Borrow the patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }
    /// Consume this result.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A reversible timeline patch guarded by exact owner and dependency source.
#[derive(Debug, Clone)]
pub struct Patch {
    scope: Scope,
    before_source: Arc<Source>,
    after_source: Arc<Source>,
    before_package: Arc<OpcPackage>,
    after_package: Arc<OpcPackage>,
    changed: bool,
}
impl Patch {
    /// Whether this patch changes no semantic content.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        !self.changed
    }
    /// Return the inverse patch, restoring the exact package image and signature state.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            scope: self.scope.clone(),
            before_source: Arc::clone(&self.after_source),
            after_source: Arc::clone(&self.before_source),
            before_package: Arc::clone(&self.after_package),
            after_package: Arc::clone(&self.before_package),
            changed: self.changed,
        }
    }
    /// Apply atomically after exact source validation.
    ///
    /// # Errors
    ///
    /// Returns an error for stale source, invalid package state, or failed staged readback.
    pub(crate) fn apply(&self, package_value: &mut OpcPackage) -> Result<Snapshot> {
        let current = Snapshot::read_at(package_value, self.scope.clone())?;
        if current.source != self.before_source {
            return Err(Error::UnsupportedFeature(
                "timeline patch source snapshot does not match".to_string(),
            ));
        }
        if !self.changed {
            return Ok(current);
        }
        let candidate = self.after_package.as_ref().clone();
        let snapshot = Snapshot::read_at(&candidate, self.scope.clone())?;
        if snapshot.source != self.after_source {
            return Err(Error::InvalidFormat(
                "timeline patch publication did not reproduce its committed source".to_string(),
            ));
        }
        *package_value = candidate;
        Ok(snapshot)
    }
}

/// Read workbook cache metadata as a source-bound snapshot.
///
/// # Errors
///
/// Returns an error when the package graph or timeline dependency closure is invalid.
pub(crate) fn read_caches(package: &OpcPackage, workbook: &PackURI) -> Result<Snapshot> {
    Snapshot::read_caches(package, workbook)
}
/// Read worksheet view metadata as a source-bound snapshot.
///
/// # Errors
///
/// Returns an error when the package graph or timeline dependency closure is invalid.
pub(crate) fn read_views(package: &OpcPackage, worksheet: &PackURI) -> Result<Snapshot> {
    Snapshot::read_views(package, worksheet)
}
/// Apply a source-checked timeline patch.
///
/// # Errors
///
/// Returns an error for stale source, invalid package state, or failed staged readback.
pub(crate) fn apply(package: &mut OpcPackage, patch: &Patch) -> Result<Snapshot> {
    patch.apply(package)
}

fn inspect(package_value: &OpcPackage, scope: &Scope) -> Result<(Content, bool)> {
    match scope {
        Scope::Caches(workbook) => {
            let parts = package::load_caches(package_value, workbook)?;
            let mut canonical = true;
            for part in &parts {
                let uri = PackURI::new(&part.part_name).map_err(Error::InvalidUri)?;
                canonical &=
                    package_value.get_part(&uri)?.blob() == write_cache(&part.cache)?.as_slice();
            }
            let mut caches = Vec::new();
            caches
                .try_reserve_exact(parts.len())
                .map_err(|source| Error::Allocation {
                    resource: "timeline cache snapshot",
                    source,
                })?;
            for part in parts {
                caches.push(part.cache);
            }
            Ok((Content::Caches(caches), canonical))
        },
        Scope::Views(worksheet) => {
            let view_part = package::load_views(package_value, worksheet)?;
            let mut canonical = true;
            if let Some(attached) = &view_part {
                let uri = PackURI::new(&attached.part_name).map_err(Error::InvalidUri)?;
                canonical = package_value.get_part(&uri)?.blob()
                    == write_views(&attached.views)?.as_slice();
            }
            Ok((
                Content::Views(view_part.map_or_else(Views::new, |attached| attached.views)),
                canonical,
            ))
        },
    }
}

fn validate_closure(package_value: &OpcPackage, scope: &Scope, content: &Content) -> Result<()> {
    let cache_names = match scope {
        Scope::Caches(_) => match content {
            Content::Caches(caches) => caches,
            Content::Views(_) => return Err(scope_error("cache")),
        },
        Scope::Views(_) => {
            let workbook = package_value.main_document_part()?.partname().clone();
            return validate_views_against_caches(package_value, &workbook, content);
        },
    };
    let names = cache_name_set(cache_names.iter().map(|cache| cache.name.as_str()))?;
    for part in package_value.iter_parts().filter(|part| {
        part.rels()
            .iter()
            .any(|relationship| relationship.reltype() == super::codec::VIEWS_RELATIONSHIP_TYPE)
    }) {
        let views = package::load_views(package_value, part.partname())?
            .map_or_else(Views::new, |attached| attached.views);
        if views
            .items
            .iter()
            .any(|view| !names.contains(&Caseless::new(&view.cache)))
        {
            return Err(Error::InvalidFormat(
                "timeline view references a missing cache".to_string(),
            ));
        }
    }
    Ok(())
}
fn validate_views_against_caches(
    package_value: &OpcPackage,
    workbook: &PackURI,
    content: &Content,
) -> Result<()> {
    let Content::Views(views) = content else {
        return Err(scope_error("view"));
    };
    let caches = package::load_caches(package_value, workbook)?;
    let names = cache_name_set(caches.iter().map(|part| part.cache.name.as_str()))?;
    if views
        .items
        .iter()
        .any(|view| !names.contains(&Caseless::new(&view.cache)))
    {
        return Err(Error::InvalidFormat(
            "timeline view references a missing cache".to_string(),
        ));
    }
    Ok(())
}
fn validate_content(content: &Content) -> Result<()> {
    match content {
        Content::Caches(caches) => validate_caches(caches),
        Content::Views(views) => write_views(views).map(|_| ()),
    }
}

fn fingerprint(content: &Content) -> Result<Vec<Vec<u8>>> {
    let count = match content {
        Content::Caches(caches) => caches.len(),
        Content::Views(_) => 1,
    };
    let mut encoded = Vec::new();
    encoded
        .try_reserve_exact(count)
        .map_err(|source| Error::Allocation {
            resource: "timeline transaction fingerprint",
            source,
        })?;
    match content {
        Content::Caches(caches) => {
            for cache in caches {
                encoded.push(write_cache(cache)?);
            }
        },
        Content::Views(views) => encoded.push(write_views(views)?),
    }
    Ok(encoded)
}
fn validate_caches(caches: &[Cache]) -> Result<()> {
    if caches.len() > super::model::MAX_CACHES {
        return Err(Error::InvalidLength {
            expected: super::model::MAX_CACHES,
            found: caches.len(),
        });
    }
    let mut names = HashSet::new();
    names
        .try_reserve(caches.len())
        .map_err(|source| Error::Allocation {
            resource: "timeline cache names",
            source,
        })?;
    for cache in caches {
        write_cache(cache)?;
        if !names.insert(Caseless::new(&cache.name)) {
            return Err(Error::InvalidFormat(
                "duplicate timeline cache name".to_string(),
            ));
        }
    }
    Ok(())
}

fn cache_name_set<'a>(
    names: impl ExactSizeIterator<Item = &'a str>,
) -> Result<HashSet<Caseless<'a>>> {
    let mut output = HashSet::new();
    output
        .try_reserve(names.len())
        .map_err(|source| Error::Allocation {
            resource: "timeline dependency names",
            source,
        })?;
    for name in names {
        output.insert(Caseless::new(name));
    }
    Ok(output)
}
fn store(package_value: &mut OpcPackage, scope: &Scope, content: &Content) -> Result<()> {
    match (scope, content) {
        (Scope::Caches(workbook), Content::Caches(caches)) => {
            package::store_caches(package_value, workbook, caches)
        },
        (Scope::Views(worksheet), Content::Views(views)) => {
            package::store_views(package_value, worksheet, views)
        },
        _ => Err(scope_error("matching")),
    }
}
fn ensure_cache_scope(scope: &Scope) -> Result<()> {
    if matches!(scope, Scope::Caches(_)) {
        Ok(())
    } else {
        Err(scope_error("cache"))
    }
}
fn ensure_view_scope(scope: &Scope) -> Result<()> {
    if matches!(scope, Scope::Views(_)) {
        Ok(())
    } else {
        Err(scope_error("view"))
    }
}
fn scope_error(expected: &str) -> Error {
    Error::UnsupportedFeature(format!("timeline transaction is not a {expected} owner"))
}
