//! Validation and storage helpers of the package edit family for inert XLSB slicer metadata.

use std::collections::HashSet;

use litchi_opc::{OpcPackage, PackURI};

use super::{Cache, Views, package, write_cache, write_views};
use crate::package::error::{Error, Result};
use crate::package::owner_transaction::Caseless;

/// The semantic value owned by a slicer snapshot.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) enum Content {
    /// Workbook-scoped slicer cache definitions.
    Caches(Vec<Cache>),
    /// Slicer views attached to one worksheet.
    Views(Views),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) enum Scope {
    Caches(PackURI),
    Views(PackURI),
}

pub(super) fn inspect(
    package_value: &OpcPackage,
    scope: &Scope,
    limits: crate::raw::Limits,
) -> Result<Content> {
    match scope {
        Scope::Caches(workbook) => {
            let parts = package::load_caches(package_value, workbook, limits)?;
            let mut caches = Vec::new();
            caches
                .try_reserve_exact(parts.len())
                .map_err(|source| Error::Allocation {
                    resource: "slicer cache snapshot",
                    source,
                })?;
            for part in parts {
                caches.push(part.cache);
            }
            Ok(Content::Caches(caches))
        },
        Scope::Views(worksheet) => {
            let view_part = package::load_views(package_value, worksheet, limits)?;
            Ok(Content::Views(
                view_part.map_or_else(Views::new, |attached| attached.views),
            ))
        },
    }
}

pub(super) fn validate_closure(
    package_value: &OpcPackage,
    scope: &Scope,
    content: &Content,
    limits: crate::raw::Limits,
) -> Result<()> {
    let cache_names = match scope {
        Scope::Caches(_) => match content {
            Content::Caches(caches) => caches,
            Content::Views(_) => return Err(scope_error("cache")),
        },
        Scope::Views(_) => {
            let workbook = package_value.main_document_part()?.partname().clone();
            return validate_views_against_caches(package_value, &workbook, content, limits);
        },
    };
    let names = cache_name_set(cache_names.iter().map(|cache| cache.name.as_str()))?;
    for part in package_value.iter_parts().filter(|part| {
        part.rels()
            .iter()
            .any(|relationship| relationship.reltype() == package::VIEWS_RELATIONSHIP_TYPE)
    }) {
        let views = package::load_views(package_value, part.partname(), limits)?
            .map_or_else(Views::new, |attached| attached.views);
        if views
            .items
            .iter()
            .any(|view| !names.contains(&Caseless::new(&view.cache_name)))
        {
            return Err(Error::InvalidFormat(
                "slicer view references a missing cache",
            ));
        }
    }
    Ok(())
}

fn validate_views_against_caches(
    package_value: &OpcPackage,
    workbook: &PackURI,
    content: &Content,
    limits: crate::raw::Limits,
) -> Result<()> {
    let Content::Views(views) = content else {
        return Err(scope_error("view"));
    };
    let caches = package::load_caches(package_value, workbook, limits)?;
    let names = cache_name_set(caches.iter().map(|part| part.cache.name.as_str()))?;
    if views
        .items
        .iter()
        .any(|view| !names.contains(&Caseless::new(&view.cache_name)))
    {
        return Err(Error::InvalidFormat(
            "slicer view references a missing cache",
        ));
    }
    Ok(())
}

pub(super) fn validate_content(content: &Content, limits: crate::raw::Limits) -> Result<()> {
    match content {
        Content::Caches(caches) => validate_caches(caches, limits),
        Content::Views(views) => write_views(views, limits).map(|_| ()),
    }
}

pub(super) fn validate_caches(caches: &[Cache], limits: crate::raw::Limits) -> Result<()> {
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
            resource: "slicer cache names",
            source,
        })?;
    for cache in caches {
        write_cache(cache, limits)?;
        if !names.insert(Caseless::new(&cache.name)) {
            return Err(Error::InvalidFormat("duplicate slicer cache name"));
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
            resource: "slicer dependency names",
            source,
        })?;
    for name in names {
        output.insert(Caseless::new(name));
    }
    Ok(output)
}

pub(super) fn store(
    package_value: &mut OpcPackage,
    scope: &Scope,
    content: &Content,
    limits: crate::raw::Limits,
) -> Result<()> {
    match (scope, content) {
        (Scope::Caches(workbook), Content::Caches(caches)) => {
            package::store_caches(package_value, workbook, caches, limits)
        },
        (Scope::Views(worksheet), Content::Views(views)) => {
            package::store_views(package_value, worksheet, views, limits)
        },
        _ => Err(scope_error("matching")),
    }
}

fn scope_error(_expected: &str) -> Error {
    Error::refused(
        litchi_core::edit::RefusalKind::UnsupportedContent,
        "edit_slicers",
        "slicer transaction is not a owner",
    )
}
