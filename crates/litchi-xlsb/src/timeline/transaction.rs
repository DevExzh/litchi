//! Validation and storage helpers of the package edit family for inert XLSB timeline metadata.

use std::collections::HashSet;

use litchi_opc::{OpcPackage, PackURI};

use super::{Cache, Views, package, write_cache, write_views};
use crate::package::error::{Error, Result};
use crate::package::owner_transaction::Caseless;

/// The semantic value owned by a timeline snapshot.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) enum Content {
    /// Workbook-scoped timeline cache definitions.
    Caches(Vec<Cache>),
    /// Timeline views attached to one worksheet.
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
    xml_limits: litchi_core::Limits,
) -> Result<(Content, bool)> {
    match scope {
        Scope::Caches(workbook) => {
            let parts = package::load_caches(package_value, workbook, xml_limits)?;
            let mut canonical = true;
            for part in &parts {
                let uri = PackURI::new(&part.part_name)?;
                canonical &= package_value.get_part(&uri)?.blob()
                    == write_cache(&part.cache, xml_limits)?.as_slice();
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
            let view_part = package::load_views(package_value, worksheet, xml_limits)?;
            let mut canonical = true;
            if let Some(attached) = &view_part {
                let uri = PackURI::new(&attached.part_name)?;
                canonical = package_value.get_part(&uri)?.blob()
                    == write_views(&attached.views, xml_limits)?.as_slice();
            }
            Ok((
                Content::Views(view_part.map_or_else(Views::new, |attached| attached.views)),
                canonical,
            ))
        },
    }
}

pub(super) fn validate_closure(
    package_value: &OpcPackage,
    scope: &Scope,
    content: &Content,
    xml_limits: litchi_core::Limits,
) -> Result<()> {
    let cache_names = match scope {
        Scope::Caches(_) => match content {
            Content::Caches(caches) => caches,
            Content::Views(_) => return Err(scope_error("cache")),
        },
        Scope::Views(_) => {
            let workbook = package_value.main_document_part()?.partname().clone();
            return validate_views_against_caches(package_value, &workbook, content, xml_limits);
        },
    };
    let names = cache_name_set(cache_names.iter().map(|cache| cache.name.as_str()))?;
    for part in package_value.iter_parts().filter(|part| {
        part.rels()
            .iter()
            .any(|relationship| relationship.reltype() == super::codec::VIEWS_RELATIONSHIP_TYPE)
    }) {
        let views = package::load_views(package_value, part.partname(), xml_limits)?
            .map_or_else(Views::new, |attached| attached.views);
        if views
            .items
            .iter()
            .any(|view| !names.contains(&Caseless::new(&view.cache)))
        {
            return Err(Error::InvalidFormat(
                "timeline view references a missing cache",
            ));
        }
    }
    Ok(())
}
fn validate_views_against_caches(
    package_value: &OpcPackage,
    workbook: &PackURI,
    content: &Content,
    xml_limits: litchi_core::Limits,
) -> Result<()> {
    let Content::Views(views) = content else {
        return Err(scope_error("view"));
    };
    let caches = package::load_caches(package_value, workbook, xml_limits)?;
    let names = cache_name_set(caches.iter().map(|part| part.cache.name.as_str()))?;
    if views
        .items
        .iter()
        .any(|view| !names.contains(&Caseless::new(&view.cache)))
    {
        return Err(Error::InvalidFormat(
            "timeline view references a missing cache",
        ));
    }
    Ok(())
}
pub(super) fn validate_content(content: &Content, xml_limits: litchi_core::Limits) -> Result<()> {
    match content {
        Content::Caches(caches) => validate_caches(caches, xml_limits),
        Content::Views(views) => write_views(views, xml_limits).map(|_| ()),
    }
}

pub(super) fn validate_caches(caches: &[Cache], xml_limits: litchi_core::Limits) -> Result<()> {
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
        write_cache(cache, xml_limits)?;
        if !names.insert(Caseless::new(&cache.name)) {
            return Err(Error::InvalidFormat("duplicate timeline cache name"));
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
pub(super) fn store(
    package_value: &mut OpcPackage,
    scope: &Scope,
    content: &Content,
    xml_limits: litchi_core::Limits,
) -> Result<()> {
    match (scope, content) {
        (Scope::Caches(workbook), Content::Caches(caches)) => {
            package::store_caches(package_value, workbook, caches, xml_limits)
        },
        (Scope::Views(worksheet), Content::Views(views)) => {
            package::store_views(package_value, worksheet, views, xml_limits)
        },
        _ => Err(scope_error("matching")),
    }
}
fn scope_error(_expected: &str) -> Error {
    Error::refused(
        litchi_core::edit::RefusalKind::UnsupportedContent,
        "edit_timelines",
        "timeline transaction is not a owner",
    )
}
