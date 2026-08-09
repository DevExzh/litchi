//! Atomic ODP package reconstruction for chart transactions.

use super::codec::{content_inline, content_xml, locate_pages};
use super::model::{Chart, Location, Storage};
use crate::core::OwnedPackage;
use litchi_core::{Error, Result, xml::escape_xml};
use litchi_odf_common::constants::ODF_CHART;
use litchi_odf_common::embedded::{Source, scan_package};
use litchi_odf_common::package::{Addition, rebuild_package, splice};
use std::collections::{BTreeMap, BTreeSet};

/// Apply one clone-staged chart transaction to an immutable ODP package.
pub(crate) fn apply(source: &OwnedPackage, original: &[Chart], draft: &[Chart]) -> Result<Vec<u8>> {
    let original_content = content_xml(source)?;
    let mut content = original_content.clone();
    let mut edits = Vec::<(usize, usize, String, usize)>::new();
    let mut additions = BTreeMap::<String, Vec<u8>>::new();
    let mut directories = Vec::<(String, String)>::new();
    let mut removed_roots = BTreeSet::<String>::new();

    for before in original {
        let after = draft
            .iter()
            .find(|candidate| candidate.same_identity(before));
        let Some(update) = after else {
            let Location::Existing {
                object_start,
                object_end,
                content_path,
                ..
            } = &before.location
            else {
                continue;
            };
            edits.push((*object_start, *object_end, String::new(), 0));
            if let Some(path) = content_path {
                removed_roots.insert(root_path(path));
            }
            continue;
        };

        if before.part().xml() == update.part().xml() {
            continue;
        }
        let Location::Existing {
            payload,
            content_path,
            ..
        } = &before.location
        else {
            continue;
        };
        if let Some(path) = content_path {
            let bytes = update.part().xml().as_bytes().to_vec();
            if let Some(previous) = additions.insert(path.clone(), bytes.clone())
                && previous != bytes
            {
                return invalid("ODP chart transaction has conflicting shared-part edits");
            }
        } else if let Some((start, end)) = payload {
            edits.push((*start, *end, content_inline(update.part().xml())?, 1));
        } else {
            return invalid("inline ODP chart payload span is missing");
        }
    }

    let pages = locate_pages(&content)?;
    let mut used_roots = source
        .files()?
        .into_iter()
        .filter_map(|path| {
            path.strip_prefix("Object_")
                .and_then(|value| value.split_once('/'))
                .map(|(number, _)| format!("Object_{number}/"))
        })
        .collect::<BTreeSet<_>>();
    let mut inserted = Vec::new();
    for chart in draft {
        let Location::Added { page_index, token } = &chart.location else {
            continue;
        };
        let page = pages
            .iter()
            .find(|page| page.index == *page_index)
            .ok_or_else(|| invalid_error("ODP chart insertion page disappeared"))?;
        let object = match chart.storage() {
            Storage::InlineXml => format!(
                "<draw:object>{}</draw:object>",
                content_inline(chart.part().xml())?
            ),
            Storage::PackageSubdocument => {
                let root = unused_root(&mut used_roots)?;
                let href = root.trim_end_matches('/');
                directories.push((root.clone(), ODF_CHART.to_string()));
                let path = format!("{root}content.xml");
                additions.insert(path, chart.part().xml().as_bytes().to_vec());
                format!(
                    "<draw:object xlink:href=\"./{href}\" xlink:type=\"simple\" xlink:show=\"embed\" xlink:actuate=\"onLoad\"/>"
                )
            },
        };
        let name = chart
            .name()
            .ok_or_else(|| invalid_error("authored ODP charts require a frame name"))?;
        let frame = format!(
            "<draw:frame xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" draw:name=\"{}\">{object}</draw:frame>",
            escape_xml(name)
        );
        inserted.push((page.end, page.end, frame, *token));
    }
    edits.extend(inserted);
    edits.sort_unstable_by(|left, right| right.0.cmp(&left.0).then_with(|| right.3.cmp(&left.3)));
    for (start, end, replacement, _) in edits {
        content = splice(&content, start, end, &replacement)?;
    }

    if !removed_roots.is_empty() {
        let package = source.package()?;
        let remaining = scan_package(&content, None, &package)?;
        removed_roots.retain(|root| {
            !remaining.iter().any(|object| {
                matches!(&object.source, Source::PackageSubdocument { root_path, .. } if root_path == root)
            })
        });
    }
    let excluded_paths = additions.keys().cloned().collect::<Vec<_>>();
    let normalized_additions: Vec<Addition> = additions
        .into_iter()
        .map(|(path, bytes)| Addition {
            path,
            bytes,
            media_type: "text/xml".to_string(),
        })
        .collect();
    let excluded_prefixes = removed_roots.into_iter().collect::<Vec<_>>();
    if normalized_additions.is_empty()
        && directories.is_empty()
        && excluded_prefixes.is_empty()
        && content == original_content
    {
        return Ok(source.as_bytes().to_vec());
    }
    rebuild_package(
        source,
        &content,
        normalized_additions,
        directories,
        excluded_paths,
        excluded_prefixes,
    )
}

fn unused_root(used: &mut BTreeSet<String>) -> Result<String> {
    for index in 1..=100_000usize {
        let root = format!("Object_{index}/");
        if used.insert(root.clone()) {
            return Ok(root);
        }
    }
    invalid("no ODP embedded chart object path is available")
}

fn root_path(content_path: &str) -> String {
    content_path
        .rsplit_once("content.xml")
        .map_or_else(|| content_path.to_string(), |(root, _)| root.to_string())
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
