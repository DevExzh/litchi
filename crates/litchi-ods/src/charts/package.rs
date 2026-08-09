//! Package-level chart-part replacement.

use super::codec::content_inline;
use super::model::{Chart, Location};
use crate::package::Package;
use litchi_core::{Error, Result};
use litchi_odf_common::package::{Addition, rebuild_package, splice};
use std::collections::BTreeMap;

/// Apply all staged chart replacements in one package rebuild.
pub(crate) fn replace(package: &Package, original: &[Chart], draft: &[Chart]) -> Result<Vec<u8>> {
    if original.len() != draft.len() {
        return invalid("ODS chart transaction changed the inventory shape");
    }

    let mut content = package.content_xml().to_string();
    let mut inline = Vec::new();
    let mut additions = BTreeMap::<String, Vec<u8>>::new();
    for (before, after) in original.iter().zip(draft) {
        if before.part().xml() == after.part().xml() {
            continue;
        }
        match &before.location {
            Location::Package { content_path } => {
                let previous =
                    additions.insert(content_path.clone(), after.part().xml().as_bytes().to_vec());
                if previous.is_some_and(|bytes| bytes != after.part().xml().as_bytes()) {
                    return invalid("ODS chart transaction has conflicting shared-part edits");
                }
            },
            Location::Inline {
                payload_start,
                payload_end,
            } => inline.push((
                *payload_start,
                *payload_end,
                content_inline(after.part().xml())?,
            )),
        }
    }

    inline.sort_unstable_by(|left, right| right.0.cmp(&left.0));
    for (start, end, replacement) in inline {
        content = splice(&content, start, end, &replacement)?;
    }

    if additions.is_empty() && content == package.content_xml() {
        return Ok(package.package().as_bytes().to_vec());
    }

    let excluded_paths: Vec<String> = additions.keys().cloned().collect();
    let additions = additions
        .into_iter()
        .map(|(path, bytes)| Addition {
            path,
            bytes,
            media_type: "text/xml".to_string(),
        })
        .collect();
    rebuild_package(
        package.package(),
        &content,
        additions,
        Vec::new(),
        excluded_paths,
        Vec::new(),
    )
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

#[cfg(test)]
mod tests {
    use crate::charts::codec::{content_inline, inline_content};

    #[test]
    fn inline_conversion_changes_only_the_document_root() {
        let content = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:chart/></office:body></office:document-content>"#;
        let inline = content_inline(content).expect("test fixture or operation should succeed");
        assert!(inline.contains(
            "office:document office:mimetype=\"application/vnd.oasis.opendocument.chart\""
        ));
        assert_eq!(
            content_inline(
                &inline_content(&inline).expect("test fixture or operation should succeed")
            )
            .expect("test fixture or operation should succeed"),
            inline
        );
    }
}
