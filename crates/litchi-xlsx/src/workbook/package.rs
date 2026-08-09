//! OPC graph ownership for the workbook snapshot.
//!
//! This layer validates the [MS-XLSX] workbook relationships and content types
//! before the semantic model exposes sheets. It keeps physical OPC identities
//! below the selector-first facade.

use std::collections::HashMap;
use std::sync::Arc;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI, Part};

use super::model::{Workbook, WorksheetKind};
use crate::error::{Result, invalid};
use crate::raw;

const CHARTSHEET_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
const STRICT_CHARTSHEET_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet";
const DIALOGSHEET_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet";
const STRICT_DIALOGSHEET_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/dialogsheet";
const MACROSHEET_REL: &str = "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet";
const INTL_MACROSHEET_REL: &str =
    "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet";
const CHARTSHEET_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";

pub(super) struct SheetPart {
    pub(super) kind: WorksheetKind,
    pub(super) uri: PackURI,
}

pub(super) fn validate_sheet_graph(
    package: &OpcPackage,
    workbook: &dyn Part,
    sheets: &[raw::Sheet],
) -> Result<Vec<SheetPart>> {
    let mut parts = Vec::with_capacity(sheets.len());
    let mut targets = HashMap::<PackURI, usize>::with_capacity(sheets.len());
    for (position, sheet) in sheets.iter().enumerate() {
        let relationship = workbook.rels().get(&sheet.relationship_id).ok_or_else(|| {
            invalid(format!(
                "sheet '{}' references missing relationship '{}'",
                sheet.name, sheet.relationship_id
            ))
        })?;
        if relationship.is_external() {
            return Err(invalid(format!(
                "sheet '{}' relationship cannot be external",
                sheet.name
            )));
        }
        let target = relationship.target_partname()?;
        let part = package.get_part(&target)?;
        let kind = match relationship.reltype() {
            rt::WORKSHEET | rt::STRICT_WORKSHEET => {
                require_content_type(sheet, part.content_type(), ct::SML_WORKSHEET)?;
                WorksheetKind::Worksheet
            },
            CHARTSHEET_REL | STRICT_CHARTSHEET_REL => {
                require_content_type(sheet, part.content_type(), CHARTSHEET_CONTENT_TYPE)?;
                WorksheetKind::Chart
            },
            DIALOGSHEET_REL | STRICT_DIALOGSHEET_REL => WorksheetKind::Dialog,
            MACROSHEET_REL | INTL_MACROSHEET_REL => WorksheetKind::Macro,
            _ => WorksheetKind::Unknown,
        };
        let uri = part.partname().clone();
        if let Some(previous) = targets.insert(uri.clone(), position) {
            return Err(invalid(format!(
                "sheet part '{uri}' is referenced by both '{}' and '{}'",
                sheets[previous].name, sheet.name
            )));
        }
        parts.push(SheetPart { kind, uri });
    }
    Ok(parts)
}

pub(super) fn validate_shared_strings(
    package: &OpcPackage,
    workbook: &dyn Part,
) -> Result<Option<PackURI>> {
    let mut found = None;
    for relationship in workbook.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            rt::SHARED_STRINGS | rt::STRICT_SHARED_STRINGS
        )
    }) {
        if found.is_some() {
            return Err(invalid("workbook has multiple shared-string relationships"));
        }
        if relationship.is_external() {
            return Err(invalid("shared-string relationship cannot be external"));
        }
        let uri = relationship.target_partname()?;
        let part = package.get_part(&uri)?;
        if part.content_type() != ct::SML_SHARED_STRINGS {
            return Err(invalid(format!(
                "shared-string part has content type '{}', expected '{}'",
                part.content_type(),
                ct::SML_SHARED_STRINGS
            )));
        }
        found = Some(uri);
    }
    Ok(found)
}

pub(super) fn validate_styles(
    package: &OpcPackage,
    workbook: &dyn Part,
) -> Result<Option<PackURI>> {
    let mut found = None;
    for relationship in workbook
        .rels()
        .iter()
        .filter(|relationship| matches!(relationship.reltype(), rt::STYLES | rt::STRICT_STYLES))
    {
        if found.is_some() {
            return Err(invalid("workbook has multiple styles relationships"));
        }
        if relationship.is_external() {
            return Err(invalid("styles relationship cannot be external"));
        }
        let uri = relationship.target_partname()?;
        let part = package.get_part(&uri)?;
        if part.content_type() != ct::SML_STYLES {
            return Err(invalid(format!(
                "styles part has content type '{}', expected '{}'",
                part.content_type(),
                ct::SML_STYLES
            )));
        }
        found = Some(uri);
    }
    Ok(found)
}

pub(super) fn same_style_table(
    source: &Workbook,
    package: &OpcPackage,
    styles_uri: Option<&PackURI>,
) -> Result<bool> {
    let (Some(source_uri), Some(styles_uri)) = (source.inner.styles_uri.as_ref(), styles_uri)
    else {
        return Ok(source.inner.styles_uri.is_none() && styles_uri.is_none());
    };
    if source_uri != styles_uri {
        return Ok(false);
    }
    let before = source.inner.package.get_part(source_uri)?;
    let after = package.get_part(styles_uri)?;
    if before.content_type() != after.content_type() {
        return Ok(false);
    }
    let before_blob = before.blob_arc();
    let after_blob = after.blob_arc();
    Ok(Arc::ptr_eq(&before_blob, &after_blob) || before_blob.as_slice() == after_blob.as_slice())
}

pub(super) fn same_shared_string_table(
    source: &Workbook,
    package: &OpcPackage,
    shared_strings_uri: Option<&PackURI>,
) -> Result<bool> {
    let (Some(source_uri), Some(shared_strings_uri)) =
        (source.inner.shared_strings_uri.as_ref(), shared_strings_uri)
    else {
        return Ok(source.inner.shared_strings_uri.is_none() && shared_strings_uri.is_none());
    };
    if source_uri != shared_strings_uri {
        return Ok(false);
    }
    let before = source.inner.package.get_part(source_uri)?;
    let after = package.get_part(shared_strings_uri)?;
    if before.content_type() != after.content_type() {
        return Ok(false);
    }
    let before_blob = before.blob_arc();
    let after_blob = after.blob_arc();
    Ok(Arc::ptr_eq(&before_blob, &after_blob) || before_blob.as_slice() == after_blob.as_slice())
}

fn require_content_type(sheet: &raw::Sheet, actual: &str, expected: &str) -> Result<()> {
    if actual != expected {
        return Err(invalid(format!(
            "sheet '{}' has content type '{actual}', expected '{expected}'",
            sheet.name
        )));
    }
    Ok(())
}
