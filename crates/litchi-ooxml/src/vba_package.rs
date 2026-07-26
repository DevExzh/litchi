//! Shared, inert MS-OFFMACRO2 package-graph mutation.

use crate::error::{OoxmlError, Result};
use crate::vba::{VbaLimits, VbaProject as ParsedVbaProject};
use litchi_cfb::OleFile;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::{OpcPackage, PackURI};
use std::io::Cursor;

const WORD_PROJECT_PART: &str = "/word/vbaProject.bin";
const WORD_PROJECT_TARGET: &str = "vbaProject.bin";
const WORD_SUPPLEMENTAL_PART: &str = "/word/vbaData.xml";
const WORD_SUPPLEMENTAL_TARGET: &str = "vbaData.xml";
const EXCEL_PROJECT_PART: &str = "/xl/vbaProject.bin";
const EXCEL_PROJECT_TARGET: &str = "vbaProject.bin";
const POWERPOINT_PROJECT_PART: &str = "/ppt/vbaProject.bin";
const POWERPOINT_PROJECT_TARGET: &str = "vbaProject.bin";

/// OOXML application family that owns a VBA project.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum VbaPackageHost {
    Word,
    Excel,
    PowerPoint,
}

#[derive(Debug, Clone)]
struct ExistingGraph {
    relationship_id: String,
    project_part: PackURI,
    supplemental: Option<(PackURI, String)>,
}

impl VbaPackageHost {
    fn project_part(self) -> &'static str {
        match self {
            Self::Word => WORD_PROJECT_PART,
            Self::Excel => EXCEL_PROJECT_PART,
            Self::PowerPoint => POWERPOINT_PROJECT_PART,
        }
    }

    fn project_target(self) -> &'static str {
        match self {
            Self::Word => WORD_PROJECT_TARGET,
            Self::Excel => EXCEL_PROJECT_TARGET,
            Self::PowerPoint => POWERPOINT_PROJECT_TARGET,
        }
    }

    fn macro_content_type(self, current: &str) -> Result<&'static str> {
        let mapped = match self {
            Self::Word => match current {
                ct::WML_DOCUMENT_MAIN | ct::WML_DOCUMENT_MACRO_MAIN => ct::WML_DOCUMENT_MACRO_MAIN,
                ct::WML_TEMPLATE_MAIN | ct::WML_TEMPLATE_MACRO_MAIN => ct::WML_TEMPLATE_MACRO_MAIN,
                _ => return Err(invalid_main_content_type(self, current)),
            },
            Self::Excel => match current {
                ct::SML_SHEET_MAIN | ct::SML_SHEET_MACRO_MAIN => ct::SML_SHEET_MACRO_MAIN,
                ct::SML_TEMPLATE_MAIN | ct::SML_TEMPLATE_MACRO_MAIN => ct::SML_TEMPLATE_MACRO_MAIN,
                _ => return Err(invalid_main_content_type(self, current)),
            },
            Self::PowerPoint => match current {
                ct::PML_PRESENTATION_MAIN | ct::PML_PRES_MACRO_MAIN => ct::PML_PRES_MACRO_MAIN,
                ct::PML_SLIDESHOW_MAIN | ct::PML_SLIDESHOW_MACRO_MAIN => {
                    ct::PML_SLIDESHOW_MACRO_MAIN
                },
                ct::PML_TEMPLATE_MAIN | ct::PML_TEMPLATE_MACRO_MAIN => ct::PML_TEMPLATE_MACRO_MAIN,
                _ => return Err(invalid_main_content_type(self, current)),
            },
        };
        Ok(mapped)
    }

    fn non_macro_content_type(self, current: &str) -> Result<&'static str> {
        let mapped = match self {
            Self::Word => match current {
                ct::WML_DOCUMENT_MAIN | ct::WML_DOCUMENT_MACRO_MAIN => ct::WML_DOCUMENT_MAIN,
                ct::WML_TEMPLATE_MAIN | ct::WML_TEMPLATE_MACRO_MAIN => ct::WML_TEMPLATE_MAIN,
                _ => return Err(invalid_main_content_type(self, current)),
            },
            Self::Excel => match current {
                ct::SML_SHEET_MAIN | ct::SML_SHEET_MACRO_MAIN => ct::SML_SHEET_MAIN,
                ct::SML_TEMPLATE_MAIN | ct::SML_TEMPLATE_MACRO_MAIN => ct::SML_TEMPLATE_MAIN,
                _ => return Err(invalid_main_content_type(self, current)),
            },
            Self::PowerPoint => match current {
                ct::PML_PRESENTATION_MAIN | ct::PML_PRES_MACRO_MAIN => ct::PML_PRESENTATION_MAIN,
                ct::PML_SLIDESHOW_MAIN | ct::PML_SLIDESHOW_MACRO_MAIN => ct::PML_SLIDESHOW_MAIN,
                ct::PML_TEMPLATE_MAIN | ct::PML_TEMPLATE_MACRO_MAIN => ct::PML_TEMPLATE_MAIN,
                _ => return Err(invalid_main_content_type(self, current)),
            },
        };
        Ok(mapped)
    }
}

/// Validate a complete OOXML VBA payload without executing any source.
pub(crate) fn validate_vba_project_payload(payload: &[u8], limits: &VbaLimits) -> Result<()> {
    let maximum = limits
        .max_total_source_bytes
        .saturating_add(limits.max_decompressed_stream_bytes.saturating_mul(4));
    if payload.len() > maximum {
        return Err(OoxmlError::InvalidFormat(format!(
            "VBA project payload exceeds package limit: {} > {maximum}",
            payload.len()
        )));
    }
    let mut compound = OleFile::open(Cursor::new(payload)).map_err(|error| {
        OoxmlError::InvalidFormat(format!("invalid VBA project compound file: {error}"))
    })?;
    ParsedVbaProject::open(&mut compound, &[], limits)
        .map_err(|error| OoxmlError::InvalidFormat(error.to_string()))?;
    Ok(())
}

/// Replace or attach the one VBA graph allowed for an OOXML main part.
pub(crate) fn store_vba_project_graph(
    package: &mut OpcPackage,
    source_part: &PackURI,
    host: VbaPackageHost,
    project_payload: Vec<u8>,
    word_supplemental_xml: Option<Vec<u8>>,
) -> Result<()> {
    let macro_content_type = {
        let source = package.get_part(source_part)?;
        host.macro_content_type(source.content_type())?
    };
    if host == VbaPackageHost::Word && word_supplemental_xml.is_none() {
        return Err(OoxmlError::InvalidFormat(
            "Word VBA project storage requires supplemental-data XML".to_string(),
        ));
    }
    if host != VbaPackageHost::Word && word_supplemental_xml.is_some() {
        return Err(OoxmlError::InvalidFormat(
            "VBA supplemental-data XML is only valid for Word".to_string(),
        ));
    }

    let existing = inspect_existing_graph(package, source_part, host)?;
    ensure_replacement_targets_available(package, host, existing.as_ref())?;
    if let Some(graph) = &existing {
        ensure_exclusive_inbound_reference(
            package,
            &graph.project_part,
            source_part,
            &graph.relationship_id,
        )?;
        if let Some((supplemental, relationship_id)) = &graph.supplemental {
            ensure_exclusive_inbound_reference(
                package,
                supplemental,
                &graph.project_part,
                relationship_id,
            )?;
        }
    }

    if let Some(graph) = existing {
        package
            .get_part_mut(source_part)?
            .rels_mut()
            .remove(&graph.relationship_id);
        package.remove_part(&graph.project_part);
        if let Some((supplemental, _)) = graph.supplemental {
            package.remove_part(&supplemental);
        }
    }

    let project_name = PackURI::new(host.project_part())
        .map_err(|error| OoxmlError::InvalidUri(error.to_string()))?;
    let mut project = BlobPart::new(
        project_name,
        ct::OFC_VBA_PROJECT.to_string(),
        project_payload,
    );
    if let Some(xml) = word_supplemental_xml {
        project.relate_to(WORD_SUPPLEMENTAL_TARGET, rt::WORD_VBA_DATA);
        let supplemental_name = PackURI::new(WORD_SUPPLEMENTAL_PART)
            .map_err(|error| OoxmlError::InvalidUri(error.to_string()))?;
        package.try_add_part(Box::new(BlobPart::new(
            supplemental_name,
            ct::WML_VBA_DATA.to_string(),
            xml,
        )))?;
    }
    package.try_add_part(Box::new(project))?;

    let source = package.get_part_mut(source_part)?;
    source.relate_to(host.project_target(), rt::VBA_PROJECT);
    source.set_content_type(macro_content_type.to_string())?;
    package.clear_digital_signatures().map_err(|error| {
        OoxmlError::InvalidFormat(format!(
            "failed to clear signatures after storing a VBA project: {error}"
        ))
    })?;
    Ok(())
}

/// Remove the complete VBA graph and restore the corresponding non-macro main type.
pub(crate) fn remove_vba_project_graph(
    package: &mut OpcPackage,
    source_part: &PackURI,
    host: VbaPackageHost,
) -> Result<bool> {
    let non_macro_content_type = {
        let source = package.get_part(source_part)?;
        host.non_macro_content_type(source.content_type())?
    };
    let existing = inspect_existing_graph(package, source_part, host)?;
    let Some(graph) = existing else {
        return Ok(false);
    };
    ensure_exclusive_inbound_reference(
        package,
        &graph.project_part,
        source_part,
        &graph.relationship_id,
    )?;
    if let Some((supplemental, relationship_id)) = &graph.supplemental {
        ensure_exclusive_inbound_reference(
            package,
            supplemental,
            &graph.project_part,
            relationship_id,
        )?;
    }

    let source = package.get_part_mut(source_part)?;
    source.rels_mut().remove(&graph.relationship_id);
    source.set_content_type(non_macro_content_type.to_string())?;
    package.remove_part(&graph.project_part);
    if let Some((supplemental, _)) = graph.supplemental {
        package.remove_part(&supplemental);
    }
    package.clear_digital_signatures().map_err(|error| {
        OoxmlError::InvalidFormat(format!(
            "failed to clear signatures after removing a VBA project: {error}"
        ))
    })?;
    Ok(true)
}

fn inspect_existing_graph(
    package: &OpcPackage,
    source_part: &PackURI,
    host: VbaPackageHost,
) -> Result<Option<ExistingGraph>> {
    let source = package.get_part(source_part)?;
    let mut matches = source
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == rt::VBA_PROJECT);
    let Some(relationship) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(OoxmlError::InvalidFormat(
            "OOXML main part has multiple VBA Project relationships".to_string(),
        ));
    }
    if relationship.is_external() {
        return Err(OoxmlError::InvalidFormat(
            "VBA Project relationship cannot be external".to_string(),
        ));
    }
    let relationship_id = relationship.r_id().to_string();
    let project_part = relationship.target_partname()?;
    let project = package.get_part(&project_part)?;
    if project.content_type() != ct::OFC_VBA_PROJECT {
        return Err(OoxmlError::InvalidContentType {
            expected: ct::OFC_VBA_PROJECT.to_string(),
            got: project.content_type().to_string(),
        });
    }

    let supplemental = if host == VbaPackageHost::Word {
        let mut supplemental = project
            .rels()
            .iter()
            .filter(|relationship| relationship.reltype() == rt::WORD_VBA_DATA);
        let Some(relationship) = supplemental.next() else {
            return Err(OoxmlError::InvalidFormat(
                "Word VBA Project part is missing supplemental data".to_string(),
            ));
        };
        if supplemental.next().is_some()
            || project
                .rels()
                .iter()
                .any(|candidate| candidate.reltype() != rt::WORD_VBA_DATA)
        {
            return Err(OoxmlError::InvalidFormat(
                "Word VBA Project part has invalid relationships".to_string(),
            ));
        }
        if relationship.is_external() {
            return Err(OoxmlError::InvalidFormat(
                "Word VBA supplemental-data relationship cannot be external".to_string(),
            ));
        }
        let part_name = relationship.target_partname()?;
        let part = package.get_part(&part_name)?;
        if part.content_type() != ct::WML_VBA_DATA {
            return Err(OoxmlError::InvalidContentType {
                expected: ct::WML_VBA_DATA.to_string(),
                got: part.content_type().to_string(),
            });
        }
        if !part.rels().is_empty() {
            return Err(OoxmlError::InvalidFormat(
                "Word VBA supplemental-data part cannot have relationships".to_string(),
            ));
        }
        Some((part_name, relationship.r_id().to_string()))
    } else {
        if !project.rels().is_empty() {
            return Err(OoxmlError::InvalidFormat(
                "Excel or PowerPoint VBA Project part cannot have relationships".to_string(),
            ));
        }
        None
    };

    Ok(Some(ExistingGraph {
        relationship_id,
        project_part,
        supplemental,
    }))
}

fn ensure_replacement_targets_available(
    package: &OpcPackage,
    host: VbaPackageHost,
    existing: Option<&ExistingGraph>,
) -> Result<()> {
    let canonical_project = PackURI::new(host.project_part())
        .map_err(|error| OoxmlError::InvalidUri(error.to_string()))?;
    if existing.is_none_or(|graph| graph.project_part != canonical_project) {
        package.validate_new_part_name(&canonical_project)?;
    }
    if host == VbaPackageHost::Word {
        let canonical_supplemental = PackURI::new(WORD_SUPPLEMENTAL_PART)
            .map_err(|error| OoxmlError::InvalidUri(error.to_string()))?;
        if existing
            .and_then(|graph| graph.supplemental.as_ref())
            .map(|(part, _)| part)
            != Some(&canonical_supplemental)
        {
            package.validate_new_part_name(&canonical_supplemental)?;
        }
    }
    Ok(())
}

pub(crate) fn ensure_exclusive_inbound_reference(
    package: &OpcPackage,
    target: &PackURI,
    expected_source: &PackURI,
    expected_relationship_id: &str,
) -> Result<()> {
    let mut unexpected = false;
    for relationship in package.rels().iter() {
        if !relationship.is_external()
            && relationship.target_partname().ok().as_ref() == Some(target)
        {
            unexpected = true;
        }
    }
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            if relationship.is_external()
                || relationship.target_partname().ok().as_ref() != Some(target)
            {
                continue;
            }
            if part.partname() != expected_source || relationship.r_id() != expected_relationship_id
            {
                unexpected = true;
            }
        }
    }
    if unexpected {
        return Err(OoxmlError::InvalidFormat(format!(
            "VBA graph part '{}' has an unexpected inbound relationship",
            target.as_str()
        )));
    }
    Ok(())
}

fn invalid_main_content_type(host: VbaPackageHost, actual: &str) -> OoxmlError {
    OoxmlError::InvalidFormat(format!(
        "{host:?} main part has unsupported content type '{actual}'"
    ))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn maps_every_supported_main_part_kind_both_directions() {
        for (host, plain, macro_enabled) in [
            (
                VbaPackageHost::Word,
                ct::WML_DOCUMENT_MAIN,
                ct::WML_DOCUMENT_MACRO_MAIN,
            ),
            (
                VbaPackageHost::Word,
                ct::WML_TEMPLATE_MAIN,
                ct::WML_TEMPLATE_MACRO_MAIN,
            ),
            (
                VbaPackageHost::Excel,
                ct::SML_SHEET_MAIN,
                ct::SML_SHEET_MACRO_MAIN,
            ),
            (
                VbaPackageHost::Excel,
                ct::SML_TEMPLATE_MAIN,
                ct::SML_TEMPLATE_MACRO_MAIN,
            ),
            (
                VbaPackageHost::PowerPoint,
                ct::PML_PRESENTATION_MAIN,
                ct::PML_PRES_MACRO_MAIN,
            ),
            (
                VbaPackageHost::PowerPoint,
                ct::PML_SLIDESHOW_MAIN,
                ct::PML_SLIDESHOW_MACRO_MAIN,
            ),
            (
                VbaPackageHost::PowerPoint,
                ct::PML_TEMPLATE_MAIN,
                ct::PML_TEMPLATE_MACRO_MAIN,
            ),
        ] {
            assert_eq!(host.macro_content_type(plain).unwrap(), macro_enabled);
            assert_eq!(
                host.macro_content_type(macro_enabled).unwrap(),
                macro_enabled
            );
            assert_eq!(host.non_macro_content_type(macro_enabled).unwrap(), plain);
            assert_eq!(host.non_macro_content_type(plain).unwrap(), plain);
        }
    }
}
