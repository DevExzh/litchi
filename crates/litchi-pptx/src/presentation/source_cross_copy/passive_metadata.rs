//! A bounded admission proof for passive native producer extensions.
//!
//! Every event still goes through the caller's XML/relationship fences. This
//! recognizer admits only the exact extension owner, URI, namespace, scalar
//! attributes and empty payload grammar. It never normalizes or drops bytes.
use super::{
    Error, MAX_XML_DEPTH, P14_NAMESPACE, Result, STRICT_DML_NAMESPACE, STRICT_PML_NAMESPACE,
    SlideCopyRefusal, TRANSITIONAL_DML_NAMESPACE, TRANSITIONAL_PML_NAMESPACE,
};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::{XmlVersion, encoding::Decoder, events::BytesStart};

const DRAWING_2014: &[u8] = b"http://schemas.microsoft.com/office/drawing/2014/main";
const THEME_2012: &[u8] = b"http://schemas.microsoft.com/office/thememl/2012/main";
const SHAPE_CREATION: &str = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}";
const P15_NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2012/main";
const SLIDE_GUIDES: &str = "{EFAFB233-063F-42B5-8137-9DF3F51BA10A}";
const THEME_FAMILY: &str = "{05A4C25C-085E-4340-85A3-A5531E510DB2}";
const ATTRIBUTE_BYTES: usize = 4096;

#[derive(Clone, Copy, PartialEq, Eq)]
enum Owner {
    Other,
    Slide,
    Shape,
    NonVisual,
    Theme,
    Presentation,
}
#[derive(Clone, Copy)]
struct List {
    owner: Owner,
    depth: usize,
    seen: bool,
    extension: bool,
    payload: bool,
}

pub(super) struct Proof<'a> {
    owners: [Owner; MAX_XML_DEPTH],
    list: Option<List>,
    strict_lists: bool,
    execution: Option<&'a litchi_core::ExecutionContext>,
}
impl<'a> Proof<'a> {
    pub(super) const fn new(
        strict_lists: bool,
        execution: Option<&'a litchi_core::ExecutionContext>,
    ) -> Self {
        Self {
            owners: [Owner::Other; MAX_XML_DEPTH],
            list: None,
            strict_lists,
            execution,
        }
    }
    pub(super) fn start(
        &mut self,
        namespace: &[u8],
        element: &BytesStart<'_>,
        depth: usize,
        empty: bool,
        decoder: Decoder,
        version: XmlVersion,
    ) -> Result<bool> {
        let pml = matches!(namespace, TRANSITIONAL_PML_NAMESPACE | STRICT_PML_NAMESPACE);
        let dml = matches!(namespace, TRANSITIONAL_DML_NAMESPACE | STRICT_DML_NAMESPACE);
        let local = element.local_name();
        let mut admitted = false;
        if let Some(mut list) = self.list {
            if depth == list.depth + 1
                && local.as_ref() == b"ext"
                && ((pml
                    && matches!(
                        list.owner,
                        Owner::Slide | Owner::Shape | Owner::Presentation
                    ))
                    || (dml && matches!(list.owner, Owner::NonVisual | Owner::Theme)))
            {
                if empty {
                    if !self.strict_lists {
                        // A master/layout/theme already admitted core-only
                        // extension markup; no foreign payload needs this proof.
                        self.list = None;
                        return Ok(false);
                    }
                    return refused();
                }
                let values = attributes(element, &[b"uri"], decoder, version, self.execution)?;
                let expected = match list.owner {
                    Owner::Slide => crate::change_tracking::CREATION_EXTENSION_URI,
                    Owner::Shape => crate::change_tracking::MODIFICATION_EXTENSION_URI,
                    Owner::NonVisual => SHAPE_CREATION,
                    Owner::Theme => THEME_FAMILY,
                    Owner::Presentation => SLIDE_GUIDES,
                    Owner::Other => return refused(),
                };
                let matches_uri = values[0].as_deref() == Some(expected)
                    || (list.owner == Owner::Theme
                        && values[0].as_deref()
                            == Some(litchi_drawingml::theme::family::part::EXTENSION_URI));
                if !matches_uri {
                    if !self.strict_lists {
                        self.list = None;
                        *self.owners.get_mut(depth).ok_or_else(error)? = Owner::Other;
                        return Ok(false);
                    }
                    return refused();
                }
                if list.seen {
                    return refused();
                }
                list.seen = true;
                list.extension = true;
            } else if depth == list.depth + 2 && list.extension && !list.payload {
                match list.owner {
                    Owner::Slide | Owner::Shape => {
                        let expected = if list.owner == Owner::Slide {
                            b"creationId".as_slice()
                        } else {
                            b"modId".as_slice()
                        };
                        if namespace != P14_NAMESPACE || local.as_ref() != expected {
                            return refused();
                        }
                        let values =
                            attributes(element, &[b"val"], decoder, version, self.execution)?;
                        let Some(value) = values[0].as_deref() else {
                            return refused();
                        };
                        if value.is_empty()
                            || !value.bytes().all(|b| b.is_ascii_digit())
                            || value.parse::<u32>().is_err()
                        {
                            return refused();
                        }
                    },
                    Owner::NonVisual => {
                        if namespace != DRAWING_2014 || local.as_ref() != b"creationId" {
                            return refused();
                        }
                        let values =
                            attributes(element, &[b"id"], decoder, version, self.execution)?;
                        // MS-ODRAWXML CT_CreationId permits an absent id.
                        if let Some(value) = values[0].as_deref() {
                            guid(value)?;
                        }
                    },
                    Owner::Presentation => {
                        // Admit the producer's empty guide-list declaration;
                        // guide entries and future extensions remain outside this proof.
                        if namespace != P15_NAMESPACE || local.as_ref() != b"sldGuideLst" {
                            return refused();
                        }
                        attributes(element, &[], decoder, version, self.execution)?;
                    },
                    Owner::Theme => {
                        if namespace != THEME_2012 || local.as_ref() != b"themeFamily" {
                            return refused();
                        }
                        let values = attributes(
                            element,
                            &[b"name", b"id", b"vid"],
                            decoder,
                            version,
                            self.execution,
                        )?;
                        if values.iter().any(Option::is_none) {
                            return refused();
                        }
                        guid(values[1].as_deref().unwrap_or_default())?;
                        guid(values[2].as_deref().unwrap_or_default())?;
                    },
                    Owner::Other => return refused(),
                }
                list.payload = true;
            } else {
                return refused();
            }
            self.list = Some(list);
            admitted = true;
        } else if local.as_ref() == b"extLst" {
            let parent = depth
                .checked_sub(1)
                .and_then(|n| self.owners.get(n))
                .copied()
                .unwrap_or(Owner::Other);
            if (pml && matches!(parent, Owner::Slide | Owner::Shape | Owner::Presentation))
                || (dml && matches!(parent, Owner::NonVisual | Owner::Theme))
            {
                attributes(element, &[], decoder, version, self.execution)?;
                if !empty {
                    self.list = Some(List {
                        owner: parent,
                        depth,
                        seen: false,
                        extension: false,
                        payload: false,
                    });
                }
                admitted = true;
            }
        }
        if !empty {
            let owner = match (pml, dml, local.as_ref()) {
                (true, _, b"presentation") => Owner::Presentation,
                (true, _, b"cSld") => Owner::Slide,
                (true, _, b"nvPr") => Owner::Shape,
                (true, _, b"cNvPr") => Owner::NonVisual,
                (_, true, b"theme") => Owner::Theme,
                _ => Owner::Other,
            };
            *self.owners.get_mut(depth).ok_or_else(error)? = owner;
        }
        Ok(admitted)
    }
    pub(super) fn end(&mut self, depth: usize) -> Result<()> {
        if let Some(mut list) = self.list {
            if depth == list.depth + 2 {
                if !list.payload {
                    return refused();
                }
                list.extension = false;
                self.list = Some(list);
            } else if depth == list.depth + 1 {
                if list.extension {
                    return refused();
                }
                self.list = None;
            }
        }
        Ok(())
    }
    pub(super) fn text(&self, bytes: &[u8]) -> Result<()> {
        if self.list.is_some() && !bytes.iter().all(u8::is_ascii_whitespace) {
            return refused();
        }
        Ok(())
    }
    pub(super) const fn active(&self) -> bool {
        self.list.is_some()
    }
}
struct Attributes {
    values: Vec<Option<String>>,
    _memory: Option<litchi_core::Reservation>,
}
impl std::ops::Deref for Attributes {
    type Target = [Option<String>];
    fn deref(&self) -> &Self::Target {
        &self.values
    }
}
fn attributes(
    element: &BytesStart<'_>,
    allowed: &[&[u8]],
    decoder: Decoder,
    version: XmlVersion,
    execution: Option<&litchi_core::ExecutionContext>,
) -> Result<Attributes> {
    let mut seen = 0_u8;
    let mut raw_bytes = 0_usize;
    for attribute in element.checked_attributes() {
        super::check_execution(execution)?;
        let attribute = attribute.map_err(|e| Error::Xml(e.to_string()))?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        let Some(index) = allowed
            .iter()
            .position(|name| *name == attribute.key.as_ref())
        else {
            return refused();
        };
        let bit = 1_u8 << index;
        if seen & bit != 0 || attribute.value.len() > ATTRIBUTE_BYTES {
            return refused();
        }
        seen |= bit;
        raw_bytes += attribute.value.len();
    }
    // At most three bounded scalars. Cover decoder/normalizer temporaries,
    // retained strings, their slots, and the shared GUID validator's value.
    let charge = if allowed.is_empty() {
        0
    } else {
        raw_bytes * 8 + allowed.len() * size_of::<Option<String>>() + super::MAX_STAGING_OVERHEAD
    };
    let memory = execution
        .map(|context| {
            context
                .reserve(litchi_core::Resource::Memory, charge as u64)
                .map_err(super::map_execution_error)
        })
        .transpose()?;
    let mut values = vec![None; allowed.len()];
    for attribute in element.checked_attributes() {
        super::check_execution(execution)?;
        let attribute = attribute.map_err(|e| Error::Xml(e.to_string()))?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        let index = allowed
            .iter()
            .position(|name| *name == attribute.key.as_ref())
            .ok_or_else(error)?;
        values[index] = Some(
            attribute
                .decoded_and_normalized_value(version, decoder)
                .map_err(|e| Error::Xml(e.to_string()))?
                .into_owned(),
        );
    }
    Ok(Attributes {
        values,
        _memory: memory,
    })
}

fn guid(value: &str) -> Result<()> {
    litchi_drawingml::theme::family::Guid::new(value)
        .map(|_| ())
        .map_err(|_| error())
}
fn error() -> Error {
    Error::slide_copy_refused(SlideCopyRefusal::UnknownSemanticSurface)
}
fn refused<T>() -> Result<T> {
    Err(error())
}

#[cfg(test)]
mod tests {
    use super::super::{Dialect, RootNamespace, validate_source_slide_xml, validate_xml};
    use super::*;
    const GUID: &str = "{AABBCCDD-1234-5678-90AB-AABBCCDDEEFF}";
    fn slide(p: &str, a: &str) -> String {
        format!(
            r#"<p:sld xmlns:p="{p}" xmlns:a="{a}" xmlns:created="http://schemas.microsoft.com/office/drawing/2014/main" xmlns:legacy="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Native"><a:extLst><a:ext uri="{SHAPE_CREATION}"><created:creationId id="{GUID}"/></a:ext></a:extLst></p:cNvPr><p:cNvSpPr/><p:nvPr><p:extLst><p:ext uri="{{D42A27DB-BD31-4B8C-83A1-F6EECF244321}}"><legacy:modId val="4294967295"></legacy:modId></p:ext></p:extLst></p:nvPr></p:nvSpPr><p:spPr/></p:sp></p:spTree><p:extLst><p:ext uri="{{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}}"><legacy:creationId val="0"/></p:ext></p:extLst></p:cSld></p:sld>"#
        )
    }
    #[test]
    fn passive_identifiers_theme_family_and_empty_guides_use_checked_owners() {
        for (p, a, dialect) in [
            (
                "http://schemas.openxmlformats.org/presentationml/2006/main",
                "http://schemas.openxmlformats.org/drawingml/2006/main",
                Dialect::Transitional,
            ),
            (
                "http://purl.oclc.org/ooxml/presentationml/main",
                "http://purl.oclc.org/ooxml/drawingml/main",
                Dialect::Strict,
            ),
        ] {
            let xml = slide(p, a);
            assert_eq!(
                validate_source_slide_xml(xml.as_bytes(), "passive metadata", None).unwrap(),
                dialect
            );
            for uri in [
                THEME_FAMILY,
                litchi_drawingml::theme::family::part::EXTENSION_URI,
            ] {
                let xml = format!(
                    r#"<a:theme xmlns:a="{a}" xmlns:family="http://schemas.microsoft.com/office/thememl/2012/main"><a:themeElements/><a:extLst><a:ext uri="{uri}"><family:themeFamily name="Native &amp; Theme" id="{GUID}" vid="{GUID}"/></a:ext></a:extLst></a:theme>"#
                );
                assert_eq!(
                    validate_xml(
                        xml.as_bytes(),
                        b"theme",
                        RootNamespace::DrawingMl,
                        false,
                        true,
                        false,
                        false,
                        false,
                        "passive theme",
                        None
                    )
                    .unwrap(),
                    dialect
                );
            }
            let xml = format!(
                r#"<p:presentation xmlns:p="{p}" xmlns:newer="http://schemas.microsoft.com/office/powerpoint/2012/main"><p:sldIdLst/><p:extLst><p:ext uri="{SLIDE_GUIDES}"><newer:sldGuideLst/></p:ext></p:extLst></p:presentation>"#
            );
            assert_eq!(
                validate_xml(
                    xml.as_bytes(),
                    b"presentation",
                    RootNamespace::PresentationMl,
                    true,
                    true,
                    true,
                    true,
                    false,
                    "empty guides",
                    None
                )
                .unwrap(),
                dialect
            );
        }
    }
    #[test]
    fn passive_extension_admission_rejects_foreign_dependencies_and_invalid_scalars() {
        let xml = slide(
            "http://schemas.openxmlformats.org/presentationml/2006/main",
            "http://schemas.openxmlformats.org/drawingml/2006/main",
        );
        let leaf = format!(r#"<created:creationId id="{GUID}"/>"#);
        let ext = format!(r#"<a:ext uri="{SHAPE_CREATION}">{leaf}</a:ext>"#);
        for changed in [
            xml.replace(SHAPE_CREATION, "urn:unrecognized"),
            xml.replace(GUID, "not-a-guid"),
            xml.replace("4294967295", "4294967296"),
            xml.replace(
                &leaf,
                &format!(r#"<created:creationId id="{GUID}" r:embed="hostile"/>"#),
            ),
            xml.replace(
                &leaf,
                &format!(r#"<created:creationId id="{GUID}"><p:pic/></created:creationId>"#),
            ),
            xml.replace("p:cNvPr", "p:spPr"),
            xml.replace(&ext, &format!("{ext}{ext}")),
            xml.replace(&leaf, "<a:unmodeled/>"),
            xml.replace(
                &leaf,
                &format!(
                    r#"<created:creationId id="{}"/>"#,
                    "X".repeat(ATTRIBUTE_BYTES + 1)
                ),
            ),
        ] {
            assert!(
                matches!(validate_source_slide_xml(changed.as_bytes(),"invalid passive metadata",None),Err(Error::Refused(r)) if r.kind()==litchi_core::edit::RefusalKind::UnsupportedContent)
            );
        }
        let guides = format!(
            r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"><p:extLst><p:ext uri="{SLIDE_GUIDES}"><p15:sldGuideLst><p15:guide/></p15:sldGuideLst></p:ext></p:extLst></p:presentation>"#
        );
        assert!(
            validate_xml(
                guides.as_bytes(),
                b"presentation",
                RootNamespace::PresentationMl,
                true,
                true,
                true,
                true,
                false,
                "nonempty guide dependency",
                None
            )
            .is_err()
        );
    }
    #[test]
    fn dependency_graph_core_extension_admission_is_not_narrowed() {
        for ext in [
            r#"<a:ext uri="urn:old-core"/>"#,
            r#"<a:ext uri="urn:old-core"><a:legacy/></a:ext>"#,
        ] {
            let xml = format!(
                r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements/><a:extLst>{ext}</a:extLst></a:theme>"#
            );
            assert!(
                validate_xml(
                    xml.as_bytes(),
                    b"theme",
                    RootNamespace::DrawingMl,
                    false,
                    true,
                    false,
                    false,
                    false,
                    "legacy core theme",
                    None
                )
                .is_ok()
            );
        }
    }
    #[test]
    fn passive_metadata_allocations_honor_the_callers_memory_budget() {
        use litchi_core::{
            Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, Resource,
        };
        use std::num::{NonZeroU64, NonZeroUsize};
        let xml = slide(
            "http://schemas.openxmlformats.org/presentationml/2006/main",
            "http://schemas.openxmlformats.org/drawingml/2006/main",
        );
        for memory in [1, 64 * 1024] {
            let budget = Budget::root(
                "passive-metadata",
                Limits::new(memory, 1 << 20, 1 << 20, 10000, 256, 1 << 20),
            );
            let (_, token) = CancellationSource::pair();
            let execution = ExecutionLimits::new(
                NonZeroUsize::new(1).unwrap(),
                NonZeroUsize::new(1).unwrap(),
                NonZeroU64::new(4096).unwrap(),
                0,
            )
            .unwrap();
            let context = ExecutionContext::new(budget.clone(), token, execution);
            let result =
                validate_source_slide_xml(xml.as_bytes(), "budgeted metadata", Some(&context));
            if memory == 1 {
                assert!(
                    matches!(result,Err(Error::Opc(litchi_opc::OpcError::Execution(litchi_core::ExecutionError::ResourceLimit(limit)))) if limit.resource==Resource::Memory)
                );
            } else {
                assert_eq!(result.unwrap(), Dialect::Transitional);
            }
            assert_eq!(budget.used(Resource::Memory), 0);
        }
    }
}
