//! Inert MS-OFFMACRO2 VBA-project relationship discovery for Word packages.
//!
//! Discovery validates OPC relationship and content-type metadata only.
//! Authoring additionally emits typed supplemental document-event and macro
//! descriptors. Neither path compiles, interprets, or executes VBA.

use crate::error::{Error, Result};
use litchi_core::xml::escape::escape_xml;
use litchi_ooxml_common::vba::{
    Host, read_project_part, remove_project_graph, store_project_graph,
};
use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::{OpcPackage, PackURI, Part};
use litchi_vba::{Limits, project::Project};
use std::sync::Arc;

const WORD_VBA_NAMESPACE: &str = "http://schemas.microsoft.com/office/word/2006/wordml";
const MAX_MACRO_NAME_CHARACTERS: usize = 255;
const MAX_SUPPLEMENTAL_MACROS: usize = 4_096;
const MAX_SUPPLEMENTAL_XML_BYTES: usize = 4 * 1024 * 1024;

/// One active Word document event recorded in VBA supplemental data.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum VbaDocumentEvent {
    New,
    Open,
    Close,
    Sync,
    XmlAfterInsert,
    XmlBeforeDelete,
    ContentControlAfterInsert,
    ContentControlBeforeDelete,
    ContentControlOnExit,
    ContentControlOnEnter,
    StoreUpdate,
    ContentControlContentUpdate,
    BuildingBlockAfterInsert,
}

impl VbaDocumentEvent {
    const ALL: [Self; 13] = [
        Self::New,
        Self::Open,
        Self::Close,
        Self::Sync,
        Self::XmlAfterInsert,
        Self::XmlBeforeDelete,
        Self::ContentControlAfterInsert,
        Self::ContentControlBeforeDelete,
        Self::ContentControlOnExit,
        Self::ContentControlOnEnter,
        Self::StoreUpdate,
        Self::ContentControlContentUpdate,
        Self::BuildingBlockAfterInsert,
    ];

    const fn element_name(self) -> &'static str {
        match self {
            Self::New => "eventDocNew",
            Self::Open => "eventDocOpen",
            Self::Close => "eventDocClose",
            Self::Sync => "eventDocSync",
            Self::XmlAfterInsert => "eventDocXmlAfterInsert",
            Self::XmlBeforeDelete => "eventDocXmlBeforeDelete",
            Self::ContentControlAfterInsert => "eventDocContentControlAfterInsert",
            Self::ContentControlBeforeDelete => "eventDocContentControlBeforeDelete",
            Self::ContentControlOnExit => "eventDocContentControlOnExit",
            Self::ContentControlOnEnter => "eventDocContentControlOnEnter",
            Self::StoreUpdate => "eventDocStoreUpdate",
            Self::ContentControlContentUpdate => "eventDocContentControlContentUpdate",
            Self::BuildingBlockAfterInsert => "eventDocBuildingBlockAfterInsert",
        }
    }
}

/// One Word macro descriptor stored in `vbaData.xml`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaMacroDescriptor {
    name: String,
    menu_help: Option<String>,
}

impl VbaMacroDescriptor {
    /// Create a descriptor from its fully qualified macro name.
    pub fn new(name: impl Into<String>) -> Result<Self> {
        let value = Self {
            name: name.into(),
            menu_help: None,
        };
        value.validate()?;
        Ok(value)
    }

    /// Add ignored legacy menu-help text while preserving it in the package.
    pub fn with_menu_help(mut self, menu_help: impl Into<String>) -> Result<Self> {
        self.menu_help = Some(menu_help.into());
        self.validate()?;
        Ok(self)
    }

    /// Return the stored, case-preserving macro name.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Return the optional legacy menu-help text.
    pub fn menu_help(&self) -> Option<&str> {
        self.menu_help.as_deref()
    }

    fn validate(&self) -> Result<()> {
        validate_supplemental_string(&self.name, "macro name")?;
        if self.name.chars().count() > MAX_MACRO_NAME_CHARACTERS {
            return Err(Error::InvalidFormat(format!(
                "Word VBA macro name exceeds {MAX_MACRO_NAME_CHARACTERS} characters"
            )));
        }
        let uppercase = self.name.to_uppercase();
        if uppercase.chars().count() > MAX_MACRO_NAME_CHARACTERS {
            return Err(Error::InvalidFormat(
                "uppercased Word VBA macro name exceeds 255 characters".to_string(),
            ));
        }
        if let Some(menu_help) = &self.menu_help {
            validate_supplemental_string(menu_help, "macro menu help")?;
            if menu_help.chars().count() > MAX_MACRO_NAME_CHARACTERS {
                return Err(Error::InvalidFormat(
                    "Word VBA macro menu help exceeds 255 characters".to_string(),
                ));
            }
        }
        Ok(())
    }
}

/// Typed MS-OFFMACRO2 Word VBA supplemental data.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct VbaSupplementalData {
    document_events: Vec<VbaDocumentEvent>,
    macros: Vec<VbaMacroDescriptor>,
}

impl VbaSupplementalData {
    /// Create empty supplemental data.
    pub fn new() -> Self {
        Self::default()
    }

    /// Mark a document event active. Repeated events are coalesced.
    pub fn add_document_event(&mut self, event: VbaDocumentEvent) -> &mut Self {
        if !self.document_events.contains(&event) {
            self.document_events.push(event);
        }
        self
    }

    /// Add one macro descriptor in stored order.
    pub fn add_macro(&mut self, descriptor: VbaMacroDescriptor) -> Result<&mut Self> {
        descriptor.validate()?;
        if self.macros.len() >= MAX_SUPPLEMENTAL_MACROS {
            return Err(Error::InvalidFormat(format!(
                "Word VBA supplemental data exceeds {MAX_SUPPLEMENTAL_MACROS} macros"
            )));
        }
        self.macros.push(descriptor);
        Ok(self)
    }

    /// Return active document events in insertion order.
    pub fn document_events(&self) -> &[VbaDocumentEvent] {
        &self.document_events
    }

    /// Return macro descriptors in stored order.
    pub fn macros(&self) -> &[VbaMacroDescriptor] {
        &self.macros
    }

    /// Serialize deterministic, schema-shaped `vbaData.xml`.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut xml = String::with_capacity(256 + self.macros.len().saturating_mul(160));
        xml.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        xml.push_str("<wne:vbaSuppData xmlns:wne=\"");
        xml.push_str(WORD_VBA_NAMESPACE);
        xml.push_str("\">");
        if !self.document_events.is_empty() {
            xml.push_str("<wne:docEvents>");
            for event in VbaDocumentEvent::ALL {
                if self.document_events.contains(&event) {
                    xml.push_str("<wne:");
                    xml.push_str(event.element_name());
                    xml.push_str("/>");
                }
            }
            xml.push_str("</wne:docEvents>");
        }
        if !self.macros.is_empty() {
            xml.push_str("<wne:mcds>");
            for descriptor in &self.macros {
                descriptor.validate()?;
                xml.push_str("<wne:mcd wne:macroName=\"");
                xml.push_str(&escape_xml(&descriptor.name.to_uppercase()));
                xml.push_str("\" wne:name=\"");
                xml.push_str(&escape_xml(&descriptor.name));
                if let Some(menu_help) = &descriptor.menu_help {
                    xml.push_str("\" wne:menuHelp=\"");
                    xml.push_str(&escape_xml(menu_help));
                }
                xml.push_str("\" wne:bEncrypt=\"00\" wne:cmg=\"56\"/>");
                if xml.len() > MAX_SUPPLEMENTAL_XML_BYTES {
                    return Err(Error::InvalidFormat(format!(
                        "Word VBA supplemental XML exceeds {MAX_SUPPLEMENTAL_XML_BYTES} bytes"
                    )));
                }
            }
            xml.push_str("</wne:mcds>");
        }
        xml.push_str("</wne:vbaSuppData>");
        if xml.len() > MAX_SUPPLEMENTAL_XML_BYTES {
            return Err(Error::InvalidFormat(format!(
                "Word VBA supplemental XML exceeds {MAX_SUPPLEMENTAL_XML_BYTES} bytes"
            )));
        }
        Ok(xml.into_bytes())
    }
}

/// Relationship metadata for the VBA project attached to a Word main document.
///
/// This describes the MS-OFFMACRO2 package topology only. The `vbaProject.bin`
/// payload and Word VBA supplemental-data XML remain inert.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaProject {
    source_part_name: PackURI,
    project_relationship_id: String,
    project_part_name: PackURI,
    supplemental_data_relationship_id: String,
    supplemental_data_part_name: PackURI,
}

impl VbaProject {
    /// Return the Word main part that owns the VBA-project relationship.
    pub fn source_part_name(&self) -> &PackURI {
        &self.source_part_name
    }

    /// Return the relationship ID from the main part to the VBA Project part.
    pub fn project_relationship_id(&self) -> &str {
        &self.project_relationship_id
    }

    /// Return the absolute OPC part name of the VBA Project binary part.
    pub fn project_part_name(&self) -> &PackURI {
        &self.project_part_name
    }

    /// Return the relationship ID from the VBA Project to Word supplemental data.
    pub fn supplemental_data_relationship_id(&self) -> &str {
        &self.supplemental_data_relationship_id
    }

    /// Return the absolute OPC part name of the Word VBA supplemental-data part.
    pub fn supplemental_data_part_name(&self) -> &PackURI {
        &self.supplemental_data_part_name
    }

    /// Parse the `vbaProject.bin` payload with default resource limits.
    pub fn project(&self, package: &OpcPackage) -> Result<Project> {
        self.project_with(package, &Limits::default())
    }

    /// Parse the `vbaProject.bin` payload with explicit resource limits.
    ///
    /// The relationship graph remains independently inspectable through this
    /// type. Parsing only decompresses and decodes source; it never compiles,
    /// interprets, or executes VBA.
    pub fn project_with(&self, package: &OpcPackage, limits: &Limits) -> Result<Project> {
        Ok(read_project_part(package, &self.project_part_name, limits)?)
    }
}

/// Discover one structurally conforming Word VBA-project relationship graph.
///
/// MS-OFFMACRO2 permits at most one VBA Project relationship from a Word main
/// document. Its binary project part must in turn have exactly one relationship
/// to the Word VBA Supplemental Data part. Both payloads stay opaque here.
pub(crate) fn discover_vba_project(
    package: &OpcPackage,
    source: &dyn Part,
) -> Result<Option<VbaProject>> {
    let mut projects = source
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == relationship_type::VBA_PROJECT);
    let Some(project_relationship) = projects.next() else {
        return Ok(None);
    };
    if projects.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "Word main part '{}' has multiple VBA Project relationships",
            source.partname().as_str()
        )));
    }
    if project_relationship.is_external() {
        return Err(Error::InvalidFormat(format!(
            "VBA Project relationship '{}' from '{}' cannot be external",
            project_relationship.r_id(),
            source.partname().as_str()
        )));
    }

    let project_part_name = project_relationship.target_partname().map_err(|error| {
        Error::InvalidFormat(format!(
            "invalid VBA Project relationship '{}' from '{}': {error}",
            project_relationship.r_id(),
            source.partname().as_str()
        ))
    })?;
    let project_part = package.get_part(&project_part_name).map_err(|error| {
        Error::PartNotFound(format!(
            "VBA Project target '{}' from '{}': {error}",
            project_part_name.as_str(),
            source.partname().as_str()
        ))
    })?;
    if project_part.content_type() != content_type::OFC_VBA_PROJECT {
        return Err(Error::InvalidContentType {
            expected: content_type::OFC_VBA_PROJECT.to_string(),
            got: project_part.content_type().to_string(),
        });
    }
    if project_part
        .rels()
        .iter()
        .any(|relationship| relationship.reltype() != relationship_type::WORD_VBA_DATA)
    {
        return Err(Error::InvalidFormat(format!(
            "VBA Project part '{}' has a forbidden relationship",
            project_part_name.as_str()
        )));
    }

    let mut supplemental_data = project_part
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == relationship_type::WORD_VBA_DATA);
    let Some(supplemental_data_relationship) = supplemental_data.next() else {
        return Err(Error::InvalidFormat(format!(
            "VBA Project part '{}' is missing its Word VBA Supplemental Data relationship",
            project_part_name.as_str()
        )));
    };
    if supplemental_data.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "VBA Project part '{}' has multiple Word VBA Supplemental Data relationships",
            project_part_name.as_str()
        )));
    }
    if supplemental_data_relationship.is_external() {
        return Err(Error::InvalidFormat(format!(
            "Word VBA Supplemental Data relationship '{}' from '{}' cannot be external",
            supplemental_data_relationship.r_id(),
            project_part_name.as_str()
        )));
    }

    let supplemental_data_part_name =
        supplemental_data_relationship
            .target_partname()
            .map_err(|error| {
                Error::InvalidFormat(format!(
                    "invalid Word VBA Supplemental Data relationship '{}' from '{}': {error}",
                    supplemental_data_relationship.r_id(),
                    project_part_name.as_str()
                ))
            })?;
    let supplemental_data_part =
        package
            .get_part(&supplemental_data_part_name)
            .map_err(|error| {
                Error::PartNotFound(format!(
                    "Word VBA Supplemental Data target '{}' from '{}': {error}",
                    supplemental_data_part_name.as_str(),
                    project_part_name.as_str()
                ))
            })?;
    if supplemental_data_part.content_type() != content_type::WML_VBA_DATA {
        return Err(Error::InvalidContentType {
            expected: content_type::WML_VBA_DATA.to_string(),
            got: supplemental_data_part.content_type().to_string(),
        });
    }

    Ok(Some(VbaProject {
        source_part_name: source.partname().clone(),
        project_relationship_id: project_relationship.r_id().to_string(),
        project_part_name,
        supplemental_data_relationship_id: supplemental_data_relationship.r_id().to_string(),
        supplemental_data_part_name,
    }))
}

pub(crate) fn store_vba_project(
    package: &mut OpcPackage,
    source: &PackURI,
    payload: Arc<Vec<u8>>,
    supplemental_xml: Vec<u8>,
) -> Result<VbaProject> {
    store_project_graph(package, source, Host::Word, payload, Some(supplemental_xml))?;
    let source = package.get_part(source)?;
    discover_vba_project(package, source)?.ok_or_else(|| {
        Error::InvalidFormat("stored Word VBA project was not discoverable".to_string())
    })
}

pub(crate) fn matching_vba_project(
    package: &OpcPackage,
    source: &PackURI,
    payload: &[u8],
    supplemental_xml: &[u8],
) -> Result<Option<VbaProject>> {
    let source_part = package.get_part(source)?;
    let Some(project) = discover_vba_project(package, source_part)? else {
        return Ok(None);
    };
    if package.get_part(project.project_part_name())?.blob() != payload
        || package
            .get_part(project.supplemental_data_part_name())?
            .blob()
            != supplemental_xml
    {
        return Ok(None);
    }
    Ok(Some(project))
}

pub(crate) fn remove_vba_project(package: &mut OpcPackage, source: &PackURI) -> Result<bool> {
    Ok(remove_project_graph(package, source, Host::Word)?)
}

fn validate_supplemental_string(value: &str, field: &str) -> Result<()> {
    if value.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "Word VBA {field} must not be empty"
        )));
    }
    if value.chars().any(|character| {
        let scalar = character as u32;
        !matches!(character, '\u{9}' | '\u{a}' | '\u{d}')
            && (character <= '\u{1f}' || (scalar & 0xffff) == 0xfffe || (scalar & 0xffff) == 0xffff)
    }) {
        return Err(Error::InvalidFormat(format!(
            "Word VBA {field} contains a character forbidden by XML 1.0"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::Package;
    use litchi_opc::part::BlobPart;
    use litchi_vba::{Limits, Payload, build};
    use std::io::Cursor;

    fn package_with_vba_project(
        main_content_type: &str,
        include_supplemental_data: bool,
    ) -> OpcPackage {
        let document_name = PackURI::new("/word/document.xml").unwrap();
        let project_name = PackURI::new("/word/vbaProject.bin").unwrap();
        let supplemental_name = PackURI::new("/word/vbaData.xml").unwrap();

        let mut document = BlobPart::new(
            document_name,
            main_content_type.to_string(),
            b"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body/></w:document>".to_vec(),
        );
        document.relate_to("vbaProject.bin", relationship_type::VBA_PROJECT);

        let mut project = BlobPart::new(
            project_name,
            content_type::OFC_VBA_PROJECT.to_string(),
            b"intentionally not a compound file".to_vec(),
        );
        if include_supplemental_data {
            project.relate_to("vbaData.xml", relationship_type::WORD_VBA_DATA);
        }

        let supplemental_data = BlobPart::new(
            supplemental_name,
            content_type::WML_VBA_DATA.to_string(),
            b"intentionally not XML".to_vec(),
        );

        let mut package = OpcPackage::new();
        package.add_part(Box::new(document));
        package.add_part(Box::new(project));
        package.add_part(Box::new(supplemental_data));
        package.relate_to("word/document.xml", relationship_type::OFFICE_DOCUMENT);
        package
    }

    #[test]
    fn discovers_macro_project_metadata_without_parsing_payloads() {
        let package = package_with_vba_project(content_type::WML_DOCUMENT_MACRO_MAIN, true);
        let source = package.main_document_part().unwrap();

        let project = discover_vba_project(&package, source).unwrap().unwrap();
        assert_eq!(project.source_part_name().as_str(), "/word/document.xml");
        assert_eq!(project.project_relationship_id(), "rId1");
        assert_eq!(project.project_part_name().as_str(), "/word/vbaProject.bin");
        assert_eq!(project.supplemental_data_relationship_id(), "rId1");
        assert_eq!(
            project.supplemental_data_part_name().as_str(),
            "/word/vbaData.xml"
        );
    }

    #[test]
    fn rejects_a_vba_project_without_required_supplemental_data() {
        let package = package_with_vba_project(content_type::WML_DOCUMENT_MACRO_MAIN, false);
        let source = package.main_document_part().unwrap();

        assert!(discover_vba_project(&package, source).is_err());
    }

    #[test]
    fn documents_without_vba_projects_return_no_metadata() {
        let package = Package::new().unwrap();

        assert!(package.vba().unwrap().is_none());
    }

    #[test]
    fn docx_package_accepts_macro_enabled_document_and_template_main_parts() {
        for main_content_type in [
            content_type::WML_DOCUMENT_MACRO_MAIN,
            content_type::WML_TEMPLATE_MACRO_MAIN,
        ] {
            let package =
                Package::from_opc_package(package_with_vba_project(main_content_type, true))
                    .unwrap();
            let project = package.vba().unwrap().unwrap();

            assert_eq!(project.project_part_name().as_str(), "/word/vbaProject.bin");
            assert!(package.document().is_ok());
        }
    }

    fn authored_project() -> build::Project {
        build::Project::new("WordProject").module(build::Module::standard(
            "Module1",
            "Public Sub Hello()\r\nEnd Sub\r\n",
        ))
    }

    #[test]
    fn serializes_typed_supplemental_data_in_schema_order() {
        let mut supplemental = VbaSupplementalData::new();
        supplemental
            .add_document_event(VbaDocumentEvent::Close)
            .add_document_event(VbaDocumentEvent::Open);
        supplemental
            .add_macro(
                VbaMacroDescriptor::new("Project.Module1.Hello")
                    .unwrap()
                    .with_menu_help("Run & greet")
                    .unwrap(),
            )
            .unwrap();

        let xml = String::from_utf8(supplemental.to_xml().unwrap()).unwrap();
        assert!(xml.find("eventDocOpen").unwrap() < xml.find("eventDocClose").unwrap());
        assert!(xml.contains("macroName=\"PROJECT.MODULE1.HELLO\""));
        assert!(xml.contains("name=\"Project.Module1.Hello\""));
        assert!(xml.contains("menuHelp=\"Run &amp; greet\""));
        assert!(xml.contains("bEncrypt=\"00\" wne:cmg=\"56\""));
    }

    #[test]
    fn stores_preserves_and_removes_word_project_graph() {
        let mut package = Package::new().unwrap();
        let project = authored_project();
        let mut supplemental = VbaSupplementalData::new();
        supplemental.add_document_event(VbaDocumentEvent::Open);
        supplemental
            .add_macro(VbaMacroDescriptor::new("WordProject.Module1.Hello").unwrap())
            .unwrap();

        package
            .set_vba_with(project, &supplemental, &Limits::default())
            .unwrap();
        package
            .document_mut()
            .unwrap()
            .add_paragraph_with_text("materialize");
        let mut bytes = Cursor::new(Vec::new());
        package.to_stream(&mut bytes).unwrap();

        bytes.set_position(0);
        let mut reopened = Package::from_reader(bytes).unwrap();
        let metadata = reopened.vba().unwrap().unwrap();
        let parsed = metadata.project(reopened.opc_package()).unwrap();
        assert_eq!(parsed.name(), "WordProject");
        assert_eq!(
            reopened
                .opc_package()
                .main_document_part()
                .unwrap()
                .content_type(),
            content_type::WML_DOCUMENT_MACRO_MAIN
        );
        let supplemental_part = reopened
            .opc_package()
            .get_part(metadata.supplemental_data_part_name())
            .unwrap();
        assert!(
            std::str::from_utf8(supplemental_part.blob())
                .unwrap()
                .contains("eventDocOpen")
        );

        assert!(reopened.clear_vba().unwrap());
        assert!(reopened.vba().unwrap().is_none());
        assert_eq!(
            reopened
                .opc_package()
                .main_document_part()
                .unwrap()
                .content_type(),
            content_type::WML_DOCUMENT_MAIN
        );
        assert!(
            reopened
                .opc_package()
                .get_part(&PackURI::new("/word/vbaData.xml").unwrap())
                .is_err()
        );
    }

    #[test]
    fn rejects_invalid_project_and_supplemental_values_atomically() {
        assert!(VbaMacroDescriptor::new("").is_err());
        assert!(VbaMacroDescriptor::new("bad\u{1}name").is_err());

        let package = Package::new().unwrap();
        assert!(Payload::read(vec![0; 64], &Limits::default()).is_err());
        assert!(package.vba().unwrap().is_none());
        assert_eq!(
            package
                .opc_package()
                .main_document_part()
                .unwrap()
                .content_type(),
            content_type::WML_DOCUMENT_MAIN
        );
    }

    #[test]
    fn template_kind_survives_attach_and_remove() {
        let mut package = Package::new().unwrap();
        let source = package
            .opc_package()
            .main_document_part()
            .unwrap()
            .partname()
            .clone();
        package
            .edit_opc(|opc| {
                opc.get_part_mut(&source)?
                    .set_content_type(content_type::WML_TEMPLATE_MAIN.to_string())?;
                Ok(())
            })
            .unwrap();

        package.set_vba(authored_project()).unwrap();
        assert_eq!(
            package
                .opc_package()
                .get_part(&source)
                .unwrap()
                .content_type(),
            content_type::WML_TEMPLATE_MACRO_MAIN
        );
        package.clear_vba().unwrap();
        assert_eq!(
            package
                .opc_package()
                .get_part(&source)
                .unwrap()
                .content_type(),
            content_type::WML_TEMPLATE_MAIN
        );
    }
}
