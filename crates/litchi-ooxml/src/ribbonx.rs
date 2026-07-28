//! Inert RibbonX customization package parts.
//!
//! RibbonX XML can name callbacks implemented by the host application. This
//! module preserves that declarative XML as bounded bytes, but never resolves,
//! invokes, or otherwise interprets callbacks, macros, images, or commands.

use crate::error::{OoxmlError, Result};
use litchi_opc::{OpcPackage, PackURI, Part, XmlPart};
use quick_xml::{
    events::Event,
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

/// Namespace for the Office 2007 Custom UI markup variant.
pub const RIBBONX_2007_NAMESPACE: &str = "http://schemas.microsoft.com/office/2006/01/customui";
/// Namespace for the Office 2010-and-later Custom UI markup variant.
pub const RIBBONX_2010_NAMESPACE: &str = "http://schemas.microsoft.com/office/2009/07/customui";
/// Legacy CustomUI2 namespace documented by MS-OI29500.
pub const RIBBONX_2007_10_NAMESPACE: &str = "http://schemas.microsoft.com/office/2007/10/customui";

/// Package relationship for the Office 2007 Custom UI markup variant.
pub const RIBBONX_2007_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility";
/// Package relationship shared by the newer Custom UI markup variants.
pub const RIBBONX_2010_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility";

/// Content type mandated for all supported RibbonX customization parts.
pub const RIBBONX_CONTENT_TYPE: &str = "application/xml";

/// Largest accepted RibbonX XML part.
pub const MAX_RIBBONX_XML_BYTES: usize = 4 * 1024 * 1024;
/// Largest accepted element nesting depth in a RibbonX XML part.
pub const MAX_RIBBONX_XML_DEPTH: usize = 128;

/// RibbonX Custom UI markup version identified by its root namespace and package relationship.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RibbonCustomizationVersion {
    /// The Office 2007 Custom UI markup vocabulary.
    Office2007,
    /// The Office 2010-and-later Ribbon and Backstage markup vocabulary.
    Office2010,
    /// The CustomUI2 root namespace documented by MS-OI29500.
    CustomUi2,
}

impl RibbonCustomizationVersion {
    /// Root namespace required by this markup version.
    pub const fn namespace(self) -> &'static str {
        match self {
            Self::Office2007 => RIBBONX_2007_NAMESPACE,
            Self::Office2010 => RIBBONX_2010_NAMESPACE,
            Self::CustomUi2 => RIBBONX_2007_10_NAMESPACE,
        }
    }

    /// Root-package relationship type required by this markup version.
    pub const fn relationship_type(self) -> &'static str {
        match self {
            Self::Office2007 => RIBBONX_2007_RELATIONSHIP_TYPE,
            Self::Office2010 | Self::CustomUi2 => RIBBONX_2010_RELATIONSHIP_TYPE,
        }
    }

    fn default_part_path(self) -> &'static str {
        match self {
            Self::Office2007 => "/customUI/customUI.xml",
            Self::Office2010 => "/customUI/customUI14.xml",
            Self::CustomUi2 => "/customUI/customUI2.xml",
        }
    }
}

/// A package-level RibbonX customization retained as opaque XML bytes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RibbonCustomization {
    part_name: PackURI,
    relationship_id: String,
    version: RibbonCustomizationVersion,
    xml: Vec<u8>,
}

impl RibbonCustomization {
    /// Absolute OPC part name of the Custom UI XML.
    pub fn part_name(&self) -> &PackURI {
        &self.part_name
    }

    /// Package-level relationship ID targeting this customization.
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Markup vocabulary identified by the root namespace and package relationship.
    pub fn version(&self) -> RibbonCustomizationVersion {
        self.version
    }

    /// Original Custom UI XML bytes.
    ///
    /// The bytes are not parsed into controls and callback names remain inert.
    pub fn xml(&self) -> &[u8] {
        &self.xml
    }
}

/// Load all package-level RibbonX customizations.
///
/// The legacy and newer relationship families may coexist, with at most one
/// part in each family. The XML is structurally validated only through its
/// document root and safe parsing boundaries.
pub fn load_ribbon_customizations(package: &OpcPackage) -> Result<Vec<RibbonCustomization>> {
    find_ribbon_relationships(package)?
        .into_iter()
        .map(|relationship| {
            let part = package.get_part(&relationship.part_name).map_err(|error| {
                OoxmlError::PartNotFound(format!(
                    "RibbonX part '{}' does not exist: {error}",
                    relationship.part_name.as_str()
                ))
            })?;
            require_content_type(part)?;
            let version = validate_ribbon_xml(part.blob(), relationship.kind)?;
            Ok(RibbonCustomization {
                part_name: relationship.part_name,
                relationship_id: relationship.relationship_id,
                version,
                xml: part.blob().to_vec(),
            })
        })
        .collect()
}

/// Load the effective RibbonX customization, if present.
///
/// When a package contains both the legacy and newer relationship families,
/// this returns the newer one, matching the precedence documented for
/// CustomUI and CustomUI2 package parts.
pub fn load_ribbon_customization(package: &OpcPackage) -> Result<Option<RibbonCustomization>> {
    Ok(load_ribbon_customizations(package)?.pop())
}

/// Store one package-level RibbonX customization as opaque XML.
///
/// An existing part in the same package-relationship family is updated in
/// place. A customization in the other family remains intact, so legacy and
/// newer Custom UI parts can coexist as the package specifications permit.
pub fn store_ribbon_customization(
    package: &mut OpcPackage,
    version: RibbonCustomizationVersion,
    xml: &[u8],
) -> Result<RibbonCustomization> {
    let kind = RibbonRelationshipKind::for_version(version);
    let parsed_version = validate_ribbon_xml(xml, kind)?;
    if parsed_version != version {
        return invalid(format!(
            "RibbonX XML root namespace does not match requested {version:?} version"
        ));
    }

    if let Some(existing) = find_ribbon_relationships(package)?
        .into_iter()
        .find(|relationship| relationship.kind == kind)
    {
        {
            let part = package.get_part_mut(&existing.part_name).map_err(|error| {
                OoxmlError::PartNotFound(format!(
                    "RibbonX part '{}' does not exist: {error}",
                    existing.part_name.as_str()
                ))
            })?;
            require_content_type(part)?;
            part.set_blob(xml.to_vec());
        }

        return Ok(RibbonCustomization {
            part_name: existing.part_name,
            relationship_id: existing.relationship_id,
            version,
            xml: xml.to_vec(),
        });
    }

    let part_name = add_new_ribbon_part(package, version, xml)?;
    let relationship_id = package.relate_to(
        part_name.as_str().trim_start_matches('/'),
        version.relationship_type(),
    );
    Ok(RibbonCustomization {
        part_name,
        relationship_id,
        version,
        xml: xml.to_vec(),
    })
}

#[derive(Debug)]
struct RibbonRelationship {
    part_name: PackURI,
    relationship_id: String,
    kind: RibbonRelationshipKind,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RibbonRelationshipKind {
    Legacy,
    Newer,
}

impl RibbonRelationshipKind {
    fn for_version(version: RibbonCustomizationVersion) -> Self {
        match version {
            RibbonCustomizationVersion::Office2007 => Self::Legacy,
            RibbonCustomizationVersion::Office2010 | RibbonCustomizationVersion::CustomUi2 => {
                Self::Newer
            },
        }
    }
}

fn find_ribbon_relationships(package: &OpcPackage) -> Result<Vec<RibbonRelationship>> {
    let mut legacy = None;
    let mut newer = None;
    for relationship in package.rels().iter() {
        let Some(kind) = ribbon_relationship_kind(relationship.reltype()) else {
            continue;
        };
        if relationship.is_external() {
            return invalid("RibbonX customization relationship must be internal".into());
        }
        let part_name = relationship.target_partname().map_err(|error| {
            OoxmlError::InvalidRelationship(format!("invalid RibbonX relationship target: {error}"))
        })?;
        let customization = RibbonRelationship {
            part_name,
            relationship_id: relationship.r_id().to_string(),
            kind,
        };
        let previous = match kind {
            RibbonRelationshipKind::Legacy => legacy.replace(customization),
            RibbonRelationshipKind::Newer => newer.replace(customization),
        };
        if previous.is_some() {
            return invalid(format!(
                "a package may contain at most one {} RibbonX customization",
                match kind {
                    RibbonRelationshipKind::Legacy => "legacy",
                    RibbonRelationshipKind::Newer => "newer",
                }
            ));
        }
    }

    Ok(legacy.into_iter().chain(newer).collect())
}

fn ribbon_relationship_kind(relationship_type: &str) -> Option<RibbonRelationshipKind> {
    match relationship_type {
        RIBBONX_2007_RELATIONSHIP_TYPE => Some(RibbonRelationshipKind::Legacy),
        RIBBONX_2010_RELATIONSHIP_TYPE => Some(RibbonRelationshipKind::Newer),
        _ => None,
    }
}

fn require_content_type(part: &dyn Part) -> Result<()> {
    if part.content_type() == RIBBONX_CONTENT_TYPE {
        return Ok(());
    }
    Err(OoxmlError::InvalidContentType {
        expected: RIBBONX_CONTENT_TYPE.to_string(),
        got: part.content_type().to_string(),
    })
}

fn add_new_ribbon_part(
    package: &mut OpcPackage,
    version: RibbonCustomizationVersion,
    xml: &[u8],
) -> Result<PackURI> {
    for suffix in 0..10_000usize {
        let path = if suffix == 0 {
            version.default_part_path().to_string()
        } else {
            format!("/customUI/customUI{suffix}.xml")
        };
        let part_name = PackURI::new(&path).map_err(|error| {
            OoxmlError::InvalidUri(format!("RibbonX part URI '{path}': {error}"))
        })?;
        let part = XmlPart::new(
            part_name.clone(),
            RIBBONX_CONTENT_TYPE.to_string(),
            xml.to_vec(),
        );
        if package.try_add_part(Box::new(part)).is_ok() {
            return Ok(part_name);
        }
    }
    invalid("unable to allocate a unique RibbonX part name".into())
}

fn validate_ribbon_xml(
    xml: &[u8],
    relationship_kind: RibbonRelationshipKind,
) -> Result<RibbonCustomizationVersion> {
    if xml.len() > MAX_RIBBONX_XML_BYTES {
        return invalid(format!(
            "RibbonX XML exceeds {MAX_RIBBONX_XML_BYTES} byte limit"
        ));
    }

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut closed_root = false;
    let mut version = None;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| OoxmlError::Xml(format!("RibbonX XML error: {error}")))?;
        let namespace = bound_namespace(&namespace).map(<[u8]>::to_vec);
        let namespace = namespace.as_deref();
        let event = event.into_owned();

        match event {
            Event::Start(element) => {
                if depth == 0 {
                    let root_version =
                        validate_root(namespace, element.local_name().as_ref(), relationship_kind)?;
                    if saw_root || closed_root {
                        return invalid("RibbonX XML has multiple document roots".into());
                    }
                    saw_root = true;
                    version = Some(root_version);
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("RibbonX XML nesting overflow".into())
                })?;
                if depth > MAX_RIBBONX_XML_DEPTH {
                    return invalid(format!("RibbonX XML depth exceeds {MAX_RIBBONX_XML_DEPTH}"));
                }
            },
            Event::Empty(element) if depth == 0 => {
                let root_version =
                    validate_root(namespace, element.local_name().as_ref(), relationship_kind)?;
                if saw_root || closed_root {
                    return invalid("RibbonX XML has multiple document roots".into());
                }
                saw_root = true;
                closed_root = true;
                version = Some(root_version);
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("RibbonX XML has an unexpected end tag".into())
                })?;
                if depth == 0 {
                    closed_root = true;
                }
            },
            Event::DocType(_) => return invalid("DTD is forbidden in RibbonX XML".into()),
            Event::GeneralRef(reference) => {
                let reference = reference.as_ref();
                let predefined = matches!(reference, b"lt" | b"gt" | b"amp" | b"apos" | b"quot");
                if !predefined && !reference.starts_with(b"#") {
                    return invalid("RibbonX XML contains a non-predefined entity".into());
                }
            },
            Event::Text(text) if depth == 0 && !is_xml_whitespace(text.as_ref()) => {
                return invalid("RibbonX XML has text outside its document root".into());
            },
            Event::CData(text) if depth == 0 && !is_xml_whitespace(text.as_ref()) => {
                return invalid("RibbonX XML has CDATA outside its document root".into());
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }

    if !saw_root || !closed_root || depth != 0 {
        return invalid("RibbonX XML has no complete customUI root".into());
    }
    version.ok_or_else(|| OoxmlError::InvalidFormat("RibbonX XML has no customUI root".into()))
}

fn validate_root(
    namespace: Option<&[u8]>,
    local_name: &[u8],
    relationship_kind: RibbonRelationshipKind,
) -> Result<RibbonCustomizationVersion> {
    if local_name != b"customUI" {
        return invalid("RibbonX XML root must be customUI".into());
    }
    match relationship_kind {
        RibbonRelationshipKind::Legacy if namespace == Some(RIBBONX_2007_NAMESPACE.as_bytes()) => {
            Ok(RibbonCustomizationVersion::Office2007)
        },
        RibbonRelationshipKind::Newer if namespace == Some(RIBBONX_2010_NAMESPACE.as_bytes()) => {
            Ok(RibbonCustomizationVersion::Office2010)
        },
        RibbonRelationshipKind::Newer
            if namespace == Some(RIBBONX_2007_10_NAMESPACE.as_bytes()) =>
        {
            Ok(RibbonCustomizationVersion::CustomUi2)
        },
        _ => invalid("RibbonX XML root namespace does not match its package relationship".into()),
    }
}

fn bound_namespace<'a>(namespace: &'a ResolveResult<'a>) -> Option<&'a [u8]> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => Some(*value),
        _ => None,
    }
}

fn is_xml_whitespace(value: &[u8]) -> bool {
    value
        .iter()
        .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
}

fn invalid<T>(message: String) -> Result<T> {
    Err(OoxmlError::InvalidFormat(message))
}
