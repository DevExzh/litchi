use super::super::codec::*;
use super::super::*;
use super::*;
/// Namespace dialect of an MS-OWEXML extension-list element.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExtKind {
    AddIn,
    TaskPane,
    DrawingMl,
    StrictDrawingMl,
}

impl ExtKind {
    pub fn namespace(self) -> &'static str {
        match self {
            Self::AddIn => WEB_EXTENSION_NAMESPACE,
            Self::TaskPane => TASK_PANES_NAMESPACE,
            Self::DrawingMl => DRAWINGML_NAMESPACE,
            Self::StrictDrawingMl => STRICT_DRAWINGML_NAMESPACE,
        }
    }

    pub(in crate::web) fn from_namespace(namespace: &str) -> Result<Self> {
        match namespace {
            WEB_EXTENSION_NAMESPACE => Ok(Self::AddIn),
            TASK_PANES_NAMESPACE => Ok(Self::TaskPane),
            DRAWINGML_NAMESPACE => Ok(Self::DrawingMl),
            STRICT_DRAWINGML_NAMESPACE => Ok(Self::StrictDrawingMl),
            _ => invalid(format!(
                "invalid web extension extLst namespace '{namespace}'"
            )),
        }
    }
}

/// A bounded, self-contained, inert `extLst` fragment.
///
/// Unknown extension payloads are retained without interpretation or resource
/// resolution. Namespace declarations inherited by the source fragment are
/// materialized on its root so it remains valid when authored elsewhere.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExtList {
    pub(in crate::web) kind: ExtKind,
    pub(in crate::web) xml: String,
}

impl ExtList {
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_WEB_EXTENSION_XML_BYTES {
            return invalid(format!(
                "web extension extLst XML exceeds {MAX_WEB_EXTENSION_XML_BYTES} bytes"
            ));
        }
        let document = parse_xml(xml)?;
        Self::from_node(document.root()?, &document)
    }

    pub fn kind(&self) -> ExtKind {
        self.kind
    }

    pub fn as_xml(&self) -> &[u8] {
        self.xml.as_bytes()
    }

    pub fn xml(&self) -> &str {
        &self.xml
    }

    pub(in crate::web) fn from_node(node: &Node, document: &XmlDocument) -> Result<Self> {
        if node.local_name != "extLst" {
            return invalid(format!(
                "web extension extension fragment root must be extLst, got {}",
                node.local_name
            ));
        }
        reject_unknown_attributes(node, &[])?;
        let kind = ExtKind::from_namespace(&node.namespace)?;
        Ok(Self {
            kind,
            xml: document.self_contained_fragment(node)?,
        })
    }
}

/// Compression state of a DrawingML `CT_Blip`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Compression {
    Email,
    Screen,
    Print,
    HighQualityPrint,
    None,
}

impl Compression {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Email => "email",
            Self::Screen => "screen",
            Self::Print => "print",
            Self::HighQualityPrint => "hqprint",
            Self::None => "none",
        }
    }

    pub(in crate::web) fn parse(value: &str) -> Result<Self> {
        match value {
            "email" => Ok(Self::Email),
            "screen" => Ok(Self::Screen),
            "print" => Ok(Self::Print),
            "hqprint" => Ok(Self::HighQualityPrint),
            "none" => Ok(Self::None),
            _ => invalid(format!("invalid snapshot compression state '{value}'")),
        }
    }
}

/// Closed effect-element choice allowed by DrawingML `CT_Blip`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EffectKind {
    AlphaBiLevel,
    AlphaCeiling,
    AlphaFloor,
    AlphaInverse,
    AlphaModulate,
    AlphaModulateFixed,
    AlphaReplace,
    BiLevel,
    Blur,
    ColorChange,
    ColorReplace,
    Duotone,
    FillOverlay,
    Grayscale,
    HueSaturationLuminance,
    Luminance,
    Tint,
}

impl EffectKind {
    pub fn local_name(self) -> &'static str {
        match self {
            Self::AlphaBiLevel => "alphaBiLevel",
            Self::AlphaCeiling => "alphaCeiling",
            Self::AlphaFloor => "alphaFloor",
            Self::AlphaInverse => "alphaInv",
            Self::AlphaModulate => "alphaMod",
            Self::AlphaModulateFixed => "alphaModFix",
            Self::AlphaReplace => "alphaRepl",
            Self::BiLevel => "biLevel",
            Self::Blur => "blur",
            Self::ColorChange => "clrChange",
            Self::ColorReplace => "clrRepl",
            Self::Duotone => "duotone",
            Self::FillOverlay => "fillOverlay",
            Self::Grayscale => "grayscl",
            Self::HueSaturationLuminance => "hsl",
            Self::Luminance => "lum",
            Self::Tint => "tint",
        }
    }

    pub(in crate::web) fn parse(local_name: &str) -> Result<Self> {
        match local_name {
            "alphaBiLevel" => Ok(Self::AlphaBiLevel),
            "alphaCeiling" => Ok(Self::AlphaCeiling),
            "alphaFloor" => Ok(Self::AlphaFloor),
            "alphaInv" => Ok(Self::AlphaInverse),
            "alphaMod" => Ok(Self::AlphaModulate),
            "alphaModFix" => Ok(Self::AlphaModulateFixed),
            "alphaRepl" => Ok(Self::AlphaReplace),
            "biLevel" => Ok(Self::BiLevel),
            "blur" => Ok(Self::Blur),
            "clrChange" => Ok(Self::ColorChange),
            "clrRepl" => Ok(Self::ColorReplace),
            "duotone" => Ok(Self::Duotone),
            "fillOverlay" => Ok(Self::FillOverlay),
            "grayscl" => Ok(Self::Grayscale),
            "hsl" => Ok(Self::HueSaturationLuminance),
            "lum" => Ok(Self::Luminance),
            "tint" => Ok(Self::Tint),
            _ => invalid(format!("invalid snapshot effect '{local_name}'")),
        }
    }
}

/// A validated, inert DrawingML effect subtree.
///
/// The subtree is retained as canonical XML. It is never interpreted as
/// executable content, and construction rejects text, CDATA, DTDs, excessive
/// depth, and roots outside the closed `CT_Blip` effect choice.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Effect {
    pub(in crate::web) kind: EffectKind,
    pub(in crate::web) xml: String,
}

impl Effect {
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_WEB_EXTENSION_XML_BYTES {
            return invalid(format!(
                "snapshot effect XML exceeds {MAX_WEB_EXTENSION_XML_BYTES} bytes"
            ));
        }
        let document = parse_xml(xml)?;
        Self::from_node(document.root()?)
    }

    pub fn kind(&self) -> EffectKind {
        self.kind
    }

    pub fn xml(&self) -> &str {
        &self.xml
    }

    pub(in crate::web) fn from_node(node: &Node) -> Result<Self> {
        if !is_drawingml_namespace(&node.namespace) {
            return invalid(format!(
                "snapshot effect {} has invalid namespace '{}'",
                node.local_name, node.namespace
            ));
        }
        let kind = EffectKind::parse(&node.local_name)?;
        Ok(Self {
            kind,
            xml: canonical_node_xml(node),
        })
    }
}
