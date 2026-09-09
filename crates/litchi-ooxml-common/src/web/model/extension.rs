use super::super::codec::{
    Node, XmlDocument, canonical_node_xml, invalid, is_drawingml_namespace, parse_xml,
    parse_xml_owned, reject_unknown_attributes,
};
use super::super::{
    DRAWINGML_NAMESPACE, Result, STRICT_DRAWINGML_NAMESPACE, TASK_PANES_NAMESPACE,
    WEB_EXTENSION_NAMESPACE,
};
use super::custom_functions::{
    CustomFunctions, custom_source_compatible, parse_custom_functions_with_limits,
    rewrite_custom_functions, validate_custom_functions,
};
use super::{Limits, MAX_WEB_EXTENSION_XML_BYTES};
/// Namespace dialect of an MS-OWEXML extension-list element.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExtKind {
    AddIn,
    TaskPane,
    DrawingMl,
    StrictDrawingMl,
}

impl ExtKind {
    #[must_use]
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
    pub(in crate::web) custom_functions: Option<CustomFunctions>,
}

impl ExtList {
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        Self::from_xml_with_limits(xml, &Limits::standard())
    }

    pub(in crate::web) fn from_xml_with_limits(xml: &[u8], limits: &Limits) -> Result<Self> {
        if xml.len() > limits.xml_bytes {
            return invalid(format!(
                "web extension extLst XML exceeds {} bytes",
                limits.xml_bytes
            ));
        }
        let document = parse_xml_owned(xml.to_vec(), limits)?;
        Self::from_node_with_limits(document.root()?, &document, limits)
    }

    /// Construct an empty extension list in one of the supported namespaces.
    ///
    /// The list remains inert until a caller explicitly adds a typed payload or
    /// supplies an opaque extension entry.
    pub fn empty(kind: ExtKind) -> Result<Self> {
        let (prefix, namespace) = match kind {
            ExtKind::AddIn => ("we", WEB_EXTENSION_NAMESPACE),
            ExtKind::TaskPane => ("wetp", TASK_PANES_NAMESPACE),
            ExtKind::DrawingMl => ("a", DRAWINGML_NAMESPACE),
            ExtKind::StrictDrawingMl => ("a", STRICT_DRAWINGML_NAMESPACE),
        };
        let xml = format!(r#"<{prefix}:extLst xmlns:{prefix}="{namespace}"/>"#);
        Self::from_xml(xml.as_bytes())
    }

    #[must_use]
    pub fn kind(&self) -> ExtKind {
        self.kind
    }

    #[must_use]
    pub fn as_xml(&self) -> &[u8] {
        self.xml.as_bytes()
    }

    #[must_use]
    pub fn xml(&self) -> &str {
        &self.xml
    }

    /// Return the inert `[MS-OWEXML]` custom-function/background metadata, if
    /// present in a direct OfficeArt extension entry.
    #[must_use]
    pub fn custom_functions(&self) -> Option<&CustomFunctions> {
        self.custom_functions.as_ref()
    }

    /// Replace the typed custom-function/background metadata while preserving
    /// unknown extension payloads and their source bytes.
    ///
    /// The metadata is valid only on the web-extension (`we:extLst`) owner.
    /// This operation never activates a runtime, invokes a callback, or
    /// executes a custom function.
    pub fn set_custom_functions(&mut self, value: Option<CustomFunctions>) -> Result<&mut Self> {
        self.set_custom_functions_with_limits(value, &Limits::standard())
    }

    pub(in crate::web) fn set_custom_functions_with_limits(
        &mut self,
        value: Option<CustomFunctions>,
        limits: &Limits,
    ) -> Result<&mut Self> {
        let value = value.filter(|value| !value.is_empty());
        if self.kind != ExtKind::AddIn {
            return invalid(
                "custom-function metadata requires a web-extension extLst owner".into(),
            );
        }
        if self.custom_functions == value {
            return Ok(self);
        }
        if let Some(value) = &value {
            validate_custom_functions(value, limits)?;
        }
        let rewritten = rewrite_custom_functions(
            self.xml.as_bytes(),
            self.custom_functions.as_ref(),
            value.as_ref(),
            limits,
        )?;
        let reparsed = Self::from_xml_with_limits(rewritten.as_bytes(), limits)?;
        if reparsed.kind != self.kind || reparsed.custom_functions != value {
            return invalid("custom-function metadata rewrite was not stable".into());
        }
        self.xml = reparsed.xml;
        self.custom_functions = reparsed.custom_functions;
        Ok(self)
    }

    /// Remove typed custom-function/background metadata, retaining surrounding
    /// extension entries and vendor payloads.
    pub fn clear_custom_functions(&mut self) -> Result<&mut Self> {
        self.set_custom_functions(None)
    }

    pub(in crate::web) fn from_node_with_limits(
        node: &Node,
        document: &XmlDocument,
        limits: &Limits,
    ) -> Result<Self> {
        if node.local_name != "extLst" {
            return invalid(format!(
                "web extension extension fragment root must be extLst, got {}",
                node.local_name
            ));
        }
        reject_unknown_attributes(node, &[])?;
        let kind = ExtKind::from_namespace(&node.namespace)?;
        let xml = document.self_contained_fragment_with_limits(node, limits)?;
        let custom_functions = if kind == ExtKind::AddIn {
            parse_custom_functions_with_limits(xml.as_bytes(), limits)?
        } else {
            None
        };
        Ok(Self {
            kind,
            xml,
            custom_functions,
        })
    }

    pub(in crate::web) fn source_compatible_with(
        &self,
        other: &Self,
        limits: &Limits,
    ) -> Result<bool> {
        if self.kind != other.kind {
            return Ok(false);
        }
        custom_source_compatible(self.xml.as_bytes(), other.xml.as_bytes(), limits)
    }
}

/// Compression state of a `DrawingML` `CT_Blip`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Compression {
    Email,
    Screen,
    Print,
    HighQualityPrint,
    None,
}

impl Compression {
    #[must_use]
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

/// Closed effect-element choice allowed by `DrawingML` `CT_Blip`.
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
    #[must_use]
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

/// A validated, inert `DrawingML` effect subtree.
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
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_WEB_EXTENSION_XML_BYTES {
            return invalid(format!(
                "snapshot effect XML exceeds {MAX_WEB_EXTENSION_XML_BYTES} bytes"
            ));
        }
        let document = parse_xml(xml)?;
        Self::from_node(document.root()?)
    }

    #[must_use]
    pub fn kind(&self) -> EffectKind {
        self.kind
    }

    #[must_use]
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
