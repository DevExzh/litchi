//! Theme part authoring for PowerPoint packages.
//!
//! This module closes the authoring gap for `a:theme` parts
//! (`/ppt/theme/themeN.xml`). Every operation keeps the presentation
//! relationship graph consistent with the read side:
//!
//! - New theme parts are written with a caller-supplied color scheme
//!   (`a:clrScheme`, twelve typed slots), a caller-supplied font scheme
//!   (`a:fontScheme`, major/minor latin typefaces plus optional East-Asian,
//!   complex-script, and per-script typefaces), and a fixed default
//!   `a:fmtScheme`, then registered with the Office theme content type.
//! - Attaching a theme to a slide master adds the theme relationship the
//!   read side requires (exactly one internal theme relationship targeting
//!   an Office theme part).
//! - Replacing a color or font scheme patches the existing theme part in
//!   place with prefix-safe XML, like the master/layout author does.
//!
//! After every mutation the master/layout/theme graph is re-validated with
//! the same rules the read side applies, so an authored package resolves
//! cleanly through `Presentation::slide_masters`, `SlideMaster::theme`,
//! `SlideMaster::theme_color`, and `Presentation::get_themes`.

use crate::error::{OoxmlError, Result};
use crate::pptx::parts::ThemePart;
use litchi_core::xml::escape_xml;
use litchi_opc::OpcPackage;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::packuri::PackURI;
use litchi_opc::part::BlobPart;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::fmt::Write as FmtWrite;

const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_THEME_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/theme";
const STRICT_THEME_OVERRIDE_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/themeOverride";

/// Bounded-input ceiling for every theme part this module parses or patches.
const MAX_PART_XML_BYTES: usize = 8 * 1024 * 1024;
/// Bounded-input ceiling for XML node counts while scanning.
const MAX_SCAN_NODES: usize = 100_000;
/// Bounded-input ceiling for XML nesting depth while scanning.
const MAX_SCAN_DEPTH: usize = 128;
/// Bounded ceiling for authored theme, scheme, and typeface names.
const MAX_NAME_CHARS: usize = 256;
/// Bounded ceiling for per-script font entries authored in a single font face.
const MAX_SCRIPT_FONTS_PER_FACE: usize = 64;
/// Bounded ceiling for a script code (`ST_Script`, for example `Jpan`).
const MAX_SCRIPT_CODE_CHARS: usize = 16;
/// Depth of `a:clrScheme` and `a:fontScheme` inside a theme part
/// (`a:theme` → `a:themeElements` → scheme element).
const SCHEME_DEPTH: usize = 3;

const XML_DECL: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";

/// Default `a:fmtScheme` (fill, line, effect, and background-fill style
/// lists) written into newly authored theme parts.
///
/// Format-scheme authoring is out of scope; this is the minified Office
/// default, which is schema-valid and uses only `phClr` placeholders.
const DEFAULT_FORMAT_SCHEME_XML: &str = "<a:fmtScheme name=\"Office\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"50000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs><a:gs pos=\"35000\"><a:schemeClr val=\"phClr\"><a:tint val=\"37000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:tint val=\"15000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"16200000\" scaled=\"1\"/></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"100000\"/><a:shade val=\"100000\"/><a:satMod val=\"130000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:tint val=\"50000\"/><a:shade val=\"100000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"16200000\" scaled=\"0\"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"><a:shade val=\"95000\"/><a:satMod val=\"105000\"/></a:schemeClr></a:solidFill><a:prstDash val=\"solid\"/></a:ln><a:ln w=\"25400\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln><a:ln w=\"38100\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"20000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"38000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"23000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"35000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"23000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"35000\"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst=\"orthographicFront\"><a:rot lat=\"0\" lon=\"0\" rev=\"0\"/></a:camera><a:lightRig rig=\"threePt\" dir=\"t\"><a:rot lat=\"0\" lon=\"0\" rev=\"1200000\"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w=\"63500\" h=\"25400\"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"40000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs><a:gs pos=\"40000\"><a:schemeClr val=\"phClr\"><a:tint val=\"45000\"/><a:shade val=\"99000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"20000\"/><a:satMod val=\"255000\"/></a:schemeClr></a:gs></a:gsLst><a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"-80000\" r=\"50000\" b=\"18000\"/></a:path></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"80000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"30000\"/><a:satMod val=\"200000\"/></a:schemeClr></a:gs></a:gsLst><a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"50000\" r=\"50000\" b=\"50000\"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme>";

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

fn limit(what: &str) -> OoxmlError {
    invalid(format!("{what} exceeds the supported safety limit"))
}

// ============================================================================
// Typed enums
// ============================================================================

/// A color slot of a DrawingML theme color scheme (`CT_ColorScheme`,
/// ECMA-376 Part 1).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ThemeColorSlot {
    Dark1,
    Light1,
    Dark2,
    Light2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Hyperlink,
    FollowedHyperlink,
}

impl ThemeColorSlot {
    /// All twelve slots in schema serialization order.
    pub const ALL: [Self; 12] = [
        Self::Dark1,
        Self::Light1,
        Self::Dark2,
        Self::Light2,
        Self::Accent1,
        Self::Accent2,
        Self::Accent3,
        Self::Accent4,
        Self::Accent5,
        Self::Accent6,
        Self::Hyperlink,
        Self::FollowedHyperlink,
    ];

    /// The spec token used as the slot's element name.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Dark1 => "dk1",
            Self::Light1 => "lt1",
            Self::Dark2 => "dk2",
            Self::Light2 => "lt2",
            Self::Accent1 => "accent1",
            Self::Accent2 => "accent2",
            Self::Accent3 => "accent3",
            Self::Accent4 => "accent4",
            Self::Accent5 => "accent5",
            Self::Accent6 => "accent6",
            Self::Hyperlink => "hlink",
            Self::FollowedHyperlink => "folHlink",
        }
    }
}

/// A system color (`ST_SystemColorVal`, ECMA-376 Part 1).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum SystemColorKind {
    ActiveBorder,
    ActiveCaption,
    AppWorkspace,
    Background,
    ButtonFace,
    ButtonHighlight,
    ButtonShadow,
    ButtonText,
    CaptionText,
    GradientActiveCaption,
    GradientInactiveCaption,
    GrayText,
    Highlight,
    HighlightText,
    HotLight,
    InactiveBorder,
    InactiveCaption,
    InactiveCaptionText,
    InfoBackground,
    InfoText,
    Menu,
    MenuBar,
    MenuHighlight,
    MenuText,
    ScrollBar,
    Window,
    WindowFrame,
    WindowText,
}

impl SystemColorKind {
    /// The spec token written to the `val` attribute of `a:sysClr`.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::ActiveBorder => "activeBorder",
            Self::ActiveCaption => "activeCaption",
            Self::AppWorkspace => "appWorkspace",
            Self::Background => "background",
            Self::ButtonFace => "btnFace",
            Self::ButtonHighlight => "btnHighlight",
            Self::ButtonShadow => "btnShadow",
            Self::ButtonText => "btnText",
            Self::CaptionText => "captionText",
            Self::GradientActiveCaption => "gradientActiveCaption",
            Self::GradientInactiveCaption => "gradientInactiveCaption",
            Self::GrayText => "grayText",
            Self::Highlight => "highlight",
            Self::HighlightText => "highlightText",
            Self::HotLight => "hotLight",
            Self::InactiveBorder => "inactiveBorder",
            Self::InactiveCaption => "inactiveCaption",
            Self::InactiveCaptionText => "inactiveCaptionText",
            Self::InfoBackground => "infoBk",
            Self::InfoText => "infoText",
            Self::Menu => "menu",
            Self::MenuBar => "menuBar",
            Self::MenuHighlight => "menuHighlight",
            Self::MenuText => "menuText",
            Self::ScrollBar => "scrollBar",
            Self::Window => "window",
            Self::WindowFrame => "windowFrame",
            Self::WindowText => "windowText",
        }
    }

    /// Parse a system color token (`ST_SystemColorVal`).
    pub fn from_token(token: &str) -> Option<Self> {
        Some(match token {
            "activeBorder" => Self::ActiveBorder,
            "activeCaption" => Self::ActiveCaption,
            "appWorkspace" => Self::AppWorkspace,
            "background" => Self::Background,
            "btnFace" => Self::ButtonFace,
            "btnHighlight" => Self::ButtonHighlight,
            "btnShadow" => Self::ButtonShadow,
            "btnText" => Self::ButtonText,
            "captionText" => Self::CaptionText,
            "gradientActiveCaption" => Self::GradientActiveCaption,
            "gradientInactiveCaption" => Self::GradientInactiveCaption,
            "grayText" => Self::GrayText,
            "highlight" => Self::Highlight,
            "highlightText" => Self::HighlightText,
            "hotLight" => Self::HotLight,
            "inactiveBorder" => Self::InactiveBorder,
            "inactiveCaption" => Self::InactiveCaption,
            "inactiveCaptionText" => Self::InactiveCaptionText,
            "infoBk" => Self::InfoBackground,
            "infoText" => Self::InfoText,
            "menu" => Self::Menu,
            "menuBar" => Self::MenuBar,
            "menuHighlight" => Self::MenuHighlight,
            "menuText" => Self::MenuText,
            "scrollBar" => Self::ScrollBar,
            "window" => Self::Window,
            "windowFrame" => Self::WindowFrame,
            "windowText" => Self::WindowText,
            _ => return None,
        })
    }
}

/// A color value authored into one theme color slot: an `a:srgbClr` or an
/// `a:sysClr` choice, per `EG_ColorChoice` (color transforms are not
/// authored).
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ThemeColorValue {
    /// An sRGB color (`a:srgbClr`), as six hexadecimal digits.
    Srgb(String),
    /// A system color (`a:sysClr`) with an optional last-known sRGB
    /// fallback (`lastClr`).
    System {
        /// The system color written to `a:sysClr/@val`.
        kind: SystemColorKind,
        /// The last-known sRGB value written to `a:sysClr/@lastClr`.
        last_color: Option<String>,
    },
}

impl ThemeColorValue {
    /// Create an sRGB color from six hexadecimal digits (normalized to
    /// uppercase).
    pub fn srgb(hex: &str) -> Result<Self> {
        Ok(Self::Srgb(require_hex_color(hex)?))
    }

    /// Create a system color with an optional last-known sRGB fallback
    /// (normalized to uppercase).
    pub fn system(kind: SystemColorKind, last_color: Option<&str>) -> Result<Self> {
        Ok(Self::System {
            kind,
            last_color: last_color.map(require_hex_color).transpose()?,
        })
    }
}

/// A complete theme color scheme: a name plus one value for each of the
/// twelve [`ThemeColorSlot`] slots.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ThemeColorScheme {
    name: String,
    colors: Vec<(ThemeColorSlot, ThemeColorValue)>,
}

impl ThemeColorScheme {
    /// Create an empty color scheme with the given name; all twelve slots
    /// must be supplied through [`with_color`](Self::with_color) before the
    /// scheme can be authored.
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            colors: Vec::with_capacity(ThemeColorSlot::ALL.len()),
        }
    }

    /// Set the value of one color slot, replacing any previous value.
    pub fn with_color(mut self, slot: ThemeColorSlot, value: ThemeColorValue) -> Self {
        if let Some(entry) = self
            .colors
            .iter_mut()
            .find(|(existing, _)| *existing == slot)
        {
            entry.1 = value;
        } else {
            self.colors.push((slot, value));
        }
        self
    }

    /// The scheme name written to `a:clrScheme/@name`.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// The value authored for one slot, if supplied.
    pub fn color(&self, slot: ThemeColorSlot) -> Option<&ThemeColorValue> {
        self.colors
            .iter()
            .find(|(existing, _)| *existing == slot)
            .map(|(_, value)| value)
    }
}

/// A per-script font entry (`a:font`) inside a major or minor font face.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ThemeScriptFont {
    /// Script code (`ST_Script`), for example `Jpan` or `Hans`.
    pub script: String,
    /// Typeface used for the script.
    pub typeface: String,
}

/// One font face (`CT_FontCollection`) of a theme font scheme.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ThemeFontFace {
    /// Latin typeface (`a:latin/@typeface`); required by the schema.
    pub latin: String,
    /// East-Asian typeface (`a:ea/@typeface`); empty when unused.
    pub ea: String,
    /// Complex-script typeface (`a:cs/@typeface`); empty when unused.
    pub cs: String,
    /// Per-script font entries, serialized in insertion order.
    pub script_fonts: Vec<ThemeScriptFont>,
}

impl ThemeFontFace {
    /// Create a font face with only a latin typeface; the East-Asian and
    /// complex-script typefaces default to empty, like the Office theme.
    pub fn new(latin: impl Into<String>) -> Self {
        Self {
            latin: latin.into(),
            ea: String::new(),
            cs: String::new(),
            script_fonts: Vec::new(),
        }
    }

    /// Set the East-Asian typeface.
    pub fn with_ea(mut self, typeface: impl Into<String>) -> Self {
        self.ea = typeface.into();
        self
    }

    /// Set the complex-script typeface.
    pub fn with_cs(mut self, typeface: impl Into<String>) -> Self {
        self.cs = typeface.into();
        self
    }

    /// Append a per-script font entry.
    pub fn with_script_font(
        mut self,
        script: impl Into<String>,
        typeface: impl Into<String>,
    ) -> Self {
        self.script_fonts.push(ThemeScriptFont {
            script: script.into(),
            typeface: typeface.into(),
        });
        self
    }
}

/// A complete theme font scheme: a name plus a major (heading) and minor
/// (body) font face.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ThemeFontScheme {
    name: String,
    major: ThemeFontFace,
    minor: ThemeFontFace,
}

impl ThemeFontScheme {
    /// Create a font scheme with the given name and font faces.
    pub fn new(name: impl Into<String>, major: ThemeFontFace, minor: ThemeFontFace) -> Self {
        Self {
            name: name.into(),
            major,
            minor,
        }
    }

    /// The scheme name written to `a:fontScheme/@name`.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// The major (heading) font face.
    pub fn major(&self) -> &ThemeFontFace {
        &self.major
    }

    /// The minor (body) font face.
    pub fn minor(&self) -> &ThemeFontFace {
        &self.minor
    }
}

/// Identity of a theme part created by [`add_theme`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AuthoredTheme {
    /// Part name of the new theme, e.g. `/ppt/theme/theme2.xml`.
    pub part_name: String,
}

// ============================================================================
// Authoring operations
// ============================================================================

/// Create a new theme part with a caller-supplied color and font scheme.
///
/// The theme is written to the next free `/ppt/theme/themeN.xml` part name
/// with the Office theme content type, the twelve-slot color scheme, the
/// major/minor font scheme, and the default format scheme. Serialization is
/// deterministic: the same inputs always produce identical bytes.
pub fn add_theme(
    package: &mut OpcPackage,
    name: &str,
    color_scheme: &ThemeColorScheme,
    font_scheme: &ThemeFontScheme,
) -> Result<AuthoredTheme> {
    require_name("theme", name)?;
    require_color_scheme(color_scheme)?;
    require_font_scheme(font_scheme)?;

    let index = next_part_index(package, "/ppt/theme/theme", ".xml")?;
    let uri = PackURI::new(format!("/ppt/theme/theme{index}.xml"))
        .map_err(|error| OoxmlError::InvalidUri(format!("theme partname: {error}")))?;
    let xml = theme_xml(name, color_scheme, font_scheme)?;
    package.add_part(Box::new(BlobPart::new(
        uri.clone(),
        ct::OFC_THEME.to_string(),
        xml.into_bytes(),
    )));

    invalidate_signatures(package)?;
    validate_theme_graph(package)?;
    Ok(AuthoredTheme {
        part_name: uri.to_string(),
    })
}

/// Attach a theme part to a slide master through a theme relationship.
///
/// The master part must exist and have the slide-master content type, and
/// the theme part must exist and have the Office theme content type. An
/// existing theme relationship on the master is replaced, so afterwards the
/// master keeps exactly one theme relationship, as the read side requires.
/// Returns the new relationship ID.
pub fn attach_theme_to_master(
    package: &mut OpcPackage,
    master_part_name: &str,
    theme_part_name: &str,
) -> Result<String> {
    let master_uri = PackURI::new(master_part_name)
        .map_err(|error| OoxmlError::InvalidUri(format!("slide master partname: {error}")))?;
    let master_part = package.get_part(&master_uri)?;
    if master_part.content_type() != ct::PML_SLIDE_MASTER {
        return Err(OoxmlError::InvalidContentType {
            expected: ct::PML_SLIDE_MASTER.to_string(),
            got: master_part.content_type().to_string(),
        });
    }

    let theme_uri = PackURI::new(theme_part_name)
        .map_err(|error| OoxmlError::InvalidUri(format!("theme partname: {error}")))?;
    let theme_part = package.get_part(&theme_uri)?;
    if theme_part.content_type() != ct::OFC_THEME {
        return Err(OoxmlError::InvalidContentType {
            expected: ct::OFC_THEME.to_string(),
            got: theme_part.content_type().to_string(),
        });
    }

    let master_dir = master_part_name
        .rsplit_once('/')
        .map(|(directory, _)| format!("{directory}/"))
        .ok_or_else(|| {
            OoxmlError::InvalidUri(format!("slide master partname: {master_part_name}"))
        })?;
    let target = relative_target(&master_dir, theme_part_name)?;
    let master_part = package.get_part_mut(&master_uri)?;
    let previous: Vec<String> = master_part
        .rels()
        .iter()
        .filter(|relationship| matches!(relationship.reltype(), rt::THEME | STRICT_THEME_REL))
        .map(|relationship| relationship.r_id().to_string())
        .collect();
    for relationship_id in previous {
        master_part.rels_mut().remove(&relationship_id);
    }
    let relationship_id = master_part.relate_to(&target, rt::THEME);

    invalidate_signatures(package)?;
    validate_theme_graph(package)?;
    Ok(relationship_id)
}

/// Replace the color scheme of an existing theme part.
///
/// The `a:clrScheme` element inside `a:themeElements` is replaced in place
/// with prefix-safe XML; the rest of the part (font scheme, format scheme,
/// extensions) is untouched. The patched part is verified through the
/// read-side theme parser before it is stored.
pub fn store_theme_color_scheme(
    package: &mut OpcPackage,
    theme_part_name: &str,
    color_scheme: &ThemeColorScheme,
) -> Result<()> {
    require_color_scheme(color_scheme)?;
    let uri = theme_part_uri(package, theme_part_name)?;
    let xml = package.get_part(&uri)?.blob().to_vec();
    let span = scan_scheme_span(&xml, "clrScheme")?
        .ok_or_else(|| invalid("theme part has no color scheme inside its theme elements"))?;
    let mut replacement = String::with_capacity(1024);
    push_color_scheme(&mut replacement, color_scheme, true);
    let patched = replace_span(&xml, &span, replacement.as_bytes())?;
    verify_patched_color_scheme(&patched, color_scheme)?;
    package.get_part_mut(&uri)?.set_blob(patched);

    invalidate_signatures(package)?;
    validate_theme_graph(package)?;
    Ok(())
}

/// Replace the font scheme of an existing theme part.
///
/// The `a:fontScheme` element inside `a:themeElements` is replaced in place
/// with prefix-safe XML; the rest of the part is untouched. The patched part
/// is verified through the read-side theme parser before it is stored.
pub fn store_theme_font_scheme(
    package: &mut OpcPackage,
    theme_part_name: &str,
    font_scheme: &ThemeFontScheme,
) -> Result<()> {
    require_font_scheme(font_scheme)?;
    let uri = theme_part_uri(package, theme_part_name)?;
    let xml = package.get_part(&uri)?.blob().to_vec();
    let span = scan_scheme_span(&xml, "fontScheme")?
        .ok_or_else(|| invalid("theme part has no font scheme inside its theme elements"))?;
    let mut replacement = String::with_capacity(1024);
    push_font_scheme(&mut replacement, font_scheme, true);
    let patched = replace_span(&xml, &span, replacement.as_bytes())?;
    verify_patched_font_scheme(&patched, font_scheme)?;
    package.get_part_mut(&uri)?.set_blob(patched);

    invalidate_signatures(package)?;
    validate_theme_graph(package)?;
    Ok(())
}

// ============================================================================
// Theme overrides
// ============================================================================

/// A theme override (`a:themeOverride`, ECMA-376 §14.2.7): the color and/or
/// font scheme a slide layout or slide applies on top of its master's theme.
///
/// Format-scheme overrides are not modeled; unknown content is tolerated on
/// read and never fabricated on write.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ThemeOverride {
    /// The overriding color scheme.
    pub color_scheme: Option<ThemeColorScheme>,
    /// The overriding font scheme.
    pub font_scheme: Option<ThemeFontScheme>,
}

impl ThemeOverride {
    /// An empty override; attach schemes with the `with_*` builders.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the overriding color scheme.
    pub fn with_color_scheme(mut self, scheme: ThemeColorScheme) -> Self {
        self.color_scheme = Some(scheme);
        self
    }

    /// Set the overriding font scheme.
    pub fn with_font_scheme(mut self, scheme: ThemeFontScheme) -> Self {
        self.font_scheme = Some(scheme);
        self
    }
}

impl ThemeColorSlot {
    /// Parse a slot element name (`dk1`, `lt1`, ...).
    pub fn from_token(token: &str) -> Option<Self> {
        Self::ALL
            .into_iter()
            .find(|slot| slot.as_str() == token)
    }
}

/// Serialize a theme override part (`a:themeOverride` root).
fn theme_override_xml(value: &ThemeOverride) -> Result<String> {
    let mut xml = String::with_capacity(1024);
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(r#"<a:themeOverride xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#);
    if let Some(scheme) = &value.color_scheme {
        push_color_scheme(&mut xml, scheme, false);
    }
    if let Some(scheme) = &value.font_scheme {
        push_font_scheme(&mut xml, scheme, false);
    }
    xml.push_str("</a:themeOverride>");
    Ok(xml)
}

/// Parse a theme override part, tolerating unknown content.
fn parse_theme_override(xml: &[u8]) -> Result<ThemeOverride> {
    const MAX_OVERRIDE_XML: usize = 8 * 1024 * 1024;
    const MAX_NODES: usize = 100_000;
    const MAX_DEPTH: usize = 64;
    const MAX_SCRIPT_FONTS: usize = 1_024;
    if xml.len() > MAX_OVERRIDE_XML {
        return Err(limit("theme override XML bytes"));
    }
    let processed = crate::common::mce::process_ooxml(xml)?;
    if processed.len() > MAX_OVERRIDE_XML {
        return Err(limit("processed theme override XML bytes"));
    }
    let mut reader = NsReader::from_reader(processed.as_ref());
    let mut value = ThemeOverride::default();
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut root_seen = false;
    // Which scheme collection is open: 1 = clrScheme, 2 = fontScheme.
    let mut open_scheme: Option<u8> = None;
    let mut current_slot: Option<ThemeColorSlot> = None;
    let mut colors: Vec<(ThemeColorSlot, ThemeColorValue)> = Vec::new();
    let mut color_scheme_name = String::new();
    let mut font_scheme_name = String::new();
    let mut major: Option<ThemeFontFace> = None;
    let mut minor: Option<ThemeFontFace> = None;
    let mut font_target: Option<bool> = None; // false = major, true = minor

    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        let is_empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                nodes += 1;
                if nodes > MAX_NODES {
                    return Err(limit("theme override XML structure"));
                }
                let local_name = element.local_name();
                let local = std::str::from_utf8(local_name.as_ref())
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                let is_dml = matches!(
                    namespace,
                    ResolveResult::Bound(Namespace(uri))
                        if uri == crate::common::xml::DRAWINGML_NAMESPACE
                            || uri == crate::common::xml::STRICT_DRAWINGML_NAMESPACE
                );
                if !root_seen {
                    if !is_dml || local != "themeOverride" {
                        return Err(invalid("invalid theme override root or namespace"));
                    }
                    root_seen = true;
                    if !is_empty {
                        depth = 1;
                    }
                    continue;
                }
                match open_scheme {
                    None if is_dml && local == "clrScheme" && depth == 1 => {
                        open_scheme = Some(1);
                        color_scheme_name = attribute_value(&element, b"name")?.unwrap_or_default();
                    },
                    None if is_dml && local == "fontScheme" && depth == 1 => {
                        open_scheme = Some(2);
                        font_scheme_name = attribute_value(&element, b"name")?.unwrap_or_default();
                    },
                    Some(1) if is_dml => {
                        if let Some(slot) = ThemeColorSlot::from_token(local) {
                            if current_slot.replace(slot).is_some() {
                                return Err(invalid("nested theme color slots"));
                            }
                        } else if let Some(slot) = current_slot {
                            let color = match local {
                                "srgbClr" => {
                                    let value = attribute_value(&element, b"val")?
                                        .ok_or_else(|| invalid("srgbClr lacks a val"))?;
                                    ThemeColorValue::srgb(&value)?
                                },
                                "sysClr" => {
                                    let token = attribute_value(&element, b"val")?
                                        .ok_or_else(|| invalid("sysClr lacks a val"))?;
                                    let kind = SystemColorKind::from_token(&token)
                                        .ok_or_else(|| invalid("unknown system color kind"))?;
                                    ThemeColorValue::system(
                                        kind,
                                        attribute_value(&element, b"lastClr")?.as_deref(),
                                    )?
                                },
                                _ => {
                                    return Err(invalid(
                                        "theme color slot has no srgbClr/sysClr color",
                                    ));
                                },
                            };
                            colors.push((slot, color));
                        }
                    },
                    Some(2) if is_dml => match local {
                        "majorFont" => font_target = Some(false),
                        "minorFont" => font_target = Some(true),
                        "latin" | "ea" | "cs" => {
                            let typeface = attribute_value(&element, b"typeface")?
                                .ok_or_else(|| invalid("theme font lacks a typeface"))?;
                            let face = if font_target == Some(true) {
                                minor.get_or_insert_with(|| ThemeFontFace::new(""))
                            } else {
                                major.get_or_insert_with(|| ThemeFontFace::new(""))
                            };
                            match local {
                                "latin" => face.latin = typeface,
                                "ea" => face.ea = typeface,
                                _ => face.cs = typeface,
                            }
                        },
                        "font" => {
                            let script = attribute_value(&element, b"script")?.unwrap_or_default();
                            let typeface = attribute_value(&element, b"typeface")?
                                .ok_or_else(|| invalid("theme script font lacks a typeface"))?;
                            let face = if font_target == Some(true) {
                                minor.get_or_insert_with(|| ThemeFontFace::new(""))
                            } else {
                                major.get_or_insert_with(|| ThemeFontFace::new(""))
                            };
                            if face.script_fonts.len() >= MAX_SCRIPT_FONTS {
                                return Err(limit("theme script font count"));
                            }
                            face.script_fonts.push(ThemeScriptFont { script, typeface });
                        },
                        _ => {},
                    },
                    _ => {},
                }
                if !is_empty {
                    depth += 1;
                    if depth > MAX_DEPTH {
                        return Err(limit("theme override XML depth"));
                    }
                }
            },
            Event::End(element) => {
                let local = element.local_name();
                if current_slot.is_some_and(|slot| slot.as_str().as_bytes() == local.as_ref()) {
                    current_slot = None;
                } else if open_scheme.is_some()
                    && matches!(local.as_ref(), b"clrScheme" | b"fontScheme")
                {
                    open_scheme = None;
                }
                if depth == 0 {
                    return Err(invalid("unbalanced theme override XML"));
                }
                depth -= 1;
            },
            Event::DocType(_) => return Err(invalid("DTDs are rejected")),
            Event::Eof => break,
            _ => {},
        }
    }
    if !root_seen || depth != 0 {
        return Err(invalid("unterminated theme override document"));
    }

    if !colors.is_empty() {
        let mut scheme = ThemeColorScheme::new(color_scheme_name);
        for (slot, color) in colors {
            scheme = scheme.with_color(slot, color);
        }
        value.color_scheme = Some(scheme);
    }
    if major.is_some() || minor.is_some() {
        let major = major.ok_or_else(|| invalid("theme font scheme lacks a major font"))?;
        let minor = minor.ok_or_else(|| invalid("theme font scheme lacks a minor font"))?;
        value.font_scheme = Some(ThemeFontScheme::new(font_scheme_name, major, minor));
    }
    Ok(value)
}

fn attribute_value(element: &BytesStart<'_>, name: &[u8]) -> Result<Option<String>> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() == name {
            let value = std::str::from_utf8(attribute.value.as_ref())
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            return Ok(Some(
                quick_xml::escape::unescape(value)
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

/// Store a theme override on a slide layout or slide part.
///
/// The override is validated (at least one scheme, complete color scheme)
/// and serialized deterministically; an existing override relationship on
/// the parent is reused (part content replaced, relationship id kept),
/// otherwise a new `/ppt/theme/themeOverrideN.xml` part and relationship
/// are created. The stored part is verified through the override parser.
/// Returns the override part name.
pub fn store_theme_override(
    package: &mut OpcPackage,
    parent_part_name: &str,
    value: &ThemeOverride,
) -> Result<String> {
    if value.color_scheme.is_none() && value.font_scheme.is_none() {
        return Err(invalid("theme override requires at least one scheme"));
    }
    if let Some(scheme) = &value.color_scheme {
        require_color_scheme(scheme)?;
    }
    if let Some(scheme) = &value.font_scheme {
        require_font_scheme(scheme)?;
    }
    let parent_uri = PackURI::new(parent_part_name)
        .map_err(|error| OoxmlError::InvalidUri(format!("theme override parent: {error}")))?;
    let parent = package.get_part(&parent_uri)?;
    if !matches!(parent.content_type(), ct::PML_SLIDE | ct::PML_SLIDE_LAYOUT) {
        return Err(invalid(
            "theme overrides attach to slide layout or slide parts only",
        ));
    }
    let xml = theme_override_xml(value)?;

    // Reuse an existing override part + relationship when present.
    let existing = parent.rels().iter().find(|relationship| {
        matches!(relationship.reltype(), rt::THEME_OVERRIDE | STRICT_THEME_OVERRIDE_REL)
            && !relationship.is_external()
    });
    if let Some(relationship) = existing {
        let uri = relationship.target_partname().map_err(|error| {
            OoxmlError::InvalidRelationship(format!("theme override target: {error}"))
        })?;
        if package.get_part(&uri).is_err() {
            return Err(OoxmlError::InvalidRelationship(
                "theme override relationship targets a missing part".to_string(),
            ));
        }
        package.get_part_mut(&uri)?.set_blob(xml.into_bytes());
        // Verify the stored part parses back to the same model.
        let stored = package.get_part(&uri)?.blob().to_vec();
        if parse_theme_override(&stored)? != *value {
            return Err(invalid("theme override failed read-back validation"));
        }
        invalidate_signatures(package)?;
        return Ok(uri.to_string());
    }

    let index = next_part_index(package, "/ppt/theme/themeOverride", ".xml")?;
    let uri = PackURI::new(format!("/ppt/theme/themeOverride{index}.xml"))
        .map_err(|error| OoxmlError::InvalidUri(format!("theme override partname: {error}")))?;
    package.add_part(Box::new(BlobPart::new(
        uri.clone(),
        ct::OFC_THEME_OVERRIDE.to_string(),
        xml.into_bytes(),
    )));
    let target = uri.relative_ref(parent_uri.base_uri());
    package
        .get_part_mut(&parent_uri)?
        .relate_to(&target, rt::THEME_OVERRIDE);
    let stored = package.get_part(&uri)?.blob().to_vec();
    if parse_theme_override(&stored)? != *value {
        return Err(invalid("theme override failed read-back validation"));
    }
    invalidate_signatures(package)?;
    Ok(uri.to_string())
}

/// Read the theme override attached to a slide layout or slide part.
pub fn theme_override(package: &OpcPackage, parent_part_name: &str) -> Result<Option<ThemeOverride>> {
    let parent_uri = PackURI::new(parent_part_name)
        .map_err(|error| OoxmlError::InvalidUri(format!("theme override parent: {error}")))?;
    let parent = package.get_part(&parent_uri)?;
    let Some(relationship) = parent.rels().iter().find(|relationship| {
        matches!(relationship.reltype(), rt::THEME_OVERRIDE | STRICT_THEME_OVERRIDE_REL)
            && !relationship.is_external()
    }) else {
        return Ok(None);
    };
    let uri = relationship.target_partname().map_err(|error| {
        OoxmlError::InvalidRelationship(format!("theme override target: {error}"))
    })?;
    let part = package.get_part(&uri)?;
    if part.content_type() != ct::OFC_THEME_OVERRIDE {
        return Err(invalid("theme override part has an invalid content type"));
    }
    Ok(Some(parse_theme_override(part.blob())?))
}

/// Remove the theme override from a slide layout or slide part, deleting
/// the override part when it becomes orphaned.
pub fn remove_theme_override(package: &mut OpcPackage, parent_part_name: &str) -> Result<bool> {
    let parent_uri = PackURI::new(parent_part_name)
        .map_err(|error| OoxmlError::InvalidUri(format!("theme override parent: {error}")))?;
    let parent = package.get_part(&parent_uri)?;
    let Some(relationship) = parent.rels().iter().find(|relationship| {
        matches!(relationship.reltype(), rt::THEME_OVERRIDE | STRICT_THEME_OVERRIDE_REL)
            && !relationship.is_external()
    }) else {
        return Ok(false);
    };
    let uri = relationship.target_partname().map_err(|error| {
        OoxmlError::InvalidRelationship(format!("theme override target: {error}"))
    })?;
    let relationship_id = relationship.r_id().to_string();
    package
        .get_part_mut(&parent_uri)?
        .rels_mut()
        .remove(&relationship_id);
    let orphaned = !package.iter_parts().any(|part| {
        part.partname() != &parent_uri
            && part.rels().iter().any(|relationship| {
                matches!(relationship.reltype(), rt::THEME_OVERRIDE | STRICT_THEME_OVERRIDE_REL)
                    && !relationship.is_external()
                    && relationship.target_partname().is_ok_and(|target| target == uri)
            })
    });
    if orphaned {
        package.remove_part(&uri);
    }
    invalidate_signatures(package)?;
    Ok(true)
}


/// Validate the master/layout/theme graph of a package.
///
/// On top of [`crate::pptx::master_layout::validate_master_layout_graph`]
/// this enforces the theme rules the read side applies in
/// `SlideMaster::theme`: every slide master part has exactly one internal
/// theme relationship, targeting an existing Office theme part that parses
/// through the read-side theme parser.
pub fn validate_theme_graph(package: &OpcPackage) -> Result<()> {
    crate::pptx::master_layout::validate_master_layout_graph(package)?;
    for part in package.iter_parts() {
        if part.content_type() != ct::PML_SLIDE_MASTER {
            continue;
        }
        let mut theme_relationships = part
            .rels()
            .iter()
            .filter(|relationship| matches!(relationship.reltype(), rt::THEME | STRICT_THEME_REL));
        let relationship = theme_relationships.next().ok_or_else(|| {
            OoxmlError::InvalidRelationship(format!(
                "slide master '{}' has no theme relationship",
                part.partname()
            ))
        })?;
        if theme_relationships.next().is_some() {
            return Err(OoxmlError::InvalidRelationship(format!(
                "slide master '{}' has multiple theme relationships",
                part.partname()
            )));
        }
        if relationship.is_external() {
            return Err(OoxmlError::InvalidRelationship(format!(
                "theme relationship '{}' of slide master '{}' must be internal",
                relationship.r_id(),
                part.partname()
            )));
        }
        let target = relationship.target_partname().map_err(|error| {
            OoxmlError::InvalidRelationship(format!(
                "invalid theme relationship '{}': {error}",
                relationship.r_id()
            ))
        })?;
        let theme_part = package.get_part(&target)?;
        if theme_part.content_type() != ct::OFC_THEME {
            return Err(OoxmlError::InvalidContentType {
                expected: ct::OFC_THEME.to_string(),
                got: theme_part.content_type().to_string(),
            });
        }
        ThemePart::from_part(theme_part)?.theme()?;
    }
    Ok(())
}

/// Resolve a theme part name to a validated theme part URI.
fn theme_part_uri(package: &OpcPackage, theme_part_name: &str) -> Result<PackURI> {
    let uri = PackURI::new(theme_part_name)
        .map_err(|error| OoxmlError::InvalidUri(format!("theme partname: {error}")))?;
    let part = package.get_part(&uri)?;
    if part.content_type() != ct::OFC_THEME {
        return Err(OoxmlError::InvalidContentType {
            expected: ct::OFC_THEME.to_string(),
            got: part.content_type().to_string(),
        });
    }
    Ok(uri)
}

// ============================================================================
// XML generation
// ============================================================================

/// Serialize a complete theme part.
fn theme_xml(
    name: &str,
    color_scheme: &ThemeColorScheme,
    font_scheme: &ThemeFontScheme,
) -> Result<String> {
    let mut xml = String::with_capacity(8192);
    xml.push_str(XML_DECL);
    let _ = write!(
        xml,
        "<a:theme xmlns:a=\"{A_NS}\" name=\"{}\"><a:themeElements>",
        escape_xml(name)
    );
    push_color_scheme(&mut xml, color_scheme, false);
    push_font_scheme(&mut xml, font_scheme, false);
    xml.push_str(DEFAULT_FORMAT_SCHEME_XML);
    xml.push_str("</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>");
    check_size(xml.as_bytes())?;
    Ok(xml)
}

/// Serialize one color scheme, in schema slot order.
///
/// When `declare_namespace` is set the scheme carries its own `xmlns:a`
/// declaration so it can be patched into a part with unknown prefix
/// bindings.
fn push_color_scheme(xml: &mut String, scheme: &ThemeColorScheme, declare_namespace: bool) {
    xml.push_str("<a:clrScheme");
    if declare_namespace {
        let _ = write!(xml, " xmlns:a=\"{A_NS}\"");
    }
    let _ = write!(xml, " name=\"{}\">", escape_xml(scheme.name()));
    for slot in ThemeColorSlot::ALL {
        let value = scheme
            .color(slot)
            .expect("color schemes are validated before serialization");
        let _ = write!(xml, "<a:{0}>", slot.as_str());
        match value {
            ThemeColorValue::Srgb(hex) => {
                let _ = write!(xml, "<a:srgbClr val=\"{hex}\"/>");
            },
            ThemeColorValue::System { kind, last_color } => {
                let _ = write!(xml, "<a:sysClr val=\"{}\"", kind.as_str());
                if let Some(last_color) = last_color {
                    let _ = write!(xml, " lastClr=\"{last_color}\"");
                }
                xml.push_str("/>");
            },
        }
        let _ = write!(xml, "</a:{0}>", slot.as_str());
    }
    xml.push_str("</a:clrScheme>");
}

/// Serialize one font scheme.
fn push_font_scheme(xml: &mut String, scheme: &ThemeFontScheme, declare_namespace: bool) {
    xml.push_str("<a:fontScheme");
    if declare_namespace {
        let _ = write!(xml, " xmlns:a=\"{A_NS}\"");
    }
    let _ = write!(xml, " name=\"{}\">", escape_xml(scheme.name()));
    push_font_face(xml, "majorFont", scheme.major());
    push_font_face(xml, "minorFont", scheme.minor());
    xml.push_str("</a:fontScheme>");
}

/// Serialize one major or minor font face.
fn push_font_face(xml: &mut String, element: &str, face: &ThemeFontFace) {
    let _ = write!(
        xml,
        "<a:{element}><a:latin typeface=\"{}\"/><a:ea typeface=\"{}\"/><a:cs typeface=\"{}\"/>",
        escape_xml(&face.latin),
        escape_xml(&face.ea),
        escape_xml(&face.cs)
    );
    for script_font in &face.script_fonts {
        let _ = write!(
            xml,
            "<a:font script=\"{}\" typeface=\"{}\"/>",
            escape_xml(&script_font.script),
            escape_xml(&script_font.typeface)
        );
    }
    let _ = write!(xml, "</a:{element}>");
}

// ============================================================================
// Bounded XML scanning and patching
// ============================================================================

/// Byte span of an XML element.
#[derive(Debug, Clone, Copy)]
struct ElementSpan {
    /// Offset of the `<` that opens the element.
    start: usize,
    /// Offset one past the `>` that closes the element.
    end: usize,
}

fn check_size(xml: &[u8]) -> Result<()> {
    if xml.len() > MAX_PART_XML_BYTES {
        return Err(invalid("part XML exceeds 8 MiB"));
    }
    Ok(())
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

/// Find the element with local name `target` that is a direct child of
/// `themeElements` at [`SCHEME_DEPTH`].
fn scan_scheme_span(xml: &[u8], target: &str) -> Result<Option<ElementSpan>> {
    check_size(xml)?;
    let mut reader = Reader::from_reader(xml);
    let mut stack: Vec<(usize, Vec<u8>)> = Vec::new();
    let mut nodes = 0usize;
    loop {
        let before = reader.buffer_position() as usize;
        match reader.read_event() {
            Ok(Event::Start(element)) => {
                nodes += 1;
                if nodes > MAX_SCAN_NODES || stack.len() >= MAX_SCAN_DEPTH {
                    return Err(invalid("part XML resource limit exceeded"));
                }
                let local = local_name(element.name().as_ref()).to_vec();
                stack.push((before, local));
            },
            Ok(Event::Empty(element)) => {
                nodes += 1;
                if nodes > MAX_SCAN_NODES {
                    return Err(invalid("part XML resource limit exceeded"));
                }
                if stack.len() + 1 == SCHEME_DEPTH
                    && stack
                        .last()
                        .is_some_and(|(_, parent)| parent == b"themeElements")
                    && local_name(element.name().as_ref()) == target.as_bytes()
                {
                    return Ok(Some(ElementSpan {
                        start: before,
                        end: reader.buffer_position() as usize,
                    }));
                }
            },
            Ok(Event::End(element)) => {
                let (start, local) = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected closing element in part XML"))?;
                if stack.len() + 1 == SCHEME_DEPTH
                    && stack
                        .last()
                        .is_some_and(|(_, parent)| parent == b"themeElements")
                    && local == target.as_bytes()
                {
                    return Ok(Some(ElementSpan {
                        start,
                        end: reader.buffer_position() as usize,
                    }));
                }
                if local_name(element.name().as_ref()) != local.as_slice() {
                    return Err(invalid("mismatched closing element in part XML"));
                }
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Ok(Event::Eof) => break,
            Err(error) => return Err(OoxmlError::Xml(error.to_string())),
            _ => {},
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated part XML"));
    }
    Ok(None)
}

fn replace_span(xml: &[u8], span: &ElementSpan, replacement: &[u8]) -> Result<Vec<u8>> {
    let mut output = Vec::with_capacity(xml.len() + replacement.len());
    output.extend_from_slice(&xml[..span.start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&xml[span.end..]);
    check_size(&output)?;
    Ok(output)
}

/// Parse a patched theme part through the read-side parser.
fn parse_patched_theme(xml: &[u8]) -> Result<crate::pptx::parts::Theme> {
    let part = BlobPart::new(
        PackURI::new("/ppt/theme/theme1.xml")
            .map_err(|error| OoxmlError::InvalidUri(format!("theme partname: {error}")))?,
        ct::OFC_THEME.to_string(),
        xml.to_vec(),
    );
    ThemePart::from_part(&part)?.theme()
}

/// The patched part must inventory every authored slot back through the
/// same parser the read side applies.
fn verify_patched_color_scheme(xml: &[u8], scheme: &ThemeColorScheme) -> Result<()> {
    let theme = parse_patched_theme(xml)?;
    for slot in ThemeColorSlot::ALL {
        let expected = scheme
            .color(slot)
            .expect("color schemes are validated before serialization");
        let parsed = theme
            .colors
            .iter()
            .find(|color| color.name == slot.as_str())
            .ok_or_else(|| {
                invalid(format!(
                    "patched color scheme lost slot '{}'",
                    slot.as_str()
                ))
            })?;
        match expected {
            ThemeColorValue::Srgb(hex) => {
                if parsed.rgb.as_deref() != Some(hex.as_str()) || parsed.system_color.is_some() {
                    return Err(invalid(format!(
                        "patched color scheme slot '{}' did not round-trip",
                        slot.as_str()
                    )));
                }
            },
            ThemeColorValue::System { kind, last_color } => {
                if parsed.system_color.as_deref() != Some(kind.as_str())
                    || parsed.rgb.as_deref() != last_color.as_deref()
                {
                    return Err(invalid(format!(
                        "patched color scheme slot '{}' did not round-trip",
                        slot.as_str()
                    )));
                }
            },
        }
    }
    Ok(())
}

/// The patched part must inventory the authored major and minor latin
/// typefaces back through the read-side parser.
fn verify_patched_font_scheme(xml: &[u8], scheme: &ThemeFontScheme) -> Result<()> {
    let theme = parse_patched_theme(xml)?;
    if theme.major_font.as_ref().map(|font| font.typeface.as_str())
        != Some(scheme.major().latin.as_str())
        || theme.minor_font.as_ref().map(|font| font.typeface.as_str())
            != Some(scheme.minor().latin.as_str())
    {
        return Err(invalid("patched font scheme did not round-trip"));
    }
    Ok(())
}

// ============================================================================
// Misc helpers and validators
// ============================================================================

/// Find the lowest free numeric suffix for a part-name pattern.
fn next_part_index(package: &OpcPackage, prefix: &str, suffix: &str) -> Result<u32> {
    let mut index = 1u32;
    loop {
        let candidate = PackURI::new(format!("{prefix}{index}{suffix}"))
            .map_err(|error| OoxmlError::InvalidUri(format!("partname allocation: {error}")))?;
        if package.get_part(&candidate).is_err() {
            return Ok(index);
        }
        index = index
            .checked_add(1)
            .ok_or_else(|| invalid("part-name index overflow"))?;
    }
}

/// Compute the relationship target for `target` relative to `source_dir`.
///
/// Both names must be absolute pack URIs; the result uses `..` segments to
/// climb out of the source directory.
fn relative_target(source_dir: &str, target: &str) -> Result<String> {
    let source = source_dir.trim_matches('/');
    let target = target.trim_start_matches('/');
    let source_segments: Vec<&str> = source.split('/').filter(|item| !item.is_empty()).collect();
    let target_segments: Vec<&str> = target.split('/').filter(|item| !item.is_empty()).collect();
    let common = source_segments
        .iter()
        .zip(target_segments.iter())
        .take_while(|(left, right)| left == right)
        .count();
    if common == 0 && !source_segments.is_empty() {
        return Err(OoxmlError::InvalidUri(format!(
            "cannot relativize '{target}' against '/{source}/'"
        )));
    }
    let mut result = String::new();
    for _ in common..source_segments.len() {
        result.push_str("../");
    }
    result.push_str(&target_segments[common..].join("/"));
    Ok(result)
}

fn require_name(label: &str, name: &str) -> Result<()> {
    if name.is_empty() {
        return Err(invalid(format!("{label} name cannot be empty")));
    }
    if name.chars().count() > MAX_NAME_CHARS {
        return Err(invalid(format!("{label} name exceeds 256 characters")));
    }
    Ok(())
}

/// Validate an sRGB hex color and normalize it to uppercase.
fn require_hex_color(value: &str) -> Result<String> {
    if value.len() == 6 && value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Ok(value.to_uppercase());
    }
    Err(invalid(format!(
        "invalid sRGB color '{value}'; expected six hexadecimal digits"
    )))
}

fn require_color_scheme(scheme: &ThemeColorScheme) -> Result<()> {
    require_name("color scheme", scheme.name())?;
    for slot in ThemeColorSlot::ALL {
        let value = scheme.color(slot).ok_or_else(|| {
            invalid(format!(
                "color scheme is missing its '{}' slot",
                slot.as_str()
            ))
        })?;
        match value {
            ThemeColorValue::Srgb(hex) => {
                require_hex_color(hex)?;
            },
            ThemeColorValue::System { last_color, .. } => {
                if let Some(last_color) = last_color {
                    require_hex_color(last_color)?;
                }
            },
        }
    }
    Ok(())
}

fn require_font_scheme(scheme: &ThemeFontScheme) -> Result<()> {
    require_name("font scheme", scheme.name())?;
    require_font_face("major font", scheme.major())?;
    require_font_face("minor font", scheme.minor())?;
    Ok(())
}

fn require_font_face(label: &str, face: &ThemeFontFace) -> Result<()> {
    require_name(label, &face.latin)?;
    if face.ea.chars().count() > MAX_NAME_CHARS || face.cs.chars().count() > MAX_NAME_CHARS {
        return Err(invalid(format!("{label} typeface exceeds 256 characters")));
    }
    if face.script_fonts.len() > MAX_SCRIPT_FONTS_PER_FACE {
        return Err(invalid(format!(
            "too many script fonts in the {label} face"
        )));
    }
    for script_font in &face.script_fonts {
        let script_len = script_font.script.chars().count();
        if script_len == 0
            || script_len > MAX_SCRIPT_CODE_CHARS
            || !script_font
                .script
                .bytes()
                .all(|byte| byte.is_ascii_alphanumeric())
        {
            return Err(invalid(format!(
                "invalid script code '{}' in the {label} face",
                script_font.script
            )));
        }
        require_name("script font", &script_font.typeface)?;
    }
    Ok(())
}

fn invalidate_signatures(package: &mut OpcPackage) -> Result<()> {
    package.clear_digital_signatures().map_err(|error| {
        OoxmlError::Other(format!("cannot invalidate package signatures: {error}"))
    })
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pptx::{ColorMapSlot, Package};
    use litchi_opc::PackageWriter;
    use std::io::Cursor;

    fn roundtrip(package: &Package) -> Package {
        let bytes = PackageWriter::to_bytes(package.opc_package()).unwrap();
        Package::from_reader(Cursor::new(bytes)).unwrap()
    }

    fn corporate_colors() -> ThemeColorScheme {
        ThemeColorScheme::new("Corporate")
            .with_color(
                ThemeColorSlot::Dark1,
                ThemeColorValue::system(SystemColorKind::WindowText, Some("000000")).unwrap(),
            )
            .with_color(
                ThemeColorSlot::Light1,
                ThemeColorValue::system(SystemColorKind::Window, Some("FFFFFF")).unwrap(),
            )
            .with_color(
                ThemeColorSlot::Dark2,
                ThemeColorValue::srgb("1F3864").unwrap(),
            )
            .with_color(
                ThemeColorSlot::Light2,
                ThemeColorValue::srgb("E7E6E6").unwrap(),
            )
            .with_color(
                ThemeColorSlot::Accent1,
                ThemeColorValue::srgb("2E75B6").unwrap(),
            )
            .with_color(
                ThemeColorSlot::Accent2,
                ThemeColorValue::srgb("C00000").unwrap(),
            )
            .with_color(
                ThemeColorSlot::Accent3,
                ThemeColorValue::srgb("70AD47").unwrap(),
            )
            .with_color(
                ThemeColorSlot::Accent4,
                ThemeColorValue::srgb("7030A0").unwrap(),
            )
            .with_color(
                ThemeColorSlot::Accent5,
                ThemeColorValue::srgb("00B0F0").unwrap(),
            )
            .with_color(
                ThemeColorSlot::Accent6,
                ThemeColorValue::srgb("FFC000").unwrap(),
            )
            .with_color(
                ThemeColorSlot::Hyperlink,
                ThemeColorValue::srgb("0563C1").unwrap(),
            )
            .with_color(
                ThemeColorSlot::FollowedHyperlink,
                ThemeColorValue::srgb("954F72").unwrap(),
            )
    }

    fn corporate_fonts() -> ThemeFontScheme {
        ThemeFontScheme::new(
            "Corporate",
            ThemeFontFace::new("Corporate Headline")
                .with_ea("ＭＳ Ｐゴシック")
                .with_script_font("Jpan", "ＭＳ Ｐゴシック"),
            ThemeFontFace::new("Corporate Body"),
        )
    }

    #[test]
    fn authored_theme_attaches_and_resolves_after_roundtrip() {
        let mut package = Package::new().unwrap();
        let master = package.add_slide_master().unwrap();
        let theme = package
            .add_theme("Corporate Theme", &corporate_colors(), &corporate_fonts())
            .unwrap();
        assert_eq!(theme.part_name, "/ppt/theme/theme2.xml");

        let relationship_id = package
            .attach_theme_to_master(&master.part_name, &theme.part_name)
            .unwrap();
        assert!(!relationship_id.is_empty());
        package.validate_theme_graph().unwrap();

        let reopened = roundtrip(&package);
        reopened.validate_theme_graph().unwrap();
        let presentation = reopened.presentation().unwrap();
        let masters = presentation.slide_masters().unwrap();
        assert_eq!(masters.len(), 2);

        // The authored master resolves the authored theme through the read side.
        let authored = masters
            .iter()
            .find(|candidate| candidate.part().part().partname().as_str() == master.part_name)
            .expect("authored master must resolve");
        let resolved = authored.theme().unwrap();
        assert_eq!(resolved.name, "Corporate Theme");
        assert_eq!(
            resolved
                .major_font
                .as_ref()
                .map(|font| font.typeface.as_str()),
            Some("Corporate Headline")
        );
        assert_eq!(
            resolved
                .minor_font
                .as_ref()
                .map(|font| font.typeface.as_str()),
            Some("Corporate Body")
        );
        assert_eq!(resolved.colors.len(), 12);
        let accent2 = resolved
            .colors
            .iter()
            .find(|color| color.name == "accent2")
            .unwrap();
        assert_eq!(accent2.rgb.as_deref(), Some("C00000"));
        assert_eq!(accent2.system_color, None);
        let dark1 = resolved
            .colors
            .iter()
            .find(|color| color.name == "dk1")
            .unwrap();
        assert_eq!(dark1.system_color.as_deref(), Some("windowText"));
        assert_eq!(dark1.rgb.as_deref(), Some("000000"));

        // Color-map resolution on the master maps slots to authored colors.
        let mapped = authored
            .theme_color(ColorMapSlot::Accent2)
            .unwrap()
            .expect("accent2 must resolve through the color map");
        assert_eq!(mapped.rgb.as_deref(), Some("C00000"));
        let text1 = authored
            .theme_color(ColorMapSlot::Text1)
            .unwrap()
            .expect("tx1 maps to dk1");
        assert_eq!(text1.system_color.as_deref(), Some("windowText"));

        // Coexistence: the default master keeps resolving the default theme.
        let default_master = masters
            .iter()
            .find(|candidate| {
                candidate.part().part().partname().as_str() == "/ppt/slideMasters/slideMaster1.xml"
            })
            .unwrap();
        let default_theme = default_master.theme().unwrap();
        assert_eq!(default_theme.name, "Office Theme");
        assert_eq!(
            default_theme
                .colors
                .iter()
                .find(|color| color.name == "accent1")
                .unwrap()
                .rgb
                .as_deref(),
            Some("4F81BD")
        );

        // The presentation theme inventory returns both themes in master order.
        let themes = presentation.get_themes().unwrap();
        assert_eq!(themes.len(), 2);
        assert_eq!(themes[0].name, "Office Theme");
        assert_eq!(themes[1].name, "Corporate Theme");
    }

    #[test]
    fn replaced_schemes_roundtrip_through_read_side() {
        let mut package = Package::new().unwrap();
        package
            .store_theme_color_scheme("/ppt/theme/theme1.xml", &corporate_colors())
            .unwrap();
        package
            .store_theme_font_scheme("/ppt/theme/theme1.xml", &corporate_fonts())
            .unwrap();
        package.validate_theme_graph().unwrap();

        let reopened = roundtrip(&package);
        reopened.validate_theme_graph().unwrap();
        let presentation = reopened.presentation().unwrap();
        let master = &presentation.slide_masters().unwrap()[0];
        let theme = master.theme().unwrap();
        assert_eq!(theme.colors.len(), 12);
        assert_eq!(
            theme
                .colors
                .iter()
                .find(|color| color.name == "accent5")
                .unwrap()
                .rgb
                .as_deref(),
            Some("00B0F0")
        );
        assert_eq!(
            theme
                .colors
                .iter()
                .find(|color| color.name == "lt1")
                .unwrap()
                .system_color
                .as_deref(),
            Some("window")
        );
        assert_eq!(
            theme.major_font.as_ref().map(|font| font.typeface.as_str()),
            Some("Corporate Headline")
        );
        assert_eq!(
            theme.minor_font.as_ref().map(|font| font.typeface.as_str()),
            Some("Corporate Body")
        );

        // Color-map resolution reflects the replaced scheme.
        let accent5 = master
            .theme_color(ColorMapSlot::Accent5)
            .unwrap()
            .expect("accent5 must resolve");
        assert_eq!(accent5.rgb.as_deref(), Some("00B0F0"));
    }

    #[test]
    fn invalid_theme_references_are_rejected() {
        let mut package = Package::new().unwrap();
        let master = package.add_slide_master().unwrap();
        let theme = package
            .add_theme("Corporate Theme", &corporate_colors(), &corporate_fonts())
            .unwrap();

        // Unknown master part.
        assert!(
            package
                .attach_theme_to_master("/ppt/slideMasters/slideMaster99.xml", &theme.part_name)
                .is_err()
        );
        // Master part name pointing at a non-master part.
        assert!(
            package
                .attach_theme_to_master("/ppt/presentation.xml", &theme.part_name)
                .is_err()
        );
        // Unknown theme part.
        assert!(
            package
                .attach_theme_to_master(&master.part_name, "/ppt/theme/theme99.xml")
                .is_err()
        );
        // Theme part name pointing at a non-theme part.
        assert!(
            package
                .attach_theme_to_master(&master.part_name, "/ppt/tableStyles.xml")
                .is_err()
        );
        // Scheme replacement on a non-theme or missing part is rejected.
        assert!(
            package
                .store_theme_color_scheme("/ppt/presentation.xml", &corporate_colors())
                .is_err()
        );
        assert!(
            package
                .store_theme_font_scheme("/ppt/theme/theme99.xml", &corporate_fonts())
                .is_err()
        );

        // Incomplete color schemes are rejected.
        let incomplete = ThemeColorScheme::new("Incomplete").with_color(
            ThemeColorSlot::Dark1,
            ThemeColorValue::srgb("000000").unwrap(),
        );
        assert!(
            package
                .add_theme("Nope", &incomplete, &corporate_fonts())
                .is_err()
        );
        // Invalid hex colors are rejected.
        assert!(ThemeColorValue::srgb("XYZ123").is_err());
        assert!(ThemeColorValue::srgb("12345").is_err());
        // Empty names are rejected.
        assert!(
            package
                .add_theme("", &corporate_colors(), &corporate_fonts())
                .is_err()
        );

        // A valid attach still succeeds after the rejections.
        package
            .attach_theme_to_master(&master.part_name, &theme.part_name)
            .unwrap();
        // Re-attaching retargets the master while keeping exactly one
        // theme relationship, as graph validation confirms.
        package
            .attach_theme_to_master(&master.part_name, "/ppt/theme/theme1.xml")
            .unwrap();
        package.validate_theme_graph().unwrap();
        let presentation = package.presentation().unwrap();
        let masters = presentation.slide_masters().unwrap();
        let retargeted = masters
            .iter()
            .find(|candidate| candidate.part().part().partname().as_str() == master.part_name)
            .unwrap();
        assert_eq!(retargeted.theme().unwrap().name, "Office Theme");
    }

    #[test]
    fn authored_and_patched_themes_serialize_deterministically() {
        let build = || {
            let mut package = Package::new().unwrap();
            let master = package.add_slide_master().unwrap();
            let theme = package
                .add_theme("Corporate Theme", &corporate_colors(), &corporate_fonts())
                .unwrap();
            package
                .attach_theme_to_master(&master.part_name, &theme.part_name)
                .unwrap();
            package
                .store_theme_color_scheme("/ppt/theme/theme1.xml", &corporate_colors())
                .unwrap();
            package
        };
        let first = build();
        let second = build();
        for part_name in [
            "/ppt/theme/theme2.xml",
            "/ppt/theme/theme1.xml",
            "/ppt/slideMasters/slideMaster2.xml",
        ] {
            let uri = PackURI::new(part_name).unwrap();
            assert_eq!(
                first.opc_package().get_part(&uri).unwrap().blob(),
                second.opc_package().get_part(&uri).unwrap().blob(),
                "part {part_name} must serialize deterministically"
            );
        }
        // Relationship targets are deterministic as well.
        let master_uri = PackURI::new("/ppt/slideMasters/slideMaster2.xml").unwrap();
        let first_rels: Vec<String> = first
            .opc_package()
            .get_part(&master_uri)
            .unwrap()
            .rels()
            .iter()
            .map(|relationship| relationship.r_id().to_string())
            .collect();
        let second_rels: Vec<String> = second
            .opc_package()
            .get_part(&master_uri)
            .unwrap()
            .rels()
            .iter()
            .map(|relationship| relationship.r_id().to_string())
            .collect();
        assert_eq!(first_rels, second_rels);
    }

    fn default_layout_name(package: &Package) -> String {
        package
            .presentation()
            .unwrap()
            .slide_masters()
            .unwrap()
            .first()
            .unwrap()
            .slide_layouts()
            .unwrap()
            .first()
            .unwrap()
            .part()
            .part()
            .partname()
            .to_string()
    }

    fn sample_override() -> ThemeOverride {
        ThemeOverride::new()
            .with_color_scheme(corporate_colors())
            .with_font_scheme(corporate_fonts())
    }

    #[test]
    fn theme_override_round_trips_through_store_and_reopen() {
        let mut package = Package::new().unwrap();
        let layout = default_layout_name(&package);
        let value = sample_override();
        let part_name = package
            .store_theme_override(&layout, &value)
            .unwrap();
        assert_eq!(part_name, "/ppt/theme/themeOverride1.xml");

        // Read back before and after a package round-trip.
        assert_eq!(package.theme_override(&layout).unwrap(), Some(value.clone()));
        let reopened = roundtrip(&package);
        assert_eq!(reopened.theme_override(&layout).unwrap(), Some(value));
        let uri = PackURI::new(&part_name).unwrap();
        assert_eq!(
            reopened.opc_package().get_part(&uri).unwrap().content_type(),
            ct::OFC_THEME_OVERRIDE
        );
    }

    #[test]
    fn theme_override_replacement_reuses_part_and_relationship() {
        let mut package = Package::new().unwrap();
        let layout = default_layout_name(&package);
        let first_name = package
            .store_theme_override(&layout, &sample_override())
            .unwrap();
        let replacement = ThemeOverride::new().with_color_scheme(corporate_colors());
        let second_name = package
            .store_theme_override(&layout, &replacement)
            .unwrap();
        assert_eq!(first_name, second_name);
        assert_eq!(
            package.theme_override(&layout).unwrap(),
            Some(replacement)
        );
    }

    #[test]
    fn theme_override_removal_deletes_orphaned_part() {
        let mut package = Package::new().unwrap();
        let layout = default_layout_name(&package);
        let part_name = package
            .store_theme_override(&layout, &sample_override())
            .unwrap();
        assert!(package.remove_theme_override(&layout).unwrap());
        assert_eq!(package.theme_override(&layout).unwrap(), None);
        let uri = PackURI::new(&part_name).unwrap();
        assert!(package.opc_package().get_part(&uri).is_err());
        // Second removal is a no-op.
        assert!(!package.remove_theme_override(&layout).unwrap());
    }

    #[test]
    fn theme_override_rejects_invalid_parents_and_empty_values() {
        let mut package = Package::new().unwrap();
        let layout = default_layout_name(&package);
        // Empty override.
        assert!(
            package
                .store_theme_override(&layout, &ThemeOverride::new())
                .is_err()
        );
        // Master part is not a valid override parent.
        let master = package.add_slide_master().unwrap();
        assert!(
            package
                .store_theme_override(&master.part_name, &sample_override())
                .is_err()
        );
        // Missing parent part.
        assert!(
            package
                .store_theme_override("/ppt/slideLayouts/nope.xml", &sample_override())
                .is_err()
        );
        // Rejections staged nothing.
        assert_eq!(package.theme_override(&layout).unwrap(), None);
    }
}
