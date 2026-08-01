//! Bounded, inert DrawingML text-box and WordArt inventory for a DOCX main document.
//!
//! A Word text box is anchored in the document body as a `w:drawing` element
//! holding a wordprocessing shape (`wps:wsp`) whose `wps:txbx` carries a rich
//! word-processing story (`w:txbxContent`) and whose `wps:bodyPr` carries the
//! text-body properties (insets, vertical anchor, direction, wrap, autofit,
//! columns). WordArt is the same shape with an `a:prstTxWarp` text warp preset
//! plus optional Word 2010 text fill/outline/effect styling on the runs.
//!
//! Documents written for compatibility wrap the DrawingML form in
//! `mc:AlternateContent`; markup-compatibility processing then surfaces the
//! legacy VML `w:pict` fallback (`v:textbox` inside a VML shape). This module
//! therefore recognizes both representations, in both the transitional and
//! the ISO Strict namespace dialects.
//!
//! [`load_text_boxes`] parses the main document part into a typed inventory.
//! Everything is treated as inert metadata: linked OLE objects, scripts, and
//! style bodies are never interpreted or followed.

use crate::docx::drawing::ShapeType;
use crate::docx::namespace::{is_wordprocessing_namespace, word_attribute_value};
use crate::error::{OoxmlError, Result};
use litchi_ooxml_common::mce::process_ooxml;
use litchi_ooxml_common::xml::{
    decode_xml_reference, is_drawingml_name, unqualified_attribute_value,
};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

const WPS_NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const WP_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const STRICT_WP_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing";
const VML_NAMESPACE: &[u8] = b"urn:schemas-microsoft-com:vml";
const WORD_2010_NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/word/2010/wordml";

/// ECMA-376 default left/right text inset (0.1 inch) when `lIns`/`rIns` are absent.
const DEFAULT_HORIZONTAL_INSET_EMU: i64 = 91440;
/// ECMA-376 default top/bottom text inset (0.05 inch) when `tIns`/`bIns` are absent.
const DEFAULT_VERTICAL_INSET_EMU: i64 = 45720;

const MAX_DOCUMENT_XML: usize = 32 * 1024 * 1024;
const MAX_TEXT_BOXES: usize = 256;
const MAX_DEPTH: usize = 128;
const MAX_NODES: usize = 200_000;
const MAX_TEXT_BYTES: usize = 1024 * 1024;

/// How a text box is anchored in the document flow.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[repr(u8)]
pub enum TextBoxAnchor {
    /// `wp:inline` — the shape flows with the surrounding text.
    Inline,
    /// `wp:anchor` — the shape floats relative to a page anchor.
    Floating,
    /// Legacy VML `w:pict` markup (markup-compatibility fallback).
    #[default]
    Vml,
}

/// Vertical anchoring of text within the shape (`wps:bodyPr@anchor`).
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[repr(u8)]
pub enum TextVerticalAnchor {
    /// Text starts at the top inset (`t`, the ECMA-376 default).
    #[default]
    Top,
    /// Text is vertically centered (`ctr`).
    Center,
    /// Text ends at the bottom inset (`b`).
    Bottom,
    /// Lines are spread to fill the shape (`just`).
    Justified,
    /// Words are spread to fill the shape (`dist`).
    Distributed,
}

impl TextVerticalAnchor {
    fn from_token(token: &str) -> Option<Self> {
        match token {
            "t" => Some(Self::Top),
            "ctr" => Some(Self::Center),
            "b" => Some(Self::Bottom),
            "just" => Some(Self::Justified),
            "dist" => Some(Self::Distributed),
            _ => None,
        }
    }

    /// The vertical-anchor token for this anchor.
    pub fn as_token(self) -> &'static str {
        match self {
            Self::Top => "t",
            Self::Center => "ctr",
            Self::Bottom => "b",
            Self::Justified => "just",
            Self::Distributed => "dist",
        }
    }
}

/// Direction of text within the shape (`wps:bodyPr@vert`).
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[repr(u8)]
pub enum TextDirection {
    /// Horizontal text (`horz`, the ECMA-376 default).
    #[default]
    Horizontal,
    /// Vertical text, each line rotated 90 degrees (`vert`).
    Vertical,
    /// Vertical text, each line rotated 270 degrees (`vert270`).
    Vertical270,
    /// WordArt vertical text, letters stacked upright (`wordArtVert`).
    WordArtVertical,
    /// East Asian vertical text (`eaVert`).
    EastAsianVertical,
    /// Mongolian vertical text (`mongolianVert`).
    MongolianVertical,
    /// Right-to-left WordArt vertical text (`wordArtVertRtl`).
    WordArtVerticalRtl,
}

impl TextDirection {
    fn from_token(token: &str) -> Option<Self> {
        match token {
            "horz" => Some(Self::Horizontal),
            "vert" => Some(Self::Vertical),
            "vert270" => Some(Self::Vertical270),
            "wordArtVert" => Some(Self::WordArtVertical),
            "eaVert" => Some(Self::EastAsianVertical),
            "mongolianVert" => Some(Self::MongolianVertical),
            "wordArtVertRtl" => Some(Self::WordArtVerticalRtl),
            _ => None,
        }
    }

    /// The ST_TextDirection token for this direction.
    pub fn as_token(self) -> &'static str {
        match self {
            Self::Horizontal => "horz",
            Self::Vertical => "vert",
            Self::Vertical270 => "vert270",
            Self::WordArtVertical => "wordArtVert",
            Self::EastAsianVertical => "eaVert",
            Self::MongolianVertical => "mongolianVert",
            Self::WordArtVerticalRtl => "wordArtVertRtl",
        }
    }
}

/// Whether text wraps inside the shape extents (`wps:bodyPr@wrap`).
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[repr(u8)]
pub enum TextWrap {
    /// Text wraps within the bounding rectangle (`square`, the ECMA-376 default).
    #[default]
    Square,
    /// Text does not wrap; the shape extents are ignored (`none`).
    None,
}

impl TextWrap {
    fn from_token(token: &str) -> Option<Self> {
        match token {
            "square" => Some(Self::Square),
            "none" => Some(Self::None),
            _ => None,
        }
    }

    /// The wrap token for this mode.
    pub fn as_token(self) -> &'static str {
        match self {
            Self::Square => "square",
            Self::None => "none",
        }
    }
}

/// Autofit behavior of the text body (`wps:bodyPr` child element).
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[repr(u8)]
pub enum TextBoxAutofit {
    /// `a:noAutofit` — text is not resized and does not resize the shape (default).
    #[default]
    NoAutofit,
    /// `a:spAutoFit` — the shape grows or shrinks to fit the text.
    ShapeAutofit,
    /// `a:normAutofit` — the text is scaled to fit the shape.
    NormalAutofit,
}

/// Standard DrawingML text warp presets (`a:prstTxWarp@prst`, ST_TextShapeType).
///
/// The `NoShape` preset (`textNoShape`) disables warping and is the value Word
/// writes on ordinary text boxes; any other preset marks WordArt.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[repr(u8)]
pub enum TextWarpPreset {
    /// `textNoShape` — no warp (ordinary text box).
    #[default]
    NoShape,
    /// `textPlain`.
    Plain,
    /// `textStop`.
    Stop,
    /// `textTriangle`.
    Triangle,
    /// `textTriangleInverted`.
    TriangleInverted,
    /// `textChevron`.
    Chevron,
    /// `textChevronInverted`.
    ChevronInverted,
    /// `textRingInside`.
    RingInside,
    /// `textRingOutside`.
    RingOutside,
    /// `textArchUp`.
    ArchUp,
    /// `textArchDown`.
    ArchDown,
    /// `textCircle`.
    Circle,
    /// `textButtonUp`.
    ButtonUp,
    /// `textButtonDown`.
    ButtonDown,
    /// `textArchUpPour`.
    ArchUpPour,
    /// `textArchDownPour`.
    ArchDownPour,
    /// `textCirclePour`.
    CirclePour,
    /// `textButtonPour`.
    ButtonPour,
    /// `textCurveUp`.
    CurveUp,
    /// `textCurveDown`.
    CurveDown,
    /// `textCanUp`.
    CanUp,
    /// `textCanDown`.
    CanDown,
    /// `textWave1`.
    Wave1,
    /// `textWave2`.
    Wave2,
    /// `textDoubleWave1`.
    DoubleWave1,
    /// `textWave4`.
    Wave4,
    /// `textInflate`.
    Inflate,
    /// `textDeflate`.
    Deflate,
    /// `textInflateBottom`.
    InflateBottom,
    /// `textDeflateBottom`.
    DeflateBottom,
    /// `textInflateTop`.
    InflateTop,
    /// `textDeflateTop`.
    DeflateTop,
    /// `textDeflateInflate`.
    DeflateInflate,
    /// `textDeflateInflateDeflate`.
    DeflateInflateDeflate,
    /// `textFadeRight`.
    FadeRight,
    /// `textFadeLeft`.
    FadeLeft,
    /// `textFadeUp`.
    FadeUp,
    /// `textFadeDown`.
    FadeDown,
    /// `textSlantUp`.
    SlantUp,
    /// `textSlantDown`.
    SlantDown,
    /// `textCascadeUp`.
    CascadeUp,
    /// `textCascadeDown`.
    CascadeDown,
}

impl TextWarpPreset {
    /// Parse a preset from its ST_TextShapeType token.
    pub fn from_token(token: &str) -> Option<Self> {
        match token {
            "textNoShape" => Some(Self::NoShape),
            "textPlain" => Some(Self::Plain),
            "textStop" => Some(Self::Stop),
            "textTriangle" => Some(Self::Triangle),
            "textTriangleInverted" => Some(Self::TriangleInverted),
            "textChevron" => Some(Self::Chevron),
            "textChevronInverted" => Some(Self::ChevronInverted),
            "textRingInside" => Some(Self::RingInside),
            "textRingOutside" => Some(Self::RingOutside),
            "textArchUp" => Some(Self::ArchUp),
            "textArchDown" => Some(Self::ArchDown),
            "textCircle" => Some(Self::Circle),
            "textButtonUp" => Some(Self::ButtonUp),
            "textButtonDown" => Some(Self::ButtonDown),
            "textArchUpPour" => Some(Self::ArchUpPour),
            "textArchDownPour" => Some(Self::ArchDownPour),
            "textCirclePour" => Some(Self::CirclePour),
            "textButtonPour" => Some(Self::ButtonPour),
            "textCurveUp" => Some(Self::CurveUp),
            "textCurveDown" => Some(Self::CurveDown),
            "textCanUp" => Some(Self::CanUp),
            "textCanDown" => Some(Self::CanDown),
            "textWave1" => Some(Self::Wave1),
            "textWave2" => Some(Self::Wave2),
            "textDoubleWave1" => Some(Self::DoubleWave1),
            "textWave4" => Some(Self::Wave4),
            "textInflate" => Some(Self::Inflate),
            "textDeflate" => Some(Self::Deflate),
            "textInflateBottom" => Some(Self::InflateBottom),
            "textDeflateBottom" => Some(Self::DeflateBottom),
            "textInflateTop" => Some(Self::InflateTop),
            "textDeflateTop" => Some(Self::DeflateTop),
            "textDeflateInflate" => Some(Self::DeflateInflate),
            "textDeflateInflateDeflate" => Some(Self::DeflateInflateDeflate),
            "textFadeRight" => Some(Self::FadeRight),
            "textFadeLeft" => Some(Self::FadeLeft),
            "textFadeUp" => Some(Self::FadeUp),
            "textFadeDown" => Some(Self::FadeDown),
            "textSlantUp" => Some(Self::SlantUp),
            "textSlantDown" => Some(Self::SlantDown),
            "textCascadeUp" => Some(Self::CascadeUp),
            "textCascadeDown" => Some(Self::CascadeDown),
            _ => None,
        }
    }
}

/// Text insets of the shape body (`wps:bodyPr` `lIns`/`tIns`/`rIns`/`bIns`), in EMUs.
///
/// Missing attributes fall back to the ECMA-376 defaults (0.1 inch horizontal,
/// 0.05 inch vertical).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TextBoxInsets {
    /// Left inset in EMUs.
    pub left_emu: i64,
    /// Top inset in EMUs.
    pub top_emu: i64,
    /// Right inset in EMUs.
    pub right_emu: i64,
    /// Bottom inset in EMUs.
    pub bottom_emu: i64,
}

impl Default for TextBoxInsets {
    fn default() -> Self {
        Self {
            left_emu: DEFAULT_HORIZONTAL_INSET_EMU,
            top_emu: DEFAULT_VERTICAL_INSET_EMU,
            right_emu: DEFAULT_HORIZONTAL_INSET_EMU,
            bottom_emu: DEFAULT_VERTICAL_INSET_EMU,
        }
    }
}

/// Text-body properties of a shape (`wps:bodyPr`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TextBoxBodyProperties {
    /// Text insets.
    pub insets: TextBoxInsets,
    /// Vertical anchoring of the text.
    pub vertical_anchor: TextVerticalAnchor,
    /// Whether the anchor point is horizontally centered (`anchorCtr`).
    pub anchor_center: bool,
    /// Text direction.
    pub direction: TextDirection,
    /// Text wrap behavior.
    pub wrap: TextWrap,
    /// Autofit behavior.
    pub autofit: TextBoxAutofit,
    /// Number of text columns (`numCol`; 1 when absent).
    pub column_count: u32,
    /// Whether paragraph spacing is ignored in the first and last paragraphs
    /// (`spcFirstLastPara`).
    pub space_first_last_paragraph: bool,
}

impl Default for TextBoxBodyProperties {
    fn default() -> Self {
        Self {
            insets: TextBoxInsets::default(),
            vertical_anchor: TextVerticalAnchor::default(),
            anchor_center: false,
            direction: TextDirection::default(),
            wrap: TextWrap::default(),
            autofit: TextBoxAutofit::default(),
            // ECMA-376 defaults `numCol` to a single column.
            column_count: 1,
            space_first_last_paragraph: false,
        }
    }
}

/// Inert WordArt styling discovered on a shape.
///
/// Styling bodies (fill/outline/effect definitions) are deliberately not
/// parsed; only their presence is recorded.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct WordArt {
    /// The text warp preset (`a:prstTxWarp`), when declared.
    pub warp: Option<TextWarpPreset>,
    /// At least one run carries Word 2010 text fill styling (`w14:textFill`).
    pub has_text_fill: bool,
    /// At least one run carries Word 2010 text outline styling (`w14:textOutline`).
    pub has_text_outline: bool,
    /// At least one run carries Word 2010 text effect styling (`w14:textEffects`).
    pub has_text_effects: bool,
}

/// A text run inside a text-box paragraph.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct TextBoxRun {
    /// Run text with tabs and breaks resolved to `\t` and `\n`.
    pub text: String,
    /// Explicit bold toggle, when declared.
    pub bold: Option<bool>,
    /// Explicit italic toggle, when declared.
    pub italic: Option<bool>,
    /// Explicit underline toggle, when declared (`w:u` present and not `none`).
    pub underline: Option<bool>,
}

/// A paragraph inside a text-box story.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct TextBoxParagraph {
    /// The paragraph's runs in document order.
    pub runs: Vec<TextBoxRun>,
}

impl TextBoxParagraph {
    /// Concatenated paragraph text.
    pub fn text(&self) -> String {
        self.runs.iter().map(|run| run.text.as_str()).collect()
    }
}

/// A typed, inert text box or WordArt shape anchored in a Word document.
#[derive(Clone, Debug)]
pub struct DocxTextBox {
    /// Drawing element ID (`wp:docPr@id` / `wps:cNvSpPr@id`), when declared.
    pub id: Option<u32>,
    /// Shape name, when declared (`wp:docPr@name`, or the VML shape `id`).
    pub name: Option<String>,
    /// How the shape is anchored in the document flow.
    pub anchor: TextBoxAnchor,
    /// Preset geometry of the shape, when declared.
    pub shape_type: Option<ShapeType>,
    /// Text-body properties (ECMA-376 defaults when `wps:bodyPr` is absent,
    /// which is always the case for the VML fallback representation).
    pub body: TextBoxBodyProperties,
    /// WordArt warp preset and styling presence flags, when the shape is WordArt.
    pub word_art: Option<WordArt>,
    /// The text-box story as paragraphs with runs.
    pub paragraphs: Vec<TextBoxParagraph>,
}

impl DocxTextBox {
    /// Whether this shape carries WordArt styling.
    pub fn is_word_art(&self) -> bool {
        self.word_art.is_some()
    }

    /// All text of the story, one line per paragraph.
    pub fn text(&self) -> String {
        self.paragraphs
            .iter()
            .map(TextBoxParagraph::text)
            .collect::<Vec<_>>()
            .join("\n")
    }
}

/// Per-shape parse state while streaming the document XML.
#[derive(Default)]
struct ShapeBuilder {
    anchor: TextBoxAnchor,
    id: Option<u32>,
    name: Option<String>,
    shape_type: Option<ShapeType>,
    body: TextBoxBodyProperties,
    saw_body_pr: bool,
    saw_content: bool,
    warp: Option<TextWarpPreset>,
    from_word_art: bool,
    legacy_word_art: bool,
    text_fill: bool,
    text_outline: bool,
    text_effects: bool,
    paragraphs: Vec<TextBoxParagraph>,
    paragraph: Option<TextBoxParagraph>,
    run: Option<TextBoxRun>,
    in_run_properties: bool,
    in_text: bool,
    in_content: bool,
}

impl ShapeBuilder {
    /// Push text into the open run, enforcing the aggregate text cap.
    fn push_text(&mut self, text: &str, total: &mut usize) -> Result<()> {
        *total = total
            .checked_add(text.len())
            .ok_or_else(|| limit("text bytes"))?;
        if *total > MAX_TEXT_BYTES {
            return Err(limit("text bytes"));
        }
        if let Some(run) = self.run.as_mut() {
            run.text.push_str(text);
        }
        Ok(())
    }

    fn finish(self) -> Option<DocxTextBox> {
        if !self.saw_content && !self.saw_body_pr && !self.legacy_word_art {
            return None;
        }
        let warped = self
            .warp
            .is_some_and(|warp| warp != TextWarpPreset::NoShape);
        let styled = self.text_fill || self.text_outline || self.text_effects;
        let word_art = if warped || self.from_word_art || self.legacy_word_art || styled {
            Some(WordArt {
                warp: self.warp,
                has_text_fill: self.text_fill,
                has_text_outline: self.text_outline,
                has_text_effects: self.text_effects,
            })
        } else {
            None
        };
        Some(DocxTextBox {
            id: self.id,
            name: self.name,
            anchor: self.anchor,
            shape_type: self.shape_type,
            body: self.body,
            word_art,
            paragraphs: self.paragraphs,
        })
    }
}

/// Load the typed, inert text-box and WordArt inventory anchored in a main
/// document part.
///
/// `xml_bytes` is the raw `word/document.xml` content; markup-compatibility
/// processing is applied so `mc:AlternateContent` fallbacks resolve to the
/// representation this inventory reads. Shapes are returned in document
/// order, with shapes nested inside another shape's story finishing first.
pub fn load_text_boxes(xml_bytes: &[u8]) -> Result<Vec<DocxTextBox>> {
    if xml_bytes.len() > MAX_DOCUMENT_XML {
        return Err(limit("document XML bytes"));
    }
    let processed = process_ooxml(xml_bytes)?;
    if processed.len() > MAX_DOCUMENT_XML {
        return Err(limit("processed document XML bytes"));
    }
    parse_text_boxes(processed.as_ref())
}

fn parse_text_boxes(xml: &[u8]) -> Result<Vec<DocxTextBox>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut stack: Vec<ShapeBuilder> = Vec::new();
    let mut text_boxes = Vec::new();
    let mut total_text = 0usize;
    let mut depth = 0usize;
    let mut nodes = 0usize;

    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| limit("XML structure"))?;
                nodes = nodes.checked_add(1).ok_or_else(|| limit("XML structure"))?;
                if nodes > MAX_NODES || depth > MAX_DEPTH {
                    return Err(limit("XML structure"));
                }
                handle_element(&namespace, &element, decoder, &resolver, &mut stack, false)?;
            },
            Event::Empty(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| limit("XML structure"))?;
                if nodes > MAX_NODES {
                    return Err(limit("XML structure"));
                }
                handle_element(&namespace, &element, decoder, &resolver, &mut stack, true)?;
            },
            Event::Text(text) => {
                if let Some(builder) = stack.last_mut()
                    && builder.in_text
                {
                    let decoded = text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                    let unescaped = quick_xml::escape::unescape(&decoded)
                        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                    builder.push_text(&unescaped, &mut total_text)?;
                }
            },
            Event::CData(text) => {
                if let Some(builder) = stack.last_mut()
                    && builder.in_text
                {
                    let decoded = text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                    builder.push_text(&decoded, &mut total_text)?;
                }
            },
            Event::GeneralRef(reference) => {
                if let Some(builder) = stack.last_mut()
                    && builder.in_text
                {
                    let decoded = decode_xml_reference(&reference)?;
                    builder.push_text(&decoded, &mut total_text)?;
                }
            },
            Event::End(element) => {
                let local = element.local_name();
                let local = local.as_ref();
                if is_wordprocessing_namespace(&namespace) {
                    match local {
                        b"t" => {
                            if let Some(builder) = stack.last_mut() {
                                builder.in_text = false;
                            }
                        },
                        b"rPr" => {
                            if let Some(builder) = stack.last_mut() {
                                builder.in_run_properties = false;
                            }
                        },
                        b"r" => {
                            if let Some(builder) = stack.last_mut()
                                && let (Some(run), Some(paragraph)) =
                                    (builder.run.take(), builder.paragraph.as_mut())
                            {
                                paragraph.runs.push(run);
                            }
                        },
                        b"p" => {
                            if let Some(builder) = stack.last_mut()
                                && let Some(paragraph) = builder.paragraph.take()
                            {
                                builder.paragraphs.push(paragraph);
                            }
                        },
                        b"txbxContent" => {
                            if let Some(builder) = stack.last_mut() {
                                builder.in_content = false;
                            }
                        },
                        b"drawing" | b"pict" => {
                            if let Some(builder) = stack.pop()
                                && let Some(text_box) = builder.finish()
                            {
                                if text_boxes.len() >= MAX_TEXT_BOXES {
                                    return Err(limit("text box count"));
                                }
                                text_boxes.push(text_box);
                            }
                        },
                        _ => {},
                    }
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("invalid document XML nesting"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof if depth != 0 || !stack.is_empty() => {
                return Err(invalid("unterminated document XML"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(text_boxes)
}

/// Dispatch one element (start or empty) against the builder stack top.
///
/// `empty` distinguishes leaf handling: container elements (`drawing`,
/// `pict`, `txbxContent`, `p`, `r`) only open on the start form.
fn handle_element(
    namespace: &ResolveResult<'_>,
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &quick_xml::name::NamespaceResolver,
    stack: &mut Vec<ShapeBuilder>,
    empty: bool,
) -> Result<()> {
    let name = element.name();
    let local = name.local_name();
    let local = local.as_ref();

    if is_wordprocessing_namespace(namespace) {
        match local {
            b"drawing" if !empty => stack.push(ShapeBuilder {
                anchor: TextBoxAnchor::Inline,
                ..ShapeBuilder::default()
            }),
            b"pict" if !empty => stack.push(ShapeBuilder {
                anchor: TextBoxAnchor::Vml,
                ..ShapeBuilder::default()
            }),
            b"txbxContent" if !empty => {
                if let Some(builder) = stack.last_mut() {
                    builder.in_content = true;
                    builder.saw_content = true;
                }
            },
            b"p" if !empty => {
                if let Some(builder) = stack.last_mut()
                    && builder.in_content
                    && builder.paragraph.is_none()
                {
                    builder.paragraph = Some(TextBoxParagraph::default());
                }
            },
            b"r" if !empty => {
                if let Some(builder) = stack.last_mut()
                    && builder.paragraph.is_some()
                {
                    builder.run = Some(TextBoxRun::default());
                }
            },
            b"rPr" if !empty => {
                if let Some(builder) = stack.last_mut()
                    && builder.run.is_some()
                {
                    builder.in_run_properties = true;
                }
            },
            b"t" if !empty => {
                if let Some(builder) = stack.last_mut()
                    && builder.run.is_some()
                {
                    builder.in_text = true;
                }
            },
            b"tab" | b"br" | b"cr" => {
                if let Some(builder) = stack.last_mut()
                    && let Some(run) = builder.run.as_mut()
                {
                    let character = if local == b"tab" { '\t' } else { '\n' };
                    run.text.push(character);
                }
            },
            b"b" | b"i" => {
                if let Some(builder) = stack.last_mut()
                    && builder.in_run_properties
                    && let Some(run) = builder.run.as_mut()
                {
                    let value = word_attribute_value(element, b"val", decoder, resolver)?
                        .is_none_or(|value| is_on(&value));
                    if local == b"b" {
                        run.bold = Some(value);
                    } else {
                        run.italic = Some(value);
                    }
                }
            },
            b"u" => {
                if let Some(builder) = stack.last_mut()
                    && builder.in_run_properties
                    && let Some(run) = builder.run.as_mut()
                {
                    let value = word_attribute_value(element, b"val", decoder, resolver)?;
                    run.underline = Some(value.as_deref() != Some("none"));
                }
            },
            _ => {},
        }
    } else if is_namespace(namespace, WP_NAMESPACE) || is_namespace(namespace, STRICT_WP_NAMESPACE)
    {
        if let Some(builder) = stack.last_mut() {
            match local {
                b"inline" => builder.anchor = TextBoxAnchor::Inline,
                b"anchor" => builder.anchor = TextBoxAnchor::Floating,
                b"docPr" => {
                    builder.id = attribute(element, b"id", decoder)?.and_then(|v| v.parse().ok());
                    builder.name = attribute(element, b"name", decoder)?;
                },
                _ => {},
            }
        }
    } else if is_namespace(namespace, WPS_NAMESPACE) {
        if let Some(builder) = stack.last_mut() {
            match local {
                b"cNvSpPr" => {
                    if builder.id.is_none() {
                        builder.id =
                            attribute(element, b"id", decoder)?.and_then(|v| v.parse().ok());
                    }
                    if builder.name.is_none() {
                        builder.name = attribute(element, b"name", decoder)?;
                    }
                },
                b"bodyPr" => parse_body_pr(builder, element, decoder)?,
                _ => {},
            }
        }
    } else if is_drawingml_name(namespace, name, local) {
        if let Some(builder) = stack.last_mut() {
            match local {
                b"prstGeom" => {
                    if let Some(preset) = attribute(element, b"prst", decoder)? {
                        builder.shape_type = Some(ShapeType::from_preset(&preset));
                    }
                },
                b"noAutofit" => builder.body.autofit = TextBoxAutofit::NoAutofit,
                b"spAutoFit" => builder.body.autofit = TextBoxAutofit::ShapeAutofit,
                b"normAutofit" => builder.body.autofit = TextBoxAutofit::NormalAutofit,
                b"prstTxWarp" => {
                    builder.warp = attribute(element, b"prst", decoder)?
                        .and_then(|preset| TextWarpPreset::from_token(&preset));
                },
                _ => {},
            }
        }
    } else if is_namespace(namespace, VML_NAMESPACE) {
        if let Some(builder) = stack.last_mut() {
            match local {
                b"textpath" => {
                    if attribute(element, b"on", decoder)?.is_some_and(|v| is_on(&v)) {
                        builder.legacy_word_art = true;
                    }
                },
                b"shapetype" => {},
                _ => {
                    // The first VML shape-like element supplies the identity
                    // of a `w:pict` fallback (typically `v:shape` or `v:rect`).
                    if builder.name.is_none()
                        && builder.anchor == TextBoxAnchor::Vml
                        && let Some(id) = attribute(element, b"id", decoder)?
                    {
                        builder.name = Some(id);
                    }
                },
            }
        }
    } else if is_namespace(namespace, WORD_2010_NAMESPACE)
        && let Some(builder) = stack.last_mut()
        && builder.in_run_properties
    {
        match local {
            b"textFill" => builder.text_fill = true,
            b"textOutline" => builder.text_outline = true,
            b"textEffects" => builder.text_effects = true,
            _ => {},
        }
    }
    Ok(())
}

/// Parse the `wps:bodyPr` attributes into the builder's body properties.
fn parse_body_pr(
    builder: &mut ShapeBuilder,
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<()> {
    builder.saw_body_pr = true;
    let body = &mut builder.body;
    if let Some(value) = attribute(element, b"lIns", decoder)?
        && let Ok(inset) = value.parse()
    {
        body.insets.left_emu = inset;
    }
    if let Some(value) = attribute(element, b"tIns", decoder)?
        && let Ok(inset) = value.parse()
    {
        body.insets.top_emu = inset;
    }
    if let Some(value) = attribute(element, b"rIns", decoder)?
        && let Ok(inset) = value.parse()
    {
        body.insets.right_emu = inset;
    }
    if let Some(value) = attribute(element, b"bIns", decoder)?
        && let Ok(inset) = value.parse()
    {
        body.insets.bottom_emu = inset;
    }
    if let Some(value) = attribute(element, b"anchor", decoder)?
        && let Some(anchor) = TextVerticalAnchor::from_token(&value)
    {
        body.vertical_anchor = anchor;
    }
    if let Some(value) = attribute(element, b"anchorCtr", decoder)? {
        body.anchor_center = is_on(&value);
    }
    if let Some(value) = attribute(element, b"vert", decoder)?
        && let Some(direction) = TextDirection::from_token(&value)
    {
        body.direction = direction;
    }
    if let Some(value) = attribute(element, b"wrap", decoder)?
        && let Some(wrap) = TextWrap::from_token(&value)
    {
        body.wrap = wrap;
    }
    if let Some(value) = attribute(element, b"numCol", decoder)?
        && let Ok(count) = value.parse()
    {
        body.column_count = count;
    }
    if let Some(value) = attribute(element, b"spcFirstLastPara", decoder)? {
        body.space_first_last_paragraph = is_on(&value);
    }
    if let Some(value) = attribute(element, b"fromWordArt", decoder)? {
        builder.from_word_art = is_on(&value);
    }
    Ok(())
}

/// OOXML boolean attribute values: `1`, `true`, and `on` are truthy.
fn is_on(value: &str) -> bool {
    matches!(value, "1" | "true" | "on")
}

fn is_namespace(namespace: &ResolveResult<'_>, uri: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == uri)
}

fn attribute(
    element: &quick_xml::events::BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<Option<String>> {
    Ok(unqualified_attribute_value(element, name, decoder)?)
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

fn limit(label: &str) -> OoxmlError {
    invalid(format!("DOCX text box {label} limit exceeded"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::OpcPackage;
    use litchi_opc::PackURI;

    const TEXT_BOX_FIXTURE: &[u8] = include_bytes!(
        "../../../../test-data/libreoffice-core/sw/qa/core/objectpositioning/data/do-not-capture-draw-objs-on-page-draw-wrap-none.docx"
    );

    /// Wrap a body fragment in a `w:document` root for the given dialect.
    fn document(strict: bool, body: &str) -> String {
        let (w, a, wp) = if strict {
            (
                "http://purl.oclc.org/ooxml/wordprocessingml/main",
                "http://purl.oclc.org/ooxml/drawingml/main",
                "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing",
            )
        } else {
            (
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                "http://schemas.openxmlformats.org/drawingml/2006/main",
                "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
            )
        };
        format!(
            "<w:document xmlns:w=\"{w}\" xmlns:a=\"{a}\" xmlns:wp=\"{wp}\" \
             xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" \
             xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" \
             xmlns:v=\"urn:schemas-microsoft-com:vml\"><w:body>{body}</w:body></w:document>"
        )
    }

    /// A floating DrawingML text box with explicit body properties.
    fn floating_text_box() -> String {
        "<w:p><w:r><w:drawing><wp:anchor>\
         <wp:extent cx=\"1828800\" cy=\"914400\"/><wp:docPr id=\"7\" name=\"Box 7\"/>\
         <a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">\
         <wps:wsp><wps:cNvSpPr id=\"7\" name=\"Box 7\"/>\
         <wps:spPr><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></wps:spPr>\
         <wps:txbx><w:txbxContent><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:txbxContent></wps:txbx>\
         <wps:bodyPr lIns=\"182880\" tIns=\"91440\" rIns=\"182880\" bIns=\"91440\" anchor=\"ctr\" \
         anchorCtr=\"1\" vert=\"vert270\" wrap=\"none\" numCol=\"2\" spcFirstLastPara=\"1\">\
         <a:spAutoFit/></wps:bodyPr>\
         </wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>"
            .to_string()
    }

    #[test]
    fn parses_floating_text_box_body_properties() {
        let xml = document(false, &floating_text_box());
        let text_boxes = load_text_boxes(xml.as_bytes()).unwrap();
        assert_eq!(text_boxes.len(), 1);
        let text_box = &text_boxes[0];
        assert_eq!(text_box.id, Some(7));
        assert_eq!(text_box.name.as_deref(), Some("Box 7"));
        assert_eq!(text_box.anchor, TextBoxAnchor::Floating);
        assert_eq!(text_box.shape_type, Some(ShapeType::Rectangle));
        assert!(!text_box.is_word_art());
        assert_eq!(text_box.text(), "Hello");
        let body = &text_box.body;
        assert_eq!(
            body.insets,
            TextBoxInsets {
                left_emu: 182880,
                top_emu: 91440,
                right_emu: 182880,
                bottom_emu: 91440,
            }
        );
        assert_eq!(body.vertical_anchor, TextVerticalAnchor::Center);
        assert!(body.anchor_center);
        assert_eq!(body.direction, TextDirection::Vertical270);
        assert_eq!(body.wrap, TextWrap::None);
        assert_eq!(body.autofit, TextBoxAutofit::ShapeAutofit);
        assert_eq!(body.column_count, 2);
        assert!(body.space_first_last_paragraph);
    }

    #[test]
    fn parses_multi_paragraph_story_with_run_formatting() {
        let body = "<w:p><w:r><w:drawing><wp:inline><wp:docPr id=\"1\" name=\"InlineBox\"/>\
             <a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">\
             <wps:wsp><wps:txbx><w:txbxContent>\
             <w:p><w:r><w:t>First</w:t></w:r></w:p>\
             <w:p>\
             <w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r>\
             <w:r><w:rPr><w:i w:val=\"0\"/><w:u w:val=\"single\"/></w:rPr><w:t>plain</w:t></w:r>\
             <w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>c</w:t></w:r>\
             </w:p>\
             </w:txbxContent></wps:txbx><wps:bodyPr><a:noAutofit/></wps:bodyPr></wps:wsp>\
             </a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";
        let xml = document(false, body);
        let text_boxes = load_text_boxes(xml.as_bytes()).unwrap();
        assert_eq!(text_boxes.len(), 1);
        let text_box = &text_boxes[0];
        assert_eq!(text_box.anchor, TextBoxAnchor::Inline);
        assert_eq!(text_box.paragraphs.len(), 2);
        assert_eq!(text_box.text(), "First\nBoldplaina\tb\nc");
        let runs = &text_box.paragraphs[1].runs;
        assert_eq!(runs[0].bold, Some(true));
        assert_eq!(runs[1].italic, Some(false));
        assert_eq!(runs[1].underline, Some(true));
        assert_eq!(text_box.body.autofit, TextBoxAutofit::NoAutofit);
        // Undeclared body properties fall back to the ECMA-376 defaults.
        assert_eq!(text_box.body.column_count, 1);
        assert_eq!(text_box.body.insets.left_emu, DEFAULT_HORIZONTAL_INSET_EMU);
        assert_eq!(text_box.body.insets.top_emu, DEFAULT_VERTICAL_INSET_EMU);
    }

    #[test]
    fn parses_word_art_warp_and_styling_flags() {
        let body = "<w:p><w:r><w:drawing><wp:anchor><wp:docPr id=\"9\" name=\"WordArt 1\"/>\
             <a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">\
             <wps:wsp>\
             <wps:spPr><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></wps:spPr>\
             <wps:txbx><w:txbxContent><w:p>\
             <w:r><w:rPr><w14:textFill><w14:solidFill><w14:srgbClr w14:val=\"FF0000\"/></w14:solidFill></w14:textFill></w:rPr>\
             <w:t>Warped</w:t></w:r>\
             </w:p></w:txbxContent></wps:txbx>\
             <wps:bodyPr fromWordArt=\"1\"><a:prstTxWarp prst=\"textArchUp\"><a:avLst/></a:prstTxWarp><a:noAutofit/></wps:bodyPr>\
             </wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>";
        let xml = document(false, body);
        let text_boxes = load_text_boxes(xml.as_bytes()).unwrap();
        assert_eq!(text_boxes.len(), 1);
        let text_box = &text_boxes[0];
        assert!(text_box.is_word_art());
        let word_art = text_box.word_art.unwrap();
        assert_eq!(word_art.warp, Some(TextWarpPreset::ArchUp));
        assert!(word_art.has_text_fill);
        assert!(!word_art.has_text_outline);
        assert!(!word_art.has_text_effects);
        assert_eq!(text_box.text(), "Warped");
    }

    #[test]
    fn text_no_shape_warp_is_not_word_art() {
        let body = "<w:p><w:r><w:drawing><wp:anchor><wp:docPr id=\"2\" name=\"Plain\"/>\
             <a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">\
             <wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:txbxContent></wps:txbx>\
             <wps:bodyPr><a:prstTxWarp prst=\"textNoShape\"><a:avLst/></a:prstTxWarp></wps:bodyPr>\
             </wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>";
        let xml = document(false, body);
        let text_boxes = load_text_boxes(xml.as_bytes()).unwrap();
        assert_eq!(text_boxes.len(), 1);
        assert!(!text_boxes[0].is_word_art());
    }

    #[test]
    fn parses_vml_text_box_fallback() {
        let body = "<w:p><w:r><w:pict>\
             <v:shape id=\"Text Box 3\" style=\"position:absolute;width:100pt;height:50pt\">\
             <v:textbox><w:txbxContent>\
             <w:p><w:r><w:t>Legacy</w:t></w:r></w:p>\
             </w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>";
        let xml = document(false, body);
        let text_boxes = load_text_boxes(xml.as_bytes()).unwrap();
        assert_eq!(text_boxes.len(), 1);
        let text_box = &text_boxes[0];
        assert_eq!(text_box.anchor, TextBoxAnchor::Vml);
        assert_eq!(text_box.name.as_deref(), Some("Text Box 3"));
        assert_eq!(text_box.text(), "Legacy");
        assert!(!text_box.is_word_art());
        // VML has no bodyPr; the typed defaults apply.
        assert_eq!(text_box.body, TextBoxBodyProperties::default());
    }

    #[test]
    fn parses_vml_textpath_word_art() {
        let body = "<w:p><w:r><w:pict>\
             <v:shape id=\"WordArt 4\" style=\"width:100pt;height:50pt\">\
             <v:textpath on=\"1\" string=\"Curved\"/>\
             <v:textbox><w:txbxContent><w:p><w:r><w:t>Curved</w:t></w:r></w:p></w:txbxContent></v:textbox>\
             </v:shape></w:pict></w:r></w:p>";
        let xml = document(false, body);
        let text_boxes = load_text_boxes(xml.as_bytes()).unwrap();
        assert_eq!(text_boxes.len(), 1);
        let text_box = &text_boxes[0];
        assert_eq!(text_box.anchor, TextBoxAnchor::Vml);
        assert!(text_box.is_word_art());
        assert_eq!(text_box.word_art.unwrap().warp, None);
        assert_eq!(text_box.text(), "Curved");
    }

    #[test]
    fn parses_strict_namespace_text_box() {
        let xml = document(true, &floating_text_box());
        let text_boxes = load_text_boxes(xml.as_bytes()).unwrap();
        assert_eq!(text_boxes.len(), 1);
        let text_box = &text_boxes[0];
        assert_eq!(text_box.id, Some(7));
        assert_eq!(text_box.anchor, TextBoxAnchor::Floating);
        assert_eq!(text_box.body.autofit, TextBoxAutofit::ShapeAutofit);
        assert_eq!(text_box.text(), "Hello");
    }

    #[test]
    fn rejects_malformed_and_oversized_input() {
        // Unterminated document XML.
        let xml = document(false, &floating_text_box());
        let truncated = &xml.as_bytes()[..xml.len() / 2];
        assert!(load_text_boxes(truncated).is_err());

        // Oversized document XML is rejected before parsing.
        let oversized = vec![b' '; MAX_DOCUMENT_XML + 1];
        assert!(load_text_boxes(&oversized).is_err());

        // DTDs are rejected.
        let dtd = "<!DOCTYPE w:document [<!ENTITY x \"y\">]><w:document \
                   xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body/></w:document>";
        assert!(load_text_boxes(dtd.as_bytes()).is_err());
    }

    #[test]
    fn ignores_drawings_without_text_stories() {
        let body = "<w:p><w:r><w:drawing><wp:inline><wp:docPr id=\"5\" name=\"Picture\"/>\
             <a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">\
             </a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";
        let xml = document(false, body);
        assert!(load_text_boxes(xml.as_bytes()).unwrap().is_empty());
    }

    #[test]
    fn reads_libreoffice_text_box_fixture() {
        let package = OpcPackage::from_bytes(TEXT_BOX_FIXTURE).unwrap();
        let document = package
            .get_part(&PackURI::new("/word/document.xml").unwrap())
            .unwrap();
        let text_boxes = load_text_boxes(document.blob()).unwrap();
        // The `mc:Choice Requires=\"wps\"` form is collapsed to its VML
        // fallback, so the shape surfaces as a VML-anchored text box.
        let text_box = text_boxes
            .iter()
            .find(|text_box| text_box.name.as_deref() == Some("Rectangle 6"))
            .expect("fixture carries a VML fallback text box");
        assert_eq!(text_box.anchor, TextBoxAnchor::Vml);
        assert_eq!(text_box.text(), "<Text>");
    }
}
