//! Automatic text and paragraph styles of `content.xml`, and the character
//! formatting edits that derive new ones (record 0308, G-ODP-1).
//!
//! Character formatting in a presentation lives in automatic styles
//! (`office:automatic-styles`): a `text:span` names a `text`-family style,
//! a `text:p` a `paragraph`-family style, and either style's
//! `style:text-properties` holds the bold, italic, underline, size, colour
//! and font. The scene indexes those styles once while it scans; a
//! [`TextPropertyUpdate`] never changes a style other markup may share, it
//! derives a new automatic style from the node's own and points the node at
//! it.

use std::ops::Range;
use std::sync::Arc;

use litchi_core::{Error, Result};

/// Character properties an automatic style sets locally
/// (`style:text-properties`), as written.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
///
/// Values are shared (`Arc<str>`), so resolving a style chain
/// ([`AutomaticStyles::resolve`]) copies no text (record 0316).
pub struct TextProperties {
    pub(crate) font_weight: Option<Arc<str>>,
    pub(crate) font_style: Option<Arc<str>>,
    pub(crate) underline: Option<Arc<str>>,
    pub(crate) font_size: Option<Arc<str>>,
    pub(crate) color: Option<Arc<str>>,
    pub(crate) font_name: Option<Arc<str>>,
    pub(crate) font_family: Option<Arc<str>>,
}

impl TextProperties {
    /// Whether the weight is bold (`fo:font-weight` `bold` or 600 and up).
    #[must_use]
    pub fn bold(&self) -> Option<bool> {
        self.font_weight.as_deref().map(|weight| match weight {
            "bold" => true,
            "normal" => false,
            numeric => numeric.parse::<u16>().is_ok_and(|weight| weight >= 600),
        })
    }

    /// Whether the posture is italic or oblique (`fo:font-style`).
    #[must_use]
    pub fn italic(&self) -> Option<bool> {
        self.font_style
            .as_deref()
            .map(|style| matches!(style, "italic" | "oblique"))
    }

    /// Whether text is underlined (`style:text-underline-style` other than
    /// `none`).
    #[must_use]
    pub fn underline(&self) -> Option<bool> {
        self.underline.as_deref().map(|style| style != "none")
    }

    /// `fo:font-size` as written: a length (`18pt`) or a percentage.
    #[must_use]
    pub fn size(&self) -> Option<&str> {
        self.font_size.as_deref()
    }

    /// `fo:color` as written (`#rrggbb`).
    #[must_use]
    pub fn color(&self) -> Option<&str> {
        self.color.as_deref()
    }

    /// The font: `fo:font-family` without its quotes, else the
    /// `style:font-name` font-face reference.
    #[must_use]
    pub fn font(&self) -> Option<&str> {
        self.font_family
            .as_deref()
            .map(|family| family.trim_matches(|character| character == '\'' || character == '"'))
            .or(self.font_name.as_deref())
    }

    /// Fill each property this style leaves unset from `parent`, the style
    /// it inherits from. The font is one property: a `style:font-name` or a
    /// `fo:font-family` here hides both of the parent's.
    pub fn inherit(&mut self, parent: &Self) {
        let fill = |field: &mut Option<Arc<str>>, value: &Option<Arc<str>>| {
            if field.is_none() {
                field.clone_from(value);
            }
        };
        fill(&mut self.font_weight, &parent.font_weight);
        fill(&mut self.font_style, &parent.font_style);
        fill(&mut self.underline, &parent.underline);
        fill(&mut self.font_size, &parent.font_size);
        fill(&mut self.color, &parent.color);
        if self.font_name.is_none() && self.font_family.is_none() {
            self.font_name.clone_from(&parent.font_name);
            self.font_family.clone_from(&parent.font_family);
        }
    }

    /// Whether every property `edit` sets already has that value here.
    pub(crate) fn satisfies(&self, edit: &TextPropertyUpdate) -> bool {
        edit.bold.is_none_or(|bold| self.bold() == Some(bold))
            && edit
                .italic
                .is_none_or(|italic| self.italic() == Some(italic))
            && edit
                .underline
                .is_none_or(|underline| self.underline() == Some(underline))
            && edit
                .size
                .as_deref()
                .is_none_or(|size| self.size() == Some(size))
            && edit.color.as_deref().is_none_or(|color| {
                self.color()
                    .is_some_and(|held| held.eq_ignore_ascii_case(color))
            })
            && edit
                .font
                .as_deref()
                .is_none_or(|font| self.font_name.is_none() && self.font() == Some(font))
    }
}

/// One `style:style` of `office:automatic-styles` in `content.xml`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AutomaticStyle {
    pub(crate) range: Range<usize>,
    pub(crate) start_tag: Range<usize>,
    pub(crate) empty: bool,
    pub(crate) name: String,
    pub(crate) family: String,
    pub(crate) parent: Option<String>,
    /// The start tag of the style's `style:text-properties`.
    pub(crate) text_tag: Option<Range<usize>>,
    /// Where a new `style:text-properties` child goes: after the last
    /// other properties child, or after the start tag.
    pub(crate) text_insertion: usize,
    pub(crate) text: TextProperties,
}

impl AutomaticStyle {
    /// The bytes of the whole element.
    #[must_use]
    pub fn range(&self) -> Range<usize> {
        self.range.clone()
    }

    /// `style:name`.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// `style:family` (`text`, `paragraph`, `graphic`, ...).
    #[must_use]
    pub fn family(&self) -> &str {
        &self.family
    }

    /// `style:parent-style-name`.
    #[must_use]
    pub fn parent(&self) -> Option<&str> {
        self.parent.as_deref()
    }

    /// The character properties the style sets locally.
    #[must_use]
    pub fn text_properties(&self) -> &TextProperties {
        &self.text
    }
}

/// The automatic styles of one presentation body, from
/// [`super::Scene::automatic_styles`].
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct AutomaticStyles {
    pub(crate) styles: Vec<AutomaticStyle>,
    /// Positions into `styles`, sorted by family and name.
    pub(crate) order: Vec<usize>,
}

impl AutomaticStyles {
    /// The styles, in document order.
    #[must_use]
    pub fn as_slice(&self) -> &[AutomaticStyle] {
        &self.styles
    }

    /// The style of `family` (`text`, `paragraph`, ...) named `name`.
    #[must_use]
    pub fn get(&self, family: &str, name: &str) -> Option<&AutomaticStyle> {
        let found = self
            .order
            .binary_search_by(|&index| {
                let style = &self.styles[index];
                (style.family.as_str(), style.name.as_str()).cmp(&(family, name))
            })
            .ok()?;
        self.styles.get(self.order[found])
    }

    /// The character properties a paragraph's automatic style sets locally
    /// (`None` when its style is a common style or absent).
    #[must_use]
    pub fn paragraph_text_properties(
        &self,
        paragraph: &super::Paragraph,
    ) -> Option<&TextProperties> {
        Some(
            self.get("paragraph", paragraph.style_name()?)?
                .text_properties(),
        )
    }

    /// The character properties a span's automatic style sets locally.
    #[must_use]
    pub fn run_text_properties(&self, run: &super::Run) -> Option<&TextProperties> {
        if run.kind() != super::RunKind::Span {
            return None;
        }
        Some(self.get("text", run.style_name()?)?.text_properties())
    }

    /// The character properties the style `family`/`name` resolves to
    /// (record 0316): its own, then those of its parents through
    /// `style:parent-style-name` (an automatic style's parent is looked up
    /// here, then in `common`), then the family's default style
    /// (`style:default-style`). A missing parent ends the chain, and a
    /// cycle is cut after 32 styles.
    #[must_use]
    pub fn resolve(
        &self,
        common: &CommonStyles,
        family: &str,
        name: Option<&str>,
    ) -> TextProperties {
        let mut resolved = TextProperties::default();
        let mut next = name;
        let mut automatic = true;
        for _ in 0..32 {
            let Some(current) = next.take() else {
                break;
            };
            // Automatic styles name common styles as parents; a common
            // style's parents are common.
            let own = if automatic {
                self.get(family, current)
            } else {
                None
            };
            automatic = own.is_some();
            let Some(style) = own.or_else(|| common.styles.get(family, current)) else {
                break;
            };
            resolved.inherit(&style.text);
            next = style.parent.as_deref();
        }
        if let Some(default) = common.styles.get(family, "") {
            resolved.inherit(&default.text);
        }
        resolved
    }

    /// The character properties a paragraph's style resolves to
    /// ([`Self::resolve`], `paragraph` family).
    #[must_use]
    pub fn resolve_paragraph(
        &self,
        common: &CommonStyles,
        paragraph: &super::Paragraph,
    ) -> TextProperties {
        self.resolve(common, "paragraph", paragraph.style_name())
    }

    /// The character properties a run's own style resolves to: a span's, a
    /// formatted field's, or a link's text span over its link style; a
    /// plain-text run and a bare field have none.
    #[must_use]
    pub fn resolve_run(&self, common: &CommonStyles, run: &super::Run) -> TextProperties {
        match run.kind() {
            super::RunKind::Span | super::RunKind::Field if run.style_name().is_some() => {
                self.resolve(common, "text", run.style_name())
            },
            super::RunKind::Link => {
                let mut resolved = match run.link_text_style_name() {
                    Some(name) => self.resolve(common, "text", Some(name)),
                    None => TextProperties::default(),
                };
                if run.style_name().is_some() {
                    resolved.inherit(&self.resolve(common, "text", run.style_name()));
                }
                resolved
            },
            _ => TextProperties::default(),
        }
    }

    /// Whether a style of any family has `name`.
    pub(crate) fn has_name(&self, name: &str) -> bool {
        self.styles.iter().any(|style| style.name == name)
    }
}

/// The common styles of a presentation: the `style:style` and
/// `style:default-style` children of `office:styles`, in `styles.xml` or a
/// flat document. Automatic styles name them as parents, and
/// [`AutomaticStyles::resolve`] reads formatting through them (record
/// 0316). A family's default style is indexed under the empty name.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct CommonStyles {
    styles: AutomaticStyles,
}

impl CommonStyles {
    /// Index the common styles of `xml`, a `styles.xml` part or a flat
    /// `office:document`. A document without `office:styles` has none.
    ///
    /// # Errors
    ///
    /// Returns [`litchi_core::Error::InvalidFormat`] for malformed XML.
    pub fn parse(xml: &str) -> Result<Self> {
        use quick_xml::events::Event;
        use quick_xml::name::ResolveResult;

        const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        let invalid = |message: String| Error::InvalidFormat(message);
        // Offsets are into `xml`; the reader starts after a byte-order mark.
        let (base, body) = match xml.strip_prefix('\u{feff}') {
            Some(rest) => ('\u{feff}'.len_utf8(), rest),
            None => (0, xml),
        };
        let mut reader = quick_xml::NsReader::from_str(body);
        let mut prefixes = super::Prefixes::default();
        let mut depth = 0_usize;
        let mut range: Option<Range<usize>> = None;
        loop {
            let start =
                usize::try_from(reader.buffer_position()).map_or(usize::MAX, |at| base + at);
            let (resolved, event) = reader
                .read_resolved_event()
                .map_err(|error| invalid(format!("invalid ODP styles XML: {error}")))?;
            let office = matches!(resolved, ResolveResult::Bound(namespace) if namespace.as_ref() == OFFICE_NS);
            let end = usize::try_from(reader.buffer_position()).map_or(usize::MAX, |at| base + at);
            match event {
                Event::Start(element) => {
                    if depth == 0 {
                        prefixes = super::scan::root_prefixes(&element)?;
                    } else if depth == 1 && office && element.local_name().as_ref() == b"styles" {
                        range = Some(end..end);
                    }
                    depth += 1;
                },
                Event::End(element) => {
                    depth = depth.saturating_sub(1);
                    if depth == 1
                        && office
                        && element.local_name().as_ref() == b"styles"
                        && let Some(open) = &mut range
                    {
                        open.end = start;
                        break;
                    }
                },
                Event::Eof => break,
                _ => {},
            }
        }
        let range = range.filter(|range| range.end >= range.start);
        Ok(Self {
            styles: parse(xml, range, &prefixes)?,
        })
    }
}

/// Most automatic styles indexed.
const MAX_STYLES: usize = 1 << 20;

/// Index the `style:style` children of `office:automatic-styles`, whose
/// content is `range` of `xml`. Element and attribute names are matched by
/// the prefixes the document root binds.
pub(crate) fn parse(
    xml: &str,
    range: Option<Range<usize>>,
    prefixes: &super::Prefixes,
) -> Result<AutomaticStyles> {
    use quick_xml::events::{BytesStart, Event};

    let (Some(range), Some(style), Some(fo)) = (range, &prefixes.style, &prefixes.fo) else {
        return Ok(AutomaticStyles::default());
    };
    let invalid = |message: String| Error::InvalidFormat(message);
    let body = xml
        .get(range.clone())
        .ok_or_else(|| invalid("ODP automatic styles are outside the content".to_string()))?;
    let qualified = |prefix: &str, local: &str| format!("{prefix}:{local}");
    let default_style = qualified(style, "default-style");
    let names = [
        qualified(style, "style"),
        qualified(style, "text-properties"),
        qualified(style, "name"),
        qualified(style, "family"),
        qualified(style, "parent-style-name"),
        qualified(fo, "font-weight"),
        qualified(fo, "font-style"),
        qualified(style, "text-underline-style"),
        qualified(fo, "font-size"),
        qualified(fo, "color"),
        qualified(style, "font-name"),
        qualified(fo, "font-family"),
    ];
    let style_prefix = format!("{style}:");
    let decoder = quick_xml::Reader::from_str("").decoder();
    let attributes = |element: &BytesStart<'_>| -> Result<Vec<(usize, String)>> {
        let mut out = Vec::new();
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| invalid(format!("invalid ODP style attribute: {error}")))?;
            let key = attribute.key.as_ref();
            if let Some(slot) = names.iter().position(|name| name.as_bytes() == key) {
                let value = attribute
                    .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
                    .map_err(|error| invalid(format!("invalid ODP style attribute: {error}")))?;
                out.push((slot, value.into_owned()));
            }
        }
        Ok(out)
    };
    let mut reader = quick_xml::Reader::from_str(body);
    reader.config_mut().check_end_names = true;
    let mut styles: Vec<AutomaticStyle> = Vec::new();
    let mut open: Option<AutomaticStyle> = None;
    let mut depth = 0_usize;
    let mut properties_open = false;
    loop {
        let start = range.start + usize::try_from(reader.buffer_position()).unwrap_or(usize::MAX);
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("invalid ODP automatic styles: {error}")))?;
        let end = range.start + usize::try_from(reader.buffer_position()).unwrap_or(usize::MAX);
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let empty = end > start && xml.as_bytes().get(end - 2) == Some(&b'/');
                let name = element.name();
                let name = name.as_ref();
                if depth == 0 && (name == names[0].as_bytes() || name == default_style.as_bytes()) {
                    let mut started = AutomaticStyle {
                        range: start..end,
                        start_tag: start..end,
                        empty,
                        name: String::new(),
                        family: String::new(),
                        parent: None,
                        text_tag: None,
                        text_insertion: end,
                        text: TextProperties::default(),
                    };
                    for (slot, value) in attributes(&element)? {
                        match slot {
                            2 => started.name = value,
                            3 => started.family = value,
                            4 => started.parent = Some(value),
                            _ => {},
                        }
                    }
                    if empty {
                        styles.push(started);
                    } else {
                        open = Some(started);
                    }
                } else if depth == 1
                    && let Some(current) = &mut open
                {
                    if name == names[1].as_bytes() {
                        if current.text_tag.is_none() {
                            current.text_tag = Some(start..end);
                            let text = &mut current.text;
                            for (slot, value) in attributes(&element)? {
                                let field = match slot {
                                    5 => &mut text.font_weight,
                                    6 => &mut text.font_style,
                                    7 => &mut text.underline,
                                    8 => &mut text.font_size,
                                    9 => &mut text.color,
                                    10 => &mut text.font_name,
                                    11 => &mut text.font_family,
                                    _ => continue,
                                };
                                *field = Some(value.into());
                            }
                        }
                    } else if name.starts_with(style_prefix.as_bytes())
                        && name.ends_with(b"-properties")
                        && current.text_tag.is_none()
                    {
                        // A start tag's end is set at its end tag.
                        if empty {
                            current.text_insertion = end;
                        } else {
                            properties_open = true;
                        }
                    }
                }
                if !empty {
                    depth += 1;
                }
            },
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 1
                    && properties_open
                    && let Some(current) = &mut open
                {
                    // A new text-properties goes after the last other
                    // properties child.
                    current.text_insertion = end;
                    properties_open = false;
                }
                if depth == 0
                    && let Some(mut finished) = open.take()
                {
                    finished.range.end = end;
                    styles.push(finished);
                }
            },
            Event::Eof => break,
            _ => {},
        }
        if styles.len() > MAX_STYLES {
            return Err(invalid(
                "ODP content exceeds the automatic style limit".to_string(),
            ));
        }
    }
    let mut order: Vec<usize> = (0..styles.len()).collect();
    order.sort_by(|&left, &right| {
        let (left, right) = (&styles[left], &styles[right]);
        (left.family.as_str(), left.name.as_str())
            .cmp(&(right.family.as_str(), right.name.as_str()))
    });
    Ok(AutomaticStyles { styles, order })
}

/// Where new automatic styles go in `content.xml`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) enum StylesInsertion {
    /// Before this offset, the `</office:automatic-styles>` end tag.
    Before(usize),
    /// Replace this empty `<office:automatic-styles/>`.
    Expand(Range<usize>),
    /// Create the element at this offset (the `office:body` start tag).
    Create(usize),
    /// The content has no place for automatic styles.
    #[default]
    None,
}

/// Character properties to set on a run or paragraph. Each field left
/// `None` keeps what the node's style had.
#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct TextPropertyUpdate {
    /// Bold (`fo:font-weight`, with its Asian and complex twins).
    pub bold: Option<bool>,
    /// Italic (`fo:font-style`, with its twins).
    pub italic: Option<bool>,
    /// Single underline, or none.
    pub underline: Option<bool>,
    /// The font size, an ODF length such as `18pt`.
    pub size: Option<String>,
    /// The colour, `#rrggbb`.
    pub color: Option<String>,
    /// The font family name.
    pub font: Option<String>,
}

impl TextPropertyUpdate {
    /// An edit that changes nothing.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Whether the edit sets nothing.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        *self == Self::default()
    }

    /// Merge `later` over this edit: its set fields win.
    pub fn merge(&mut self, later: &Self) {
        if later.bold.is_some() {
            self.bold = later.bold;
        }
        if later.italic.is_some() {
            self.italic = later.italic;
        }
        if later.underline.is_some() {
            self.underline = later.underline;
        }
        if later.size.is_some() {
            self.size.clone_from(&later.size);
        }
        if later.color.is_some() {
            self.color.clone_from(&later.color);
        }
        if later.font.is_some() {
            self.font.clone_from(&later.font);
        }
    }

    pub(crate) fn validate(&self) -> Result<()> {
        if let Some(size) = &self.size {
            let number = size.strip_suffix("pt").unwrap_or("");
            let valid = number
                .parse::<f64>()
                .is_ok_and(|value| value.is_finite() && value > 0.0 && value <= 4000.0);
            if !valid {
                return Err(Error::InvalidFormat(format!(
                    "ODP font sizes are positive point lengths such as `18pt`, not `{size}`"
                )));
            }
        }
        if let Some(color) = &self.color {
            let valid = color.len() == 7
                && color.starts_with('#')
                && color[1..].bytes().all(|byte| byte.is_ascii_hexdigit());
            if !valid {
                return Err(Error::InvalidFormat(format!(
                    "ODP colours are `#rrggbb`, not `{color}`"
                )));
            }
        }
        if let Some(font) = &self.font {
            let valid = !font.is_empty()
                && font.len() <= 256
                && !font
                    .chars()
                    .any(|character| character.is_control() || matches!(character, '\'' | '"'));
            if !valid {
                return Err(Error::InvalidFormat(
                    "ODP font names hold 1 to 256 bytes without quotes or control characters"
                        .to_string(),
                ));
            }
        }
        Ok(())
    }

    /// The `(local name with its prefix, value)` attribute changes of a
    /// `style:text-properties` start tag; `None` removes the attribute.
    pub(crate) fn attributes(&self, style: &str, fo: &str) -> Vec<(String, Option<String>)> {
        let mut out = Vec::new();
        let mut set = |prefix: &str, local: &str, value: Option<String>| {
            out.push((format!("{prefix}:{local}"), value));
        };
        if let Some(bold) = self.bold {
            let value = if bold { "bold" } else { "normal" };
            set(fo, "font-weight", Some(value.to_string()));
            set(style, "font-weight-asian", Some(value.to_string()));
            set(style, "font-weight-complex", Some(value.to_string()));
        }
        if let Some(italic) = self.italic {
            let value = if italic { "italic" } else { "normal" };
            set(fo, "font-style", Some(value.to_string()));
            set(style, "font-style-asian", Some(value.to_string()));
            set(style, "font-style-complex", Some(value.to_string()));
        }
        if let Some(underline) = self.underline {
            if underline {
                set(style, "text-underline-style", Some("solid".to_string()));
                set(style, "text-underline-width", Some("auto".to_string()));
                set(
                    style,
                    "text-underline-color",
                    Some("font-color".to_string()),
                );
            } else {
                set(style, "text-underline-style", Some("none".to_string()));
            }
        }
        if let Some(size) = &self.size {
            set(fo, "font-size", Some(size.clone()));
            set(style, "font-size-asian", Some(size.clone()));
            set(style, "font-size-complex", Some(size.clone()));
        }
        if let Some(color) = &self.color {
            set(fo, "color", Some(color.to_ascii_lowercase()));
        }
        if let Some(font) = &self.font {
            let quoted = if font.contains(' ') {
                format!("'{font}'")
            } else {
                font.clone()
            };
            set(fo, "font-family", Some(quoted));
            // A font-face reference would win over the family.
            set(style, "font-name", None);
        }
        out
    }
}
