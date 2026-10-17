//! A source-located catalog of the styles part (record 0319, L-DOCX-R6).
//!
//! [`Catalog`](crate::styles::catalog::Catalog) reads a `styles.xml` payload once and keeps, for every
//! `w:style`, its identity, its UI flags and the byte range of its element.
//! The style's paragraph and run properties are read on request through the
//! same typed readers that read a paragraph's or run's direct properties
//! ([`crate::properties`]), so a style answers the same keys a paragraph or
//! run does.
//!
//! The catalog also computes the edited payloads that the main-document
//! editor publishes ([`crate::document::Edit::add_style`] and its siblings):
//! each edit splices the raw part, so every other byte stays as authored.

use std::sync::Arc;

use crate::error::{Error, Result};
use crate::part_scan::{
    self, Child, Tree, escape, on_off, on_off_element, qualified, respliced_paragraph,
    respliced_run, splice_error, val_element, wrap,
};
use crate::properties;

use super::Type;

/// The WordprocessingML namespace of ISO/IEC 29500 Strict documents.
const STRICT_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";

/// The `CT_Style` child order, which every insertion keeps.
const CHILD_ORDER: &[&str] = &[
    "name",
    "aliases",
    "basedOn",
    "next",
    "link",
    "autoRedefine",
    "hidden",
    "uiPriority",
    "semiHidden",
    "unhideWhenUsed",
    "qFormat",
    "locked",
    "personal",
    "personalCompose",
    "personalReply",
    "rsid",
    "pPr",
    "rPr",
    "tblPr",
    "trPr",
    "tcPr",
    "tblStylePr",
];

/// A style's on/off UI flags, each a `CT_OnOff` child of `w:style`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Flag {
    /// `w:autoRedefine`: the style follows direct formatting of its paragraphs.
    AutoRedefine,
    /// `w:hidden`: the style is hidden from the user interface.
    Hidden,
    /// `w:semiHidden`: the style is hidden from the main gallery.
    SemiHidden,
    /// `w:unhideWhenUsed`: a semi-hidden style is shown once it is used.
    UnhideWhenUsed,
    /// `w:qFormat`: the style is a primary (quick) style.
    QuickFormat,
    /// `w:locked`: the style cannot be applied under formatting protection.
    Locked,
}

impl Flag {
    /// Every flag, in `CT_Style` order.
    pub const ALL: [Self; 6] = [
        Self::AutoRedefine,
        Self::Hidden,
        Self::SemiHidden,
        Self::UnhideWhenUsed,
        Self::QuickFormat,
        Self::Locked,
    ];

    /// The flag's element name.
    #[must_use]
    pub const fn local_name(self) -> &'static str {
        match self {
            Self::AutoRedefine => "autoRedefine",
            Self::Hidden => "hidden",
            Self::SemiHidden => "semiHidden",
            Self::UnhideWhenUsed => "unhideWhenUsed",
            Self::QuickFormat => "qFormat",
            Self::Locked => "locked",
        }
    }

    const fn slot(self) -> usize {
        match self {
            Self::AutoRedefine => 0,
            Self::Hidden => 1,
            Self::SemiHidden => 2,
            Self::UnhideWhenUsed => 3,
            Self::QuickFormat => 4,
            Self::Locked => 5,
        }
    }
}

/// One `w:style` of a [`Catalog`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Entry {
    element: usize,
    id: String,
    kind: Option<Type>,
    default: bool,
    custom: Option<bool>,
    name: Option<String>,
    based_on: Option<String>,
    next: Option<String>,
    link: Option<String>,
    flags: [Option<bool>; 6],
    ui_priority: Option<i64>,
}

impl Entry {
    /// The style identifier (`w:styleId`).
    #[must_use]
    pub fn id(&self) -> &str {
        &self.id
    }

    /// The display name (`w:name`).
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// The style type (`w:type`), `None` when the attribute is absent or
    /// not a style type.
    #[must_use]
    pub const fn kind(&self) -> Option<Type> {
        self.kind
    }

    /// Whether the style is its type's default (`w:default`).
    #[must_use]
    pub const fn is_default(&self) -> bool {
        self.default
    }

    /// The `w:customStyle` value, `None` when it is absent.
    #[must_use]
    pub const fn custom(&self) -> Option<bool> {
        self.custom
    }

    /// The parent style identifier (`w:basedOn`).
    #[must_use]
    pub fn based_on(&self) -> Option<&str> {
        self.based_on.as_deref()
    }

    /// The next paragraph's style identifier (`w:next`).
    #[must_use]
    pub fn next(&self) -> Option<&str> {
        self.next.as_deref()
    }

    /// The linked style identifier (`w:link`).
    #[must_use]
    pub fn link(&self) -> Option<&str> {
        self.link.as_deref()
    }

    /// A flag's value: `None` when the style does not state it.
    #[must_use]
    pub const fn flag(&self, flag: Flag) -> Option<bool> {
        self.flags[flag.slot()]
    }

    /// The `w:uiPriority` value.
    #[must_use]
    pub const fn ui_priority(&self) -> Option<i64> {
        self.ui_priority
    }
}

/// The styles part, located: every style and the byte range of its element.
///
/// ```
/// use std::sync::Arc;
/// use litchi_docx::styles::catalog::Catalog;
///
/// let xml = br#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/></w:pPr><w:rPr><w:b/></w:rPr></w:style></w:styles>"#;
/// let catalog = Catalog::from_xml(Arc::new(xml.to_vec()))?;
/// let style = catalog.get("Heading1").unwrap();
/// assert_eq!(style.name(), Some("heading 1"));
/// assert_eq!(catalog.resolve("heading 1").map(|style| style.id()), Some("Heading1"));
/// assert_eq!(catalog.run_properties(style)?.bold, Some(true));
/// assert_eq!(catalog.paragraph_properties(style)?.keep_next, Some(true));
/// # Ok::<(), litchi_docx::Error>(())
/// ```
#[derive(Debug, Clone)]
pub struct Catalog {
    xml: Arc<Vec<u8>>,
    tree: Tree,
    entries: Vec<Entry>,
    strict: bool,
}

impl Catalog {
    /// Read a styles part payload.
    ///
    /// # Errors
    ///
    /// Returns an error for XML that is not well-formed, a root that is not
    /// `w:styles`, or a style without a `w:styleId`.
    pub fn from_xml(xml: Arc<Vec<u8>>) -> Result<Self> {
        // Only the styles are located; latent styles and document defaults
        // are checked and skipped.
        let tree = part_scan::scan_kept(&xml, 2, |local| local == b"style")?;
        if !tree.root()?.is(b"styles") {
            return Err(Error::InvalidFormat(
                "the styles part root is not w:styles".into(),
            ));
        }
        let strict = tree
            .bindings(0)
            .iter()
            .any(|(_, namespace)| namespace.as_slice() == STRICT_NAMESPACE);
        let mut entries = Vec::new();
        for element in tree.children(0).collect::<Vec<_>>() {
            if !tree.elements[element].is(b"style") {
                continue;
            }
            entries.push(entry(&tree, element)?);
        }
        Ok(Self {
            xml,
            tree,
            entries,
            strict,
        })
    }

    /// The payload the catalog was read from.
    #[must_use]
    pub fn xml(&self) -> &[u8] {
        &self.xml
    }

    /// Every style, in part order.
    #[must_use]
    pub fn entries(&self) -> &[Entry] {
        &self.entries
    }

    /// The style with identifier `id` (exact).
    #[must_use]
    pub fn get(&self, id: &str) -> Option<&Entry> {
        self.entries.iter().find(|entry| entry.id == id)
    }

    /// The first style, in part order, whose identifier or display name is
    /// `key` (exact), as Word's style references and OfficeCLI's
    /// `/styles/<key>` paths resolve.
    #[must_use]
    pub fn resolve(&self, key: &str) -> Option<&Entry> {
        self.entries
            .iter()
            .find(|entry| entry.id == key || entry.name.as_deref() == Some(key))
    }

    /// The style's paragraph properties (its `w:pPr`), empty when it has
    /// none.
    ///
    /// # Errors
    ///
    /// Returns an error for a malformed `w:pPr`.
    pub fn paragraph_properties(&self, entry: &Entry) -> Result<properties::Paragraph> {
        match self.wrapped(entry, b"pPr", "p")? {
            Some((wrapper, bindings)) => properties::parse_paragraph(&wrapper, &bindings),
            None => Ok(properties::Paragraph::default()),
        }
    }

    /// The style's run properties (its `w:rPr`), empty when it has none.
    ///
    /// # Errors
    ///
    /// Returns an error for a malformed `w:rPr`.
    pub fn run_properties(&self, entry: &Entry) -> Result<properties::Run> {
        match self.wrapped(entry, b"rPr", "r")? {
            Some((wrapper, bindings)) => properties::parse_run(&wrapper, &bindings),
            None => Ok(properties::Run::default()),
        }
    }

    /// Whether the part uses the Strict WordprocessingML namespace.
    #[must_use]
    pub const fn is_strict(&self) -> bool {
        self.strict
    }

    /// The `container` child of the style wrapped in a `root` element of the
    /// style's prefix, with the bindings in scope: a fragment the property
    /// readers and splicers accept.
    fn wrapped(
        &self,
        entry: &Entry,
        container: &[u8],
        root: &str,
    ) -> Result<Option<part_scan::Wrapped>> {
        let Some(child) = self.tree.child(entry.element, container) else {
            return Ok(None);
        };
        let element = &self.tree.elements[child];
        let prefix = &self.tree.elements[entry.element].prefix;
        Ok(Some((
            wrap(prefix, root, &self.xml[element.start..element.end]),
            self.tree.bindings(entry.element),
        )))
    }

    /// The payload with `style` appended before the root's end tag.
    pub(crate) fn with_added(&self, style: &NewStyle) -> Result<Vec<u8>> {
        if self.get(&style.id).is_some() {
            return Err(Error::InvalidFormat("the style identifier is taken".into()));
        }
        let root = self.tree.root()?;
        let prefix = root.prefix.to_vec();
        let element = style.element(&prefix, self.strict)?;
        if root.empty {
            let open = &self.xml[root.start..root.tag_end - 2];
            let mut replacement = open.to_vec();
            replacement.push(b'>');
            replacement.extend_from_slice(&element);
            replacement
                .extend_from_slice(format!("</{}>", qualified(&prefix, "styles")).as_bytes());
            return Ok(part_scan::splice(
                &self.xml,
                root.start,
                root.end,
                &replacement,
            ));
        }
        Ok(part_scan::splice(
            &self.xml, root.close, root.close, &element,
        ))
    }

    /// The payload with the style `id` removed.
    pub(crate) fn with_removed(&self, id: &str) -> Result<Vec<u8>> {
        let entry = self
            .get(id)
            .ok_or_else(|| Error::InvalidFormat("no style has this identifier".into()))?;
        let element = &self.tree.elements[entry.element];
        Ok(part_scan::splice(
            &self.xml,
            element.start,
            element.end,
            &[],
        ))
    }

    /// The payload with `change` applied to the style `id`, or `None` when
    /// it changes nothing.
    pub(crate) fn with_changed(&self, id: &str, change: &StyleChange) -> Result<Option<Vec<u8>>> {
        let entry = self
            .get(id)
            .ok_or_else(|| Error::InvalidFormat("no style has this identifier".into()))?;
        let style = &self.tree.elements[entry.element];
        let prefix = style.prefix.to_vec();
        let mut children: Vec<Child> = Vec::new();
        let mut set = |local: &'static str, bytes: Option<Vec<u8>>| {
            children.push(Child { local, bytes });
        };
        if let Some(name) = &change.name {
            set("name", Some(val_element(&prefix, "name", Some(name))));
        }
        for (local, value) in [
            ("basedOn", &change.based_on),
            ("next", &change.next),
            ("link", &change.link),
        ] {
            if let Some(value) = value {
                set(
                    local,
                    value
                        .as_deref()
                        .map(|value| val_element(&prefix, local, Some(value))),
                );
            }
        }
        for flag in Flag::ALL {
            if let Some(value) = change.flags[flag.slot()] {
                set(
                    flag.local_name(),
                    value.map(|on| on_off_element(&prefix, flag.local_name(), on)),
                );
            }
        }
        if let Some(priority) = change.ui_priority {
            set(
                "uiPriority",
                priority.map(|value| val_element(&prefix, "uiPriority", Some(&value.to_string()))),
            );
        }
        let bindings = self.tree.bindings(entry.element);
        if let Some(target) = &change.paragraph {
            let current = self.paragraph_properties(entry)?;
            if !properties::same_paragraph(&current, target) {
                let before = self.wrapped(entry, b"pPr", "p")?;
                let bytes =
                    respliced_paragraph(&prefix, before, &current, target, self.strict, &bindings)?;
                set("pPr", bytes);
            }
        }
        if let Some(target) = &change.run {
            let current = self.run_properties(entry)?;
            if !properties::same_run(&current, target) {
                let before = self.wrapped(entry, b"rPr", "r")?;
                let bytes = respliced_run(&prefix, before, &current, target, &bindings)?;
                set("rPr", bytes);
            }
        }
        let output = part_scan::rewrite_children(
            &self.xml,
            &self.tree,
            entry.element,
            "style",
            &children,
            CHILD_ORDER,
        )?;
        Ok((output != self.xml.as_slice()).then_some(output))
    }
}

fn entry(tree: &Tree, element: usize) -> Result<Entry> {
    let style = &tree.elements[element];
    let id = style
        .attribute(b"styleId")
        .ok_or_else(|| Error::InvalidFormat("a style has no w:styleId".into()))?
        .to_owned();
    let val = |local: &[u8]| {
        tree.child(element, local)
            .and_then(|child| tree.elements[child].attribute(b"val"))
            .map(str::to_owned)
    };
    let mut flags = [None; 6];
    for flag in Flag::ALL {
        if let Some(child) = tree.child(element, flag.local_name().as_bytes()) {
            flags[flag.slot()] = on_off(tree.elements[child].attribute(b"val"));
        }
    }
    Ok(Entry {
        element,
        kind: style.attribute(b"type").and_then(Type::from_xml),
        default: style
            .attribute(b"default")
            .and_then(|value| on_off(Some(value)))
            == Some(true),
        custom: style
            .attribute(b"customStyle")
            .and_then(|value| on_off(Some(value))),
        name: val(b"name"),
        based_on: val(b"basedOn"),
        next: val(b"next"),
        link: val(b"link"),
        ui_priority: val(b"uiPriority").and_then(|value| value.trim().parse().ok()),
        flags,
        id,
    })
}

/// A style to add.
///
/// ```
/// use litchi_docx::styles::Type;
/// use litchi_docx::styles::catalog::{NewStyle, StyleChange};
///
/// let style = NewStyle::new("Accent", Type::Paragraph)
///     .with_change(StyleChange::default().with_name("Accent").with_based_on(Some("Normal".into())));
/// assert_eq!(style.id(), "Accent");
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct NewStyle {
    id: String,
    kind: Type,
    default: bool,
    custom: Option<bool>,
    change: StyleChange,
}

impl NewStyle {
    /// A style `id` of type `kind`, with no other content.
    #[must_use]
    pub fn new(id: impl Into<String>, kind: Type) -> Self {
        Self {
            id: id.into(),
            kind,
            default: false,
            custom: None,
            change: StyleChange::default(),
        }
    }

    /// Mark the style as its type's default.
    #[must_use]
    pub const fn with_default(mut self, default: bool) -> Self {
        self.default = default;
        self
    }

    /// Write `w:customStyle`.
    #[must_use]
    pub const fn with_custom(mut self, custom: bool) -> Self {
        self.custom = Some(custom);
        self
    }

    /// The style's name, links, flags and properties.
    #[must_use]
    pub fn with_change(mut self, change: StyleChange) -> Self {
        self.change = change;
        self
    }

    /// The identifier.
    #[must_use]
    pub fn id(&self) -> &str {
        &self.id
    }

    fn element(&self, prefix: &[u8], strict: bool) -> Result<Vec<u8>> {
        if self.id.is_empty() {
            return Err(Error::InvalidFormat(
                "a style identifier cannot be empty".into(),
            ));
        }
        let q = |local: &str| qualified(prefix, local);
        let mut xml = format!(
            r#"<{} {}="{}" {}="{}""#,
            q("style"),
            q("type"),
            self.kind.to_xml(),
            q("styleId"),
            escape(&self.id)
        );
        if self.default {
            xml.push_str(&format!(r#" {}="1""#, q("default")));
        }
        if let Some(custom) = self.custom {
            xml.push_str(&format!(r#" {}="{}""#, q("customStyle"), u8::from(custom)));
        }
        xml.push('>');
        let mut bytes = xml.into_bytes();
        let change = &self.change;
        if let Some(name) = &change.name {
            bytes.extend_from_slice(&val_element(prefix, "name", Some(name)));
        }
        for (local, value) in [
            ("basedOn", &change.based_on),
            ("next", &change.next),
            ("link", &change.link),
        ] {
            if let Some(Some(value)) = value {
                bytes.extend_from_slice(&val_element(prefix, local, Some(value)));
            }
        }
        for flag in [Flag::AutoRedefine, Flag::Hidden] {
            if let Some(Some(on)) = change.flags[flag.slot()] {
                bytes.extend_from_slice(&on_off_element(prefix, flag.local_name(), on));
            }
        }
        if let Some(Some(priority)) = change.ui_priority {
            bytes.extend_from_slice(&val_element(
                prefix,
                "uiPriority",
                Some(&priority.to_string()),
            ));
        }
        for flag in [
            Flag::SemiHidden,
            Flag::UnhideWhenUsed,
            Flag::QuickFormat,
            Flag::Locked,
        ] {
            if let Some(Some(on)) = change.flags[flag.slot()] {
                bytes.extend_from_slice(&on_off_element(prefix, flag.local_name(), on));
            }
        }
        if let Some(paragraph) = &change.paragraph {
            if let Some(element) =
                properties::paragraph_properties_element(prefix, paragraph, strict)
                    .map_err(splice_error)?
            {
                bytes.extend_from_slice(&element);
            }
        }
        if let Some(run) = &change.run
            && let Some(element) =
                properties::run_properties_element(prefix, run).map_err(splice_error)?
        {
            bytes.extend_from_slice(&element);
        }
        bytes.extend_from_slice(format!("</{}>", q("style")).as_bytes());
        Ok(bytes)
    }
}

/// Changes to a style's name, links, flags, priority and properties; each
/// field left unset is kept as authored.
///
/// ```
/// use litchi_docx::styles::catalog::{Flag, StyleChange};
///
/// let change = StyleChange::default()
///     .with_name("Accent")
///     .with_flag(Flag::QuickFormat, Some(true))
///     .with_next(None);
/// assert!(!change.is_empty());
/// assert!(StyleChange::default().is_empty());
/// ```
#[derive(Debug, Clone, Default, PartialEq)]
pub struct StyleChange {
    name: Option<String>,
    based_on: Option<Option<String>>,
    next: Option<Option<String>>,
    link: Option<Option<String>>,
    flags: [Option<Option<bool>>; 6],
    ui_priority: Option<Option<i64>>,
    paragraph: Option<properties::Paragraph>,
    run: Option<properties::Run>,
}

impl StyleChange {
    /// Set the display name.
    #[must_use]
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// Set (`Some`) or remove (`None`) the parent style.
    #[must_use]
    pub fn with_based_on(mut self, based_on: Option<String>) -> Self {
        self.based_on = Some(based_on);
        self
    }

    /// Set or remove the next paragraph's style.
    #[must_use]
    pub fn with_next(mut self, next: Option<String>) -> Self {
        self.next = Some(next);
        self
    }

    /// Set or remove the linked style.
    #[must_use]
    pub fn with_link(mut self, link: Option<String>) -> Self {
        self.link = Some(link);
        self
    }

    /// Set (`Some`) or remove (`None`) a flag.
    #[must_use]
    pub const fn with_flag(mut self, flag: Flag, value: Option<bool>) -> Self {
        self.flags[flag.slot()] = Some(value);
        self
    }

    /// Set or remove the UI priority.
    #[must_use]
    pub const fn with_ui_priority(mut self, priority: Option<i64>) -> Self {
        self.ui_priority = Some(priority);
        self
    }

    /// Make the style's paragraph properties read as `paragraph`.
    #[must_use]
    pub fn with_paragraph(mut self, paragraph: properties::Paragraph) -> Self {
        self.paragraph = Some(paragraph);
        self
    }

    /// Make the style's run properties read as `run`.
    #[must_use]
    pub fn with_run(mut self, run: properties::Run) -> Self {
        self.run = Some(run);
        self
    }

    /// Whether the change sets nothing.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        *self == Self::default()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    fn bold() -> properties::Run {
        let mut run = properties::Run::default();
        run.bold = Some(true);
        run
    }

    fn italic() -> properties::Run {
        let mut run = properties::Run::default();
        run.italic = Some(true);
        run
    }

    fn keep_next() -> properties::Paragraph {
        let mut paragraph = properties::Paragraph::default();
        paragraph.keep_next = Some(true);
        paragraph
    }

    fn catalog(body: &str) -> Catalog {
        let xml = format!(r#"<?xml version="1.0"?><w:styles xmlns:w="{W}">{body}</w:styles>"#);
        Catalog::from_xml(Arc::new(xml.into_bytes())).unwrap()
    }

    #[test]
    fn reads_identity_flags_and_properties() {
        let catalog = catalog(
            r#"<w:docDefaults/><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style><w:style w:type="character" w:customStyle="1" w:styleId="Strong"><w:name w:val="Strong Text"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="22"/><w:semiHidden w:val="0"/><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr></w:style>"#,
        );
        assert_eq!(catalog.entries().len(), 2);
        let normal = catalog.get("Normal").unwrap();
        assert!(normal.is_default());
        assert_eq!(normal.flag(Flag::QuickFormat), Some(true));
        assert_eq!(normal.kind(), Some(Type::Paragraph));
        let strong = catalog.resolve("Strong Text").unwrap();
        assert_eq!(strong.id(), "Strong");
        assert_eq!(strong.custom(), Some(true));
        assert_eq!(strong.ui_priority(), Some(22));
        assert_eq!(strong.flag(Flag::SemiHidden), Some(false));
        assert_eq!(strong.based_on(), Some("DefaultParagraphFont"));
        assert_eq!(catalog.run_properties(strong).unwrap().bold, Some(true));
        assert_eq!(
            catalog.paragraph_properties(strong).unwrap(),
            properties::Paragraph::default()
        );
    }

    #[test]
    fn adds_changes_and_removes_in_place() {
        let catalog = catalog(
            r#"<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="24"/></w:rPr></w:style>"#,
        );
        let added = catalog
            .with_added(
                &NewStyle::new("Accent", Type::Paragraph).with_change(
                    StyleChange::default()
                        .with_name("Accent")
                        .with_based_on(Some("Normal".into()))
                        .with_flag(Flag::QuickFormat, Some(true))
                        .with_run(bold()),
                ),
            )
            .unwrap();
        let next = Catalog::from_xml(Arc::new(added)).unwrap();
        let accent = next.get("Accent").unwrap();
        assert_eq!(accent.name(), Some("Accent"));
        assert_eq!(accent.based_on(), Some("Normal"));
        assert_eq!(accent.flag(Flag::QuickFormat), Some(true));
        assert_eq!(next.run_properties(accent).unwrap().bold, Some(true));
        assert!(
            catalog
                .with_added(&NewStyle::new("Normal", Type::Paragraph))
                .is_err()
        );

        let mut run = next.run_properties(next.get("Normal").unwrap()).unwrap();
        run.italic = Some(true);
        let changed = next
            .with_changed(
                "Normal",
                &StyleChange::default()
                    .with_next(Some("Accent".into()))
                    .with_flag(Flag::Locked, Some(false))
                    .with_paragraph(keep_next())
                    .with_run(run),
            )
            .unwrap()
            .unwrap();
        let text = String::from_utf8(changed.clone()).unwrap();
        // CT_Style order: name, next, locked, pPr, rPr.
        let order = ["w:name", "w:next", "w:locked", "w:pPr", "w:rPr"]
            .map(|name| text.find(&format!("<{name}")).unwrap());
        assert!(order.windows(2).all(|pair| pair[0] < pair[1]), "{text}");
        let changed = Catalog::from_xml(Arc::new(changed)).unwrap();
        let normal = changed.get("Normal").unwrap();
        assert_eq!(normal.next(), Some("Accent"));
        assert_eq!(normal.flag(Flag::Locked), Some(false));
        assert_eq!(
            changed.paragraph_properties(normal).unwrap().keep_next,
            Some(true)
        );
        assert_eq!(changed.run_properties(normal).unwrap().italic, Some(true));
        // Accent is untouched byte for byte.
        let accent_bytes = |catalog: &Catalog| {
            let element = &catalog.tree.elements[catalog.get("Accent").unwrap().element];
            catalog.xml()[element.start..element.end].to_vec()
        };
        assert_eq!(accent_bytes(&changed), accent_bytes(&next));
        assert!(
            next.with_changed("Normal", &StyleChange::default())
                .unwrap()
                .is_none()
        );

        let removed = Catalog::from_xml(Arc::new(changed.with_removed("Accent").unwrap())).unwrap();
        assert!(removed.get("Accent").is_none());
        assert!(removed.get("Normal").is_some());
        assert!(removed.with_removed("Accent").is_err());
    }

    #[test]
    fn clears_properties_to_no_container_and_fills_an_empty_style() {
        let catalog = catalog(r#"<w:style w:type="character" w:styleId="E"/>"#);
        let changed = catalog
            .with_changed("E", &StyleChange::default().with_name("Empty"))
            .unwrap()
            .unwrap();
        let changed = Catalog::from_xml(Arc::new(changed)).unwrap();
        assert_eq!(changed.get("E").unwrap().name(), Some("Empty"));
        let with_run = changed
            .with_changed("E", &StyleChange::default().with_run(italic()))
            .unwrap()
            .unwrap();
        let with_run = Catalog::from_xml(Arc::new(with_run)).unwrap();
        let cleared = with_run
            .with_changed(
                "E",
                &StyleChange::default().with_run(properties::Run::default()),
            )
            .unwrap()
            .unwrap();
        assert!(!String::from_utf8(cleared).unwrap().contains("rPr"));
    }
}
