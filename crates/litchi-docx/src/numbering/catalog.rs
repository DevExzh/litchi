//! A source-located catalog of the numbering part (record 0319,
//! L-DOCX-R6).
//!
//! [`Catalog`] reads a `numbering.xml` payload once and keeps every
//! abstract definition (`w:abstractNum`) with its levels (`w:lvl`) and every
//! numbering instance (`w:num`), each with the byte range of its element.
//! Values are kept as the part spells them, as OfficeCLI reports them; a
//! level's paragraph and run properties are read on request through the
//! typed direct-property readers ([`crate::properties`]).
//!
//! The catalog also computes the edited payloads the main-document editor
//! publishes ([`crate::document::Edit::set_level`] and its siblings). Each
//! splices the raw part, so every other byte stays as authored.

use std::sync::Arc;

use crate::error::{Error, Result};
use crate::part_scan::{
    self, Child, Tree, on_off, on_off_element, qualified, respliced_paragraph, respliced_run,
    val_element, wrap,
};
use crate::properties;

/// The WordprocessingML namespace of ISO/IEC 29500 Strict documents.
const STRICT_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";

/// The `CT_AbstractNum` child order.
const ABSTRACT_ORDER: &[&str] = &[
    "nsid",
    "multiLevelType",
    "tmpl",
    "name",
    "styleLink",
    "numStyleLink",
    "lvl",
];

/// The `CT_Lvl` child order.
const LEVEL_ORDER: &[&str] = &[
    "start",
    "numFmt",
    "lvlRestart",
    "pStyle",
    "isLgl",
    "suff",
    "lvlText",
    "lvlPicBulletId",
    "legacy",
    "lvlJc",
    "pPr",
    "rPr",
];

/// The `CT_Num` child order.
const NUM_ORDER: &[&str] = &["abstractNumId", "lvlOverride"];

/// One level (`w:lvl`) of an abstract definition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Lvl {
    element: usize,
    ilvl: u8,
    start: Option<String>,
    format: Option<String>,
    text: Option<String>,
    justification: Option<String>,
    suffix: Option<String>,
    restart: Option<String>,
    legal: Option<bool>,
    paragraph_style: Option<String>,
}

impl Lvl {
    /// The level index (`w:ilvl`, 0–8).
    #[must_use]
    pub const fn ilvl(&self) -> u8 {
        self.ilvl
    }

    /// `w:start/@w:val`, as written.
    #[must_use]
    pub fn start(&self) -> Option<&str> {
        self.start.as_deref()
    }

    /// `w:numFmt/@w:val` (`decimal`, `bullet`, …), as written.
    #[must_use]
    pub fn format(&self) -> Option<&str> {
        self.format.as_deref()
    }

    /// `w:lvlText/@w:val` (`%1.`, `•`).
    #[must_use]
    pub fn text(&self) -> Option<&str> {
        self.text.as_deref()
    }

    /// `w:lvlJc/@w:val`, as written.
    #[must_use]
    pub fn justification(&self) -> Option<&str> {
        self.justification.as_deref()
    }

    /// `w:suff/@w:val` (`tab`, `space`, `nothing`).
    #[must_use]
    pub fn suffix(&self) -> Option<&str> {
        self.suffix.as_deref()
    }

    /// `w:lvlRestart/@w:val`, as written.
    #[must_use]
    pub fn restart(&self) -> Option<&str> {
        self.restart.as_deref()
    }

    /// `w:isLgl`: legal numbering, `None` when not stated.
    #[must_use]
    pub const fn legal(&self) -> Option<bool> {
        self.legal
    }

    /// `w:pStyle/@w:val`: the paragraph style linked to the level.
    #[must_use]
    pub fn paragraph_style(&self) -> Option<&str> {
        self.paragraph_style.as_deref()
    }
}

/// One abstract numbering definition (`w:abstractNum`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AbstractNum {
    element: usize,
    id: u32,
    multi_level_type: Option<String>,
    name: Option<String>,
    style_link: Option<String>,
    num_style_link: Option<String>,
    levels: Vec<Lvl>,
}

impl AbstractNum {
    /// `w:abstractNumId`.
    #[must_use]
    pub const fn id(&self) -> u32 {
        self.id
    }

    /// `w:multiLevelType/@w:val`, as written.
    #[must_use]
    pub fn multi_level_type(&self) -> Option<&str> {
        self.multi_level_type.as_deref()
    }

    /// `w:name/@w:val`.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// `w:styleLink/@w:val`.
    #[must_use]
    pub fn style_link(&self) -> Option<&str> {
        self.style_link.as_deref()
    }

    /// `w:numStyleLink/@w:val`.
    #[must_use]
    pub fn num_style_link(&self) -> Option<&str> {
        self.num_style_link.as_deref()
    }

    /// The levels, in part order.
    #[must_use]
    pub fn levels(&self) -> &[Lvl] {
        &self.levels
    }

    /// The level with index `ilvl`.
    #[must_use]
    pub fn level(&self, ilvl: u8) -> Option<&Lvl> {
        self.levels.iter().find(|level| level.ilvl == ilvl)
    }
}

/// One numbering instance (`w:num`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Num {
    element: usize,
    id: u32,
    abstract_num_id: Option<u32>,
    overrides: Vec<(u8, Option<String>)>,
}

impl Num {
    /// `w:numId`.
    #[must_use]
    pub const fn id(&self) -> u32 {
        self.id
    }

    /// The abstract definition it instantiates (`w:abstractNumId/@w:val`).
    #[must_use]
    pub const fn abstract_num_id(&self) -> Option<u32> {
        self.abstract_num_id
    }

    /// The level overrides: each level index with its `w:startOverride`
    /// value, as written.
    #[must_use]
    pub fn overrides(&self) -> &[(u8, Option<String>)] {
        &self.overrides
    }
}

/// The numbering part, located.
///
/// ```
/// use std::sync::Arc;
/// use litchi_docx::numbering::catalog::Catalog;
///
/// let xml = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#;
/// let catalog = Catalog::from_xml(Arc::new(xml.to_vec()))?;
/// let definition = catalog.abstract_num(0).unwrap();
/// assert_eq!(definition.level(0).unwrap().format(), Some("decimal"));
/// assert_eq!(catalog.num(1).unwrap().abstract_num_id(), Some(0));
/// assert_eq!(catalog.format_of(1, 0), Some("decimal"));
/// # Ok::<(), litchi_docx::Error>(())
/// ```
#[derive(Debug, Clone)]
pub struct Catalog {
    xml: Arc<Vec<u8>>,
    tree: Tree,
    abstracts: Vec<AbstractNum>,
    nums: Vec<Num>,
    strict: bool,
}

impl Catalog {
    /// Read a numbering part payload.
    ///
    /// # Errors
    ///
    /// Returns an error for XML that is not well-formed, a root that is not
    /// `w:numbering`, or an identifier that is not a number.
    pub fn from_xml(xml: Arc<Vec<u8>>) -> Result<Self> {
        let tree = part_scan::scan(&xml, 3)?;
        if !tree.root()?.is(b"numbering") {
            return Err(Error::InvalidFormat(
                "the numbering part root is not w:numbering".into(),
            ));
        }
        let strict = tree
            .bindings(0)
            .iter()
            .any(|(_, namespace)| namespace.as_slice() == STRICT_NAMESPACE);
        let mut abstracts = Vec::new();
        let mut nums = Vec::new();
        for element in tree.children(0).collect::<Vec<_>>() {
            let node = &tree.elements[element];
            if node.is(b"abstractNum") {
                abstracts.push(abstract_num(&tree, element)?);
            } else if node.is(b"num") {
                nums.push(num(&tree, element)?);
            }
        }
        Ok(Self {
            xml,
            tree,
            abstracts,
            nums,
            strict,
        })
    }

    /// The payload the catalog was read from.
    #[must_use]
    pub fn xml(&self) -> &[u8] {
        &self.xml
    }

    /// Every abstract definition, in part order.
    #[must_use]
    pub fn abstract_nums(&self) -> &[AbstractNum] {
        &self.abstracts
    }

    /// Every numbering instance, in part order.
    #[must_use]
    pub fn nums(&self) -> &[Num] {
        &self.nums
    }

    /// The abstract definition `id`.
    #[must_use]
    pub fn abstract_num(&self, id: u32) -> Option<&AbstractNum> {
        self.abstracts.iter().find(|definition| definition.id == id)
    }

    /// The numbering instance `id`.
    #[must_use]
    pub fn num(&self, id: u32) -> Option<&Num> {
        self.nums.iter().find(|num| num.id == id)
    }

    /// The level `ilvl` that numbering instance `num_id` displays: its
    /// abstract definition's level, following a `w:numStyleLink` once
    /// through the style's own definition when the definition has no level
    /// of its own. Level overrides that replace a whole level are not
    /// followed.
    #[must_use]
    pub fn level_of(&self, num_id: u32, ilvl: u8) -> Option<&Lvl> {
        let definition = self.abstract_num(self.num(num_id)?.abstract_num_id?)?;
        if let Some(level) = definition.level(ilvl) {
            return Some(level);
        }
        let link = definition.num_style_link.as_deref()?;
        self.abstracts
            .iter()
            .find(|other| other.style_link.as_deref() == Some(link))
            .and_then(|other| other.level(ilvl))
    }

    /// The number format (`w:numFmt`) of level `ilvl` of instance `num_id`.
    #[must_use]
    pub fn format_of(&self, num_id: u32, ilvl: u8) -> Option<&str> {
        self.level_of(num_id, ilvl)?.format()
    }

    /// A level's paragraph properties (its `w:pPr`), empty when it has none.
    ///
    /// # Errors
    ///
    /// Returns an error for a malformed `w:pPr`.
    pub fn level_paragraph_properties(&self, level: &Lvl) -> Result<properties::Paragraph> {
        match self.wrapped(level.element, b"pPr", "p") {
            Some((wrapper, bindings)) => properties::parse_paragraph(&wrapper, &bindings),
            None => Ok(properties::Paragraph::default()),
        }
    }

    /// A level's marker run properties (its `w:rPr`), empty when it has
    /// none.
    ///
    /// # Errors
    ///
    /// Returns an error for a malformed `w:rPr`.
    pub fn level_run_properties(&self, level: &Lvl) -> Result<properties::Run> {
        match self.wrapped(level.element, b"rPr", "r") {
            Some((wrapper, bindings)) => properties::parse_run(&wrapper, &bindings),
            None => Ok(properties::Run::default()),
        }
    }

    /// Whether the part uses the Strict WordprocessingML namespace.
    #[must_use]
    pub const fn is_strict(&self) -> bool {
        self.strict
    }

    /// The smallest abstract definition identifier above every one in use.
    #[must_use]
    pub fn next_abstract_num_id(&self) -> u32 {
        self.abstracts
            .iter()
            .map(|definition| definition.id.saturating_add(1))
            .max()
            .unwrap_or(0)
    }

    /// The smallest instance identifier above every one in use (from 1).
    #[must_use]
    pub fn next_num_id(&self) -> u32 {
        self.nums
            .iter()
            .map(|num| num.id.saturating_add(1))
            .max()
            .unwrap_or(1)
            .max(1)
    }

    fn wrapped(&self, element: usize, container: &[u8], root: &str) -> Option<part_scan::Wrapped> {
        let child = self.tree.child(element, container)?;
        let node = &self.tree.elements[child];
        let prefix = &self.tree.elements[element].prefix;
        Some((
            wrap(prefix, root, &self.xml[node.start..node.end]),
            self.tree.bindings(element),
        ))
    }

    fn prefix(&self) -> Result<Vec<u8>> {
        Ok(self.tree.root()?.prefix.to_vec())
    }

    fn level_element(&self, abstract_num_id: u32, ilvl: u8) -> Result<(&AbstractNum, &Lvl)> {
        let definition = self
            .abstract_num(abstract_num_id)
            .ok_or_else(|| missing("no abstract definition has this identifier"))?;
        let level = definition
            .level(ilvl)
            .ok_or_else(|| missing("the abstract definition has no such level"))?;
        Ok((definition, level))
    }

    /// The payload with `change` applied to level `ilvl` of the abstract
    /// definition `abstract_num_id`, or `None` when nothing changes.
    pub(crate) fn with_level_changed(
        &self,
        abstract_num_id: u32,
        ilvl: u8,
        change: &LevelChange,
    ) -> Result<Option<Vec<u8>>> {
        let (_, level) = self.level_element(abstract_num_id, ilvl)?;
        let prefix = self.tree.elements[level.element].prefix.to_vec();
        let mut children = Vec::new();
        for (local, value) in [
            ("start", &change.start),
            ("numFmt", &change.format),
            ("lvlRestart", &change.restart),
            ("pStyle", &change.paragraph_style),
            ("suff", &change.suffix),
            ("lvlText", &change.text),
            ("lvlJc", &change.justification),
        ] {
            if let Some(value) = value {
                children.push(Child {
                    local,
                    bytes: value
                        .as_deref()
                        .map(|value| val_element(&prefix, local, Some(value))),
                });
            }
        }
        if let Some(legal) = change.legal {
            children.push(Child {
                local: "isLgl",
                bytes: legal.map(|on| on_off_element(&prefix, "isLgl", on)),
            });
        }
        let bindings = self.tree.bindings(level.element);
        if let Some(target) = &change.paragraph {
            let current = self.level_paragraph_properties(level)?;
            if !properties::same_paragraph(&current, target) {
                let before = self.wrapped(level.element, b"pPr", "p");
                children.push(Child {
                    local: "pPr",
                    bytes: respliced_paragraph(
                        &prefix,
                        before,
                        &current,
                        target,
                        self.strict,
                        &bindings,
                    )?,
                });
            }
        }
        if let Some(target) = &change.run {
            let current = self.level_run_properties(level)?;
            if !properties::same_run(&current, target) {
                let before = self.wrapped(level.element, b"rPr", "r");
                children.push(Child {
                    local: "rPr",
                    bytes: respliced_run(&prefix, before, &current, target, &bindings)?,
                });
            }
        }
        let output = part_scan::rewrite_children(
            &self.xml,
            &self.tree,
            level.element,
            "lvl",
            &children,
            LEVEL_ORDER,
        )?;
        Ok((output != self.xml.as_slice()).then_some(output))
    }

    /// The payload with `change` applied to the abstract definition `id`,
    /// or `None` when nothing changes.
    pub(crate) fn with_abstract_num_changed(
        &self,
        id: u32,
        change: &AbstractNumChange,
    ) -> Result<Option<Vec<u8>>> {
        let definition = self
            .abstract_num(id)
            .ok_or_else(|| missing("no abstract definition has this identifier"))?;
        let prefix = self.tree.elements[definition.element].prefix.to_vec();
        let mut children = Vec::new();
        for (local, value) in [
            ("multiLevelType", &change.multi_level_type),
            ("name", &change.name),
            ("styleLink", &change.style_link),
            ("numStyleLink", &change.num_style_link),
        ] {
            if let Some(value) = value {
                children.push(Child {
                    local,
                    bytes: value
                        .as_deref()
                        .map(|value| val_element(&prefix, local, Some(value))),
                });
            }
        }
        let output = part_scan::rewrite_children(
            &self.xml,
            &self.tree,
            definition.element,
            "abstractNum",
            &children,
            ABSTRACT_ORDER,
        )?;
        Ok((output != self.xml.as_slice()).then_some(output))
    }

    /// The payload with instance `id` pointed at abstract definition
    /// `abstract_num_id`, or `None` when it already is.
    pub(crate) fn with_num_retargeted(
        &self,
        id: u32,
        abstract_num_id: u32,
    ) -> Result<Option<Vec<u8>>> {
        let num = self
            .num(id)
            .ok_or_else(|| missing("no numbering instance has this identifier"))?;
        if self.abstract_num(abstract_num_id).is_none() {
            return Err(missing("no abstract definition has this identifier"));
        }
        let prefix = self.tree.elements[num.element].prefix.to_vec();
        let children = [Child {
            local: "abstractNumId",
            bytes: Some(val_element(
                &prefix,
                "abstractNumId",
                Some(&abstract_num_id.to_string()),
            )),
        }];
        let output = part_scan::rewrite_children(
            &self.xml,
            &self.tree,
            num.element,
            "num",
            &children,
            NUM_ORDER,
        )?;
        Ok((output != self.xml.as_slice()).then_some(output))
    }

    /// The payload with a new numbering instance appended after the last
    /// one (before `w:numIdMacAtCleanup`), and its identifier.
    pub(crate) fn with_num_added(&self, num: &NewNum) -> Result<(Vec<u8>, u32)> {
        if self.abstract_num(num.abstract_num_id).is_none() {
            return Err(missing("no abstract definition has this identifier"));
        }
        let id = match num.id {
            Some(id) if self.num(id).is_some() || id == 0 => {
                return Err(Error::InvalidFormat(
                    "the numbering identifier is taken".into(),
                ));
            },
            Some(id) => id,
            None => self.next_num_id(),
        };
        let prefix = self.prefix()?;
        let q = |local: &str| qualified(&prefix, local);
        let mut element = format!(r#"<{} {}="{id}">"#, q("num"), q("numId"));
        element.push_str(&String::from_utf8_lossy(&val_element(
            &prefix,
            "abstractNumId",
            Some(&num.abstract_num_id.to_string()),
        )));
        for (ilvl, start) in &num.overrides {
            element.push_str(&format!(r#"<{} {}="{ilvl}">"#, q("lvlOverride"), q("ilvl")));
            element.push_str(&String::from_utf8_lossy(&val_element(
                &prefix,
                "startOverride",
                Some(&start.to_string()),
            )));
            element.push_str(&format!("</{}>", q("lvlOverride")));
        }
        element.push_str(&format!("</{}>", q("num")));
        let anchor = self.insertion_anchor(b"num", &[b"numIdMacAtCleanup"])?;
        Ok((self.inserted(anchor, element.as_bytes())?, id))
    }

    /// The payload with a new abstract definition inserted after the last
    /// one (before the first `w:num`), and its identifier.
    pub(crate) fn with_abstract_num_added(
        &self,
        definition: &NewAbstractNum,
    ) -> Result<(Vec<u8>, u32)> {
        let id = match definition.id {
            Some(id) if self.abstract_num(id).is_some() => {
                return Err(Error::InvalidFormat(
                    "the abstract definition identifier is taken".into(),
                ));
            },
            Some(id) => id,
            None => self.next_abstract_num_id(),
        };
        let prefix = self.prefix()?;
        let bytes = definition.element(&prefix, id, self.strict)?;
        let anchor = self.insertion_anchor(b"abstractNum", &[b"num", b"numIdMacAtCleanup"])?;
        Ok((self.inserted(anchor, &bytes)?, id))
    }

    /// The payload with `level` added to abstract definition `id`,
    /// replacing the level of the same index when there is one.
    pub(crate) fn with_level_added(&self, id: u32, level: &NewLevel) -> Result<Vec<u8>> {
        let definition = self
            .abstract_num(id)
            .ok_or_else(|| missing("no abstract definition has this identifier"))?;
        let prefix = self.tree.elements[definition.element].prefix.to_vec();
        let bytes = level.element(&prefix, self.strict)?;
        if let Some(existing) = definition.level(level.ilvl) {
            let node = &self.tree.elements[existing.element];
            return Ok(part_scan::splice(&self.xml, node.start, node.end, &bytes));
        }
        // After the last level with a smaller index, else before the first
        // with a larger one, else at the end of the definition.
        let node = &self.tree.elements[definition.element];
        let offset = definition
            .levels
            .iter()
            .rev()
            .find(|other| other.ilvl < level.ilvl)
            .map(|other| self.tree.elements[other.element].end)
            .or_else(|| {
                definition
                    .levels
                    .iter()
                    .find(|other| other.ilvl > level.ilvl)
                    .map(|other| self.tree.elements[other.element].start)
            });
        match offset {
            Some(offset) => Ok(part_scan::splice(&self.xml, offset, offset, &bytes)),
            None if node.empty => {
                let mut replacement = self.xml[node.start..node.tag_end - 2].to_vec();
                replacement.push(b'>');
                replacement.extend_from_slice(&bytes);
                replacement.extend_from_slice(
                    format!("</{}>", qualified(&prefix, "abstractNum")).as_bytes(),
                );
                Ok(part_scan::splice(
                    &self.xml,
                    node.start,
                    node.end,
                    &replacement,
                ))
            },
            None => Ok(part_scan::splice(&self.xml, node.close, node.close, &bytes)),
        }
    }

    /// The payload without numbering instance `id`.
    pub(crate) fn with_num_removed(&self, id: u32) -> Result<Vec<u8>> {
        let num = self
            .num(id)
            .ok_or_else(|| missing("no numbering instance has this identifier"))?;
        let node = &self.tree.elements[num.element];
        Ok(part_scan::splice(&self.xml, node.start, node.end, &[]))
    }

    /// The payload without abstract definition `id` and the numbering
    /// instances of it, which cannot exist without it.
    pub(crate) fn with_abstract_num_removed(&self, id: u32) -> Result<Vec<u8>> {
        let definition = self
            .abstract_num(id)
            .ok_or_else(|| missing("no abstract definition has this identifier"))?;
        let mut ranges: Vec<(usize, usize)> = self
            .nums
            .iter()
            .filter(|num| num.abstract_num_id == Some(id))
            .map(|num| {
                let node = &self.tree.elements[num.element];
                (node.start, node.end)
            })
            .collect();
        let node = &self.tree.elements[definition.element];
        ranges.push((node.start, node.end));
        ranges.sort_unstable_by(|left, right| right.cmp(left));
        let mut output = self.xml.to_vec();
        for (start, end) in ranges {
            output = part_scan::splice(&output, start, end, &[]);
        }
        Ok(output)
    }

    /// The payload without level `ilvl` of abstract definition `id`.
    pub(crate) fn with_level_removed(&self, id: u32, ilvl: u8) -> Result<Vec<u8>> {
        let (_, level) = self.level_element(id, ilvl)?;
        let node = &self.tree.elements[level.element];
        Ok(part_scan::splice(&self.xml, node.start, node.end, &[]))
    }

    /// Where a new direct child of the root named `local` goes: after the
    /// last child of that name, else before the first child named in
    /// `before`, else before the root's end tag. `None` for an empty root.
    fn insertion_anchor(&self, local: &[u8], before: &[&[u8]]) -> Result<Option<usize>> {
        let root = self.tree.root()?;
        if root.empty {
            return Ok(None);
        }
        let children: Vec<usize> = self.tree.children(0).collect();
        if let Some(last) = children
            .iter()
            .rev()
            .find(|child| self.tree.elements[**child].is(local))
        {
            return Ok(Some(self.tree.elements[*last].end));
        }
        if let Some(first) = children.iter().find(|child| {
            before
                .iter()
                .any(|name| self.tree.elements[**child].is(name))
        }) {
            return Ok(Some(self.tree.elements[*first].start));
        }
        Ok(Some(root.close))
    }

    fn inserted(&self, anchor: Option<usize>, bytes: &[u8]) -> Result<Vec<u8>> {
        match anchor {
            Some(offset) => Ok(part_scan::splice(&self.xml, offset, offset, bytes)),
            None => {
                let root = self.tree.root()?;
                let mut replacement = self.xml[root.start..root.tag_end - 2].to_vec();
                replacement.push(b'>');
                replacement.extend_from_slice(bytes);
                replacement.extend_from_slice(
                    format!("</{}>", qualified(&root.prefix, "numbering")).as_bytes(),
                );
                Ok(part_scan::splice(
                    &self.xml,
                    root.start,
                    root.end,
                    &replacement,
                ))
            },
        }
    }
}

fn missing(message: &str) -> Error {
    Error::InvalidFormat(message.to_owned())
}

fn number(value: Option<&str>, what: &str) -> Result<u32> {
    value
        .and_then(|value| value.trim().parse().ok())
        .ok_or_else(|| Error::InvalidFormat(format!("a numbering {what} is not a number")))
}

fn abstract_num(tree: &Tree, element: usize) -> Result<AbstractNum> {
    let node = &tree.elements[element];
    let id = number(node.attribute(b"abstractNumId"), "abstractNumId")?;
    let val = |local: &[u8]| {
        tree.child(element, local)
            .and_then(|child| tree.elements[child].attribute(b"val"))
            .map(str::to_owned)
    };
    let mut levels = Vec::new();
    for child in tree.children(element).collect::<Vec<_>>() {
        if tree.elements[child].is(b"lvl") {
            levels.push(lvl(tree, child)?);
        }
    }
    Ok(AbstractNum {
        element,
        id,
        multi_level_type: val(b"multiLevelType"),
        name: val(b"name"),
        style_link: val(b"styleLink"),
        num_style_link: val(b"numStyleLink"),
        levels,
    })
}

fn lvl(tree: &Tree, element: usize) -> Result<Lvl> {
    let node = &tree.elements[element];
    let ilvl = node
        .attribute(b"ilvl")
        .and_then(|value| value.trim().parse::<u8>().ok())
        .filter(|ilvl| *ilvl <= 8)
        .ok_or_else(|| Error::InvalidFormat("a numbering level has no valid w:ilvl".into()))?;
    let val = |local: &[u8]| {
        tree.child(element, local)
            .and_then(|child| tree.elements[child].attribute(b"val"))
            .map(str::to_owned)
    };
    Ok(Lvl {
        element,
        ilvl,
        start: val(b"start"),
        format: val(b"numFmt"),
        text: val(b"lvlText"),
        justification: val(b"lvlJc"),
        suffix: val(b"suff"),
        restart: val(b"lvlRestart"),
        legal: tree
            .child(element, b"isLgl")
            .and_then(|child| on_off(tree.elements[child].attribute(b"val"))),
        paragraph_style: val(b"pStyle"),
    })
}

fn num(tree: &Tree, element: usize) -> Result<Num> {
    let node = &tree.elements[element];
    let id = number(node.attribute(b"numId"), "numId")?;
    let abstract_num_id = tree
        .child(element, b"abstractNumId")
        .and_then(|child| tree.elements[child].attribute(b"val"))
        .and_then(|value| value.trim().parse().ok());
    let mut overrides = Vec::new();
    for child in tree.children(element).collect::<Vec<_>>() {
        let override_node = &tree.elements[child];
        if !override_node.is(b"lvlOverride") {
            continue;
        }
        let Some(ilvl) = override_node
            .attribute(b"ilvl")
            .and_then(|value| value.trim().parse::<u8>().ok())
        else {
            continue;
        };
        let start = tree
            .child(child, b"startOverride")
            .and_then(|start| tree.elements[start].attribute(b"val"))
            .map(str::to_owned);
        overrides.push((ilvl, start));
    }
    Ok(Num {
        element,
        id,
        abstract_num_id,
        overrides,
    })
}

/// Changes to one level; each field left unset is kept as authored, and an
/// inner `None` removes the element.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct LevelChange {
    start: Option<Option<String>>,
    format: Option<Option<String>>,
    restart: Option<Option<String>>,
    paragraph_style: Option<Option<String>>,
    legal: Option<Option<bool>>,
    suffix: Option<Option<String>>,
    text: Option<Option<String>>,
    justification: Option<Option<String>>,
    paragraph: Option<properties::Paragraph>,
    run: Option<properties::Run>,
}

impl LevelChange {
    /// Set or remove `w:start`.
    #[must_use]
    pub fn with_start(mut self, start: Option<i64>) -> Self {
        self.start = Some(start.map(|start| start.to_string()));
        self
    }

    /// Set or remove `w:numFmt` (`decimal`, `bullet`, …).
    #[must_use]
    pub fn with_format(mut self, format: Option<String>) -> Self {
        self.format = Some(format);
        self
    }

    /// Set or remove `w:lvlRestart`.
    #[must_use]
    pub fn with_restart(mut self, restart: Option<i64>) -> Self {
        self.restart = Some(restart.map(|restart| restart.to_string()));
        self
    }

    /// Set or remove `w:isLgl`.
    #[must_use]
    pub const fn with_legal(mut self, legal: Option<bool>) -> Self {
        self.legal = Some(legal);
        self
    }

    /// Set or remove `w:suff` (`tab`, `space`, `nothing`).
    #[must_use]
    pub fn with_suffix(mut self, suffix: Option<String>) -> Self {
        self.suffix = Some(suffix);
        self
    }

    /// Set or remove `w:lvlText`.
    #[must_use]
    pub fn with_text(mut self, text: Option<String>) -> Self {
        self.text = Some(text);
        self
    }

    /// Set or remove `w:lvlJc` (`left`, `center`, `right`, …).
    #[must_use]
    pub fn with_justification(mut self, justification: Option<String>) -> Self {
        self.justification = Some(justification);
        self
    }

    /// Make the level's paragraph properties read as `paragraph`.
    #[must_use]
    pub fn with_paragraph(mut self, paragraph: properties::Paragraph) -> Self {
        self.paragraph = Some(paragraph);
        self
    }

    /// Make the level's marker run properties read as `run`.
    #[must_use]
    pub fn with_run(mut self, run: properties::Run) -> Self {
        self.run = Some(run);
        self
    }
}

/// Changes to an abstract definition's type, name and style links.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct AbstractNumChange {
    multi_level_type: Option<Option<String>>,
    name: Option<Option<String>>,
    style_link: Option<Option<String>>,
    num_style_link: Option<Option<String>>,
}

impl AbstractNumChange {
    /// Set or remove `w:multiLevelType`.
    #[must_use]
    pub fn with_multi_level_type(mut self, value: Option<String>) -> Self {
        self.multi_level_type = Some(value);
        self
    }

    /// Set or remove `w:name`.
    #[must_use]
    pub fn with_name(mut self, value: Option<String>) -> Self {
        self.name = Some(value);
        self
    }

    /// Set or remove `w:styleLink`.
    #[must_use]
    pub fn with_style_link(mut self, value: Option<String>) -> Self {
        self.style_link = Some(value);
        self
    }

    /// Set or remove `w:numStyleLink`.
    #[must_use]
    pub fn with_num_style_link(mut self, value: Option<String>) -> Self {
        self.num_style_link = Some(value);
        self
    }
}

/// A numbering instance to add.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NewNum {
    id: Option<u32>,
    abstract_num_id: u32,
    overrides: Vec<(u8, i64)>,
}

impl NewNum {
    /// An instance of abstract definition `abstract_num_id`.
    #[must_use]
    pub const fn new(abstract_num_id: u32) -> Self {
        Self {
            id: None,
            abstract_num_id,
            overrides: Vec::new(),
        }
    }

    /// Use identifier `id` instead of the next free one.
    #[must_use]
    pub const fn with_id(mut self, id: u32) -> Self {
        self.id = Some(id);
        self
    }

    /// Restart level `ilvl` at `start` (`w:lvlOverride/w:startOverride`).
    #[must_use]
    pub fn with_start_override(mut self, ilvl: u8, start: i64) -> Self {
        self.overrides.retain(|(level, _)| *level != ilvl);
        self.overrides.push((ilvl, start));
        self.overrides.sort_by_key(|(level, _)| *level);
        self
    }
}

/// One level of a new abstract definition, or a level to add.
#[derive(Debug, Clone, PartialEq)]
pub struct NewLevel {
    ilvl: u8,
    start: i64,
    format: String,
    text: String,
    justification: String,
    suffix: Option<String>,
    paragraph: properties::Paragraph,
    run: properties::Run,
}

impl NewLevel {
    /// Level `ilvl` (0–8) starting at 1, formatted `format`, with marker
    /// `text`, left-justified.
    #[must_use]
    pub fn new(ilvl: u8, format: impl Into<String>, text: impl Into<String>) -> Self {
        Self {
            ilvl,
            start: 1,
            format: format.into(),
            text: text.into(),
            justification: "left".to_owned(),
            suffix: None,
            paragraph: properties::Paragraph::default(),
            run: properties::Run::default(),
        }
    }

    /// OfficeCLI's default level `ilvl` of a definition: bullets `•`, `◦`,
    /// `▪` or numbers `%N.` in decimal, lower letter and lower Roman by
    /// level, indented 720 twips per level with a 360-twip hanging indent.
    #[must_use]
    pub fn template(ilvl: u8, bullet: bool) -> Self {
        let (format, text) = if bullet {
            (
                "bullet",
                ["\u{2022}", "\u{25E6}", "\u{25AA}"][usize::from(ilvl % 3)].to_owned(),
            )
        } else {
            (
                ["decimal", "lowerLetter", "lowerRoman"][usize::from(ilvl % 3)],
                format!("%{}.", ilvl + 1),
            )
        };
        let mut paragraph = properties::Paragraph::default();
        paragraph.indentation = Some(properties::Indentation {
            left: Some(properties::Twips(720 * (i64::from(ilvl) + 1))),
            hanging: Some(properties::Twips(360)),
            ..Default::default()
        });
        Self::new(ilvl, format, text).with_paragraph(paragraph)
    }

    /// The level index.
    #[must_use]
    pub const fn ilvl(&self) -> u8 {
        self.ilvl
    }

    /// The level's start value.
    #[must_use]
    pub const fn start(&self) -> i64 {
        self.start
    }

    /// The level's paragraph properties (its indentation).
    #[must_use]
    pub const fn paragraph(&self) -> &properties::Paragraph {
        &self.paragraph
    }

    /// Start at `start`.
    #[must_use]
    pub const fn with_start(mut self, start: i64) -> Self {
        self.start = start;
        self
    }

    /// Use number format `format`.
    #[must_use]
    pub fn with_format(mut self, format: impl Into<String>) -> Self {
        self.format = format.into();
        self
    }

    /// Use marker text `text`.
    #[must_use]
    pub fn with_text(mut self, text: impl Into<String>) -> Self {
        self.text = text.into();
        self
    }

    /// Justify the marker `justification` (`left`, `center`, `right`).
    #[must_use]
    pub fn with_justification(mut self, justification: impl Into<String>) -> Self {
        self.justification = justification.into();
        self
    }

    /// Follow the marker with `suffix` (`tab`, `space`, `nothing`).
    #[must_use]
    pub fn with_suffix(mut self, suffix: impl Into<String>) -> Self {
        self.suffix = Some(suffix.into());
        self
    }

    /// The level's paragraph properties (indentation).
    #[must_use]
    pub fn with_paragraph(mut self, paragraph: properties::Paragraph) -> Self {
        self.paragraph = paragraph;
        self
    }

    /// The marker's run properties.
    #[must_use]
    pub fn with_run(mut self, run: properties::Run) -> Self {
        self.run = run;
        self
    }

    fn element(&self, prefix: &[u8], strict: bool) -> Result<Vec<u8>> {
        if self.ilvl > 8 {
            return Err(Error::InvalidFormat(
                "a numbering level index is above 8".into(),
            ));
        }
        let q = |local: &str| qualified(prefix, local);
        let mut bytes = format!(r#"<{} {}="{}">"#, q("lvl"), q("ilvl"), self.ilvl).into_bytes();
        bytes.extend_from_slice(&val_element(prefix, "start", Some(&self.start.to_string())));
        bytes.extend_from_slice(&val_element(prefix, "numFmt", Some(&self.format)));
        if let Some(suffix) = &self.suffix {
            bytes.extend_from_slice(&val_element(prefix, "suff", Some(suffix)));
        }
        bytes.extend_from_slice(&val_element(prefix, "lvlText", Some(&self.text)));
        let justification = match (strict, self.justification.as_str()) {
            (true, "left") => "start",
            (true, "right") => "end",
            (_, other) => other,
        };
        bytes.extend_from_slice(&val_element(prefix, "lvlJc", Some(justification)));
        if let Some(element) =
            properties::paragraph_properties_element(prefix, &self.paragraph, strict)
                .map_err(part_scan::splice_error)?
        {
            bytes.extend_from_slice(&element);
        }
        if let Some(element) = properties::run_properties_element(prefix, &self.run)
            .map_err(part_scan::splice_error)?
        {
            bytes.extend_from_slice(&element);
        }
        bytes.extend_from_slice(format!("</{}>", q("lvl")).as_bytes());
        Ok(bytes)
    }
}

/// An abstract definition to add.
#[derive(Debug, Clone, PartialEq)]
pub struct NewAbstractNum {
    id: Option<u32>,
    multi_level_type: String,
    name: Option<String>,
    style_link: Option<String>,
    num_style_link: Option<String>,
    levels: Vec<NewLevel>,
}

impl NewAbstractNum {
    /// A `hybridMultilevel` definition with OfficeCLI's nine template
    /// levels ([`NewLevel::template`]).
    #[must_use]
    pub fn new(bullet: bool) -> Self {
        Self {
            id: None,
            multi_level_type: "hybridMultilevel".to_owned(),
            name: None,
            style_link: None,
            num_style_link: None,
            levels: (0..9)
                .map(|ilvl| NewLevel::template(ilvl, bullet))
                .collect(),
        }
    }

    /// Use identifier `id` instead of the next free one.
    #[must_use]
    pub const fn with_id(mut self, id: u32) -> Self {
        self.id = Some(id);
        self
    }

    /// Use `w:multiLevelType` `value`.
    #[must_use]
    pub fn with_multi_level_type(mut self, value: impl Into<String>) -> Self {
        self.multi_level_type = value.into();
        self
    }

    /// Name the definition (`w:name`).
    #[must_use]
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// Link a numbering style that this definition defines (`w:styleLink`).
    #[must_use]
    pub fn with_style_link(mut self, link: impl Into<String>) -> Self {
        self.style_link = Some(link.into());
        self
    }

    /// Defer to a numbering style's definition (`w:numStyleLink`).
    #[must_use]
    pub fn with_num_style_link(mut self, link: impl Into<String>) -> Self {
        self.num_style_link = Some(link.into());
        self
    }

    /// Replace the level of `level`'s index.
    #[must_use]
    pub fn with_level(mut self, level: NewLevel) -> Self {
        match self
            .levels
            .iter_mut()
            .find(|other| other.ilvl == level.ilvl)
        {
            Some(slot) => *slot = level,
            None => {
                self.levels.push(level);
                self.levels.sort_by_key(|level| level.ilvl);
            },
        }
        self
    }

    /// The level of index `ilvl`, to adjust before [`Self::with_level`].
    #[must_use]
    pub fn level(&self, ilvl: u8) -> Option<&NewLevel> {
        self.levels.iter().find(|level| level.ilvl == ilvl)
    }

    fn element(&self, prefix: &[u8], id: u32, strict: bool) -> Result<Vec<u8>> {
        let q = |local: &str| qualified(prefix, local);
        let mut bytes =
            format!(r#"<{} {}="{id}">"#, q("abstractNum"), q("abstractNumId")).into_bytes();
        bytes.extend_from_slice(&val_element(
            prefix,
            "multiLevelType",
            Some(&self.multi_level_type),
        ));
        for (local, value) in [
            ("name", &self.name),
            ("styleLink", &self.style_link),
            ("numStyleLink", &self.num_style_link),
        ] {
            if let Some(value) = value {
                bytes.extend_from_slice(&val_element(prefix, local, Some(value)));
            }
        }
        for level in &self.levels {
            bytes.extend_from_slice(&level.element(prefix, strict)?);
        }
        bytes.extend_from_slice(format!("</{}>", q("abstractNum")).as_bytes());
        Ok(bytes)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    fn catalog(body: &str) -> Catalog {
        let xml = format!(r#"<w:numbering xmlns:w="{W}">{body}</w:numbering>"#);
        Catalog::from_xml(Arc::new(xml.into_bytes())).unwrap()
    }

    #[test]
    fn reads_definitions_levels_and_instances() {
        let catalog = catalog(
            r#"<w:abstractNum w:abstractNumId="3"><w:multiLevelType w:val="multilevel"/><w:name w:val="Outline"/><w:lvl w:ilvl="0"><w:start w:val="2"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1)"/><w:lvlJc w:val="center"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:b/></w:rPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="4"><w:numStyleLink w:val="OutlineList"/></w:abstractNum><w:abstractNum w:abstractNumId="5"><w:styleLink w:val="OutlineList"/><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="3"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num><w:num w:numId="2"><w:abstractNumId w:val="4"/></w:num>"#,
        );
        let definition = catalog.abstract_num(3).unwrap();
        assert_eq!(definition.name(), Some("Outline"));
        assert_eq!(definition.multi_level_type(), Some("multilevel"));
        let level = definition.level(0).unwrap();
        assert_eq!(level.start(), Some("2"));
        assert_eq!(level.justification(), Some("center"));
        let paragraph = catalog.level_paragraph_properties(level).unwrap();
        assert_eq!(
            paragraph.indentation.unwrap().hanging,
            Some(properties::Twips(360))
        );
        assert_eq!(
            catalog.level_run_properties(level).unwrap().bold,
            Some(true)
        );
        assert_eq!(
            catalog.num(1).unwrap().overrides(),
            &[(0, Some("5".to_owned()))]
        );
        // numStyleLink is followed to the style's own definition.
        assert_eq!(catalog.format_of(2, 0), Some("bullet"));
        assert_eq!(catalog.format_of(1, 0), Some("upperRoman"));
        assert_eq!(catalog.next_num_id(), 3);
        assert_eq!(catalog.next_abstract_num_id(), 6);
    }

    #[test]
    fn edits_levels_definitions_and_instances_in_place() {
        let catalog = catalog(
            r#"<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:numIdMacAtCleanup w:val="0"/>"#,
        );
        let mut paragraph = properties::Paragraph::default();
        paragraph.indentation = Some(properties::Indentation {
            left: Some(properties::Twips(1440)),
            ..Default::default()
        });
        let changed = catalog
            .with_level_changed(
                0,
                0,
                &LevelChange::default()
                    .with_format(Some("lowerLetter".into()))
                    .with_justification(Some("right".into()))
                    .with_paragraph(paragraph),
            )
            .unwrap()
            .unwrap();
        let changed = Catalog::from_xml(Arc::new(changed)).unwrap();
        let level = changed.abstract_num(0).unwrap().level(0).unwrap();
        assert_eq!(level.format(), Some("lowerLetter"));
        assert_eq!(level.justification(), Some("right"));
        assert_eq!(
            changed
                .level_paragraph_properties(level)
                .unwrap()
                .indentation
                .unwrap()
                .left,
            Some(properties::Twips(1440))
        );

        let (added, id) = changed
            .with_abstract_num_added(&NewAbstractNum::new(true).with_name("Bullets"))
            .unwrap();
        assert_eq!(id, 1);
        let added = Catalog::from_xml(Arc::new(added)).unwrap();
        assert_eq!(added.abstract_num(1).unwrap().levels().len(), 9);
        assert_eq!(
            added.abstract_num(1).unwrap().level(1).unwrap().text(),
            Some("\u{25E6}")
        );
        let text = String::from_utf8(added.xml().to_vec()).unwrap();
        assert!(text.find("abstractNumId=\"1\"").unwrap() < text.find("<w:num ").unwrap());

        let (with_num, num_id) = added
            .with_num_added(&NewNum::new(1).with_start_override(0, 4))
            .unwrap();
        assert_eq!(num_id, 2);
        let with_num = Catalog::from_xml(Arc::new(with_num)).unwrap();
        assert_eq!(
            with_num.num(2).unwrap().overrides(),
            &[(0, Some("4".to_owned()))]
        );
        let text = String::from_utf8(with_num.xml().to_vec()).unwrap();
        assert!(text.find("w:numId=\"2\"").unwrap() < text.find("numIdMacAtCleanup").unwrap());

        let retargeted = with_num.with_num_retargeted(2, 0).unwrap().unwrap();
        let retargeted = Catalog::from_xml(Arc::new(retargeted)).unwrap();
        assert_eq!(retargeted.num(2).unwrap().abstract_num_id(), Some(0));
        // Removing a definition removes its instances.
        let cascaded =
            Catalog::from_xml(Arc::new(retargeted.with_abstract_num_removed(0).unwrap())).unwrap();
        assert!(cascaded.abstract_num(0).is_none());
        assert!(cascaded.num(1).is_none() && cascaded.num(2).is_none());
        let removed =
            Catalog::from_xml(Arc::new(retargeted.with_abstract_num_removed(1).unwrap())).unwrap();
        assert!(removed.abstract_num(1).is_none());
        let removed = Catalog::from_xml(Arc::new(removed.with_num_removed(2).unwrap())).unwrap();
        assert!(removed.num(2).is_none());
        let level_removed =
            Catalog::from_xml(Arc::new(removed.with_level_removed(0, 0).unwrap())).unwrap();
        assert!(level_removed.abstract_num(0).unwrap().levels().is_empty());
        let level_added = level_removed
            .with_level_added(0, &NewLevel::new(2, "decimal", "%3."))
            .unwrap();
        let level_added = Catalog::from_xml(Arc::new(level_added)).unwrap();
        assert_eq!(
            level_added
                .abstract_num(0)
                .unwrap()
                .level(2)
                .unwrap()
                .text(),
            Some("%3.")
        );
        let renamed = level_added
            .with_abstract_num_changed(
                0,
                &AbstractNumChange::default().with_name(Some("Renamed".into())),
            )
            .unwrap()
            .unwrap();
        let renamed = Catalog::from_xml(Arc::new(renamed)).unwrap();
        assert_eq!(renamed.abstract_num(0).unwrap().name(), Some("Renamed"));
    }
}
