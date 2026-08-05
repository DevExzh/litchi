//! Typed WordprocessingML tracked-change authoring.

use super::content_control::MutableContentControl;
use super::field::MutableField;
use super::hyperlink::MutableHyperlink;
use super::paragraph::{MutableParagraph, ParagraphElement, ParagraphProperties};
use super::run::{MutableRun, RunProperties};
use crate::error::{Error, Result};
use chrono::DateTime;
use litchi_core::xml::escape_xml;
use std::fmt::Write as _;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RevisionKind {
    Insert,
    Delete,
    MoveFrom,
    MoveTo,
}

/// Whole-table insertion or deletion marker written in table properties.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableRevisionKind {
    Insert,
    Delete,
}

impl TableRevisionKind {
    pub(crate) fn element(self) -> &'static str {
        match self {
            Self::Insert => "tblIns",
            Self::Delete => "tblDel",
        }
    }
}

/// Table-row insertion or deletion marker written in row properties.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RowRevisionKind {
    Insert,
    Delete,
}

impl RowRevisionKind {
    pub(crate) fn element(self) -> &'static str {
        match self {
            Self::Insert => "ins",
            Self::Delete => "del",
        }
    }
}

/// Table-cell insertion or deletion marker written in cell properties.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellRevisionKind {
    Insert,
    Delete,
}

impl CellRevisionKind {
    pub(crate) fn element(self) -> &'static str {
        match self {
            Self::Insert => "cellIns",
            Self::Delete => "cellDel",
        }
    }
}

/// Vertical-merge state used specifically by `w:cellMerge` revisions.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableCellMergeRevisionState {
    Rest,
    Continue,
}

impl TableCellMergeRevisionState {
    pub(crate) fn token(self) -> &'static str {
        match self {
            Self::Rest => "rest",
            Self::Continue => "cont",
        }
    }
}

impl RevisionKind {
    fn element(self) -> &'static str {
        match self {
            Self::Insert => "ins",
            Self::Delete => "del",
            Self::MoveFrom => "moveFrom",
            Self::MoveTo => "moveTo",
        }
    }
    pub(crate) fn text_mode(self) -> RevisionTextMode {
        match self {
            Self::Delete | Self::MoveFrom => RevisionTextMode::Deleted,
            _ => RevisionTextMode::Normal,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum RevisionTextMode {
    Normal,
    Deleted,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RevisionMetadata {
    id: u32,
    author: String,
    date: Option<String>,
    user_id: Option<String>,
}

impl RevisionMetadata {
    pub fn new(id: impl AsRef<str>, author: impl Into<String>) -> Result<Self> {
        let id = id.as_ref();
        if id.is_empty() || !id.bytes().all(|b| b.is_ascii_digit()) {
            return Err(Error::InvalidFormat(
                "revision ID must be a nonnegative decimal integer".into(),
            ));
        }
        let id = id.parse::<u32>().map_err(|_| {
            Error::InvalidFormat("revision ID exceeds the supported u32 range".into())
        })?;
        let author = author.into();
        validate_nonempty_xml("revision author", &author)?;
        Ok(Self {
            id,
            author,
            date: None,
            user_id: None,
        })
    }
    pub fn id(&self) -> u32 {
        self.id
    }
    pub fn author(&self) -> &str {
        &self.author
    }
    pub fn date(&self) -> Option<&str> {
        self.date.as_deref()
    }
    pub fn user_id(&self) -> Option<&str> {
        self.user_id.as_deref()
    }
    pub fn set_date(&mut self, date: Option<impl Into<String>>) -> Result<&mut Self> {
        let candidate = date.map(Into::into);
        if let Some(value) = &candidate {
            validate_nonempty_xml("revision date", value)?;
            DateTime::parse_from_rfc3339(value).map_err(|_| {
                Error::InvalidFormat("revision date must be valid W3CDTF/RFC 3339".into())
            })?;
        }
        self.date = candidate;
        Ok(self)
    }
    pub fn set_user_id(&mut self, user_id: Option<impl Into<String>>) -> Result<&mut Self> {
        let candidate = user_id.map(Into::into);
        if let Some(value) = &candidate {
            validate_nonempty_xml("revision user ID", value)?;
        }
        self.user_id = candidate;
        Ok(self)
    }
    pub(crate) fn write_attributes(&self, xml: &mut String) -> Result<()> {
        write!(
            xml,
            " w:id=\"{}\" w:author=\"{}\"",
            self.id,
            escape_xml(&self.author)
        )?;
        if let Some(date) = &self.date {
            write!(xml, " w:date=\"{}\"", escape_xml(date))?;
        }
        if let Some(user_id) = &self.user_id {
            write!(xml, " w:userId=\"{}\"", escape_xml(user_id))?;
        }
        Ok(())
    }
}

fn validate_nonempty_xml(label: &str, value: &str) -> Result<()> {
    if value.trim().is_empty() {
        return Err(Error::InvalidFormat(format!("{label} must not be empty")));
    }
    if value.chars().any(|ch| !matches!(ch, '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')) {
        return Err(Error::InvalidFormat(format!("{label} contains a character forbidden by XML 1.0")));
    }
    Ok(())
}

#[derive(Debug)]
enum RevisionChild {
    Element(ParagraphElement),
    CommentRangeStart(u32),
    CommentRangeEnd(u32),
    CommentReference(u32),
    ContentControl(RevisionContentControl),
}

#[derive(Debug)]
pub struct RevisionContentControl {
    control: MutableContentControl,
    elements: Vec<ParagraphElement>,
}

impl RevisionContentControl {
    pub fn add_run(&mut self) -> &mut MutableRun {
        self.elements.push(ParagraphElement::Run(MutableRun::new()));
        match self.elements.last_mut() {
            Some(ParagraphElement::Run(run)) => run,
            _ => unreachable!(),
        }
    }
    pub fn add_run_with_text(&mut self, text: &str) -> &mut MutableRun {
        let run = self.add_run();
        run.set_text(text);
        run
    }
    pub fn add_field(&mut self, field: MutableField) -> &mut Self {
        self.elements.push(ParagraphElement::Field(field));
        self
    }
    fn write_placeholder(
        &self,
        xml: &mut String,
        mode: RevisionTextMode,
        hi: &mut usize,
        ii: &mut usize,
    ) -> Result<()> {
        xml.push_str(&self.control.to_xml_start()?);
        for e in &self.elements {
            e.write_placeholder_mode(xml, hi, ii, mode)?;
        }
        xml.push_str(MutableContentControl::to_xml_end());
        Ok(())
    }
    fn write_with_rels(
        &self,
        xml: &mut String,
        mode: RevisionTextMode,
        mapper: &super::relmap::RelationshipMapper,
        hi: &mut usize,
        ii: &mut usize,
    ) -> Result<()> {
        xml.push_str(&self.control.to_xml_start()?);
        for e in &self.elements {
            e.write_with_rels_mode(xml, mapper, hi, ii, mode)?;
        }
        xml.push_str(MutableContentControl::to_xml_end());
        Ok(())
    }
}

#[derive(Debug)]
pub struct MutableRevision {
    kind: RevisionKind,
    metadata: RevisionMetadata,
    children: Vec<RevisionChild>,
}

impl MutableRevision {
    pub fn new(kind: RevisionKind, metadata: RevisionMetadata) -> Self {
        Self {
            kind,
            metadata,
            children: Vec::new(),
        }
    }
    pub fn kind(&self) -> RevisionKind {
        self.kind
    }
    pub fn metadata(&self) -> &RevisionMetadata {
        &self.metadata
    }
    pub fn child_count(&self) -> usize {
        self.children.len()
    }
    pub fn add_run(&mut self) -> &mut MutableRun {
        self.children
            .push(RevisionChild::Element(ParagraphElement::Run(
                MutableRun::new(),
            )));
        match self.children.last_mut() {
            Some(RevisionChild::Element(ParagraphElement::Run(run))) => run,
            _ => unreachable!(),
        }
    }
    pub fn add_run_with_text(&mut self, text: &str) -> &mut MutableRun {
        let run = self.add_run();
        run.set_text(text);
        run
    }
    pub fn add_hyperlink(&mut self, url: &str, text: &str) -> &mut MutableHyperlink {
        self.children
            .push(RevisionChild::Element(ParagraphElement::Hyperlink(
                MutableHyperlink::new(url, text),
            )));
        match self.children.last_mut() {
            Some(RevisionChild::Element(ParagraphElement::Hyperlink(link))) => link,
            _ => unreachable!(),
        }
    }
    pub fn add_field(&mut self, field: MutableField) -> &mut Self {
        self.children
            .push(RevisionChild::Element(ParagraphElement::Field(field)));
        self
    }
    pub fn add_bookmark_start(&mut self, id: u32, name: &str) -> &mut Self {
        self.children
            .push(RevisionChild::Element(ParagraphElement::BookmarkStart(
                super::bookmark::MutableBookmark::new(id, name.to_owned()),
            )));
        self
    }
    pub fn add_bookmark_end(&mut self, id: u32) -> &mut Self {
        self.children
            .push(RevisionChild::Element(ParagraphElement::BookmarkEnd(id)));
        self
    }
    pub fn add_comment_range_start(&mut self, id: u32) -> &mut Self {
        self.children.push(RevisionChild::CommentRangeStart(id));
        self
    }
    pub fn add_comment_range_end(&mut self, id: u32) -> &mut Self {
        self.children.push(RevisionChild::CommentRangeEnd(id));
        self
    }
    pub fn add_comment_reference(&mut self, id: u32) -> &mut Self {
        self.children.push(RevisionChild::CommentReference(id));
        self
    }
    pub fn add_content_control(
        &mut self,
        control: MutableContentControl,
    ) -> &mut RevisionContentControl {
        self.children
            .push(RevisionChild::ContentControl(RevisionContentControl {
                control,
                elements: Vec::new(),
            }));
        match self.children.last_mut() {
            Some(RevisionChild::ContentControl(control)) => control,
            _ => unreachable!(),
        }
    }
    pub fn add_revision(&mut self, _nested: MutableRevision) -> Result<&mut Self> {
        Err(Error::InvalidFormat(
            "nested WordprocessingML revision wrappers are not supported".into(),
        ))
    }
    fn write_start(&self, xml: &mut String) -> Result<()> {
        write!(xml, "<w:{}", self.kind.element())?;
        self.metadata.write_attributes(xml)?;
        xml.push('>');
        Ok(())
    }
    pub(crate) fn write_placeholder(
        &self,
        xml: &mut String,
        hi: &mut usize,
        ii: &mut usize,
    ) -> Result<()> {
        self.write_start(xml)?;
        let mode = self.kind.text_mode();
        for c in &self.children {
            c.write_placeholder(xml, mode, hi, ii)?;
        }
        write!(xml, "</w:{}>", self.kind.element())?;
        Ok(())
    }
    pub(crate) fn write_with_rels(
        &self,
        xml: &mut String,
        mapper: &super::relmap::RelationshipMapper,
        hi: &mut usize,
        ii: &mut usize,
    ) -> Result<()> {
        self.write_start(xml)?;
        let mode = self.kind.text_mode();
        for c in &self.children {
            c.write_with_rels(xml, mode, mapper, hi, ii)?;
        }
        write!(xml, "</w:{}>", self.kind.element())?;
        Ok(())
    }
    pub(crate) fn collect_hyperlink_urls(&self, urls: &mut Vec<String>) {
        for c in &self.children {
            c.collect_hyperlink_urls(urls);
        }
    }
    pub(crate) fn collect_images<'a>(
        &'a self,
        images: &mut Vec<(&'a [u8], crate::format::ImageFormat)>,
    ) {
        for c in &self.children {
            c.collect_images(images);
        }
    }
    pub(crate) fn append_run_text(&self, text: &mut String) {
        for c in &self.children {
            c.append_run_text(text);
        }
    }
}

impl RevisionChild {
    fn write_placeholder(
        &self,
        xml: &mut String,
        mode: RevisionTextMode,
        hi: &mut usize,
        ii: &mut usize,
    ) -> Result<()> {
        match self {
            Self::Element(e) => e.write_placeholder_mode(xml, hi, ii, mode),
            Self::CommentRangeStart(id) => {
                write!(xml, "<w:commentRangeStart w:id=\"{id}\"/>")?;
                Ok(())
            },
            Self::CommentRangeEnd(id) => {
                write!(xml, "<w:commentRangeEnd w:id=\"{id}\"/>")?;
                Ok(())
            },
            Self::CommentReference(id) => {
                write!(xml, "<w:r><w:commentReference w:id=\"{id}\"/></w:r>")?;
                Ok(())
            },
            Self::ContentControl(c) => c.write_placeholder(xml, mode, hi, ii),
        }
    }
    fn write_with_rels(
        &self,
        xml: &mut String,
        mode: RevisionTextMode,
        mapper: &super::relmap::RelationshipMapper,
        hi: &mut usize,
        ii: &mut usize,
    ) -> Result<()> {
        match self {
            Self::Element(e) => e.write_with_rels_mode(xml, mapper, hi, ii, mode),
            Self::CommentRangeStart(id) => {
                write!(xml, "<w:commentRangeStart w:id=\"{id}\"/>")?;
                Ok(())
            },
            Self::CommentRangeEnd(id) => {
                write!(xml, "<w:commentRangeEnd w:id=\"{id}\"/>")?;
                Ok(())
            },
            Self::CommentReference(id) => {
                write!(xml, "<w:r><w:commentReference w:id=\"{id}\"/></w:r>")?;
                Ok(())
            },
            Self::ContentControl(c) => c.write_with_rels(xml, mode, mapper, hi, ii),
        }
    }
    fn collect_hyperlink_urls(&self, urls: &mut Vec<String>) {
        match self {
            Self::Element(e) => e.collect_hyperlink_urls(urls),
            Self::ContentControl(c) => c
                .elements
                .iter()
                .for_each(|e| e.collect_hyperlink_urls(urls)),
            _ => {},
        }
    }
    fn collect_images<'a>(&'a self, images: &mut Vec<(&'a [u8], crate::format::ImageFormat)>) {
        match self {
            Self::Element(e) => e.collect_images(images),
            Self::ContentControl(c) => c.elements.iter().for_each(|e| e.collect_images(images)),
            _ => {},
        }
    }
    fn append_run_text(&self, text: &mut String) {
        match self {
            Self::Element(e) => e.append_run_text(text),
            Self::ContentControl(c) => c.elements.iter().for_each(|e| e.append_run_text(text)),
            _ => {},
        }
    }
}

#[derive(Debug, Clone)]
pub(crate) struct RunPropertyChange {
    metadata: RevisionMetadata,
    previous: RunProperties,
}
impl RunPropertyChange {
    pub(crate) fn snapshot(metadata: RevisionMetadata, previous: &MutableRun) -> Self {
        Self {
            metadata,
            previous: previous.properties.clone(),
        }
    }
    pub(crate) fn write_xml(&self, xml: &mut String) -> Result<()> {
        xml.push_str("<w:rPrChange");
        self.metadata.write_attributes(xml)?;
        xml.push_str("><w:rPr>");
        self.previous.write_values(xml)?;
        xml.push_str("</w:rPr></w:rPrChange>");
        Ok(())
    }
}

#[derive(Debug, Clone)]
pub(crate) struct ParagraphPropertyChange {
    metadata: RevisionMetadata,
    style: Option<String>,
    previous: ParagraphProperties,
}
impl ParagraphPropertyChange {
    pub(crate) fn snapshot(metadata: RevisionMetadata, previous: &MutableParagraph) -> Self {
        Self {
            metadata,
            style: previous.style.clone(),
            previous: previous.properties.clone(),
        }
    }
    pub(crate) fn write_xml(
        &self,
        xml: &mut String,
        mapper: Option<&super::relmap::RelationshipMapper>,
    ) -> Result<()> {
        xml.push_str("<w:pPrChange");
        self.metadata.write_attributes(xml)?;
        xml.push_str("><w:pPr>");
        self.previous
            .write_values(xml, self.style.as_deref(), mapper)?;
        xml.push_str("</w:pPr></w:pPrChange>");
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::paragraph::Paragraph;
    use crate::revision::RevisionType;
    fn md(id: &str) -> RevisionMetadata {
        let mut m = RevisionMetadata::new(id, "A & B").unwrap();
        m.set_date(Some("2026-07-19T10:30:00+08:00")).unwrap();
        m.set_user_id(Some("u<&\"1")).unwrap();
        m
    }
    #[test]
    fn wrappers_escape_delete_and_round_trip() {
        let mut p = MutableParagraph::new();
        for (i, k) in [
            RevisionKind::Insert,
            RevisionKind::Delete,
            RevisionKind::MoveFrom,
            RevisionKind::MoveTo,
        ]
        .into_iter()
        .enumerate()
        {
            p.add_revision(k, md(&i.to_string()))
                .add_run_with_text("x<&");
        }
        let mut x = String::new();
        p.to_xml(&mut x).unwrap();
        assert_eq!(x.matches("<w:delText").count(), 2);
        assert!(x.contains("A &amp; B"));
        assert!(x.contains("u&lt;&amp;&quot;1"));
        let r = Paragraph::new(x.into_bytes()).revisions().unwrap();
        assert_eq!(r.len(), 4);
        assert_eq!(r[0].revision_type(), RevisionType::Insert);
        assert_eq!(r[1].revision_type(), RevisionType::Delete);
    }
    #[test]
    fn validates_metadata_and_rolls_back() {
        assert!(RevisionMetadata::new("-1", "a").is_err());
        assert!(RevisionMetadata::new("4294967296", "a").is_err());
        assert!(RevisionMetadata::new("1", " ").is_err());
        let mut m = RevisionMetadata::new("1", "a").unwrap();
        assert!(m.set_date(Some("bad")).is_err());
        assert_eq!(m.date(), None);
        assert!(m.set_user_id(Some("\u{1}")).is_err());
        assert_eq!(m.user_id(), None);
    }
    #[test]
    fn deleted_fields_use_deleted_elements() {
        let mut p = MutableParagraph::new();
        p.add_revision(RevisionKind::Delete, md("1"))
            .add_field(MutableField::with_result("REF A".into(), "old".into()));
        let mut x = String::new();
        p.to_xml(&mut x).unwrap();
        assert!(x.contains("<w:delInstrText>REF A</w:delInstrText>"));
        assert!(x.contains("<w:delText>old</w:delText>"));
    }
    #[test]
    fn property_changes_are_last() {
        let mut oldr = MutableRun::new();
        oldr.bold(true).font_name("Old");
        let mut p = MutableParagraph::new();
        p.add_run_with_text("n")
            .italic(true)
            .set_property_change(md("2"), &oldr);
        let mut oldp = MutableParagraph::new();
        oldp.set_alignment(crate::ParagraphAlignment::Center);
        p.set_property_change(md("3"), &oldp);
        let mut x = String::new();
        p.to_xml(&mut x).unwrap();
        assert!(x.contains("<w:rPr><w:i/><w:rPrChange"));
        assert!(x.contains("<w:pPrChange"));
        assert!(x.contains("<w:pPr><w:jc w:val=\"center\"/></w:pPr></w:pPrChange></w:pPr>"));
    }
    #[test]
    fn rejects_nested_atomically() {
        let mut a = MutableRevision::new(RevisionKind::Insert, md("1"));
        a.add_run_with_text("x");
        let n = a.child_count();
        assert!(
            a.add_revision(MutableRevision::new(RevisionKind::Insert, md("2")))
                .is_err()
        );
        assert_eq!(a.child_count(), n);
    }
    #[test]
    fn legal_anchors_controls_and_strict_context() {
        let mut p = MutableParagraph::new();
        let r = p.add_revision(RevisionKind::Insert, md("1"));
        r.add_bookmark_start(4, "b").add_comment_range_start(7);
        r.add_content_control(MutableContentControl::plain_text(9, Some("t")))
            .add_run_with_text("v");
        r.add_comment_range_end(7)
            .add_comment_reference(7)
            .add_bookmark_end(4);
        let mut x = String::new();
        p.to_xml(&mut x).unwrap();
        let strict = format!(
            "<w:document xmlns:w=\"http://purl.oclc.org/ooxml/wordprocessingml/main\">{x}</w:document>"
        );
        assert!(strict.contains("<w:ins"));
        assert!(strict.contains("<w:sdt>"));
        assert_eq!(Paragraph::new(x.into_bytes()).revisions().unwrap().len(), 1);
    }
}
