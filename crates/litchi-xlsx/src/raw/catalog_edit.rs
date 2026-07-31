//! Narrow, lossless surgery for workbook sheet-tab properties.
//!
//! Only directly modeled `sheet/@state` and the first workbook view's
//! `activeTab` are regenerated. All untouched workbook bytes remain exact.

use litchi_core::xml::escape_xml;
use litchi_ooxml_common::xml::unqualified_attribute_value;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use crate::error::{Error, Result, TabEditBlock, invalid};
use crate::raw::namespace::{is_spreadsheetml_name, relationship_attribute_value};

const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const MAX_ACTIVE_TAB: usize = 32_766;

/// Recognized sheet states that are safe to author.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum State {
    Visible,
    Hidden,
    VeryHidden,
}

impl State {
    const fn attribute(self) -> Option<&'static str> {
        match self {
            Self::Visible => None,
            Self::Hidden => Some("hidden"),
            Self::VeryHidden => Some("veryHidden"),
        }
    }
}

/// One borrowed semantic tab change. Physical relationship IDs never escape
/// this low-level boundary.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Tab<'a> {
    pub(crate) sheet: &'a str,
    pub(crate) position: usize,
    pub(crate) relationship_id: &'a str,
    pub(crate) state: State,
}

/// Move-only workbook rewrite plan.
#[derive(Debug)]
pub(crate) struct Plan<'a> {
    pub(crate) tabs: Vec<Tab<'a>>,
    /// A replacement for the first workbook view's active tab. `None` leaves
    /// every workbook view byte-exact.
    pub(crate) active: Option<usize>,
}

#[derive(Debug, Clone, Copy)]
struct Span {
    start: usize,
    end: usize,
}

#[derive(Debug, Clone)]
struct Attribute {
    name: Box<str>,
    value: Box<str>,
}

#[derive(Debug, Clone)]
struct Tag {
    name: Box<str>,
    attributes: Box<[Attribute]>,
}

#[derive(Debug)]
struct Slot {
    span: Span,
    tag_end: usize,
    close_start: usize,
    tag: Tag,
    empty: bool,
}

#[derive(Debug)]
struct SheetSlot {
    relationship_id: Box<str>,
    slot: Slot,
}

#[derive(Debug)]
struct Container {
    slot: Slot,
    payload: bool,
}

#[derive(Debug)]
struct Layout {
    sheets: Container,
    sheet_slots: Box<[SheetSlot]>,
    book_views: Option<Container>,
    workbook_view: Option<Slot>,
    protected: bool,
    alternate_content: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Kind {
    Workbook,
    Sheets,
    Sheet,
    BookViews,
    WorkbookView,
    Other,
}

#[derive(Debug, Clone, Copy)]
struct Frame {
    kind: Kind,
}

#[derive(Debug)]
struct Pending {
    start: usize,
    tag_end: usize,
    tag: Tag,
}

#[derive(Debug, Default)]
struct Scanner {
    root_seen: bool,
    sheets: Option<Container>,
    pending_sheets: Option<Pending>,
    sheet_slots: Vec<SheetSlot>,
    pending_sheet: Option<(Pending, Box<str>)>,
    book_views: Option<Container>,
    pending_book_views: Option<Pending>,
    workbook_view: Option<Slot>,
    pending_workbook_view: Option<Pending>,
    book_views_payload: bool,
    protected: bool,
    alternate_content: bool,
}

#[derive(Debug)]
struct Replacement {
    span: Span,
    bytes: Vec<u8>,
}

/// Rewrite recognized tab states and, when requested, the active workbook
/// view. The caller reparses and verifies the semantic result before publish.
pub(crate) fn rewrite(content: &[u8], plan: Plan<'_>) -> Result<Vec<u8>> {
    if plan.tabs.is_empty() && plan.active.is_none() {
        return Ok(content.to_vec());
    }
    let layout = scan(content)?;
    let first = plan.tabs.first().copied();
    if layout.protected {
        return Err(block(first, TabEditBlock::ProtectedWorkbook));
    }
    if plan.active.is_some_and(|active| active > MAX_ACTIVE_TAB) {
        return Err(invalid(format!(
            "workbook active tab exceeds the Office limit {MAX_ACTIVE_TAB}"
        )));
    }

    let mut replacements = Vec::new();
    replacements
        .try_reserve(plan.tabs.len().saturating_add(1))
        .map_err(|error| invalid(format!("cannot reserve workbook edit plan: {error}")))?;
    for requested in plan.tabs {
        let mut matches = layout
            .sheet_slots
            .iter()
            .filter(|slot| slot.relationship_id.as_ref() == requested.relationship_id);
        let Some(found) = matches.next() else {
            return Err(Error::TabEditBlocked {
                sheet: requested.sheet.to_owned(),
                position: requested.position,
                reason: TabEditBlock::MarkupCompatibility,
            });
        };
        if matches.next().is_some() {
            return Err(invalid(format!(
                "duplicate direct workbook sheet relationship '{}' during edit",
                requested.relationship_id
            )));
        }
        let appended = requested
            .state
            .attribute()
            .map(|value| ("state", value.to_owned()))
            .into_iter()
            .collect::<Vec<_>>();
        replacements.push(Replacement {
            span: found.slot.span,
            bytes: rewrite_slot(content, &found.slot, &["state"], &appended),
        });
    }

    if let Some(active) = plan.active {
        replacements.push(active_replacement(content, &layout, active, first)?);
    }
    replacements.sort_unstable_by_key(|replacement| replacement.span.start);
    if replacements
        .windows(2)
        .any(|pair| pair[0].span.end > pair[1].span.start)
    {
        return Err(invalid("overlapping workbook edit replacements"));
    }

    let output_len = replacements
        .iter()
        .try_fold(content.len(), |size, replacement| {
            let removed = replacement.span.end.checked_sub(replacement.span.start)?;
            size.checked_sub(removed)?
                .checked_add(replacement.bytes.len())
        })
        .ok_or_else(|| invalid("workbook edit output size overflow"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|error| invalid(format!("cannot reserve workbook edit output: {error}")))?;
    let mut cursor = 0usize;
    for replacement in replacements {
        output.extend_from_slice(&content[cursor..replacement.span.start]);
        output.extend_from_slice(&replacement.bytes);
        cursor = replacement.span.end;
    }
    output.extend_from_slice(&content[cursor..]);
    Ok(output)
}

fn block(tab: Option<Tab<'_>>, reason: TabEditBlock) -> Error {
    tab.map_or_else(
        || invalid("active-tab rewrite has no associated tab change"),
        |tab| Error::TabEditBlocked {
            sheet: tab.sheet.to_owned(),
            position: tab.position,
            reason,
        },
    )
}

fn active_replacement(
    source: &[u8],
    layout: &Layout,
    active: usize,
    tab: Option<Tab<'_>>,
) -> Result<Replacement> {
    let appended = [("activeTab", active.to_string())];
    if let Some(view) = &layout.workbook_view {
        return Ok(Replacement {
            span: view.span,
            bytes: rewrite_slot(source, view, &["activeTab"], &appended),
        });
    }
    if layout.alternate_content {
        return Err(block(tab, TabEditBlock::MarkupCompatibility));
    }
    if let Some(book_views) = &layout.book_views {
        if book_views.payload {
            return Err(block(tab, TabEditBlock::MarkupCompatibility));
        }
        let name = sibling_name(&book_views.slot.tag.name, "workbookView");
        let view = Tag {
            name: name.into_boxed_str(),
            attributes: Box::new([]),
        };
        let mut bytes = Vec::new();
        if book_views.slot.empty {
            write_tag(&mut bytes, &book_views.slot.tag, false, &[], &[]);
            write_tag(&mut bytes, &view, true, &[], &appended);
            write_close(&mut bytes, &book_views.slot.tag.name);
        } else {
            bytes.extend_from_slice(
                &source[book_views.slot.span.start..book_views.slot.close_start],
            );
            write_tag(&mut bytes, &view, true, &[], &appended);
            bytes.extend_from_slice(&source[book_views.slot.close_start..book_views.slot.span.end]);
        }
        return Ok(Replacement {
            span: book_views.slot.span,
            bytes,
        });
    }

    let book_views_name = sibling_name(&layout.sheets.slot.tag.name, "bookViews");
    let workbook_view_name = sibling_name(&book_views_name, "workbookView");
    let book_views = Tag {
        name: book_views_name.clone().into_boxed_str(),
        attributes: Box::new([]),
    };
    let workbook_view = Tag {
        name: workbook_view_name.into_boxed_str(),
        attributes: Box::new([]),
    };
    let mut bytes = Vec::new();
    write_tag(&mut bytes, &book_views, false, &[], &[]);
    write_tag(&mut bytes, &workbook_view, true, &[], &appended);
    write_close(&mut bytes, &book_views_name);
    Ok(Replacement {
        span: Span {
            start: layout.sheets.slot.span.start,
            end: layout.sheets.slot.span.start,
        },
        bytes,
    })
}

fn rewrite_slot(
    source: &[u8],
    slot: &Slot,
    removed: &[&str],
    appended: &[(&str, String)],
) -> Vec<u8> {
    let mut output = Vec::new();
    write_tag(&mut output, &slot.tag, slot.empty, removed, appended);
    if !slot.empty {
        output.extend_from_slice(&source[slot.tag_end..slot.span.end]);
    }
    output
}

fn scan(content: &[u8]) -> Result<Layout> {
    let mut reader = NsReader::from_reader(content);
    let mut scanner = Scanner::default();
    let mut stack = Vec::<Frame>::new();
    loop {
        let event_start = position(&reader)?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| invalid(error.to_string()))?
            .into_owned();
        let event_end = position(&reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                let parent = stack.last().map(|frame| frame.kind);
                let kind = scanner.start(
                    parent,
                    &namespace,
                    &element,
                    decoder,
                    &resolver,
                    event_start,
                    event_end,
                )?;
                stack.push(Frame { kind });
            },
            Event::Empty(element) => scanner.empty(
                stack.last().map(|frame| frame.kind),
                &namespace,
                &element,
                decoder,
                &resolver,
                Span {
                    start: event_start,
                    end: event_end,
                },
            )?,
            Event::End(_) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("workbook edit scan has an unmatched closing tag"))?;
                scanner.finish(frame, event_start, event_end)?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !stack.is_empty() {
        return Err(invalid("workbook edit scan ended inside an element"));
    }
    scanner.finish_layout()
}

impl Scanner {
    #[allow(clippy::too_many_arguments)]
    fn start(
        &mut self,
        parent: Option<Kind>,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
        start: usize,
        end: usize,
    ) -> Result<Kind> {
        if parent.is_none() {
            if self.root_seen || !is_spreadsheetml_name(namespace, element.name(), b"workbook") {
                return Err(invalid("workbook edit requires one SpreadsheetML root"));
            }
            self.root_seen = true;
            return Ok(Kind::Workbook);
        }
        self.observe_guard(parent, namespace, element, decoder)?;
        if parent == Some(Kind::Workbook)
            && is_spreadsheetml_name(namespace, element.name(), b"sheets")
        {
            if self.sheets.is_some() || self.pending_sheets.is_some() {
                return Err(invalid("duplicate direct sheets element during edit"));
            }
            self.pending_sheets = Some(Pending {
                start,
                tag_end: end,
                tag: tag(element, decoder)?,
            });
            return Ok(Kind::Sheets);
        }
        if parent == Some(Kind::Sheets)
            && is_spreadsheetml_name(namespace, element.name(), b"sheet")
        {
            if self.pending_sheet.is_some() {
                return Err(invalid("nested direct sheet element during edit"));
            }
            let relationship_id = sheet_relationship(element, decoder, resolver)?;
            self.pending_sheet = Some((
                Pending {
                    start,
                    tag_end: end,
                    tag: tag(element, decoder)?,
                },
                relationship_id,
            ));
            return Ok(Kind::Sheet);
        }
        if parent == Some(Kind::Workbook)
            && is_spreadsheetml_name(namespace, element.name(), b"bookViews")
        {
            if self.book_views.is_some() || self.pending_book_views.is_some() {
                return Err(invalid("duplicate direct bookViews element during edit"));
            }
            self.pending_book_views = Some(Pending {
                start,
                tag_end: end,
                tag: tag(element, decoder)?,
            });
            return Ok(Kind::BookViews);
        }
        if parent == Some(Kind::BookViews)
            && is_spreadsheetml_name(namespace, element.name(), b"workbookView")
        {
            if self.workbook_view.is_none() && self.pending_workbook_view.is_none() {
                self.pending_workbook_view = Some(Pending {
                    start,
                    tag_end: end,
                    tag: tag(element, decoder)?,
                });
                return Ok(Kind::WorkbookView);
            }
            return Ok(Kind::Other);
        }
        if parent == Some(Kind::BookViews) {
            self.book_views_payload = true;
        }
        Ok(Kind::Other)
    }

    #[allow(clippy::too_many_arguments)]
    fn empty(
        &mut self,
        parent: Option<Kind>,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
        span: Span,
    ) -> Result<()> {
        self.observe_guard(parent, namespace, element, decoder)?;
        if parent == Some(Kind::Workbook)
            && is_spreadsheetml_name(namespace, element.name(), b"sheets")
        {
            if self.sheets.is_some() || self.pending_sheets.is_some() {
                return Err(invalid("duplicate direct sheets element during edit"));
            }
            self.sheets = Some(Container {
                slot: Slot {
                    span,
                    tag_end: span.end,
                    close_start: span.end,
                    tag: tag(element, decoder)?,
                    empty: true,
                },
                payload: false,
            });
        } else if parent == Some(Kind::Sheets)
            && is_spreadsheetml_name(namespace, element.name(), b"sheet")
        {
            self.sheet_slots.push(SheetSlot {
                relationship_id: sheet_relationship(element, decoder, resolver)?,
                slot: Slot {
                    span,
                    tag_end: span.end,
                    close_start: span.end,
                    tag: tag(element, decoder)?,
                    empty: true,
                },
            });
        } else if parent == Some(Kind::Workbook)
            && is_spreadsheetml_name(namespace, element.name(), b"bookViews")
        {
            if self.book_views.is_some() || self.pending_book_views.is_some() {
                return Err(invalid("duplicate direct bookViews element during edit"));
            }
            self.book_views = Some(Container {
                slot: Slot {
                    span,
                    tag_end: span.end,
                    close_start: span.end,
                    tag: tag(element, decoder)?,
                    empty: true,
                },
                payload: false,
            });
        } else if parent == Some(Kind::BookViews)
            && is_spreadsheetml_name(namespace, element.name(), b"workbookView")
        {
            if self.workbook_view.is_none() && self.pending_workbook_view.is_none() {
                self.workbook_view = Some(Slot {
                    span,
                    tag_end: span.end,
                    close_start: span.end,
                    tag: tag(element, decoder)?,
                    empty: true,
                });
            }
        } else if parent == Some(Kind::BookViews) {
            self.book_views_payload = true;
        }
        Ok(())
    }

    fn observe_guard(
        &mut self,
        parent: Option<Kind>,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<()> {
        if is_spreadsheetml_name(namespace, element.name(), b"workbookProtection") {
            self.protected |= optional_bool(element, b"lockStructure", decoder)?.unwrap_or(false);
        }
        if matches!(parent, Some(Kind::Workbook | Kind::BookViews))
            && is_mce_name(namespace, element, b"AlternateContent")
        {
            self.alternate_content = true;
        }
        Ok(())
    }

    fn finish(&mut self, frame: Frame, close_start: usize, end: usize) -> Result<()> {
        match frame.kind {
            Kind::Sheet => {
                let (pending, relationship_id) = self
                    .pending_sheet
                    .take()
                    .ok_or_else(|| invalid("sheet close without workbook edit state"))?;
                self.sheet_slots.push(SheetSlot {
                    relationship_id,
                    slot: Slot {
                        span: Span {
                            start: pending.start,
                            end,
                        },
                        tag_end: pending.tag_end,
                        close_start,
                        tag: pending.tag,
                        empty: false,
                    },
                });
            },
            Kind::Sheets => {
                let pending = self
                    .pending_sheets
                    .take()
                    .ok_or_else(|| invalid("sheets close without workbook edit state"))?;
                self.sheets = Some(Container {
                    slot: Slot {
                        span: Span {
                            start: pending.start,
                            end,
                        },
                        tag_end: pending.tag_end,
                        close_start,
                        tag: pending.tag,
                        empty: false,
                    },
                    payload: false,
                });
            },
            Kind::WorkbookView => {
                let pending = self
                    .pending_workbook_view
                    .take()
                    .ok_or_else(|| invalid("workbookView close without edit state"))?;
                self.workbook_view = Some(Slot {
                    span: Span {
                        start: pending.start,
                        end,
                    },
                    tag_end: pending.tag_end,
                    close_start,
                    tag: pending.tag,
                    empty: false,
                });
            },
            Kind::BookViews => {
                let pending = self
                    .pending_book_views
                    .take()
                    .ok_or_else(|| invalid("bookViews close without workbook edit state"))?;
                self.book_views = Some(Container {
                    slot: Slot {
                        span: Span {
                            start: pending.start,
                            end,
                        },
                        tag_end: pending.tag_end,
                        close_start,
                        tag: pending.tag,
                        empty: false,
                    },
                    payload: self.book_views_payload,
                });
            },
            _ => {},
        }
        Ok(())
    }

    fn finish_layout(self) -> Result<Layout> {
        if !self.root_seen {
            return Err(invalid("workbook edit scan did not find a root"));
        }
        let sheets = self
            .sheets
            .ok_or_else(|| invalid("tab edits require a direct sheets element"))?;
        Ok(Layout {
            sheets,
            sheet_slots: self.sheet_slots.into_boxed_slice(),
            book_views: self.book_views,
            workbook_view: self.workbook_view,
            protected: self.protected,
            alternate_content: self.alternate_content,
        })
    }
}

fn sheet_relationship(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Box<str>> {
    relationship_attribute_value(element, b"id", decoder, resolver)?
        .filter(|value| !value.is_empty())
        .map(String::into_boxed_str)
        .ok_or_else(|| invalid("direct workbook sheet is missing a relationship ID during edit"))
}

fn optional_bool(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<Option<bool>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| match value.as_str() {
            "1" | "true" => Ok(true),
            "0" | "false" => Ok(false),
            _ => Err(invalid(format!(
                "invalid workbook protection boolean '{value}' during edit"
            ))),
        })
        .transpose()
}

fn write_tag(
    output: &mut Vec<u8>,
    tag: &Tag,
    empty: bool,
    removed: &[&str],
    appended: &[(&str, String)],
) {
    output.extend_from_slice(b"<");
    output.extend_from_slice(tag.name.as_bytes());
    for attribute in &tag.attributes {
        if removed.iter().any(|name| *name == attribute.name.as_ref()) {
            continue;
        }
        output.extend_from_slice(b" ");
        output.extend_from_slice(attribute.name.as_bytes());
        output.extend_from_slice(b"=\"");
        output.extend_from_slice(escape_xml(&attribute.value).as_bytes());
        output.extend_from_slice(b"\"");
    }
    for (name, value) in appended {
        output.extend_from_slice(b" ");
        output.extend_from_slice(name.as_bytes());
        output.extend_from_slice(b"=\"");
        output.extend_from_slice(escape_xml(value).as_bytes());
        output.extend_from_slice(b"\"");
    }
    if empty {
        output.extend_from_slice(b"/>");
    } else {
        output.extend_from_slice(b">");
    }
}

fn write_close(output: &mut Vec<u8>, name: &str) {
    output.extend_from_slice(b"</");
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b">");
}

fn tag(element: &BytesStart<'_>, decoder: Decoder) -> Result<Tag> {
    let name = std::str::from_utf8(element.name().as_ref())
        .map_err(|error| invalid(format!("workbook element name is not UTF-8: {error}")))?
        .to_owned();
    let mut attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| invalid(format!("workbook attribute name is not UTF-8: {error}")))?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| invalid(error.to_string()))?
            .into_owned();
        attributes.push(Attribute {
            name: name.into_boxed_str(),
            value: value.into_boxed_str(),
        });
    }
    Ok(Tag {
        name: name.into_boxed_str(),
        attributes: attributes.into_boxed_slice(),
    })
}

fn sibling_name(name: &str, local: &str) -> String {
    name.split_once(':').map_or_else(
        || local.to_owned(),
        |(prefix, _)| format!("{prefix}:{local}"),
    )
}

fn is_mce_name(namespace: &ResolveResult<'_>, element: &BytesStart<'_>, local: &[u8]) -> bool {
    element.name().local_name().as_ref() == local
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == MCE)
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid("workbook XML position does not fit usize"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::raw::{Visibility, parse_catalog};

    const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    fn plan<'a>(tabs: Vec<Tab<'a>>, active: Option<usize>) -> Plan<'a> {
        Plan { tabs, active }
    }

    #[test]
    fn rewrites_only_selected_states_and_first_view() {
        let source = format!(
            r#"<?xml version="1.0"?><x:workbook xmlns:x="{S}" xmlns:rel="{R}" xmlns:z="urn:future"><x:bookViews><x:workbookView activeTab="0" z:keep="yes"/><x:workbookView activeTab="0"/></x:bookViews><x:sheets z:container="exact"><x:sheet name="One" sheetId="1" rel:id="r1" z:keep="yes"/><x:sheet name="Two" sheetId="2" state="hidden" rel:id="r2"/></x:sheets><x:extLst><z:data value="exact"/></x:extLst></x:workbook>"#
        );
        let output = rewrite(
            source.as_bytes(),
            plan(
                vec![
                    Tab {
                        sheet: "One",
                        position: 0,
                        relationship_id: "r1",
                        state: State::Hidden,
                    },
                    Tab {
                        sheet: "Two",
                        position: 1,
                        relationship_id: "r2",
                        state: State::Visible,
                    },
                ],
                Some(1),
            ),
        )
        .expect("rewrite");
        let text = std::str::from_utf8(&output).expect("UTF-8");
        assert!(text.contains(
            r#"<x:sheet name="One" sheetId="1" rel:id="r1" z:keep="yes" state="hidden"/>"#
        ));
        assert!(text.contains(r#"<x:sheet name="Two" sheetId="2" rel:id="r2"/>"#));
        assert!(text.contains(r#"<x:workbookView z:keep="yes" activeTab="1"/>"#));
        assert!(text.contains(r#"<x:workbookView activeTab="0"/>"#));
        assert!(text.contains(r#"<x:extLst><z:data value="exact"/></x:extLst>"#));
        let catalog = parse_catalog(&output).expect("catalog");
        assert_eq!(catalog.active_sheet_index, 1);
        assert_eq!(catalog.sheets[0].visibility, Visibility::Hidden);
        assert_eq!(catalog.sheets[1].visibility, Visibility::Visible);
    }

    #[test]
    fn inserts_prefixed_book_views_before_sheets() {
        let source = format!(
            r#"<s:workbook xmlns:s="{S}" xmlns:r="{R}"><s:sheets><s:sheet name="One" sheetId="1" state="hidden" r:id="r1"/><s:sheet name="Two" sheetId="2" r:id="r2"/></s:sheets></s:workbook>"#
        );
        let output = rewrite(
            source.as_bytes(),
            plan(
                vec![Tab {
                    sheet: "One",
                    position: 0,
                    relationship_id: "r1",
                    state: State::Hidden,
                }],
                Some(1),
            ),
        )
        .expect("rewrite");
        let text = std::str::from_utf8(&output).expect("UTF-8");
        assert!(
            text.contains(
                r#"<s:bookViews><s:workbookView activeTab="1"/></s:bookViews><s:sheets>"#
            )
        );
        assert_eq!(
            parse_catalog(&output).expect("catalog").active_sheet_index,
            1
        );
    }

    #[test]
    fn expands_an_empty_book_views_container() {
        let source = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}"><bookViews data="kept"/><sheets><sheet name="One" sheetId="1" r:id="r1"/><sheet name="Two" sheetId="2" r:id="r2"/></sheets></workbook>"#
        );
        let output = rewrite(
            source.as_bytes(),
            plan(
                vec![Tab {
                    sheet: "One",
                    position: 0,
                    relationship_id: "r1",
                    state: State::Hidden,
                }],
                Some(1),
            ),
        )
        .expect("rewrite");
        assert!(
            std::str::from_utf8(&output)
                .expect("UTF-8")
                .contains(r#"<bookViews data="kept"><workbookView activeTab="1"/></bookViews>"#)
        );
    }

    #[test]
    fn blocks_structure_protection_but_preserves_unrelated_alternate_content() {
        let protected = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}"><workbookProtection lockStructure="1"/><sheets><sheet name="One" sheetId="1" r:id="r1"/></sheets></workbook>"#
        );
        let tab = Tab {
            sheet: "One",
            position: 0,
            relationship_id: "r1",
            state: State::Hidden,
        };
        assert!(matches!(
            rewrite(protected.as_bytes(), plan(vec![tab], None)),
            Err(Error::TabEditBlocked {
                reason: TabEditBlock::ProtectedWorkbook,
                ..
            })
        ));

        let mce = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}" xmlns:mc="{mce}"><mc:AlternateContent><mc:Fallback/></mc:AlternateContent><sheets><sheet name="One" sheetId="1" r:id="r1"/></sheets></workbook>"#,
            mce = String::from_utf8_lossy(MCE)
        );
        let rewritten =
            rewrite(mce.as_bytes(), plan(vec![tab], None)).expect("unrelated compatibility XML");
        let text = std::str::from_utf8(&rewritten).expect("UTF-8");
        assert!(text.contains("mc:AlternateContent"));
        assert!(text.contains(r#"state="hidden""#));

        let nested_protection = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}" xmlns:mc="{mce}"><mc:AlternateContent><mc:Fallback><workbookProtection lockStructure="1"/></mc:Fallback></mc:AlternateContent><sheets><sheet name="One" sheetId="1" r:id="r1"/></sheets></workbook>"#,
            mce = String::from_utf8_lossy(MCE)
        );
        assert!(matches!(
            rewrite(nested_protection.as_bytes(), plan(vec![tab], None)),
            Err(Error::TabEditBlocked {
                reason: TabEditBlock::ProtectedWorkbook,
                ..
            })
        ));
    }

    #[test]
    fn blocks_active_view_insertion_beside_alternate_content() {
        let source = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}" xmlns:mc="{mce}"><mc:AlternateContent><mc:Fallback/></mc:AlternateContent><sheets><sheet name="One" sheetId="1" state="hidden" r:id="r1"/><sheet name="Two" sheetId="2" r:id="r2"/></sheets></workbook>"#,
            mce = String::from_utf8_lossy(MCE)
        );
        assert!(matches!(
            rewrite(
                source.as_bytes(),
                plan(
                    vec![Tab {
                        sheet: "One",
                        position: 0,
                        relationship_id: "r1",
                        state: State::Hidden,
                    }],
                    Some(1),
                )
            ),
            Err(Error::TabEditBlocked {
                reason: TabEditBlock::MarkupCompatibility,
                ..
            })
        ));
    }

    #[test]
    fn blocks_an_effective_sheet_without_a_direct_slot() {
        let source = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="One" sheetId="1" r:id="r1"/></sheets></workbook>"#
        );
        assert!(matches!(
            rewrite(
                source.as_bytes(),
                plan(
                    vec![Tab {
                        sheet: "Fallback",
                        position: 1,
                        relationship_id: "mce-rel",
                        state: State::VeryHidden,
                    }],
                    None,
                )
            ),
            Err(Error::TabEditBlocked {
                reason: TabEditBlock::MarkupCompatibility,
                ..
            })
        ));
    }
}
