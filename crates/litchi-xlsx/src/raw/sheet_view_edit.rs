//! Lossless synchronization of a sheet's first workbook-view selection bit.
//!
//! `workbookView/@activeTab` and each sheet part's
//! `sheetView[@workbookViewId=0]/@tabSelected` describe one logical selection.
//! This module edits only that boolean and inserts the minimal view structure
//! when no direct view exists and schema order is unambiguous.

use litchi_core::xml::escape_xml;
use litchi_ooxml_common::xml::unqualified_attribute_value;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use crate::error::{Error, Result, TabEditBlock, allocation, invalid};
use crate::raw::namespace::is_spreadsheetml_name;

const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";

#[derive(Debug, Clone, Copy)]
pub(crate) struct Context<'a> {
    pub(crate) sheet: &'a str,
    pub(crate) position: usize,
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
struct Container {
    slot: Slot,
    payload: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RootKind {
    Tabular,
    Chart,
}

#[derive(Debug)]
struct Layout {
    root: Slot,
    views: Option<Container>,
    view: Option<Slot>,
    insertion_before: Option<usize>,
    alternate_content: bool,
    root_payload: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Kind {
    Root,
    Views,
    View,
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
    root_kind: Option<RootKind>,
    root: Option<Slot>,
    pending_root: Option<Pending>,
    views: Option<Container>,
    pending_views: Option<Pending>,
    view: Option<Slot>,
    pending_view: Option<Pending>,
    views_payload: bool,
    insertion_before: Option<usize>,
    alternate_content: bool,
    root_payload: bool,
}

/// Set or clear selection for workbook view zero. A missing false state is
/// already canonical and remains byte-exact.
pub(crate) fn rewrite(content: &[u8], selected: bool, context: Context<'_>) -> Result<Vec<u8>> {
    let layout = scan(content)?;
    let appended = selected
        .then(|| ("tabSelected", "1".to_owned()))
        .into_iter()
        .collect::<Vec<_>>();
    let output = if let Some(view) = &layout.view {
        rewrite_slot(content, view, &["tabSelected"], &appended)
    } else {
        if layout.alternate_content || layout.views.as_ref().is_some_and(|views| views.payload) {
            return Err(block(context));
        }
        if !selected {
            return Ok(content.to_vec());
        }
        if let Some(views) = &layout.views {
            insert_view(content, views)
        } else {
            if layout.root_payload {
                return Err(block(context));
            }
            insert_views(content, &layout)?
        }
    };
    verify(&output, selected, context)?;
    Ok(output)
}

fn block(context: Context<'_>) -> Error {
    Error::TabEditBlocked {
        sheet: context.sheet.to_owned(),
        position: context.position,
        reason: TabEditBlock::MarkupCompatibility,
    }
}

fn insert_view(source: &[u8], views: &Container) -> Vec<u8> {
    let view_name = sibling_name(&views.slot.tag.name, "sheetView");
    let view = Tag {
        name: view_name.into_boxed_str(),
        attributes: Box::new([]),
    };
    let appended = [
        ("tabSelected", "1".to_owned()),
        ("workbookViewId", "0".to_owned()),
    ];
    let mut output = Vec::new();
    if views.slot.empty {
        write_tag(&mut output, &views.slot.tag, false, &[], &[]);
        write_tag(&mut output, &view, true, &[], &appended);
        write_close(&mut output, &views.slot.tag.name);
    } else {
        output.extend_from_slice(&source[views.slot.span.start..views.slot.close_start]);
        write_tag(&mut output, &view, true, &[], &appended);
        output.extend_from_slice(&source[views.slot.close_start..views.slot.span.end]);
    }
    replace(source, views.slot.span, &output)
}

fn insert_views(source: &[u8], layout: &Layout) -> Result<Vec<u8>> {
    let views_name = sibling_name(&layout.root.tag.name, "sheetViews");
    let view_name = sibling_name(&views_name, "sheetView");
    let views = Tag {
        name: views_name.clone().into_boxed_str(),
        attributes: Box::new([]),
    };
    let view = Tag {
        name: view_name.into_boxed_str(),
        attributes: Box::new([]),
    };
    let appended = [
        ("tabSelected", "1".to_owned()),
        ("workbookViewId", "0".to_owned()),
    ];
    let mut inserted = Vec::new();
    write_tag(&mut inserted, &views, false, &[], &[]);
    write_tag(&mut inserted, &view, true, &[], &appended);
    write_close(&mut inserted, &views_name);

    if layout.root.empty {
        let mut replacement = Vec::new();
        write_tag(&mut replacement, &layout.root.tag, false, &[], &[]);
        replacement.extend_from_slice(&inserted);
        write_close(&mut replacement, &layout.root.tag.name);
        return Ok(replace(source, layout.root.span, &replacement));
    }
    let at = layout.insertion_before.unwrap_or(layout.root.close_start);
    if at < layout.root.tag_end || at > layout.root.close_start {
        return Err(invalid("sheet-view insertion point lies outside the root"));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len().saturating_add(inserted.len()))
        .map_err(|source| allocation("sheet-view output", source))?;
    output.extend_from_slice(&source[..at]);
    output.extend_from_slice(&inserted);
    output.extend_from_slice(&source[at..]);
    Ok(output)
}

fn rewrite_slot(
    source: &[u8],
    slot: &Slot,
    removed: &[&str],
    appended: &[(&str, String)],
) -> Vec<u8> {
    let mut replacement = Vec::new();
    write_tag(&mut replacement, &slot.tag, slot.empty, removed, appended);
    if !slot.empty {
        replacement.extend_from_slice(&source[slot.tag_end..slot.span.end]);
    }
    replace(source, slot.span, &replacement)
}

fn replace(source: &[u8], span: Span, replacement: &[u8]) -> Vec<u8> {
    let removed = span.end.saturating_sub(span.start);
    let capacity = source
        .len()
        .saturating_sub(removed)
        .saturating_add(replacement.len());
    let mut output = Vec::with_capacity(capacity);
    output.extend_from_slice(&source[..span.start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&source[span.end..]);
    output
}

fn verify(content: &[u8], selected: bool, context: Context<'_>) -> Result<()> {
    let layout = scan(content)?;
    let actual = layout
        .view
        .as_ref()
        .map(|view| attribute_bool(&view.tag, "tabSelected"))
        .transpose()?
        .flatten()
        .unwrap_or(false);
    if actual == selected {
        Ok(())
    } else {
        Err(invalid(format!(
            "sheet tab selection verification failed at '{}'",
            context.sheet
        )))
    }
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
                let kind = scanner.start(
                    stack.last().map(|frame| frame.kind),
                    &namespace,
                    &element,
                    decoder,
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
                Span {
                    start: event_start,
                    end: event_end,
                },
            )?,
            Event::End(_) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("sheet-view scan has an unmatched closing tag"))?;
                scanner.finish(frame, event_start, event_end)?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !stack.is_empty() {
        return Err(invalid("sheet-view scan ended inside an element"));
    }
    scanner.finish_layout()
}

impl Scanner {
    fn start(
        &mut self,
        parent: Option<Kind>,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        start: usize,
        end: usize,
    ) -> Result<Kind> {
        if parent.is_none() {
            if self.root_kind.is_some() {
                return Err(invalid("sheet-view edit requires one root"));
            }
            self.root_kind = Some(root_kind(namespace, element)?);
            self.pending_root = Some(Pending {
                start,
                tag_end: end,
                tag: tag(element, decoder)?,
            });
            return Ok(Kind::Root);
        }
        self.observe_root_child(parent, namespace, element, start);
        if parent == Some(Kind::Root)
            && is_spreadsheetml_name(namespace, element.name(), b"sheetViews")
        {
            if self.views.is_some() || self.pending_views.is_some() {
                return Err(invalid("duplicate direct sheetViews during tab edit"));
            }
            self.pending_views = Some(Pending {
                start,
                tag_end: end,
                tag: tag(element, decoder)?,
            });
            return Ok(Kind::Views);
        }
        if parent == Some(Kind::Views)
            && is_spreadsheetml_name(namespace, element.name(), b"sheetView")
        {
            let id = required_u32(element, b"workbookViewId", decoder)?;
            if id == 0 {
                if self.view.is_some() || self.pending_view.is_some() {
                    return Err(invalid("duplicate sheetView for workbook view zero"));
                }
                self.pending_view = Some(Pending {
                    start,
                    tag_end: end,
                    tag: tag(element, decoder)?,
                });
                return Ok(Kind::View);
            }
            return Ok(Kind::Other);
        }
        if parent == Some(Kind::Views) {
            self.views_payload = true;
        }
        Ok(Kind::Other)
    }

    fn empty(
        &mut self,
        parent: Option<Kind>,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        span: Span,
    ) -> Result<()> {
        if parent.is_none() {
            if self.root_kind.is_some() {
                return Err(invalid("sheet-view edit requires one root"));
            }
            self.root_kind = Some(root_kind(namespace, element)?);
            self.root = Some(Slot {
                span,
                tag_end: span.end,
                close_start: span.end,
                tag: tag(element, decoder)?,
                empty: true,
            });
            return Ok(());
        }
        self.observe_root_child(parent, namespace, element, span.start);
        if parent == Some(Kind::Root)
            && is_spreadsheetml_name(namespace, element.name(), b"sheetViews")
        {
            if self.views.is_some() || self.pending_views.is_some() {
                return Err(invalid("duplicate direct sheetViews during tab edit"));
            }
            self.views = Some(Container {
                slot: Slot {
                    span,
                    tag_end: span.end,
                    close_start: span.end,
                    tag: tag(element, decoder)?,
                    empty: true,
                },
                payload: false,
            });
        } else if parent == Some(Kind::Views)
            && is_spreadsheetml_name(namespace, element.name(), b"sheetView")
        {
            let id = required_u32(element, b"workbookViewId", decoder)?;
            if id == 0 {
                if self.view.is_some() || self.pending_view.is_some() {
                    return Err(invalid("duplicate sheetView for workbook view zero"));
                }
                self.view = Some(Slot {
                    span,
                    tag_end: span.end,
                    close_start: span.end,
                    tag: tag(element, decoder)?,
                    empty: true,
                });
            }
        } else if parent == Some(Kind::Views) {
            self.views_payload = true;
        }
        Ok(())
    }

    fn observe_root_child(
        &mut self,
        parent: Option<Kind>,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        start: usize,
    ) {
        if parent != Some(Kind::Root) {
            return;
        }
        if is_mce_name(namespace, element, b"AlternateContent") {
            self.alternate_content = true;
            self.root_payload = true;
            return;
        }
        let local = element.name().local_name();
        if is_spreadsheetml_name(namespace, element.name(), local.as_ref()) {
            if local.as_ref() != b"sheetViews"
                && self.insertion_before.is_none()
                && self
                    .root_kind
                    .is_some_and(|kind| follows_views(kind, local.as_ref()))
            {
                self.insertion_before = Some(start);
            }
        } else {
            self.root_payload = true;
        }
    }

    fn finish(&mut self, frame: Frame, close_start: usize, end: usize) -> Result<()> {
        match frame.kind {
            Kind::Root => {
                let pending = self
                    .pending_root
                    .take()
                    .ok_or_else(|| invalid("sheet root closed without edit state"))?;
                self.root = Some(Slot {
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
            Kind::Views => {
                let pending = self
                    .pending_views
                    .take()
                    .ok_or_else(|| invalid("sheetViews closed without edit state"))?;
                self.views = Some(Container {
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
                    payload: self.views_payload,
                });
            },
            Kind::View => {
                let pending = self
                    .pending_view
                    .take()
                    .ok_or_else(|| invalid("sheetView closed without edit state"))?;
                self.view = Some(Slot {
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
            Kind::Other => {},
        }
        Ok(())
    }

    fn finish_layout(self) -> Result<Layout> {
        if self.root_kind.is_none() {
            return Err(invalid("sheet-view scan did not find a root"));
        }
        let root = self
            .root
            .ok_or_else(|| invalid("sheet-view scan did not finish its root"))?;
        Ok(Layout {
            root,
            views: self.views,
            view: self.view,
            insertion_before: self.insertion_before,
            alternate_content: self.alternate_content,
            root_payload: self.root_payload,
        })
    }
}

fn root_kind(namespace: &ResolveResult<'_>, element: &BytesStart<'_>) -> Result<RootKind> {
    let local = element.name().local_name();
    if !is_spreadsheetml_name(namespace, element.name(), local.as_ref()) {
        return Err(invalid("tab selection requires a SpreadsheetML sheet root"));
    }
    match local.as_ref() {
        b"worksheet" | b"dialogsheet" | b"macrosheet" => Ok(RootKind::Tabular),
        b"chartsheet" => Ok(RootKind::Chart),
        name => Err(invalid(format!(
            "unsupported SpreadsheetML sheet root '{}' during tab selection",
            String::from_utf8_lossy(name)
        ))),
    }
}

fn follows_views(kind: RootKind, local: &[u8]) -> bool {
    match kind {
        RootKind::Tabular => !matches!(local, b"sheetPr" | b"dimension"),
        RootKind::Chart => local != b"sheetPr",
    }
}

fn required_u32(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<u32> {
    let value = unqualified_attribute_value(element, name, decoder)?
        .ok_or_else(|| invalid("sheetView is missing workbookViewId during tab edit"))?;
    value
        .parse::<u32>()
        .map_err(|_| invalid(format!("invalid workbookViewId '{value}' during tab edit")))
}

fn attribute_bool(tag: &Tag, name: &str) -> Result<Option<bool>> {
    tag.attributes
        .iter()
        .find(|attribute| attribute.name.as_ref() == name)
        .map(|attribute| match attribute.value.as_ref() {
            "1" | "true" => Ok(true),
            "0" | "false" => Ok(false),
            value => Err(invalid(format!(
                "invalid sheet-view boolean '{value}' during verification"
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
        .map_err(|error| invalid(format!("sheet element name is not UTF-8: {error}")))?
        .to_owned();
    let mut attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| invalid(format!("sheet attribute name is not UTF-8: {error}")))?
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
        .map_err(|_| invalid("sheet XML position does not fit usize"))
}

#[cfg(test)]
mod tests {
    use super::*;

    const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    fn context() -> Context<'static> {
        Context {
            sheet: "Sheet2",
            position: 1,
        }
    }

    #[test]
    fn toggles_only_the_matching_prefixed_view() {
        let source = format!(
            r#"<x:worksheet xmlns:x="{S}" xmlns:z="urn:future"><x:dimension ref="A1"/><x:sheetViews><x:sheetView tabSelected="1" workbookViewId="0" z:keep="yes"><x:selection activeCell="C2"/></x:sheetView><x:sheetView tabSelected="1" workbookViewId="2"/></x:sheetViews><x:sheetData/></x:worksheet>"#
        );
        let cleared = rewrite(source.as_bytes(), false, context()).expect("clear");
        let text = std::str::from_utf8(&cleared).expect("UTF-8");
        assert!(text.contains(
            r#"<x:sheetView workbookViewId="0" z:keep="yes"><x:selection activeCell="C2"/></x:sheetView>"#
        ));
        assert!(text.contains(r#"<x:sheetView tabSelected="1" workbookViewId="2"/>"#));
        let selected = rewrite(&cleared, true, context()).expect("select");
        assert!(
            std::str::from_utf8(&selected)
                .expect("UTF-8")
                .contains(r#"z:keep="yes" tabSelected="1""#)
        );
    }

    #[test]
    fn inserts_views_in_worksheet_schema_order() {
        let source = format!(
            r#"<worksheet xmlns="{S}"><sheetPr/><dimension ref="A1"/><sheetData exact="yes"/></worksheet>"#
        );
        let selected = rewrite(source.as_bytes(), true, context()).expect("select");
        let text = std::str::from_utf8(&selected).expect("UTF-8");
        assert!(text.contains(
            r#"<dimension ref="A1"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetData exact="yes"/>"#
        ));
        assert_eq!(
            rewrite(source.as_bytes(), false, context()).expect("false no-op"),
            source.as_bytes()
        );
    }

    #[test]
    fn expands_empty_views_and_supports_chartsheets() {
        let empty = format!(
            r#"<worksheet xmlns="{S}"><dimension ref="A1"/><sheetViews data="kept"/><sheetData/></worksheet>"#
        );
        let selected = rewrite(empty.as_bytes(), true, context()).expect("empty views");
        assert!(std::str::from_utf8(&selected).expect("UTF-8").contains(
            r#"<sheetViews data="kept"><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>"#
        ));

        let chart = format!(
            r#"<c:chartsheet xmlns:c="{S}"><c:sheetPr/><c:drawing r:id="d1" xmlns:r="urn:rel"/></c:chartsheet>"#
        );
        let selected = rewrite(chart.as_bytes(), true, context()).expect("chart view");
        assert!(std::str::from_utf8(&selected).expect("UTF-8").contains(
            r#"<c:sheetPr/><c:sheetViews><c:sheetView tabSelected="1" workbookViewId="0"/></c:sheetViews><c:drawing"#
        ));
    }

    #[test]
    fn blocks_missing_views_when_compatibility_payload_may_own_them() {
        let source = format!(
            r#"<worksheet xmlns="{S}" xmlns:mc="{mce}"><mc:AlternateContent><mc:Fallback/></mc:AlternateContent><sheetData/></worksheet>"#,
            mce = String::from_utf8_lossy(MCE)
        );
        assert!(matches!(
            rewrite(source.as_bytes(), true, context()),
            Err(Error::TabEditBlocked {
                reason: TabEditBlock::MarkupCompatibility,
                ..
            })
        ));
    }
}
