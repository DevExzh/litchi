//! `SpreadsheetML` scanning and exact-span XML primitives for catalog edits.

use litchi_core::xml::escape_xml;
use litchi_ooxml_common::xml::unqualified_attribute_value;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::{
    Attribute, Container, DefinedNameSlot, Dialect, Layout, SheetSlot, Slot, Span, Tag, ViewSlot,
};
use crate::error::{Result, invalid};
use crate::raw::namespace::{
    STRICT_SPREADSHEETML_NAMESPACE, is_spreadsheetml_name, relationship_attribute_value,
};

pub(crate) const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Kind {
    Workbook,
    Sheets,
    Sheet,
    BookViews,
    WorkbookView,
    DefinedNames,
    DefinedName,
    Other,
}

#[derive(Debug, Clone, Copy)]
struct Frame {
    kind: Kind,
    alternate_content: bool,
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
    root: Option<Tag>,
    dialect: Option<Dialect>,
    sheets: Option<Container>,
    pending_sheets: Option<Pending>,
    sheet_slots: Vec<SheetSlot>,
    pending_sheet: Option<(Pending, Box<str>)>,
    sheets_payload: bool,
    book_views: Option<Container>,
    pending_book_views: Option<Pending>,
    workbook_views: Vec<ViewSlot>,
    pending_workbook_view: Option<(Pending, Option<usize>, Option<u32>)>,
    book_views_payload: bool,
    defined_names: Option<Container>,
    pending_defined_names: Option<Pending>,
    defined_name_slots: Vec<DefinedNameSlot>,
    pending_defined_name: Option<(Pending, Option<usize>)>,
    defined_names_payload: bool,
    protected: bool,
    alternate_content: bool,
    alternate_dependencies: bool,
}

pub(crate) fn dialect(content: &[u8]) -> Result<Dialect> {
    Ok(scan(content)?.dialect)
}

pub(super) fn relationship_attribute_name(sheet: &SheetSlot) -> Option<&str> {
    let mut found = sheet.slot.tag.attributes.iter().filter(|attribute| {
        attribute.value.as_ref() == sheet.relationship_id.as_ref()
            && attribute
                .name
                .rsplit_once(':')
                .is_some_and(|(_, local)| local == "id")
    });
    let name = found.next()?.name.as_ref();
    found.next().is_none().then_some(name)
}

pub(super) fn relationship_attribute_from_namespaces(root: &Tag) -> Option<String> {
    root.attributes.iter().find_map(|attribute| {
        let prefix = attribute.name.strip_prefix("xmlns:")?;
        matches!(
            attribute.value.as_ref(),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                | "http://purl.oclc.org/ooxml/officeDocument/relationships"
        )
        .then(|| format!("{prefix}:id"))
    })
}

pub(super) fn rewrite_slot(
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

pub(super) fn scan(content: &[u8]) -> Result<Layout> {
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
                let alternate_content = stack.last().is_some_and(|frame| frame.alternate_content)
                    || is_mce_name(&namespace, &element, b"AlternateContent");
                let kind = scanner.start(
                    parent,
                    &namespace,
                    &element,
                    decoder,
                    &resolver,
                    alternate_content,
                    event_start,
                    event_end,
                )?;
                stack.push(Frame {
                    kind,
                    alternate_content,
                });
            },
            Event::Empty(element) => {
                let alternate_content = stack.last().is_some_and(|frame| frame.alternate_content)
                    || is_mce_name(&namespace, &element, b"AlternateContent");
                scanner.empty(
                    stack.last().map(|frame| frame.kind),
                    &namespace,
                    &element,
                    decoder,
                    &resolver,
                    alternate_content,
                    Span {
                        start: event_start,
                        end: event_end,
                    },
                )?;
            },
            Event::End(_) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("workbook edit scan has an unmatched closing tag"))?;
                scanner.finish(frame, event_start, event_end)?;
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if !stack.is_empty() {
        return Err(invalid("workbook edit scan ended inside an element"));
    }
    scanner.finish_layout()
}

impl Scanner {
    #[allow(
        clippy::too_many_arguments,
        reason = "arguments are the complete catalog rewrite state"
    )]
    fn start(
        &mut self,
        parent: Option<Kind>,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
        alternate_content: bool,
        start: usize,
        end: usize,
    ) -> Result<Kind> {
        if parent.is_none() {
            if self.root_seen || !is_spreadsheetml_name(namespace, element.name(), b"workbook") {
                return Err(invalid("workbook edit requires one SpreadsheetML root"));
            }
            self.root_seen = true;
            self.root = Some(tag(element, decoder)?);
            self.dialect = Some(
                matches!(
                    namespace,
                    ResolveResult::Bound(Namespace(value))
                        if *value == STRICT_SPREADSHEETML_NAMESPACE
                )
                .then_some(Dialect::Strict)
                .unwrap_or(Dialect::Transitional),
            );
            return Ok(Kind::Workbook);
        }
        self.observe_guard(parent, namespace, element, decoder, alternate_content)?;
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
            if self.pending_workbook_view.is_some() {
                return Err(invalid("nested direct workbookView element during edit"));
            }
            self.pending_workbook_view = Some((
                Pending {
                    start,
                    tag_end: end,
                    tag: tag(element, decoder)?,
                },
                optional_usize(element, b"activeTab", decoder, "workbook activeTab")?,
                optional_u32(element, b"firstSheet", decoder, "workbook firstSheet")?,
            ));
            return Ok(Kind::WorkbookView);
        }
        if parent == Some(Kind::Workbook)
            && is_spreadsheetml_name(namespace, element.name(), b"definedNames")
        {
            if self.defined_names.is_some() || self.pending_defined_names.is_some() {
                return Err(invalid("duplicate direct definedNames element during edit"));
            }
            self.pending_defined_names = Some(Pending {
                start,
                tag_end: end,
                tag: tag(element, decoder)?,
            });
            return Ok(Kind::DefinedNames);
        }
        if parent == Some(Kind::DefinedNames)
            && is_spreadsheetml_name(namespace, element.name(), b"definedName")
        {
            if self.pending_defined_name.is_some() {
                return Err(invalid("nested direct definedName element during edit"));
            }
            self.pending_defined_name = Some((
                Pending {
                    start,
                    tag_end: end,
                    tag: tag(element, decoder)?,
                },
                optional_usize(
                    element,
                    b"localSheetId",
                    decoder,
                    "defined name localSheetId",
                )?,
            ));
            return Ok(Kind::DefinedName);
        }
        if parent == Some(Kind::BookViews) {
            self.book_views_payload = true;
        } else if parent == Some(Kind::Sheets) {
            self.sheets_payload = true;
        } else if parent == Some(Kind::DefinedNames) {
            self.defined_names_payload = true;
        }
        Ok(Kind::Other)
    }

    #[allow(
        clippy::too_many_arguments,
        reason = "arguments are the complete catalog relationship rewrite state"
    )]
    fn empty(
        &mut self,
        parent: Option<Kind>,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
        alternate_content: bool,
        span: Span,
    ) -> Result<()> {
        self.observe_guard(parent, namespace, element, decoder, alternate_content)?;
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
            self.workbook_views.push(ViewSlot {
                slot: Slot {
                    span,
                    tag_end: span.end,
                    close_start: span.end,
                    tag: tag(element, decoder)?,
                    empty: true,
                },
                active: optional_usize(element, b"activeTab", decoder, "workbook activeTab")?,
                first: optional_u32(element, b"firstSheet", decoder, "workbook firstSheet")?,
            });
        } else if parent == Some(Kind::Workbook)
            && is_spreadsheetml_name(namespace, element.name(), b"definedNames")
        {
            if self.defined_names.is_some() || self.pending_defined_names.is_some() {
                return Err(invalid("duplicate direct definedNames element during edit"));
            }
            self.defined_names = Some(Container {
                slot: Slot {
                    span,
                    tag_end: span.end,
                    close_start: span.end,
                    tag: tag(element, decoder)?,
                    empty: true,
                },
                payload: false,
            });
        } else if parent == Some(Kind::DefinedNames)
            && is_spreadsheetml_name(namespace, element.name(), b"definedName")
        {
            self.defined_name_slots.push(DefinedNameSlot {
                slot: Slot {
                    span,
                    tag_end: span.end,
                    close_start: span.end,
                    tag: tag(element, decoder)?,
                    empty: true,
                },
                local_sheet_id: optional_usize(
                    element,
                    b"localSheetId",
                    decoder,
                    "defined name localSheetId",
                )?,
            });
        } else if parent == Some(Kind::BookViews) {
            self.book_views_payload = true;
        } else if parent == Some(Kind::Sheets) {
            self.sheets_payload = true;
        } else if parent == Some(Kind::DefinedNames) {
            self.defined_names_payload = true;
        }
        Ok(())
    }

    fn observe_guard(
        &mut self,
        parent: Option<Kind>,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        alternate_content: bool,
    ) -> Result<()> {
        if is_spreadsheetml_name(namespace, element.name(), b"workbookProtection") {
            self.protected |= optional_bool(element, b"lockStructure", decoder)?.unwrap_or(false);
        }
        if matches!(
            parent,
            Some(Kind::Workbook | Kind::Sheets | Kind::BookViews | Kind::DefinedNames)
        ) && is_mce_name(namespace, element, b"AlternateContent")
        {
            self.alternate_content = true;
        }
        if alternate_content && is_order_dependency_name(namespace, element) {
            self.alternate_dependencies = true;
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
                    payload: self.sheets_payload,
                });
            },
            Kind::WorkbookView => {
                let (pending, active, first) = self
                    .pending_workbook_view
                    .take()
                    .ok_or_else(|| invalid("workbookView close without edit state"))?;
                self.workbook_views.push(ViewSlot {
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
                    active,
                    first,
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
            Kind::DefinedName => {
                let (pending, local_sheet_id) = self
                    .pending_defined_name
                    .take()
                    .ok_or_else(|| invalid("definedName close without edit state"))?;
                self.defined_name_slots.push(DefinedNameSlot {
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
                    local_sheet_id,
                });
            },
            Kind::DefinedNames => {
                let pending = self
                    .pending_defined_names
                    .take()
                    .ok_or_else(|| invalid("definedNames close without edit state"))?;
                self.defined_names = Some(Container {
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
                    payload: self.defined_names_payload,
                });
            },
            Kind::Workbook | Kind::Other => {},
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
        let root = self
            .root
            .ok_or_else(|| invalid("workbook edit scan lost its root element"))?;
        let dialect = self
            .dialect
            .ok_or_else(|| invalid("workbook edit scan lost its XML dialect"))?;
        Ok(Layout {
            root,
            dialect,
            sheets,
            sheet_slots: self.sheet_slots.into_boxed_slice(),
            book_views: self.book_views,
            workbook_views: self.workbook_views.into_boxed_slice(),
            defined_names: self.defined_names,
            defined_name_slots: self.defined_name_slots.into_boxed_slice(),
            protected: self.protected,
            alternate_content: self.alternate_content,
            alternate_dependencies: self.alternate_dependencies,
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

fn optional_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Option<u32>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| {
            value
                .parse::<u32>()
                .map_err(|_source| invalid(format!("invalid {description} '{value}' during edit")))
        })
        .transpose()
}

fn optional_usize(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Option<usize>> {
    optional_u32(element, name, decoder, description)?
        .map(|value| {
            usize::try_from(value)
                .map_err(|_source| invalid(format!("{description} does not fit usize during edit")))
        })
        .transpose()
}

pub(super) fn write_tag(
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

pub(super) fn write_close(output: &mut Vec<u8>, name: &str) {
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

pub(super) fn sibling_name(name: &str, local: &str) -> String {
    name.split_once(':').map_or_else(
        || local.to_owned(),
        |(prefix, _)| format!("{prefix}:{local}"),
    )
}

fn is_mce_name(namespace: &ResolveResult<'_>, element: &BytesStart<'_>, local: &[u8]) -> bool {
    element.name().local_name().as_ref() == local
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == MCE)
}

fn is_order_dependency_name(namespace: &ResolveResult<'_>, element: &BytesStart<'_>) -> bool {
    [
        b"sheets".as_slice(),
        b"sheet".as_slice(),
        b"bookViews".as_slice(),
        b"workbookView".as_slice(),
        b"definedNames".as_slice(),
        b"definedName".as_slice(),
    ]
    .iter()
    .any(|local| is_spreadsheetml_name(namespace, element.name(), local))
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_source| invalid("workbook XML position does not fit usize"))
}
