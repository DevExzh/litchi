//! Borrowed DrawingML table views embedded in PresentationML graphic frames.
//!
//! The table, row, and cell values retain the caller's XML allocation. Only
//! compact byte spans are allocated while indexing, so reading a large table
//! does not duplicate its XML subtrees.

use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::mce::{InScopeNamespaces, Limits as MceLimits};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_ooxml_common::xml::{
    DRAWINGML_NAMESPACE, STRICT_DRAWINGML_NAMESPACE, decode_xml_reference,
    unqualified_attribute_value,
};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, NamespaceResolver, PrefixDeclaration, QName, ResolveResult};
use quick_xml::reader::NsReader;

use crate::{Error, Result};

const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_NODES: usize = 1_000_000;
const MAX_ROWS: usize = 100_000;
const MAX_CELLS: usize = 1_000_000;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Span {
    start: usize,
    end: usize,
}

impl Span {
    fn get(self, xml: &[u8]) -> Result<&[u8]> {
        xml.get(self.start..self.end)
            .ok_or_else(|| invalid("table byte span is outside its owner XML"))
    }
}

/// The namespace scope at the root of a fragment cut from an owner XML.
///
/// Table rows and cells are indexed from progressively smaller borrowed
/// fragments. Keeping the owner and the absolute start of each fragment lets
/// every scan restore the declarations that fragment inherits before its own
/// root declarations are processed. The shared bounded owner walk is repeated
/// for nested projections; this keeps the bridge allocation-free while leaving
/// scope caching as a separate optimization.
#[derive(Debug, Clone, Copy)]
struct NamespaceContext<'a> {
    owner: &'a [u8],
    offset: usize,
}

impl NamespaceContext<'_> {
    fn resolver(self) -> Result<NamespaceResolver> {
        let scope = InScopeNamespaces::at_offset(self.owner, self.offset, &MceLimits::default())?;
        let mut resolver = NamespaceResolver::default();
        for (prefix, namespace) in scope.iter() {
            let declaration = if prefix.is_empty() {
                PrefixDeclaration::Default
            } else {
                PrefixDeclaration::Named(prefix)
            };
            resolver
                .add(declaration, Namespace(namespace))
                .map_err(|_| invalid("table XML namespace scope is invalid"))?;
        }
        Ok(resolver)
    }

    fn child(self, relative: usize) -> Result<Self> {
        let offset = self
            .offset
            .checked_add(relative)
            .ok_or_else(|| invalid("table XML namespace scope offset overflows usize"))?;
        Ok(Self {
            owner: self.owner,
            offset,
        })
    }
}

/// Table style switches and the referenced table-style GUID.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Properties {
    pub first_row: bool,
    pub first_col: bool,
    pub last_row: bool,
    pub last_col: bool,
    pub band_row: bool,
    pub band_col: bool,
    pub style_id: Option<String>,
}

/// A borrowed `DrawingML` table (`a:tbl`).
#[derive(Debug, Clone, Copy)]
pub struct Table<'a> {
    xml: &'a [u8],
    span: Span,
    context: Option<NamespaceContext<'a>>,
}

impl<'a> Table<'a> {
    /// Index a standalone `a:tbl` owner without copying its bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_xml(xml: &'a [u8]) -> Result<Self> {
        Self::from_xml_with_context(xml, None)
    }

    fn from_xml_with_context(xml: &'a [u8], context: Option<NamespaceContext<'a>>) -> Result<Self> {
        let span = scan_with_context(xml, b"tbl", context)?
            .into_iter()
            .next()
            .ok_or_else(|| Error::PartNotFound("DrawingML table not found".into()))?;
        let context = context
            .map(|context| context.child(span.start))
            .transpose()?;
        Ok(Self { xml, span, context })
    }

    /// Locate the first `DrawingML` table inside a graphic-frame owner.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_graphic_frame(xml: &'a [u8]) -> Result<Self> {
        Self::from_xml(xml)
    }

    /// Locate a `DrawingML` table while retaining the namespace scope of the
    /// graphic-frame fragment in its owner XML.
    pub(crate) fn from_graphic_frame_with_context(
        xml: &'a [u8],
        owner: &'a [u8],
        owner_offset: usize,
    ) -> Result<Self> {
        Self::from_xml_with_context(
            xml,
            Some(NamespaceContext {
                owner,
                offset: owner_offset,
            }),
        )
    }

    /// Borrow the exact table element.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[inline]
    pub fn xml(self) -> Result<&'a [u8]> {
        self.span.get(self.xml)
    }

    /// Count table rows without allocating row values.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn row_count(self) -> Result<usize> {
        let table = self.xml()?;
        let count = scan_with_context(table, b"tr", self.context)?.len();
        if count > MAX_ROWS {
            return Err(limit("table row count", MAX_ROWS));
        }
        Ok(count)
    }

    /// Count columns from the first row, or return zero for an empty table.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn column_count(self) -> Result<usize> {
        Ok(self.rows()?.first().map_or(0, |row| row.cell_count()))
    }

    /// Every cell's local properties and text body, by row and position in
    /// the row, as [`Cell::properties`] and [`Cell::text_body`] read each
    /// (record 0346). A table that declares no namespace has one namespace
    /// scope, resolved once, so reading every cell is proportional to the
    /// table rather than to the part per row and cell.
    ///
    /// # Errors
    ///
    /// Returns the errors of [`Self::rows`], [`Row::cells`],
    /// [`Cell::properties`] and [`Cell::text_body`].
    #[allow(
        clippy::type_complexity,
        reason = "rows of (properties, body) pairs, the shape the callers index"
    )]
    pub fn cell_contents(
        self,
    ) -> Result<Vec<Vec<(CellProperties, Option<litchi_drawingml::text::body::Body>)>>> {
        let table = self.xml()?;
        let Some(context) = self.context.filter(|_| !contains(table, b"xmlns")) else {
            let mut out = Vec::new();
            for row in self.rows()? {
                let mut cells = Vec::new();
                for cell in row.cells()? {
                    cells.push((cell.properties()?, cell.text_body()?));
                }
                out.push(cells);
            }
            return Ok(out);
        };
        let scope = context.resolver()?;
        let rows = scan_with_scope(table, b"tr", Some(scope.clone()), false)?;
        if rows.len() > MAX_ROWS {
            return Err(limit("table row count", MAX_ROWS));
        }
        let mut out = Vec::new();
        for row_span in rows {
            let row = row_span.get(table)?;
            let spans = scan_with_scope(row, b"tc", Some(scope.clone()), false)?;
            if spans.len() > MAX_CELLS {
                return Err(limit("table cell count", MAX_CELLS));
            }
            let row_context = context.child(row_span.start)?;
            let mut cells = Vec::new();
            for span in spans {
                let cell = Cell {
                    xml: row,
                    context: Some(row_context.child(span.start)?),
                    span,
                };
                cells.push((cell.properties()?, cell.text_body_in(Some(&scope))?));
            }
            out.push(cells);
        }
        Ok(out)
    }

    /// Index rows as borrowed views over the table's source XML.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn rows(self) -> Result<Vec<Row<'a>>> {
        let table = self.xml()?;
        let spans = scan_with_context(table, b"tr", self.context)?;
        if spans.len() > MAX_ROWS {
            return Err(limit("table row count", MAX_ROWS));
        }
        spans
            .into_iter()
            .map(|span| {
                let context = self
                    .context
                    .map(|context| context.child(span.start))
                    .transpose()?;
                Ok(Row {
                    xml: table,
                    span,
                    context,
                })
            })
            .collect()
    }

    /// The number of cells of every row, in one pass over the table
    /// (record 0333): the counts [`Row::cell_count`] gives, without
    /// resolving each row's namespace scope again.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed table XML or more rows than the
    /// reader admits.
    pub fn row_cell_counts(self) -> Result<Vec<usize>> {
        let Some(context) = self.context else {
            return Ok(self.rows()?.into_iter().map(Row::cell_count).collect());
        };
        let table = self.xml()?;
        if table.len() > MAX_XML_BYTES {
            return Err(limit("table XML bytes", MAX_XML_BYTES));
        }
        let mut reader = NsReader::from_reader(table);
        *reader.resolver_mut() = context.resolver()?;
        // DrawingML local names of the open elements.
        let mut stack: Vec<Option<Vec<u8>>> = Vec::new();
        let mut counts = Vec::new();
        let drawing = |namespace: &ResolveResult<'_>| {
            matches!(namespace, ResolveResult::Bound(Namespace(value))
                if *value == DRAWINGML_NAMESPACE || *value == STRICT_DRAWINGML_NAMESPACE)
        };
        loop {
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            let empty = matches!(event, Event::Empty(_));
            match event {
                Event::Start(element) | Event::Empty(element) => {
                    let local = drawing(&namespace).then(|| element.local_name().as_ref().to_vec());
                    let parent = stack.last().and_then(Option::as_deref);
                    let grandparent = stack
                        .len()
                        .checked_sub(2)
                        .and_then(|index| stack[index].as_deref());
                    match (local.as_deref(), parent, grandparent) {
                        (Some(b"tr"), Some(b"tbl"), _) => {
                            if counts.len() >= MAX_ROWS {
                                return Err(limit("table row count", MAX_ROWS));
                            }
                            counts.push(0);
                        },
                        (Some(b"tc"), Some(b"tr"), Some(b"tbl")) => {
                            if let Some(count) = counts.last_mut() {
                                *count += 1;
                            }
                        },
                        _ => {},
                    }
                    if !empty {
                        stack.push(local);
                    }
                },
                Event::End(_) => {
                    stack.pop();
                },
                Event::Eof => break,
                _ => {},
            }
        }
        Ok(counts)
    }

    /// Get one cell by zero-based row and column position.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn cell(self, row: usize, column: usize) -> Result<Option<Cell<'a>>> {
        let Some(row) = self.rows()?.get(row).copied() else {
            return Ok(None);
        };
        Ok(row.cells()?.get(column).copied())
    }

    /// The widths of the table's grid columns (`a:tblGrid/a:gridCol/@w`),
    /// in EMU, in grid order (record 0321).
    ///
    /// # Errors
    ///
    /// Returns an error for a malformed grid or a column without a
    /// non-negative integer width.
    pub fn grid_columns(self) -> Result<Vec<i64>> {
        let table = self.xml()?;
        match self.context {
            Some(context) => {
                let end = context
                    .offset
                    .checked_add(table.len())
                    .ok_or_else(|| invalid("table XML span overflows usize"))?;
                super::columns::widths(context.owner, context.offset..end)
            },
            None => super::columns::widths(table, 0..table.len()),
        }
    }

    /// Read the optional `a:tblPr` values.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn properties(self) -> Result<Option<Properties>> {
        parse_properties(self.xml()?, self.context)
    }
}

/// A borrowed table row.
#[derive(Debug, Clone, Copy)]
pub struct Row<'a> {
    xml: &'a [u8],
    span: Span,
    context: Option<NamespaceContext<'a>>,
}

impl<'a> Row<'a> {
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[inline]
    pub fn xml(self) -> Result<&'a [u8]> {
        self.span.get(self.xml)
    }

    #[must_use]
    pub fn cell_count(self) -> usize {
        scan_with_context(self.xml().unwrap_or_default(), b"tc", self.context)
            .map_or(0, |spans| spans.len())
    }

    /// The row's height (`a:tr@h`) in EMU; `None` when the attribute is
    /// absent (record 0333).
    ///
    /// # Errors
    ///
    /// Returns an error for malformed row XML or a height that is not a
    /// non-negative integer.
    pub fn height(self) -> Result<Option<i64>> {
        let xml = self.xml()?;
        let mut reader = quick_xml::Reader::from_reader(xml);
        loop {
            match reader
                .read_event()
                .map_err(|_| invalid("malformed table row XML"))?
            {
                Event::Start(start) | Event::Empty(start) => {
                    for attribute in start.checked_attributes() {
                        let attribute =
                            attribute.map_err(|_| invalid("malformed table row XML"))?;
                        if attribute.key.as_ref() == b"h" {
                            return std::str::from_utf8(&attribute.value)
                                .ok()
                                .and_then(|value| value.trim().parse::<i64>().ok())
                                .filter(|value| *value >= 0)
                                .map(Some)
                                .ok_or_else(|| {
                                    invalid("a table row height is not a non-negative integer")
                                });
                        }
                    }
                    return Ok(None);
                },
                Event::Eof => return Ok(None),
                _ => {},
            }
        }
    }

    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn cells(self) -> Result<Vec<Cell<'a>>> {
        let row = self.xml()?;
        let spans = scan_with_context(row, b"tc", self.context)?;
        if spans.len() > MAX_CELLS {
            return Err(limit("table cell count", MAX_CELLS));
        }
        spans
            .into_iter()
            .map(|span| {
                let context = self
                    .context
                    .map(|context| context.child(span.start))
                    .transpose()?;
                Ok(Cell {
                    xml: row,
                    span,
                    context,
                })
            })
            .collect()
    }
}

/// The transitional DrawingML main namespace.
const DRAWINGML_MAIN: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";

/// A borrowed table cell.
#[derive(Debug, Clone, Copy)]
pub struct Cell<'a> {
    xml: &'a [u8],
    span: Span,
    context: Option<NamespaceContext<'a>>,
}

impl<'a> Cell<'a> {
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[inline]
    pub fn xml(self) -> Result<&'a [u8]> {
        self.span.get(self.xml)
    }

    /// Decode `DrawingML` text, preserving paragraph, break, and tab markers.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn text(self) -> Result<String> {
        extract_text(self.xml()?, self.context)
    }

    /// The cell's text body (`a:txBody`) with its local properties, read
    /// as a shape's text body is; `None` for a cell without one. Paragraph
    /// and run positions count as [`crate::opened::TextTarget::CellRun`]
    /// addresses them (record 0285).
    ///
    /// # Errors
    ///
    /// Returns an error for malformed cell XML or a text body the reader
    /// refuses.
    pub fn text_body(self) -> Result<Option<litchi_drawingml::text::body::Body>> {
        self.text_body_in(None)
    }

    /// [`Self::text_body`], reusing `shared`, the namespace scope every
    /// offset of the owning table has (record 0346): a table without
    /// namespace declarations has one scope, which
    /// [`Table::text_bodies`] computes once.
    fn text_body_in(
        self,
        shared: Option<&NamespaceResolver>,
    ) -> Result<Option<litchi_drawingml::text::body::Body>> {
        let xml = self.xml()?;
        let mut reader = quick_xml::Reader::from_reader(xml);
        let origin = ReaderOrigin::of(xml);
        let mut depth = 0usize;
        let mut body: Option<(usize, usize)> = None;
        loop {
            let start = origin
                .offset(reader.buffer_position())
                .ok_or_else(|| invalid("table cell offset overflows usize"))?;
            let event = reader
                .read_event()
                .map_err(|_| invalid("malformed table cell XML"))?;
            match event {
                Event::Start(element) => {
                    if depth == 1 && element.local_name().as_ref() == b"txBody" {
                        if body.is_some() {
                            return Err(invalid("table cell holds two text bodies"));
                        }
                        body = Some((start, 0));
                    }
                    depth += 1;
                },
                Event::Empty(element) => {
                    if depth == 1 && element.local_name().as_ref() == b"txBody" {
                        return Ok(None);
                    }
                },
                Event::End(_) => {
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("malformed table cell XML"))?;
                    if depth == 1
                        && let Some((start, 0)) = body
                    {
                        let end = origin
                            .offset(reader.buffer_position())
                            .ok_or_else(|| invalid("table cell offset overflows usize"))?;
                        body = Some((start, end));
                    }
                },
                Event::Eof => break,
                _ => {},
            }
        }
        let Some((start, end)) = body.filter(|(_, end)| *end > 0) else {
            return Ok(None);
        };
        let scope = match (shared, self.context) {
            (Some(shared), _) => shared.clone(),
            (None, Some(context)) => context.child(start)?.resolver()?,
            // A fragment cut from its part: its own prefix is DrawingML's,
            // as the cell text reader assumes.
            (None, None) => {
                let mut resolver = NamespaceResolver::default();
                let tag = &xml[start..end];
                let name_end = tag
                    .iter()
                    .position(|byte| matches!(byte, b' ' | b'>' | b'/' | b'\t' | b'\n' | b'\r'))
                    .unwrap_or(tag.len());
                let name = &tag[1.min(name_end)..name_end];
                let declaration = match name.iter().position(|byte| *byte == b':') {
                    Some(colon) => PrefixDeclaration::Named(&name[..colon]),
                    None => PrefixDeclaration::Default,
                };
                resolver
                    .add(declaration, Namespace(DRAWINGML_MAIN))
                    .map_err(|_| invalid("table cell namespace scope is invalid"))?;
                resolver
            },
        };
        let body = litchi_drawingml::text::body::reader::read_with_options(
            &xml[start..end],
            scope,
            litchi_drawingml::text::body::reader::Limits::DEFAULT,
        )?;
        Ok(Some(body))
    }

    /// The cell's local properties (`a:tcPr`, record 0333): its fill (a
    /// solid RGB or theme colour, or no fill; other fills read as `None`),
    /// vertical anchor and insets.
    ///
    /// # Errors
    ///
    /// Returns an error if the cell XML is malformed, an inset is not an
    /// integer, or an RGB value is not six hexadecimal digits.
    pub fn properties(self) -> Result<CellProperties> {
        let xml = self.xml()?;
        let mut reader = quick_xml::Reader::from_reader(xml);
        reader.config_mut().check_end_names = true;
        let mut properties = CellProperties::default();
        // Local names below the cell root.
        let mut path: Vec<Vec<u8>> = Vec::new();
        loop {
            let event = reader
                .read_event()
                .map_err(|_| invalid("malformed table cell XML"))?;
            let (start, empty) = match &event {
                Event::Start(start) => (start, false),
                Event::Empty(start) => (start, true),
                Event::End(_) => {
                    path.pop();
                    continue;
                },
                Event::Eof => return Ok(properties),
                _ => continue,
            };
            let local = start.local_name().as_ref().to_vec();
            let value = |name: &[u8]| -> Result<Option<Vec<u8>>> {
                for attribute in start.checked_attributes() {
                    let attribute = attribute.map_err(|_| invalid("malformed table cell XML"))?;
                    if attribute.key.as_ref() == name {
                        return Ok(Some(attribute.value.into_owned()));
                    }
                }
                Ok(None)
            };
            match (path.len(), path.get(1).map(Vec::as_slice), local.as_slice()) {
                (1, _, b"tcPr") => {
                    if let Some(anchor) = value(b"anchor")? {
                        properties.anchor = std::str::from_utf8(&anchor)
                            .ok()
                            .and_then(|token| token.parse().ok());
                    }
                    for (index, name) in
                        [&b"marL"[..], b"marR", b"marT", b"marB"].iter().enumerate()
                    {
                        if let Some(margin) = value(name)? {
                            properties.margins[index] = Some(
                                std::str::from_utf8(&margin)
                                    .ok()
                                    .and_then(|value| value.trim().parse::<i64>().ok())
                                    .ok_or_else(|| {
                                        invalid("a table cell inset is not an integer")
                                    })?,
                            );
                        }
                    }
                },
                (2, Some(b"tcPr"), b"noFill") => {
                    properties.fill = Some(crate::opened::CellFill::None);
                },
                (3, Some(b"tcPr"), b"srgbClr") if path[2].as_slice() == b"solidFill" => {
                    let rgb =
                        value(b"val")?.ok_or_else(|| invalid("table cell colour has no value"))?;
                    properties.fill = Some(crate::opened::CellFill::Rgb(crate::opened::Rgb(
                        parse_rgb(&rgb)?,
                    )));
                },
                (3, Some(b"tcPr"), b"schemeClr") if path[2].as_slice() == b"solidFill" => {
                    properties.fill = value(b"val")?
                        .and_then(|token| {
                            litchi_drawingml::color::Scheme::from_token(
                                std::str::from_utf8(&token).ok()?,
                            )
                        })
                        .map(crate::opened::CellFill::Scheme);
                },
                _ => {},
            }
            if !empty {
                path.push(local);
            }
        }
    }

    /// The cell's local solid RGB fill (`a:tcPr/a:solidFill/a:srgbClr`).
    ///
    /// Scheme, preset and system colours, gradient, picture and pattern
    /// fills, and fills inherited from a table style return `None`.
    ///
    /// # Errors
    ///
    /// Returns an error if the cell XML is malformed or its colour value is
    /// not six hexadecimal digits.
    pub fn fill(self) -> Result<Option<[u8; 3]>> {
        let xml = self.xml()?;
        let mut reader = quick_xml::Reader::from_reader(xml);
        reader.config_mut().check_end_names = true;
        // Path of local names below the cell root.
        let mut path: Vec<Vec<u8>> = Vec::new();
        loop {
            let event = reader
                .read_event()
                .map_err(|_| invalid("malformed table cell XML"))?;
            let (start, empty) = match &event {
                Event::Start(start) => (Some(start), false),
                Event::Empty(start) => (Some(start), true),
                Event::End(_) => {
                    path.pop();
                    continue;
                },
                Event::Eof => return Ok(None),
                _ => continue,
            };
            let Some(start) = start else { continue };
            let local = start.local_name().as_ref().to_vec();
            let inside = path.len() == 3
                && path[1].as_slice() == b"tcPr"
                && path[2].as_slice() == b"solidFill";
            if inside && local.as_slice() == b"srgbClr" {
                for attribute in start.checked_attributes() {
                    let attribute = attribute.map_err(|_| invalid("malformed table cell XML"))?;
                    if attribute.key.as_ref() == b"val" {
                        return parse_rgb(&attribute.value).map(Some);
                    }
                }
                return Err(invalid("table cell colour has no value"));
            }
            if !empty {
                path.push(local);
            }
        }
    }
}

/// A table cell's local properties (`a:tcPr`, record 0333).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[non_exhaustive]
pub struct CellProperties {
    /// The local fill: solid RGB or theme colour, or no fill.
    pub fill: Option<crate::opened::CellFill>,
    /// The vertical text anchor (`anchor`).
    pub anchor: Option<litchi_drawingml::text::Anchor>,
    /// The insets in EMU, in `marL`, `marR`, `marT`, `marB` order.
    pub margins: [Option<i64>; 4],
}

fn parse_rgb(value: &[u8]) -> Result<[u8; 3]> {
    if value.len() != 6 {
        return Err(invalid("table cell colour is not six hexadecimal digits"));
    }
    let mut rgb = [0u8; 3];
    for (index, slot) in rgb.iter_mut().enumerate() {
        let pair = std::str::from_utf8(&value[index * 2..index * 2 + 2])
            .map_err(|_| invalid("table cell colour is not hexadecimal"))?;
        *slot = u8::from_str_radix(pair, 16)
            .map_err(|_| invalid("table cell colour is not hexadecimal"))?;
    }
    Ok(rgb)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}

fn is_drawingml(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    local: &[u8],
    fragment_prefix: &Option<Option<Vec<u8>>>,
    allow_fragment_fallback: bool,
) -> bool {
    if name.local_name().as_ref() != local {
        return false;
    }
    match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            *value == DRAWINGML_NAMESPACE || *value == STRICT_DRAWINGML_NAMESPACE
        },
        ResolveResult::Unknown(prefix) if allow_fragment_fallback => {
            prefix.as_slice() == b"a"
                || fragment_prefix.as_ref().and_then(Option::as_deref) == Some(prefix.as_slice())
        },
        ResolveResult::Unbound if allow_fragment_fallback => fragment_prefix == &Some(None),
        ResolveResult::Unknown(_) | ResolveResult::Unbound => false,
    }
}

fn scan_with_context(
    xml: &[u8],
    target: &[u8],
    context: Option<NamespaceContext<'_>>,
) -> Result<Vec<Span>> {
    let scope = context.map(NamespaceContext::resolver).transpose()?;
    scan_with_scope(xml, target, scope, context.is_none())
}

/// [`scan_with_context`] with the scope at the start of `xml` already
/// resolved (record 0346).
fn scan_with_scope(
    xml: &[u8],
    target: &[u8],
    scope: Option<NamespaceResolver>,
    allow_fragment_fallback: bool,
) -> Result<Vec<Span>> {
    enum ScanEvent {
        Start { local: Vec<u8>, matched: bool },
        Empty { matched: bool },
        End { local: Vec<u8> },
        Other,
        Forbidden,
        Eof,
    }

    if xml.len() > MAX_XML_BYTES {
        return Err(limit("table XML bytes", MAX_XML_BYTES));
    }
    let mut reader = NsReader::from_reader(xml);
    if let Some(scope) = scope {
        *reader.resolver_mut() = scope;
    }
    let origin = ReaderOrigin::of(xml);
    let mut stack: Vec<(usize, Vec<u8>, bool)> = Vec::new();
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut spans = Vec::new();
    let mut nodes = 0usize;

    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("table XML offset exceeds usize"))?;
        let event = {
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            match event {
                Event::Start(element) => {
                    if fragment_prefix.is_none() && !matches!(namespace, ResolveResult::Bound(_)) {
                        fragment_prefix =
                            Some(element.name().prefix().map(|p| p.into_inner().to_vec()));
                    }
                    let matched = is_drawingml(
                        &namespace,
                        element.name(),
                        target,
                        &fragment_prefix,
                        allow_fragment_fallback,
                    );
                    ScanEvent::Start {
                        local: element.local_name().as_ref().to_vec(),
                        matched,
                    }
                },
                Event::Empty(element) => {
                    if fragment_prefix.is_none() && !matches!(namespace, ResolveResult::Bound(_)) {
                        fragment_prefix =
                            Some(element.name().prefix().map(|p| p.into_inner().to_vec()));
                    }
                    ScanEvent::Empty {
                        matched: is_drawingml(
                            &namespace,
                            element.name(),
                            target,
                            &fragment_prefix,
                            allow_fragment_fallback,
                        ),
                    }
                },
                Event::End(element) => ScanEvent::End {
                    local: element.local_name().as_ref().to_vec(),
                },
                Event::DocType(_) | Event::PI(_) => ScanEvent::Forbidden,
                Event::Eof => ScanEvent::Eof,
                _ => ScanEvent::Other,
            }
        };
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("table XML offset exceeds usize"))?;
        nodes = nodes
            .checked_add(1)
            .ok_or_else(|| invalid("table XML node count overflow"))?;
        if nodes > MAX_NODES {
            return Err(limit("table XML nodes", MAX_NODES));
        }

        match event {
            ScanEvent::Start { local, matched } => {
                let depth = stack
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| invalid("table XML depth overflow"))?;
                if depth > MAX_DEPTH {
                    return Err(limit("table XML depth", MAX_DEPTH));
                }
                stack.push((start, local, matched));
            },
            ScanEvent::Empty { matched } => {
                if matched {
                    spans.push(Span { start, end });
                }
            },
            ScanEvent::End { local: closing } => {
                let Some((open, local, matched)) = stack.pop() else {
                    return Err(invalid("table XML has an unexpected closing element"));
                };
                if local.as_slice() != closing.as_slice() {
                    return Err(invalid("table XML closing element does not match"));
                }
                if matched {
                    spans.push(Span { start: open, end });
                }
            },
            ScanEvent::Forbidden => return Err(invalid("table XML contains forbidden markup")),
            ScanEvent::Eof => break,
            ScanEvent::Other => {},
        }
    }
    if !stack.is_empty() {
        return Err(invalid("table XML is unterminated"));
    }
    Ok(spans)
}

fn parse_properties(
    xml: &[u8],
    context: Option<NamespaceContext<'_>>,
) -> Result<Option<Properties>> {
    let allow_fragment_fallback = context.is_none();
    let mut reader = NsReader::from_reader(xml);
    if let Some(context) = context {
        *reader.resolver_mut() = context.resolver()?;
    }
    let mut fragment_prefix = None;
    let mut depth = 0usize;
    let mut properties = None;
    let mut property_depth = None;
    let mut style_depth = None;
    let mut style = String::new();

    loop {
        let decoder = reader.decoder();
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                if fragment_prefix.is_none() && !matches!(namespace, ResolveResult::Bound(_)) {
                    fragment_prefix =
                        Some(element.name().prefix().map(|p| p.into_inner().to_vec()));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("table XML depth overflow"))?;
                if depth == 1
                    && !is_drawingml(
                        &namespace,
                        element.name(),
                        b"tbl",
                        &fragment_prefix,
                        allow_fragment_fallback,
                    )
                {
                    return Ok(None);
                }
                if depth == 2
                    && is_drawingml(
                        &namespace,
                        element.name(),
                        b"tblPr",
                        &fragment_prefix,
                        allow_fragment_fallback,
                    )
                {
                    if properties.is_some() {
                        return Err(invalid("table has multiple tblPr elements"));
                    }
                    properties = Some(Properties {
                        first_row: on_off(&element, b"firstRow", decoder)?,
                        first_col: on_off(&element, b"firstCol", decoder)?,
                        last_row: on_off(&element, b"lastRow", decoder)?,
                        last_col: on_off(&element, b"lastCol", decoder)?,
                        band_row: on_off(&element, b"bandRow", decoder)?,
                        band_col: on_off(&element, b"bandCol", decoder)?,
                        style_id: None,
                    });
                    property_depth = Some(depth);
                } else if property_depth == Some(depth - 1)
                    && is_drawingml(
                        &namespace,
                        element.name(),
                        b"tableStyleId",
                        &fragment_prefix,
                        allow_fragment_fallback,
                    )
                    && style_depth.replace(depth).is_some()
                {
                    return Err(invalid("table has multiple tableStyleId elements"));
                }
            },
            Event::Empty(element) => {
                if fragment_prefix.is_none() && !matches!(namespace, ResolveResult::Bound(_)) {
                    fragment_prefix =
                        Some(element.name().prefix().map(|p| p.into_inner().to_vec()));
                }
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("table XML depth overflow"))?;
                if child_depth == 1
                    && !is_drawingml(
                        &namespace,
                        element.name(),
                        b"tbl",
                        &fragment_prefix,
                        allow_fragment_fallback,
                    )
                {
                    return Ok(None);
                }
                if child_depth == 2
                    && is_drawingml(
                        &namespace,
                        element.name(),
                        b"tblPr",
                        &fragment_prefix,
                        allow_fragment_fallback,
                    )
                {
                    if properties.is_some() {
                        return Err(invalid("table has multiple tblPr elements"));
                    }
                    properties = Some(Properties {
                        first_row: on_off(&element, b"firstRow", decoder)?,
                        first_col: on_off(&element, b"firstCol", decoder)?,
                        last_row: on_off(&element, b"lastRow", decoder)?,
                        last_col: on_off(&element, b"lastCol", decoder)?,
                        band_row: on_off(&element, b"bandRow", decoder)?,
                        band_col: on_off(&element, b"bandCol", decoder)?,
                        style_id: None,
                    });
                }
            },
            Event::Text(text) if style_depth.is_some() => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                style.push_str(
                    &quick_xml::escape::unescape(&decoded)
                        .map_err(|error| Error::Xml(error.to_string()))?,
                );
            },
            Event::End(element) => {
                if style_depth == Some(depth)
                    && is_drawingml(
                        &namespace,
                        element.name(),
                        b"tableStyleId",
                        &fragment_prefix,
                        allow_fragment_fallback,
                    )
                {
                    let value = style.trim();
                    if value.is_empty() {
                        return Err(invalid("tableStyleId must not be empty"));
                    }
                    if let Some(properties) = properties.as_mut() {
                        properties.style_id = Some(value.to_owned());
                    }
                    style_depth = None;
                    style.clear();
                }
                if property_depth == Some(depth)
                    && is_drawingml(
                        &namespace,
                        element.name(),
                        b"tblPr",
                        &fragment_prefix,
                        allow_fragment_fallback,
                    )
                {
                    property_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("table XML nesting underflow"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("table XML contains forbidden markup"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || property_depth.is_some() || style_depth.is_some() {
        return Err(invalid("table XML is unterminated"));
    }
    Ok(properties)
}

fn on_off(
    element: &quick_xml::events::BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<bool> {
    match unqualified_attribute_value(element, name, decoder)? {
        None => Ok(false),
        Some(value) => match value.as_str() {
            "true" | "1" | "on" => Ok(true),
            "false" | "0" | "off" => Ok(false),
            _ => Err(invalid(format!("invalid table property value '{value}'"))),
        },
    }
}

fn extract_text(xml: &[u8], context: Option<NamespaceContext<'_>>) -> Result<String> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("table cell XML bytes", MAX_XML_BYTES));
    }
    let allow_fragment_fallback = context.is_none();
    let mut reader = NsReader::from_reader(xml);
    if let Some(context) = context {
        *reader.resolver_mut() = context.resolver()?;
    }
    let mut fragment_prefix = None;
    let mut depth = 0usize;
    let mut text_depth = None;
    // Depth of an open `a:tabLst`: its `a:tab` children are tab stops.
    let mut tab_list_depth = None;
    let mut result = String::new();
    let mut paragraph_seen = false;

    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                if fragment_prefix.is_none() && !matches!(namespace, ResolveResult::Bound(_)) {
                    fragment_prefix =
                        Some(element.name().prefix().map(|p| p.into_inner().to_vec()));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("table cell XML depth overflow"))?;
                if depth > MAX_DEPTH {
                    return Err(limit("table cell XML depth", MAX_DEPTH));
                }
                if is_drawingml(
                    &namespace,
                    element.name(),
                    b"p",
                    &fragment_prefix,
                    allow_fragment_fallback,
                ) {
                    if paragraph_seen && !result.is_empty() && !result.ends_with(' ') {
                        result.push(' ');
                    }
                    paragraph_seen = true;
                } else if text_depth.is_none()
                    && is_drawingml(
                        &namespace,
                        element.name(),
                        b"t",
                        &fragment_prefix,
                        allow_fragment_fallback,
                    )
                {
                    text_depth = Some(depth);
                } else if is_drawingml(
                    &namespace,
                    element.name(),
                    b"br",
                    &fragment_prefix,
                    allow_fragment_fallback,
                ) {
                    result.push('\n');
                } else if tab_list_depth.is_none()
                    && is_drawingml(
                        &namespace,
                        element.name(),
                        b"tabLst",
                        &fragment_prefix,
                        allow_fragment_fallback,
                    )
                {
                    tab_list_depth = Some(depth);
                } else if tab_list_depth.is_none()
                    && is_drawingml(
                        &namespace,
                        element.name(),
                        b"tab",
                        &fragment_prefix,
                        allow_fragment_fallback,
                    )
                {
                    result.push('\t');
                }
            },
            Event::Empty(element) => {
                if fragment_prefix.is_none() && !matches!(namespace, ResolveResult::Bound(_)) {
                    fragment_prefix =
                        Some(element.name().prefix().map(|p| p.into_inner().to_vec()));
                }
                if is_drawingml(
                    &namespace,
                    element.name(),
                    b"br",
                    &fragment_prefix,
                    allow_fragment_fallback,
                ) {
                    result.push('\n');
                } else if tab_list_depth.is_none()
                    && is_drawingml(
                        &namespace,
                        element.name(),
                        b"tab",
                        &fragment_prefix,
                        allow_fragment_fallback,
                    )
                {
                    // An `a:tab` of an `a:tabLst` defines a tab stop; only
                    // run content is a tab character.
                    result.push('\t');
                }
            },
            Event::Text(text) if text_depth.is_some() => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                result.push_str(
                    &quick_xml::escape::unescape(&decoded)
                        .map_err(|error| Error::Xml(error.to_string()))?,
                );
            },
            Event::CData(text) if text_depth.is_some() => {
                result.push_str(
                    &text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| Error::Xml(error.to_string()))?,
                );
            },
            Event::GeneralRef(reference) if text_depth.is_some() => {
                result.push_str(&decode_xml_reference(&reference)?);
            },
            Event::End(element) => {
                if text_depth == Some(depth)
                    && is_drawingml(
                        &namespace,
                        element.name(),
                        b"t",
                        &fragment_prefix,
                        allow_fragment_fallback,
                    )
                {
                    text_depth = None;
                }
                if tab_list_depth == Some(depth) {
                    tab_list_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("table cell XML nesting underflow"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("table cell XML contains forbidden markup"));
            },
            Event::Eof if depth != 0 || text_depth.is_some() => {
                return Err(invalid("table cell XML is unterminated"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(result.trim().to_owned())
}

/// Whether `haystack` holds `needle`.
fn contains(haystack: &[u8], needle: &[u8]) -> bool {
    haystack
        .windows(needle.len())
        .any(|window| window == needle)
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

    #[test]
    fn table_views_borrow_rows_and_cells() {
        let xml = format!(
            r#"<a:graphicFrame xmlns:a="{DML}"><a:tbl><a:tblPr firstRow="1"><a:tableStyleId>{{id}}</a:tableStyleId></a:tblPr><a:tr><a:tc><a:txBody><a:p><a:r><a:t>A &amp; B</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicFrame>"#
        );
        let table = Table::from_graphic_frame(xml.as_bytes()).unwrap();
        assert_eq!(table.row_count().unwrap(), 1);
        assert_eq!(table.column_count().unwrap(), 1);
        assert_eq!(
            table.rows().unwrap()[0].cells().unwrap()[0].text().unwrap(),
            "A & B"
        );
        assert!(table.properties().unwrap().unwrap().first_row);
    }

    #[test]
    fn cell_tab_stops_are_not_text() {
        let xml = format!(
            r#"<a:tbl xmlns:a="{DML}"><a:tr><a:tc><a:txBody><a:bodyPr/><a:p><a:pPr><a:tabLst><a:tab pos="914400" algn="l"/></a:tabLst></a:pPr><a:r><a:t>x</a:t></a:r><a:r><a:tab/><a:t>y</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl>"#
        );
        let table = Table::from_xml(xml.as_bytes()).unwrap();
        assert_eq!(
            table.rows().unwrap()[0].cells().unwrap()[0].text().unwrap(),
            "x\ty"
        );
    }

    #[test]
    fn inherited_prefixes_and_foreign_names_are_checked() {
        let xml = br#"<a:tbl xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:foreign"><x:tr/><a:tr><a:tc/></a:tr></a:tbl>"#;
        let table = Table::from_xml(xml).unwrap();
        assert_eq!(table.row_count().unwrap(), 1);
    }
}
