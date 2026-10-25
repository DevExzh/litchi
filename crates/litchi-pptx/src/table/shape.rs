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
            return Err(limit(
                crate::error::CodecLimitKind::TableRows,
                count,
                MAX_ROWS,
            ));
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
            return Err(limit(
                crate::error::CodecLimitKind::TableRows,
                rows.len(),
                MAX_ROWS,
            ));
        }
        let mut out = Vec::new();
        for row_span in rows {
            let row = row_span.get(table)?;
            let spans = scan_with_scope(row, b"tc", Some(scope.clone()), false)?;
            if spans.len() > MAX_CELLS {
                return Err(limit(
                    crate::error::CodecLimitKind::TableCells,
                    spans.len(),
                    MAX_CELLS,
                ));
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
            return Err(limit(
                crate::error::CodecLimitKind::TableRows,
                spans.len(),
                MAX_ROWS,
            ));
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
            return Err(limit(
                crate::error::CodecLimitKind::XmlBytes,
                table.len(),
                MAX_XML_BYTES,
            ));
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
                                return Err(limit(
                                    crate::error::CodecLimitKind::TableRows,
                                    counts.len().checked_add(1).unwrap_or(usize::MAX),
                                    MAX_ROWS,
                                ));
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

    /// The paragraphs and text runs of every cell (record 0348): what a
    /// batch planner needs to resolve cell, cell-paragraph and cell-run
    /// addresses, without reading the cells' properties and text. The
    /// index equals the one [`Self::cell_contents`] gives, from one pass
    /// over the table; a table holding markup this pass does not vouch for
    /// (markup compatibility, a nested table, a second text body in a cell,
    /// a fragment without its owner) is indexed from [`Self::cell_contents`]
    /// instead, with its errors.
    ///
    /// The pass reads structure only: an attribute value the text-body
    /// reader would refuse is not refused here.
    ///
    /// # Errors
    ///
    /// Returns an allocation failure, or the errors of
    /// [`Self::cell_contents`] when the table is indexed from it.
    pub fn cell_text_index(self) -> Result<CellTextIndex> {
        if let Some(index) = self.scan_cell_text_index()? {
            return Ok(index);
        }
        CellTextIndex::from_cells(self)
    }

    /// The one-pass [`Self::cell_text_index`], or `None` when the table
    /// holds markup it does not vouch for.
    pub(crate) fn scan_cell_text_index(self) -> Result<Option<CellTextIndex>> {
        let Some(context) = self.context else {
            return Ok(None);
        };
        scan_cell_text_fragment(self.xml()?, context)
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

/// The namespace of an element, as the cell text index tells them apart.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Namespaced {
    DrawingMl,
    MarkupCompatibility,
    Other,
}

impl Namespaced {
    fn of(namespace: &ResolveResult<'_>) -> Self {
        match namespace {
            ResolveResult::Bound(Namespace(value))
                if *value == DRAWINGML_NAMESPACE || *value == STRICT_DRAWINGML_NAMESPACE =>
            {
                Self::DrawingMl
            },
            ResolveResult::Bound(Namespace(value))
                if *value == litchi_ooxml_common::mce::NAMESPACE.as_bytes() =>
            {
                Self::MarkupCompatibility
            },
            _ => Self::Other,
        }
    }
}

/// What a cell text scan does after an event.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Step {
    Next,
    /// The table holds markup the scan does not vouch for.
    Unvouched,
}

/// The open elements of a cell text scan, as the index sees them.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Mark {
    /// An element of a frame around the table.
    Outside,
    Table,
    Row,
    Cell,
    Body,
    Paragraph,
    Other,
}

/// The state of one pass building a [`CellTextIndex`] (record 0348). It
/// takes the rows, cells, bodies, paragraphs and runs as
/// [`Table::cell_contents`] reads them, and gives up on markup that reader
/// treats otherwise.
#[derive(Default)]
struct CellTextScan {
    index: CellTextIndex,
    stack: Vec<Mark>,
    /// Whether the current cell has a `txBody` child already (the reader
    /// takes the first; a second is left to it).
    bodied: bool,
    /// The current paragraph's run children so far.
    children: usize,
    nodes: usize,
    /// Whether the fragment's table has started.
    tabled: bool,
    /// The byte span of each cell's text body in the fragment, by cell
    /// (record 0348): bodies are validated after the pass.
    bodies: Vec<(usize, std::ops::Range<usize>)>,
    /// Where the current cell's text body starts.
    body_start: usize,
    /// Whether the fragment is one text body the caller has read, scanned
    /// as one cell, whose body is not validated again.
    read: bool,
}

impl CellTextScan {
    /// A scan of one text body element, taken as the one cell of one row;
    /// its body is not validated (see [`scan_read_body_text_index`]).
    fn body() -> Result<Self> {
        let mut scan = Self {
            stack: vec![Mark::Table, Mark::Row, Mark::Cell],
            tabled: true,
            read: true,
            ..Self::default()
        };
        scan.index.open_row()?;
        scan.index.open_cell()?;
        Ok(scan)
    }

    fn start(
        &mut self,
        namespace: Namespaced,
        local: &[u8],
        empty: bool,
        open: usize,
    ) -> Result<Step> {
        self.nodes += 1;
        if self.nodes > MAX_NODES || namespace == Namespaced::MarkupCompatibility {
            return Ok(Step::Unvouched);
        }
        let dml = namespace == Namespaced::DrawingMl;
        let index = &mut self.index;
        let mark = match (self.stack.last().copied(), dml, local) {
            // The fragment's one table; a frame's other elements are outside.
            (None | Some(Mark::Outside), true, b"tbl") => {
                if self.tabled {
                    return Ok(Step::Unvouched);
                }
                self.tabled = true;
                Mark::Table
            },
            (Some(_), true, b"tbl") => return Ok(Step::Unvouched),
            (None | Some(Mark::Outside), ..) => Mark::Outside,
            (Some(Mark::Table), true, b"tr") => {
                index.open_row()?;
                Mark::Row
            },
            (Some(Mark::Row), true, b"tc") => {
                index.open_cell()?;
                self.bodied = false;
                Mark::Cell
            },
            (_, true, b"tr" | b"tc") => return Ok(Step::Unvouched),
            (Some(Mark::Cell), _, b"txBody") => {
                if self.bodied {
                    return Ok(Step::Unvouched);
                }
                self.bodied = true;
                if empty {
                    Mark::Other
                } else {
                    index.open_body();
                    self.body_start = open;
                    Mark::Body
                }
            },
            (Some(Mark::Body), true, b"p") => {
                index.open_paragraph()?;
                self.children = 0;
                Mark::Paragraph
            },
            (Some(Mark::Paragraph), true, b"r" | b"br" | b"fld") => {
                if local == b"r" {
                    index.push_text_run(self.children)?;
                }
                self.children += 1;
                Mark::Other
            },
            _ => Mark::Other,
        };
        if empty {
            return Ok(match mark {
                Mark::Paragraph => {
                    index.close_paragraph();
                    Step::Next
                },
                Mark::Cell => {
                    index.close_cell();
                    Step::Next
                },
                Mark::Row => {
                    index.close_row();
                    Step::Next
                },
                Mark::Table => Step::Unvouched,
                Mark::Body | Mark::Other | Mark::Outside => Step::Next,
            });
        }
        if self.stack.len() >= MAX_DEPTH {
            return Ok(Step::Unvouched);
        }
        self.stack.try_reserve(1).map_err(index_allocation)?;
        self.stack.push(mark);
        Ok(Step::Next)
    }

    fn end(&mut self, close: usize) -> Result<Step> {
        match self.stack.pop() {
            None => return Ok(Step::Unvouched),
            Some(Mark::Paragraph) => self.index.close_paragraph(),
            Some(Mark::Body) => {
                self.index.close_cell();
                self.bodies.try_reserve(1).map_err(index_allocation)?;
                let cell = self.index.cells.len().saturating_sub(1);
                self.bodies.push((cell, self.body_start..close));
            },
            Some(Mark::Cell) => self.index.close_cell(),
            Some(Mark::Row) => self.index.close_row(),
            Some(_) => {},
        }
        Ok(Step::Next)
    }

    fn finish(mut self) -> Option<CellTextIndex> {
        if self.read {
            // The body's own element closed; the cell it fills ends here.
            if self.stack != [Mark::Table, Mark::Row, Mark::Cell]
                || !matches!(self.index.cells.last(), Some(Some(_)))
            {
                return None;
            }
            self.stack.clear();
            self.index.close_cell();
            self.index.close_row();
        }
        (self.tabled
            && self.stack.is_empty()
            && self.index.rows.len() <= MAX_ROWS
            && self.index.cells.len() <= MAX_CELLS)
            .then_some(self.index)
    }
}

/// [`Table::cell_text_index`] of the one `DrawingML` table of the graphic
/// frame at `offset` of `owner`, from one pass over the frame (record
/// 0348); `None` when the frame holds markup the pass does not vouch for
/// (or no table, or more than one), which the caller then reads through
/// [`Table::from_graphic_frame_with_context`].
pub(crate) fn scan_frame_cell_text_index(
    frame: &[u8],
    owner: &[u8],
    offset: usize,
) -> Result<Option<CellTextIndex>> {
    scan_cell_text_fragment(frame, NamespaceContext { owner, offset })
}

/// The cell text index of the one `DrawingML` table in `xml` (the table
/// itself or a frame holding it), whose root inherits `context`'s scope.
fn scan_cell_text_fragment(
    xml: &[u8],
    context: NamespaceContext<'_>,
) -> Result<Option<CellTextIndex>> {
    if xml.len() > MAX_XML_BYTES || memchr::memmem::find(xml, b"AlternateContent").is_some() {
        return Ok(None);
    }
    scan_text_fragment(xml, context.resolver()?, CellTextScan::default())
}

/// The paragraphs and text runs of the text body element `xml` (a shape's
/// `p:txBody`), whose root inherits `scope`, as the one cell of one row,
/// from one pass (record 0348); `None` when the body holds markup the pass
/// does not vouch for. The body is not validated: the caller has read it
/// with the text-body reader.
pub(crate) fn scan_read_body_text_index(
    xml: &[u8],
    scope: NamespaceResolver,
) -> Result<Option<CellTextIndex>> {
    if xml.len() > MAX_XML_BYTES || memchr::memmem::find(xml, b"AlternateContent").is_some() {
        return Ok(None);
    }
    scan_text_fragment(xml, scope, CellTextScan::body()?)
}

/// One pass of `scan` over `xml`, whose root inherits `scope`.
fn scan_text_fragment(
    xml: &[u8],
    mut scope: NamespaceResolver,
    mut scan: CellTextScan,
) -> Result<Option<CellTextIndex>> {
    // Declarations on the root start tag only: one namespace scope for every
    // descendant, so each prefix is resolved once and the tags are found
    // without a namespace reader (the fragment is a span of a part the
    // scene reader has read whole).
    if let Some(root_end) = root_tag_end(xml)
        && memchr::memmem::find(&xml[root_end..], b"xmlns").is_none()
    {
        let mut reader = quick_xml::Reader::from_reader(&xml[..root_end]);
        let Ok(Event::Start(root)) = reader.read_event() else {
            return Ok(None);
        };
        if scope.push(&root).is_err() {
            return Ok(None);
        }
        let mut prefixes: Vec<(Vec<u8>, Namespaced)> = Vec::new();
        let classify = |name: &[u8]| -> Result<Namespaced> {
            let prefix = name
                .iter()
                .position(|byte| *byte == b':')
                .map_or(&b""[..], |colon| &name[..colon]);
            if let Some((_, namespaced)) = prefixes
                .iter()
                .find(|(known, _)| known.as_slice() == prefix)
            {
                return Ok(*namespaced);
            }
            let namespaced = Namespaced::of(&scope.resolve_element(QName(name)).0);
            prefixes.try_reserve(1).map_err(index_allocation)?;
            prefixes.push((prefix.to_vec(), namespaced));
            Ok(namespaced)
        };
        if scan_tags(xml, &mut scan, classify)? == Step::Unvouched {
            return Ok(None);
        }
        let bodies = if scan.read {
            Vec::new()
        } else {
            std::mem::take(&mut scan.bodies)
        };
        let Some(mut index) = scan.finish() else {
            return Ok(None);
        };
        // A body the text-body reader refuses is indexed as no body, as
        // `Cell::text_body` sees it fail (record 0348): its paragraph and
        // run addresses then take the item-by-item path and its refusal.
        // The reader's verdict depends on the body's markup, not on the
        // plain text of its runs, so bodies alike but for that text are
        // read once.
        let mut verdicts: std::collections::HashMap<Vec<u8>, bool> =
            std::collections::HashMap::new();
        let mut key = Vec::new();
        let mut dropped = false;
        for (cell, range) in bodies {
            let body = &xml[range];
            let readable = if body_key(body, &mut key) {
                match verdicts.get(key.as_slice()) {
                    Some(readable) => *readable,
                    None => {
                        let readable = body_readable(body, &scope);
                        verdicts.try_reserve(1).map_err(index_allocation)?;
                        verdicts.insert(key.clone(), readable);
                        readable
                    },
                }
            } else {
                body_readable(body, &scope)
            };
            if !readable {
                index.cells[cell] = None;
                dropped = true;
            }
        }
        return if dropped {
            index.compact().map(Some)
        } else {
            Ok(Some(index))
        };
    }
    // Declarations below the root: the cells are read one by one.
    Ok(None)
}

/// The end of the root start tag of `xml`, when the fragment begins with a
/// start tag (not an empty element).
fn root_tag_end(xml: &[u8]) -> Option<usize> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    match reader.read_event() {
        Ok(Event::Start(_)) => usize::try_from(reader.buffer_position()).ok(),
        _ => None,
    }
}

/// Feed the elements of a well-formed fragment without namespace
/// declarations to `scan`: start and empty tags with their namespace (by
/// `classify`, from the qualified name) and local name, end tags checked
/// against their start. Comments and character data sections are skipped;
/// a processing instruction, a declaration or a malformed tag is
/// [`Step::Unvouched`].
fn scan_tags(
    xml: &[u8],
    scan: &mut CellTextScan,
    mut classify: impl FnMut(&[u8]) -> Result<Namespaced>,
) -> Result<Step> {
    // The qualified names of the open elements.
    let mut names: Vec<std::ops::Range<usize>> = Vec::new();
    let mut at = 0;
    while let Some(found) = memchr::memchr(b'<', &xml[at..]) {
        let open = at + found;
        let rest = &xml[open + 1..];
        match rest.first() {
            Some(b'/') => {
                let Some(close) = memchr::memchr(b'>', rest) else {
                    return Ok(Step::Unvouched);
                };
                let name = rest[1..close].trim_ascii_end();
                match names.pop() {
                    Some(range) if xml[range.clone()] == *name => {},
                    _ => return Ok(Step::Unvouched),
                }
                if scan.end(open + 1 + close + 1)? == Step::Unvouched {
                    return Ok(Step::Unvouched);
                }
                at = open + 1 + close + 1;
            },
            Some(b'!') => {
                let (start, terminator): (usize, &[u8]) = if rest.starts_with(b"!--") {
                    (open + 4, b"-->")
                } else if rest.starts_with(b"![CDATA[") {
                    (open + 9, b"]]>")
                } else {
                    return Ok(Step::Unvouched);
                };
                let Some(end) = memchr::memmem::find(&xml[start..], terminator) else {
                    return Ok(Step::Unvouched);
                };
                at = start + end + terminator.len();
            },
            Some(byte) if !matches!(byte, b'?' | b' ' | b'\t' | b'\r' | b'\n' | b'>') => {
                let Some(length) = rest
                    .iter()
                    .position(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n' | b'/' | b'>'))
                else {
                    return Ok(Step::Unvouched);
                };
                let name = open + 1..open + 1 + length;
                // The tag's `>`, past quoted attribute values.
                let mut cursor = name.end;
                let close = loop {
                    let Some(found) = memchr::memchr3(b'>', b'"', b'\'', &xml[cursor..]) else {
                        return Ok(Step::Unvouched);
                    };
                    let position = cursor + found;
                    let quote = xml[position];
                    if quote == b'>' {
                        break position;
                    }
                    let Some(end) = memchr::memchr(quote, &xml[position + 1..]) else {
                        return Ok(Step::Unvouched);
                    };
                    cursor = position + 1 + end + 1;
                };
                let empty = xml[close - 1] == b'/';
                let qualified = &xml[name.clone()];
                let local = qualified
                    .iter()
                    .position(|byte| *byte == b':')
                    .map_or(qualified, |colon| &qualified[colon + 1..]);
                if scan.start(classify(qualified)?, local, empty, open)? == Step::Unvouched {
                    return Ok(Step::Unvouched);
                }
                if !empty {
                    names.try_reserve(1).map_err(index_allocation)?;
                    names.push(name);
                }
                at = close + 1;
            },
            _ => return Ok(Step::Unvouched),
        }
    }
    Ok(if names.is_empty() {
        Step::Next
    } else {
        Step::Unvouched
    })
}

/// Whether the shared text-body reader accepts the body element `body` in
/// `scope`, as [`Cell::text_body`] reads it (record 0348).
fn body_readable(body: &[u8], scope: &NamespaceResolver) -> bool {
    litchi_drawingml::text::body::reader::read(
        body,
        scope.clone(),
        litchi_drawingml::text::body::reader::Limits::DEFAULT,
    )
    .is_ok()
}

/// The text-body reader's view of `body` with the plain character data
/// of its elements reduced to its class, into `key` (record 0348): markup
/// and attribute values are kept byte for byte; a run of character data
/// without references, control characters or more than 1 KiB, valid
/// UTF-8, becomes one byte telling whitespace from text. `false` when the
/// body is not keyed so (comments, character data sections, a large body):
/// it is then read on its own.
fn body_key(body: &[u8], key: &mut Vec<u8>) -> bool {
    const LARGEST: usize = 64 * 1024;
    key.clear();
    if body.len() > LARGEST || memchr::memmem::find(body, b"<!").is_some() {
        return false;
    }
    let mut at = 0;
    while at < body.len() {
        let next = memchr::memchr(b'<', &body[at..]).map_or(body.len(), |found| at + found);
        let text = &body[at..next];
        if !text.is_empty() {
            let plain = text.len() <= 1024
                && !text.contains(&b'&')
                && text
                    .iter()
                    .all(|byte| *byte >= 0x20 || matches!(byte, b'\t' | b'\n' | b'\r'))
                && std::str::from_utf8(text).is_ok();
            if plain {
                let blank = text
                    .iter()
                    .all(|byte| matches!(byte, b' ' | b'\t' | b'\n' | b'\r'));
                key.push(if blank { 0x01 } else { 0x02 });
            } else {
                key.push(0x03);
                key.extend_from_slice(text);
                key.push(0x03);
            }
        }
        if next == body.len() {
            break;
        }
        // The tag, past quoted attribute values.
        let mut cursor = next + 1;
        let close = loop {
            let Some(found) = memchr::memchr3(b'>', b'"', b'\'', &body[cursor..]) else {
                return false;
            };
            let position = cursor + found;
            let quote = body[position];
            if quote == b'>' {
                break position;
            }
            let Some(end) = memchr::memchr(quote, &body[position + 1..]) else {
                return false;
            };
            cursor = position + 1 + end + 1;
        };
        key.extend_from_slice(&body[next..=close]);
        at = close + 1;
    }
    true
}

fn index_allocation(source: std::collections::TryReserveError) -> Error {
    Error::Allocation {
        resource: "table cell text index",
        source,
    }
}

/// The paragraphs and text runs of a table's cells, by row and position in
/// the row (record 0348); see [`Table::cell_text_index`].
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct CellTextIndex {
    /// Each row's cells, as a range of `cells`.
    rows: Vec<std::ops::Range<usize>>,
    /// Each cell's paragraphs, as a range of `paragraphs`; `None` for a
    /// cell without a text body.
    cells: Vec<Option<std::ops::Range<usize>>>,
    /// Each paragraph's text runs, as a range of `runs`.
    paragraphs: Vec<std::ops::Range<usize>>,
    /// The positions of text runs (`a:r`) among their paragraph's run
    /// children (`a:r`, `a:br`, `a:fld`), as
    /// [`crate::opened::TextTarget::CellRun`] counts them.
    runs: Vec<usize>,
}

impl CellTextIndex {
    /// The number of rows.
    #[must_use]
    pub fn row_count(&self) -> usize {
        self.rows.len()
    }

    /// The number of cells of row `row`; `None` past the last row.
    #[must_use]
    pub fn cell_count(&self, row: usize) -> Option<usize> {
        self.rows.get(row).map(ExactSizeIterator::len)
    }

    /// Whether the table has a cell at `row` and `column`.
    #[must_use]
    pub fn has_cell(&self, row: usize, column: usize) -> bool {
        self.cell(row, column).is_some()
    }

    /// The number of paragraphs of the cell at `row` and `column`; `None`
    /// for a missing cell or a cell without a text body.
    #[must_use]
    pub fn paragraph_count(&self, row: usize, column: usize) -> Option<usize> {
        self.cell(row, column)?.as_ref().map(ExactSizeIterator::len)
    }

    /// The position among its paragraph's run children of the
    /// `ordinal`-th text run (zero-based) of paragraph `paragraph` of the
    /// cell at `row` and `column`; `None` when there is no such run.
    #[must_use]
    pub fn text_run(
        &self,
        row: usize,
        column: usize,
        paragraph: usize,
        ordinal: usize,
    ) -> Option<usize> {
        let paragraphs = self.cell(row, column)?.as_ref()?;
        let runs = self.paragraphs.get(paragraphs.clone())?.get(paragraph)?;
        self.runs.get(runs.clone())?.get(ordinal).copied()
    }

    fn cell(&self, row: usize, column: usize) -> Option<&Option<std::ops::Range<usize>>> {
        let cells = self.rows.get(row)?;
        self.cells.get(cells.clone())?.get(column)
    }

    /// The index of one text body as the one cell of one row (record
    /// 0348).
    pub(crate) fn from_body(body: &litchi_drawingml::text::body::Body) -> Result<Self> {
        let mut index = Self::default();
        index.open_row()?;
        index.open_cell()?;
        index.open_body();
        for paragraph in &body.paragraphs {
            index.open_paragraph()?;
            for (position, run) in paragraph.runs.iter().enumerate() {
                if run.is_text() {
                    index.push_text_run(position)?;
                }
            }
            index.close_paragraph();
        }
        index.close_cell();
        index.close_row();
        Ok(index)
    }

    /// The index of the cells read one by one: a cell whose text body
    /// [`Cell::text_body`] refuses is indexed as one without a body (record
    /// 0348).
    pub(crate) fn from_cells(table: Table<'_>) -> Result<Self> {
        let mut index = Self::default();
        for row in table.rows()? {
            index.open_row()?;
            for cell in row.cells()? {
                index.open_cell()?;
                if let Ok(Some(body)) = cell.text_body() {
                    index.open_body();
                    for paragraph in &body.paragraphs {
                        index.open_paragraph()?;
                        for (position, run) in paragraph.runs.iter().enumerate() {
                            if run.is_text() {
                                index.push_text_run(position)?;
                            }
                        }
                        index.close_paragraph();
                    }
                }
                index.close_cell();
            }
            index.close_row();
        }
        Ok(index)
    }

    /// The index of contents [`Table::cell_contents`] read.
    #[cfg(test)]
    pub(crate) fn from_contents(
        rows: &[Vec<(CellProperties, Option<litchi_drawingml::text::body::Body>)>],
    ) -> Result<Self> {
        let mut index = Self::default();
        for row in rows {
            index.open_row()?;
            for (_, body) in row {
                index.open_cell()?;
                if let Some(body) = body {
                    index.open_body();
                    for paragraph in &body.paragraphs {
                        index.open_paragraph()?;
                        for (position, run) in paragraph.runs.iter().enumerate() {
                            if run.is_text() {
                                index.push_text_run(position)?;
                            }
                        }
                        index.close_paragraph();
                    }
                }
                index.close_cell();
            }
            index.close_row();
        }
        Ok(index)
    }

    /// The index without the paragraphs of dropped bodies.
    fn compact(&self) -> Result<Self> {
        let mut index = Self::default();
        for cells in &self.rows {
            index.open_row()?;
            for cell in &self.cells[cells.clone()] {
                index.open_cell()?;
                if let Some(paragraphs) = cell {
                    index.open_body();
                    for runs in &self.paragraphs[paragraphs.clone()] {
                        index.open_paragraph()?;
                        for position in &self.runs[runs.clone()] {
                            index.push_text_run(*position)?;
                        }
                        index.close_paragraph();
                    }
                }
                index.close_cell();
            }
            index.close_row();
        }
        Ok(index)
    }

    fn grow<T>(items: &mut Vec<T>) -> Result<()> {
        items.try_reserve(1).map_err(index_allocation)
    }

    fn open_row(&mut self) -> Result<()> {
        Self::grow(&mut self.rows)?;
        self.rows.push(self.cells.len()..self.cells.len());
        Ok(())
    }

    fn close_row(&mut self) {
        let end = self.cells.len();
        if let Some(row) = self.rows.last_mut() {
            row.end = end;
        }
    }

    fn open_cell(&mut self) -> Result<()> {
        Self::grow(&mut self.cells)?;
        self.cells.push(None);
        Ok(())
    }

    fn open_body(&mut self) {
        let start = self.paragraphs.len();
        if let Some(cell) = self.cells.last_mut() {
            *cell = Some(start..start);
        }
    }

    fn close_cell(&mut self) {
        let end = self.paragraphs.len();
        if let Some(Some(cell)) = self.cells.last_mut() {
            cell.end = end;
        }
    }

    fn open_paragraph(&mut self) -> Result<()> {
        Self::grow(&mut self.paragraphs)?;
        self.paragraphs.push(self.runs.len()..self.runs.len());
        Ok(())
    }

    fn close_paragraph(&mut self) {
        let end = self.runs.len();
        if let Some(paragraph) = self.paragraphs.last_mut() {
            paragraph.end = end;
        }
    }

    fn push_text_run(&mut self, position: usize) -> Result<()> {
        Self::grow(&mut self.runs)?;
        self.runs.push(position);
        Ok(())
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
            return Err(limit(
                crate::error::CodecLimitKind::TableCells,
                spans.len(),
                MAX_CELLS,
            ));
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
        let body = litchi_drawingml::text::body::reader::read(
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

fn limit(kind: crate::error::CodecLimitKind, observed: usize, maximum: usize) -> Error {
    Error::Limit {
        kind: crate::error::LimitKind::Codec(kind),
        observed,
        maximum,
    }
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
        return Err(limit(
            crate::error::CodecLimitKind::XmlBytes,
            xml.len(),
            MAX_XML_BYTES,
        ));
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
            return Err(limit(
                crate::error::CodecLimitKind::XmlNodes,
                nodes,
                MAX_NODES,
            ));
        }

        match event {
            ScanEvent::Start { local, matched } => {
                let depth = stack
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| invalid("table XML depth overflow"))?;
                if depth > MAX_DEPTH {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        depth,
                        MAX_DEPTH,
                    ));
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
        return Err(limit(
            crate::error::CodecLimitKind::XmlBytes,
            xml.len(),
            MAX_XML_BYTES,
        ));
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
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        depth,
                        MAX_DEPTH,
                    ));
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

    /// The one-pass cell text index of `owner` (a graphic frame or table
    /// whose root declares the namespaces), next to the index of the full
    /// cell reader.
    fn indexes(owner: &str) -> (Option<CellTextIndex>, CellTextIndex) {
        let table =
            Table::from_graphic_frame_with_context(owner.as_bytes(), owner.as_bytes(), 0).unwrap();
        let scanned = table.scan_cell_text_index().unwrap();
        let read = CellTextIndex::from_cells(table).unwrap();
        if let Ok(contents) = table.cell_contents() {
            assert_eq!(CellTextIndex::from_contents(&contents).unwrap(), read);
        }
        assert_eq!(table.cell_text_index().unwrap(), read);
        (scanned, read)
    }

    #[test]
    fn cell_text_index_equals_the_cell_reader() {
        // Record 0348: runs, breaks and fields; empty paragraphs; cells
        // without a body or with an empty one; foreign and unmodeled
        // children; a cell after an empty row.
        let field = r#"<a:fld id="{5C2F1A2B-0000-4000-8000-000000000001}" type="slidenum"><a:t>1</a:t></a:fld>"#;
        let rows = [
            format!(
                r#"<a:tr h="1"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>a</a:t></a:r><a:br/><a:r><a:rPr b="1"/><a:t>b</a:t></a:r>{field}<a:r><a:t>c</a:t></a:r><a:endParaRPr/></a:p><a:p/><a:p><a:pPr algn="ctr"/></a:p></a:txBody><a:tcPr/></a:tc><a:tc/><a:tc><a:txBody/></a:tc></a:tr>"#
            ),
            "<a:tr/>".to_owned(),
            // Comments, character data sections and quoted `>` and `/`.
            r#"<a:tr><!-- <a:tc> --><a:tc ><a:txBody><a:p><a:r><a:rPr altLang="a>b/" lang='x"y'/><a:t><![CDATA[<a:r/>]]></a:t></a:r ><a:r/></a:p></a:txBody></a:tc></a:tr>"#.to_owned(),
            r#"<a:tr><a:tc><a:txBody><a:p><x:note/><a:r><a:t>x</a:t></a:r><a:unknown/><a:r/></a:p><x:p><a:r/></x:p></a:txBody><a:tcPr><a:solidFill><a:srgbClr val="FFEECC"/></a:solidFill></a:tcPr></a:tc><a:tc><x:txBody><a:p><a:r><a:t>y</a:t></a:r></a:p></x:txBody></a:tc></a:tr>"#.to_owned(),
        ];
        for namespace in [DML, "http://purl.oclc.org/ooxml/drawingml/main"] {
            let owner = format!(
                r#"<p:graphicFrame xmlns:p="urn:p" xmlns:a="{namespace}" xmlns:x="urn:foreign"><a:graphic><a:graphicData><a:tbl><a:tblGrid><a:gridCol w="1"/></a:tblGrid>{}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#,
                rows.concat()
            );
            let (scanned, read) = indexes(&owner);
            assert_eq!(scanned.as_ref(), Some(&read));
            assert_eq!(read.row_count(), 4);
            assert_eq!(read.cell_count(1), Some(0));
            assert_eq!(read.text_run(2, 0, 0, 1), Some(1));
            assert_eq!(read.paragraph_count(0, 0), Some(3));
            assert_eq!(read.paragraph_count(0, 1), None);
            assert_eq!(read.paragraph_count(0, 2), None);
            assert!(read.has_cell(0, 2) && !read.has_cell(0, 3) && !read.has_cell(4, 0));
            // `a` at 0, `b` at 2, `c` at 4 (the break and field count).
            assert_eq!(read.text_run(0, 0, 0, 1), Some(2));
            assert_eq!(read.text_run(0, 0, 0, 2), Some(4));
            assert_eq!(read.text_run(0, 0, 0, 3), None);
            assert_eq!(read.text_run(0, 0, 1, 0), None);
            assert_eq!(read.paragraph_count(3, 0), Some(1));
            assert_eq!(read.text_run(3, 0, 0, 1), Some(1));
            assert_eq!(read.paragraph_count(3, 1), Some(1));
        }
        // A table declaring namespaces below its root: read cell by cell.
        let owner = format!(
            r#"<a:tbl xmlns:a="{DML}"><a:tr><a:tc><a:txBody xmlns:b="{DML}" xmlns:a="urn:other"><b:p><b:r><b:t>z</b:t></b:r><a:r/><b:r/></b:p></a:txBody></a:tc></a:tr></a:tbl>"#
        );
        let (scanned, read) = indexes(&owner);
        assert_eq!(scanned, None);
        assert_eq!(read.text_run(0, 0, 0, 1), Some(1));
    }

    #[test]
    fn cell_text_index_drops_the_bodies_the_reader_refuses() {
        // Record 0348: a body the text-body reader refuses is no body, as
        // `Cell::text_body` sees it fail; bodies alike but for their text
        // share one reading, others are read on their own.
        let cell =
            |content: &str| format!("<a:tc><a:txBody><a:p>{content}</a:p></a:txBody></a:tc>");
        let cells = [
            cell("<a:r><a:t>fine</a:t></a:r>"),
            cell(r#"<a:pPr algn="bogus"/><a:r><a:t>a</a:t></a:r>"#),
            cell(r#"<a:r><a:rPr sz="abc"/><a:t>b</a:t></a:r>"#),
            cell("<a:r><a:rPr/><a:rPr/><a:t>c</a:t></a:r>"),
            cell(r#"<a:r><a:rPr sz="1800"/><a:t>d</a:t></a:r>"#),
            cell(r#"<a:r><a:rPr sz="1800"/><a:t>another text</a:t></a:r>"#),
            cell(r#"<a:r><a:rPr sz="1800"/><a:t>x &amp; y</a:t></a:r>"#),
            cell(r#"<a:r><a:rPr sz="1800"/><a:t>x &bogus; y</a:t></a:r>"#),
            cell(r#"<a:r><a:rPr altLang="a>b" sz="1800"/><a:t>e</a:t></a:r>"#),
            cell(r#"<a:r><a:rPr altLang="a>b" sz="18x0"/><a:t>e</a:t></a:r>"#),
            cell("<!-- c --><a:r><a:t>f</a:t></a:r>"),
            cell("stray text<a:r><a:t>g</a:t></a:r>"),
            cell("<a:r><a:t>h</a:t></a:r>"),
        ];
        let owner = format!(
            r#"<a:tbl xmlns:a="{DML}"><a:tr>{}</a:tr></a:tbl>"#,
            cells.concat()
        );
        let table =
            Table::from_graphic_frame_with_context(owner.as_bytes(), owner.as_bytes(), 0).unwrap();
        let read = CellTextIndex::from_cells(table).unwrap();
        let rows = table.rows().unwrap();
        let expected: Vec<bool> = rows[0]
            .cells()
            .unwrap()
            .into_iter()
            .map(|cell| matches!(cell.text_body(), Ok(Some(_))))
            .collect();
        assert_eq!(
            expected.iter().filter(|ok| !**ok).count(),
            5,
            "{expected:?}"
        );
        for (column, ok) in expected.iter().enumerate() {
            assert_eq!(read.paragraph_count(0, column).is_some(), *ok, "{column}");
        }
        assert_eq!(table.scan_cell_text_index().unwrap().as_ref(), Some(&read));
        assert_eq!(table.cell_text_index().unwrap(), read);
    }

    #[test]
    fn cell_text_index_defers_markup_it_does_not_vouch_for() {
        let cell = r#"<a:tc><a:txBody><a:p><a:r><a:t>v</a:t></a:r></a:p></a:txBody></a:tc>"#;
        let frame = |rows: &str| {
            format!(
                r#"<p:graphicFrame xmlns:p="urn:p" xmlns:a="{DML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><a:graphic><a:graphicData><a:tbl>{rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#
            )
        };
        for rows in [
            // Markup compatibility.
            r#"<a:tr><a:tc><a:txBody><a:p><mc:AlternateContent><mc:Fallback><a:r><a:t>x</a:t></a:r></mc:Fallback></mc:AlternateContent></a:p></a:txBody></a:tc></a:tr>"#
                .to_owned(),
            // A second text body.
            r#"<a:tr><a:tc><a:txBody><a:p/></a:txBody><a:txBody><a:p/></a:txBody></a:tc></a:tr>"#
                .to_owned(),
            // A row outside the table's own rows.
            format!(r#"<a:tr>{cell}<a:extLst><a:tr/></a:extLst></a:tr>"#),
        ] {
            let owner = frame(&rows);
            let table =
                Table::from_graphic_frame_with_context(owner.as_bytes(), owner.as_bytes(), 0)
                    .unwrap();
            assert_eq!(table.scan_cell_text_index().unwrap(), None, "{rows}");
            assert_eq!(
                table.cell_text_index().unwrap(),
                CellTextIndex::from_cells(table).unwrap()
            );
        }
        // A fragment without its owner's scope.
        let owner = frame(&format!("<a:tr>{cell}</a:tr>"));
        let table = Table::from_graphic_frame(owner.as_bytes()).unwrap();
        assert_eq!(table.scan_cell_text_index().unwrap(), None);
        assert_eq!(
            table.cell_text_index().unwrap().paragraph_count(0, 0),
            Some(1)
        );
    }

    #[test]
    fn inherited_prefixes_and_foreign_names_are_checked() {
        let xml = br#"<a:tbl xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:foreign"><x:tr/><a:tr><a:tc/></a:tr></a:tbl>"#;
        let table = Table::from_xml(xml).unwrap();
        assert_eq!(table.row_count().unwrap(), 1);
    }
}
