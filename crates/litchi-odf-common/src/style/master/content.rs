//! Typed, ordered content extracted from ODF master-page regions.

use std::collections::HashMap;

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesRef, BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

use super::region::Kind;

const STYLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const NUMBER_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";

const MAX_XML_DEPTH: usize = 128;
const MAX_BLOCKS: usize = 4_096;
const MAX_INLINE_TOKENS: usize = 65_536;
const MAX_FIELDS: usize = 16_384;
const MAX_FIELD_ATTRIBUTES: usize = 128;
pub(crate) const MAX_EXPANDED_SPACES: usize = 1_000_000;

/// The ODF 1.3 `style:region-*` column region wrapping a header/footer block.
///
/// Multi-column headers and footers replace their plain paragraphs with
/// `style:region-left`, `style:region-center`, and `style:region-right`
/// wrappers, each containing the paragraphs of that column.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Column {
    Left,
    Center,
    Right,
}

impl Column {
    fn parse(local_name: &[u8]) -> Option<Self> {
        match local_name {
            b"region-left" => Some(Self::Left),
            b"region-center" => Some(Self::Center),
            b"region-right" => Some(Self::Right),
            _ => None,
        }
    }

    /// Normative document order of the region wrappers.
    fn order(self) -> u8 {
        match self {
            Self::Left => 0,
            Self::Center => 1,
            Self::Right => 2,
        }
    }
}

/// One paragraph or heading in a header/footer region.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Block {
    /// The paragraph's `text:style-name`, when present.
    pub style_name: Option<String>,
    /// The column region containing this block, for multi-column headers/footers.
    pub column_region: Option<Column>,
    /// Inline content in document order.
    pub content: Vec<Inline>,
}

/// Ordered inline header/footer content.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Inline {
    Text(String),
    Space { count: usize },
    Tab,
    LineBreak,
    Field(Field),
}

/// A dynamic ODF text field without evaluating its value.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Field {
    pub kind: FieldKind,
    /// Cached/displayed child text stored in the document.
    pub displayed_text: String,
    pub fixed: Option<bool>,
    pub data_style_name: Option<String>,
    /// Attributes in source order with canonical namespace names.
    pub attributes: Vec<(String, String)>,
}

/// Supported field semantics. Fields are read as cached document metadata only.
///
/// The parser never evaluates formulas, resolves DDE connections, loads external
/// content, or invokes macros. Unknown text-namespace fields remain typed and
/// lossless.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum FieldKind {
    PageNumber,
    PageCount,
    PageContinuation,
    PageVariableSet,
    PageVariableGet,
    ParagraphCount,
    WordCount,
    CharacterCount,
    TableCount,
    ImageCount,
    ObjectCount,
    Reference,
    SequenceReference,
    BookmarkReference,
    NoteReference,
    VariableSet,
    VariableGet,
    VariableInput,
    UserFieldGet,
    UserFieldInput,
    Sequence,
    Expression,
    TextInput,
    Placeholder,
    ConditionalText,
    HiddenText,
    HiddenParagraph,
    /// Cached DDE field metadata. The named connection is never opened or refreshed.
    DdeConnection,
    Measure,
    TableFormula,
    MetaField,
    /// Database field metadata. Its source is never opened or queried.
    DatabaseDisplay,
    /// Database field metadata. Its source is never opened or queried.
    DatabaseNext,
    /// Database field metadata. Its source is never opened or queried.
    DatabaseRowSelect,
    /// Database field metadata. Its source is never opened or queried.
    DatabaseRowNumber,
    /// Database field metadata. Its source is never opened or queried.
    DatabaseName,
    Title,
    Subject,
    AuthorName,
    AuthorInitials,
    Date,
    Time,
    FileName,
    TemplateName,
    SheetName,
    Chapter,
    InitialCreator,
    Description,
    PrintedBy,
    Keywords,
    Creator,
    CreationDate,
    CreationTime,
    ModificationDate,
    ModificationTime,
    PrintDate,
    PrintTime,
    EditingCycles,
    EditingDuration,
    UserDefined,
    /// A sender identity or contact field defined by ODF's `text:sender-*` elements.
    Sender(SenderKind),
    /// Inert `text:script` metadata. Its URI and payload are never opened or executed.
    Script,
    /// Inert `text:execute-macro` metadata. Its named macro is never invoked.
    ExecuteMacro,
    Unknown {
        namespace: String,
        local_name: String,
    },
}

/// One of the ODF sender identity/contact fields available in header/footer content.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SenderKind {
    FirstName,
    LastName,
    Initials,
    Title,
    Position,
    Email,
    PrivatePhone,
    Fax,
    Company,
    WorkPhone,
    Street,
    City,
    PostalCode,
    Country,
    StateOrProvince,
}

impl SenderKind {
    /// The local name of the corresponding ODF `text:sender-*` element.
    #[must_use]
    pub const fn element_name(self) -> &'static str {
        match self {
            Self::FirstName => "sender-firstname",
            Self::LastName => "sender-lastname",
            Self::Initials => "sender-initials",
            Self::Title => "sender-title",
            Self::Position => "sender-position",
            Self::Email => "sender-email",
            Self::PrivatePhone => "sender-phone-private",
            Self::Fax => "sender-fax",
            Self::Company => "sender-company",
            Self::WorkPhone => "sender-phone-work",
            Self::Street => "sender-street",
            Self::City => "sender-city",
            Self::PostalCode => "sender-postal-code",
            Self::Country => "sender-country",
            Self::StateOrProvince => "sender-state-or-province",
        }
    }
}

struct Master {
    name: String,
    depth: usize,
}

struct Region {
    master_name: String,
    kind: Kind,
    depth: usize,
    blocks: Vec<Block>,
    block: Option<ActiveBlock>,
    field: Option<ActiveField>,
    active_column: Option<ActiveColumn>,
    last_column_region_order: Option<u8>,
    has_plain_blocks: bool,
    token_count: usize,
    field_count: usize,
    expanded_spaces: usize,
}

struct ActiveBlock {
    depth: usize,
    block: Block,
}

struct ActiveColumn {
    kind: Column,
    depth: usize,
}

struct ActiveField {
    depth: usize,
    field: Field,
}

impl Region {
    fn new(master_name: String, kind: Kind, depth: usize) -> Self {
        Self {
            master_name,
            kind,
            depth,
            blocks: Vec::new(),
            block: None,
            field: None,
            active_column: None,
            last_column_region_order: None,
            has_plain_blocks: false,
            token_count: 0,
            field_count: 0,
            expanded_spaces: 0,
        }
    }

    fn start_element(
        &mut self,
        reader: &NsReader<&[u8]>,
        namespace: Option<&[u8]>,
        element: &BytesStart<'_>,
        depth: usize,
        empty: bool,
    ) -> Result<()> {
        let local = element.local_name();
        let starts_column_region = self.block.is_none() && namespace == Some(STYLE_NAMESPACE);
        if let (true, Some(kind)) = (starts_column_region, Column::parse(local.as_ref())) {
            self.start_column_region(kind, depth, empty)?;
            return Ok(());
        }
        if namespace == Some(TEXT_NAMESPACE) && matches!(local.as_ref(), b"p" | b"h") {
            if self.block.is_some() {
                return Err(Error::InvalidFormat(
                    "nested header/footer paragraph or heading".to_string(),
                ));
            }
            if local.as_ref() == b"h" && self.active_column.is_some() {
                return Err(Error::InvalidFormat(
                    "style:region-* column regions may contain only text:p".to_string(),
                ));
            }
            let column_region = self.active_column.as_ref().map(|region| region.kind);
            if column_region.is_none() {
                if self.last_column_region_order.is_some() {
                    return Err(Error::InvalidFormat(
                        "header/footer mixes plain blocks with style:region-* column regions"
                            .to_string(),
                    ));
                }
                self.has_plain_blocks = true;
            }
            if self.blocks.len() >= MAX_BLOCKS {
                return Err(limit_error("block"));
            }
            self.block = Some(ActiveBlock {
                depth,
                block: Block {
                    style_name: namespaced_attr(reader, element, TEXT_NAMESPACE, b"style-name")?,
                    column_region,
                    content: Vec::new(),
                },
            });
            if empty {
                self.finish_block()?;
            }
            return Ok(());
        }
        if self.block.is_none() {
            return Ok(());
        }
        if self.field.is_some() {
            if namespace == Some(TEXT_NAMESPACE) {
                self.append_field_control(reader, element)?;
            }
            return Ok(());
        }
        if namespace == Some(TEXT_NAMESPACE) {
            match local.as_ref() {
                b"s" => {
                    let count = text_space_count(reader, element)?.unwrap_or(1);
                    self.add_spaces(count)?;
                    self.push_token(Inline::Space { count })?;
                },
                b"tab" => self.push_token(Inline::Tab)?,
                b"line-break" => self.push_token(Inline::LineBreak)?,
                field_local
                    if field_kind(field_local).is_some() || is_unknown_field(field_local) =>
                {
                    self.field_count = self
                        .field_count
                        .checked_add(1)
                        .ok_or_else(|| limit_error("field"))?;
                    if self.field_count > MAX_FIELDS {
                        return Err(limit_error("field"));
                    }
                    let field = parse_field(reader, element, field_local)?;
                    if empty {
                        self.push_token(Inline::Field(field))?;
                    } else {
                        self.field = Some(ActiveField { depth, field });
                    }
                },
                _ => {},
            }
        }
        Ok(())
    }

    fn start_column_region(&mut self, kind: Column, depth: usize, empty: bool) -> Result<()> {
        if self.active_column.is_some() {
            return Err(Error::InvalidFormat(
                "nested style:region-* column regions".to_string(),
            ));
        }
        if self.has_plain_blocks {
            return Err(Error::InvalidFormat(
                "header/footer mixes plain blocks with style:region-* column regions".to_string(),
            ));
        }
        if self
            .last_column_region_order
            .is_some_and(|order| kind.order() <= order)
        {
            return Err(Error::InvalidFormat(
                "style:region-* column regions are duplicated or out of order".to_string(),
            ));
        }
        self.last_column_region_order = Some(kind.order());
        if !empty {
            self.active_column = Some(ActiveColumn { kind, depth });
        }
        Ok(())
    }

    fn end_element(&mut self, namespace: Option<&[u8]>, local: &[u8], depth: usize) -> Result<()> {
        if self
            .field
            .as_ref()
            .is_some_and(|field| field.depth == depth)
        {
            let field = self.field.take().ok_or_else(|| {
                Error::InvalidFormat("missing active header/footer field".to_string())
            })?;
            self.push_token(Inline::Field(field.field))?;
        }
        if self
            .block
            .as_ref()
            .is_some_and(|block| block.depth == depth)
        {
            if namespace != Some(TEXT_NAMESPACE) || !matches!(local, b"p" | b"h") {
                return Err(Error::InvalidFormat(
                    "malformed header/footer block nesting".to_string(),
                ));
            }
            if self.field.is_some() {
                return Err(Error::InvalidFormat(
                    "unterminated header/footer field".to_string(),
                ));
            }
            self.finish_block()?;
        }
        if self
            .active_column
            .as_ref()
            .is_some_and(|region| region.depth == depth)
        {
            let column_kind = self
                .active_column
                .as_ref()
                .map(|region| region.kind)
                .ok_or_else(|| {
                    Error::InvalidFormat("missing active header/footer column region".to_string())
                })?;
            if namespace != Some(STYLE_NAMESPACE) || Column::parse(local) != Some(column_kind) {
                return Err(Error::InvalidFormat(
                    "malformed style:region-* column region nesting".to_string(),
                ));
            }
            if self.block.is_some() || self.field.is_some() {
                return Err(Error::InvalidFormat(
                    "unterminated header/footer block inside a column region".to_string(),
                ));
            }
            self.active_column = None;
        }
        Ok(())
    }

    fn finish_block(&mut self) -> Result<()> {
        let block = self
            .block
            .take()
            .ok_or_else(|| Error::InvalidFormat("missing header/footer block".to_string()))?;
        self.blocks.push(block.block);
        Ok(())
    }

    fn push_text(&mut self, text: &str) -> Result<()> {
        if text.is_empty() || self.block.is_none() {
            return Ok(());
        }
        if let Some(field) = self.field.as_mut() {
            field.field.displayed_text.push_str(text);
            return Ok(());
        }
        if let Some(Inline::Text(existing)) = self.active_block_mut()?.block.content.last_mut() {
            existing.push_str(text);
            return Ok(());
        }
        self.push_token(Inline::Text(text.to_string()))
    }

    fn append_field_control(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<()> {
        match element.local_name().as_ref() {
            b"s" => {
                let count = text_space_count(reader, element)?.unwrap_or(1);
                self.add_spaces(count)?;
                self.active_field_mut()?
                    .field
                    .displayed_text
                    .extend(std::iter::repeat_n(' ', count));
            },
            b"tab" => self.active_field_mut()?.field.displayed_text.push('\t'),
            b"line-break" => self.active_field_mut()?.field.displayed_text.push('\n'),
            _ => {},
        }
        Ok(())
    }

    fn add_spaces(&mut self, count: usize) -> Result<()> {
        self.expanded_spaces = self
            .expanded_spaces
            .checked_add(count)
            .ok_or_else(|| limit_error("expanded-space"))?;
        if self.expanded_spaces > MAX_EXPANDED_SPACES {
            return Err(limit_error("expanded-space"));
        }
        Ok(())
    }

    fn push_token(&mut self, token: Inline) -> Result<()> {
        self.token_count = self
            .token_count
            .checked_add(1)
            .ok_or_else(|| limit_error("inline-token"))?;
        if self.token_count > MAX_INLINE_TOKENS {
            return Err(limit_error("inline-token"));
        }
        self.active_block_mut()?.block.content.push(token);
        Ok(())
    }

    fn active_block_mut(&mut self) -> Result<&mut ActiveBlock> {
        self.block.as_mut().ok_or_else(|| {
            Error::InvalidFormat("header/footer token requires an active block".to_string())
        })
    }

    fn active_field_mut(&mut self) -> Result<&mut ActiveField> {
        self.field.as_mut().ok_or_else(|| {
            Error::InvalidFormat("header/footer control requires an active field".to_string())
        })
    }
}

pub(crate) fn parse(xml: &str) -> Result<HashMap<(String, Kind), Vec<Block>>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut master: Option<Master> = None;
    let mut region: Option<Region> = None;
    let mut regions = HashMap::new();

    loop {
        let (resolved, parsed_event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("styles.xml parsing error: {error}")))?;
        let resolved_bytes = resolved_namespace(&resolved).map(<[u8]>::to_vec);
        let namespace = resolved_bytes.as_deref();
        let event = parsed_event.into_owned();
        match event {
            Event::Start(element) => {
                let element_depth = depth.checked_add(1).ok_or_else(depth_error)?;
                if element_depth > MAX_XML_DEPTH {
                    return Err(depth_error());
                }
                if namespace == Some(STYLE_NAMESPACE)
                    && element.local_name().as_ref() == b"master-page"
                {
                    if master.is_some() {
                        return Err(Error::InvalidFormat(
                            "nested style:master-page element".to_string(),
                        ));
                    }
                    let name = namespaced_attr(&reader, &element, STYLE_NAMESPACE, b"name")?
                        .ok_or_else(|| {
                            Error::InvalidFormat(
                                "style:master-page is missing style:name".to_string(),
                            )
                        })?;
                    master = Some(Master {
                        name,
                        depth: element_depth,
                    });
                } else if let (Some(master_page), Some(kind)) =
                    (master.as_ref(), Kind::parse(element.local_name().as_ref()))
                {
                    if region.is_none() && namespace == Some(STYLE_NAMESPACE) {
                        region = Some(Region::new(master_page.name.clone(), kind, element_depth));
                    } else if let Some(active) = region.as_mut() {
                        active.start_element(&reader, namespace, &element, element_depth, false)?;
                    }
                } else if let Some(active) = region.as_mut() {
                    active.start_element(&reader, namespace, &element, element_depth, false)?;
                }
                depth = element_depth;
            },
            Event::Empty(element) => {
                let element_depth = depth.checked_add(1).ok_or_else(depth_error)?;
                if element_depth > MAX_XML_DEPTH {
                    return Err(depth_error());
                }
                if let (Some(master_page), Some(kind)) =
                    (master.as_ref(), Kind::parse(element.local_name().as_ref()))
                {
                    if region.is_none() && namespace == Some(STYLE_NAMESPACE) {
                        insert_region(&mut regions, &master_page.name, kind, Vec::new())?;
                    } else if let Some(active) = region.as_mut() {
                        active.start_element(&reader, namespace, &element, element_depth, true)?;
                    }
                } else if let Some(active) = region.as_mut() {
                    active.start_element(&reader, namespace, &element, element_depth, true)?;
                }
            },
            Event::Text(value) => {
                if let Some(active) = region.as_mut() {
                    let decoded = value
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| {
                            Error::InvalidFormat(format!("invalid header text: {error}"))
                        })?;
                    active.push_text(&decoded)?;
                }
            },
            Event::GeneralRef(reference) => {
                if let Some(active) = region.as_mut() {
                    active.push_text(&decode_reference(&reference)?)?;
                }
            },
            Event::CData(value) => {
                if let Some(active) = region.as_mut() {
                    let decoded = value
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| {
                            Error::InvalidFormat(format!("invalid header CDATA: {error}"))
                        })?;
                    active.push_text(&decoded)?;
                }
            },
            Event::End(element) => {
                if let Some(active) = region.as_mut() {
                    active.end_element(namespace, element.local_name().as_ref(), depth)?;
                }
                if region.as_ref().is_some_and(|active| active.depth == depth) {
                    let active = region.take().ok_or_else(|| {
                        Error::InvalidFormat("missing active header/footer region".to_string())
                    })?;
                    if active.active_column.is_some() || active.block.is_some() {
                        return Err(Error::InvalidFormat(
                            "unterminated header/footer column region or block".to_string(),
                        ));
                    }
                    if namespace != Some(STYLE_NAMESPACE)
                        || Kind::parse(element.local_name().as_ref()) != Some(active.kind)
                    {
                        return Err(Error::InvalidFormat(
                            "malformed header/footer region nesting".to_string(),
                        ));
                    }
                    insert_region(
                        &mut regions,
                        &active.master_name,
                        active.kind,
                        active.blocks,
                    )?;
                }
                if master.as_ref().is_some_and(|active| active.depth == depth) {
                    master = None;
                }
                depth = depth.checked_sub(1).ok_or_else(depth_error)?;
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DTD is not allowed in ODF styles.xml".to_string(),
                ));
            },
            Event::Eof => break,
            Event::Decl(_) | Event::PI(_) | Event::Comment(_) => {},
        }
        buffer.clear();
    }
    if master.is_some() || region.is_some() || depth != 0 {
        return Err(Error::InvalidFormat(
            "unterminated master-page header/footer".to_string(),
        ));
    }
    Ok(regions)
}

fn parse_field(reader: &NsReader<&[u8]>, element: &BytesStart<'_>, local: &[u8]) -> Result<Field> {
    let mut attributes = Vec::new();
    let mut fixed = None;
    let mut data_style_name = None;
    for raw_attribute in element.attributes() {
        if attributes.len() >= MAX_FIELD_ATTRIBUTES {
            return Err(limit_error("field-attribute"));
        }
        let attribute = raw_attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid header/footer field attribute: {error}"))
        })?;
        let (resolved_attribute, attr_local) = reader.resolver().resolve_attribute(attribute.key);
        let attribute_namespace = resolved_namespace(&resolved_attribute);
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid header/footer field attribute: {error}"))
            })?
            .into_owned();
        let key = canonical_attribute_name(attribute_namespace, attr_local.as_ref())?;
        if attribute_namespace == Some(TEXT_NAMESPACE) && attr_local.as_ref() == b"fixed" {
            fixed = Some(parse_bool(&value)?);
        }
        if attribute_namespace == Some(STYLE_NAMESPACE) && attr_local.as_ref() == b"data-style-name"
        {
            data_style_name = Some(value.clone());
        }
        attributes.push((key, value));
    }
    Ok(Field {
        kind: field_kind(local).unwrap_or_else(|| FieldKind::Unknown {
            namespace: String::from_utf8_lossy(TEXT_NAMESPACE).into_owned(),
            local_name: String::from_utf8_lossy(local).into_owned(),
        }),
        displayed_text: String::new(),
        fixed,
        data_style_name,
        attributes,
    })
}

fn field_kind(local: &[u8]) -> Option<FieldKind> {
    Some(match local {
        b"page-number" => FieldKind::PageNumber,
        b"page-count" => FieldKind::PageCount,
        b"page-continuation" => FieldKind::PageContinuation,
        b"page-variable-set" => FieldKind::PageVariableSet,
        b"page-variable-get" => FieldKind::PageVariableGet,
        b"paragraph-count" => FieldKind::ParagraphCount,
        b"word-count" => FieldKind::WordCount,
        b"character-count" => FieldKind::CharacterCount,
        b"table-count" => FieldKind::TableCount,
        b"image-count" => FieldKind::ImageCount,
        b"object-count" => FieldKind::ObjectCount,
        b"reference-ref" => FieldKind::Reference,
        b"sequence-ref" => FieldKind::SequenceReference,
        b"bookmark-ref" => FieldKind::BookmarkReference,
        b"note-ref" => FieldKind::NoteReference,
        b"variable-set" => FieldKind::VariableSet,
        b"variable-get" => FieldKind::VariableGet,
        b"variable-input" => FieldKind::VariableInput,
        b"user-field-get" => FieldKind::UserFieldGet,
        b"user-field-input" => FieldKind::UserFieldInput,
        b"sequence" => FieldKind::Sequence,
        b"expression" => FieldKind::Expression,
        b"text-input" => FieldKind::TextInput,
        b"placeholder" => FieldKind::Placeholder,
        b"conditional-text" => FieldKind::ConditionalText,
        b"hidden-text" => FieldKind::HiddenText,
        b"hidden-paragraph" => FieldKind::HiddenParagraph,
        b"dde-connection" => FieldKind::DdeConnection,
        b"measure" => FieldKind::Measure,
        b"table-formula" => FieldKind::TableFormula,
        b"meta-field" => FieldKind::MetaField,
        b"database-display" => FieldKind::DatabaseDisplay,
        b"database-next" => FieldKind::DatabaseNext,
        b"database-row-select" => FieldKind::DatabaseRowSelect,
        b"database-row-number" => FieldKind::DatabaseRowNumber,
        b"database-name" => FieldKind::DatabaseName,
        b"title" => FieldKind::Title,
        b"subject" => FieldKind::Subject,
        b"author-name" => FieldKind::AuthorName,
        b"author-initials" => FieldKind::AuthorInitials,
        b"date" => FieldKind::Date,
        b"time" => FieldKind::Time,
        b"file-name" => FieldKind::FileName,
        b"template-name" => FieldKind::TemplateName,
        b"sheet-name" => FieldKind::SheetName,
        b"chapter" => FieldKind::Chapter,
        b"initial-creator" => FieldKind::InitialCreator,
        b"description" => FieldKind::Description,
        b"printed-by" => FieldKind::PrintedBy,
        b"keywords" => FieldKind::Keywords,
        b"creator" => FieldKind::Creator,
        b"creation-date" => FieldKind::CreationDate,
        b"creation-time" => FieldKind::CreationTime,
        b"modification-date" => FieldKind::ModificationDate,
        b"modification-time" => FieldKind::ModificationTime,
        b"print-date" => FieldKind::PrintDate,
        b"print-time" => FieldKind::PrintTime,
        b"editing-cycles" => FieldKind::EditingCycles,
        b"editing-duration" => FieldKind::EditingDuration,
        b"user-defined" => FieldKind::UserDefined,
        b"sender-firstname" => FieldKind::Sender(SenderKind::FirstName),
        b"sender-lastname" => FieldKind::Sender(SenderKind::LastName),
        b"sender-initials" => FieldKind::Sender(SenderKind::Initials),
        b"sender-title" => FieldKind::Sender(SenderKind::Title),
        b"sender-position" => FieldKind::Sender(SenderKind::Position),
        b"sender-email" => FieldKind::Sender(SenderKind::Email),
        b"sender-phone-private" => FieldKind::Sender(SenderKind::PrivatePhone),
        b"sender-fax" => FieldKind::Sender(SenderKind::Fax),
        b"sender-company" => FieldKind::Sender(SenderKind::Company),
        b"sender-phone-work" => FieldKind::Sender(SenderKind::WorkPhone),
        b"sender-street" => FieldKind::Sender(SenderKind::Street),
        b"sender-city" => FieldKind::Sender(SenderKind::City),
        b"sender-postal-code" => FieldKind::Sender(SenderKind::PostalCode),
        b"sender-country" => FieldKind::Sender(SenderKind::Country),
        b"sender-state-or-province" => FieldKind::Sender(SenderKind::StateOrProvince),
        b"script" => FieldKind::Script,
        b"execute-macro" => FieldKind::ExecuteMacro,
        _ => return None,
    })
}

fn is_unknown_field(local: &[u8]) -> bool {
    !matches!(
        local,
        b"p" | b"h"
            | b"span"
            | b"a"
            | b"s"
            | b"tab"
            | b"line-break"
            | b"soft-page-break"
            | b"bookmark"
            | b"bookmark-start"
            | b"bookmark-end"
            | b"reference-mark"
            | b"reference-mark-start"
            | b"reference-mark-end"
            | b"alphabetical-index-mark"
            | b"alphabetical-index-mark-start"
            | b"alphabetical-index-mark-end"
            | b"toc-mark"
            | b"toc-mark-start"
            | b"toc-mark-end"
            | b"change"
            | b"change-start"
            | b"change-end"
            | b"note"
            | b"note-citation"
            | b"note-body"
            | b"ruby"
            | b"ruby-base"
            | b"ruby-text"
            | b"meta"
    )
}

fn insert_region(
    regions: &mut HashMap<(String, Kind), Vec<Block>>,
    master_name: &str,
    kind: Kind,
    blocks: Vec<Block>,
) -> Result<()> {
    if regions
        .insert((master_name.to_owned(), kind), blocks)
        .is_some()
    {
        return Err(Error::InvalidFormat(format!(
            "duplicate {kind:?} in master page '{master_name}'"
        )));
    }
    Ok(())
}

fn namespaced_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected_namespace: &[u8],
    local_name: &[u8],
) -> Result<Option<String>> {
    for raw_attribute in element.attributes() {
        let attribute = raw_attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid header/footer attribute: {error}"))
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if resolved_namespace(&namespace) == Some(expected_namespace)
            && local.as_ref() == local_name
        {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map(|value| Some(value.into_owned()))
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid header/footer attribute: {error}"))
                });
        }
    }
    Ok(None)
}

fn text_space_count(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Option<usize>> {
    let count_lexical = namespaced_attr(reader, element, TEXT_NAMESPACE, b"c")?;
    count_lexical
        .map(|lexical_value| {
            lexical_value.parse::<usize>().map_err(|error| {
                Error::InvalidFormat(format!("invalid text:c count in header/footer: {error}"))
            })
        })
        .transpose()
}

fn canonical_attribute_name(namespace_bytes: Option<&[u8]>, local_bytes: &[u8]) -> Result<String> {
    let local_name = std::str::from_utf8(local_bytes)
        .map_err(|error| Error::InvalidFormat(format!("invalid field attribute name: {error}")))?;
    Ok(match namespace_bytes {
        Some(TEXT_NAMESPACE) => format!("text:{local_name}"),
        Some(STYLE_NAMESPACE) => format!("style:{local_name}"),
        Some(OFFICE_NAMESPACE) => format!("office:{local_name}"),
        Some(NUMBER_NAMESPACE) => format!("number:{local_name}"),
        Some(other_namespace) => format!(
            "{{{}}}{local_name}",
            String::from_utf8_lossy(other_namespace)
        ),
        None => local_name.to_string(),
    })
}

fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(Error::InvalidFormat(format!(
            "invalid text:fixed value '{value}'"
        ))),
    }
}

fn resolved_namespace<'a>(namespace: &'a ResolveResult<'a>) -> Option<&'a [u8]> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => Some(*value),
        ResolveResult::Unbound | ResolveResult::Unknown(_) => None,
    }
}

fn decode_reference(reference: &BytesRef<'_>) -> Result<String> {
    if let Some(character) = reference.resolve_char_ref().map_err(|error| {
        Error::InvalidFormat(format!("invalid header character reference: {error}"))
    })? {
        return Ok(character.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| Error::InvalidFormat(format!("invalid header entity: {error}")))?;
    match name.as_ref() {
        "amp" => Ok("&".to_string()),
        "lt" => Ok("<".to_string()),
        "gt" => Ok(">".to_string()),
        "quot" => Ok("\"".to_string()),
        "apos" => Ok("'".to_string()),
        _ => Err(Error::InvalidFormat(format!(
            "unsupported header entity '&{name};'"
        ))),
    }
}

fn depth_error() -> Error {
    Error::InvalidFormat("header/footer XML depth exceeds safety limit".to_string())
}

fn limit_error(resource: &str) -> Error {
    Error::InvalidFormat(format!(
        "header/footer {resource} count exceeds safety limit"
    ))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn test_ok<T>(result: Result<T>) -> T {
        match result {
            Ok(value) => value,
            Err(error) => panic!("test operation failed: {error}"),
        }
    }

    fn field_items(block: &Block) -> Vec<&Field> {
        block
            .content
            .iter()
            .filter_map(|inline| {
                if let Inline::Field(field) = inline {
                    Some(field)
                } else {
                    None
                }
            })
            .collect()
    }

    #[test]
    fn parses_ordered_blocks_controls_and_fields_with_arbitrary_prefixes() {
        let xml = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:master-styles><s:master-page s:name="A"><s:header><t:p t:style-name="Header">Page <t:page-number t:select-page="current" t:page-adjust="2" t:fixed="false" s:data-style-name="N1">7</t:page-number><t:s t:c="2"/><t:tab/><t:line-break/><t:sender-company t:fixed="true">Example</t:sender-company></t:p><t:h>Heading</t:h></s:header></s:master-page></o:master-styles></o:document-styles>"#;
        let regions = test_ok(parse(xml));
        let blocks = &regions[&(String::from("A"), Kind::Header)];
        assert_eq!(blocks.len(), 2);
        assert_eq!(blocks[0].style_name.as_deref(), Some("Header"));
        assert_eq!(blocks[0].content[0], Inline::Text("Page ".into()));
        let Inline::Field(page) = &blocks[0].content[1] else {
            panic!("expected page field");
        };
        assert_eq!(page.kind, FieldKind::PageNumber);
        assert_eq!(page.displayed_text, "7");
        assert_eq!(page.fixed, Some(false));
        assert_eq!(page.data_style_name.as_deref(), Some("N1"));
        assert!(
            page.attributes
                .contains(&("text:page-adjust".into(), "2".into()))
        );
        assert_eq!(blocks[0].content[2], Inline::Space { count: 2 });
        assert_eq!(blocks[0].content[3], Inline::Tab);
        assert_eq!(blocks[0].content[4], Inline::LineBreak);
        let Inline::Field(sender) = &blocks[0].content[5] else {
            panic!("expected sender field");
        };
        assert_eq!(sender.kind, FieldKind::Sender(SenderKind::Company));
        assert_eq!(sender.displayed_text, "Example");
        assert_eq!(sender.fixed, Some(true));
        assert_eq!(blocks[1].content, vec![Inline::Text("Heading".into())]);
    }

    #[test]
    fn classifies_all_standard_sender_field_names() {
        let cases = [
            ("sender-firstname", SenderKind::FirstName),
            ("sender-lastname", SenderKind::LastName),
            ("sender-initials", SenderKind::Initials),
            ("sender-title", SenderKind::Title),
            ("sender-position", SenderKind::Position),
            ("sender-email", SenderKind::Email),
            ("sender-phone-private", SenderKind::PrivatePhone),
            ("sender-fax", SenderKind::Fax),
            ("sender-company", SenderKind::Company),
            ("sender-phone-work", SenderKind::WorkPhone),
            ("sender-street", SenderKind::Street),
            ("sender-city", SenderKind::City),
            ("sender-postal-code", SenderKind::PostalCode),
            ("sender-country", SenderKind::Country),
            ("sender-state-or-province", SenderKind::StateOrProvince),
        ];

        for (local_name, sender_kind) in cases {
            assert_eq!(
                field_kind(local_name.as_bytes()),
                Some(FieldKind::Sender(sender_kind))
            );
            assert_eq!(sender_kind.element_name(), local_name);
        }
    }

    #[test]
    fn retains_inert_script_and_macro_metadata() {
        let xml = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:x="http://www.w3.org/1999/xlink"><o:master-styles><s:master-page s:name="A"><s:header><t:p><t:script x:type="simple" x:href="https://example.invalid/never-open">payload</t:script><t:execute-macro t:name="Standard.Module1.Main">button</t:execute-macro></t:p></s:header></s:master-page></o:master-styles></o:document-styles>"#;
        let regions = test_ok(parse(xml));
        let blocks = &regions[&(String::from("A"), Kind::Header)];
        let fields = field_items(&blocks[0]);

        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].kind, FieldKind::Script);
        assert_eq!(fields[0].displayed_text, "payload");
        assert!(fields[0].attributes.contains(&(
            "{http://www.w3.org/1999/xlink}href".into(),
            "https://example.invalid/never-open".into(),
        )));
        assert_eq!(fields[1].kind, FieldKind::ExecuteMacro);
        assert_eq!(fields[1].displayed_text, "button");
        assert!(
            fields[1]
                .attributes
                .contains(&("text:name".into(), "Standard.Module1.Main".into()))
        );
    }

    #[test]
    fn classifies_cached_document_identity_and_revision_fields() {
        let xml = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:master-styles><s:master-page s:name="A"><s:footer><t:p><t:author-initials>AI</t:author-initials><t:template-name>Letter</t:template-name><t:sheet-name>Sheet1</t:sheet-name><t:initial-creator>Initial</t:initial-creator><t:description>Summary</t:description><t:printed-by>Printer</t:printed-by><t:keywords>one,two</t:keywords><t:creator>Creator</t:creator><t:creation-date t:fixed="true" s:data-style-name="D1">2026-07-22</t:creation-date><t:creation-time>12:00</t:creation-time><t:modification-time>13:00</t:modification-time><t:print-date>2026-07-22</t:print-date><t:print-time>14:00</t:print-time><t:editing-cycles>3</t:editing-cycles><t:editing-duration t:duration="PT1H">1 hour</t:editing-duration></t:p></s:footer></s:master-page></o:master-styles></o:document-styles>"#;
        let regions = test_ok(parse(xml));
        let blocks = &regions[&(String::from("A"), Kind::Footer)];
        let fields = field_items(&blocks[0]);

        assert_eq!(
            fields.iter().map(|field| &field.kind).collect::<Vec<_>>(),
            vec![
                &FieldKind::AuthorInitials,
                &FieldKind::TemplateName,
                &FieldKind::SheetName,
                &FieldKind::InitialCreator,
                &FieldKind::Description,
                &FieldKind::PrintedBy,
                &FieldKind::Keywords,
                &FieldKind::Creator,
                &FieldKind::CreationDate,
                &FieldKind::CreationTime,
                &FieldKind::ModificationTime,
                &FieldKind::PrintDate,
                &FieldKind::PrintTime,
                &FieldKind::EditingCycles,
                &FieldKind::EditingDuration,
            ]
        );
        assert_eq!(fields[8].displayed_text, "2026-07-22");
        assert_eq!(fields[8].fixed, Some(true));
        assert_eq!(fields[8].data_style_name.as_deref(), Some("D1"));
        assert!(
            fields[14]
                .attributes
                .contains(&("text:duration".into(), "PT1H".into()))
        );
    }

    #[test]
    fn classifies_cached_page_navigation_and_statistic_fields() {
        let xml = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:master-styles><s:master-page s:name="A"><s:footer><t:p><t:page-continuation t:select-page="previous" t:string-value="Continued">Prev</t:page-continuation><t:page-variable-set t:active="true" t:page-adjust="2">3</t:page-variable-set><t:page-variable-get s:num-format="1">3</t:page-variable-get><t:paragraph-count>4</t:paragraph-count><t:word-count>5</t:word-count><t:character-count>6</t:character-count><t:table-count>7</t:table-count><t:image-count>8</t:image-count><t:object-count>9</t:object-count></t:p></s:footer></s:master-page></o:master-styles></o:document-styles>"#;
        let regions = test_ok(parse(xml));
        let blocks = &regions[&(String::from("A"), Kind::Footer)];
        let fields = field_items(&blocks[0]);

        assert_eq!(
            fields.iter().map(|field| &field.kind).collect::<Vec<_>>(),
            vec![
                &FieldKind::PageContinuation,
                &FieldKind::PageVariableSet,
                &FieldKind::PageVariableGet,
                &FieldKind::ParagraphCount,
                &FieldKind::WordCount,
                &FieldKind::CharacterCount,
                &FieldKind::TableCount,
                &FieldKind::ImageCount,
                &FieldKind::ObjectCount,
            ]
        );
        assert_eq!(fields[0].displayed_text, "Prev");
        assert!(
            fields[0]
                .attributes
                .contains(&("text:string-value".into(), "Continued".into()))
        );
        assert!(
            fields[1]
                .attributes
                .contains(&("text:page-adjust".into(), "2".into()))
        );
        assert!(
            fields[2]
                .attributes
                .contains(&("style:num-format".into(), "1".into()))
        );
    }

    #[test]
    fn classifies_cached_reference_sequence_and_variable_fields() {
        let xml = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:master-styles><s:master-page s:name="A"><s:footer><t:p><t:reference-ref t:ref-name="reference" t:reference-format="text">Reference</t:reference-ref><t:sequence-ref t:ref-name="sequence">Sequence reference</t:sequence-ref><t:bookmark-ref t:ref-name="bookmark">Bookmark</t:bookmark-ref><t:note-ref t:ref-name="note" t:note-class="footnote">Note</t:note-ref><t:variable-set t:name="variable" t:formula="of:=1+1">2</t:variable-set><t:variable-get t:name="variable">2</t:variable-get><t:variable-input t:name="input" t:description="Input">Input</t:variable-input><t:user-field-get t:name="user">User</t:user-field-get><t:user-field-input t:name="user-input" t:description="User input">User input</t:user-field-input><t:sequence t:name="Figure" t:formula="ooow:Figure+1">1</t:sequence><t:expression t:formula="of:=2+2">4</t:expression><t:text-input t:description="Prompt">Answer</t:text-input></t:p></s:footer></s:master-page></o:master-styles></o:document-styles>"#;
        let regions = test_ok(parse(xml));
        let blocks = &regions[&(String::from("A"), Kind::Footer)];
        let fields = field_items(&blocks[0]);

        assert_eq!(
            fields.iter().map(|field| &field.kind).collect::<Vec<_>>(),
            vec![
                &FieldKind::Reference,
                &FieldKind::SequenceReference,
                &FieldKind::BookmarkReference,
                &FieldKind::NoteReference,
                &FieldKind::VariableSet,
                &FieldKind::VariableGet,
                &FieldKind::VariableInput,
                &FieldKind::UserFieldGet,
                &FieldKind::UserFieldInput,
                &FieldKind::Sequence,
                &FieldKind::Expression,
                &FieldKind::TextInput,
            ]
        );
        assert_eq!(fields[0].displayed_text, "Reference");
        assert!(
            fields[0]
                .attributes
                .contains(&("text:ref-name".into(), "reference".into()))
        );
        assert!(
            fields[4]
                .attributes
                .contains(&("text:formula".into(), "of:=1+1".into()))
        );
        assert!(
            fields[11]
                .attributes
                .contains(&("text:description".into(), "Prompt".into()))
        );
    }

    #[test]
    fn classifies_cached_conditional_dde_formula_and_metadata_fields() {
        let xml = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:master-styles><s:master-page s:name="A"><s:footer><t:p><t:placeholder t:placeholder-type="text">Hint</t:placeholder><t:conditional-text t:condition="of:=1=1" t:string-value-if-true="yes" t:string-value-if-false="no">yes</t:conditional-text><t:hidden-text t:condition="of:=0=1" t:string-value="hidden">hidden</t:hidden-text><t:hidden-paragraph t:condition="of:=0=1">paragraph</t:hidden-paragraph><t:dde-connection t:connection-name="NeverOpen">cached DDE</t:dde-connection><t:measure t:kind="unit">cm</t:measure><t:table-formula t:formula="of:=SUM([.A1:.A2])">42</t:table-formula><t:meta-field xml:id="meta1">meta <t:span>text</t:span></t:meta-field></t:p></s:footer></s:master-page></o:master-styles></o:document-styles>"#;
        let regions = test_ok(parse(xml));
        let blocks = &regions[&(String::from("A"), Kind::Footer)];
        let fields = field_items(&blocks[0]);

        assert_eq!(
            fields.iter().map(|field| &field.kind).collect::<Vec<_>>(),
            vec![
                &FieldKind::Placeholder,
                &FieldKind::ConditionalText,
                &FieldKind::HiddenText,
                &FieldKind::HiddenParagraph,
                &FieldKind::DdeConnection,
                &FieldKind::Measure,
                &FieldKind::TableFormula,
                &FieldKind::MetaField,
            ]
        );
        assert!(
            fields[1]
                .attributes
                .contains(&("text:condition".into(), "of:=1=1".into()))
        );
        assert!(
            fields[4]
                .attributes
                .contains(&("text:connection-name".into(), "NeverOpen".into()))
        );
        assert!(
            fields[6]
                .attributes
                .contains(&("text:formula".into(), "of:=SUM([.A1:.A2])".into()))
        );
        assert_eq!(fields[7].displayed_text, "meta text");
    }

    #[test]
    fn classifies_inert_database_field_metadata() {
        let xml = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:master-styles><s:master-page s:name="A"><s:footer><t:p><t:database-display t:database-name="NeverOpen" t:table-name="Records" t:column-name="Name">Ada</t:database-display><t:database-next t:database-name="NeverOpen" t:table-name="Records" t:condition="of:=TRUE">Next</t:database-next><t:database-row-select t:database-name="NeverOpen" t:table-name="Records" t:condition="of:=TRUE">Select</t:database-row-select><t:database-row-number t:database-name="NeverOpen" t:table-name="Records" t:value="7">7</t:database-row-number><t:database-name t:database-name="NeverOpen" t:table-name="Records">NeverOpen</t:database-name></t:p></s:footer></s:master-page></o:master-styles></o:document-styles>"#;
        let regions = test_ok(parse(xml));
        let blocks = &regions[&(String::from("A"), Kind::Footer)];
        let fields = field_items(&blocks[0]);

        assert_eq!(
            fields.iter().map(|field| &field.kind).collect::<Vec<_>>(),
            vec![
                &FieldKind::DatabaseDisplay,
                &FieldKind::DatabaseNext,
                &FieldKind::DatabaseRowSelect,
                &FieldKind::DatabaseRowNumber,
                &FieldKind::DatabaseName,
            ]
        );
        assert_eq!(fields[0].displayed_text, "Ada");
        assert!(
            fields[0]
                .attributes
                .contains(&("text:database-name".into(), "NeverOpen".into()))
        );
        assert!(
            fields[1]
                .attributes
                .contains(&("text:condition".into(), "of:=TRUE".into()))
        );
        assert!(
            fields[3]
                .attributes
                .contains(&("text:value".into(), "7".into()))
        );
    }

    #[test]
    fn parses_libreoffice_title_field_regression_fixture() {
        let xml = include_str!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/odf/odt/title-field-invalidate.fodt"
        ));
        let regions = test_ok(parse(xml));
        let blocks = &regions[&(String::from("Standard"), Kind::Footer)];
        let fields: Vec<_> = field_items(&blocks[0])
            .into_iter()
            .map(|field| (&field.kind, field.displayed_text.as_str()))
            .collect();
        assert_eq!(
            fields,
            vec![
                (&FieldKind::Subject, "mysubject"),
                (&FieldKind::Title, "mytitle"),
                (&FieldKind::UserDefined, "1.1"),
                (&FieldKind::ModificationDate, "May 18, 2021"),
            ]
        );
    }

    #[test]
    fn rejects_invalid_boolean_dtd_depth_and_cumulative_spaces() {
        let wrap = |body: &str| {
            format!(
                r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:master-styles><s:master-page s:name="A"><s:header><t:p>{body}</t:p></s:header></s:master-page></o:master-styles></o:document-styles>"#
            )
        };
        assert!(parse(&wrap("<t:date t:fixed=\"yes\"/>")).is_err());
        assert!(parse(&format!("<!DOCTYPE x>{}", wrap("x"))).is_err());
        assert!(parse(&wrap("<t:s t:c=\"600000\"/><t:s t:c=\"600000\"/>")).is_err());
        let nested = format!("{}x{}", "<t:span>".repeat(129), "</t:span>".repeat(129));
        assert!(parse(&wrap(&nested)).is_err());
    }

    fn column_region_document(content: &str) -> String {
        format!(
            r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:master-styles><s:master-page s:name="A"><s:header>{content}</s:header></s:master-page></o:master-styles></o:document-styles>"#
        )
    }

    #[test]
    fn attributes_blocks_to_ordered_column_regions() {
        let xml = column_region_document(
            r#"<s:region-left><t:p t:style-name="Left">Left <t:page-number>1</t:page-number></t:p></s:region-left><s:region-center><t:p>Center</t:p><t:p>Second</t:p></s:region-center><s:region-right><t:p>Right</t:p></s:region-right>"#,
        );
        let regions = test_ok(parse(&xml));
        let blocks = &regions[&(String::from("A"), Kind::Header)];
        assert_eq!(blocks.len(), 4);
        assert_eq!(blocks[0].column_region, Some(Column::Left));
        assert_eq!(blocks[0].style_name.as_deref(), Some("Left"));
        assert_eq!(blocks[1].column_region, Some(Column::Center));
        assert_eq!(blocks[2].column_region, Some(Column::Center));
        assert_eq!(blocks[3].column_region, Some(Column::Right));
        assert!(matches!(&blocks[1].content[0], Inline::Text(text) if text == "Center"));
    }

    #[test]
    fn accepts_empty_and_skipped_column_regions() {
        let xml =
            column_region_document(r"<s:region-left/><s:region-right><t:p/></s:region-right>");
        let regions = test_ok(parse(&xml));
        let blocks = &regions[&(String::from("A"), Kind::Header)];
        assert_eq!(blocks.len(), 1);
        assert_eq!(blocks[0].column_region, Some(Column::Right));
    }

    #[test]
    fn rejects_malformed_column_region_usage() {
        // Duplicate or out-of-order regions.
        assert!(
            parse(&column_region_document(
                r"<s:region-right/><s:region-left><t:p/></s:region-left>"
            ))
            .is_err()
        );
        assert!(parse(&column_region_document(r"<s:region-left/><s:region-left/>")).is_err());
        // Plain blocks must not be mixed with column regions.
        assert!(
            parse(&column_region_document(
                r"<t:p>Plain</t:p><s:region-left><t:p/></s:region-left>"
            ))
            .is_err()
        );
        assert!(
            parse(&column_region_document(
                r"<s:region-left><t:p/></s:region-left><t:p>Plain</t:p>"
            ))
            .is_err()
        );
        // Regions cannot nest and may contain only paragraphs.
        assert!(
            parse(&column_region_document(
                r"<s:region-left><s:region-center><t:p/></s:region-center></s:region-left>"
            ))
            .is_err()
        );
        assert!(
            parse(&column_region_document(
                r"<s:region-left><t:h>Heading</t:h></s:region-left>"
            ))
            .is_err()
        );
        // Unterminated regions are malformed.
        assert!(parse(&column_region_document(r"<s:region-left><t:p>open</t:p>")).is_err());
    }
}
