//! Semantic catalog edit plans and the physical spans they project onto.

/// `[MS-OE376]` section 2.1.622(c) requires `activeTab` in 0..=32,766.
pub(crate) const MAX_ACTIVE_TAB: usize = 32_766;
/// `[MS-OE376]` section 2.1.613(a) limits `<sheet>` to 32,767 occurrences.
pub(crate) const MAX_SHEETS: usize = 32_767;
/// `[MS-OE376]` section 2.1.612(b) requires `sheetId` in 1..=65,534.
pub(crate) const MAX_SHEET_ID: u32 = 65_534;
/// `[MS-OE376]` section 2.1.612(c) limits relationship IDs to 255 characters.
pub(super) const MAX_RELATIONSHIP_ID_CHARS: usize = 255;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Dialect {
    Transitional,
    Strict,
}

impl Dialect {
    pub(crate) const fn worksheet_namespace(self) -> &'static str {
        match self {
            Self::Transitional => "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            Self::Strict => "http://purl.oclc.org/ooxml/spreadsheetml/main",
        }
    }
}

/// Recognized sheet states that are safe to author.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum State {
    Visible,
    Hidden,
    VeryHidden,
}

impl State {
    pub(super) const fn attribute(self) -> Option<&'static str> {
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

/// One checked semantic sheet-name change.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Rename<'a> {
    pub(crate) sheet: &'a str,
    pub(crate) position: usize,
    pub(crate) relationship_id: &'a str,
    pub(crate) name: &'a str,
}

/// Semantic active-tab target. The physical workbook view remains private to
/// this low-level boundary.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Active<'a> {
    pub(crate) sheet: &'a str,
    pub(crate) position: usize,
}

/// One physical catalog record synthesized below the semantic facade.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Create<'a> {
    pub(crate) sheet: &'a str,
    pub(crate) position: usize,
    pub(crate) sheet_id: u32,
    pub(crate) relationship_id: &'a str,
    pub(crate) state: State,
}

/// Checked worksheet removals plus their final active-tab disposition.
/// Relationship IDs stay confined to this physical catalog boundary.
#[derive(Debug)]
pub(crate) struct Remove<'a> {
    pub(crate) sheet: &'a str,
    pub(crate) position: usize,
    pub(crate) relationship_ids: Vec<&'a str>,
    pub(crate) active: Active<'a>,
    pub(crate) local_scopes: usize,
}

/// Final relationship order plus semantic error context. Relationship IDs are
/// borrowed only inside this physical rewrite boundary.
#[derive(Debug)]
pub(crate) struct Order<'a> {
    pub(crate) sheet: &'a str,
    pub(crate) position: usize,
    pub(crate) relationship_ids: Vec<&'a str>,
    pub(crate) local_scopes: usize,
}

/// Move-only workbook rewrite plan.
#[derive(Debug)]
pub(crate) struct Plan<'a> {
    pub(crate) tabs: Vec<Tab<'a>>,
    pub(crate) renames: Vec<Rename<'a>>,
    /// A replacement for the first workbook view's active tab. `None` leaves
    /// its active sheet unchanged unless an order edit remaps its position.
    pub(crate) active: Option<Active<'a>>,
    /// Final sheet relationship order. `None` leaves order-dependent fields
    /// byte-exact.
    pub(crate) order: Option<Order<'a>>,
}

#[derive(Debug, Clone, Copy)]
pub(super) struct Span {
    pub(super) start: usize,
    pub(super) end: usize,
}

#[derive(Debug, Clone)]
pub(super) struct Attribute {
    pub(super) name: Box<str>,
    pub(super) value: Box<str>,
}

#[derive(Debug, Clone)]
pub(super) struct Tag {
    pub(super) name: Box<str>,
    pub(super) attributes: Box<[Attribute]>,
}

#[derive(Debug)]
pub(super) struct Slot {
    pub(super) span: Span,
    pub(super) tag_end: usize,
    pub(super) close_start: usize,
    pub(super) tag: Tag,
    pub(super) empty: bool,
}

#[derive(Debug)]
pub(super) struct SheetSlot {
    pub(super) relationship_id: Box<str>,
    pub(super) slot: Slot,
}

#[derive(Debug)]
pub(super) struct ViewSlot {
    pub(super) slot: Slot,
    pub(super) active: Option<usize>,
    pub(super) first: Option<u32>,
}

#[derive(Debug)]
pub(super) struct DefinedNameSlot {
    pub(super) slot: Slot,
    pub(super) local_sheet_id: Option<usize>,
}

#[derive(Debug)]
pub(super) struct Container {
    pub(super) slot: Slot,
    pub(super) payload: bool,
}

#[derive(Debug)]
pub(super) struct Layout {
    pub(super) root: Tag,
    pub(super) dialect: Dialect,
    pub(super) sheets: Container,
    pub(super) sheet_slots: Box<[SheetSlot]>,
    pub(super) book_views: Option<Container>,
    pub(super) workbook_views: Box<[ViewSlot]>,
    pub(super) defined_names: Option<Container>,
    pub(super) defined_name_slots: Box<[DefinedNameSlot]>,
    pub(super) protected: bool,
    pub(super) alternate_content: bool,
    pub(super) alternate_dependencies: bool,
}

#[derive(Debug)]
pub(super) struct Replacement {
    pub(super) span: Span,
    pub(super) bytes: Vec<u8>,
}

pub(crate) const FIRST_SHEET_SENTINEL: u32 = 4_294_967_286;
