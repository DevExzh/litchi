//! Bounded scan state used by the dynamic-text content codec.

#[derive(Debug)]
pub(crate) struct Span {
    pub(super) start: usize,
    pub(super) end: usize,
}

#[derive(Debug)]
pub(crate) enum ParagraphSite {
    BeforeEnd(usize),
    Empty {
        start: usize,
        end: usize,
        qualified_name: String,
    },
}

#[derive(Debug, Default)]
pub(crate) struct Scan {
    pub(super) fields: Vec<Span>,
    pub(super) database_fields: Vec<Span>,
    pub(super) paragraph: Option<ParagraphSite>,
}
