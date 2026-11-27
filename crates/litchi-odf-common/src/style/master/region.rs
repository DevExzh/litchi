//! Header/footer region vocabulary for an ODF master page.

use super::content::Block;

/// One of the six standard header/footer regions supported by an ODF master.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Kind {
    Header,
    HeaderFirst,
    HeaderLeft,
    Footer,
    FooterFirst,
    FooterLeft,
}

impl Kind {
    pub(crate) fn parse(local_name: &[u8]) -> Option<Self> {
        match local_name {
            b"header" => Some(Self::Header),
            b"header-first" => Some(Self::HeaderFirst),
            b"header-left" => Some(Self::HeaderLeft),
            b"footer" => Some(Self::Footer),
            b"footer-first" => Some(Self::FooterFirst),
            b"footer-left" => Some(Self::FooterLeft),
            _ => None,
        }
    }

    pub(crate) const fn element_name(self) -> &'static str {
        match self {
            Self::Header => "header",
            Self::HeaderFirst => "header-first",
            Self::HeaderLeft => "header-left",
            Self::Footer => "footer",
            Self::FooterFirst => "footer-first",
            Self::FooterLeft => "footer-left",
        }
    }

    pub(crate) const fn order(self) -> u8 {
        match self {
            Self::Header => 0,
            Self::HeaderLeft => 1,
            Self::HeaderFirst => 2,
            Self::Footer => 3,
            Self::FooterLeft => 4,
            Self::FooterFirst => 5,
        }
    }
}

/// Losslessly retained content of one master-page header/footer region.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Region {
    pub kind: Kind,
    /// Exact element bytes from the styles part.
    pub xml: String,
    /// Best-effort visible literal text. Dynamic field values remain in `xml`.
    pub text: String,
    /// Ordered paragraphs/headings with explicit inline text, whitespace, and fields.
    pub blocks: Vec<Block>,
}
