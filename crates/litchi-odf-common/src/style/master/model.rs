//! Master-page semantic model.

use super::region::{Kind, Region};

/// An ODF master page and its losslessly retained regions and direct children.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Master {
    pub name: String,
    pub display_name: Option<String>,
    pub page_layout_name: Option<String>,
    pub drawing_style_name: Option<String>,
    pub next_style_name: Option<String>,
    pub regions: Vec<Region>,
    /// Direct children classified in normative ODF order.
    pub children: Vec<Child>,
    /// Exact master-page element bytes, including shapes and extensions.
    pub xml: String,
}

/// Typed classification of one direct `style:master-page` child.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum ChildKind {
    Region(Kind),
    LayerSet,
    Forms,
    Shape,
    Animation,
    Notes,
}

impl ChildKind {
    pub(crate) const fn order(self) -> u8 {
        match self {
            Self::Region(Kind::Header) => 0,
            Self::Region(Kind::HeaderLeft) => 1,
            Self::Region(Kind::HeaderFirst) => 2,
            Self::Region(Kind::Footer) => 3,
            Self::Region(Kind::FooterLeft) => 4,
            Self::Region(Kind::FooterFirst) => 5,
            Self::LayerSet => 6,
            Self::Forms => 7,
            Self::Shape => 8,
            Self::Animation => 9,
            Self::Notes => 10,
        }
    }
}

/// Exact inert XML for one classified direct master-page child.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Child {
    pub kind: ChildKind,
    pub xml: String,
}

impl Master {
    /// Return a particular header/footer region when it exists.
    pub fn region(&self, kind: Kind) -> Option<&Region> {
        self.regions.iter().find(|region| region.kind == kind)
    }
}
