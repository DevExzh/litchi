//! Package-independent extended-guide values.

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Orientation {
    Horizontal,
    Vertical,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ColorKind {
    ScRgb,
    Srgb,
    Hsl,
    System,
    Scheme,
    Preset,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Color {
    pub kind: ColorKind,
    /// Inert DrawingML color XML, including transforms.
    pub xml: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Guide {
    pub id: u32,
    pub name: Option<String>,
    pub orientation: Option<Orientation>,
    pub position: Option<i32>,
    pub user_drawn: Option<bool>,
    pub color: Color,
    /// Optional, inert `p:extLst` permitted by `CT_ExtendedGuide`.
    pub extension_xml: Option<Vec<u8>>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct List {
    pub guides: Vec<Guide>,
    /// Optional, inert `p:extLst` permitted by `CT_ExtendedGuideList`.
    pub extension_xml: Option<Vec<u8>>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Guides {
    pub slide: Option<List>,
    pub notes: Option<List>,
}
