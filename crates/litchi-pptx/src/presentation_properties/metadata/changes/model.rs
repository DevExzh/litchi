//! Package-independent Changes Information values.

/// A namespace declaration retained for an opaque command fragment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Namespace {
    /// Empty means the default namespace.
    pub prefix: String,
    pub uri: String,
}

/// Changes Information author metadata.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Data {
    pub name: Option<String>,
    pub user_id: Option<String>,
    pub provider_id: Option<String>,
    pub client_id: Option<String>,
    pub email: Option<String>,
    pub date_time: Option<String>,
    pub version: Option<u32>,
    pub change_id: Option<String>,
    pub action_id: Option<i32>,
    /// Optional complete DrawingML `a:extLst` fragment.
    pub extension_xml: Option<Vec<u8>>,
}

/// A document-change bit from `ST_DocumentChangeBit`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    CustomSelection,
    AddSlide,
    DeleteSlide,
    ModifySlide,
    SlideOrder,
    ModifyMainMaster,
    ModifyNotesMaster,
    ModifyHandoutMaster,
    AddSection,
    DeleteSection,
    ModifySection,
}

/// An inert `pc:docChg` descriptor and its typed change bits.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Descriptor {
    pub change_kinds: Vec<Kind>,
    /// Complete `pc:docChg` fragment with nested commands kept inert.
    pub xml: Vec<u8>,
}

/// One `pc:docChgLst` entry.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct List {
    pub author: Option<Data>,
    pub changes: Vec<Descriptor>,
    /// Optional complete PresentationML `p:extLst` fragment.
    pub extension_xml: Option<Vec<u8>>,
}

/// Typed view of one Changes Information part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Info {
    pub command_prefix: String,
    pub namespace_declarations: Vec<Namespace>,
    pub change_lists: Vec<List>,
}

impl Default for Info {
    fn default() -> Self {
        Self {
            command_prefix: "pc".into(),
            namespace_declarations: Vec::new(),
            change_lists: Vec::new(),
        }
    }
}

/// Changes Information bound to its PresentationML relationship.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Part {
    pub relationship_id: String,
    pub part_name: String,
    pub changes_information: Info,
}
