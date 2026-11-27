//! Package-independent Revision Information values.

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Namespace {
    /// Empty means the default namespace.
    pub prefix: String,
    pub uri: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Client {
    pub client_id: String,
    /// Omitted values retain the schema default of zero without losing lexical presence.
    pub revision: Option<u32>,
    /// Omitted values retain the schema default of zero without losing lexical presence.
    pub wet_revision: Option<u32>,
    pub date_time: String,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Info {
    pub clients: Vec<Client>,
    /// Namespace declarations inherited by the opaque extension fragment.
    pub namespace_declarations: Vec<Namespace>,
    /// Optional complete `p:extLst` fragment, retained inertly.
    pub extension_xml: Option<Vec<u8>>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Part {
    pub relationship_id: String,
    pub part_name: String,
    pub revision_information: Info,
}
