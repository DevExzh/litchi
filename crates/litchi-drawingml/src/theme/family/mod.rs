//! Source-backed support for the DrawingML 2012 `themeFamily` extension.
//!
//! The extension is a small fragment owned by the shared DrawingML vocabulary.
//! A host format decides where the fragment is reached (normally an `a:ext`
//! child of a theme part); this module deliberately has no package or graph
//! knowledge.  Parsed values retain their complete source fragment, including
//! attributes and extension children that are outside the typed projection.

pub mod codec;
mod model;
pub mod part;
mod transaction;

pub use codec::{read, read_shared, write, write_to};
pub use model::{Family, Guid, ValueError};
pub use transaction::{Commit, Edit, Patch, Snapshot};

/// XML namespace defined by `[MS-ODRAWXML]` §2.4 and schema §5.17.
pub const NAMESPACE: &str = "http://schemas.microsoft.com/office/thememl/2012/main";
/// Transitional DrawingML namespace used by `a:ext` children in the typed
/// `CT_OfficeArtExtensionList` value contained by the optional family-
/// namespace `extLst` child.
pub const DRAWINGML_NAMESPACE: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
/// Strict DrawingML namespace used by `a:ext` children in an extension list.
pub const DRAWINGML_NAMESPACE_STRICT: &str = "http://purl.oclc.org/ooxml/drawingml/main";
/// XML namespace reserved for the `xml` prefix.
pub const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";
/// Namespace used by namespace declaration attributes.
pub const XMLNS_NAMESPACE: &str = "http://www.w3.org/2000/xmlns/";

/// Maximum complete fragment bytes retained by this bounded codec.
pub const MAX_XML_BYTES: usize = 1 << 20;
/// Maximum decoded `name` bytes accepted by the typed model.
pub const MAX_NAME_BYTES: usize = 64 * 1024;
/// Maximum lexical GUID bytes accepted by `ST_Guid` values.
pub const MAX_GUID_BYTES: usize = 38;
/// Maximum attributes inspected on one element.
pub const MAX_ATTRIBUTES: usize = 256;
/// Maximum namespace declarations inspected on one element.
pub const MAX_NAMESPACE_DECLARATIONS: usize = 256;
/// Maximum element nesting depth accepted by the fragment scanner.
pub const MAX_DEPTH: usize = 128;
/// Maximum element nodes accepted by the fragment scanner.
pub const MAX_NODES: usize = 100_000;

/// Maximum decoded value bytes retained for one unknown attribute.
pub const MAX_ATTRIBUTE_VALUE_BYTES: usize = 64 * 1024;
/// Maximum namespace prefix or URI bytes inspected on one element.
pub const MAX_NAMESPACE_BYTES: usize = 4_096;
