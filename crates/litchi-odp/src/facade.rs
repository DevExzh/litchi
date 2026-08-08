//! Primary API for presentation consumers.

pub use crate::authoring::Builder;
pub use crate::authoring::flat::FlatPresentation;
pub use crate::package::{MasterPage, Presentation};

/// Source-checked presentation snapshots, transactions, commits, and patches.
pub mod edit {
    pub use crate::authoring::edit::{Commit, Patch, Selector, Snapshot, Transaction};
}

/// Source-bound flat-presentation reads and checked edits.
pub mod flat {
    pub use crate::authoring::flat::{
        Commit, FlatPresentation, Patch, Selector, Snapshot, Transaction,
    };
}

/// Slide-owned semantic values.
pub mod slide {
    pub use crate::model::slide::{
        DrawingAttribute, DrawingAttributeNamespace, DrawingShapeKind, EnhancedGeometry,
        EnhancedGeometryChild, EnhancedGeometryChildKind, Shape, Slide,
    };
}

/// Named presentation page-layout definitions and their XML codec.
pub mod layout {
    pub use crate::model::page_layout::{
        Collection, Layout, Measure, Placeholder, Role, Unit, parse, remove_xml, set_xml,
    };
}

/// Static metadata attached to presentation pages.
pub mod page {
    pub use crate::model::page_metadata::{Collection, Page, parse};
}

/// Inert slide-show settings and ordered custom shows.
pub mod settings {
    pub use crate::model::settings::{CustomShow, Settings, parse};
}

/// Validated image-authoring values shared by builders and transactions.
pub mod image {
    pub use litchi_odf_common::drawing::authoring::Length;
    pub use litchi_odf_common::media::authoring::Format;
}

/// Presentation master pages and their shared ODF regions/children.
pub mod master {
    pub use crate::package::MasterPage;
    pub use litchi_odf_common::style::master::{Child, ChildKind, Kind, Master, Region};
}
