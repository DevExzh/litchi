//! Primary API for presentation consumers.

/// Unified source-checked presentation package transactions and history.
pub mod edit {
    pub use crate::authoring::edit::{
        Commit, Domain, History, MergePlan, Patch, SecurityPolicy, Selector, Snapshot, Transaction,
    };
}

/// Bounded, inert embedded-object discovery.
pub mod embedded {
    pub use litchi_odf_common::embedded::{Kind, Object, Parameter, Root, Source};
}

/// Source-bound flat-presentation reads and checked edits.
pub mod flat {
    #[allow(
        clippy::module_name_repetitions,
        reason = "re-exporting the flat presentation type under its own module is the intended public API shape"
    )]
    pub use crate::authoring::flat::{
        Commit, FlatPresentation, Patch, Selector, Snapshot, Transaction,
    };
}

/// Validated image-authoring values shared by builders and transactions.
pub mod image {
    pub use litchi_odf_common::drawing::authoring::Length;
    pub use litchi_odf_common::media::authoring::Format;
}

/// Named presentation page-layout definitions and their XML codec.
pub mod layout {
    pub use crate::model::page_layout::{
        Collection, Layout, Measure, Placeholder, Role, Unit, parse, remove_xml, set_xml,
    };
}

/// Presentation master pages and their shared ODF regions/children.
pub mod master {
    #[allow(
        clippy::module_name_repetitions,
        reason = "re-exporting the master page type under its own module is the intended public API shape"
    )]
    pub use crate::package::MasterPage;
    pub use litchi_odf_common::style::master::{Child, ChildKind, Kind, Master, Region};
}

/// Static metadata attached to presentation pages.
pub mod page {
    pub use crate::model::page_metadata::{Collection, Page, parse};
}

/// Inert slide-show settings and ordered custom shows.
pub mod settings {
    pub use crate::model::settings::{CustomShow, Settings, parse};
}

/// Slide-owned semantic values.
pub mod slide {
    pub use crate::model::slide::{
        DrawingAttribute, DrawingAttributeNamespace, DrawingShapeKind, EnhancedGeometry,
        EnhancedGeometryChild, EnhancedGeometryChildKind, Shape, Slide,
    };
}

pub use crate::authoring::Builder;
pub use crate::authoring::flat::FlatPresentation;
pub use crate::package::{MasterPage, Presentation};
