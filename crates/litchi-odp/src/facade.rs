//! Primary API for presentation consumers.

pub use crate::authoring::Builder;
pub use crate::package::{MasterPage, Presentation};

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

/// Presentation master pages and their shared ODF regions/children.
pub mod master {
    pub use crate::package::MasterPage;
    pub use litchi_odf_common::style::master::{Child, ChildKind, Kind, Master, Region};
}
