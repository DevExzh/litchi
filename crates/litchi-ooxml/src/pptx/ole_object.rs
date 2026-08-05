//! Host error-boundary adapter for standalone PPTX OLE authoring.
//!
//! OLE/package XML scanning, payload validation, relationship creation, and
//! slide-tree mutation are owned by `litchi_pptx::presentation::embedded::ole`.
//! This module only preserves the host entry-point spelling and translates the
//! owner error into the aggregate OOXML error.

use crate::error::Result;
use crate::pptx::media_parts::map_pptx_error;
use crate::pptx::ole::PayloadKind;
use litchi_opc::OpcPackage;

pub use litchi_pptx::presentation::embedded::ole::{
    Authored as AuthoredOleObject, Frame as OleObjectFrame,
};

/// Add an inert embedded OLE/package payload to a slide.
pub fn add_ole_object(
    package: &mut OpcPackage,
    slide_part_name: &str,
    kind: PayloadKind,
    prog_id: Option<&str>,
    name: Option<&str>,
    frame: OleObjectFrame,
    payload: &[u8],
) -> Result<AuthoredOleObject> {
    litchi_pptx::presentation::embedded::ole::add(
        package,
        slide_part_name,
        kind,
        prog_id,
        name,
        frame,
        payload,
    )
    .map_err(map_pptx_error)
}
