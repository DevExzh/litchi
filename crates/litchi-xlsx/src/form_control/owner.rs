//! Worksheet/package ownership for SpreadsheetML form controls.
//!
//! The leaf [`super::Properties`] codec owns one `formControlPr` part.  This
//! module proves the worksheet relationship and the DrawingML/VML identity
//! closure before exposing that leaf through a worksheet-facing, selector
//! first view.  It intentionally has no mutation or publication API.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::too_many_lines,
    clippy::type_complexity,
    reason = "the scanner follows the worksheet, relationship, drawing, and VML wire order"
)]

use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::mce::{Capabilities, Limits as MceLimits, process_markup_compatibility};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    EffectiveTopology, OpcPackage, OwnedRelationships, PackURI, PartData, ReadLimits, ReadResource,
    Relationships, SourceBackedPackage, TargetMode,
};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::borrow::Cow;
use std::collections::{HashMap, HashSet};
use std::fmt::{self, Write};
use std::io::Read;
use std::mem::size_of;
use std::ops::Range;
#[cfg(any(unix, windows))]
use std::path::Path;
use std::sync::Arc;

use super::{
    CONTROL_PROPERTIES_CONTENT_TYPE, CONTROL_PROPERTIES_RELATIONSHIP_TYPE, ControlSelector,
    FORM_CONTROL_NAMESPACE, FormControlError, Limits as LeafLimits, Properties,
    budget::RetainedBudgetHold,
};
use crate::error::{Error, Result, invalid};
use crate::raw;
use crate::source_payload::SourcePayload;
use crate::workbook::source::validate_sheet_graph;
use crate::{Selector, WorksheetKind};

const SML: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const REL: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const XDR: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const STRICT_XDR: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
const A14: &[u8] = b"http://schemas.microsoft.com/office/drawing/2010/main";
const VML: &[u8] = b"urn:schemas-microsoft-com:vml";
const OFFICE: &[u8] = b"urn:schemas-microsoft-com:office:office";
const EXCEL: &[u8] = b"urn:schemas-microsoft-com:office:excel";
const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const DRAWING: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWING: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
const MATH: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/math";
const STRICT_MATH: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/math";
const XML: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const XMLNS: &[u8] = b"http://www.w3.org/2000/xmlns/";

const MAX_SHAPE_ID: u64 = 67_098_623;
const MIN_SP_ID: u64 = 1_025;
const MAX_SP_ID: u64 = 268_435_456;
const MAX_CONTROLS_PROFILE: usize = 65_535;
const DEFAULT_MAX_DRAWING_BYTES: usize = 16 * 1024 * 1024;
const DEFAULT_MAX_VML_BYTES: usize = 16 * 1024 * 1024;
const DEFAULT_MAX_MCE_BYTES: usize = 32 * 1024 * 1024;
const DEFAULT_MAX_RELATIONSHIP_EDGES: usize = 1_000_000;
const DEFAULT_MAX_SHAPES: usize = 65_535;
const DEFAULT_MAX_MIRROR_NODES: usize = 65_536;
const DEFAULT_MAX_NAME_BYTES: usize = 4 * 1024;
const DEFAULT_MAX_SEMANTIC_BYTES: usize = 64 * 1024 * 1024;
const DEFAULT_MAX_PROJECTION_BYTES: usize = 64 * 1024 * 1024;
const DEFAULT_MAX_READ_SET_BYTES: usize = 64 * 1024 * 1024;
const DEFAULT_MAX_SIDECAR_TARGET_BYTES: usize = 32 * 1024 * 1024;
const DEFAULT_MAX_SCALAR_OPERATIONS: usize = 256;
const DEFAULT_MAX_CHANGED_PARTS: usize = 2;
const DEFAULT_MAX_OUTPUT_BYTES: usize = 64 * 1024 * 1024;
const DEFAULT_MAX_STAGING_BYTES: usize = 128 * 1024 * 1024;
const MAX_AUTHORED_NAME_CHARS: usize = 32;
const UNREFERENCED_DIAGNOSTIC_PREFIX: &str =
    "unreferenced control-properties part was retained as opaque source: ";
const ACTIVE_X_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control";
const STRICT_ACTIVE_X_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/control";

/// Bounded host policy for worksheet form-control reads.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct OwnerLimits {
    max_controls: usize,
    max_scalar_operations: usize,
    max_changed_parts: usize,
    max_output_bytes: usize,
    max_staging_bytes: usize,
    max_mce_branches: usize,
    max_mce_depth: usize,
    max_mce_events: usize,
    max_mce_bytes: usize,
    max_relationship_edges: usize,
    max_drawing_bytes: usize,
    max_vml_bytes: usize,
    max_sidecar_relationship_bytes: usize,
    max_control_properties_relationship_bytes: usize,
    max_shapes: usize,
    max_mirror_nodes: usize,
    max_name_bytes: usize,
    max_semantic_bytes: usize,
    max_projection_bytes: usize,
    max_read_set_bytes: usize,
    max_sidecar_target_bytes: usize,
}

impl Default for OwnerLimits {
    fn default() -> Self {
        Self {
            max_controls: MAX_CONTROLS_PROFILE,
            max_scalar_operations: DEFAULT_MAX_SCALAR_OPERATIONS,
            max_changed_parts: DEFAULT_MAX_CHANGED_PARTS,
            max_output_bytes: DEFAULT_MAX_OUTPUT_BYTES,
            max_staging_bytes: DEFAULT_MAX_STAGING_BYTES,
            max_mce_branches: 65_536,
            max_mce_depth: 256,
            max_mce_events: super::MAX_XML_EVENTS,
            max_mce_bytes: DEFAULT_MAX_MCE_BYTES,
            max_relationship_edges: DEFAULT_MAX_RELATIONSHIP_EDGES,
            max_drawing_bytes: DEFAULT_MAX_DRAWING_BYTES,
            max_vml_bytes: DEFAULT_MAX_VML_BYTES,
            max_sidecar_relationship_bytes: 4 * 1024 * 1024,
            max_control_properties_relationship_bytes: 4 * 1024 * 1024,
            max_shapes: DEFAULT_MAX_SHAPES,
            max_mirror_nodes: DEFAULT_MAX_MIRROR_NODES,
            max_name_bytes: DEFAULT_MAX_NAME_BYTES,
            max_semantic_bytes: DEFAULT_MAX_SEMANTIC_BYTES,
            max_projection_bytes: DEFAULT_MAX_PROJECTION_BYTES,
            max_read_set_bytes: DEFAULT_MAX_READ_SET_BYTES,
            max_sidecar_target_bytes: DEFAULT_MAX_SIDECAR_TARGET_BYTES,
        }
    }
}

impl OwnerLimits {
    /// Create the standard bounded host policy.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            max_controls: MAX_CONTROLS_PROFILE,
            max_scalar_operations: DEFAULT_MAX_SCALAR_OPERATIONS,
            max_changed_parts: DEFAULT_MAX_CHANGED_PARTS,
            max_output_bytes: DEFAULT_MAX_OUTPUT_BYTES,
            max_staging_bytes: DEFAULT_MAX_STAGING_BYTES,
            max_mce_branches: 65_536,
            max_mce_depth: 256,
            max_mce_events: super::MAX_XML_EVENTS,
            max_mce_bytes: DEFAULT_MAX_MCE_BYTES,
            max_relationship_edges: DEFAULT_MAX_RELATIONSHIP_EDGES,
            max_drawing_bytes: DEFAULT_MAX_DRAWING_BYTES,
            max_vml_bytes: DEFAULT_MAX_VML_BYTES,
            max_sidecar_relationship_bytes: 4 * 1024 * 1024,
            max_control_properties_relationship_bytes: 4 * 1024 * 1024,
            max_shapes: DEFAULT_MAX_SHAPES,
            max_mirror_nodes: DEFAULT_MAX_MIRROR_NODES,
            max_name_bytes: DEFAULT_MAX_NAME_BYTES,
            max_semantic_bytes: DEFAULT_MAX_SEMANTIC_BYTES,
            max_projection_bytes: DEFAULT_MAX_PROJECTION_BYTES,
            max_read_set_bytes: DEFAULT_MAX_READ_SET_BYTES,
            max_sidecar_target_bytes: DEFAULT_MAX_SIDECAR_TARGET_BYTES,
        }
    }

    /// Lower the effective-control ceiling.  The profile ceiling is 65,535.
    #[must_use]
    pub const fn with_max_controls(mut self, value: usize) -> Self {
        self.max_controls = if value < MAX_CONTROLS_PROFILE {
            value
        } else {
            MAX_CONTROLS_PROFILE
        };
        self
    }

    /// Lower the number of scalar operations admitted by one edit.
    #[must_use]
    pub const fn with_max_scalar_operations(mut self, value: usize) -> Self {
        self.max_scalar_operations = if value < DEFAULT_MAX_SCALAR_OPERATIONS {
            value
        } else {
            DEFAULT_MAX_SCALAR_OPERATIONS
        };
        self
    }

    /// Lower the number of changed source parts admitted by one paired edit.
    #[must_use]
    pub const fn with_max_changed_parts(mut self, value: usize) -> Self {
        self.max_changed_parts = if value < DEFAULT_MAX_CHANGED_PARTS {
            value
        } else {
            DEFAULT_MAX_CHANGED_PARTS
        };
        self
    }

    /// Lower the aggregate generated bytes admitted by one paired edit.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, value: usize) -> Self {
        self.max_output_bytes = if value < DEFAULT_MAX_OUTPUT_BYTES {
            value
        } else {
            DEFAULT_MAX_OUTPUT_BYTES
        };
        self
    }

    /// Lower temporary source/output staging bytes admitted by one edit.
    #[must_use]
    pub const fn with_max_staging_bytes(mut self, value: usize) -> Self {
        self.max_staging_bytes = if value < DEFAULT_MAX_STAGING_BYTES {
            value
        } else {
            DEFAULT_MAX_STAGING_BYTES
        };
        self
    }

    /// Lower the selected/ignored MCE branch ceiling.
    #[must_use]
    pub const fn with_max_mce_branches(mut self, value: usize) -> Self {
        self.max_mce_branches = if value < 65_536 { value } else { 65_536 };
        self
    }

    /// Lower the MCE nesting ceiling.
    #[must_use]
    pub const fn with_max_mce_depth(mut self, value: usize) -> Self {
        self.max_mce_depth = if value < 256 { value } else { 256 };
        self
    }

    /// Lower the XML-event ceiling applied to every owner scanner and leaf.
    #[must_use]
    pub const fn with_max_mce_events(mut self, value: usize) -> Self {
        self.max_mce_events = if value < super::MAX_XML_EVENTS {
            value
        } else {
            super::MAX_XML_EVENTS
        };
        self
    }

    /// Lower the MCE source/output-byte ceiling.
    #[must_use]
    pub const fn with_max_mce_bytes(mut self, value: usize) -> Self {
        self.max_mce_bytes = if value < DEFAULT_MAX_MCE_BYTES {
            value
        } else {
            DEFAULT_MAX_MCE_BYTES
        };
        self
    }

    /// Lower the relationship-edge ceiling.
    #[must_use]
    pub const fn with_max_relationship_edges(mut self, value: usize) -> Self {
        self.max_relationship_edges = if value < DEFAULT_MAX_RELATIONSHIP_EDGES {
            value
        } else {
            DEFAULT_MAX_RELATIONSHIP_EDGES
        };
        self
    }

    /// Lower the retained DrawingML byte ceiling.
    #[must_use]
    pub const fn with_max_drawing_bytes(mut self, value: usize) -> Self {
        self.max_drawing_bytes = if value < DEFAULT_MAX_DRAWING_BYTES {
            value
        } else {
            DEFAULT_MAX_DRAWING_BYTES
        };
        self
    }

    /// Lower the retained VML byte ceiling.
    #[must_use]
    pub const fn with_max_vml_bytes(mut self, value: usize) -> Self {
        self.max_vml_bytes = if value < DEFAULT_MAX_VML_BYTES {
            value
        } else {
            DEFAULT_MAX_VML_BYTES
        };
        self
    }

    /// Lower the DrawingML/VML sidecar relationship-member ceiling.
    #[must_use]
    pub const fn with_max_sidecar_relationship_bytes(mut self, value: usize) -> Self {
        self.max_sidecar_relationship_bytes = if value < 4 * 1024 * 1024 {
            value
        } else {
            4 * 1024 * 1024
        };
        self
    }

    /// Lower the selected control-properties relationship-member ceiling.
    #[must_use]
    pub const fn with_max_control_properties_relationship_bytes(mut self, value: usize) -> Self {
        self.max_control_properties_relationship_bytes = if value < 4 * 1024 * 1024 {
            value
        } else {
            4 * 1024 * 1024
        };
        self
    }

    /// Lower the retained shape-identity ceiling.
    #[must_use]
    pub const fn with_max_shapes(mut self, value: usize) -> Self {
        self.max_shapes = if value < DEFAULT_MAX_SHAPES {
            value
        } else {
            DEFAULT_MAX_SHAPES
        };
        self
    }

    /// Lower the mirror-node ceiling.
    #[must_use]
    pub const fn with_max_mirror_nodes(mut self, value: usize) -> Self {
        self.max_mirror_nodes = if value < DEFAULT_MAX_MIRROR_NODES {
            value
        } else {
            DEFAULT_MAX_MIRROR_NODES
        };
        self
    }

    /// Lower the bound for one authored control, shape, or VML identity name.
    #[must_use]
    pub const fn with_max_name_bytes(mut self, value: usize) -> Self {
        self.max_name_bytes = if value < DEFAULT_MAX_NAME_BYTES {
            value
        } else {
            DEFAULT_MAX_NAME_BYTES
        };
        self
    }

    /// Lower the aggregate semantic model-byte ceiling for selected
    /// properties.  This is independent of the per-part leaf ceiling.
    #[must_use]
    pub const fn with_max_semantic_bytes(mut self, value: usize) -> Self {
        self.max_semantic_bytes = if value < DEFAULT_MAX_SEMANTIC_BYTES {
            value
        } else {
            DEFAULT_MAX_SEMANTIC_BYTES
        };
        self
    }

    /// Lower the aggregate worksheet-control projection-byte ceiling.
    #[must_use]
    pub const fn with_max_projection_bytes(mut self, value: usize) -> Self {
        self.max_projection_bytes = if value < DEFAULT_MAX_PROJECTION_BYTES {
            value
        } else {
            DEFAULT_MAX_PROJECTION_BYTES
        };
        self
    }

    /// Lower the aggregate retained source read-set-byte ceiling.
    #[must_use]
    pub const fn with_max_read_set_bytes(mut self, value: usize) -> Self {
        self.max_read_set_bytes = if value < DEFAULT_MAX_READ_SET_BYTES {
            value
        } else {
            DEFAULT_MAX_READ_SET_BYTES
        };
        self
    }

    /// Lower the aggregate sidecar-target payload-byte ceiling.
    #[must_use]
    pub const fn with_max_sidecar_target_bytes(mut self, value: usize) -> Self {
        self.max_sidecar_target_bytes = if value < DEFAULT_MAX_SIDECAR_TARGET_BYTES {
            value
        } else {
            DEFAULT_MAX_SIDECAR_TARGET_BYTES
        };
        self
    }

    /// Derive an owner policy from the caller's OPC read policy.
    #[must_use]
    pub fn from_read_limits(read_limits: ReadLimits) -> Self {
        let mut limits = Self::default();
        limits.max_drawing_bytes =
            lower_u64(limits.max_drawing_bytes, read_limits.max_part_bytes());
        limits.max_vml_bytes = lower_u64(limits.max_vml_bytes, read_limits.max_part_bytes());
        limits.max_mce_bytes = lower_u64(limits.max_mce_bytes, read_limits.max_part_bytes());
        limits.max_sidecar_relationship_bytes = lower_usize(
            limits.max_sidecar_relationship_bytes,
            read_limits.max_relationship_xml_bytes(),
        );
        limits.max_control_properties_relationship_bytes = lower_usize(
            limits.max_control_properties_relationship_bytes,
            read_limits.max_relationship_xml_bytes(),
        );
        limits.max_relationship_edges = limits
            .max_relationship_edges
            .min(read_limits.max_relationship_graph_nodes());
        limits.max_semantic_bytes = lower_u64(
            limits.max_semantic_bytes,
            read_limits.max_total_part_bytes(),
        );
        limits.max_projection_bytes = lower_u64(
            limits.max_projection_bytes,
            read_limits.max_total_part_bytes(),
        );
        limits.max_read_set_bytes = lower_u64(
            limits.max_read_set_bytes,
            read_limits.max_total_part_bytes(),
        );
        limits.max_sidecar_target_bytes = lower_u64(
            limits.max_sidecar_target_bytes,
            read_limits.max_total_part_bytes(),
        );
        limits.max_output_bytes =
            lower_u64(limits.max_output_bytes, read_limits.max_total_part_bytes());
        limits.max_staging_bytes =
            lower_u64(limits.max_staging_bytes, read_limits.max_total_part_bytes());
        limits
    }

    /// Maximum effective controls admitted by this policy.
    #[must_use]
    pub const fn max_controls(self) -> usize {
        self.max_controls
    }

    /// Maximum scalar operations admitted by one edit.
    #[must_use]
    pub const fn max_scalar_operations(self) -> usize {
        self.max_scalar_operations
    }

    /// Maximum changed source parts admitted by one paired edit.
    #[must_use]
    pub const fn max_changed_parts(self) -> usize {
        self.max_changed_parts
    }

    /// Maximum aggregate generated bytes admitted by one paired edit.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    /// Maximum temporary source/output staging bytes admitted by one edit.
    #[must_use]
    pub const fn max_staging_bytes(self) -> usize {
        self.max_staging_bytes
    }

    /// Maximum MCE branches admitted by this policy.
    #[must_use]
    pub const fn max_mce_branches(self) -> usize {
        self.max_mce_branches
    }

    /// Maximum MCE nesting depth admitted by this policy.
    #[must_use]
    pub const fn max_mce_depth(self) -> usize {
        self.max_mce_depth
    }

    /// Maximum XML events admitted by this policy.
    #[must_use]
    pub const fn max_mce_events(self) -> usize {
        self.max_mce_events
    }

    /// Maximum bytes retained for MCE input/output.
    #[must_use]
    pub const fn max_mce_bytes(self) -> usize {
        self.max_mce_bytes
    }

    /// Maximum relationship edges admitted by this policy.
    #[must_use]
    pub const fn max_relationship_edges(self) -> usize {
        self.max_relationship_edges
    }

    /// Maximum DrawingML bytes retained by this policy.
    #[must_use]
    pub const fn max_drawing_bytes(self) -> usize {
        self.max_drawing_bytes
    }

    /// Maximum VML bytes retained by this policy.
    #[must_use]
    pub const fn max_vml_bytes(self) -> usize {
        self.max_vml_bytes
    }

    /// Maximum shape identities admitted by this policy.
    #[must_use]
    pub const fn max_shapes(self) -> usize {
        self.max_shapes
    }

    /// Maximum mirror nodes admitted by this policy.
    #[must_use]
    pub const fn max_mirror_nodes(self) -> usize {
        self.max_mirror_nodes
    }

    /// Maximum retained DrawingML/VML sidecar relationship bytes.
    #[must_use]
    pub const fn max_sidecar_relationship_bytes(self) -> usize {
        self.max_sidecar_relationship_bytes
    }

    /// Maximum retained selected control-properties relationship bytes.
    #[must_use]
    pub const fn max_control_properties_relationship_bytes(self) -> usize {
        self.max_control_properties_relationship_bytes
    }

    /// Maximum bytes in one authored identity/name token.
    #[must_use]
    pub const fn max_name_bytes(self) -> usize {
        self.max_name_bytes
    }

    /// Maximum aggregate semantic model bytes retained across controls.
    #[must_use]
    pub const fn max_semantic_bytes(self) -> usize {
        self.max_semantic_bytes
    }

    /// Maximum aggregate control projection bytes retained by one collection.
    #[must_use]
    pub const fn max_projection_bytes(self) -> usize {
        self.max_projection_bytes
    }

    /// Maximum aggregate bytes retained by one source read set.
    #[must_use]
    pub const fn max_read_set_bytes(self) -> usize {
        self.max_read_set_bytes
    }

    /// Maximum aggregate payload bytes retained from sidecar targets.
    #[must_use]
    pub const fn max_sidecar_target_bytes(self) -> usize {
        self.max_sidecar_target_bytes
    }
}

fn lower_usize(current: usize, requested: usize) -> usize {
    if requested < current {
        requested
    } else {
        current
    }
}

fn lower_u64(current: usize, requested: u64) -> usize {
    current.min(usize::try_from(requested).unwrap_or(usize::MAX))
}

/// A stable host profile identity for the first canonical worksheet dialect.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct OwnerProfile {
    /// Profile version.
    pub version: u16,
    /// Expanded worksheet capability URI.
    pub worksheet_capability: &'static str,
    /// Expanded DrawingML capability URI.
    pub drawing_capability: &'static str,
}

impl OwnerProfile {
    /// The admitted canonical x14/a14 profile.
    #[must_use]
    pub const fn canonical() -> Self {
        Self {
            version: 1,
            worksheet_capability: FORM_CONTROL_NAMESPACE,
            drawing_capability: "http://schemas.microsoft.com/office/drawing/2010/main",
        }
    }
}

#[derive(Clone, Debug)]
enum RetainedPayload {
    Owned(Arc<Vec<u8>>),
    Source(PartData),
    Relationships(OwnedRelationships),
}

impl RetainedPayload {
    fn owned(bytes: Vec<u8>) -> Self {
        Self::Owned(Arc::new(bytes))
    }

    fn source(data: PartData) -> Self {
        Self::Source(data)
    }

    fn relationships(token: OwnedRelationships) -> Self {
        Self::Relationships(token)
    }

    fn as_bytes(&self) -> &[u8] {
        match self {
            Self::Owned(bytes) => bytes,
            Self::Source(data) => data.as_bytes(),
            Self::Relationships(token) => token.bytes(),
        }
    }

    fn source_payload(&self) -> SourcePayload {
        match self {
            Self::Owned(bytes) => SourcePayload::Owned(Arc::clone(bytes)),
            Self::Source(data) => SourcePayload::Managed(data.clone()),
            Self::Relationships(token) => SourcePayload::Owned(Arc::new(token.bytes().to_vec())),
        }
    }

    const fn is_source_backed(&self) -> bool {
        matches!(self, Self::Source(_))
    }
}

impl PartialEq for RetainedPayload {
    fn eq(&self, other: &Self) -> bool {
        self.as_bytes() == other.as_bytes()
    }
}

impl Eq for RetainedPayload {}

/// Bounded provenance for one MCE-selected owner part.
#[derive(Clone, Debug)]
pub struct MceProvenance {
    raw: RetainedPayload,
    selected: Arc<[u8]>,
    ignored: Arc<[u8]>,
    selected_ranges: Arc<[Range<usize>]>,
    ignored_ranges: Arc<[Range<usize>]>,
    wrapper_ranges: Arc<[Range<usize>]>,
    selected_parent_ranges: Arc<[Range<usize>]>,
    ignored_parent_ranges: Arc<[Range<usize>]>,
    selected_choices: usize,
    selected_fallbacks: usize,
    ignored_elements: usize,
    ignored_attributes: usize,
    budget_hold: Option<Arc<OwnerBudgetHold>>,
}

impl PartialEq for MceProvenance {
    fn eq(&self, other: &Self) -> bool {
        self.raw == other.raw
            && self.selected == other.selected
            && self.ignored == other.ignored
            && self.selected_ranges == other.selected_ranges
            && self.ignored_ranges == other.ignored_ranges
            && self.wrapper_ranges == other.wrapper_ranges
            && self.selected_parent_ranges == other.selected_parent_ranges
            && self.ignored_parent_ranges == other.ignored_parent_ranges
            && self.selected_choices == other.selected_choices
            && self.selected_fallbacks == other.selected_fallbacks
            && self.ignored_elements == other.ignored_elements
            && self.ignored_attributes == other.ignored_attributes
    }
}

impl Eq for MceProvenance {}

impl MceProvenance {
    /// Exact bounded source bytes presented to the MCE resolver.
    #[must_use]
    pub fn raw_bytes(&self) -> &[u8] {
        self.raw.as_bytes()
    }

    /// Exact bounded semantic bytes emitted by the resolver.
    #[must_use]
    pub fn selected_bytes(&self) -> &[u8] {
        &self.selected
    }

    /// Source bytes retained as inactive/ignored provenance.
    #[must_use]
    pub fn ignored_bytes(&self) -> &[u8] {
        &self.ignored
    }

    /// Source ranges of selected MCE branches, relative to [`Self::raw_bytes`].
    #[must_use]
    pub fn selected_ranges(&self) -> &[Range<usize>] {
        &self.selected_ranges
    }

    /// Source ranges of inactive MCE branches, relative to [`Self::raw_bytes`].
    #[must_use]
    pub fn ignored_ranges(&self) -> &[Range<usize>] {
        &self.ignored_ranges
    }

    /// Source ranges of all `mc:AlternateContent` wrappers, including nested
    /// wrappers, relative to [`Self::raw_bytes`].
    #[must_use]
    pub fn wrapper_ranges(&self) -> &[Range<usize>] {
        &self.wrapper_ranges
    }

    /// Wrapper ranges whose effective branch was selected.
    #[must_use]
    pub fn selected_parent_ranges(&self) -> &[Range<usize>] {
        &self.selected_parent_ranges
    }

    /// Wrapper ranges whose effective branch was ignored.
    #[must_use]
    pub fn ignored_parent_ranges(&self) -> &[Range<usize>] {
        &self.ignored_parent_ranges
    }

    /// Number of selected `mc:Choice` branches.
    #[must_use]
    pub const fn selected_choices(&self) -> usize {
        self.selected_choices
    }

    /// Number of selected `mc:Fallback` branches.
    #[must_use]
    pub const fn selected_fallbacks(&self) -> usize {
        self.selected_fallbacks
    }

    /// Number of ignored source elements.
    #[must_use]
    pub const fn ignored_elements(&self) -> usize {
        self.ignored_elements
    }

    /// Number of ignored source attributes.
    #[must_use]
    pub const fn ignored_attributes(&self) -> usize {
        self.ignored_attributes
    }

    /// Whether this provenance clone keeps the owner execution lease alive.
    #[must_use]
    pub fn has_execution_budget(&self) -> bool {
        self.budget_hold
            .as_ref()
            .is_some_and(|hold| hold.reservation_count() != 0)
    }
}

/// A bounded sidecar relationship-member fingerprint.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct RelationshipFingerprint {
    present: bool,
    bytes: usize,
    digest: u64,
    edges: usize,
}

impl RelationshipFingerprint {
    /// Whether a physical `.rels` member was present.
    #[must_use]
    pub const fn present(self) -> bool {
        self.present
    }

    /// Bounded canonical relationship-byte size used for the fingerprint.
    #[must_use]
    pub const fn bytes(self) -> usize {
        self.bytes
    }

    /// Stable FNV-1a digest of the parsed relationship graph.
    #[must_use]
    pub const fn digest(self) -> u64 {
        self.digest
    }

    /// Number of parsed relationship edges.
    #[must_use]
    pub const fn edges(self) -> usize {
        self.edges
    }
}

/// One retained package owner payload in a source read set.
#[derive(Clone, Debug)]
pub struct FormControlPartRead {
    /// Physical part identity retained for the host editor.  This never
    /// crosses the ordinary public selector API; it is only used by the
    /// source-backed transaction that owns the complete read set.
    pub(crate) part_name: Option<PackURI>,
    payload: RetainedPayload,
    relationship_payload: Option<RetainedPayload>,
    source_version: Option<litchi_core::SourceVersion>,
    relationship_source_version: Option<litchi_core::SourceVersion>,
    range: Range<usize>,
    relationship_range: Option<Range<usize>>,
    relationships: RelationshipFingerprint,
}

impl PartialEq for FormControlPartRead {
    fn eq(&self, other: &Self) -> bool {
        self.part_name == other.part_name
            && self.payload == other.payload
            && self.relationship_payload == other.relationship_payload
            && self.source_version == other.source_version
            && self.relationship_source_version == other.relationship_source_version
            && self.range == other.range
            && self.relationship_range == other.relationship_range
            && self.relationships == other.relationships
    }
}

impl Eq for FormControlPartRead {}

impl FormControlPartRead {
    pub(crate) fn part_name(&self) -> Option<&PackURI> {
        self.part_name.as_ref()
    }
    /// Borrow the retained owner bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.payload.as_bytes()
    }

    pub(crate) fn source_payload(&self) -> SourcePayload {
        self.payload.source_payload()
    }

    /// Source version captured when this payload was read, if deferred.
    #[must_use]
    pub const fn source_version(&self) -> Option<litchi_core::SourceVersion> {
        self.source_version
    }

    /// Source version captured for the retained physical relationship
    /// member, if one was present in a deferred read set.
    #[must_use]
    pub const fn relationship_source_version(&self) -> Option<litchi_core::SourceVersion> {
        self.relationship_source_version
    }

    /// Exact range retained from the owner payload.
    #[must_use]
    pub fn range(&self) -> Range<usize> {
        self.range.clone()
    }

    /// Exact range retained for the physical relationship member, if present.
    #[must_use]
    pub fn relationship_range(&self) -> Option<Range<usize>> {
        self.relationship_range.clone()
    }

    /// Fingerprint of the owner's physical relationship member.
    #[must_use]
    pub const fn relationships(&self) -> RelationshipFingerprint {
        self.relationships
    }

    /// Borrow exact retained relationship-member bytes when a physical member
    /// was present.
    #[must_use]
    pub fn relationship_bytes(&self) -> Option<&[u8]> {
        self.relationship_payload
            .as_ref()
            .map(RetainedPayload::as_bytes)
    }

    /// Whether the retained relationship member is backed by managed
    /// source-owned `PartData`.
    #[must_use]
    pub fn relationship_is_source_backed(&self) -> bool {
        self.relationship_payload
            .as_ref()
            .is_some_and(RetainedPayload::is_source_backed)
    }

    /// Whether this payload is retained through source-backed `PartData`.
    #[must_use]
    pub const fn is_source_backed(&self) -> bool {
        self.payload.is_source_backed()
    }
}

/// Scoped owner read set retained by a form-control collection.
#[derive(Clone, Debug)]
pub struct FormControlReadSet {
    source_version: Option<litchi_core::SourceVersion>,
    workbook: Option<FormControlPartRead>,
    package_relationship_payload: Option<RetainedPayload>,
    package_relationships: RelationshipFingerprint,
    content_types_payload: Option<RetainedPayload>,
    signature_present: bool,
    incoming_relationships: RelationshipFingerprint,
    worksheet: FormControlPartRead,
    drawing: Option<FormControlPartRead>,
    vml: Option<FormControlPartRead>,
    properties: Arc<[FormControlPartRead]>,
    sidecar_targets: Arc<[FormControlPartRead]>,
    semantic_holds: Arc<[Arc<litchi_core::Reservation>]>,
    worksheet_mce: MceProvenance,
    drawing_mce: MceProvenance,
    budget_hold: Option<Arc<OwnerBudgetHold>>,
}

impl PartialEq for FormControlReadSet {
    fn eq(&self, other: &Self) -> bool {
        self.source_version == other.source_version
            && self.workbook == other.workbook
            && self.package_relationship_payload == other.package_relationship_payload
            && self.package_relationships == other.package_relationships
            && self.content_types_payload == other.content_types_payload
            && self.signature_present == other.signature_present
            && self.incoming_relationships == other.incoming_relationships
            && self.worksheet == other.worksheet
            && self.drawing == other.drawing
            && self.vml == other.vml
            && self.properties == other.properties
            && self.sidecar_targets == other.sidecar_targets
            && self.worksheet_mce == other.worksheet_mce
            && self.drawing_mce == other.drawing_mce
    }
}

impl Eq for FormControlReadSet {}

impl FormControlReadSet {
    pub(crate) fn worksheet_part_name(&self) -> Option<&PackURI> {
        self.worksheet.part_name()
    }

    pub(crate) fn vml_part_name(&self) -> Option<&PackURI> {
        self.vml.as_ref().and_then(FormControlPartRead::part_name)
    }

    pub(crate) fn property_part_name(&self, position: usize) -> Option<&PackURI> {
        self.properties
            .get(position)
            .and_then(FormControlPartRead::part_name)
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.source_version == other.source_version
            && self.workbook == other.workbook
            && self.package_relationship_payload == other.package_relationship_payload
            && self.package_relationships == other.package_relationships
            && self.content_types_payload == other.content_types_payload
            && self.signature_present == other.signature_present
            && self.incoming_relationships == other.incoming_relationships
            && self.worksheet == other.worksheet
            && self.drawing == other.drawing
            && self.vml == other.vml
            && self.properties == other.properties
            && self.sidecar_targets == other.sidecar_targets
            && self.worksheet_mce == other.worksheet_mce
            && self.drawing_mce == other.drawing_mce
    }
    /// Source version shared by every retained source payload.
    #[must_use]
    pub const fn source_version(&self) -> Option<litchi_core::SourceVersion> {
        self.source_version
    }

    /// Worksheet XML and relationship-member fingerprint.
    #[must_use]
    pub const fn worksheet(&self) -> &FormControlPartRead {
        &self.worksheet
    }

    /// Workbook payload used to resolve the selected worksheet catalog.
    #[must_use]
    pub fn workbook(&self) -> Option<&FormControlPartRead> {
        self.workbook.as_ref()
    }

    /// Package-root relationship fingerprint used by workbook resolution.
    #[must_use]
    pub const fn package_relationships(&self) -> RelationshipFingerprint {
        self.package_relationships
    }

    /// Exact `[Content_Types].xml` bytes retained with this source read set.
    #[must_use]
    pub fn content_types_bytes(&self) -> Option<&[u8]> {
        self.content_types_payload
            .as_ref()
            .map(RetainedPayload::as_bytes)
    }

    /// Whether the source package contained signature infrastructure when it
    /// was read.  This includes an incomplete or orphaned signature graph.
    #[must_use]
    pub const fn signature_present(&self) -> bool {
        self.signature_present
    }

    /// Fingerprint of every retained package relationship edge incoming to
    /// the workbook/worksheet/control closure.
    #[must_use]
    pub const fn incoming_relationships(&self) -> RelationshipFingerprint {
        self.incoming_relationships
    }

    /// Selected DrawingML payload and sidecar fingerprint.
    #[must_use]
    pub const fn drawing(&self) -> Option<&FormControlPartRead> {
        self.drawing.as_ref()
    }

    /// Selected VML payload and sidecar fingerprint.
    #[must_use]
    pub const fn vml(&self) -> Option<&FormControlPartRead> {
        self.vml.as_ref()
    }

    /// Selected control-properties payloads in effective control order.
    #[must_use]
    pub fn properties(&self) -> &[FormControlPartRead] {
        &self.properties
    }

    /// Internal targets reached from retained DrawingML, VML, or selected
    /// properties sidecar relationship members.  Each target is retained as a
    /// bounded source payload so the sidecar edge cannot silently disappear
    /// between a read and a later source-backed operation.
    #[must_use]
    pub fn sidecar_targets(&self) -> &[FormControlPartRead] {
        &self.sidecar_targets
    }

    /// Worksheet MCE raw/selected/inactive provenance.
    #[must_use]
    pub const fn worksheet_mce(&self) -> &MceProvenance {
        &self.worksheet_mce
    }

    /// Drawing MCE raw/selected/inactive provenance.
    #[must_use]
    pub const fn drawing_mce(&self) -> &MceProvenance {
        &self.drawing_mce
    }

    /// Whether this read set keeps parser-owned semantic reservations alive.
    #[must_use]
    pub fn has_execution_budget(&self) -> bool {
        self.budget_hold
            .as_ref()
            .is_some_and(|hold| hold.reservation_count() != 0)
            || !self.semantic_holds.is_empty()
    }
}

/// Check every retained physical member and relationship sidecar against an
/// in-memory OPC package before applying a source-backed patch.  The source
/// owner intentionally retains the exact relationship XML, including an
/// explicit empty member versus absence, so a semantically equivalent graph
/// rewrite is still a stale read-set refusal.
pub(crate) fn read_set_matches_package_except(
    package: &OpcPackage,
    read_set: &FormControlReadSet,
    excluded: &[&PackURI],
) -> Result<bool> {
    fn matches_part(
        package: &OpcPackage,
        part: &FormControlPartRead,
        excluded: &[&PackURI],
    ) -> Result<bool> {
        let Some(name) = part.part_name() else {
            return Ok(false);
        };
        let payload_excluded = excluded.contains(&name);
        // A source-backed patch must turn a disappeared owner member into a
        // normal stale-read result.  Propagating `PartNotFound` would expose
        // an OPC lookup failure instead of the patch-conflict contract, and
        // would make the missing-member path inconsistent with byte or
        // relationship mismatches below.
        let current = match package.get_part(name) {
            Ok(current) => current,
            Err(litchi_opc::OpcError::PartNotFound(_)) => return Ok(false),
            Err(error) => return Err(error.into()),
        };
        if !payload_excluded && current.blob() != part.bytes() {
            return Ok(false);
        }
        let current_relationships =
            package.source_relationships_with_limits(name, package.read_limits())?;
        if current_relationships.member_present() != part.relationship_bytes().is_some() {
            return Ok(false);
        }
        if let Some(expected) = part.relationship_bytes()
            && current_relationships.bytes() != expected
        {
            return Ok(false);
        }
        Ok(true)
    }

    if !read_set
        .workbook
        .as_ref()
        .map_or(Ok(true), |part| matches_part(package, part, excluded))?
        || !matches_part(package, &read_set.worksheet, excluded)?
        || !read_set
            .drawing
            .as_ref()
            .map_or(Ok(true), |part| matches_part(package, part, excluded))?
        || !read_set
            .vml
            .as_ref()
            .map_or(Ok(true), |part| matches_part(package, part, excluded))?
    {
        return Ok(false);
    }
    let root = PackURI::new("/").map_err(|error| invalid(error.to_string()))?;
    let root_relationships =
        package.source_relationships_with_limits(&root, package.read_limits())?;
    if root_relationships.member_present() != read_set.package_relationship_payload.is_some()
        || read_set
            .package_relationship_payload
            .as_ref()
            .is_some_and(|expected| root_relationships.bytes() != expected.as_bytes())
        || root_relationships.member_present() != read_set.package_relationships.present()
        || root_relationships.bytes().len() != read_set.package_relationships.bytes()
    {
        return Ok(false);
    }
    if read_set.signature_present
        != (package.is_signed() || package.requires_signature_edit_policy())
    {
        return Ok(false);
    }
    if let Some(expected) = read_set.content_types_payload.as_ref() {
        let current = package.source_content_types_with_limits(package.read_limits())?;
        if current.bytes() != expected.as_bytes() {
            return Ok(false);
        }
    }
    if incoming_fingerprint_from_package(package, read_set).map_err(owner_to_xlsx)?
        != read_set.incoming_relationships
    {
        return Ok(false);
    }
    for part in read_set.properties.iter() {
        if !matches_part(package, part, excluded)? {
            return Ok(false);
        }
    }
    for part in read_set.sidecar_targets.iter() {
        if !matches_part(package, part, excluded)? {
            return Ok(false);
        }
    }
    Ok(true)
}

#[derive(Clone, Debug, Eq, Ord, PartialEq, PartialOrd)]
struct IncomingEdgeKey {
    owner: String,
    target: String,
    relationship_id: String,
    relationship_type: String,
    target_ref: String,
    target_mode: u8,
}

struct IncomingFingerprintBuilder {
    edges: Vec<IncomingEdgeKey>,
    maximum: usize,
}

impl IncomingFingerprintBuilder {
    fn new(maximum: usize) -> OwnerResult<Self> {
        let mut edges = Vec::new();
        edges
            .try_reserve(maximum.min(16))
            .map_err(|source| owner_alloc("incoming relationship census", source))?;
        Ok(Self { edges, maximum })
    }

    fn push(
        &mut self,
        owner: &PackURI,
        target: &PackURI,
        relationship: &litchi_opc::Relationship,
    ) -> OwnerResult<()> {
        let observed = self
            .edges
            .len()
            .checked_add(1)
            .ok_or_else(|| owner_invalid("incoming relationship edge count overflow"))?;
        if observed > self.maximum {
            return Err(owner_limit(
                "incoming relationship edges",
                observed,
                self.maximum,
            ));
        }
        self.edges
            .try_reserve(1)
            .map_err(|source| owner_alloc("incoming relationship census", source))?;
        self.edges.push(IncomingEdgeKey {
            owner: owner.to_string(),
            target: target.to_string(),
            relationship_id: relationship.r_id().to_owned(),
            relationship_type: relationship.reltype().to_owned(),
            target_ref: relationship.target_ref().to_owned(),
            target_mode: match relationship.target_mode() {
                TargetMode::Internal => 0,
                TargetMode::External => 1,
            },
        });
        Ok(())
    }

    fn finish(mut self) -> RelationshipFingerprint {
        self.edges.sort_unstable();
        let mut digest = 0xcbf29ce484222325u64;
        for edge in &self.edges {
            digest = fnv_bytes(digest, edge.owner.as_bytes());
            digest = fnv_bytes(digest, &[0]);
            digest = fnv_bytes(digest, edge.target.as_bytes());
            digest = fnv_bytes(digest, &[0]);
            digest = fnv_bytes(digest, edge.relationship_id.as_bytes());
            digest = fnv_bytes(digest, &[0]);
            digest = fnv_bytes(digest, edge.relationship_type.as_bytes());
            digest = fnv_bytes(digest, &[0]);
            digest = fnv_bytes(digest, edge.target_ref.as_bytes());
            digest = fnv_bytes(digest, &[edge.target_mode]);
        }
        RelationshipFingerprint {
            present: true,
            bytes: 0,
            digest,
            edges: self.edges.len(),
        }
    }
}

fn read_set_contains_target(
    read_set: &FormControlReadSet,
    workbook_name: Option<&PackURI>,
    target: &PackURI,
) -> bool {
    workbook_name.is_some_and(|name| name == target)
        || read_set
            .workbook
            .as_ref()
            .and_then(FormControlPartRead::part_name)
            .is_some_and(|name| name == target)
        || read_set
            .worksheet
            .part_name()
            .is_some_and(|name| name == target)
        || read_set
            .drawing
            .as_ref()
            .and_then(FormControlPartRead::part_name)
            .is_some_and(|name| name == target)
        || read_set
            .vml
            .as_ref()
            .and_then(FormControlPartRead::part_name)
            .is_some_and(|name| name == target)
        || read_set
            .properties
            .iter()
            .any(|part| part.part_name().is_some_and(|name| name == target))
        || read_set
            .sidecar_targets
            .iter()
            .any(|part| part.part_name().is_some_and(|name| name == target))
}

fn incoming_fingerprint_from_source(
    package: &SourceBackedPackage,
    read_set: &FormControlReadSet,
    workbook_name: Option<&PackURI>,
    maximum: usize,
) -> OwnerResult<RelationshipFingerprint> {
    let mut builder = IncomingFingerprintBuilder::new(maximum)?;
    let package_uri = PackURI::new("/").map_err(|error| owner_invalid(error.to_string()))?;
    for relationship in package.rels().iter() {
        if relationship.target_mode() == TargetMode::Internal {
            let target = relationship.target_partname()?;
            if read_set_contains_target(read_set, workbook_name, &target) {
                builder.push(&package_uri, &target, relationship)?;
            }
        }
    }
    for part in package.iter_parts() {
        let owner = part.partname();
        for relationship in part.rels().iter() {
            if relationship.target_mode() == TargetMode::Internal {
                let target = relationship.target_partname()?;
                if read_set_contains_target(read_set, workbook_name, &target) {
                    builder.push(owner, &target, relationship)?;
                }
            }
        }
    }
    Ok(builder.finish())
}

fn incoming_fingerprint_from_package(
    package: &OpcPackage,
    read_set: &FormControlReadSet,
) -> OwnerResult<RelationshipFingerprint> {
    let maximum = package.read_limits().max_relationship_graph_nodes();
    let mut builder = IncomingFingerprintBuilder::new(maximum)?;
    let package_uri = PackURI::new("/").map_err(|error| owner_invalid(error.to_string()))?;
    for relationship in package.rels().iter() {
        if relationship.target_mode() == TargetMode::Internal {
            let target = relationship.target_partname()?;
            if read_set_contains_target(read_set, None, &target) {
                builder.push(&package_uri, &target, relationship)?;
            }
        }
    }
    for part in package.iter_parts() {
        let owner = part.partname();
        for relationship in part.rels().iter() {
            if relationship.target_mode() == TargetMode::Internal {
                let target = relationship.target_partname()?;
                if read_set_contains_target(read_set, None, &target) {
                    builder.push(owner, &target, relationship)?;
                }
            }
        }
    }
    Ok(builder.finish())
}

/// Diagnostic category carried by an owner collection or one control.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum FormControlDiagnosticCode {
    /// A control-properties part has no effective incoming worksheet owner.
    UnreferencedControlProperties,
    /// A graph branch was retained but is not admitted to the typed profile.
    OpaqueGraph,
    /// A shape or relationship identity is ambiguous.
    AmbiguousIdentity,
    /// A source mirror is not proven by the admitted profile.
    UnprovenMirror,
    /// An MCE wrapper could not be selected deterministically.
    MarkupCompatibility,
    /// A strict or mixed worksheet dialect was refused.
    UnsupportedDialect,
    /// The optional SpreadsheetML control-pr anchor was absent or malformed.
    ControlPrAnchor,
}

/// One inert diagnostic.  Details are explanatory and never interpreted as a
/// selector or package path by the public facade.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FormControlDiagnostic {
    code: FormControlDiagnosticCode,
    detail: Box<str>,
}

impl FormControlDiagnostic {
    fn new(code: FormControlDiagnosticCode, detail: impl Into<Box<str>>) -> Self {
        Self {
            code,
            detail: detail.into(),
        }
    }

    /// Diagnostic category.
    #[must_use]
    pub const fn code(&self) -> FormControlDiagnosticCode {
        self.code
    }

    /// Human-readable inert detail.
    #[must_use]
    pub fn detail(&self) -> &str {
        &self.detail
    }
}

/// DrawingML/VML identity closure for one effective control.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ShapeClosure {
    shape_id: u64,
    drawing_name: Option<Box<str>>,
    compat_spid: Box<str>,
    vml_id: Box<str>,
    vml_spid: Option<Box<str>>,
    object_type: Box<str>,
    vml_object_type: Box<str>,
    drawing_relationships_present: bool,
    vml_relationships_present: bool,
    drawing_relationships: RelationshipFingerprint,
    vml_relationships: RelationshipFingerprint,
}

impl ShapeClosure {
    /// Worksheet/DrawingML numeric identity.
    #[must_use]
    pub const fn shape_id(&self) -> u64 {
        self.shape_id
    }

    /// DrawingML `cNvPr/@name`, if authored.
    #[must_use]
    pub fn drawing_name(&self) -> Option<&str> {
        self.drawing_name.as_deref()
    }

    /// Selected `a14:compatExt/@spid` lexical value.
    #[must_use]
    pub fn compat_spid(&self) -> &str {
        &self.compat_spid
    }

    /// VML `v:shape/@id` lexical value.
    #[must_use]
    pub fn vml_id(&self) -> &str {
        &self.vml_id
    }

    /// VML `o:spid`, if authored.
    #[must_use]
    pub fn vml_spid(&self) -> Option<&str> {
        self.vml_spid.as_deref()
    }

    /// Typed x14 `formControlPr/@objectType` token.
    #[must_use]
    pub fn object_type(&self) -> &str {
        &self.object_type
    }

    /// VML `ClientData/@ObjectType` token retained by the closure.
    #[must_use]
    pub fn vml_object_type(&self) -> &str {
        &self.vml_object_type
    }

    /// Whether the DrawingML part had a physical `.rels` member.
    #[must_use]
    pub const fn drawing_relationships_present(&self) -> bool {
        self.drawing_relationships_present
    }

    /// Whether the VML part had a physical `.rels` member.
    #[must_use]
    pub const fn vml_relationships_present(&self) -> bool {
        self.vml_relationships_present
    }

    /// Fingerprint of the DrawingML sidecar relationship member.
    #[must_use]
    pub const fn drawing_relationships(&self) -> RelationshipFingerprint {
        self.drawing_relationships
    }

    /// Fingerprint of the VML sidecar relationship member.
    #[must_use]
    pub const fn vml_relationships(&self) -> RelationshipFingerprint {
        self.vml_relationships
    }
}

/// One typed, source-backed form-control projection.
#[derive(Clone, Debug)]
pub struct FormControlView {
    position: usize,
    property_part: Option<PackURI>,
    vml_part: Option<PackURI>,
    name: Option<Arc<str>>,
    anchor_profile: Option<Arc<str>>,
    properties: Arc<Properties>,
    properties_relationships_present: bool,
    properties_relationships: RelationshipFingerprint,
    shape: Arc<ShapeClosure>,
    diagnostics: Arc<[FormControlDiagnostic]>,
    read_set: Option<Arc<FormControlReadSet>>,
    budget_hold: Option<Arc<OwnerBudgetHold>>,
}

impl PartialEq for FormControlView {
    fn eq(&self, other: &Self) -> bool {
        self.position == other.position
            && self.name == other.name
            && self.anchor_profile == other.anchor_profile
            && self.properties == other.properties
            && self.properties_relationships_present == other.properties_relationships_present
            && self.properties_relationships == other.properties_relationships
            && self.shape == other.shape
            && self.diagnostics == other.diagnostics
    }
}

impl Eq for FormControlView {}

impl FormControlView {
    pub(crate) fn source_property_bytes(&self) -> Option<&[u8]> {
        self.properties.source_bytes()
    }

    pub(crate) fn source_property_payload(&self) -> Option<SourcePayload> {
        self.properties.source_payload()
    }

    pub(crate) fn property_part_name(&self) -> Option<&PackURI> {
        self.property_part.as_ref().or_else(|| {
            self.read_set
                .as_deref()
                .and_then(|read_set| read_set.property_part_name(self.position))
        })
    }

    pub(crate) fn vml_part_name(&self) -> Option<&PackURI> {
        self.vml_part.as_ref().or_else(|| {
            self.read_set
                .as_deref()
                .and_then(FormControlReadSet::vml_part_name)
        })
    }

    /// Effective zero-based selector position.
    #[must_use]
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Return the selector represented by this entry.
    #[must_use]
    pub fn selector(&self) -> ControlSelector<'_> {
        ControlSelector::Position(self.position)
    }

    /// Worksheet `control/@name`, if present.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Matched SpreadsheetML control-pr anchor ancestry, if the optional
    /// `sml:controlPr/sml:anchor` shape is a canonical `LoSmlAnchorV1`.
    #[must_use]
    pub fn anchor_profile(&self) -> Option<&str> {
        self.anchor_profile.as_deref()
    }

    /// Typed `formControlPr` properties.
    #[must_use]
    pub fn properties(&self) -> &Properties {
        self.properties.as_ref()
    }

    /// Whether the selected properties part had a physical `.rels` member.
    #[must_use]
    pub const fn properties_relationships_present(&self) -> bool {
        self.properties_relationships_present
    }

    /// Fingerprint of the selected control-properties relationship member.
    #[must_use]
    pub const fn properties_relationships(&self) -> RelationshipFingerprint {
        self.properties_relationships
    }

    /// DrawingML/VML identity closure.
    #[must_use]
    pub fn shape(&self) -> &ShapeClosure {
        self.shape.as_ref()
    }

    /// Inert diagnostics for this source projection.
    #[must_use]
    pub fn diagnostics(&self) -> &[FormControlDiagnostic] {
        self.diagnostics.as_ref()
    }

    /// Scoped source read set retained by this singular view, when the view
    /// came from a deferred source-backed package.
    #[must_use]
    pub fn read_set(&self) -> Option<&FormControlReadSet> {
        self.read_set.as_deref()
    }

    /// Whether this view keeps caller execution reservations alive.
    #[must_use]
    pub fn has_execution_budget(&self) -> bool {
        self.budget_hold
            .as_ref()
            .is_some_and(|hold| hold.reservation_count() != 0)
    }
}

/// Immutable selector-first collection of effective worksheet controls.
#[derive(Clone, Debug)]
pub struct FormControlCollection {
    profile: OwnerProfile,
    controls: Arc<[FormControlView]>,
    diagnostics: Arc<[FormControlDiagnostic]>,
    read_set: Option<Arc<FormControlReadSet>>,
    budget_hold: Option<Arc<OwnerBudgetHold>>,
    generated_budget: Option<Arc<RetainedBudgetHold>>,
}

impl PartialEq for FormControlCollection {
    fn eq(&self, other: &Self) -> bool {
        self.profile == other.profile
            && self.controls == other.controls
            && self.diagnostics == other.diagnostics
    }
}

impl Eq for FormControlCollection {}

impl FormControlCollection {
    fn attach_source_package_readset(
        mut self,
        package: SourcePackageReadSet,
        incoming_relationships: RelationshipFingerprint,
        limits: OwnerLimits,
    ) -> OwnerResult<Self> {
        let Some(existing) = self.read_set.as_deref() else {
            return Ok(self);
        };
        let added = retained_part_bytes(&package.workbook)?
            .checked_add(
                package
                    .package_relationship_payload
                    .as_ref()
                    .map_or(0, |payload| payload.as_bytes().len()),
            )
            .and_then(|bytes| bytes.checked_add(package.content_types_payload.as_bytes().len()))
            .ok_or_else(|| owner_invalid("package read-set byte count overflow"))?;
        let mut retained = retained_read_set_bytes(existing)?;
        retained = retained
            .checked_add(added)
            .ok_or_else(|| owner_invalid("package read-set byte count overflow"))?;
        if retained > limits.max_read_set_bytes {
            return Err(owner_limit(
                "form-control read-set bytes",
                retained,
                limits.max_read_set_bytes,
            ));
        }
        let mut updated = existing.clone();
        updated.workbook = Some(package.workbook.into_public());
        updated.package_relationship_payload = package.package_relationship_payload;
        updated.package_relationships = package.package_relationships;
        updated.content_types_payload = Some(package.content_types_payload);
        updated.signature_present = package.signature_present;
        updated.incoming_relationships = incoming_relationships;
        let read_set = Arc::new(updated);
        let mut controls = self.controls.to_vec();
        for control in &mut controls {
            control.read_set = Some(Arc::clone(&read_set));
        }
        self.controls = Arc::from(controls.into_boxed_slice());
        self.read_set = Some(read_set);
        Ok(self)
    }

    pub(crate) fn with_replaced_properties(
        &self,
        position: usize,
        properties: Properties,
    ) -> Option<Self> {
        let mut controls = self.controls.to_vec();
        let control = controls.get_mut(position)?;
        control.properties = Arc::new(properties);
        Some(Self {
            profile: self.profile,
            controls: Arc::from(controls.into_boxed_slice()),
            diagnostics: Arc::clone(&self.diagnostics),
            read_set: self.read_set.clone(),
            budget_hold: self.budget_hold.clone(),
            generated_budget: self.generated_budget.clone(),
        })
    }

    pub(crate) fn with_generated_budget(mut self, budget: Option<Arc<RetainedBudgetHold>>) -> Self {
        self.generated_budget = budget;
        self
    }

    pub(crate) fn generated_clone_uri_storage_bytes(&self) -> Option<usize> {
        self.controls.iter().try_fold(0usize, |total, control| {
            let property = control
                .property_part
                .as_ref()
                .map_or(0, |part| part.as_str().len() + size_of::<String>());
            let vml = control
                .vml_part
                .as_ref()
                .map_or(0, |part| part.as_str().len() + size_of::<String>());
            total.checked_add(property)?.checked_add(vml)
        })
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        match (self.read_set.as_deref(), other.read_set.as_deref()) {
            (Some(left), Some(right)) => left.same_source(right),
            (None, None) => self == other,
            _ => false,
        }
    }

    /// Number of effective typed controls.
    #[must_use]
    pub fn len(&self) -> usize {
        let _ = self
            .budget_hold
            .as_ref()
            .map(|hold| hold.reservation_count());
        self.controls.len()
    }

    /// Whether no effective controls were admitted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.controls.is_empty()
    }

    /// Profile used for the projection.
    #[must_use]
    pub const fn profile(&self) -> OwnerProfile {
        self.profile
    }

    /// Collection-level graph diagnostics.
    #[must_use]
    pub fn diagnostics(&self) -> &[FormControlDiagnostic] {
        self.diagnostics.as_ref()
    }

    /// Scoped source read set, present for deferred source-backed reads.
    #[must_use]
    pub fn read_set(&self) -> Option<&FormControlReadSet> {
        self.read_set.as_deref()
    }

    /// Source version captured by this collection, if source-backed.
    #[must_use]
    pub fn source_version(&self) -> Option<litchi_core::SourceVersion> {
        self.read_set
            .as_deref()
            .and_then(FormControlReadSet::source_version)
    }

    /// Whether this collection retains caller execution reservations for its
    /// shared semantic, projection, and source read-set state.
    #[must_use]
    pub fn has_execution_budget(&self) -> bool {
        self.budget_hold
            .as_ref()
            .is_some_and(|hold| hold.reservation_count() != 0)
            || self
                .read_set
                .as_deref()
                .is_some_and(FormControlReadSet::has_execution_budget)
            || self.generated_budget.is_some()
    }

    /// Iterate controls in effective worksheet source order.
    pub fn iter(
        &self,
    ) -> impl ExactSizeIterator<Item = &FormControlView> + DoubleEndedIterator + '_ {
        self.controls.iter()
    }

    /// Resolve a semantic position or exact name.
    pub fn get<'a>(
        &self,
        selector: impl Into<ControlSelector<'a>>,
    ) -> Result<Option<&FormControlView>> {
        match selector.into() {
            ControlSelector::Position(position) => Ok(self.controls.get(position)),
            ControlSelector::Name(name) => {
                let mut found = None;
                for control in self.controls.iter() {
                    if control.name() == Some(name) {
                        if found.is_some() {
                            return Err(invalid("form-control name selector is ambiguous"));
                        }
                        found = Some(control);
                    }
                }
                Ok(found)
            },
        }
    }
}

fn retained_read_set_bytes(read_set: &FormControlReadSet) -> OwnerResult<usize> {
    let mut total = retained_part_bytes_public(&read_set.worksheet)?;
    if let Some(part) = read_set.workbook.as_ref() {
        total = total
            .checked_add(retained_part_bytes_public(part)?)
            .ok_or_else(|| owner_invalid("form-control read-set byte count overflow"))?;
    }
    if let Some(part) = read_set.drawing.as_ref() {
        total = total
            .checked_add(retained_part_bytes_public(part)?)
            .ok_or_else(|| owner_invalid("form-control read-set byte count overflow"))?;
    }
    if let Some(part) = read_set.vml.as_ref() {
        total = total
            .checked_add(retained_part_bytes_public(part)?)
            .ok_or_else(|| owner_invalid("form-control read-set byte count overflow"))?;
    }
    for part in read_set
        .properties
        .iter()
        .chain(read_set.sidecar_targets.iter())
    {
        total = total
            .checked_add(retained_part_bytes_public(part)?)
            .ok_or_else(|| owner_invalid("form-control read-set byte count overflow"))?;
    }
    total = total
        .checked_add(
            read_set
                .package_relationship_payload
                .as_ref()
                .map_or(0, |payload| payload.as_bytes().len()),
        )
        .ok_or_else(|| owner_invalid("form-control read-set byte count overflow"))?;
    total = total
        .checked_add(
            read_set
                .content_types_payload
                .as_ref()
                .map_or(0, |payload| payload.as_bytes().len()),
        )
        .ok_or_else(|| owner_invalid("form-control read-set byte count overflow"))?;
    Ok(total)
}

fn retained_part_bytes_public(part: &FormControlPartRead) -> OwnerResult<usize> {
    part.bytes()
        .len()
        .checked_add(part.relationship_bytes().map_or(0, <[u8]>::len))
        .ok_or_else(|| owner_invalid("form-control read-set byte count overflow"))
}

impl<'a> IntoIterator for &'a FormControlCollection {
    type Item = &'a FormControlView;
    type IntoIter = std::slice::Iter<'a, FormControlView>;

    fn into_iter(self) -> Self::IntoIter {
        self.controls.iter()
    }
}

/// Failure returned by the worksheet owner scanner.
#[derive(Debug, thiserror::Error)]
#[non_exhaustive]
pub enum FormControlOwnerError {
    /// A bounded owner graph is malformed or outside the admitted profile.
    #[error("invalid XLSX form-control owner graph: {0}")]
    Invalid(String),
    /// A host resource ceiling was exceeded before retaining more data.
    #[error("XLSX form-control owner {resource} exceeds {maximum} (observed {observed})")]
    Limit {
        resource: &'static str,
        observed: usize,
        maximum: usize,
    },
    /// A bounded owner allocation could not be reserved.
    #[error("could not reserve memory for XLSX form-control owner {resource}: {source}")]
    Allocation {
        resource: &'static str,
        #[source]
        source: std::collections::TryReserveError,
    },
    /// The leaf properties codec refused a selected part.
    #[error(transparent)]
    Leaf(#[from] FormControlError),
    /// The underlying OPC source refused a bounded read.
    #[error(transparent)]
    Package(#[from] litchi_opc::OpcError),
    /// A caller-owned execution context cancelled or bounded this scan.
    #[error(transparent)]
    Execution(#[from] litchi_core::ExecutionError),
    /// The source-backed XLSX facade returned a worksheet error.
    #[error(transparent)]
    Xlsx(#[from] Error),
}

/// Result type for direct owner operations.
pub type OwnerResult<T> = std::result::Result<T, FormControlOwnerError>;

/// Outstanding charges retained by a source-backed owner result.  Each
/// resource has one merged reservation so cloning a collection or one of its
/// views never allocates a second budget ledger.
#[derive(Debug)]
struct OwnerBudgetHold {
    memory: Option<litchi_core::Reservation>,
    input_bytes: Option<litchi_core::Reservation>,
    output_bytes: Option<litchi_core::Reservation>,
    objects: Option<litchi_core::Reservation>,
    depth: Option<litchi_core::Reservation>,
}

impl OwnerBudgetHold {
    fn reservation_count(&self) -> usize {
        usize::from(self.memory.is_some())
            + usize::from(self.input_bytes.is_some())
            + usize::from(self.output_bytes.is_some())
            + usize::from(self.objects.is_some())
            + usize::from(self.depth.is_some())
    }
}

struct OwnerExecution {
    context: Option<litchi_core::ExecutionContext>,
    memory: Option<litchi_core::Reservation>,
    input_bytes: Option<litchi_core::Reservation>,
    output_bytes: Option<litchi_core::Reservation>,
    objects: Option<litchi_core::Reservation>,
    depth: Option<litchi_core::Reservation>,
    semantic_bytes: usize,
    projection_bytes: usize,
    read_set_bytes: usize,
    sidecar_target_bytes: usize,
}

#[derive(Debug)]
struct EagerUnreferencedInventory {
    parts: Vec<String>,
    // These fields intentionally have no accessors: retaining the reservation
    // for the wrapper's lifetime keeps the eager strings charged while the
    // owner scans the same inventory.
    _memory: Option<litchi_core::Reservation>,
    _objects: Option<litchi_core::Reservation>,
    projection_bytes: usize,
}

impl EagerUnreferencedInventory {
    fn as_slice(&self) -> &[String] {
        &self.parts
    }

    fn projection_bytes(&self) -> usize {
        self.projection_bytes
    }
}

#[derive(Clone, Copy)]
enum OwnerMemoryCategory {
    Semantic,
    Projection,
    ReadSet,
    SidecarTarget,
}

impl OwnerExecution {
    fn new(
        package: Option<&SourceBackedPackage>,
        explicit_context: Option<litchi_core::ExecutionContext>,
    ) -> Self {
        Self {
            context: explicit_context
                .or_else(|| package.and_then(SourceBackedPackage::execution_context)),
            memory: None,
            input_bytes: None,
            output_bytes: None,
            objects: None,
            depth: None,
            semantic_bytes: 0,
            projection_bytes: 0,
            read_set_bytes: 0,
            sidecar_target_bytes: 0,
        }
    }

    fn check(&self) -> OwnerResult<()> {
        if let Some(context) = self.context.as_ref() {
            context.check().map_err(FormControlOwnerError::from)?;
        }
        Ok(())
    }

    fn context(&self) -> Option<&litchi_core::ExecutionContext> {
        self.context.as_ref()
    }

    fn work(&self, amount: usize) -> OwnerResult<()> {
        self.check()?;
        if let Some(context) = self.context.as_ref() {
            context
                .consume(
                    litchi_core::Resource::Work,
                    u64::try_from(amount).unwrap_or(u64::MAX),
                )
                .map_err(FormControlOwnerError::from)?;
        }
        Ok(())
    }

    fn reserve(&mut self, resource: litchi_core::Resource, amount: usize) -> OwnerResult<()> {
        self.check()?;
        if amount == 0 {
            return Ok(());
        }
        let Some(context) = self.context.as_ref() else {
            return Ok(());
        };
        let reservation = context
            .reserve(resource, u64::try_from(amount).unwrap_or(u64::MAX))
            .map_err(FormControlOwnerError::from)?;
        let slot = match resource {
            litchi_core::Resource::Memory => &mut self.memory,
            litchi_core::Resource::InputBytes => &mut self.input_bytes,
            litchi_core::Resource::OutputBytes => &mut self.output_bytes,
            litchi_core::Resource::Objects => &mut self.objects,
            litchi_core::Resource::Depth => &mut self.depth,
            litchi_core::Resource::Work => {
                drop(reservation);
                return self.work(amount);
            },
            _ => {
                drop(reservation);
                return Err(owner_invalid("unsupported owner execution resource"));
            },
        };
        if let Some(existing) = slot.as_mut() {
            if existing.try_merge(reservation).is_err() {
                return Err(owner_invalid(
                    "owner execution reservations do not share a budget",
                ));
            }
        } else {
            *slot = Some(reservation);
        }
        Ok(())
    }

    fn reserve_memory_category(
        &mut self,
        category: OwnerMemoryCategory,
        amount: usize,
        maximum: usize,
        resource: &'static str,
    ) -> OwnerResult<()> {
        let current = match category {
            OwnerMemoryCategory::Semantic => self.semantic_bytes,
            OwnerMemoryCategory::Projection => self.projection_bytes,
            OwnerMemoryCategory::ReadSet => self.read_set_bytes,
            OwnerMemoryCategory::SidecarTarget => self.sidecar_target_bytes,
        };
        let observed = current
            .checked_add(amount)
            .ok_or_else(|| owner_invalid(format!("{resource} byte count overflow")))?;
        if observed > maximum {
            return Err(owner_limit(resource, observed, maximum));
        }
        self.reserve(litchi_core::Resource::Memory, amount)?;
        match category {
            OwnerMemoryCategory::Semantic => self.semantic_bytes = observed,
            OwnerMemoryCategory::Projection => self.projection_bytes = observed,
            OwnerMemoryCategory::ReadSet => self.read_set_bytes = observed,
            OwnerMemoryCategory::SidecarTarget => self.sidecar_target_bytes = observed,
        }
        Ok(())
    }

    /// Account memory that was already charged by a shared child parser.
    ///
    /// Source-backed leaf parsing owns its semantic reservation so that the
    /// `Properties` clone can keep that reservation alive.  The owner still
    /// needs the bytes in its aggregate semantic ceiling, but must not reserve
    /// the same execution budget a second time.  This method updates only the
    /// owner category ledger; callers use [`Self::reserve`] for allocations
    /// that are owned directly by this scanner.
    fn account_memory_category(
        &mut self,
        category: OwnerMemoryCategory,
        amount: usize,
        maximum: usize,
        resource: &'static str,
    ) -> OwnerResult<()> {
        let current = match category {
            OwnerMemoryCategory::Semantic => self.semantic_bytes,
            OwnerMemoryCategory::Projection => self.projection_bytes,
            OwnerMemoryCategory::ReadSet => self.read_set_bytes,
            OwnerMemoryCategory::SidecarTarget => self.sidecar_target_bytes,
        };
        let observed = current
            .checked_add(amount)
            .ok_or_else(|| owner_invalid(format!("{resource} byte count overflow")))?;
        if observed > maximum {
            return Err(owner_limit(resource, observed, maximum));
        }
        match category {
            OwnerMemoryCategory::Semantic => self.semantic_bytes = observed,
            OwnerMemoryCategory::Projection => self.projection_bytes = observed,
            OwnerMemoryCategory::ReadSet => self.read_set_bytes = observed,
            OwnerMemoryCategory::SidecarTarget => self.sidecar_target_bytes = observed,
        }
        Ok(())
    }

    fn remaining_memory_category(&self, category: OwnerMemoryCategory, maximum: usize) -> usize {
        let used = match category {
            OwnerMemoryCategory::Semantic => self.semantic_bytes,
            OwnerMemoryCategory::Projection => self.projection_bytes,
            OwnerMemoryCategory::ReadSet => self.read_set_bytes,
            OwnerMemoryCategory::SidecarTarget => self.sidecar_target_bytes,
        };
        maximum.saturating_sub(used)
    }

    fn reserve_capacity<T>(
        &mut self,
        category: OwnerMemoryCategory,
        capacity: usize,
        maximum: usize,
        resource: &'static str,
    ) -> OwnerResult<()> {
        let bytes = capacity
            .checked_mul(size_of::<T>())
            .ok_or_else(|| owner_invalid(format!("{resource} capacity overflow")))?;
        self.reserve_memory_category(category, bytes, maximum, resource)
    }

    /// Reconcile a preflight capacity charge with the allocator's actual
    /// post-reservation capacity.  `try_reserve_exact` guarantees a minimum,
    /// but an allocator may legally return a larger backing allocation; that
    /// excess must enter the same aggregate owner ledger before the vector is
    /// published.
    fn reconcile_capacity<T>(
        &mut self,
        category: OwnerMemoryCategory,
        preflight_capacity: usize,
        actual_capacity: usize,
        maximum: usize,
        resource: &'static str,
    ) -> OwnerResult<()> {
        if actual_capacity <= preflight_capacity {
            return Ok(());
        }
        self.reserve_capacity::<T>(
            category,
            actual_capacity - preflight_capacity,
            maximum,
            resource,
        )
    }

    fn finish(&mut self) -> Option<Arc<OwnerBudgetHold>> {
        self.context.as_ref()?;
        Some(Arc::new(OwnerBudgetHold {
            memory: self.memory.take(),
            input_bytes: self.input_bytes.take(),
            output_bytes: self.output_bytes.take(),
            objects: self.objects.take(),
            depth: self.depth.take(),
        }))
    }
}

/// Source-backed read-only owner for one XLSX package.
pub struct SourceBackedFormControlOwner {
    package: SourceBackedPackage,
    limits: OwnerLimits,
}

impl SourceBackedFormControlOwner {
    /// Open a deferred OPC package from a reader and retain the source-backed
    /// package for subsequent worksheet selections.
    pub fn from_reader<R: Read>(reader: R) -> OwnerResult<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_reader(reader)?)
    }

    /// Open a deferred OPC package from a filesystem path.
    #[cfg(any(unix, windows))]
    pub fn from_path(path: impl AsRef<Path>) -> OwnerResult<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_path(path)?)
    }

    /// Build an owner from an already opened deferred OPC package.
    pub fn from_source_backed_package(package: SourceBackedPackage) -> OwnerResult<Self> {
        package.check_execution()?;
        if package.has_encrypted_entries() {
            return Err(FormControlOwnerError::Invalid(
                "encrypted XLSX form-control source is not admitted".into(),
            ));
        }
        let limits = OwnerLimits::from_read_limits(package.read_limits());
        Ok(Self { package, limits })
    }

    /// Set a lower host policy before selecting a worksheet.
    #[must_use]
    pub fn with_limits(mut self, limits: OwnerLimits) -> Self {
        self.limits = limits;
        self
    }

    /// Capture one worksheet's effective form-control collection.
    pub fn form_controls<'a>(
        &self,
        selector: impl Into<Selector<'a>>,
    ) -> OwnerResult<FormControlCollection> {
        self.package.check_execution()?;
        let workbook = self.package.main_document_part()?;
        let workbook_data =
            source_part_data_with_limit(&workbook, self.limits.max_mce_bytes, "workbook bytes")?;
        let catalog =
            raw::parse_catalog(workbook_data.as_bytes()).map_err(FormControlOwnerError::Xlsx)?;
        let sheet_parts = validate_sheet_graph(&self.package, &workbook, &catalog.sheets)
            .map_err(FormControlOwnerError::Xlsx)?;
        let position = resolve_sheet_position(&catalog.sheets, selector.into())?;
        let sheet = catalog
            .sheets
            .get(position)
            .ok_or_else(|| owner_invalid("worksheet selector did not resolve"))?;
        let part = sheet_parts
            .get(position)
            .ok_or_else(|| owner_invalid("worksheet graph position is absent"))?;
        if part.kind != WorksheetKind::Worksheet {
            return Err(FormControlOwnerError::Invalid(format!(
                "sheet '{}' is not a worksheet",
                sheet.name
            )));
        }
        let source_version = self.package.source_version()?;
        let worksheet = self.package.part(&part.uri)?;
        preflight_minimum_mirror_nodes(worksheet.rels(), &self.limits)?;
        let relationships = worksheet.rels().clone();
        let worksheet_data =
            source_part_data_with_limit(&worksheet, self.limits.max_mce_bytes, "worksheet bytes")?;
        let (relationship_member, relationship_payload) = source_relationship_payload(
            &self.package,
            &part.uri,
            &relationships,
            self.limits.max_sidecar_relationship_bytes,
            self.limits.max_relationship_edges,
            "worksheet relationship bytes",
        )?;
        let worksheet_part = source_loaded_part(
            Some(part.uri.clone()),
            worksheet.content_type(),
            worksheet_data,
            relationships.len(),
            relationship_member,
            source_version,
            &relationships,
            self.limits.max_sidecar_relationship_bytes,
            "worksheet bytes",
            self.limits.max_mce_bytes,
            relationship_payload,
        )?;
        let loader =
            |uri: &PackURI, maximum: usize, resource: &'static str| -> OwnerResult<LoadedPart> {
                let part = self.package.part(uri)?;
                let data = source_part_data_with_limit(&part, maximum, resource)?;
                let relationships = part.rels().clone();
                let relationship_limit = if resource == "control-properties bytes" {
                    self.limits.max_control_properties_relationship_bytes
                } else if resource == "sidecar target bytes" {
                    maximum
                        .checked_sub(data.as_bytes().len())
                        .ok_or_else(|| owner_invalid("sidecar target byte budget underflow"))?
                        .min(self.limits.max_sidecar_relationship_bytes)
                } else {
                    self.limits.max_sidecar_relationship_bytes
                };
                let (relationship_member, relationship_payload) = source_relationship_payload(
                    &self.package,
                    uri,
                    &relationships,
                    relationship_limit,
                    self.limits.max_relationship_edges,
                    "source relationship bytes",
                )?;
                let fingerprint = fingerprint_relationships(
                    &relationships,
                    relationship_member,
                    relationship_limit,
                    resource,
                    relationship_payload.as_ref().map(RetainedPayload::as_bytes),
                )?;
                Ok(LoadedPart {
                    part_name: Some(uri.clone()),
                    content_type: part.content_type().to_owned(),
                    payload: RetainedPayload::source(data),
                    source_version: Some(source_version),
                    relationship_count: relationships.len(),
                    relationships: fingerprint,
                    relationship_payload,
                    relationship_graph: relationships,
                })
            };
        scan_owner(
            worksheet_part,
            &relationships,
            &loader,
            &self.limits,
            Some(&self.package),
            None,
            &[],
            0,
        )
    }

    /// Resolve one semantic control in a selected worksheet.
    pub fn form_control<'a>(
        &self,
        sheet: impl Into<Selector<'a>>,
        selector: impl Into<ControlSelector<'a>>,
    ) -> OwnerResult<Option<FormControlView>> {
        let collection = self.form_controls(sheet)?;
        Ok(collection.get(selector)?.cloned())
    }

    /// Expose source cache diagnostics without exposing package identities.
    #[must_use]
    pub fn cache_diagnostics(&self) -> litchi_opc::SourceCacheDiagnostics {
        self.package.cache_diagnostics()
    }
}

/// Read one eager worksheet while retaining an explicitly supplied execution
/// context for the whole owner scan.  `OpcPackage` predates managed execution
/// contexts, so callers that already own a context must use this seam rather
/// than silently running the owner without cancellation or aggregate charges.
#[allow(
    dead_code,
    reason = "reserved for the explicit eager host execution hook"
)]
pub(crate) fn eager_form_controls_with_execution_context(
    package: &OpcPackage,
    worksheet_uri: &PackURI,
    limits: OwnerLimits,
    execution_context: Option<litchi_core::ExecutionContext>,
) -> OwnerResult<FormControlCollection> {
    if let Some(context) = execution_context.as_ref() {
        context.check().map_err(FormControlOwnerError::from)?;
    }
    let worksheet = package
        .get_part(worksheet_uri)
        .map_err(FormControlOwnerError::from)?;
    if let Some(context) = execution_context.as_ref() {
        context.check().map_err(FormControlOwnerError::from)?;
    }
    preflight_minimum_mirror_nodes(worksheet.rels(), &limits)?;
    let relationships = worksheet.rels().clone();
    let worksheet_bytes = bounded_owned(worksheet.blob(), limits.max_mce_bytes, "worksheet bytes")?;
    let unreferenced_parts =
        eager_unreferenced_properties(package, &limits, execution_context.as_ref())?;
    let (relationship_member, relationship_payload) = eager_relationship_payload(
        package,
        worksheet_uri,
        &relationships,
        limits.max_sidecar_relationship_bytes,
        limits.max_relationship_edges,
        "worksheet relationship bytes",
    )?;
    let worksheet_part = eager_loaded_part(
        Some(worksheet_uri.clone()),
        worksheet.content_type(),
        worksheet_bytes,
        worksheet.rels().len(),
        relationship_member,
        &relationships,
        limits.max_sidecar_relationship_bytes,
        "worksheet bytes",
        relationship_payload,
    )?;
    let loader =
        |uri: &PackURI, maximum: usize, resource: &'static str| -> OwnerResult<LoadedPart> {
            let part = package.get_part(uri).map_err(FormControlOwnerError::from)?;
            let bytes = bounded_owned(part.blob(), maximum, resource)?;
            let relationship_limit = if resource == "control-properties bytes" {
                limits.max_control_properties_relationship_bytes
            } else if resource == "sidecar target bytes" {
                maximum
                    .checked_sub(bytes.len())
                    .ok_or_else(|| owner_invalid("sidecar target byte budget underflow"))?
                    .min(limits.max_sidecar_relationship_bytes)
            } else {
                limits.max_sidecar_relationship_bytes
            };
            let relationships = part.rels().clone();
            let (relationship_member, relationship_payload) = eager_relationship_payload(
                package,
                uri,
                &relationships,
                relationship_limit,
                limits.max_relationship_edges,
                "sidecar relationship bytes",
            )?;
            Ok(LoadedPart {
                part_name: Some(uri.clone()),
                content_type: part.content_type().to_owned(),
                payload: RetainedPayload::owned(bytes),
                source_version: None,
                relationship_count: relationships.len(),
                relationships: fingerprint_relationships(
                    &relationships,
                    relationship_member,
                    relationship_limit,
                    resource,
                    relationship_payload.as_ref().map(RetainedPayload::as_bytes),
                )?,
                relationship_payload,
                relationship_graph: relationships,
            })
        };
    scan_owner(
        worksheet_part,
        &relationships,
        &loader,
        &limits,
        None,
        execution_context,
        unreferenced_parts.as_slice(),
        unreferenced_parts.projection_bytes(),
    )
}

/// Read-only host hook for an eager worksheet handle.
pub(crate) fn eager_form_controls_for_sheet(
    package: &OpcPackage,
    worksheet_uri: &PackURI,
) -> Result<FormControlCollection> {
    eager_form_controls_for_sheet_with_limits(package, worksheet_uri, OwnerLimits::default())
        .map_err(owner_to_xlsx)
}

pub(crate) fn eager_form_controls_for_sheet_with_limits(
    package: &OpcPackage,
    worksheet_uri: &PackURI,
    limits: OwnerLimits,
) -> OwnerResult<FormControlCollection> {
    eager_form_controls_with_execution_context(package, worksheet_uri, limits, None)
}

/// Re-scan a prepared source-backed candidate through the complete owner
/// graph.  The topology callback deliberately exposes no mutable package, so
/// this adapter turns only the bounded effective parts needed by the scanner
/// into `LoadedPart` values while retaining the source-backed lazy reads for
/// every unchanged member.
pub(crate) fn eager_form_controls_for_effective_topology(
    topology: &EffectiveTopology<'_>,
    worksheet_uri: &PackURI,
    limits: OwnerLimits,
) -> OwnerResult<FormControlCollection> {
    let worksheet = topology
        .part(worksheet_uri)
        .map_err(FormControlOwnerError::from)?;
    let relationships = worksheet.relationships().clone();
    preflight_minimum_mirror_nodes(&relationships, &limits)?;
    let worksheet_part = effective_loaded_part(
        topology,
        worksheet_uri,
        limits.max_mce_bytes,
        "worksheet bytes",
        limits.max_sidecar_relationship_bytes,
    )?;
    let unreferenced = effective_unreferenced_properties(topology, &limits)?;
    let loader = |uri: &PackURI, maximum: usize, resource: &'static str| {
        effective_loaded_part(
            topology,
            uri,
            maximum,
            resource,
            if resource == "control-properties bytes" {
                limits.max_control_properties_relationship_bytes
            } else if resource == "sidecar target bytes" {
                maximum.min(limits.max_sidecar_relationship_bytes)
            } else {
                limits.max_sidecar_relationship_bytes
            },
        )
    };
    scan_owner(
        worksheet_part,
        &relationships,
        &loader,
        &limits,
        None,
        None,
        &unreferenced.0,
        unreferenced.1,
    )
}

fn effective_loaded_part(
    topology: &EffectiveTopology<'_>,
    uri: &PackURI,
    maximum: usize,
    resource: &'static str,
    relationship_limit: usize,
) -> OwnerResult<LoadedPart> {
    let part = topology.part(uri).map_err(FormControlOwnerError::from)?;
    let data = part.data().map_err(FormControlOwnerError::from)?;
    let bytes = bounded_owned(data.as_bytes(), maximum, resource)?;
    let relationships = part.relationships().clone();
    if relationships.len() > relationship_limit.max(1) {
        return Err(owner_limit(
            "sidecar relationship edges",
            relationships.len(),
            relationship_limit.max(1),
        ));
    }
    let relationship_uri = uri
        .rels_uri()
        .map_err(|error| owner_invalid(error.to_string()))?;
    let relationship_member = topology.has_physical_member(relationship_uri.as_str())?;
    let fingerprint = fingerprint_relationships(
        &relationships,
        relationship_member,
        relationship_limit,
        resource,
        None,
    )?;
    Ok(LoadedPart {
        part_name: Some(uri.clone()),
        content_type: part.content_type().as_str().to_owned(),
        payload: RetainedPayload::owned(bytes),
        source_version: None,
        relationship_count: relationships.len(),
        relationships: fingerprint,
        relationship_payload: None,
        relationship_graph: relationships,
    })
}

fn effective_unreferenced_properties(
    topology: &EffectiveTopology<'_>,
    limits: &OwnerLimits,
) -> OwnerResult<(Vec<String>, usize)> {
    let mut names = Vec::new();
    let mut projection_bytes = 0usize;
    for part in topology.parts() {
        if part.content_type().as_str() != CONTROL_PROPERTIES_CONTENT_TYPE {
            continue;
        }
        let name = part.partname().as_str();
        if name.len() > limits.max_name_bytes {
            return Err(owner_limit(
                "unreferenced control-properties name bytes",
                name.len(),
                limits.max_name_bytes,
            ));
        }
        if names.len() >= limits.max_mirror_nodes {
            return Err(owner_limit(
                "unreferenced control-properties inventory",
                names.len() + 1,
                limits.max_mirror_nodes,
            ));
        }
        projection_bytes = projection_bytes
            .checked_add(name.len())
            .ok_or_else(|| owner_invalid("unreferenced control-properties name overflow"))?;
        let mut owned = String::new();
        owned
            .try_reserve_exact(name.len())
            .map_err(|source| owner_alloc("unreferenced control-properties name", source))?;
        owned.push_str(name);
        names.push(owned);
    }
    names.sort_unstable();
    projection_bytes = projection_bytes
        .checked_add(
            names
                .len()
                .checked_mul(size_of::<String>())
                .ok_or_else(|| {
                    owner_invalid("unreferenced control-properties capacity overflow")
                })?,
        )
        .ok_or_else(|| owner_invalid("unreferenced control-properties memory overflow"))?;
    Ok((names, projection_bytes))
}

/// Execution-capable eager worksheet hook used by hosts that retain a caller
/// budget while reading an in-memory OPC package.
#[allow(
    dead_code,
    reason = "reserved for workbook execution-context integration"
)]
pub(crate) fn eager_form_controls_for_sheet_with_limits_and_execution_context(
    package: &OpcPackage,
    worksheet_uri: &PackURI,
    limits: OwnerLimits,
    execution_context: litchi_core::ExecutionContext,
) -> OwnerResult<FormControlCollection> {
    eager_form_controls_with_execution_context(
        package,
        worksheet_uri,
        limits,
        Some(execution_context),
    )
}

/// Read-only host hook for a source-backed worksheet handle.
pub(crate) fn source_form_controls_for_sheet(
    package: &SourceBackedPackage,
    worksheet_uri: &PackURI,
) -> Result<FormControlCollection> {
    source_form_controls_for_sheet_with_limits(
        package,
        worksheet_uri,
        OwnerLimits::from_read_limits(package.read_limits()),
    )
}

pub(crate) fn source_form_controls_for_sheet_with_limits(
    package: &SourceBackedPackage,
    worksheet_uri: &PackURI,
    limits: OwnerLimits,
) -> Result<FormControlCollection> {
    package
        .check_execution()
        .map_err(FormControlOwnerError::from)
        .map_err(owner_to_xlsx)?;
    let worksheet = package
        .part(worksheet_uri)
        .map_err(FormControlOwnerError::from)
        .map_err(owner_to_xlsx)?;
    let source_version = package
        .source_version()
        .map_err(FormControlOwnerError::from)
        .map_err(owner_to_xlsx)?;
    preflight_minimum_mirror_nodes(worksheet.rels(), &limits).map_err(owner_to_xlsx)?;
    let data = source_part_data_with_limit(&worksheet, limits.max_mce_bytes, "worksheet bytes")
        .map_err(owner_to_xlsx)?;
    let relationships = worksheet.rels().clone();
    let (relationship_member, relationship_payload) = source_relationship_payload(
        package,
        worksheet_uri,
        &relationships,
        limits.max_sidecar_relationship_bytes,
        limits.max_relationship_edges,
        "worksheet relationship bytes",
    )
    .map_err(owner_to_xlsx)?;
    let worksheet_part = source_loaded_part(
        Some(worksheet_uri.clone()),
        worksheet.content_type(),
        data,
        relationships.len(),
        relationship_member,
        source_version,
        &relationships,
        limits.max_sidecar_relationship_bytes,
        "worksheet bytes",
        limits.max_mce_bytes,
        relationship_payload,
    )
    .map_err(owner_to_xlsx)?;
    let loader =
        |uri: &PackURI, maximum: usize, resource: &'static str| -> OwnerResult<LoadedPart> {
            let part = package.part(uri)?;
            let data = source_part_data_with_limit(&part, maximum, resource)?;
            let relationships = part.rels().clone();
            let relationship_limit = if resource == "control-properties bytes" {
                limits.max_control_properties_relationship_bytes
            } else if resource == "sidecar target bytes" {
                maximum
                    .checked_sub(data.as_bytes().len())
                    .ok_or_else(|| owner_invalid("sidecar target byte budget underflow"))?
                    .min(limits.max_sidecar_relationship_bytes)
            } else {
                limits.max_sidecar_relationship_bytes
            };
            let (relationship_member, relationship_payload) = source_relationship_payload(
                package,
                uri,
                &relationships,
                relationship_limit,
                limits.max_relationship_edges,
                "source relationship bytes",
            )?;
            Ok(LoadedPart {
                part_name: Some(uri.clone()),
                content_type: part.content_type().to_owned(),
                payload: RetainedPayload::source(data),
                source_version: Some(source_version),
                relationship_count: relationships.len(),
                relationships: fingerprint_relationships(
                    &relationships,
                    relationship_member,
                    relationship_limit,
                    resource,
                    relationship_payload.as_ref().map(RetainedPayload::as_bytes),
                )?,
                relationship_payload,
                relationship_graph: relationships,
            })
        };
    let collection = scan_owner(
        worksheet_part,
        &relationships,
        &loader,
        &limits,
        Some(package),
        None,
        &[],
        0,
    )
    .map_err(owner_to_xlsx)?;
    let package_readset = source_package_readset(package, &limits).map_err(owner_to_xlsx)?;
    let incoming_relationships = incoming_fingerprint_from_source(
        package,
        collection
            .read_set()
            .ok_or_else(|| invalid("form-control source read set was not retained"))?,
        package_readset.workbook.part_name.as_ref(),
        limits.max_relationship_edges,
    )
    .map_err(owner_to_xlsx)?;
    collection
        .attach_source_package_readset(package_readset, incoming_relationships, limits)
        .map_err(owner_to_xlsx)
}

pub(crate) fn owner_to_xlsx(error: FormControlOwnerError) -> Error {
    match error {
        FormControlOwnerError::Xlsx(error) => error,
        FormControlOwnerError::Package(error) => Error::Package(error),
        FormControlOwnerError::Execution(litchi_core::ExecutionError::ResourceLimit(limit)) => {
            Error::ResourceLimit(limit)
        },
        FormControlOwnerError::Execution(litchi_core::ExecutionError::Cancelled) => {
            Error::Package(litchi_opc::OpcError::Cancelled)
        },
        FormControlOwnerError::Execution(error) => {
            Error::Package(litchi_opc::OpcError::Execution(error))
        },
        FormControlOwnerError::Leaf(error) => Error::FormControl(error),
        FormControlOwnerError::Invalid(error) => Error::Invalid(error),
        FormControlOwnerError::Limit {
            resource,
            observed,
            maximum,
        } => Error::ResourceLimit(litchi_core::ResourceLimit {
            resource: owner_resource(resource),
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            limit: u64::try_from(maximum).unwrap_or(u64::MAX),
            scope: Arc::from(format!("form-control owner {resource}")),
        }),
        FormControlOwnerError::Allocation { resource, source } => {
            Error::Allocation { resource, source }
        },
    }
}

fn owner_resource(resource: &str) -> litchi_core::Resource {
    if resource.contains("projection")
        || resource.contains("semantic")
        || resource.contains("read-set")
        || resource.contains("read set")
        || resource.contains("sidecar target")
        || resource.contains("MCE range capacity")
        || resource.contains("MCE alias")
    {
        litchi_core::Resource::Memory
    } else if resource.contains("output") || resource.contains("selected provenance") {
        litchi_core::Resource::OutputBytes
    } else if resource.contains("input")
        || resource.contains("raw provenance")
        || resource.contains("ignored provenance")
        || resource.contains("bytes")
    {
        litchi_core::Resource::InputBytes
    } else if resource.contains("event") {
        litchi_core::Resource::Work
    } else if resource.contains("depth") {
        litchi_core::Resource::Depth
    } else if resource.contains("edge")
        || resource.contains("control")
        || resource.contains("shape")
        || resource.contains("node")
        || resource.contains("item")
        || resource.contains("branch")
        || resource.contains("wrapper")
        || resource.contains("inventory")
        || resource.contains("diagnostic")
    {
        litchi_core::Resource::Objects
    } else {
        litchi_core::Resource::Memory
    }
}

fn is_active_x_relationship(reltype: &str) -> bool {
    matches!(reltype, ACTIVE_X_REL | STRICT_ACTIVE_X_REL)
}

#[derive(Clone, Debug)]
struct LoadedPart {
    part_name: Option<PackURI>,
    content_type: String,
    payload: RetainedPayload,
    source_version: Option<litchi_core::SourceVersion>,
    relationship_count: usize,
    relationships: RelationshipFingerprint,
    relationship_payload: Option<RetainedPayload>,
    relationship_graph: Relationships,
}

struct SourcePackageReadSet {
    workbook: LoadedPart,
    package_relationship_payload: Option<RetainedPayload>,
    package_relationships: RelationshipFingerprint,
    content_types_payload: RetainedPayload,
    signature_present: bool,
}

impl LoadedPart {
    fn bytes(&self) -> &[u8] {
        self.payload.as_bytes()
    }

    fn into_public(self) -> FormControlPartRead {
        let range = 0..self.bytes().len();
        let relationship_source_version =
            self.relationship_payload.as_ref().and(self.source_version);
        let relationship_range = self
            .relationship_payload
            .as_ref()
            .map(|payload| 0..payload.as_bytes().len());
        FormControlPartRead {
            part_name: self.part_name,
            payload: self.payload,
            relationship_payload: self.relationship_payload,
            source_version: self.source_version,
            relationship_source_version,
            range,
            relationship_range,
            relationships: self.relationships,
        }
    }
}

fn retained_part_bytes(part: &LoadedPart) -> OwnerResult<usize> {
    part.bytes()
        .len()
        .checked_add(
            part.relationship_payload
                .as_ref()
                .map_or(0, |payload| payload.as_bytes().len()),
        )
        .ok_or_else(|| owner_invalid("retained source byte count overflow"))
}

fn eager_loaded_part(
    part_name: Option<PackURI>,
    content_type: &str,
    bytes: Vec<u8>,
    relationship_count: usize,
    relationship_member: bool,
    relationships: &Relationships,
    relationship_limit: usize,
    resource: &'static str,
    relationship_payload: Option<RetainedPayload>,
) -> OwnerResult<LoadedPart> {
    Ok(LoadedPart {
        part_name,
        content_type: content_type.to_owned(),
        payload: RetainedPayload::owned(bytes),
        source_version: None,
        relationship_count,
        relationships: fingerprint_relationships(
            relationships,
            relationship_member,
            relationship_limit,
            resource,
            relationship_payload.as_ref().map(RetainedPayload::as_bytes),
        )?,
        relationship_payload,
        relationship_graph: relationships.clone(),
    })
}

fn source_loaded_part(
    part_name: Option<PackURI>,
    content_type: &str,
    data: PartData,
    relationship_count: usize,
    relationship_member: bool,
    source_version: litchi_core::SourceVersion,
    relationships: &Relationships,
    relationship_limit: usize,
    resource: &'static str,
    maximum: usize,
    relationship_payload: Option<RetainedPayload>,
) -> OwnerResult<LoadedPart> {
    if data.as_bytes().len() > maximum {
        return Err(owner_limit(resource, data.as_bytes().len(), maximum));
    }
    Ok(LoadedPart {
        part_name,
        content_type: content_type.to_owned(),
        payload: RetainedPayload::source(data),
        source_version: Some(source_version),
        relationship_count,
        relationships: fingerprint_relationships(
            relationships,
            relationship_member,
            relationship_limit,
            resource,
            relationship_payload.as_ref().map(RetainedPayload::as_bytes),
        )?,
        relationship_payload,
        relationship_graph: relationships.clone(),
    })
}

fn source_part_data_with_limit(
    part: &litchi_opc::PartView<'_>,
    maximum: usize,
    resource: &'static str,
) -> OwnerResult<PartData> {
    let declared = part.declared_uncompressed_size()?;
    let maximum_u64 = u64::try_from(maximum).unwrap_or(u64::MAX);
    if declared > maximum_u64 {
        return Err(owner_limit(
            resource,
            usize::try_from(declared).unwrap_or(usize::MAX),
            maximum,
        ));
    }
    let data = part.data()?;
    if data.as_bytes().len() > maximum {
        return Err(owner_limit(resource, data.as_bytes().len(), maximum));
    }
    Ok(data)
}

fn source_content_types_data_with_limit(
    package: &SourceBackedPackage,
    maximum: usize,
) -> OwnerResult<PartData> {
    let caller_maximum = u64::try_from(maximum).unwrap_or(u64::MAX);
    match package.content_types_data_with_limit(maximum) {
        Ok(data) => Ok(data),
        Err(litchi_opc::OpcError::ReadLimit {
            resource: ReadResource::ContentTypesBytes,
            actual,
            maximum: observed_maximum,
        }) if observed_maximum == caller_maximum => Err(owner_limit(
            "MCE input bytes",
            usize::try_from(actual).unwrap_or(usize::MAX),
            maximum,
        )),
        Err(error) => Err(FormControlOwnerError::Package(error)),
    }
}

fn fingerprint_relationships(
    relationships: &Relationships,
    present: bool,
    maximum: usize,
    resource: &'static str,
    raw: Option<&[u8]>,
) -> OwnerResult<RelationshipFingerprint> {
    if !present {
        return Ok(RelationshipFingerprint {
            present: false,
            bytes: 0,
            digest: 0,
            edges: 0,
        });
    }
    if let Some(raw) = raw {
        if raw.len() > maximum {
            return Err(owner_limit(resource, raw.len(), maximum));
        }
        return Ok(RelationshipFingerprint {
            present: true,
            bytes: raw.len(),
            digest: fnv_bytes(0xcbf29ce484222325, raw),
            edges: relationships.len(),
        });
    }
    let mut entries = Vec::new();
    entries
        .try_reserve(relationships.len())
        .map_err(|source| owner_alloc("relationship fingerprint entries", source))?;
    entries.extend(relationships.iter());
    entries.sort_unstable_by(|left, right| left.r_id().cmp(right.r_id()));
    let mut bytes = b"<Relationships>".len();
    let mut digest = 0xcbf29ce484222325u64;
    for relationship in entries {
        let edge_bytes = relationship
            .r_id()
            .len()
            .checked_add(relationship.reltype().len())
            .and_then(|value| value.checked_add(relationship.target_ref().len()))
            .and_then(|value| value.checked_add(16))
            .ok_or_else(|| owner_invalid("relationship fingerprint byte count overflow"))?;
        bytes = bytes
            .checked_add(edge_bytes)
            .ok_or_else(|| owner_invalid("relationship fingerprint byte count overflow"))?;
        if bytes > maximum {
            return Err(owner_limit(resource, bytes, maximum));
        }
        digest = fnv_bytes(digest, relationship.r_id().as_bytes());
        digest = fnv_bytes(digest, relationship.reltype().as_bytes());
        digest = fnv_bytes(digest, relationship.target_ref().as_bytes());
        digest = fnv_bytes(
            digest,
            match relationship.target_mode() {
                TargetMode::Internal => b"I",
                TargetMode::External => b"E",
            },
        );
    }
    Ok(RelationshipFingerprint {
        present: true,
        bytes,
        digest,
        edges: relationships.len(),
    })
}

fn eager_relationship_payload(
    package: &OpcPackage,
    owner: &PackURI,
    relationships: &Relationships,
    maximum: usize,
    edge_limit: usize,
    resource: &'static str,
) -> OwnerResult<(bool, Option<RetainedPayload>)> {
    if relationships.len() > edge_limit {
        return Err(owner_limit(
            "relationship edges",
            relationships.len(),
            edge_limit,
        ));
    }
    let package_limits = package.read_limits();
    let plan = package.plan_source_relationships_with_limits(owner, package_limits)?;
    if plan.final_len() > maximum {
        return Err(owner_limit(resource, plan.final_len(), maximum));
    }
    let token = package.source_relationships_with_limits(owner, package_limits)?;
    let present = token.member_present();
    let payload = present.then(|| RetainedPayload::relationships(token));
    Ok((present, payload))
}

fn source_relationship_payload(
    package: &SourceBackedPackage,
    owner: &PackURI,
    relationships: &Relationships,
    maximum: usize,
    edge_limit: usize,
    resource: &'static str,
) -> OwnerResult<(bool, Option<RetainedPayload>)> {
    if relationships.len() > edge_limit {
        return Err(owner_limit(
            "relationship edges",
            relationships.len(),
            edge_limit,
        ));
    }
    let data = package.relationships_data_for_with_limit(owner, maximum)?;
    let Some(data) = data else {
        return Ok((false, None));
    };
    if data.as_bytes().len() > maximum {
        return Err(owner_limit(resource, data.as_bytes().len(), maximum));
    }
    Ok((true, Some(RetainedPayload::source(data))))
}

fn source_package_readset(
    package: &SourceBackedPackage,
    limits: &OwnerLimits,
) -> OwnerResult<SourcePackageReadSet> {
    let content_types_payload = RetainedPayload::source(source_content_types_data_with_limit(
        package,
        limits.max_mce_bytes,
    )?);
    let signature_present = package.has_signature_infrastructure();
    let workbook = package.main_document_part()?;
    let workbook_data =
        source_part_data_with_limit(&workbook, limits.max_mce_bytes, "workbook bytes")?;
    let workbook_relationships = workbook.rels().clone();
    let (workbook_relationship_member, workbook_relationship_payload) =
        source_relationship_payload(
            package,
            workbook.partname(),
            &workbook_relationships,
            limits.max_sidecar_relationship_bytes,
            limits.max_relationship_edges,
            "workbook relationship bytes",
        )?;
    let workbook_part = source_loaded_part(
        Some(workbook.partname().clone()),
        workbook.content_type(),
        workbook_data,
        workbook_relationships.len(),
        workbook_relationship_member,
        package.source_version()?,
        &workbook_relationships,
        limits.max_sidecar_relationship_bytes,
        "workbook bytes",
        limits.max_mce_bytes,
        workbook_relationship_payload,
    )?;

    let package_uri = PackURI::new("/").map_err(|error| owner_invalid(error.to_string()))?;
    let package_relationships = package.rels().clone();
    let (package_relationship_member, package_relationship_payload) = source_relationship_payload(
        package,
        &package_uri,
        &package_relationships,
        limits.max_sidecar_relationship_bytes,
        limits.max_relationship_edges,
        "package relationship bytes",
    )?;
    let package_relationships = fingerprint_relationships(
        &package_relationships,
        package_relationship_member,
        limits.max_sidecar_relationship_bytes,
        "package relationship bytes",
        package_relationship_payload
            .as_ref()
            .map(RetainedPayload::as_bytes),
    )?;
    Ok(SourcePackageReadSet {
        workbook: workbook_part,
        package_relationship_payload,
        package_relationships,
        content_types_payload,
        signature_present,
    })
}

/// A canonical worksheet `ctrlProp` edge necessarily contributes one
/// worksheet control, one DrawingML shape, and one VML shape to the admitted
/// closure.  Refuse an owner cap below that floor from relationship metadata
/// before the worksheet payload is decompressed or cached.  This is a
/// conservative metadata preflight for malformed/unreferenced edges; the
/// full XML census remains authoritative once the cap can admit the minimum
/// closure.
fn preflight_minimum_mirror_nodes(
    relationships: &Relationships,
    limits: &OwnerLimits,
) -> OwnerResult<()> {
    if limits.max_mirror_nodes < 3
        && relationships
            .iter()
            .any(|relationship| relationship.reltype() == CONTROL_PROPERTIES_RELATIONSHIP_TYPE)
    {
        return Err(owner_limit(
            "form-control mirror nodes",
            3,
            limits.max_mirror_nodes,
        ));
    }
    Ok(())
}

fn retain_sidecar_targets<F>(
    owner: &PackURI,
    part: &LoadedPart,
    load: &F,
    limits: &OwnerLimits,
    seen: &mut HashSet<String>,
    retained: &mut Vec<LoadedPart>,
    graph_edges: &mut usize,
    depth: usize,
    mut execution: Option<&mut OwnerExecution>,
) -> OwnerResult<()>
where
    F: Fn(&PackURI, usize, &'static str) -> OwnerResult<LoadedPart>,
{
    if part.relationship_count != part.relationship_graph.len() {
        return Err(owner_invalid(
            "sidecar relationship graph count is inconsistent",
        ));
    }
    if part.relationship_graph.len() > limits.max_relationship_edges {
        return Err(owner_limit(
            "sidecar relationship edges",
            part.relationship_graph.len(),
            limits.max_relationship_edges,
        ));
    }
    let mut new_targets = 0usize;
    for relationship in part.relationship_graph.iter() {
        if relationship.target_mode() != TargetMode::Internal {
            return Err(owner_invalid(format!(
                "sidecar relationship from {} has an external target",
                owner.as_str()
            )));
        }
        let target = relationship.target_partname()?;
        if !seen.contains(target.as_str()) {
            new_targets = new_targets
                .checked_add(1)
                .ok_or_else(|| owner_invalid("sidecar target count overflow"))?;
        }
    }
    let projected_targets = retained
        .len()
        .checked_add(depth)
        .and_then(|count| count.checked_add(new_targets))
        .ok_or_else(|| owner_invalid("sidecar target count overflow"))?;
    if projected_targets > limits.max_mirror_nodes {
        return Err(owner_limit(
            "sidecar target nodes",
            projected_targets,
            limits.max_mirror_nodes,
        ));
    }
    *graph_edges = graph_edges
        .checked_add(part.relationship_graph.len())
        .ok_or_else(|| owner_invalid("form-control relationship edge count overflow"))?;
    if *graph_edges > limits.max_relationship_edges {
        return Err(owner_limit(
            "relationship edges",
            *graph_edges,
            limits.max_relationship_edges,
        ));
    }
    if let Some(execution) = execution.as_deref_mut() {
        execution.reserve(litchi_core::Resource::Objects, new_targets)?;
        execution.reserve_capacity::<String>(
            OwnerMemoryCategory::ReadSet,
            new_targets,
            limits.max_read_set_bytes,
            "sidecar target identity capacity",
        )?;
        execution.reserve_capacity::<LoadedPart>(
            OwnerMemoryCategory::ReadSet,
            new_targets,
            limits.max_read_set_bytes,
            "sidecar target read-set capacity",
        )?;
    }
    seen.try_reserve(new_targets)
        .map_err(|source| owner_alloc("sidecar target identity index", source))?;
    retained
        .try_reserve(new_targets)
        .map_err(|source| owner_alloc("form-control sidecar target read set", source))?;
    for relationship in part.relationship_graph.iter() {
        if let Some(execution) = execution.as_deref_mut() {
            execution.check()?;
            execution.work(1)?;
        }
        if relationship.target_mode() != TargetMode::Internal {
            return Err(owner_invalid(format!(
                "sidecar relationship from {} has an external target",
                owner.as_str()
            )));
        }
        let target = relationship.target_partname()?;
        if seen.contains(target.as_str()) {
            continue;
        }
        seen.insert(target.as_str().to_owned());
        let projected_with_target = retained
            .len()
            .checked_add(depth)
            .and_then(|count| count.checked_add(1))
            .ok_or_else(|| owner_invalid("sidecar target count overflow"))?;
        if projected_with_target > limits.max_mirror_nodes {
            return Err(owner_limit(
                "sidecar target nodes",
                projected_with_target,
                limits.max_mirror_nodes,
            ));
        }
        let child_depth = depth
            .checked_add(1)
            .ok_or_else(|| owner_invalid("sidecar relationship depth overflow"))?;
        if child_depth > limits.max_mce_depth {
            return Err(owner_limit(
                "sidecar relationship depth",
                child_depth,
                limits.max_mce_depth,
            ));
        }
        // Lower the part read ceiling to the remaining aggregate sidecar
        // budget before asking the package loader to decompress the target.
        // Managed source reads therefore refuse on the declared size without
        // first retaining a payload that the owner will discard.
        let sidecar_remaining =
            execution
                .as_deref()
                .map_or(limits.max_sidecar_target_bytes, |execution| {
                    execution.remaining_memory_category(
                        OwnerMemoryCategory::SidecarTarget,
                        limits.max_sidecar_target_bytes,
                    )
                });
        if sidecar_remaining == 0 {
            return Err(owner_limit(
                "sidecar target bytes",
                limits.max_sidecar_target_bytes.saturating_add(1),
                limits.max_sidecar_target_bytes,
            ));
        }
        let target_limit = limits
            .max_mce_bytes
            .min(super::MAX_PART_BYTES)
            .min(sidecar_remaining);
        let target_part = load(&target, target_limit, "sidecar target bytes")?;
        let target_bytes = target_part
            .bytes()
            .len()
            .checked_add(
                target_part
                    .relationship_payload
                    .as_ref()
                    .map_or(0, |payload| payload.as_bytes().len()),
            )
            .ok_or_else(|| owner_invalid("sidecar target byte count overflow"))?;
        if let Some(execution) = execution.as_deref_mut() {
            execution.reserve_memory_category(
                OwnerMemoryCategory::SidecarTarget,
                target_bytes,
                limits.max_sidecar_target_bytes,
                "sidecar target bytes",
            )?;
            execution.reserve_memory_category(
                OwnerMemoryCategory::ReadSet,
                target_bytes,
                limits.max_read_set_bytes,
                "form-control read-set bytes",
            )?;
        }
        retain_sidecar_targets(
            &target,
            &target_part,
            load,
            limits,
            seen,
            retained,
            graph_edges,
            child_depth,
            execution.as_deref_mut(),
        )?;
        retained.push(target_part);
    }
    Ok(())
}

const fn fnv_bytes(mut digest: u64, bytes: &[u8]) -> u64 {
    let mut index = 0;
    while index < bytes.len() {
        digest ^= bytes[index] as u64;
        digest = digest.wrapping_mul(0x100000001b3);
        index += 1;
    }
    digest
}

fn scan_owner<F>(
    worksheet_part: LoadedPart,
    worksheet_rels: &Relationships,
    load: &F,
    limits: &OwnerLimits,
    source_package: Option<&SourceBackedPackage>,
    execution_context: Option<litchi_core::ExecutionContext>,
    unreferenced_parts: &[String],
    retained_projection_bytes: usize,
) -> OwnerResult<FormControlCollection>
where
    F: Fn(&PackURI, usize, &'static str) -> OwnerResult<LoadedPart>,
{
    let mut execution = OwnerExecution::new(source_package, execution_context);
    execution.account_memory_category(
        OwnerMemoryCategory::Projection,
        retained_projection_bytes,
        limits.max_projection_bytes,
        "form-control collection projection bytes",
    )?;
    execution.check()?;
    let expected_version = if let Some(package) = source_package {
        package.check_execution()?;
        Some(package.source_version()?)
    } else {
        None
    };
    if expected_version != worksheet_part.source_version {
        if expected_version.is_some() {
            return Err(owner_invalid(
                "worksheet source version changed before owner scan",
            ));
        }
    }
    let result = scan_owner_inner(
        worksheet_part,
        worksheet_rels,
        load,
        limits,
        source_package,
        unreferenced_parts,
        &mut execution,
    );
    let fence = if let Some(package) = source_package {
        package
            .check_execution()
            .map_err(FormControlOwnerError::from)
            .and_then(|()| {
                package
                    .source_version()
                    .map_err(FormControlOwnerError::from)
            })
            .and_then(|current_version| {
                if Some(current_version) != expected_version {
                    Err(owner_invalid(
                        "worksheet source version changed during owner scan",
                    ))
                } else {
                    Ok(())
                }
            })
    } else {
        execution.check()
    };
    match (result, fence) {
        (Err(error), _) => Err(error),
        (Ok(_), Err(error)) => Err(error),
        (Ok(collection), Ok(())) => Ok(collection),
    }
}

fn scan_owner_inner<F>(
    worksheet_part: LoadedPart,
    worksheet_rels: &Relationships,
    load: &F,
    limits: &OwnerLimits,
    source_package: Option<&SourceBackedPackage>,
    unreferenced_parts: &[String],
    execution: &mut OwnerExecution,
) -> OwnerResult<FormControlCollection>
where
    F: Fn(&PackURI, usize, &'static str) -> OwnerResult<LoadedPart>,
{
    execution.check()?;
    execution.reserve(
        litchi_core::Resource::InputBytes,
        worksheet_part.bytes().len(),
    )?;
    execution.reserve(litchi_core::Resource::Objects, worksheet_rels.len())?;
    // Reserve the parser's worst-case nesting stack once for this scan.  The
    // same execution lease covers worksheet, DrawingML, VML, and MCE passes;
    // it is released with the detached result instead of being recreated per
    // parser call.
    execution.reserve(litchi_core::Resource::Depth, limits.max_mce_depth)?;
    if source_package.is_some() {
        let worksheet_read_bytes = worksheet_part
            .bytes()
            .len()
            .checked_add(
                worksheet_part
                    .relationship_payload
                    .as_ref()
                    .map_or(0, |payload| payload.as_bytes().len()),
            )
            .ok_or_else(|| owner_invalid("worksheet read-set byte count overflow"))?;
        execution.reserve_memory_category(
            OwnerMemoryCategory::ReadSet,
            worksheet_read_bytes,
            limits.max_read_set_bytes,
            "form-control read-set bytes",
        )?;
    }
    validate_canonical_worksheet_root(worksheet_part.bytes(), Some(execution))?;
    if worksheet_rels.len() > limits.max_relationship_edges {
        return Err(owner_limit(
            "worksheet relationship edges",
            worksheet_rels.len(),
            limits.max_relationship_edges,
        ));
    }
    let selected = select_mce(
        worksheet_part.bytes(),
        FORM_CONTROL_NAMESPACE,
        limits,
        "worksheet",
        Some(&worksheet_part.payload),
        Some(execution),
    )?;
    reject_mixed_dialect(
        &selected.xml,
        "worksheet",
        &[STRICT_SML, STRICT_REL],
        Some(execution),
    )?;
    let worksheet_control_count = preflight_shape_nodes(
        &selected.xml,
        SML,
        b"control",
        limits.max_controls,
        limits,
        "effective controls",
        Some(execution),
    )?;
    // Resolve the worksheet relationship type during the allocation-free
    // census.  ActiveX and opaque worksheet controls do not consume this
    // form-control mirror budget, so counting every `<control>` would make a
    // small cap reject an otherwise empty form-control collection.
    let preflight_form_control_count =
        preflight_form_control_nodes(&selected.xml, worksheet_rels, limits, Some(execution))?;
    let minimum_mirror_nodes = preflight_form_control_count
        .checked_mul(3)
        .ok_or_else(|| owner_invalid("form-control mirror-node count overflow"))?;
    if minimum_mirror_nodes > limits.max_mirror_nodes {
        return Err(owner_limit(
            "form-control mirror nodes",
            minimum_mirror_nodes,
            limits.max_mirror_nodes,
        ));
    }
    execution.reserve_capacity::<SourceControl>(
        OwnerMemoryCategory::Projection,
        worksheet_control_count,
        limits.max_projection_bytes,
        "worksheet control projection capacity",
    )?;
    let worksheet = parse_worksheet(&selected.xml, limits, Some(execution))?;
    execution.reconcile_capacity::<SourceControl>(
        OwnerMemoryCategory::Projection,
        worksheet_control_count,
        worksheet.controls.capacity(),
        limits.max_projection_bytes,
        "worksheet control projection capacity",
    )?;
    let worksheet_owner_objects = worksheet
        .controls
        .len()
        .checked_mul(3)
        .ok_or_else(|| owner_invalid("worksheet owner object count overflow"))?;
    execution.reserve(litchi_core::Resource::Objects, worksheet_owner_objects)?;
    let mut active_x_shapes = HashSet::new();
    active_x_shapes
        .try_reserve(worksheet.controls.len())
        .map_err(|source| owner_alloc("ActiveX control identity index", source))?;
    let mut form_shapes = HashSet::new();
    form_shapes
        .try_reserve(worksheet.controls.len())
        .map_err(|source| owner_alloc("form-control identity index", source))?;
    let mut form_control_count = 0usize;
    let mut has_unclassified_control = false;
    for source_control in &worksheet.controls {
        let relationship = worksheet_rels
            .get(&source_control.relationship_id)
            .ok_or_else(|| owner_invalid("form-control worksheet relationship is missing"))?;
        if is_active_x_relationship(relationship.reltype()) {
            if !active_x_shapes.insert(source_control.shape_id) {
                return Err(owner_invalid(
                    "multiple effective ActiveX controls claim one shape identity",
                ));
            }
        } else if relationship.reltype() == CONTROL_PROPERTIES_RELATIONSHIP_TYPE {
            if !form_shapes.insert(source_control.shape_id) {
                return Err(owner_invalid(
                    "multiple effective form controls claim one shape identity",
                ));
            }
            form_control_count = form_control_count
                .checked_add(1)
                .ok_or_else(|| owner_invalid("form-control inventory count overflow"))?;
        } else {
            has_unclassified_control = true;
        }
    }
    if form_shapes
        .iter()
        .any(|shape_id| active_x_shapes.contains(shape_id))
    {
        return Err(owner_invalid(
            "one effective worksheet control has both form-control and ActiveX persistence owners",
        ));
    }
    if form_control_count == 0 && !has_unclassified_control {
        return empty_collection(
            source_package,
            &HashSet::new(),
            worksheet_part,
            selected.provenance,
            unreferenced_parts,
            limits,
            execution,
        );
    }
    // The parsed classification must agree with the allocation-free census;
    // a mismatch would mean the relationship graph changed while scanning.
    if form_control_count != preflight_form_control_count {
        return Err(owner_invalid(
            "worksheet form-control census changed during owner scan",
        ));
    }
    let drawing_uri = relationship_target(
        worksheet_rels,
        worksheet.drawing_relationship_id.as_deref(),
        rt::DRAWING,
        "worksheet drawing",
    )?;
    let vml_uri = relationship_target(
        worksheet_rels,
        worksheet.vml_relationship_id.as_deref(),
        rt::VML_DRAWING,
        "worksheet legacyDrawing",
    )?;
    let mut retained_sidecar_targets = Vec::new();
    execution.reserve_capacity::<LoadedPart>(
        OwnerMemoryCategory::ReadSet,
        form_control_count.min(limits.max_mirror_nodes),
        limits.max_read_set_bytes,
        "form-control sidecar target read-set capacity",
    )?;
    retained_sidecar_targets
        .try_reserve(form_control_count.min(limits.max_mirror_nodes))
        .map_err(|source| owner_alloc("form-control sidecar target read set", source))?;
    let mut retained_sidecar_target_uris = HashSet::<String>::new();
    execution.reserve_memory_category(
        OwnerMemoryCategory::ReadSet,
        form_control_count
            .min(limits.max_mirror_nodes)
            .checked_mul(size_of::<String>())
            .ok_or_else(|| owner_invalid("sidecar target index capacity overflow"))?,
        limits.max_read_set_bytes,
        "form-control sidecar target index capacity",
    )?;
    retained_sidecar_target_uris
        .try_reserve(form_control_count.min(limits.max_mirror_nodes))
        .map_err(|source| owner_alloc("form-control sidecar target index", source))?;
    let drawing_part = load(&drawing_uri, limits.max_drawing_bytes, "DrawingML bytes")?;
    // Worksheet edges are already loaded before the sidecar walk.  The walk
    // itself charges each DrawingML/VML/properties sidecar edge exactly once,
    // including recursively retained sidecar targets.
    let mut graph_edges = worksheet_rels.len();
    if drawing_part.content_type != ct::OFC_DRAWING {
        return Err(owner_invalid(
            "drawing relationship target has an unexpected content type",
        ));
    }
    if source_package.is_some() {
        let drawing_read_bytes = retained_part_bytes(&drawing_part)?;
        execution.reserve_memory_category(
            OwnerMemoryCategory::ReadSet,
            drawing_read_bytes,
            limits.max_read_set_bytes,
            "form-control read-set bytes",
        )?;
    }
    retain_sidecar_targets(
        &drawing_uri,
        &drawing_part,
        load,
        limits,
        &mut retained_sidecar_target_uris,
        &mut retained_sidecar_targets,
        &mut graph_edges,
        0,
        Some(execution),
    )?;
    let drawing_selected = select_mce(
        drawing_part.bytes(),
        "http://schemas.microsoft.com/office/drawing/2010/main",
        limits,
        "drawing",
        Some(&drawing_part.payload),
        Some(execution),
    )?;
    reject_mixed_dialect(
        &drawing_selected.xml,
        "drawing",
        &[STRICT_XDR, STRICT_REL],
        Some(execution),
    )?;
    validate_canonical_drawing_root(&drawing_selected.xml, Some(execution))?;
    let drawing_shape_count = preflight_shape_nodes(
        &drawing_selected.xml,
        XDR,
        b"sp",
        limits.max_shapes,
        limits,
        "DrawingML shape identities",
        Some(execution),
    )?;
    let projected_drawing_nodes = form_control_count
        .checked_add(drawing_shape_count)
        .and_then(|count| count.checked_add(retained_sidecar_targets.len()))
        .ok_or_else(|| owner_invalid("form-control mirror-node count overflow"))?;
    if projected_drawing_nodes > limits.max_mirror_nodes {
        return Err(owner_limit(
            "form-control mirror nodes",
            projected_drawing_nodes,
            limits.max_mirror_nodes,
        ));
    }
    execution.reserve_capacity::<DrawingShape>(
        OwnerMemoryCategory::Projection,
        drawing_shape_count,
        limits.max_projection_bytes,
        "DrawingML shape projection capacity",
    )?;
    validate_drawing_anchors(&drawing_selected.xml, limits, Some(execution))?;
    let drawing_shapes = parse_drawing(&drawing_selected.xml, limits, Some(execution))?;
    execution.reconcile_capacity::<DrawingShape>(
        OwnerMemoryCategory::Projection,
        drawing_shape_count,
        drawing_shapes.capacity(),
        limits.max_projection_bytes,
        "DrawingML shape projection capacity",
    )?;
    let vml_part = load(&vml_uri, limits.max_vml_bytes, "VML bytes")?;
    if vml_part.content_type != ct::OFC_VML_DRAWING {
        return Err(owner_invalid(
            "VML relationship target has an unexpected content type",
        ));
    }
    validate_namespace_bindings(vml_part.bytes(), limits, Some(execution))?;
    if source_package.is_some() {
        let vml_read_bytes = retained_part_bytes(&vml_part)?;
        execution.reserve_memory_category(
            OwnerMemoryCategory::ReadSet,
            vml_read_bytes,
            limits.max_read_set_bytes,
            "form-control read-set bytes",
        )?;
    }
    retain_sidecar_targets(
        &vml_uri,
        &vml_part,
        load,
        limits,
        &mut retained_sidecar_target_uris,
        &mut retained_sidecar_targets,
        &mut graph_edges,
        0,
        Some(execution),
    )?;
    let vml_shape_count = preflight_shape_nodes(
        vml_part.bytes(),
        VML,
        b"shape",
        limits.max_shapes,
        limits,
        "VML shape identities",
        Some(execution),
    )?;
    let projected_mirror_nodes = form_control_count
        .checked_add(drawing_shape_count)
        .and_then(|count| count.checked_add(vml_shape_count))
        .and_then(|count| count.checked_add(retained_sidecar_targets.len()))
        .ok_or_else(|| owner_invalid("form-control mirror-node count overflow"))?;
    if projected_mirror_nodes > limits.max_mirror_nodes {
        return Err(owner_limit(
            "form-control mirror nodes",
            projected_mirror_nodes,
            limits.max_mirror_nodes,
        ));
    }
    execution.reserve_capacity::<VmlShape>(
        OwnerMemoryCategory::Projection,
        vml_shape_count,
        limits.max_projection_bytes,
        "VML shape projection capacity",
    )?;
    execution.reserve(litchi_core::Resource::InputBytes, vml_part.bytes().len())?;
    let vml_shapes = parse_vml(vml_part.bytes(), limits, Some(execution))?;
    execution.reconcile_capacity::<VmlShape>(
        OwnerMemoryCategory::Projection,
        vml_shape_count,
        vml_shapes.capacity(),
        limits.max_projection_bytes,
        "VML shape projection capacity",
    )?;
    let mirror_nodes = form_control_count
        .checked_add(drawing_shapes.len())
        .and_then(|count| count.checked_add(vml_shapes.len()))
        .and_then(|count| count.checked_add(retained_sidecar_targets.len()))
        .ok_or_else(|| owner_invalid("form-control mirror-node count overflow"))?;
    if mirror_nodes > limits.max_mirror_nodes {
        return Err(owner_limit(
            "form-control mirror nodes",
            mirror_nodes,
            limits.max_mirror_nodes,
        ));
    }
    let mut owner_parts = HashMap::<String, usize>::new();
    let index_objects = form_control_count
        .checked_mul(4)
        .ok_or_else(|| owner_invalid("form-control owner object count overflow"))?;
    execution.reserve(litchi_core::Resource::Objects, index_objects)?;
    execution.reserve_memory_category(
        OwnerMemoryCategory::Projection,
        form_control_count
            .checked_mul(size_of::<FormControlView>())
            .ok_or_else(|| owner_invalid("form-control projection capacity overflow"))?,
        limits.max_projection_bytes,
        "form-control projection capacity",
    )?;
    execution.reserve_memory_category(
        OwnerMemoryCategory::ReadSet,
        form_control_count
            .checked_mul(size_of::<FormControlPartRead>())
            .ok_or_else(|| owner_invalid("form-control read-set capacity overflow"))?,
        limits.max_read_set_bytes,
        "form-control property read-set capacity",
    )?;
    owner_parts
        .try_reserve(form_control_count)
        .map_err(|source| owner_alloc("control owner index", source))?;
    let mut controls = Vec::new();
    controls
        .try_reserve_exact(form_control_count)
        .map_err(|source| owner_alloc("form-control inventory", source))?;
    let mut retained_properties = Vec::new();
    retained_properties
        .try_reserve_exact(form_control_count)
        .map_err(|source| owner_alloc("form-control source read set", source))?;
    let mut retained_semantic_holds = Vec::new();
    if source_package.is_some() {
        retained_semantic_holds
            .try_reserve_exact(form_control_count)
            .map_err(|source| owner_alloc("form-control semantic lease read set", source))?;
    }
    let mut incoming_targets = HashSet::<String>::new();
    incoming_targets
        .try_reserve(form_control_count)
        .map_err(|source| owner_alloc("form-control incoming index", source))?;
    let drawing_index = index_drawing_shapes(&drawing_shapes, Some(execution))?;
    let vml_index = index_vml_shapes(&vml_shapes, Some(execution))?;
    for source_control in &worksheet.controls {
        let relationship = worksheet_rels
            .get(&source_control.relationship_id)
            .ok_or_else(|| owner_invalid("form-control worksheet relationship is missing"))?;
        if is_active_x_relationship(relationship.reltype()) {
            // ActiveX has a separate inert owner.  Its worksheet edge and
            // shape remain part of the opaque source graph, but it must not
            // force a form-control read to inspect or load the descriptor.
            continue;
        }
        let position = controls.len();
        if position >= limits.max_controls {
            return Err(owner_limit(
                "effective controls",
                position + 1,
                limits.max_controls,
            ));
        }
        if relationship.target_mode() != TargetMode::Internal {
            return Err(owner_invalid(
                "form-control relationship target is external",
            ));
        }
        if relationship.reltype() != CONTROL_PROPERTIES_RELATIONSHIP_TYPE {
            return Err(owner_invalid(
                "form-control relationship type is not canonical ctrlProp",
            ));
        }
        let target = relationship.target_partname()?;
        if !incoming_targets.insert(target.as_str().to_owned()) {
            return Err(owner_invalid(
                "form-control properties target has multiple effective owners",
            ));
        }
        let part = load(
            &target,
            limits.max_mce_bytes.min(super::MAX_PART_BYTES),
            "control-properties bytes",
        )?;
        let properties_bytes = part.bytes().len();
        execution.reserve_memory_category(
            OwnerMemoryCategory::Semantic,
            properties_bytes,
            limits.max_semantic_bytes,
            "form-control semantic bytes",
        )?;
        if source_package.is_some() {
            execution.reserve_memory_category(
                OwnerMemoryCategory::ReadSet,
                retained_part_bytes(&part)?,
                limits.max_read_set_bytes,
                "form-control read-set bytes",
            )?;
        }
        execution.reserve(litchi_core::Resource::InputBytes, part.bytes().len())?;
        execution.reserve(litchi_core::Resource::Objects, 1)?;
        if part.content_type != CONTROL_PROPERTIES_CONTENT_TYPE {
            return Err(owner_invalid(
                "form-control properties content type is not canonical",
            ));
        }
        validate_namespace_bindings(part.bytes(), limits, Some(execution))?;
        retain_sidecar_targets(
            &target,
            &part,
            load,
            limits,
            &mut retained_sidecar_target_uris,
            &mut retained_sidecar_targets,
            &mut graph_edges,
            0,
            Some(execution),
        )?;
        let leaf = leaf_limits(limits);
        let properties = super::codec::parse_source_with_limits_and_context(
            part.payload.source_payload(),
            &leaf,
            execution.context(),
        )?;
        // The leaf parser retains its own execution reservation so cloned
        // `Properties` values keep their semantic allocations alive.  Fold
        // that already-charged amount into the owner's aggregate semantic
        // ceiling without charging the same core budget twice.
        if let Some(lease) = properties.retained_lease.as_ref() {
            let lease_bytes = usize::try_from(lease.amount()).unwrap_or(usize::MAX);
            execution.account_memory_category(
                OwnerMemoryCategory::Semantic,
                lease_bytes,
                limits.max_semantic_bytes,
                "form-control leaf semantic lease",
            )?;
        }
        retained_properties.push(part.clone().into_public());
        if source_package.is_some() {
            if let Some(lease) = properties.retained_lease.clone() {
                retained_semantic_holds.push(lease);
            }
        }
        let drawing = source_control.shape_id;
        let shape = match drawing_index.get(&drawing) {
            Some(Some(index)) => drawing_shapes
                .get(*index)
                .ok_or_else(|| owner_invalid("DrawingML identity index is inconsistent"))?,
            Some(None) => return Err(owner_invalid("DrawingML identity is ambiguous")),
            None => {
                return Err(owner_invalid(
                    "DrawingML cNvPr identity does not match control shapeId",
                ));
            },
        };
        if shape.compat_spid.is_empty() {
            return Err(owner_invalid("DrawingML cNvPr is missing a14:compatExt"));
        }
        if numeric_spid(&shape.compat_spid) != Some(shape.id) {
            return Err(owner_invalid(
                "DrawingML a14:compatExt/@spid does not match cNvPr/@id",
            ));
        }
        if let (Some(control_name), Some(drawing_name)) =
            (source_control.name.as_deref(), shape.name.as_deref())
        {
            if control_name != drawing_name {
                return Err(owner_invalid(
                    "DrawingML cNvPr name does not match worksheet control name",
                ));
            }
        }
        let object_type = properties
            .object_type()
            .map(|value| match value {
                super::KnownOrUnknown::Known(value) => value.to_string(),
                super::KnownOrUnknown::Unknown(value) => value.clone(),
            })
            .unwrap_or_default();
        let expected_vml_type = match object_type.as_str() {
            "CheckBox" => Some("Checkbox"),
            "Button" => Some("Button"),
            "Radio" => Some("Radio"),
            _ => None,
        };
        let vml = match_vml(
            &vml_shapes,
            &vml_index,
            shape,
            expected_vml_type,
            source_control.name.as_deref(),
            limits,
        )?;
        let mut projection_bytes = size_of::<ShapeClosure>();
        for value in [
            source_control.name.as_deref(),
            source_control.anchor_profile.as_deref(),
            shape.name.as_deref(),
            Some(shape.compat_spid.as_str()),
            Some(vml.id.as_str()),
            vml.spid.as_deref(),
            Some(object_type.as_str()),
            Some(vml.object_type.as_str()),
        ]
        .into_iter()
        .flatten()
        {
            projection_bytes = projection_bytes
                .checked_add(value.len())
                .ok_or_else(|| owner_invalid("form-control projection bytes overflow"))?;
        }
        execution.reserve_memory_category(
            OwnerMemoryCategory::Projection,
            projection_bytes,
            limits.max_projection_bytes,
            "form-control projection bytes",
        )?;
        let mut diagnostics = DiagnosticBuffer::new(execution, limits);
        if let Some(detail) = source_control.anchor_diagnostic.as_deref() {
            diagnostics.push_fmt(
                FormControlDiagnosticCode::ControlPrAnchor,
                format_args!("{detail}"),
            )?;
        }
        if object_type.is_empty()
            || properties
                .object_type()
                .is_some_and(|value| value.unknown().is_some())
        {
            diagnostics.push_fmt(
                FormControlDiagnosticCode::OpaqueGraph,
                format_args!(
                    "form-control objectType is absent or unknown; the VML mirror remains opaque"
                ),
            )?;
        }
        if expected_vml_type.is_none()
            || vml.object_type.is_empty()
            || expected_vml_type != Some(vml.object_type.as_str())
        {
            diagnostics.push_fmt(
                FormControlDiagnosticCode::UnprovenMirror,
                format_args!(
                    "VML ClientData does not prove the typed form-control objectType mirror"
                ),
            )?;
        }
        if properties.source_bytes().is_none() {
            diagnostics.push_fmt(
                FormControlDiagnosticCode::OpaqueGraph,
                format_args!("properties source bytes were not retained"),
            )?;
        }
        vml_mirror_diagnostics(&properties, vml, limits, &mut diagnostics)?;
        let diagnostics = diagnostics.into_vec();
        let closure = ShapeClosure {
            shape_id: drawing,
            drawing_name: shape.name.clone().map(Into::into),
            compat_spid: shape.compat_spid.clone().into(),
            vml_id: vml.id.clone().into(),
            vml_spid: vml.spid.clone().map(Into::into),
            object_type: object_type.clone().into(),
            vml_object_type: vml.object_type.clone().into(),
            drawing_relationships_present: drawing_part.relationships.present(),
            vml_relationships_present: vml_part.relationships.present(),
            drawing_relationships: drawing_part.relationships,
            vml_relationships: vml_part.relationships,
        };
        if owner_parts
            .insert(target.as_str().to_owned(), position)
            .is_some()
        {
            return Err(owner_invalid(
                "form-control properties part is shared by multiple owners",
            ));
        }
        controls.push(FormControlView {
            position,
            property_part: Some(target.clone()),
            vml_part: Some(vml_uri.clone()),
            name: source_control.name.clone().map(Into::into),
            anchor_profile: source_control.anchor_profile.clone().map(Into::into),
            properties: Arc::new(properties),
            properties_relationships_present: part.relationships.present(),
            properties_relationships: part.relationships,
            shape: Arc::new(closure),
            diagnostics: Arc::from(diagnostics.into_boxed_slice()),
            read_set: None,
            budget_hold: None,
        });
    }
    execution.reconcile_capacity::<LoadedPart>(
        OwnerMemoryCategory::ReadSet,
        form_control_count.min(limits.max_mirror_nodes),
        retained_sidecar_targets.capacity(),
        limits.max_read_set_bytes,
        "form-control sidecar target read-set capacity",
    )?;
    execution.reconcile_capacity::<FormControlView>(
        OwnerMemoryCategory::Projection,
        form_control_count,
        controls.capacity(),
        limits.max_projection_bytes,
        "form-control projection capacity",
    )?;
    execution.reconcile_capacity::<FormControlPartRead>(
        OwnerMemoryCategory::ReadSet,
        form_control_count,
        retained_properties.capacity(),
        limits.max_read_set_bytes,
        "form-control property read-set capacity",
    )?;
    let final_mirror_nodes = form_control_count
        .checked_add(drawing_shapes.len())
        .and_then(|count| count.checked_add(vml_shapes.len()))
        .and_then(|count| count.checked_add(retained_sidecar_targets.len()))
        .ok_or_else(|| owner_invalid("form-control mirror-node count overflow"))?;
    if final_mirror_nodes > limits.max_mirror_nodes {
        return Err(owner_limit(
            "form-control mirror nodes",
            final_mirror_nodes,
            limits.max_mirror_nodes,
        ));
    }
    let diagnostics = collection_diagnostics(
        source_package,
        &incoming_targets,
        unreferenced_parts,
        limits,
        Some(execution),
    )?;
    let retained_structure_bytes = if source_package.is_some() {
        let sidecar_bytes = retained_sidecar_targets
            .len()
            .checked_mul(size_of::<FormControlPartRead>())
            .ok_or_else(|| owner_invalid("sidecar read-set capacity overflow"))?;
        let semantic_lease_bytes = retained_semantic_holds
            .len()
            .checked_mul(size_of::<Arc<litchi_core::Reservation>>())
            .ok_or_else(|| owner_invalid("semantic lease read-set capacity overflow"))?;
        size_of::<FormControlReadSet>()
            .checked_add(sidecar_bytes)
            .and_then(|bytes| bytes.checked_add(semantic_lease_bytes))
            .ok_or_else(|| owner_invalid("form-control read-set capacity overflow"))?
    } else {
        0
    };
    execution.reserve_memory_category(
        OwnerMemoryCategory::ReadSet,
        retained_structure_bytes,
        limits.max_read_set_bytes,
        "form-control read-set capacity",
    )?;
    let budget_hold = execution.finish();
    let read_set = if source_package.is_some() {
        let mut worksheet_mce = selected.provenance;
        worksheet_mce.budget_hold = budget_hold.clone();
        let mut drawing_mce = drawing_selected.provenance;
        drawing_mce.budget_hold = budget_hold.clone();
        let mut sidecar_targets = Vec::new();
        sidecar_targets
            .try_reserve_exact(retained_sidecar_targets.len())
            .map_err(|source| owner_alloc("form-control sidecar target read set", source))?;
        for part in retained_sidecar_targets {
            sidecar_targets.push(part.into_public());
        }
        Some(FormControlReadSet {
            source_version: worksheet_part.source_version,
            workbook: None,
            package_relationship_payload: None,
            package_relationships: RelationshipFingerprint {
                present: false,
                bytes: 0,
                digest: 0,
                edges: 0,
            },
            content_types_payload: None,
            signature_present: false,
            incoming_relationships: RelationshipFingerprint {
                present: false,
                bytes: 0,
                digest: 0,
                edges: 0,
            },
            worksheet: worksheet_part.into_public(),
            drawing: Some(drawing_part.into_public()),
            vml: Some(vml_part.into_public()),
            properties: Arc::from(retained_properties.into_boxed_slice()),
            sidecar_targets: Arc::from(sidecar_targets.into_boxed_slice()),
            semantic_holds: Arc::from(retained_semantic_holds.into_boxed_slice()),
            worksheet_mce,
            drawing_mce,
            budget_hold: budget_hold.clone(),
        })
    } else {
        None
    };
    let read_set = read_set.map(Arc::new);
    for control in &mut controls {
        control.read_set = read_set.clone();
        control.budget_hold = budget_hold.clone();
    }
    Ok(FormControlCollection {
        profile: OwnerProfile::canonical(),
        controls: Arc::from(controls.into_boxed_slice()),
        diagnostics: Arc::from(diagnostics.into_boxed_slice()),
        read_set,
        budget_hold,
        generated_budget: None,
    })
}

fn empty_collection(
    source_package: Option<&SourceBackedPackage>,
    referenced: &HashSet<String>,
    worksheet_part: LoadedPart,
    worksheet_mce: MceProvenance,
    unreferenced_parts: &[String],
    limits: &OwnerLimits,
    execution: &mut OwnerExecution,
) -> OwnerResult<FormControlCollection> {
    let diagnostics = collection_diagnostics(
        source_package,
        referenced,
        unreferenced_parts,
        limits,
        Some(execution),
    )?;
    if source_package.is_some() {
        execution.reserve_memory_category(
            OwnerMemoryCategory::ReadSet,
            size_of::<FormControlReadSet>(),
            limits.max_read_set_bytes,
            "form-control read-set capacity",
        )?;
    }
    let mut read_set = source_package.is_some().then(|| FormControlReadSet {
        source_version: worksheet_part.source_version,
        workbook: None,
        package_relationship_payload: None,
        package_relationships: RelationshipFingerprint {
            present: false,
            bytes: 0,
            digest: 0,
            edges: 0,
        },
        content_types_payload: None,
        signature_present: false,
        incoming_relationships: RelationshipFingerprint {
            present: false,
            bytes: 0,
            digest: 0,
            edges: 0,
        },
        worksheet: worksheet_part.into_public(),
        drawing: None,
        vml: None,
        properties: Arc::from([]),
        sidecar_targets: Arc::from([]),
        semantic_holds: Arc::from([]),
        worksheet_mce,
        drawing_mce: empty_mce_provenance(),
        budget_hold: None,
    });
    let budget_hold = execution.finish();
    if let Some(read_set) = read_set.as_mut() {
        read_set.worksheet_mce.budget_hold = budget_hold.clone();
        read_set.budget_hold = budget_hold.clone();
    }
    let read_set = read_set.map(Arc::new);
    Ok(FormControlCollection {
        profile: OwnerProfile::canonical(),
        controls: Arc::from(Vec::<FormControlView>::new().into_boxed_slice()),
        diagnostics: Arc::from(diagnostics.into_boxed_slice()),
        read_set,
        budget_hold,
        generated_budget: None,
    })
}

fn collection_diagnostics(
    source_package: Option<&SourceBackedPackage>,
    referenced: &HashSet<String>,
    eager_parts: &[String],
    limits: &OwnerLimits,
    mut execution: Option<&mut OwnerExecution>,
) -> OwnerResult<Vec<FormControlDiagnostic>> {
    // Count and charge the deterministic opaque inventory before allocating
    // the name and diagnostic vectors.  The count is independent of source
    // ordering, so the later sort cannot turn an unbounded package inventory
    // into an uncharged collection.
    let mut candidate_count = 0usize;
    let mut candidate_name_bytes = 0usize;
    let mut candidate_detail_bytes = 0usize;
    if let Some(package) = source_package {
        for part in package.iter_parts() {
            if let Some(execution) = execution.as_deref_mut() {
                execution.check()?;
                execution.work(1)?;
            }
            if part.content_type() == CONTROL_PROPERTIES_CONTENT_TYPE
                && !referenced.contains(part.partname().as_str())
            {
                if part.partname().as_str().len() > limits.max_name_bytes {
                    return Err(owner_limit(
                        "form-control diagnostic name bytes",
                        part.partname().as_str().len(),
                        limits.max_name_bytes,
                    ));
                }
                candidate_count = candidate_count
                    .checked_add(1)
                    .ok_or_else(|| owner_invalid("form-control diagnostic count overflow"))?;
                candidate_name_bytes = candidate_name_bytes
                    .checked_add(part.partname().as_str().len())
                    .ok_or_else(|| owner_invalid("form-control diagnostic name bytes overflow"))?;
                candidate_detail_bytes = candidate_detail_bytes
                    .checked_add(
                        UNREFERENCED_DIAGNOSTIC_PREFIX
                            .len()
                            .checked_add(part.partname().as_str().len())
                            .ok_or_else(|| {
                                owner_invalid("form-control diagnostic detail bytes overflow")
                            })?,
                    )
                    .ok_or_else(|| {
                        owner_invalid("form-control diagnostic detail bytes overflow")
                    })?;
            }
        }
    }
    for partname in eager_parts {
        if let Some(execution) = execution.as_deref_mut() {
            execution.check()?;
            execution.work(1)?;
        }
        if !referenced.contains(partname) {
            if partname.len() > limits.max_name_bytes {
                return Err(owner_limit(
                    "form-control diagnostic name bytes",
                    partname.len(),
                    limits.max_name_bytes,
                ));
            }
            candidate_count = candidate_count
                .checked_add(1)
                .ok_or_else(|| owner_invalid("form-control diagnostic count overflow"))?;
            candidate_name_bytes = candidate_name_bytes
                .checked_add(partname.len())
                .ok_or_else(|| owner_invalid("form-control diagnostic name bytes overflow"))?;
            candidate_detail_bytes = candidate_detail_bytes
                .checked_add(
                    UNREFERENCED_DIAGNOSTIC_PREFIX
                        .len()
                        .checked_add(partname.len())
                        .ok_or_else(|| {
                            owner_invalid("form-control diagnostic detail bytes overflow")
                        })?,
                )
                .ok_or_else(|| owner_invalid("form-control diagnostic detail bytes overflow"))?;
        }
    }
    if candidate_count > limits.max_mirror_nodes {
        return Err(owner_limit(
            "form-control collection diagnostics",
            candidate_count,
            limits.max_mirror_nodes,
        ));
    }
    if let Some(execution) = execution.as_deref_mut() {
        let diagnostic_objects = candidate_count
            .checked_mul(2)
            .ok_or_else(|| owner_invalid("form-control diagnostic object count overflow"))?;
        execution.reserve(litchi_core::Resource::Objects, diagnostic_objects)?;
        let mut diagnostic_memory = candidate_name_bytes
            .checked_add(
                candidate_count
                    .checked_mul(size_of::<String>())
                    .ok_or_else(|| owner_invalid("form-control diagnostic capacity overflow"))?,
            )
            .ok_or_else(|| owner_invalid("form-control diagnostic memory overflow"))?;
        diagnostic_memory = diagnostic_memory
            .checked_add(
                candidate_count
                    .checked_mul(size_of::<FormControlDiagnostic>())
                    .ok_or_else(|| owner_invalid("form-control diagnostic capacity overflow"))?,
            )
            .ok_or_else(|| owner_invalid("form-control diagnostic memory overflow"))?;
        diagnostic_memory = diagnostic_memory
            .checked_add(candidate_detail_bytes)
            .ok_or_else(|| owner_invalid("form-control diagnostic memory overflow"))?;
        execution.reserve_memory_category(
            OwnerMemoryCategory::Projection,
            diagnostic_memory,
            limits.max_projection_bytes,
            "form-control collection projection bytes",
        )?;
    }
    let mut names = Vec::new();
    names
        .try_reserve_exact(candidate_count)
        .map_err(|source| owner_alloc("form-control unreferenced inventory", source))?;
    if let Some(execution) = execution.as_deref_mut() {
        execution.reconcile_capacity::<String>(
            OwnerMemoryCategory::Projection,
            candidate_count,
            names.capacity(),
            limits.max_projection_bytes,
            "form-control collection projection bytes",
        )?;
    }
    if let Some(package) = source_package {
        for part in package.iter_parts() {
            if let Some(execution) = execution.as_deref_mut() {
                execution.check()?;
                execution.work(1)?;
            }
            if part.content_type() == CONTROL_PROPERTIES_CONTENT_TYPE
                && !referenced.contains(part.partname().as_str())
            {
                if names.len() >= limits.max_mirror_nodes {
                    return Err(owner_limit(
                        "form-control collection diagnostics",
                        names.len() + 1,
                        limits.max_mirror_nodes,
                    ));
                }
                if names.len() >= names.capacity() {
                    return Err(owner_invalid(
                        "form-control unreferenced inventory census changed",
                    ));
                }
                let name =
                    collection_owned_text("", part.partname().as_str(), &mut execution, limits)?;
                names.push(name);
            }
        }
    }
    for partname in eager_parts {
        if let Some(execution) = execution.as_deref_mut() {
            execution.check()?;
            execution.work(1)?;
        }
        if !referenced.contains(partname) {
            if names.len() >= limits.max_mirror_nodes {
                return Err(owner_limit(
                    "form-control collection diagnostics",
                    names.len() + 1,
                    limits.max_mirror_nodes,
                ));
            }
            if names.len() >= names.capacity() {
                return Err(owner_invalid(
                    "form-control unreferenced inventory census changed",
                ));
            }
            let name = collection_owned_text("", partname, &mut execution, limits)?;
            names.push(name);
        }
    }
    names.sort_unstable();
    let mut diagnostics = Vec::new();
    diagnostics
        .try_reserve_exact(names.len().min(limits.max_mirror_nodes))
        .map_err(|source| owner_alloc("form-control collection diagnostics", source))?;
    if let Some(execution) = execution.as_deref_mut() {
        execution.reconcile_capacity::<FormControlDiagnostic>(
            OwnerMemoryCategory::Projection,
            names.len().min(limits.max_mirror_nodes),
            diagnostics.capacity(),
            limits.max_projection_bytes,
            "form-control collection projection bytes",
        )?;
    }
    for partname in names {
        if let Some(execution) = execution.as_deref_mut() {
            execution.check()?;
            execution.work(1)?;
        }
        if diagnostics.len() >= diagnostics.capacity() {
            return Err(owner_invalid(
                "form-control collection diagnostic census changed",
            ));
        }
        let detail = collection_owned_text(
            UNREFERENCED_DIAGNOSTIC_PREFIX,
            &partname,
            &mut execution,
            limits,
        )?;
        diagnostics.push(FormControlDiagnostic::new(
            FormControlDiagnosticCode::UnreferencedControlProperties,
            detail.into_boxed_str(),
        ));
    }
    Ok(diagnostics)
}

fn collection_owned_text(
    prefix: &str,
    suffix: &str,
    execution: &mut Option<&mut OwnerExecution>,
    limits: &OwnerLimits,
) -> OwnerResult<String> {
    let length = prefix
        .len()
        .checked_add(suffix.len())
        .ok_or_else(|| owner_invalid("form-control collection text length overflow"))?;
    let mut value = String::new();
    value
        .try_reserve_exact(length)
        .map_err(|source| owner_alloc("form-control collection text", source))?;
    if let Some(execution) = execution.as_deref_mut() {
        execution.reconcile_capacity::<u8>(
            OwnerMemoryCategory::Projection,
            length,
            value.capacity(),
            limits.max_projection_bytes,
            "form-control collection projection bytes",
        )?;
    }
    value.push_str(prefix);
    value.push_str(suffix);
    Ok(value)
}

fn eager_unreferenced_properties(
    package: &OpcPackage,
    limits: &OwnerLimits,
    execution: Option<&litchi_core::ExecutionContext>,
) -> OwnerResult<EagerUnreferencedInventory> {
    let mut candidate_count = 0usize;
    let mut candidate_name_bytes = 0usize;
    for part in package.iter_parts() {
        if let Some(execution) = execution {
            execution.check().map_err(FormControlOwnerError::from)?;
            execution
                .consume(litchi_core::Resource::Work, 1)
                .map_err(FormControlOwnerError::from)?;
        }
        if part.content_type() == CONTROL_PROPERTIES_CONTENT_TYPE {
            if part.partname().as_str().len() > limits.max_name_bytes {
                return Err(owner_limit(
                    "unreferenced control-properties name bytes",
                    part.partname().as_str().len(),
                    limits.max_name_bytes,
                ));
            }
            candidate_count = candidate_count
                .checked_add(1)
                .ok_or_else(|| owner_invalid("unreferenced control-properties count overflow"))?;
            if candidate_count > limits.max_mirror_nodes {
                return Err(owner_limit(
                    "unreferenced control-properties inventory",
                    candidate_count,
                    limits.max_mirror_nodes,
                ));
            }
            candidate_name_bytes = candidate_name_bytes
                .checked_add(part.partname().as_str().len())
                .ok_or_else(|| owner_invalid("unreferenced control-properties name overflow"))?;
        }
    }
    let inventory_memory = candidate_count
        .checked_mul(size_of::<String>())
        .and_then(|bytes| bytes.checked_add(candidate_name_bytes))
        .ok_or_else(|| owner_invalid("unreferenced control-properties memory overflow"))?;
    let mut charged_memory = 0usize;
    let mut memory = None;
    reserve_eager_inventory_memory(
        execution,
        &mut memory,
        &mut charged_memory,
        inventory_memory,
        limits.max_projection_bytes,
    )?;
    let mut objects = None;
    if let Some(execution) = execution {
        let reservation = execution
            .reserve(
                litchi_core::Resource::Objects,
                u64::try_from(candidate_count).unwrap_or(u64::MAX),
            )
            .map_err(FormControlOwnerError::from)?;
        objects = Some(reservation);
    }
    let mut parts = Vec::new();
    parts
        .try_reserve_exact(candidate_count)
        .map_err(|source| owner_alloc("unreferenced control-properties inventory", source))?;
    if parts.capacity() > candidate_count {
        let extra = parts
            .capacity()
            .checked_sub(candidate_count)
            .and_then(|count| count.checked_mul(size_of::<String>()))
            .ok_or_else(|| owner_invalid("unreferenced control-properties capacity overflow"))?;
        reserve_eager_inventory_memory(
            execution,
            &mut memory,
            &mut charged_memory,
            extra,
            limits.max_projection_bytes,
        )?;
    }
    for part in package.iter_parts() {
        if let Some(execution) = execution {
            execution.check().map_err(FormControlOwnerError::from)?;
            execution
                .consume(litchi_core::Resource::Work, 1)
                .map_err(FormControlOwnerError::from)?;
        }
        if part.content_type() == CONTROL_PROPERTIES_CONTENT_TYPE {
            if part.partname().as_str().len() > limits.max_name_bytes {
                return Err(owner_limit(
                    "unreferenced control-properties name bytes",
                    part.partname().as_str().len(),
                    limits.max_name_bytes,
                ));
            }
            let value = part.partname().as_str();
            let mut owned = String::new();
            owned
                .try_reserve_exact(value.len())
                .map_err(|source| owner_alloc("unreferenced control-properties name", source))?;
            if owned.capacity() > value.len() {
                reserve_eager_inventory_memory(
                    execution,
                    &mut memory,
                    &mut charged_memory,
                    owned.capacity() - value.len(),
                    limits.max_projection_bytes,
                )?;
            }
            owned.push_str(value);
            if parts.len() >= parts.capacity() {
                return Err(owner_invalid(
                    "unreferenced control-properties census changed",
                ));
            }
            parts.push(owned);
        }
    }
    parts.sort_unstable();
    Ok(EagerUnreferencedInventory {
        parts,
        _memory: memory,
        _objects: objects,
        projection_bytes: charged_memory,
    })
}

fn merge_owner_reservation(
    target: &mut Option<litchi_core::Reservation>,
    reservation: litchi_core::Reservation,
) -> OwnerResult<()> {
    if let Some(existing) = target.as_mut() {
        if existing.try_merge(reservation).is_err() {
            return Err(owner_invalid(
                "eager inventory reservations do not share a budget",
            ));
        }
    } else {
        *target = Some(reservation);
    }
    Ok(())
}

fn reserve_eager_inventory_memory(
    execution: Option<&litchi_core::ExecutionContext>,
    reservation: &mut Option<litchi_core::Reservation>,
    charged: &mut usize,
    amount: usize,
    maximum: usize,
) -> OwnerResult<()> {
    let observed = charged
        .checked_add(amount)
        .ok_or_else(|| owner_invalid("unreferenced control-properties memory overflow"))?;
    if observed > maximum {
        return Err(owner_limit(
            "form-control collection projection bytes",
            observed,
            maximum,
        ));
    }
    if let Some(execution) = execution {
        let next = execution
            .reserve(
                litchi_core::Resource::Memory,
                u64::try_from(amount).unwrap_or(u64::MAX),
            )
            .map_err(FormControlOwnerError::from)?;
        merge_owner_reservation(reservation, next)?;
    }
    *charged = observed;
    Ok(())
}

fn resolve_sheet_position(sheets: &[raw::Sheet], selector: Selector<'_>) -> OwnerResult<usize> {
    match selector {
        litchi_core::Selector::Position(position) => Ok(position.get()),
        litchi_core::Selector::Name(name) => sheets
            .iter()
            .position(|sheet| sheet.name.eq_ignore_ascii_case(name.as_ref()))
            .ok_or_else(|| owner_invalid("worksheet selector did not resolve")),
        litchi_core::Selector::Id(_) => {
            Err(owner_invalid("worksheet ID selectors are not admitted"))
        },
        _ => Err(owner_invalid("worksheet selector is unsupported")),
    }
}

fn validate_canonical_worksheet_root(
    xml: &[u8],
    execution: Option<&OwnerExecution>,
) -> OwnerResult<()> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    loop {
        if let Some(execution) = execution {
            execution.check()?;
            execution.work(1)?;
        }
        let event = reader
            .read_event()
            .map_err(|error| owner_invalid(error.to_string()))?;
        match event {
            Event::Start(root) | Event::Empty(root) => {
                let (namespace, local) = reader.resolver().resolve_element(root.name());
                if !is_ns(&namespace, SML) || local.as_ref() != b"worksheet" {
                    if is_ns(&namespace, STRICT_SML) {
                        return Err(FormControlOwnerError::Invalid(
                            "strict worksheet form-control dialect is not admitted".into(),
                        ));
                    }
                    return Err(owner_invalid(
                        "worksheet root is not canonical SpreadsheetML",
                    ));
                }
                return Ok(());
            },
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => return Err(owner_invalid("worksheet XML has no root element")),
            _ => return Err(owner_invalid("worksheet XML has no root element")),
        }
    }
}

/// Validate namespace declaration bindings after XML attribute-value
/// normalization.  quick-xml rejects the common raw spellings of reserved
/// bindings, but its resolver intentionally keeps entity references lexical;
/// this pass closes the escaped-URI hole before any owner or MCE projection is
/// admitted.
fn validate_namespace_bindings(
    xml: &[u8],
    limits: &OwnerLimits,
    execution: Option<&OwnerExecution>,
) -> OwnerResult<()> {
    validate_namespace_bindings_and_count_events::<false>(xml, limits, execution).map(|_| ())
}

/// Validate namespace declarations while counting the XML events consumed by
/// the MCE selector.  The selector historically admitted the same source
/// through a namespace pass and a second event-count pass.  Keep the two
/// execution-work charges for that caller so its work/cancellation limits and
/// failure boundaries remain unchanged, but read each XML event only once.
fn validate_namespace_bindings_and_count_events<const PRESERVE_DUPLICATE_WORK_CHARGE: bool>(
    xml: &[u8],
    limits: &OwnerLimits,
    execution: Option<&OwnerExecution>,
) -> OwnerResult<usize> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut events = 0usize;
    loop {
        if let Some(execution) = execution {
            execution.check()?;
            execution.work(1)?;
        }
        events = events
            .checked_add(1)
            .ok_or_else(|| owner_invalid("namespace binding event count overflow"))?;
        if events > limits.max_mce_events {
            return Err(owner_limit(
                "namespace binding events",
                events,
                limits.max_mce_events,
            ));
        }
        match reader
            .read_event()
            .map_err(|error| owner_invalid(error.to_string()))?
        {
            Event::Start(element) | Event::Empty(element) => {
                validate_namespace_declarations(&element)?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if PRESERVE_DUPLICATE_WORK_CHARGE {
        if let Some(execution) = execution {
            for _ in 0..events {
                execution.check()?;
                execution.work(1)?;
            }
        }
    }
    Ok(events)
}

fn validate_namespace_declarations(element: &BytesStart<'_>) -> OwnerResult<()> {
    for attribute in element.unchecked_attributes() {
        let attribute = attribute.map_err(|error| owner_invalid(error.to_string()))?;
        let Some(binding) = attribute.key.as_namespace_binding() else {
            continue;
        };
        let value = attribute.value.as_ref();
        let binds_xml = namespace_uri_matches(value, XML);
        let binds_xmlns = namespace_uri_matches(value, XMLNS);
        match binding {
            quick_xml::name::PrefixDeclaration::Default if binds_xml || binds_xmlns => {
                return Err(owner_invalid(
                    "default namespace cannot bind a reserved XML namespace",
                ));
            },
            quick_xml::name::PrefixDeclaration::Named(prefix) if prefix == b"xml" && !binds_xml => {
                return Err(owner_invalid(
                    "xml prefix is bound to a namespace other than the XML namespace",
                ));
            },
            quick_xml::name::PrefixDeclaration::Named(prefix)
                if prefix == b"xmlns" || (prefix != b"xml" && binds_xml) || binds_xmlns =>
            {
                return Err(owner_invalid(
                    "namespace binding uses a reserved XML prefix/URI",
                ));
            },
            _ => {},
        }
    }
    Ok(())
}

fn select_mce(
    xml: &[u8],
    capability: &str,
    limits: &OwnerLimits,
    subject: &str,
    source: Option<&RetainedPayload>,
    mut execution: Option<&mut OwnerExecution>,
) -> OwnerResult<MceSelection> {
    let mce_events =
        validate_namespace_bindings_and_count_events::<true>(xml, limits, execution.as_deref())?;
    if let Some(execution) = execution.as_deref_mut() {
        execution.reserve(litchi_core::Resource::Objects, mce_events)?;
        execution.work(1)?;
    }
    let mut capabilities = Capabilities::default();
    capabilities.understand_namespace(capability.to_owned());
    let mce_limits = MceLimits {
        max_input_bytes: limits.max_mce_bytes,
        max_output_bytes: limits.max_mce_bytes,
        max_depth: limits.max_mce_depth,
        max_choices_per_alternate: limits.max_mce_branches,
        ..MceLimits::default()
    };
    match process_markup_compatibility(xml, &capabilities, &mce_limits) {
        Ok(output) => Ok(mce_selection(
            xml,
            output.xml,
            output.report,
            capability,
            limits,
            source,
            execution,
        )?),
        Err(first_error) => {
            // Some native producers rename the capability prefix on the
            // selected branch but leave an inherited `mc:Ignorable` value
            // using the old spelling.  The expanded namespace is still
            // unambiguous, so bind that stale spelling only for MCE
            // selection; the owner never publishes this normalized buffer.
            if let Some(execution) = execution.as_deref_mut() {
                execution.check()?;
                execution.work(1)?;
                // The repair is a bounded replacement over the source XML;
                // reserve the source-sized peak before `replace` allocates.
                execution.reserve_memory_category(
                    OwnerMemoryCategory::ReadSet,
                    xml.len(),
                    limits.max_read_set_bytes,
                    "form-control MCE alias scratch",
                )?;
            }
            let Some(alias_input) = repair_renamed_capability_prefix(xml, capability) else {
                return Err(mce_to_owner_error(subject, first_error, limits));
            };
            if let Some(execution) = execution.as_deref_mut() {
                execution.check()?;
                execution.work(alias_input.len().max(1))?;
            }
            process_markup_compatibility(&alias_input, &capabilities, &mce_limits)
                .map_err(|error| mce_to_owner_error(subject, error, limits))
                .and_then(move |output| {
                    mce_selection(
                        xml,
                        output.xml,
                        output.report,
                        capability,
                        limits,
                        source,
                        execution,
                    )
                })
        },
    }
}

fn mce_to_owner_error(
    subject: &str,
    error: litchi_ooxml_common::mce::Error,
    limits: &OwnerLimits,
) -> FormControlOwnerError {
    match error {
        litchi_ooxml_common::mce::Error::LimitExceeded(resource) => {
            let (resource_name, maximum) = if resource.contains("depth") {
                ("MCE depth", limits.max_mce_depth)
            } else if resource.contains("choice") {
                ("MCE branches", limits.max_mce_branches)
            } else if resource.contains("input") {
                ("MCE input bytes", limits.max_mce_bytes)
            } else if resource.contains("output") {
                ("MCE output bytes", limits.max_mce_bytes)
            } else {
                ("MCE branches", limits.max_mce_branches)
            };
            owner_limit(resource_name, maximum.saturating_add(1), maximum)
        },
        litchi_ooxml_common::mce::Error::Allocation { resource, source } => {
            owner_alloc(resource, source)
        },
        error => FormControlOwnerError::Invalid(format!("{subject} MCE selection failed: {error}")),
    }
}

struct MceSelection {
    xml: Arc<[u8]>,
    provenance: MceProvenance,
}

fn mce_selection(
    raw: &[u8],
    selected: Cow<'_, [u8]>,
    report: litchi_ooxml_common::mce::Report,
    capability: &str,
    limits: &OwnerLimits,
    source: Option<&RetainedPayload>,
    mut execution: Option<&mut OwnerExecution>,
) -> OwnerResult<MceSelection> {
    if let Some(execution) = execution.as_deref_mut() {
        let range_capacity = limits
            .max_mce_branches
            .checked_mul(5)
            .and_then(|count| count.checked_mul(size_of::<Range<usize>>()))
            .ok_or_else(|| owner_invalid("MCE provenance range capacity overflow"))?;
        execution.reserve_memory_category(
            OwnerMemoryCategory::ReadSet,
            range_capacity,
            limits.max_read_set_bytes,
            "form-control MCE range capacity",
        )?;
    }
    let ranges = mce_branch_ranges(raw, capability, limits, execution.as_deref())?;
    let ignored_len = ranges
        .ignored
        .iter()
        .try_fold(0usize, |total, range| {
            total
                .checked_add(range.end.checked_sub(range.start).ok_or(())?)
                .ok_or(())
        })
        .map_err(|()| owner_invalid("MCE ignored provenance byte count overflow"))?;
    let retained_len = raw
        .len()
        .checked_add(selected.len())
        .and_then(|length| length.checked_add(ignored_len))
        .ok_or_else(|| owner_invalid("MCE retained provenance byte count overflow"))?;
    if retained_len > limits.max_mce_bytes {
        return Err(owner_limit(
            "MCE retained provenance",
            retained_len,
            limits.max_mce_bytes,
        ));
    }
    if let Some(execution) = execution.as_deref_mut() {
        execution.reserve(litchi_core::Resource::InputBytes, raw.len())?;
        execution.reserve(litchi_core::Resource::InputBytes, ignored_len)?;
        execution.reserve(litchi_core::Resource::OutputBytes, selected.len())?;
        let range_objects = ranges
            .selected
            .len()
            .checked_add(ranges.ignored.len())
            .and_then(|count| count.checked_add(ranges.wrappers.len()))
            .and_then(|count| count.checked_add(ranges.selected_parents.len()))
            .and_then(|count| count.checked_add(ranges.ignored_parents.len()))
            .ok_or_else(|| owner_invalid("MCE provenance object count overflow"))?;
        execution.reserve(litchi_core::Resource::Objects, range_objects)?;
        let range_bytes = range_objects
            .checked_mul(size_of::<Range<usize>>())
            .ok_or_else(|| owner_invalid("MCE provenance range bytes overflow"))?;
        let provenance_bytes = retained_len
            .checked_add(range_bytes)
            .ok_or_else(|| owner_invalid("MCE provenance read-set bytes overflow"))?;
        execution.reserve_memory_category(
            OwnerMemoryCategory::ReadSet,
            provenance_bytes,
            limits.max_read_set_bytes,
            "form-control MCE read-set bytes",
        )?;
    }
    let mut ignored = Vec::new();
    ignored
        .try_reserve_exact(ignored_len)
        .map_err(|source| owner_alloc("MCE ignored provenance", source))?;
    for range in &ranges.ignored {
        ignored.extend_from_slice(
            raw.get(range.clone())
                .ok_or_else(|| owner_invalid("MCE ignored provenance range is invalid"))?,
        );
    }
    if let Some(execution) = execution.as_mut() {
        execution.reconcile_capacity::<u8>(
            OwnerMemoryCategory::ReadSet,
            ignored_len,
            ignored.capacity(),
            limits.max_read_set_bytes,
            "form-control MCE ignored provenance capacity",
        )?;
    }
    let selected = match selected {
        Cow::Borrowed(bytes) => Arc::<[u8]>::from(bounded_owned(
            bytes,
            limits.max_mce_bytes,
            "MCE selected provenance",
        )?),
        Cow::Owned(bytes) => Arc::<[u8]>::from(bytes),
    };
    let raw = if let Some(source) = source {
        source.clone()
    } else {
        RetainedPayload::owned(bounded_owned(
            raw,
            limits.max_mce_bytes,
            "MCE raw provenance",
        )?)
    };
    Ok(MceSelection {
        xml: Arc::clone(&selected),
        provenance: MceProvenance {
            raw,
            selected,
            ignored: Arc::from(ignored),
            selected_ranges: Arc::from(ranges.selected.into_boxed_slice()),
            ignored_ranges: Arc::from(ranges.ignored.into_boxed_slice()),
            wrapper_ranges: Arc::from(ranges.wrappers.into_boxed_slice()),
            selected_parent_ranges: Arc::from(ranges.selected_parents.into_boxed_slice()),
            ignored_parent_ranges: Arc::from(ranges.ignored_parents.into_boxed_slice()),
            selected_choices: report.selected_choices,
            selected_fallbacks: report.selected_fallbacks,
            ignored_elements: report.ignored_elements,
            ignored_attributes: report.ignored_attributes,
            budget_hold: None,
        },
    })
}

fn empty_mce_provenance() -> MceProvenance {
    MceProvenance {
        raw: RetainedPayload::Owned(Arc::new(Vec::new())),
        selected: Arc::from([]),
        ignored: Arc::from([]),
        selected_ranges: Arc::from([]),
        ignored_ranges: Arc::from([]),
        wrapper_ranges: Arc::from([]),
        selected_parent_ranges: Arc::from([]),
        ignored_parent_ranges: Arc::from([]),
        selected_choices: 0,
        selected_fallbacks: 0,
        ignored_elements: 0,
        ignored_attributes: 0,
        budget_hold: None,
    }
}

struct MceBranchRanges {
    selected: Vec<Range<usize>>,
    ignored: Vec<Range<usize>>,
    wrappers: Vec<Range<usize>>,
    selected_parents: Vec<Range<usize>>,
    ignored_parents: Vec<Range<usize>>,
}

#[derive(Clone, Copy)]
enum MceBranchKind {
    Choice,
    Fallback,
}

struct MceBranchRange {
    start: usize,
    end: usize,
    kind: MceBranchKind,
    supported: bool,
}

struct MceAlternateFrame {
    start: usize,
    branches: Vec<MceBranchRange>,
}

struct MceElementFrame {
    alternate: Option<usize>,
    branch: Option<(usize, usize)>,
    ignored_ancestor: bool,
    namespace: Vec<u8>,
    local: Vec<u8>,
}

fn mce_branch_ranges(
    xml: &[u8],
    capability: &str,
    limits: &OwnerLimits,
    execution: Option<&OwnerExecution>,
) -> OwnerResult<MceBranchRanges> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    // Provenance ranges are byte ranges of `xml`, whose leading byte-order
    // mark precedes reader position zero.
    let origin = ReaderOrigin::of(xml);
    let mut elements = Vec::<MceElementFrame>::new();
    let mut alternates = Vec::<Option<MceAlternateFrame>>::new();
    let mut selected = Vec::new();
    let mut ignored = Vec::new();
    let mut wrappers = Vec::new();
    let mut selected_parents = Vec::new();
    let mut ignored_parents = Vec::new();
    let mut events = 0usize;
    loop {
        if let Some(execution) = execution {
            execution.check()?;
            execution.work(1)?;
        }
        events = events
            .checked_add(1)
            .ok_or_else(|| owner_invalid("MCE provenance event count overflow"))?;
        if events > limits.max_mce_events {
            return Err(owner_limit(
                "MCE provenance events",
                events,
                limits.max_mce_events,
            ));
        }
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| owner_invalid("MCE source position overflow"))?;
        let event = reader
            .read_event()
            .map_err(|error| owner_invalid(error.to_string()))?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| owner_invalid("MCE source position overflow"))?;
        let resolver = reader.resolver().clone();
        match event {
            Event::Start(element) => {
                let (namespace, local) = resolver.resolve_element(element.name());
                let namespace = namespace_bytes(namespace);
                let local = local.as_ref().to_vec();
                let alternate =
                    if namespace.as_slice() == MCE && local.as_slice() == b"AlternateContent" {
                        if alternates.len() >= limits.max_mce_branches {
                            return Err(owner_limit(
                                "MCE provenance wrappers",
                                alternates.len() + 1,
                                limits.max_mce_branches,
                            ));
                        }
                        alternates
                            .try_reserve(1)
                            .map_err(|source| owner_alloc("MCE provenance branches", source))?;
                        alternates.push(Some(MceAlternateFrame {
                            start,
                            branches: Vec::new(),
                        }));
                        Some(alternates.len() - 1)
                    } else {
                        None
                    };
                let parent_alternate = elements.last().and_then(|frame| frame.alternate);
                let inherited_ignored = elements.last().is_some_and(|frame| frame.ignored_ancestor);
                let mut branch_ignored = inherited_ignored;
                let branch = if let Some(parent_alternate) = parent_alternate {
                    if namespace.as_slice() == MCE
                        && matches!(local.as_slice(), b"Choice" | b"Fallback")
                    {
                        let kind = if local.as_slice() == b"Choice" {
                            MceBranchKind::Choice
                        } else {
                            MceBranchKind::Fallback
                        };
                        let supported = match kind {
                            MceBranchKind::Fallback => true,
                            MceBranchKind::Choice => mce_choice_supported(
                                &element,
                                &resolver,
                                reader.decoder(),
                                capability,
                                mce_baseline_namespaces(capability),
                                limits,
                            )?,
                        };
                        branch_ignored = inherited_ignored
                            || alternates
                                .get(parent_alternate)
                                .and_then(Option::as_ref)
                                .is_some_and(|state| match kind {
                                    MceBranchKind::Choice => {
                                        !supported
                                            || state.branches.iter().any(|branch| {
                                                matches!(branch.kind, MceBranchKind::Choice)
                                                    && branch.supported
                                            })
                                    },
                                    MceBranchKind::Fallback => {
                                        state.branches.iter().any(|branch| {
                                            matches!(branch.kind, MceBranchKind::Choice)
                                                && branch.supported
                                        })
                                    },
                                });
                        let parent = alternates
                            .get_mut(parent_alternate)
                            .and_then(Option::as_mut)
                            .ok_or_else(|| {
                                owner_invalid("MCE provenance alternate state is invalid")
                            })?;
                        if parent.branches.len() >= limits.max_mce_branches {
                            return Err(owner_limit(
                                "MCE provenance branches",
                                parent.branches.len() + 1,
                                limits.max_mce_branches,
                            ));
                        }
                        parent
                            .branches
                            .try_reserve(1)
                            .map_err(|source| owner_alloc("MCE provenance branches", source))?;
                        parent.branches.push(MceBranchRange {
                            start,
                            end,
                            kind,
                            supported,
                        });
                        Some((parent_alternate, parent.branches.len() - 1))
                    } else {
                        None
                    }
                } else {
                    None
                };
                if elements.len() >= limits.max_mce_depth {
                    return Err(owner_limit(
                        "MCE provenance depth",
                        elements.len() + 1,
                        limits.max_mce_depth,
                    ));
                }
                elements
                    .try_reserve(1)
                    .map_err(|source| owner_alloc("MCE provenance element stack", source))?;
                elements.push(MceElementFrame {
                    alternate,
                    branch,
                    ignored_ancestor: branch_ignored,
                    namespace,
                    local,
                });
            },
            Event::Empty(element) => {
                let (namespace, local) = resolver.resolve_element(element.name());
                let namespace = namespace_bytes(namespace);
                let local = local.as_ref().to_vec();
                if namespace.as_slice() == MCE && local.as_slice() == b"AlternateContent" {
                    return Err(owner_invalid("MCE AlternateContent cannot be empty"));
                }
                if let Some(parent) = elements.last().and_then(|frame| {
                    frame
                        .alternate
                        .and_then(|index| alternates.get_mut(index))
                        .and_then(Option::as_mut)
                }) {
                    if namespace.as_slice() == MCE
                        && matches!(local.as_slice(), b"Choice" | b"Fallback")
                    {
                        let kind = if local.as_slice() == b"Choice" {
                            MceBranchKind::Choice
                        } else {
                            MceBranchKind::Fallback
                        };
                        let supported = match kind {
                            MceBranchKind::Fallback => true,
                            MceBranchKind::Choice => mce_choice_supported(
                                &element,
                                &resolver,
                                reader.decoder(),
                                capability,
                                mce_baseline_namespaces(capability),
                                limits,
                            )?,
                        };
                        if parent.branches.len() >= limits.max_mce_branches {
                            return Err(owner_limit(
                                "MCE provenance branches",
                                parent.branches.len() + 1,
                                limits.max_mce_branches,
                            ));
                        }
                        parent
                            .branches
                            .try_reserve(1)
                            .map_err(|source| owner_alloc("MCE provenance branches", source))?;
                        parent.branches.push(MceBranchRange {
                            start,
                            end,
                            kind,
                            supported,
                        });
                    }
                }
            },
            Event::End(end_element) => {
                let (namespace, local) = resolver.resolve_element(end_element.name());
                let actual = (namespace_bytes(namespace), local.as_ref().to_vec());
                let frame = elements
                    .pop()
                    .ok_or_else(|| owner_invalid("MCE provenance closing tag is unmatched"))?;
                if frame.namespace != actual.0 || frame.local != actual.1 {
                    return Err(owner_invalid(
                        "MCE provenance closing QName does not match its opening QName",
                    ));
                }
                if let Some((alternate, branch)) = frame.branch {
                    if let Some(Some(state)) = alternates.get_mut(alternate) {
                        let item = state.branches.get_mut(branch).ok_or_else(|| {
                            owner_invalid("MCE provenance branch index is invalid")
                        })?;
                        item.end = end;
                    }
                }
                if let Some(alternate) = frame.alternate {
                    let Some(Some(state)) = alternates.get_mut(alternate) else {
                        return Err(owner_invalid("MCE provenance alternate state is invalid"));
                    };
                    let selected_index = state
                        .branches
                        .iter()
                        .position(|branch| {
                            matches!(branch.kind, MceBranchKind::Choice) && branch.supported
                        })
                        .or_else(|| {
                            state
                                .branches
                                .iter()
                                .position(|branch| matches!(branch.kind, MceBranchKind::Fallback))
                        });
                    let selected_index = selected_index.ok_or_else(|| {
                        owner_invalid("MCE provenance has no selected Choice or Fallback")
                    })?;
                    let wrapper_range = state.start..end;
                    wrappers
                        .try_reserve(1)
                        .map_err(|source| owner_alloc("MCE wrapper ranges", source))?;
                    wrappers.push(wrapper_range.clone());
                    if !frame.ignored_ancestor
                        && (state.branches[selected_index].supported
                            || matches!(
                                state.branches[selected_index].kind,
                                MceBranchKind::Fallback
                            ))
                    {
                        selected_parents
                            .try_reserve(1)
                            .map_err(|source| owner_alloc("MCE selected parent ranges", source))?;
                        selected_parents.push(wrapper_range.clone());
                    }
                    let wrapper_ignored = frame.ignored_ancestor;
                    let mut has_ignored_branch = wrapper_ignored;
                    for (index, branch) in state.branches.iter().enumerate() {
                        let range = branch.start..branch.end;
                        if !wrapper_ignored && index == selected_index {
                            selected
                                .try_reserve(1)
                                .map_err(|source| owner_alloc("MCE selected ranges", source))?;
                            selected.push(range);
                        } else {
                            has_ignored_branch = true;
                            ignored
                                .try_reserve(1)
                                .map_err(|source| owner_alloc("MCE ignored ranges", source))?;
                            ignored.push(range);
                        }
                    }
                    if has_ignored_branch {
                        ignored_parents
                            .try_reserve(1)
                            .map_err(|source| owner_alloc("MCE ignored parent ranges", source))?;
                        ignored_parents.push(wrapper_range);
                    }
                    alternates[alternate] = None;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !elements.is_empty() {
        return Err(owner_invalid("MCE provenance XML is unterminated"));
    }
    coalesce_nested_ranges(&mut selected);
    coalesce_nested_ranges(&mut ignored);
    Ok(MceBranchRanges {
        selected,
        ignored,
        wrappers,
        selected_parents,
        ignored_parents,
    })
}

fn mce_choice_supported(
    element: &BytesStart<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    capability: &str,
    baseline: &[&[u8]],
    limits: &OwnerLimits,
) -> OwnerResult<bool> {
    let mut requires = None;
    for attribute in element.unchecked_attributes() {
        let attribute = attribute.map_err(|error| owner_invalid(error.to_string()))?;
        if attribute.key.as_ref() == b"Requires" {
            requires = Some(decode_attribute_bounded(
                attribute.value.as_ref(),
                decoder,
                limits,
                "MCE Requires value",
            )?);
            break;
        }
    }
    let requires = requires.ok_or_else(|| owner_invalid("MCE Choice requires Requires"))?;
    let mut tokens = requires.split_whitespace();
    let mut supported = true;
    let mut count = 0usize;
    for prefix in tokens.by_ref() {
        count = count
            .checked_add(1)
            .ok_or_else(|| owner_invalid("MCE Requires token count overflow"))?;
        let qualified = format!("{prefix}:x");
        let (namespace, _) = resolver.resolve_element(quick_xml::name::QName(qualified.as_bytes()));
        let namespace = match namespace {
            ResolveResult::Bound(Namespace(value)) => value,
            _ => return Err(owner_invalid("MCE Choice Requires prefix is unbound")),
        };
        supported &= namespace_uri_matches(namespace, capability.as_bytes())
            || baseline
                .iter()
                .any(|wanted| namespace_uri_matches(namespace, wanted));
    }
    if count == 0 {
        return Err(owner_invalid("MCE Choice Requires is empty"));
    }
    Ok(supported)
}

fn mce_baseline_namespaces(capability: &str) -> &'static [&'static [u8]] {
    const SML_BASELINE: &[&[u8]] = &[
        SML,
        STRICT_SML,
        REL,
        STRICT_REL,
        XDR,
        STRICT_XDR,
        DRAWING,
        STRICT_DRAWING,
        MATH,
        STRICT_MATH,
        VML,
        OFFICE,
        XML,
        MCE,
    ];
    let _ = capability;
    SML_BASELINE
}

fn coalesce_nested_ranges(ranges: &mut Vec<Range<usize>>) {
    ranges.sort_unstable_by_key(|range| (range.start, std::cmp::Reverse(range.end)));
    let mut write = 0usize;
    for read in 0..ranges.len() {
        if write != 0 && ranges[write - 1].end >= ranges[read].end {
            continue;
        }
        if write != read {
            ranges.swap(write, read);
        }
        write += 1;
    }
    ranges.truncate(write);
}

fn validate_canonical_drawing_root(
    xml: &[u8],
    execution: Option<&OwnerExecution>,
) -> OwnerResult<()> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    loop {
        if let Some(execution) = execution {
            execution.check()?;
            execution.work(1)?;
        }
        let event = reader
            .read_event()
            .map_err(|error| owner_invalid(error.to_string()))?;
        match event {
            Event::Start(root) | Event::Empty(root) => {
                let (namespace, local) = reader.resolver().resolve_element(root.name());
                if is_ns(&namespace, STRICT_XDR) {
                    return Err(FormControlOwnerError::Invalid(
                        "strict DrawingML form-control dialect is not admitted".into(),
                    ));
                }
                if !is_ns(&namespace, XDR) || local.as_ref() != b"wsDr" {
                    return Err(owner_invalid(
                        "drawing root is not canonical SpreadsheetDrawing",
                    ));
                }
                return Ok(());
            },
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) => {},
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            Event::Eof => return Err(owner_invalid("drawing XML has no root element")),
            _ => return Err(owner_invalid("drawing XML has no root element")),
        }
    }
}

fn reject_mixed_dialect(
    xml: &[u8],
    subject: &'static str,
    strict_namespaces: &[&[u8]],
    execution: Option<&OwnerExecution>,
) -> OwnerResult<()> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    loop {
        if let Some(execution) = execution {
            execution.check()?;
            execution.work(1)?;
        }
        let event = reader
            .read_event()
            .map_err(|error| owner_invalid(error.to_string()))?;
        let resolver = reader.resolver().clone();
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let (namespace, _) = resolver.resolve_element(element.name());
                let mut strict_attribute = false;
                for attribute in element.unchecked_attributes() {
                    let attribute = attribute.map_err(|error| owner_invalid(error.to_string()))?;
                    let (namespace, _) = resolver.resolve_attribute(attribute.key);
                    if is_ns(&namespace, STRICT_REL) {
                        strict_attribute = true;
                        break;
                    }
                }
                if strict_namespaces
                    .iter()
                    .any(|wanted| is_ns(&namespace, wanted))
                    || strict_attribute
                {
                    return Err(FormControlOwnerError::Invalid(format!(
                        "{subject} contains a strict/mixed dialect branch"
                    )));
                }
            },
            Event::Eof => return Ok(()),
            _ => {},
        }
    }
}

fn validate_drawing_anchors(
    xml: &[u8],
    limits: &OwnerLimits,
    execution: Option<&OwnerExecution>,
) -> OwnerResult<()> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut stack: Vec<(Vec<u8>, Vec<u8>)> = Vec::new();
    let mut anchors = Vec::<(Vec<u8>, usize, usize, usize, usize, usize)>::new();
    let mut root_seen = false;
    let mut root_closed = false;
    let mut events = 0usize;
    loop {
        if let Some(execution) = execution {
            execution.check()?;
            execution.work(1)?;
        }
        events = events
            .checked_add(1)
            .ok_or_else(|| owner_invalid("drawing anchor event count overflow"))?;
        if events > limits.max_mce_events {
            return Err(owner_limit(
                "drawing anchor events",
                events,
                limits.max_mce_events,
            ));
        }
        let event = reader
            .read_event()
            .map_err(|error| owner_invalid(error.to_string()))?;
        let resolver = reader.resolver().clone();
        match event {
            Event::Start(element) => {
                let (namespace, local) = resolver.resolve_element(element.name());
                let namespace = namespace_bytes(namespace);
                let local = local.as_ref().to_vec();
                if stack.is_empty() {
                    if root_seen || namespace.as_slice() != XDR || local.as_slice() != b"wsDr" {
                        return Err(owner_invalid(
                            "DrawingML XML has an unexpected or duplicate root element",
                        ));
                    }
                    root_seen = true;
                } else if root_closed {
                    return Err(owner_invalid("DrawingML XML has content after its root"));
                }
                if namespace.as_slice() == XDR
                    && matches!(
                        local.as_slice(),
                        b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor"
                    )
                {
                    if !stack_ends_with(&stack, &[(XDR, b"wsDr")]) {
                        return Err(owner_invalid("DrawingML anchor is not a direct wsDr child"));
                    }
                    if anchors.len() >= limits.max_shapes {
                        return Err(owner_limit(
                            "DrawingML anchor identities",
                            anchors.len() + 1,
                            limits.max_shapes,
                        ));
                    }
                    anchors
                        .try_reserve(1)
                        .map_err(|source| owner_alloc("DrawingML anchor identities", source))?;
                    anchors.push((local.clone(), 0, 0, 0, 0, 0));
                } else if namespace.as_slice() == XDR {
                    if let Some(anchor) = anchors.last_mut() {
                        match local.as_slice() {
                            b"from"
                                if stack.last().is_some_and(|(ns, parent)| {
                                    ns.as_slice() == XDR && parent.as_slice() == anchor.0.as_slice()
                                }) =>
                            {
                                anchor.1 = anchor.1.saturating_add(1)
                            },
                            b"to"
                                if stack.last().is_some_and(|(ns, parent)| {
                                    ns.as_slice() == XDR && parent.as_slice() == anchor.0.as_slice()
                                }) =>
                            {
                                anchor.2 = anchor.2.saturating_add(1)
                            },
                            b"ext"
                                if stack.last().is_some_and(|(ns, parent)| {
                                    ns.as_slice() == XDR && parent.as_slice() == anchor.0.as_slice()
                                }) =>
                            {
                                anchor.3 = anchor.3.saturating_add(1)
                            },
                            b"clientData"
                                if stack.last().is_some_and(|(ns, parent)| {
                                    ns.as_slice() == XDR && parent.as_slice() == anchor.0.as_slice()
                                }) =>
                            {
                                anchor.4 = anchor.4.saturating_add(1)
                            },
                            b"pos"
                                if stack.last().is_some_and(|(ns, parent)| {
                                    ns.as_slice() == XDR && parent.as_slice() == anchor.0.as_slice()
                                }) =>
                            {
                                anchor.5 = anchor.5.saturating_add(1)
                            },
                            _ => {},
                        }
                    }
                }
                if stack.len() >= limits.max_mce_depth {
                    return Err(owner_limit(
                        "drawing anchor depth",
                        stack.len() + 1,
                        limits.max_mce_depth,
                    ));
                }
                stack
                    .try_reserve(1)
                    .map_err(|source| owner_alloc("drawing anchor stack", source))?;
                stack.push((namespace, local));
            },
            Event::Empty(element) => {
                let (namespace, local) = resolver.resolve_element(element.name());
                if is_ns(&namespace, XDR) {
                    if let Some(anchor) = anchors.last_mut() {
                        match local.as_ref() {
                            b"from"
                                if stack.last().is_some_and(|(ns, parent)| {
                                    ns.as_slice() == XDR && parent.as_slice() == anchor.0.as_slice()
                                }) =>
                            {
                                anchor.1 = anchor.1.saturating_add(1)
                            },
                            b"to"
                                if stack.last().is_some_and(|(ns, parent)| {
                                    ns.as_slice() == XDR && parent.as_slice() == anchor.0.as_slice()
                                }) =>
                            {
                                anchor.2 = anchor.2.saturating_add(1)
                            },
                            b"ext"
                                if stack.last().is_some_and(|(ns, parent)| {
                                    ns.as_slice() == XDR && parent.as_slice() == anchor.0.as_slice()
                                }) =>
                            {
                                anchor.3 = anchor.3.saturating_add(1)
                            },
                            b"clientData"
                                if stack.last().is_some_and(|(ns, parent)| {
                                    ns.as_slice() == XDR && parent.as_slice() == anchor.0.as_slice()
                                }) =>
                            {
                                anchor.4 = anchor.4.saturating_add(1)
                            },
                            b"pos"
                                if stack.last().is_some_and(|(ns, parent)| {
                                    ns.as_slice() == XDR && parent.as_slice() == anchor.0.as_slice()
                                }) =>
                            {
                                anchor.5 = anchor.5.saturating_add(1)
                            },
                            _ => {},
                        }
                    }
                    if matches!(
                        local.as_ref(),
                        b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor"
                    ) {
                        return Err(owner_invalid("DrawingML anchor is empty"));
                    }
                }
            },
            Event::End(end) => {
                let closed = pop_checked_end(&mut stack, &end, &resolver)?;
                if closed.0.as_slice() == XDR && {
                    let local = closed.1.as_slice();
                    matches!(
                        local,
                        b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor"
                    )
                } {
                    let (kind, from, to, ext, client_data, pos) = anchors
                        .pop()
                        .ok_or_else(|| owner_invalid("DrawingML anchor state is inconsistent"))?;
                    let valid = match kind.as_slice() {
                        b"twoCellAnchor" => from == 1 && to == 1 && client_data == 1,
                        b"oneCellAnchor" => from == 1 && ext == 1 && client_data == 1,
                        b"absoluteAnchor" => pos == 1 && ext == 1 && client_data == 1,
                        _ => false,
                    };
                    if !valid {
                        return Err(owner_invalid(
                            "DrawingML anchor closure is incomplete or ambiguous",
                        ));
                    }
                } else if closed.0.as_slice() == XDR && closed.1.as_slice() == b"wsDr" {
                    root_closed = true;
                }
            },
            Event::Eof => break,
            Event::Decl(_)
            | Event::Comment(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(owner_invalid("DrawingML XML is unterminated"));
    }
    if !anchors.is_empty() {
        return Err(owner_invalid("DrawingML anchor is unterminated"));
    }
    Ok(())
}

fn repair_renamed_capability_prefix(xml: &[u8], capability: &str) -> Option<Vec<u8>> {
    let source = std::str::from_utf8(xml).ok()?;
    if !source.contains("mc:Ignorable=\"a14\"") {
        return None;
    }
    let declaration = format!("=\"{capability}\"");
    let mut alias = None;
    let mut search = 0;
    while let Some(offset) = source[search..].find("xmlns:") {
        let start = search + offset + "xmlns:".len();
        let equals = source[start..].find('=')? + start;
        let prefix = &source[start..equals];
        if !prefix.is_empty() && source[equals..].starts_with(&declaration) {
            alias = Some(prefix);
            break;
        }
        search = equals.saturating_add(1);
    }
    let alias = alias?;
    if alias == "a14" {
        return None;
    }
    let replacement = format!("mc:Ignorable=\"{alias}\"");
    Some(
        source
            .replace("mc:Ignorable=\"a14\"", &replacement)
            .into_bytes(),
    )
}

#[derive(Clone, Debug)]
struct SourceControl {
    shape_id: u64,
    relationship_id: String,
    name: Option<String>,
    anchor_profile: Option<String>,
    anchor_diagnostic: Option<String>,
}

struct WorksheetScan {
    controls: Vec<SourceControl>,
    drawing_relationship_id: Option<String>,
    vml_relationship_id: Option<String>,
}

fn stack_ends_with(stack: &[(Vec<u8>, Vec<u8>)], path: &[(&[u8], &[u8])]) -> bool {
    stack.len() >= path.len()
        && stack[stack.len() - path.len()..].iter().zip(path).all(
            |((namespace, local), (wanted_namespace, wanted_local))| {
                namespace.as_slice() == *wanted_namespace && local.as_slice() == *wanted_local
            },
        )
}

fn stack_starts_with(stack: &[(Vec<u8>, Vec<u8>)], path: &[(&[u8], &[u8])]) -> bool {
    stack.len() >= path.len()
        && stack[..path.len()].iter().zip(path).all(
            |((namespace, local), (wanted_namespace, wanted_local))| {
                namespace.as_slice() == *wanted_namespace && local.as_slice() == *wanted_local
            },
        )
}

fn stack_is_under_admitted_cnvpr(stack: &[(Vec<u8>, Vec<u8>)]) -> bool {
    stack_starts_with(
        stack,
        &[
            (XDR, b"wsDr"),
            (XDR, b"twoCellAnchor"),
            (XDR, b"sp"),
            (XDR, b"nvSpPr"),
            (XDR, b"cNvPr"),
        ],
    ) || stack_starts_with(
        stack,
        &[
            (XDR, b"wsDr"),
            (XDR, b"oneCellAnchor"),
            (XDR, b"sp"),
            (XDR, b"nvSpPr"),
            (XDR, b"cNvPr"),
        ],
    ) || stack_starts_with(
        stack,
        &[
            (XDR, b"wsDr"),
            (XDR, b"absoluteAnchor"),
            (XDR, b"sp"),
            (XDR, b"nvSpPr"),
            (XDR, b"cNvPr"),
        ],
    )
}

fn parse_worksheet(
    xml: &[u8],
    limits: &OwnerLimits,
    execution: Option<&OwnerExecution>,
) -> OwnerResult<WorksheetScan> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut stack: Vec<(Vec<u8>, Vec<u8>)> = Vec::new();
    let mut controls = Vec::new();
    controls
        .try_reserve(8.min(limits.max_controls))
        .map_err(|source| owner_alloc("worksheet controls", source))?;
    let mut drawing = None;
    let mut vml = None;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut controls_container_seen = false;
    let mut current_control: Option<(SourceControl, bool, bool, usize, usize)> = None;
    let mut events = 0usize;
    loop {
        if let Some(execution) = execution {
            execution.check()?;
            execution.work(1)?;
        }
        events = events
            .checked_add(1)
            .ok_or_else(|| owner_invalid("worksheet event count overflow"))?;
        if events > limits.max_mce_events {
            return Err(owner_limit(
                "worksheet XML events",
                events,
                limits.max_mce_events,
            ));
        }
        let event = reader
            .read_event()
            .map_err(|error| owner_invalid(error.to_string()))?;
        let resolver = reader.resolver().clone();
        match event {
            Event::Start(element) => {
                let (namespace, local) = resolver.resolve_element(element.name());
                let namespace_bytes = namespace_bytes(namespace.clone());
                let local_bytes = local.as_ref().to_vec();
                if stack.is_empty() {
                    if root_seen || namespace_bytes.as_slice() != SML || local_bytes != b"worksheet"
                    {
                        return Err(owner_invalid(
                            "worksheet XML has an unexpected or duplicate root element",
                        ));
                    }
                    root_seen = true;
                } else if root_closed {
                    return Err(owner_invalid("worksheet XML has content after its root"));
                }
                if is_ns(&namespace, SML) {
                    if local.as_ref() == b"controls" {
                        if !stack_ends_with(&stack, &[(SML, b"worksheet")])
                            || controls_container_seen
                        {
                            return Err(owner_invalid(
                                "worksheet controls container is not a unique direct child",
                            ));
                        }
                        controls_container_seen = true;
                    }
                    if local.as_ref() == b"control" {
                        if !stack_ends_with(&stack, &[(SML, b"worksheet"), (SML, b"controls")]) {
                            return Err(owner_invalid(
                                "worksheet control is not a direct child of controls",
                            ));
                        }
                        if current_control.is_some() {
                            return Err(owner_invalid("worksheet controls are nested"));
                        }
                        current_control = Some((
                            parse_control(&element, &resolver, reader.decoder(), limits)?,
                            false,
                            false,
                            0,
                            0,
                        ));
                    }
                    if let Some((_, control_pr, anchor, from, to)) = current_control.as_mut() {
                        let direct_control_pr = stack_ends_with(
                            &stack,
                            &[(SML, b"worksheet"), (SML, b"controls"), (SML, b"control")],
                        );
                        let direct_anchor = stack_ends_with(
                            &stack,
                            &[
                                (SML, b"worksheet"),
                                (SML, b"controls"),
                                (SML, b"control"),
                                (SML, b"controlPr"),
                            ],
                        );
                        let direct_endpoint = stack_ends_with(
                            &stack,
                            &[
                                (SML, b"worksheet"),
                                (SML, b"controls"),
                                (SML, b"control"),
                                (SML, b"controlPr"),
                                (SML, b"anchor"),
                            ],
                        );
                        match local.as_ref() {
                            b"controlPr" if direct_control_pr => {
                                if *control_pr {
                                    return Err(owner_invalid(
                                        "worksheet control has duplicate controlPr elements",
                                    ));
                                }
                                *control_pr = true;
                            },
                            b"controlPr" => {
                                return Err(owner_invalid(
                                    "worksheet controlPr is not a direct control child",
                                ));
                            },
                            b"anchor" if direct_anchor => {
                                if *anchor {
                                    return Err(owner_invalid(
                                        "worksheet controlPr has duplicate anchors",
                                    ));
                                }
                                *anchor = true;
                            },
                            b"anchor" => {
                                return Err(owner_invalid(
                                    "worksheet control anchor is not a direct controlPr child",
                                ));
                            },
                            b"from" if direct_endpoint => *from = from.saturating_add(1),
                            b"to" if direct_endpoint => *to = to.saturating_add(1),
                            b"from" | b"to" => {
                                return Err(owner_invalid(
                                    "worksheet anchor endpoint is not a direct anchor child",
                                ));
                            },
                            _ => {},
                        }
                    }
                    if local.as_ref() == b"drawing" {
                        if !stack_ends_with(&stack, &[(SML, b"worksheet")]) {
                            return Err(owner_invalid(
                                "worksheet drawing is not a direct worksheet child",
                            ));
                        }
                        if drawing.is_some() {
                            return Err(owner_invalid(
                                "worksheet has duplicate drawing references",
                            ));
                        }
                        drawing = attr_rel_id(&element, &resolver, reader.decoder(), limits)?;
                    }
                    if local.as_ref() == b"legacyDrawing" {
                        if !stack_ends_with(&stack, &[(SML, b"worksheet")]) {
                            return Err(owner_invalid(
                                "worksheet legacyDrawing is not a direct worksheet child",
                            ));
                        }
                        if vml.is_some() {
                            return Err(owner_invalid(
                                "worksheet has duplicate legacyDrawing references",
                            ));
                        }
                        vml = attr_rel_id(&element, &resolver, reader.decoder(), limits)?;
                    }
                }
                if stack.len() >= limits.max_mce_depth {
                    return Err(owner_limit(
                        "worksheet XML depth",
                        stack.len() + 1,
                        limits.max_mce_depth,
                    ));
                }
                stack
                    .try_reserve(1)
                    .map_err(|source| owner_alloc("worksheet XML stack", source))?;
                stack.push((namespace_bytes, local_bytes));
            },
            Event::Empty(element) => {
                let (namespace, local) = resolver.resolve_element(element.name());
                let namespace_bytes = namespace_bytes(namespace.clone());
                let local_bytes = local.as_ref();
                if stack.is_empty() {
                    if root_seen || namespace_bytes.as_slice() != SML || local_bytes != b"worksheet"
                    {
                        return Err(owner_invalid(
                            "worksheet XML has an unexpected or duplicate root element",
                        ));
                    }
                    root_seen = true;
                    root_closed = true;
                    continue;
                }
                if root_closed {
                    return Err(owner_invalid("worksheet XML has content after its root"));
                }
                if is_ns(&namespace, SML) {
                    if local.as_ref() == b"control" {
                        if !stack_ends_with(&stack, &[(SML, b"worksheet"), (SML, b"controls")]) {
                            return Err(owner_invalid(
                                "worksheet control is not a direct child of controls",
                            ));
                        }
                        let mut control =
                            parse_control(&element, &resolver, reader.decoder(), limits)?;
                        control.anchor_diagnostic =
                            Some("worksheet control has no LoSmlAnchorV1 controlPr anchor".into());
                        if controls.len() >= limits.max_controls {
                            return Err(owner_limit(
                                "effective controls",
                                controls.len() + 1,
                                limits.max_controls,
                            ));
                        }
                        controls
                            .try_reserve(1)
                            .map_err(|source| owner_alloc("worksheet controls", source))?;
                        controls.push(control);
                    }
                    if local.as_ref() == b"controls" {
                        return Err(owner_invalid(
                            "worksheet controls container cannot be empty",
                        ));
                    }
                    if let Some((_, control_pr, anchor, from, to)) = current_control.as_mut() {
                        let direct_control_pr = stack_ends_with(
                            &stack,
                            &[(SML, b"worksheet"), (SML, b"controls"), (SML, b"control")],
                        );
                        let direct_anchor = stack_ends_with(
                            &stack,
                            &[
                                (SML, b"worksheet"),
                                (SML, b"controls"),
                                (SML, b"control"),
                                (SML, b"controlPr"),
                            ],
                        );
                        let direct_endpoint = stack_ends_with(
                            &stack,
                            &[
                                (SML, b"worksheet"),
                                (SML, b"controls"),
                                (SML, b"control"),
                                (SML, b"controlPr"),
                                (SML, b"anchor"),
                            ],
                        );
                        match local.as_ref() {
                            b"controlPr" if direct_control_pr => {
                                if *control_pr {
                                    return Err(owner_invalid(
                                        "worksheet control has duplicate controlPr elements",
                                    ));
                                }
                                *control_pr = true;
                            },
                            b"controlPr" => {
                                return Err(owner_invalid(
                                    "worksheet controlPr is not a direct control child",
                                ));
                            },
                            b"anchor" if direct_anchor => {
                                if *anchor {
                                    return Err(owner_invalid(
                                        "worksheet controlPr has duplicate anchors",
                                    ));
                                }
                                *anchor = true;
                            },
                            b"anchor" => {
                                return Err(owner_invalid(
                                    "worksheet control anchor is not a direct controlPr child",
                                ));
                            },
                            b"from" if direct_endpoint => *from = from.saturating_add(1),
                            b"to" if direct_endpoint => *to = to.saturating_add(1),
                            b"from" | b"to" => {
                                return Err(owner_invalid(
                                    "worksheet anchor endpoint is not a direct anchor child",
                                ));
                            },
                            _ => {},
                        }
                    }
                    if local.as_ref() == b"drawing" {
                        if !stack_ends_with(&stack, &[(SML, b"worksheet")]) {
                            return Err(owner_invalid(
                                "worksheet drawing is not a direct worksheet child",
                            ));
                        }
                        if drawing.is_some() {
                            return Err(owner_invalid(
                                "worksheet has duplicate drawing references",
                            ));
                        }
                        drawing = attr_rel_id(&element, &resolver, reader.decoder(), limits)?;
                    }
                    if local.as_ref() == b"legacyDrawing" {
                        if !stack_ends_with(&stack, &[(SML, b"worksheet")]) {
                            return Err(owner_invalid(
                                "worksheet legacyDrawing is not a direct worksheet child",
                            ));
                        }
                        if vml.is_some() {
                            return Err(owner_invalid(
                                "worksheet has duplicate legacyDrawing references",
                            ));
                        }
                        vml = attr_rel_id(&element, &resolver, reader.decoder(), limits)?;
                    }
                }
            },
            Event::End(end) => {
                let closed = pop_checked_end(&mut stack, &end, &resolver)?;
                if closed.0.as_slice() == SML && closed.1.as_slice() == b"worksheet" {
                    root_closed = true;
                }
                if closed.0.as_slice() == SML && closed.1.as_slice() == b"control" {
                    let (mut control, _control_pr, anchor, from, to) = current_control
                        .take()
                        .ok_or_else(|| owner_invalid("worksheet control state is inconsistent"))?;
                    if anchor && from == 1 && to == 1 {
                        control.anchor_profile = Some("LoSmlAnchorV1".into());
                    } else {
                        control.anchor_diagnostic = Some(
                            "worksheet controlPr anchor is incomplete or ambiguous for LoSmlAnchorV1".into(),
                        );
                    }
                    if controls.len() >= limits.max_controls {
                        return Err(owner_limit(
                            "effective controls",
                            controls.len() + 1,
                            limits.max_controls,
                        ));
                    }
                    controls
                        .try_reserve(1)
                        .map_err(|source| owner_alloc("worksheet controls", source))?;
                    controls.push(control);
                }
            },
            Event::Eof => break,
            Event::Decl(_)
            | Event::Comment(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if current_control.is_some() {
        return Err(owner_invalid("worksheet control is unterminated"));
    }
    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(owner_invalid(
            "worksheet XML is unterminated or has no root",
        ));
    }
    if controls.len() > limits.max_controls {
        return Err(owner_limit(
            "effective controls",
            controls.len(),
            limits.max_controls,
        ));
    }
    let anchor_diagnostics = validate_losml_anchor_values(xml, limits, execution)?;
    if anchor_diagnostics.len() != controls.len() {
        return Err(owner_invalid(
            "worksheet LoSmlAnchorV1 control inventory is inconsistent",
        ));
    }
    for (control, diagnostic) in controls.iter_mut().zip(anchor_diagnostics) {
        if let Some(diagnostic) = diagnostic {
            control.anchor_profile = None;
            control.anchor_diagnostic = Some(diagnostic);
        }
    }
    Ok(WorksheetScan {
        controls,
        drawing_relationship_id: drawing,
        vml_relationship_id: vml,
    })
}

const MAX_ANCHOR_COLUMN: u64 = 16_383;
const MAX_ANCHOR_ROW: u64 = 1_048_575;
const MAX_ANCHOR_OFFSET: u64 = 2_147_483_647;

#[derive(Default)]
struct AnchorEndpoint {
    kind: Vec<u8>,
    fields: [Option<String>; 4],
    active_field: Option<(usize, Vec<u8>)>,
}

#[derive(Default)]
struct AnchorValidation {
    seen: bool,
    from: Option<AnchorEndpoint>,
    to: Option<AnchorEndpoint>,
    endpoint: Option<AnchorEndpoint>,
    diagnostic: Option<String>,
}

fn validate_losml_anchor_values(
    xml: &[u8],
    limits: &OwnerLimits,
    execution: Option<&OwnerExecution>,
) -> OwnerResult<Vec<Option<String>>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut stack: Vec<(Vec<u8>, Vec<u8>)> = Vec::new();
    let mut validations: Vec<Option<String>> = Vec::new();
    let mut current: Option<AnchorValidation> = None;
    let mut events = 0usize;
    loop {
        if let Some(execution) = execution {
            execution.check()?;
            execution.work(1)?;
        }
        events = events
            .checked_add(1)
            .ok_or_else(|| owner_invalid("worksheet anchor event count overflow"))?;
        if events > limits.max_mce_events {
            return Err(owner_limit(
                "worksheet anchor events",
                events,
                limits.max_mce_events,
            ));
        }
        let event = reader
            .read_event()
            .map_err(|error| owner_invalid(error.to_string()))?;
        let resolver = reader.resolver().clone();
        match event {
            Event::Start(element) => {
                let (namespace, local) = resolver.resolve_element(element.name());
                let namespace = namespace_bytes(namespace);
                let local = local.as_ref().to_vec();
                if namespace.as_slice() == SML && local.as_slice() == b"control" {
                    if current.is_some() {
                        return Err(owner_invalid("worksheet controls are nested"));
                    }
                    current = Some(AnchorValidation::default());
                } else if let Some(state) = current.as_mut() {
                    if namespace.as_slice() == SML
                        && local.as_slice() == b"anchor"
                        && stack.last().is_some_and(|(parent_ns, parent)| {
                            parent_ns.as_slice() == SML && parent.as_slice() == b"controlPr"
                        })
                    {
                        if state.seen {
                            state.diagnostic = Some(
                                "worksheet controlPr contains duplicate LoSmlAnchorV1 anchors"
                                    .into(),
                            );
                        }
                        state.seen = true;
                    } else if namespace.as_slice() == SML
                        && matches!(local.as_slice(), b"from" | b"to")
                        && stack.last().is_some_and(|(parent_ns, parent)| {
                            parent_ns.as_slice() == SML && parent.as_slice() == b"anchor"
                        })
                    {
                        if state.endpoint.is_some() {
                            state.diagnostic =
                                Some("worksheet LoSmlAnchorV1 has nested anchor endpoints".into());
                        } else {
                            state.endpoint = Some(AnchorEndpoint {
                                kind: local.clone(),
                                ..AnchorEndpoint::default()
                            });
                        }
                    } else if namespace.as_slice() == XDR
                        && anchor_field_index(local.as_slice()).is_some()
                        && state.endpoint.as_ref().is_some_and(|endpoint| {
                            stack.last().is_some_and(|(parent_ns, parent)| {
                                parent_ns.as_slice() == SML
                                    && parent.as_slice() == endpoint.kind.as_slice()
                            })
                        })
                    {
                        let index = anchor_field_index(local.as_slice()).expect("checked above");
                        if let Some(endpoint) = state.endpoint.as_mut() {
                            if endpoint.fields[index].is_some() || endpoint.active_field.is_some() {
                                state.diagnostic = Some(format!(
                                    "worksheet LoSmlAnchorV1 endpoint field {} is duplicated",
                                    String::from_utf8_lossy(&local)
                                ));
                            } else {
                                endpoint.active_field = Some((index, local.clone()));
                            }
                        }
                    }
                }
                if stack.len() >= limits.max_mce_depth {
                    return Err(owner_limit(
                        "worksheet anchor depth",
                        stack.len() + 1,
                        limits.max_mce_depth,
                    ));
                }
                stack
                    .try_reserve(1)
                    .map_err(|source| owner_alloc("worksheet anchor stack", source))?;
                stack.push((namespace, local));
            },
            Event::Empty(element) => {
                let (namespace, local) = resolver.resolve_element(element.name());
                let namespace = namespace_bytes(namespace);
                let local = local.as_ref().to_vec();
                if namespace.as_slice() == SML && local.as_slice() == b"control" {
                    if current.is_some() {
                        return Err(owner_invalid("worksheet controls are nested"));
                    }
                    if validations.len() >= limits.max_controls {
                        return Err(owner_limit(
                            "worksheet anchor validations",
                            validations.len() + 1,
                            limits.max_controls,
                        ));
                    }
                    validations
                        .try_reserve(1)
                        .map_err(|source| owner_alloc("worksheet anchor validations", source))?;
                    validations.push(None);
                } else if let Some(state) = current.as_mut() {
                    if namespace.as_slice() == XDR && anchor_field_index(local.as_slice()).is_some()
                    {
                        state.diagnostic = Some(format!(
                            "worksheet LoSmlAnchorV1 field {} has no numeric lexical value",
                            String::from_utf8_lossy(&local)
                        ));
                    }
                }
            },
            Event::Text(text) => append_anchor_text(
                current.as_mut(),
                text.as_ref(),
                reader.decoder(),
                limits.max_name_bytes,
            )?,
            Event::CData(text) => append_anchor_text(
                current.as_mut(),
                text.as_ref(),
                reader.decoder(),
                limits.max_name_bytes,
            )?,
            Event::End(end) => {
                let closed = pop_checked_end(&mut stack, &end, &resolver)?;
                if let Some(state) = current.as_mut() {
                    if closed.0.as_slice() == XDR
                        && anchor_field_index(closed.1.as_slice()).is_some()
                    {
                        if let Some(endpoint) = state.endpoint.as_mut() {
                            if let Some((index, _)) = endpoint.active_field.take() {
                                if endpoint.fields[index].is_none() {
                                    state.diagnostic = Some(
                                        "worksheet LoSmlAnchorV1 numeric field is empty".into(),
                                    );
                                }
                            }
                        }
                    } else if closed.0.as_slice() == SML
                        && matches!(closed.1.as_slice(), b"from" | b"to")
                    {
                        if let Some(endpoint) = state.endpoint.take() {
                            if endpoint.fields.iter().any(Option::is_none) {
                                state.diagnostic = Some(format!(
                                    "worksheet LoSmlAnchorV1 {} endpoint is incomplete",
                                    String::from_utf8_lossy(&endpoint.kind)
                                ));
                            } else {
                                for (index, value) in endpoint.fields.iter().enumerate() {
                                    let value = value.as_deref().unwrap_or_default();
                                    let maximum = match index {
                                        0 => MAX_ANCHOR_COLUMN,
                                        1 => MAX_ANCHOR_OFFSET,
                                        2 => MAX_ANCHOR_ROW,
                                        _ => MAX_ANCHOR_OFFSET,
                                    };
                                    if !valid_anchor_integer(value, maximum) {
                                        state.diagnostic = Some(format!(
                                            "worksheet LoSmlAnchorV1 {} field {} is outside its decimal range",
                                            String::from_utf8_lossy(&endpoint.kind),
                                            anchor_field_name(index)
                                        ));
                                    }
                                }
                            }
                            if endpoint.kind.as_slice() == b"from" {
                                if state.from.is_some() {
                                    state.diagnostic = Some(
                                        "worksheet LoSmlAnchorV1 has duplicate from endpoints"
                                            .into(),
                                    );
                                }
                                state.from = Some(endpoint);
                            } else {
                                if state.to.is_some() {
                                    state.diagnostic = Some(
                                        "worksheet LoSmlAnchorV1 has duplicate to endpoints".into(),
                                    );
                                }
                                state.to = Some(endpoint);
                            }
                        }
                    } else if closed.0.as_slice() == SML && closed.1.as_slice() == b"anchor" {
                        if state.from.is_none() || state.to.is_none() {
                            state.diagnostic = Some(
                                "worksheet LoSmlAnchorV1 anchor is missing from/to endpoints"
                                    .into(),
                            );
                        }
                    } else if closed.0.as_slice() == SML && closed.1.as_slice() == b"control" {
                        let state = current.take().ok_or_else(|| {
                            owner_invalid("worksheet anchor state is inconsistent")
                        })?;
                        if validations.len() >= limits.max_controls {
                            return Err(owner_limit(
                                "worksheet anchor validations",
                                validations.len() + 1,
                                limits.max_controls,
                            ));
                        }
                        validations.try_reserve(1).map_err(|source| {
                            owner_alloc("worksheet anchor validations", source)
                        })?;
                        validations.push(state.diagnostic);
                    }
                }
            },
            Event::Eof => break,
            Event::Decl(_)
            | Event::Comment(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if current.is_some() || !stack.is_empty() {
        return Err(owner_invalid("worksheet anchor XML is unterminated"));
    }
    Ok(validations)
}

fn append_anchor_text(
    current: Option<&mut AnchorValidation>,
    bytes: &[u8],
    decoder: quick_xml::encoding::Decoder,
    maximum: usize,
) -> OwnerResult<()> {
    let Some(state) = current else {
        return Ok(());
    };
    let Some(endpoint) = state.endpoint.as_mut() else {
        return Ok(());
    };
    let Some((index, _)) = endpoint.active_field.take() else {
        return Ok(());
    };
    let decoded = decoder
        .decode(bytes)
        .map_err(|error| owner_invalid(error.to_string()))?;
    if decoded.len() > maximum {
        state.diagnostic =
            Some("worksheet LoSmlAnchorV1 numeric lexical value exceeds the owner bound".into());
    } else if endpoint.fields[index].is_some() {
        state.diagnostic =
            Some("worksheet LoSmlAnchorV1 numeric field has multiple text nodes".into());
    } else {
        endpoint.fields[index] = Some(decoded.into_owned());
    }
    endpoint.active_field = Some((index, Vec::new()));
    Ok(())
}

fn anchor_field_index(local: &[u8]) -> Option<usize> {
    match local {
        b"col" => Some(0),
        b"colOff" => Some(1),
        b"row" => Some(2),
        b"rowOff" => Some(3),
        _ => None,
    }
}

const fn anchor_field_name(index: usize) -> &'static str {
    match index {
        0 => "col",
        1 => "colOff",
        2 => "row",
        _ => "rowOff",
    }
}

fn valid_anchor_integer(value: &str, maximum: u64) -> bool {
    !value.is_empty()
        && value.bytes().all(|byte| byte.is_ascii_digit())
        && value.parse::<u64>().is_ok_and(|value| value <= maximum)
}

fn relationship_target(
    relationships: &Relationships,
    relationship_id: Option<&str>,
    expected_type: &'static str,
    subject: &'static str,
) -> OwnerResult<PackURI> {
    let relationship_id = relationship_id
        .ok_or_else(|| owner_invalid(format!("{subject} relationship is missing")))?;
    let relationship = relationships
        .get(relationship_id)
        .ok_or_else(|| owner_invalid(format!("{subject} relationship is missing")))?;
    if relationship.target_mode() != TargetMode::Internal {
        return Err(owner_invalid(format!(
            "{subject} relationship target is external"
        )));
    }
    if relationship.reltype() != expected_type {
        return Err(owner_invalid(format!(
            "{subject} relationship type is not canonical"
        )));
    }
    relationship
        .target_partname()
        .map_err(FormControlOwnerError::from)
}

fn parse_control(
    element: &BytesStart<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    limits: &OwnerLimits,
) -> OwnerResult<SourceControl> {
    let mut shape = None;
    let mut rel_id = None;
    let mut name = None;
    for attribute in element.unchecked_attributes() {
        let attribute = attribute.map_err(|error| owner_invalid(error.to_string()))?;
        let value = bounded_name(
            decode_attribute_bounded(
                attribute.value.as_ref(),
                decoder,
                limits,
                "worksheet control identity",
            )?,
            limits,
            "worksheet control identity",
        )?;
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Unbound) && local.as_ref() == b"shapeId" {
            shape = Some(
                value
                    .parse::<u64>()
                    .map_err(|_| owner_invalid("control shapeId is not decimal"))?,
            )
        } else if local.as_ref() == b"id" && is_ns(&namespace, REL) {
            rel_id = Some(value);
        } else if matches!(namespace, ResolveResult::Unbound) && local.as_ref() == b"name" {
            name = Some(bounded_display_name(
                value,
                limits,
                "worksheet control name",
            )?)
        } else if local.as_ref() == b"id" && is_ns(&namespace, STRICT_REL) {
            return Err(FormControlOwnerError::Invalid(
                "strict worksheet relationship attribute is not admitted".into(),
            ));
        }
    }
    let shape_id = shape.ok_or_else(|| owner_invalid("control shapeId is missing"))?;
    if !(1..=MAX_SHAPE_ID).contains(&shape_id) {
        return Err(owner_invalid("control shapeId is outside the Office range"));
    }
    Ok(SourceControl {
        shape_id,
        relationship_id: rel_id.ok_or_else(|| owner_invalid("control r:id is missing"))?,
        name,
        anchor_profile: None,
        anchor_diagnostic: None,
    })
}

fn attr_rel_id(
    element: &BytesStart<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    limits: &OwnerLimits,
) -> OwnerResult<Option<String>> {
    for attribute in element.unchecked_attributes() {
        let attribute = attribute.map_err(|error| owner_invalid(error.to_string()))?;
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if local.as_ref() == b"id" {
            if is_ns(&namespace, STRICT_REL) {
                return Err(FormControlOwnerError::Invalid(
                    "strict worksheet relationship attribute is not admitted".into(),
                ));
            }
            if is_ns(&namespace, REL) {
                return Ok(Some(bounded_name(
                    decode_attribute_bounded(
                        attribute.value.as_ref(),
                        decoder,
                        limits,
                        "worksheet relationship identity",
                    )?,
                    limits,
                    "worksheet relationship identity",
                )?));
            }
        }
    }
    Ok(None)
}

#[derive(Clone, Debug)]
struct DrawingShape {
    id: u64,
    name: Option<String>,
    compat_spid: String,
}

fn index_drawing_shapes(
    shapes: &[DrawingShape],
    execution: Option<&OwnerExecution>,
) -> OwnerResult<HashMap<u64, Option<usize>>> {
    let mut index = HashMap::new();
    index
        .try_reserve(shapes.len())
        .map_err(|source| owner_alloc("DrawingML identity index", source))?;
    for (position, shape) in shapes.iter().enumerate() {
        if let Some(execution) = execution {
            execution.check()?;
            execution.work(1)?;
        }
        match index.entry(shape.id) {
            std::collections::hash_map::Entry::Vacant(entry) => {
                entry.insert(Some(position));
            },
            std::collections::hash_map::Entry::Occupied(mut entry) => {
                entry.insert(None);
            },
        }
    }
    Ok(index)
}

fn is_admitted_sp_cnvpr(stack: &[(Vec<u8>, Vec<u8>)]) -> bool {
    // The owner only admits the non-visual properties of a shape (`sp`) when
    // the cNvPr is its direct child through `nvSpPr`, and that shape is itself
    // directly descended from one of the three admitted anchor kinds.  A
    // global cNvPr scan would otherwise accidentally adopt a picture or an
    // unrelated shape with the same numeric identity.
    if stack.len() < 4 {
        return false;
    }
    let parent = &stack[stack.len() - 1];
    let grandparent = &stack[stack.len() - 2];
    if parent.0.as_slice() != XDR
        || parent.1.as_slice() != b"nvSpPr"
        || grandparent.0.as_slice() != XDR
        || grandparent.1.as_slice() != b"sp"
    {
        return false;
    }
    let anchor = &stack[stack.len() - 3];
    let drawing_root = &stack[stack.len() - 4];
    anchor.0.as_slice() == XDR
        && matches!(
            anchor.1.as_slice(),
            b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor"
        )
        && drawing_root.0.as_slice() == XDR
        && drawing_root.1.as_slice() == b"wsDr"
}

/// Count shape candidates without constructing shape identities.  The
/// conservative count is charged before the full parser allocates names,
/// fields, or indexes; the strict ancestry parser still decides admission.
fn preflight_shape_nodes(
    xml: &[u8],
    namespace: &[u8],
    local_name: &[u8],
    maximum: usize,
    limits: &OwnerLimits,
    resource: &'static str,
    mut execution: Option<&mut OwnerExecution>,
) -> OwnerResult<usize> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut events = 0usize;
    let mut count = 0usize;
    loop {
        if let Some(execution) = execution.as_deref_mut() {
            execution.check()?;
            execution.work(1)?;
        }
        events = events
            .checked_add(1)
            .ok_or_else(|| owner_invalid("shape preflight event count overflow"))?;
        if events > limits.max_mce_events {
            return Err(owner_limit(
                "shape preflight events",
                events,
                limits.max_mce_events,
            ));
        }
        let event = reader
            .read_event()
            .map_err(|error| owner_invalid(error.to_string()))?;
        let resolver = reader.resolver().clone();
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let (resolved_namespace, local) = resolver.resolve_element(element.name());
                if is_ns(&resolved_namespace, namespace) && local.as_ref() == local_name {
                    count = count
                        .checked_add(1)
                        .ok_or_else(|| owner_invalid("shape identity count overflow"))?;
                    if count > maximum {
                        return Err(owner_limit(resource, count, maximum));
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if let Some(execution) = execution {
        execution.reserve(litchi_core::Resource::Objects, count)?;
    }
    Ok(count)
}

/// Count only worksheet controls whose relationship edge is the admitted
/// `ctrlProp` owner.  This pass deliberately retains no control identity or
/// name: the full worksheet parser remains responsible for ancestry, lexical
/// bounds, and duplicate-field validation after the mirror lower bound has
/// been checked.
fn preflight_form_control_nodes(
    xml: &[u8],
    relationships: &Relationships,
    limits: &OwnerLimits,
    mut execution: Option<&mut OwnerExecution>,
) -> OwnerResult<usize> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut events = 0usize;
    let mut count = 0usize;
    loop {
        if let Some(execution) = execution.as_deref_mut() {
            execution.check()?;
            execution.work(1)?;
        }
        events = events
            .checked_add(1)
            .ok_or_else(|| owner_invalid("form-control census event count overflow"))?;
        if events > limits.max_mce_events {
            return Err(owner_limit(
                "form-control census events",
                events,
                limits.max_mce_events,
            ));
        }
        let event = reader
            .read_event()
            .map_err(|error| owner_invalid(error.to_string()))?;
        let resolver = reader.resolver().clone();
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let (namespace, local) = resolver.resolve_element(element.name());
                if !is_ns(&namespace, SML) || local.as_ref() != b"control" {
                    continue;
                }
                let Some(relationship_id) = preflight_control_relationship_id(
                    &element,
                    &resolver,
                    reader.decoder(),
                    limits,
                )?
                else {
                    continue;
                };
                if relationships
                    .get(&relationship_id)
                    .is_some_and(|relationship| {
                        relationship.reltype() == CONTROL_PROPERTIES_RELATIONSHIP_TYPE
                    })
                {
                    count = count
                        .checked_add(1)
                        .ok_or_else(|| owner_invalid("form-control census count overflow"))?;
                    if count > limits.max_controls {
                        return Err(owner_limit(
                            "effective controls",
                            count,
                            limits.max_controls,
                        ));
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(count)
}

fn preflight_control_relationship_id(
    element: &BytesStart<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    limits: &OwnerLimits,
) -> OwnerResult<Option<String>> {
    for attribute in element.unchecked_attributes() {
        let attribute = attribute.map_err(|error| owner_invalid(error.to_string()))?;
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if local.as_ref() != b"id" {
            continue;
        }
        if is_ns(&namespace, STRICT_REL) {
            return Err(FormControlOwnerError::Invalid(
                "strict worksheet relationship attribute is not admitted".into(),
            ));
        }
        if is_ns(&namespace, REL) {
            return Ok(Some(bounded_name(
                decode_attribute_bounded(
                    attribute.value.as_ref(),
                    decoder,
                    limits,
                    "worksheet relationship identity",
                )?,
                limits,
                "worksheet relationship identity",
            )?));
        }
    }
    Ok(None)
}

fn parse_drawing(
    xml: &[u8],
    limits: &OwnerLimits,
    execution: Option<&OwnerExecution>,
) -> OwnerResult<Vec<DrawingShape>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut stack: Vec<(Vec<u8>, Vec<u8>)> = Vec::new();
    let mut shapes = Vec::new();
    shapes
        .try_reserve(8.min(limits.max_shapes))
        .map_err(|source| owner_alloc("drawing shape identities", source))?;
    let mut current_cnvpr = None;
    let mut shape_cnvpr_seen = false;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut events = 0usize;
    loop {
        if let Some(execution) = execution {
            execution.check()?;
            execution.work(1)?;
        }
        events = events
            .checked_add(1)
            .ok_or_else(|| owner_invalid("drawing XML event count overflow"))?;
        if events > limits.max_mce_events {
            return Err(owner_limit(
                "drawing XML events",
                events,
                limits.max_mce_events,
            ));
        }
        let event = reader
            .read_event()
            .map_err(|error| owner_invalid(error.to_string()))?;
        let resolver = reader.resolver().clone();
        match event {
            Event::Start(element) => {
                let (namespace, local) = resolver.resolve_element(element.name());
                let namespace = namespace_bytes(namespace);
                let local = local.as_ref().to_vec();
                if stack.is_empty() {
                    if root_seen || namespace.as_slice() != XDR || local.as_slice() != b"wsDr" {
                        return Err(owner_invalid(
                            "DrawingML XML has an unexpected or duplicate root element",
                        ));
                    }
                    root_seen = true;
                } else if root_closed {
                    return Err(owner_invalid("DrawingML XML has content after its root"));
                }
                if namespace.as_slice() == XDR
                    && local.as_slice() == b"sp"
                    && !stack_ends_with(&stack, &[(XDR, b"wsDr"), (XDR, b"twoCellAnchor")])
                    && !stack_ends_with(&stack, &[(XDR, b"wsDr"), (XDR, b"oneCellAnchor")])
                    && !stack_ends_with(&stack, &[(XDR, b"wsDr"), (XDR, b"absoluteAnchor")])
                {
                    return Err(owner_invalid(
                        "DrawingML xdr:sp is not a direct admitted-anchor child",
                    ));
                }
                if namespace.as_slice() == XDR && local.as_slice() == b"sp" {
                    shape_cnvpr_seen = false;
                }
                if namespace.as_slice() == XDR && local.as_slice() == b"cNvPr" {
                    let has_shape = stack
                        .iter()
                        .any(|(ns, local)| ns.as_slice() == XDR && local.as_slice() == b"sp");
                    if !has_shape {
                        if stack.len() >= limits.max_mce_depth {
                            return Err(owner_limit(
                                "drawing XML depth",
                                stack.len() + 1,
                                limits.max_mce_depth,
                            ));
                        }
                        stack
                            .try_reserve(1)
                            .map_err(|source| owner_alloc("drawing XML stack", source))?;
                        stack.push((namespace, local));
                        continue;
                    }
                    if !is_admitted_sp_cnvpr(&stack) {
                        return Err(owner_invalid(
                            "DrawingML xdr:sp/cNvPr is outside its admitted anchor",
                        ));
                    }
                    if current_cnvpr.is_some() {
                        return Err(owner_invalid("DrawingML cNvPr elements are nested"));
                    }
                    if shape_cnvpr_seen {
                        return Err(owner_invalid(
                            "DrawingML admitted xdr:sp has duplicate cNvPr elements",
                        ));
                    }
                    if shapes.len() >= limits.max_shapes {
                        return Err(owner_limit(
                            "shape identities",
                            shapes.len() + 1,
                            limits.max_shapes,
                        ));
                    }
                    shapes
                        .try_reserve(1)
                        .map_err(|source| owner_alloc("drawing shape identities", source))?;
                    shapes.push(parse_drawing_shape(
                        &element,
                        &resolver,
                        reader.decoder(),
                        limits,
                    )?);
                    if shapes.len() > limits.max_shapes {
                        return Err(owner_limit(
                            "shape identities",
                            shapes.len(),
                            limits.max_shapes,
                        ));
                    }
                    current_cnvpr = Some(shapes.len() - 1);
                    shape_cnvpr_seen = true;
                } else if namespace.as_slice() == A14 && local.as_slice() == b"compatExt" {
                    if let Some(index) = current_cnvpr {
                        if stack_is_under_admitted_cnvpr(&stack) {
                            let spid = attr_unqualified(
                                &element,
                                b"spid",
                                &resolver,
                                reader.decoder(),
                                limits,
                                "DrawingML compat spid",
                            )?
                            .ok_or_else(|| {
                                owner_invalid("DrawingML a14:compatExt/@spid is missing")
                            })?;
                            if !shapes[index].compat_spid.is_empty() {
                                return Err(owner_invalid(
                                    "DrawingML cNvPr has duplicate a14:compatExt",
                                ));
                            }
                            shapes[index].compat_spid = spid;
                        }
                    }
                }
                if stack.len() >= limits.max_mce_depth {
                    return Err(owner_limit(
                        "drawing XML depth",
                        stack.len() + 1,
                        limits.max_mce_depth,
                    ));
                }
                stack
                    .try_reserve(1)
                    .map_err(|source| owner_alloc("drawing XML stack", source))?;
                stack.push((namespace, local));
            },
            Event::Empty(element) => {
                let (namespace, local) = resolver.resolve_element(element.name());
                if stack.is_empty() {
                    return Err(owner_invalid("DrawingML empty root is not admitted"));
                }
                if root_closed {
                    return Err(owner_invalid("DrawingML XML has content after its root"));
                }
                if is_ns(&namespace, XDR) && local.as_ref() == b"sp" {
                    let direct =
                        stack_ends_with(&stack, &[(XDR, b"wsDr"), (XDR, b"twoCellAnchor")])
                            || stack_ends_with(&stack, &[(XDR, b"wsDr"), (XDR, b"oneCellAnchor")])
                            || stack_ends_with(&stack, &[(XDR, b"wsDr"), (XDR, b"absoluteAnchor")]);
                    if !direct {
                        return Err(owner_invalid(
                            "DrawingML xdr:sp is not a direct admitted-anchor child",
                        ));
                    }
                } else if is_ns(&namespace, XDR) && local.as_ref() == b"cNvPr" {
                    let has_shape = stack
                        .iter()
                        .any(|(ns, local)| ns.as_slice() == XDR && local.as_slice() == b"sp");
                    if !has_shape {
                        continue;
                    }
                    if !is_admitted_sp_cnvpr(&stack) {
                        return Err(owner_invalid(
                            "DrawingML xdr:sp/cNvPr is outside its admitted anchor",
                        ));
                    }
                    if shape_cnvpr_seen {
                        return Err(owner_invalid(
                            "DrawingML admitted xdr:sp has duplicate cNvPr elements",
                        ));
                    }
                    if shapes.len() >= limits.max_shapes {
                        return Err(owner_limit(
                            "shape identities",
                            shapes.len() + 1,
                            limits.max_shapes,
                        ));
                    }
                    shapes
                        .try_reserve(1)
                        .map_err(|source| owner_alloc("drawing shape identities", source))?;
                    let shape = parse_drawing_shape(&element, &resolver, reader.decoder(), limits)?;
                    shapes.push(shape);
                    shape_cnvpr_seen = true;
                    if shapes.len() > limits.max_shapes {
                        return Err(owner_limit(
                            "shape identities",
                            shapes.len(),
                            limits.max_shapes,
                        ));
                    }
                } else if is_ns(&namespace, A14) && local.as_ref() == b"compatExt" {
                    if let Some(index) = current_cnvpr
                        && stack_is_under_admitted_cnvpr(&stack)
                    {
                        let spid = attr_unqualified(
                            &element,
                            b"spid",
                            &resolver,
                            reader.decoder(),
                            limits,
                            "DrawingML compat spid",
                        )?
                        .ok_or_else(|| owner_invalid("DrawingML a14:compatExt/@spid is missing"))?;
                        if !shapes[index].compat_spid.is_empty() {
                            return Err(owner_invalid(
                                "DrawingML cNvPr has duplicate a14:compatExt",
                            ));
                        }
                        shapes[index].compat_spid = spid;
                    }
                }
            },
            Event::End(end) => {
                let closed = pop_checked_end(&mut stack, &end, &resolver)?;
                if closed.0.as_slice() == XDR && closed.1.as_slice() == b"cNvPr" {
                    let Some(_index) = current_cnvpr.take() else {
                        continue;
                    };
                } else if closed.0.as_slice() == XDR && closed.1.as_slice() == b"sp" {
                    shape_cnvpr_seen = false;
                } else if closed.0.as_slice() == XDR && closed.1.as_slice() == b"wsDr" {
                    root_closed = true;
                }
            },
            Event::Eof => break,
            Event::Decl(_)
            | Event::Comment(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(owner_invalid(
            "DrawingML XML is unterminated or has no root",
        ));
    }
    Ok(shapes)
}

fn parse_drawing_shape(
    element: &BytesStart<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    limits: &OwnerLimits,
) -> OwnerResult<DrawingShape> {
    let mut id = None;
    let mut name = None;
    for attribute in element.unchecked_attributes() {
        let attribute = attribute.map_err(|error| owner_invalid(error.to_string()))?;
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Unbound) {
            continue;
        }
        let value = bounded_name(
            decode_attribute_bounded(
                attribute.value.as_ref(),
                decoder,
                limits,
                "DrawingML identity",
            )?,
            limits,
            "DrawingML identity",
        )?;
        match local.as_ref() {
            b"id" => {
                id = Some(
                    value
                        .parse::<u64>()
                        .map_err(|_| owner_invalid("DrawingML cNvPr id is not decimal"))?,
                )
            },
            b"name" => name = Some(bounded_display_name(value, limits, "DrawingML shape name")?),
            _ => {},
        }
    }
    let id = id.ok_or_else(|| owner_invalid("DrawingML cNvPr id is missing"))?;
    if !(1..=MAX_SHAPE_ID).contains(&id) {
        return Err(owner_invalid(
            "DrawingML cNvPr id is outside the Office range",
        ));
    }
    Ok(DrawingShape {
        id,
        name,
        compat_spid: String::new(),
    })
}

#[derive(Clone, Debug)]
struct VmlShape {
    id: String,
    spid: Option<String>,
    object_type: String,
    client_fields: Vec<VmlField>,
}

#[derive(Clone, Debug)]
struct VmlField {
    name: String,
    value: String,
    empty: bool,
}

struct VmlIdentityIndex {
    by_numeric: HashMap<u64, Option<usize>>,
    mismatched: HashSet<u64>,
}

fn index_vml_shapes(
    shapes: &[VmlShape],
    execution: Option<&OwnerExecution>,
) -> OwnerResult<VmlIdentityIndex> {
    let mut by_numeric = HashMap::new();
    by_numeric
        .try_reserve(shapes.len())
        .map_err(|source| owner_alloc("VML identity index", source))?;
    let mut mismatched = HashSet::new();
    mismatched
        .try_reserve(
            shapes
                .len()
                .checked_mul(2)
                .ok_or_else(|| owner_invalid("VML identity mismatch index size overflow"))?,
        )
        .map_err(|source| owner_alloc("VML identity mismatch index", source))?;
    for (position, shape) in shapes.iter().enumerate() {
        if let Some(execution) = execution {
            execution.check()?;
            execution.work(1)?;
        }
        let id_number = numeric_spid(&shape.id);
        let spid_number = shape.spid.as_deref().and_then(numeric_spid);
        if let (Some(id_number), Some(spid_number)) = (id_number, spid_number) {
            if id_number != spid_number {
                mismatched.insert(id_number);
                mismatched.insert(spid_number);
                continue;
            }
        }
        let Some(canonical) = id_number.or(spid_number) else {
            continue;
        };
        match by_numeric.entry(canonical) {
            std::collections::hash_map::Entry::Vacant(entry) => {
                entry.insert(Some(position));
            },
            std::collections::hash_map::Entry::Occupied(mut entry) => {
                entry.insert(None);
            },
        }
    }
    Ok(VmlIdentityIndex {
        by_numeric,
        mismatched,
    })
}

fn validate_unique_vml_attributes(
    element: &BytesStart<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
    limits: &OwnerLimits,
) -> OwnerResult<()> {
    let mut seen = HashSet::<Vec<u8>>::new();
    seen.try_reserve(8.min(super::MAX_ATTRIBUTES))
        .map_err(|source| owner_alloc("VML attribute identity index", source))?;
    let mut count = 0usize;
    for attribute in element.unchecked_attributes() {
        let attribute = attribute.map_err(|error| owner_invalid(error.to_string()))?;
        count = count
            .checked_add(1)
            .ok_or_else(|| owner_invalid("VML attribute count overflow"))?;
        if count > super::MAX_ATTRIBUTES {
            return Err(owner_limit(
                "VML attribute count",
                count,
                super::MAX_ATTRIBUTES,
            ));
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        let namespace = namespace_bytes(namespace);
        let key_len = namespace
            .len()
            .checked_add(1)
            .and_then(|length| length.checked_add(local.as_ref().len()))
            .ok_or_else(|| owner_invalid("VML attribute identity length overflow"))?;
        if key_len > limits.max_name_bytes {
            return Err(owner_limit(
                "VML attribute identity",
                key_len,
                limits.max_name_bytes,
            ));
        }
        let mut key = Vec::new();
        key.try_reserve_exact(key_len)
            .map_err(|source| owner_alloc("VML attribute identity", source))?;
        key.extend_from_slice(&namespace);
        key.push(0);
        key.extend_from_slice(local.as_ref());
        if !seen.insert(key) {
            return Err(owner_invalid("VML element contains duplicate attributes"));
        }
    }
    Ok(())
}

fn parse_vml(
    xml: &[u8],
    limits: &OwnerLimits,
    execution: Option<&OwnerExecution>,
) -> OwnerResult<Vec<VmlShape>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut stack: Vec<(Vec<u8>, Vec<u8>)> = Vec::new();
    let mut shapes = Vec::new();
    shapes
        .try_reserve(8.min(limits.max_shapes))
        .map_err(|source| owner_alloc("VML shape identities", source))?;
    let mut current: Option<(VmlShape, usize)> = None;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut events = 0usize;
    loop {
        if let Some(execution) = execution {
            execution.check()?;
            execution.work(1)?;
        }
        events = events
            .checked_add(1)
            .ok_or_else(|| owner_invalid("VML XML event count overflow"))?;
        if events > limits.max_mce_events {
            return Err(owner_limit("VML XML events", events, limits.max_mce_events));
        }
        let event = reader
            .read_event()
            .map_err(|error| owner_invalid(error.to_string()))?;
        let resolver = reader.resolver().clone();
        match event {
            Event::Start(element) => {
                validate_unique_vml_attributes(&element, &resolver, limits)?;
                let (namespace, local) = resolver.resolve_element(element.name());
                let namespace = namespace_bytes(namespace);
                let local = local.as_ref().to_vec();
                if !root_seen {
                    if !namespace.is_empty() || local.as_slice() != b"xml" {
                        return Err(owner_invalid(
                            "VML root must be the unqualified xml element",
                        ));
                    }
                    root_seen = true;
                } else if stack.is_empty() {
                    return Err(owner_invalid("VML XML has more than one root element"));
                } else if root_closed {
                    return Err(owner_invalid("VML XML has content after its root"));
                }
                if namespace.as_slice() == VML && local.as_slice() == b"shape" {
                    if !(stack.len() == 1
                        && stack[0].0.is_empty()
                        && stack[0].1.as_slice() == b"xml")
                    {
                        return Err(owner_invalid("VML shape is not a direct xml child"));
                    }
                    if current.is_some() {
                        return Err(owner_invalid("VML control shapes are nested"));
                    }
                    if shapes.len() >= limits.max_shapes {
                        return Err(owner_limit(
                            "VML shape identities",
                            shapes.len() + 1,
                            limits.max_shapes,
                        ));
                    }
                    let id = attr_unqualified(
                        &element,
                        b"id",
                        &resolver,
                        reader.decoder(),
                        limits,
                        "VML identity",
                    )?
                    .ok_or_else(|| owner_invalid("VML control shape id is missing"))?;
                    let spid = attr_qualified(
                        &element,
                        OFFICE,
                        b"spid",
                        &resolver,
                        reader.decoder(),
                        limits,
                        "VML o:spid",
                    )?;
                    if spid
                        .as_deref()
                        .is_some_and(|value| numeric_spid(value).is_none())
                    {
                        return Err(owner_invalid("VML o:spid is not a valid numeric identity"));
                    }
                    current = Some((
                        VmlShape {
                            id,
                            spid,
                            object_type: String::new(),
                            client_fields: Vec::new(),
                        },
                        stack.len(),
                    ));
                } else if namespace.as_slice() == EXCEL && local.as_slice() == b"ClientData" {
                    if let Some((shape, _)) = current.as_mut()
                        && stack_ends_with(&stack, &[(VML, b"shape")])
                    {
                        shape.object_type = attr_unqualified(
                            &element,
                            b"ObjectType",
                            &resolver,
                            reader.decoder(),
                            limits,
                            "VML ClientData ObjectType",
                        )?
                        .unwrap_or_default();
                    }
                }
                if stack.len() >= limits.max_mce_depth {
                    return Err(owner_limit(
                        "VML XML depth",
                        stack.len() + 1,
                        limits.max_mce_depth,
                    ));
                }
                stack
                    .try_reserve(1)
                    .map_err(|source| owner_alloc("VML XML stack", source))?;
                stack.push((namespace, local));
            },
            Event::Empty(element) => {
                validate_unique_vml_attributes(&element, &resolver, limits)?;
                let (namespace, local) = resolver.resolve_element(element.name());
                let namespace = namespace_bytes(namespace);
                let local = local.as_ref();
                if !root_seen || stack.is_empty() {
                    return Err(owner_invalid(
                        "VML empty element is outside the root element",
                    ));
                }
                if namespace.as_slice() == VML && local == b"shape" {
                    if !(stack.len() == 1
                        && stack[0].0.is_empty()
                        && stack[0].1.as_slice() == b"xml")
                    {
                        return Err(owner_invalid("VML shape is not a direct xml child"));
                    }
                    let id = attr_unqualified(
                        &element,
                        b"id",
                        &resolver,
                        reader.decoder(),
                        limits,
                        "VML identity",
                    )?
                    .ok_or_else(|| owner_invalid("VML control shape id is missing"))?;
                    let spid = attr_qualified(
                        &element,
                        OFFICE,
                        b"spid",
                        &resolver,
                        reader.decoder(),
                        limits,
                        "VML o:spid",
                    )?;
                    if spid
                        .as_deref()
                        .is_some_and(|value| numeric_spid(value).is_none())
                    {
                        return Err(owner_invalid("VML o:spid is not a valid numeric identity"));
                    }
                    if shapes.len() >= limits.max_shapes {
                        return Err(owner_limit(
                            "VML shape identities",
                            shapes.len() + 1,
                            limits.max_shapes,
                        ));
                    }
                    shapes
                        .try_reserve(1)
                        .map_err(|source| owner_alloc("VML shape identities", source))?;
                    shapes.push(VmlShape {
                        id,
                        spid,
                        object_type: String::new(),
                        client_fields: Vec::new(),
                    });
                } else if namespace.as_slice() == EXCEL && local == b"ClientData" {
                    if let Some((shape, _)) = current.as_mut()
                        && stack_ends_with(&stack, &[(VML, b"shape")])
                    {
                        shape.object_type = attr_unqualified(
                            &element,
                            b"ObjectType",
                            &resolver,
                            reader.decoder(),
                            limits,
                            "VML ClientData ObjectType",
                        )?
                        .unwrap_or_default();
                    }
                }
            },
            Event::End(end) => {
                let closed = pop_checked_end(&mut stack, &end, &resolver)?;
                if let Some((_, depth)) = current.as_ref() {
                    if closed.0.as_slice() == VML
                        && closed.1.as_slice() == b"shape"
                        && *depth == stack.len()
                    {
                        if let Some((shape, _)) = current.take() {
                            if shapes.len() >= limits.max_shapes {
                                return Err(owner_limit(
                                    "VML shape identities",
                                    shapes.len() + 1,
                                    limits.max_shapes,
                                ));
                            }
                            shapes
                                .try_reserve(1)
                                .map_err(|source| owner_alloc("VML shape identities", source))?;
                            shapes.push(shape);
                        }
                    }
                }
                if closed.0.is_empty() && closed.1.as_slice() == b"xml" {
                    root_closed = true;
                }
            },
            Event::Eof => break,
            Event::Text(text)
                if (!root_seen || root_closed)
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(owner_invalid(
                    "VML XML has non-whitespace content outside its root",
                ));
            },
            Event::CData(text)
                if (!root_seen || root_closed)
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(owner_invalid(
                    "VML XML has non-whitespace content outside its root",
                ));
            },
            Event::Decl(_)
            | Event::Comment(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
        if shapes.len() > limits.max_shapes {
            return Err(owner_limit(
                "VML shape identities",
                shapes.len(),
                limits.max_shapes,
            ));
        }
    }
    if !root_seen || !root_closed {
        return Err(owner_invalid("VML XML has no root element"));
    }
    if !stack.is_empty() || current.is_some() {
        return Err(owner_invalid("VML XML is unterminated"));
    }
    let fields = parse_vml_client_fields(xml, limits, execution)?;
    if fields.len() != shapes.len() {
        return Err(owner_invalid(
            "VML ClientData shape inventory does not match VML shape identities",
        ));
    }
    for (shape, fields) in shapes.iter_mut().zip(fields) {
        shape.client_fields = fields;
    }
    Ok(shapes)
}

fn parse_vml_client_fields(
    xml: &[u8],
    limits: &OwnerLimits,
    execution: Option<&OwnerExecution>,
) -> OwnerResult<Vec<Vec<VmlField>>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut stack: Vec<(Vec<u8>, Vec<u8>)> = Vec::new();
    let mut fields_by_shape: Vec<Vec<VmlField>> = Vec::new();
    let mut current_shape = None;
    let mut shape_depth = 0usize;
    let mut client_depth = None;
    let mut client_seen = false;
    let mut active_field: Option<(usize, usize, usize, Vec<u8>)> = None;
    let mut field_nodes = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut events = 0usize;
    loop {
        if let Some(execution) = execution {
            execution.check()?;
            execution.work(1)?;
        }
        events = events
            .checked_add(1)
            .ok_or_else(|| owner_invalid("VML ClientData event count overflow"))?;
        if events > limits.max_mce_events {
            return Err(owner_limit(
                "VML ClientData events",
                events,
                limits.max_mce_events,
            ));
        }
        let event = reader
            .read_event()
            .map_err(|error| owner_invalid(error.to_string()))?;
        let resolver = reader.resolver().clone();
        match event {
            Event::Start(element) => {
                let (namespace, local) = resolver.resolve_element(element.name());
                let namespace = namespace_bytes(namespace);
                let local = local.as_ref().to_vec();
                if stack.is_empty() {
                    if root_seen || !namespace.is_empty() || local.as_slice() != b"xml" {
                        return Err(owner_invalid(
                            "VML ClientData XML has an unexpected or duplicate root",
                        ));
                    }
                    root_seen = true;
                } else if root_closed {
                    return Err(owner_invalid(
                        "VML ClientData XML has content after its root",
                    ));
                }
                if namespace.as_slice() == VML && local.as_slice() == b"shape" {
                    if !(stack.len() == 1
                        && stack[0].0.is_empty()
                        && stack[0].1.as_slice() == b"xml")
                    {
                        return Err(owner_invalid(
                            "VML ClientData shape is not a direct xml child",
                        ));
                    }
                    if current_shape.is_some() {
                        return Err(owner_invalid("VML ClientData shapes are nested"));
                    }
                    if fields_by_shape.len() >= limits.max_shapes {
                        return Err(owner_limit(
                            "VML ClientData shapes",
                            fields_by_shape.len() + 1,
                            limits.max_shapes,
                        ));
                    }
                    fields_by_shape
                        .try_reserve(1)
                        .map_err(|source| owner_alloc("VML ClientData shapes", source))?;
                    fields_by_shape.push(Vec::new());
                    current_shape = Some(fields_by_shape.len() - 1);
                    shape_depth = stack.len();
                    client_depth = None;
                    client_seen = false;
                } else if namespace.as_slice() == EXCEL
                    && local.as_slice() == b"ClientData"
                    && current_shape.is_some()
                    && stack.last().is_some_and(|(parent_ns, parent)| {
                        parent_ns.as_slice() == VML && parent.as_slice() == b"shape"
                    })
                {
                    if client_seen {
                        return Err(owner_invalid(
                            "VML shape has duplicate Excel ClientData elements",
                        ));
                    }
                    client_seen = true;
                    client_depth = Some(stack.len());
                } else if namespace.as_slice() == EXCEL
                    && current_shape.is_some()
                    && client_depth.is_some_and(|depth| stack.len() == depth + 1)
                {
                    let shape = match current_shape {
                        Some(shape) => shape,
                        None => {
                            return Err(owner_invalid("VML ClientData field has no owning shape"));
                        },
                    };
                    field_nodes = field_nodes
                        .checked_add(1)
                        .ok_or_else(|| owner_invalid("VML ClientData field count overflow"))?;
                    if field_nodes > limits.max_mirror_nodes {
                        return Err(owner_limit(
                            "VML ClientData fields",
                            field_nodes,
                            limits.max_mirror_nodes,
                        ));
                    }
                    let fields = fields_by_shape
                        .get_mut(shape)
                        .ok_or_else(|| owner_invalid("VML ClientData shape index is invalid"))?;
                    fields
                        .try_reserve(1)
                        .map_err(|source| owner_alloc("VML ClientData fields", source))?;
                    fields.push(VmlField {
                        name: bounded_local_name(&local, limits, "VML ClientData field name")?,
                        value: String::new(),
                        empty: false,
                    });
                    active_field = Some((shape, fields.len() - 1, stack.len(), local.clone()));
                }
                if stack.len() >= limits.max_mce_depth {
                    return Err(owner_limit(
                        "VML ClientData depth",
                        stack.len() + 1,
                        limits.max_mce_depth,
                    ));
                }
                stack
                    .try_reserve(1)
                    .map_err(|source| owner_alloc("VML ClientData XML stack", source))?;
                stack.push((namespace, local));
            },
            Event::Empty(element) => {
                let (namespace, local) = resolver.resolve_element(element.name());
                let namespace = namespace_bytes(namespace);
                let local = local.as_ref().to_vec();
                if stack.is_empty() {
                    return Err(owner_invalid("VML ClientData empty root is not admitted"));
                }
                if root_closed {
                    return Err(owner_invalid(
                        "VML ClientData XML has content after its root",
                    ));
                }
                if namespace.as_slice() == VML && local.as_slice() == b"shape" {
                    if !(stack.len() == 1
                        && stack[0].0.is_empty()
                        && stack[0].1.as_slice() == b"xml")
                    {
                        return Err(owner_invalid(
                            "VML ClientData shape is not a direct xml child",
                        ));
                    }
                    if fields_by_shape.len() >= limits.max_shapes {
                        return Err(owner_limit(
                            "VML ClientData shapes",
                            fields_by_shape.len() + 1,
                            limits.max_shapes,
                        ));
                    }
                    fields_by_shape
                        .try_reserve(1)
                        .map_err(|source| owner_alloc("VML ClientData shapes", source))?;
                    fields_by_shape.push(Vec::new());
                } else if namespace.as_slice() == EXCEL
                    && local.as_slice() == b"ClientData"
                    && current_shape.is_some()
                    && stack.last().is_some_and(|(parent_ns, parent)| {
                        parent_ns.as_slice() == VML && parent.as_slice() == b"shape"
                    })
                {
                    if client_seen {
                        return Err(owner_invalid(
                            "VML shape has duplicate Excel ClientData elements",
                        ));
                    }
                    client_seen = true;
                } else if current_shape.is_some()
                    && namespace.as_slice() == EXCEL
                    && client_depth.is_some_and(|depth| stack.len() == depth + 1)
                {
                    let shape = current_shape
                        .ok_or_else(|| owner_invalid("VML ClientData field has no owning shape"))?;
                    field_nodes = field_nodes
                        .checked_add(1)
                        .ok_or_else(|| owner_invalid("VML ClientData field count overflow"))?;
                    if field_nodes > limits.max_mirror_nodes {
                        return Err(owner_limit(
                            "VML ClientData fields",
                            field_nodes,
                            limits.max_mirror_nodes,
                        ));
                    }
                    let fields = fields_by_shape
                        .get_mut(shape)
                        .ok_or_else(|| owner_invalid("VML ClientData shape index is invalid"))?;
                    fields
                        .try_reserve(1)
                        .map_err(|source| owner_alloc("VML ClientData fields", source))?;
                    fields.push(VmlField {
                        name: bounded_local_name(&local, limits, "VML ClientData field name")?,
                        value: String::new(),
                        empty: true,
                    });
                }
            },
            Event::Text(text) => {
                if (!root_seen || root_closed) && !text.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(owner_invalid(
                        "VML ClientData XML has non-whitespace content outside its root",
                    ));
                }
                append_vml_field_text(
                    &mut active_field,
                    &mut fields_by_shape,
                    text.as_ref(),
                    reader.decoder(),
                    limits.max_name_bytes,
                )?;
            },
            Event::CData(text) => {
                if (!root_seen || root_closed) && !text.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(owner_invalid(
                        "VML ClientData XML has non-whitespace content outside its root",
                    ));
                }
                append_vml_field_text(
                    &mut active_field,
                    &mut fields_by_shape,
                    text.as_ref(),
                    reader.decoder(),
                    limits.max_name_bytes,
                )?;
            },
            Event::End(end) => {
                let closed = pop_checked_end(&mut stack, &end, &resolver)?;
                if let Some((_, _, depth, name)) = active_field.as_ref() {
                    if *depth == stack.len() && closed.1 == *name {
                        active_field = None;
                    }
                }
                if closed.0.as_slice() == EXCEL
                    && closed.1.as_slice() == b"ClientData"
                    && client_depth == Some(stack.len())
                {
                    client_depth = None;
                }
                if closed.0.as_slice() == VML
                    && closed.1.as_slice() == b"shape"
                    && shape_depth == stack.len()
                {
                    current_shape = None;
                    client_depth = None;
                    client_seen = false;
                    active_field = None;
                }
                if closed.0.is_empty() && closed.1.as_slice() == b"xml" {
                    root_closed = true;
                }
            },
            Event::Eof => break,
            Event::Decl(_)
            | Event::Comment(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if !root_seen
        || !root_closed
        || !stack.is_empty()
        || current_shape.is_some()
        || active_field.is_some()
    {
        return Err(owner_invalid("VML ClientData XML is unterminated"));
    }
    Ok(fields_by_shape)
}

fn append_vml_field_text(
    active_field: &mut Option<(usize, usize, usize, Vec<u8>)>,
    fields_by_shape: &mut [Vec<VmlField>],
    bytes: &[u8],
    decoder: quick_xml::encoding::Decoder,
    maximum: usize,
) -> OwnerResult<()> {
    let Some((shape, field, _, _)) = active_field.as_ref() else {
        return Ok(());
    };
    let decoded = decoder
        .decode(bytes)
        .map_err(|error| owner_invalid(error.to_string()))?;
    let decoded =
        quick_xml::escape::unescape(&decoded).map_err(|error| owner_invalid(error.to_string()))?;
    let target = fields_by_shape
        .get_mut(*shape)
        .and_then(|fields| fields.get_mut(*field))
        .ok_or_else(|| owner_invalid("VML ClientData field index is invalid"))?;
    let total = target
        .value
        .len()
        .checked_add(decoded.len())
        .ok_or_else(|| owner_invalid("VML ClientData field value length overflow"))?;
    if total > maximum {
        return Err(owner_limit("VML ClientData field value", total, maximum));
    }
    target
        .value
        .try_reserve(decoded.len())
        .map_err(|source| owner_alloc("VML ClientData field value", source))?;
    target.value.push_str(&decoded);
    Ok(())
}

struct DiagnosticBuffer<'a> {
    diagnostics: Vec<FormControlDiagnostic>,
    execution: &'a mut OwnerExecution,
    limits: &'a OwnerLimits,
}

impl<'a> DiagnosticBuffer<'a> {
    fn new(execution: &'a mut OwnerExecution, limits: &'a OwnerLimits) -> Self {
        Self {
            diagnostics: Vec::new(),
            execution,
            limits,
        }
    }

    fn push_fmt(
        &mut self,
        code: FormControlDiagnosticCode,
        detail: fmt::Arguments<'_>,
    ) -> OwnerResult<()> {
        if self.diagnostics.len() >= self.limits.max_mirror_nodes {
            return Err(owner_limit(
                "form-control diagnostics",
                self.diagnostics.len() + 1,
                self.limits.max_mirror_nodes,
            ));
        }
        let detail_bytes = formatted_detail_len(detail)?;
        let needs_capacity = self.diagnostics.len() == self.diagnostics.capacity();
        let requested_capacity_bytes = if needs_capacity {
            size_of::<FormControlDiagnostic>()
        } else {
            0
        };
        let requested_bytes = requested_capacity_bytes
            .checked_add(detail_bytes)
            .ok_or_else(|| owner_invalid("form-control diagnostic bytes overflow"))?;
        self.execution.reserve_memory_category(
            OwnerMemoryCategory::Projection,
            requested_bytes,
            self.limits.max_projection_bytes,
            "form-control projection bytes",
        )?;

        if needs_capacity {
            let previous_capacity = self.diagnostics.capacity();
            self.diagnostics
                .try_reserve_exact(1)
                .map_err(|source| owner_alloc("form-control diagnostics", source))?;
            let actual_capacity = self.diagnostics.capacity();
            let expected_capacity = previous_capacity
                .checked_add(1)
                .ok_or_else(|| owner_invalid("form-control diagnostic capacity overflow"))?;
            if actual_capacity > expected_capacity {
                self.execution.reserve_memory_category(
                    OwnerMemoryCategory::Projection,
                    (actual_capacity - expected_capacity)
                        .checked_mul(size_of::<FormControlDiagnostic>())
                        .ok_or_else(|| {
                            owner_invalid("form-control diagnostic capacity overflow")
                        })?,
                    self.limits.max_projection_bytes,
                    "form-control projection bytes",
                )?;
            }
        }

        let mut detail_string = String::new();
        detail_string
            .try_reserve_exact(detail_bytes)
            .map_err(|source| owner_alloc("form-control diagnostic detail", source))?;
        let actual_detail_capacity = detail_string.capacity();
        if actual_detail_capacity > detail_bytes {
            self.execution.reserve_memory_category(
                OwnerMemoryCategory::Projection,
                actual_detail_capacity - detail_bytes,
                self.limits.max_projection_bytes,
                "form-control projection bytes",
            )?;
        }
        detail_string
            .write_fmt(detail)
            .map_err(|_| owner_invalid("form-control diagnostic formatting failed"))?;
        if detail_string.len() != detail_bytes {
            return Err(owner_invalid(
                "form-control diagnostic detail length changed during formatting",
            ));
        }
        self.diagnostics.push(FormControlDiagnostic::new(
            code,
            detail_string.into_boxed_str(),
        ));
        Ok(())
    }

    fn into_vec(self) -> Vec<FormControlDiagnostic> {
        self.diagnostics
    }
}

struct DetailLengthWriter {
    length: usize,
}

impl Write for DetailLengthWriter {
    fn write_str(&mut self, value: &str) -> fmt::Result {
        self.length = self.length.checked_add(value.len()).ok_or(fmt::Error)?;
        Ok(())
    }
}

fn formatted_detail_len(detail: fmt::Arguments<'_>) -> OwnerResult<usize> {
    let mut writer = DetailLengthWriter { length: 0 };
    writer
        .write_fmt(detail)
        .map_err(|_| owner_invalid("form-control diagnostic detail length overflow"))?;
    Ok(writer.length)
}

macro_rules! push_diagnostic {
    ($diagnostics:expr, $code:expr, $($arg:tt)*) => {{
        $diagnostics.push_fmt($code, format_args!($($arg)*))?;
    }};
}

fn vml_mirror_diagnostics(
    properties: &Properties,
    shape: &VmlShape,
    limits: &OwnerLimits,
    diagnostics: &mut DiagnosticBuffer<'_>,
) -> OwnerResult<()> {
    if let Some(value) = properties.checked() {
        let expected = match value {
            super::KnownOrUnknown::Known(super::Checked::Unchecked) => Some(0),
            super::KnownOrUnknown::Known(super::Checked::Checked) => Some(1),
            super::KnownOrUnknown::Known(super::Checked::Mixed) => Some(2),
            super::KnownOrUnknown::Unknown(_) => None,
        };
        check_vml_number(diagnostics, shape, "Checked", expected, "checked", limits)?;
    }
    check_vml_bool(
        diagnostics,
        shape,
        "Colored",
        properties.colored(),
        "colored",
        limits,
    )?;
    check_vml_number(
        diagnostics,
        shape,
        "DropLines",
        properties.drop_lines(),
        "dropLines",
        limits,
    )?;
    check_vml_token(
        diagnostics,
        shape,
        "DropStyle",
        properties.drop_style().and_then(vml_drop_style),
        "dropStyle",
        limits,
    )?;
    check_vml_number(diagnostics, shape, "Dx", properties.dx(), "dx", limits)?;
    check_vml_bool(
        diagnostics,
        shape,
        "FirstButton",
        properties.first_button(),
        "firstButton",
        limits,
    )?;
    check_vml_formula(
        diagnostics,
        shape,
        "FmlaGroup",
        properties.fmla_group(),
        "fmlaGroup",
        limits,
    )?;
    check_vml_formula(
        diagnostics,
        shape,
        "FmlaLink",
        properties.fmla_link(),
        "fmlaLink",
        limits,
    )?;
    check_vml_formula(
        diagnostics,
        shape,
        "FmlaRange",
        properties.fmla_range(),
        "fmlaRange",
        limits,
    )?;
    check_vml_formula(
        diagnostics,
        shape,
        "FmlaTxbx",
        properties.fmla_txbx(),
        "fmlaTxbx",
        limits,
    )?;
    check_vml_effective_bool(
        diagnostics,
        shape,
        "Horiz",
        properties.effective_horiz(),
        "horiz",
        limits,
    )?;
    check_vml_number(diagnostics, shape, "Inc", properties.inc(), "inc", limits)?;
    check_vml_bool(
        diagnostics,
        shape,
        "JustLastX",
        properties.just_last_x(),
        "justLastX",
        limits,
    )?;
    check_vml_lock_text(
        diagnostics,
        shape,
        "LockText",
        properties.effective_lock_text(),
        "lockText",
        limits,
    )?;
    check_vml_number(diagnostics, shape, "Max", properties.max(), "max", limits)?;
    check_vml_effective_number(
        diagnostics,
        shape,
        "Min",
        properties.effective_min(),
        "min",
        limits,
    )?;
    check_vml_string(
        diagnostics,
        shape,
        "MultiSel",
        properties.multi_sel(),
        "multiSel",
        limits,
    )?;
    check_vml_bool(
        diagnostics,
        shape,
        "NoThreeD",
        properties.no_three_d(),
        "noThreeD",
        limits,
    )?;
    check_vml_bool(
        diagnostics,
        shape,
        "NoThreeD2",
        properties.no_three_d2(),
        "noThreeD2",
        limits,
    )?;
    check_vml_number(
        diagnostics,
        shape,
        "Page",
        properties.page(),
        "page",
        limits,
    )?;
    check_vml_number(diagnostics, shape, "Sel", properties.sel(), "sel", limits)?;
    check_vml_effective_token(
        diagnostics,
        shape,
        "SelType",
        vml_selection_type(&properties.effective_seltype()),
        "seltype",
        limits,
    )?;
    check_vml_effective_token(
        diagnostics,
        shape,
        "TextHAlign",
        vml_horizontal_alignment(&properties.effective_text_h_align()),
        "textHAlign",
        limits,
    )?;
    check_vml_effective_token(
        diagnostics,
        shape,
        "TextVAlign",
        vml_vertical_alignment(&properties.effective_text_v_align()),
        "textVAlign",
        limits,
    )?;
    check_vml_effective_number(
        diagnostics,
        shape,
        "Val",
        properties.effective_val(),
        "val",
        limits,
    )?;
    check_vml_number(
        diagnostics,
        shape,
        "WidthMin",
        properties.width_min(),
        "widthMin",
        limits,
    )?;
    check_vml_edit_val(diagnostics, shape, properties, limits)?;
    check_vml_bool(
        diagnostics,
        shape,
        "MultiLine",
        properties.multi_line(),
        "multiLine",
        limits,
    )?;
    check_vml_effective_bool(
        diagnostics,
        shape,
        "VScroll",
        properties.effective_vertical_bar(),
        "verticalBar",
        limits,
    )?;
    check_vml_bool(
        diagnostics,
        shape,
        "SecretEdit",
        properties.password_edit(),
        "passwordEdit",
        limits,
    )?;
    check_vml_items(
        diagnostics,
        shape,
        properties.item_list(),
        properties.fmla_range(),
        limits,
    )?;
    Ok(())
}

fn vml_field<'a>(shape: &'a VmlShape, name: &str) -> (Option<&'a VmlField>, bool) {
    let mut found = None;
    let mut duplicate = false;
    for field in &shape.client_fields {
        if field.name == name {
            if found.is_some() {
                duplicate = true;
            } else {
                found = Some(field);
            }
        }
    }
    (found, duplicate)
}

fn check_vml_lexical(
    diagnostics: &mut DiagnosticBuffer<'_>,
    shape: &VmlShape,
    field_name: &str,
    expected: Option<&str>,
    read_only: bool,
    property: &str,
    _limits: &OwnerLimits,
) -> OwnerResult<()> {
    let (field, duplicate) = vml_field(shape, field_name);
    if duplicate {
        push_diagnostic!(
            diagnostics,
            FormControlDiagnosticCode::UnprovenMirror,
            "VML mirror {field_name} occurs more than once for {property}"
        );
    } else if read_only {
        push_diagnostic!(
            diagnostics,
            FormControlDiagnosticCode::UnprovenMirror,
            "{property}↔VML {field_name} mapping is read-only in this profile"
        );
    } else {
        match (field, expected) {
            (None, Some(_)) => push_diagnostic!(
                diagnostics,
                FormControlDiagnosticCode::UnprovenMirror,
                "VML mirror {field_name} is absent for {property}"
            ),
            (Some(field), Some(expected)) if field.value != expected => push_diagnostic!(
                diagnostics,
                FormControlDiagnosticCode::UnprovenMirror,
                "VML mirror {field_name} disagrees with x14 {property}"
            ),
            _ => {},
        }
    }
    Ok(())
}

fn check_vml_string(
    diagnostics: &mut DiagnosticBuffer<'_>,
    shape: &VmlShape,
    field_name: &str,
    expected: Option<&str>,
    property: &str,
    limits: &OwnerLimits,
) -> OwnerResult<()> {
    check_vml_lexical(
        diagnostics,
        shape,
        field_name,
        expected,
        false,
        property,
        limits,
    )
}

fn check_vml_number(
    diagnostics: &mut DiagnosticBuffer<'_>,
    shape: &VmlShape,
    field_name: &str,
    expected: Option<u32>,
    property: &str,
    _limits: &OwnerLimits,
) -> OwnerResult<()> {
    let (field, duplicate) = vml_field(shape, field_name);
    if duplicate {
        push_diagnostic!(
            diagnostics,
            FormControlDiagnosticCode::UnprovenMirror,
            "VML mirror {field_name} occurs more than once for {property}"
        );
    } else {
        match (field, expected) {
            // This field has no proven VML omission default.  Its authored
            // value is compared only when the mirror is also present.
            (None, _) => {},
            (Some(field), None) if vml_unsigned(field).is_none() => push_diagnostic!(
                diagnostics,
                FormControlDiagnosticCode::UnprovenMirror,
                "VML mirror {field_name} has an unknown numeric token for {property}"
            ),
            (Some(field), Some(expected)) => match vml_unsigned(field) {
                Some(actual) if actual == expected => {},
                Some(_) => push_diagnostic!(
                    diagnostics,
                    FormControlDiagnosticCode::UnprovenMirror,
                    "VML mirror {field_name} disagrees with x14 {property}"
                ),
                None => push_diagnostic!(
                    diagnostics,
                    FormControlDiagnosticCode::UnprovenMirror,
                    "VML mirror {field_name} has an unknown numeric token for {property}"
                ),
            },
            _ => {},
        }
    }
    Ok(())
}

fn check_vml_bool(
    diagnostics: &mut DiagnosticBuffer<'_>,
    shape: &VmlShape,
    field_name: &str,
    expected: Option<bool>,
    property: &str,
    _limits: &OwnerLimits,
) -> OwnerResult<()> {
    let (field, duplicate) = vml_field(shape, field_name);
    if duplicate {
        push_diagnostic!(
            diagnostics,
            FormControlDiagnosticCode::UnprovenMirror,
            "VML mirror {field_name} occurs more than once for {property}"
        );
    } else {
        match (field, expected) {
            // Only compare authored booleans when the VML mirror is present;
            // the VML omission default is not proven for this field.
            (None, _) => {},
            (Some(field), None) if vml_boolean(field).is_none() => push_diagnostic!(
                diagnostics,
                FormControlDiagnosticCode::UnprovenMirror,
                "VML mirror {field_name} has an unknown boolean token for {property}"
            ),
            (Some(field), Some(expected)) => match vml_boolean(field) {
                Some(actual) if actual == expected => {},
                Some(_) => push_diagnostic!(
                    diagnostics,
                    FormControlDiagnosticCode::UnprovenMirror,
                    "VML mirror {field_name} disagrees with x14 {property}"
                ),
                None => push_diagnostic!(
                    diagnostics,
                    FormControlDiagnosticCode::UnprovenMirror,
                    "VML mirror {field_name} has an unknown boolean token for {property}"
                ),
            },
            _ => {},
        }
    }
    Ok(())
}

fn check_vml_effective_bool(
    diagnostics: &mut DiagnosticBuffer<'_>,
    shape: &VmlShape,
    field_name: &str,
    expected: bool,
    property: &str,
    _limits: &OwnerLimits,
) -> OwnerResult<()> {
    let (field, duplicate) = vml_field(shape, field_name);
    if duplicate {
        push_diagnostic!(
            diagnostics,
            FormControlDiagnosticCode::UnprovenMirror,
            "VML mirror {field_name} occurs more than once for {property}"
        );
    } else {
        match field {
            // The VML omission is the proven false effective default for
            // horiz/verticalBar.
            None if !expected => {},
            None => push_diagnostic!(
                diagnostics,
                FormControlDiagnosticCode::UnprovenMirror,
                "VML mirror {field_name} disagrees with effective x14 {property}"
            ),
            Some(field) => match vml_boolean(field) {
                Some(actual) if actual == expected => {},
                Some(_) => push_diagnostic!(
                    diagnostics,
                    FormControlDiagnosticCode::UnprovenMirror,
                    "VML mirror {field_name} disagrees with effective x14 {property}"
                ),
                None => push_diagnostic!(
                    diagnostics,
                    FormControlDiagnosticCode::UnprovenMirror,
                    "VML mirror {field_name} has an unknown boolean token for {property}"
                ),
            },
        }
    }
    Ok(())
}

fn check_vml_lock_text(
    diagnostics: &mut DiagnosticBuffer<'_>,
    shape: &VmlShape,
    field_name: &str,
    expected: bool,
    property: &str,
    _limits: &OwnerLimits,
) -> OwnerResult<()> {
    let (field, duplicate) = vml_field(shape, field_name);
    if duplicate {
        push_diagnostic!(
            diagnostics,
            FormControlDiagnosticCode::UnprovenMirror,
            "VML mirror {field_name} occurs more than once for {property}"
        );
    } else {
        // Unlike the other effective mirrors, the VML default for lockText
        // is true.  Compare that effective value with the x14 effective value
        // while retaining authored presence separately in the leaf data.
        match field {
            None if expected => {},
            None => push_diagnostic!(
                diagnostics,
                FormControlDiagnosticCode::UnprovenMirror,
                "VML mirror {field_name} disagrees with effective x14 {property}"
            ),
            Some(field) => match vml_boolean(field) {
                Some(actual) if actual == expected => {},
                Some(_) => push_diagnostic!(
                    diagnostics,
                    FormControlDiagnosticCode::UnprovenMirror,
                    "VML mirror {field_name} disagrees with effective x14 {property}"
                ),
                None => push_diagnostic!(
                    diagnostics,
                    FormControlDiagnosticCode::UnprovenMirror,
                    "VML mirror {field_name} has an unknown boolean token for {property}"
                ),
            },
        }
    }
    Ok(())
}

fn check_vml_effective_number(
    diagnostics: &mut DiagnosticBuffer<'_>,
    shape: &VmlShape,
    field_name: &str,
    expected: u32,
    property: &str,
    _limits: &OwnerLimits,
) -> OwnerResult<()> {
    let (field, duplicate) = vml_field(shape, field_name);
    if duplicate {
        push_diagnostic!(
            diagnostics,
            FormControlDiagnosticCode::UnprovenMirror,
            "VML mirror {field_name} occurs more than once for {property}"
        );
    } else {
        match field {
            // The VML omission is the proven zero effective default for
            // min/val.
            None if expected == 0 => {},
            None => push_diagnostic!(
                diagnostics,
                FormControlDiagnosticCode::UnprovenMirror,
                "VML mirror {field_name} disagrees with effective x14 {property}"
            ),
            Some(field) => match vml_unsigned(field) {
                Some(actual) if actual == expected => {},
                Some(_) => push_diagnostic!(
                    diagnostics,
                    FormControlDiagnosticCode::UnprovenMirror,
                    "VML mirror {field_name} disagrees with effective x14 {property}"
                ),
                None => push_diagnostic!(
                    diagnostics,
                    FormControlDiagnosticCode::UnprovenMirror,
                    "VML mirror {field_name} has an unknown numeric token for {property}"
                ),
            },
        }
    }
    Ok(())
}

fn check_vml_formula(
    diagnostics: &mut DiagnosticBuffer<'_>,
    shape: &VmlShape,
    field_name: &str,
    expected: Option<&super::FormControlFormula>,
    property: &str,
    limits: &OwnerLimits,
) -> OwnerResult<()> {
    check_vml_lexical(
        diagnostics,
        shape,
        field_name,
        expected.map(super::FormControlFormula::as_str),
        false,
        property,
        limits,
    )
}

fn check_vml_token(
    diagnostics: &mut DiagnosticBuffer<'_>,
    shape: &VmlShape,
    field_name: &str,
    expected: Option<&str>,
    property: &str,
    _limits: &OwnerLimits,
) -> OwnerResult<()> {
    let (field, duplicate) = vml_field(shape, field_name);
    if duplicate {
        push_diagnostic!(
            diagnostics,
            FormControlDiagnosticCode::UnprovenMirror,
            "VML mirror {field_name} occurs more than once for {property}"
        );
    } else {
        match (field, expected) {
            // Token omission is only meaningful for the effective defaults
            // supplied by the caller.  Other token fields are authored-only.
            (None, _) => {},
            (Some(field), None) if !valid_vml_token(field_name, field.value.as_str()) => {
                push_diagnostic!(
                    diagnostics,
                    FormControlDiagnosticCode::UnprovenMirror,
                    "VML mirror {field_name} has an unknown token for {property}"
                )
            },
            (Some(field), Some(expected)) if field.value != expected => push_diagnostic!(
                diagnostics,
                FormControlDiagnosticCode::UnprovenMirror,
                "VML mirror {field_name} disagrees with x14 {property}"
            ),
            _ => {},
        }
    }
    Ok(())
}

fn check_vml_effective_token(
    diagnostics: &mut DiagnosticBuffer<'_>,
    shape: &VmlShape,
    field_name: &str,
    expected: Option<&str>,
    property: &str,
    _limits: &OwnerLimits,
) -> OwnerResult<()> {
    let (field, duplicate) = vml_field(shape, field_name);
    if duplicate {
        push_diagnostic!(
            diagnostics,
            FormControlDiagnosticCode::UnprovenMirror,
            "VML mirror {field_name} occurs more than once for {property}"
        );
    } else if let Some(field) = field {
        if !valid_vml_token(field_name, field.value.as_str()) {
            push_diagnostic!(
                diagnostics,
                FormControlDiagnosticCode::UnprovenMirror,
                "VML mirror {field_name} has an unknown token for {property}"
            );
        } else if let Some(expected) = expected {
            if field.value != expected {
                push_diagnostic!(
                    diagnostics,
                    FormControlDiagnosticCode::UnprovenMirror,
                    "VML mirror {field_name} disagrees with effective x14 {property}"
                );
            }
        }
    } else {
        let default = match field_name {
            "SelType" => Some("Single"),
            "TextHAlign" => Some("Left"),
            "TextVAlign" => Some("Top"),
            _ => None,
        };
        match (expected, default) {
            (Some(expected), Some(default)) if expected != default => push_diagnostic!(
                diagnostics,
                FormControlDiagnosticCode::UnprovenMirror,
                "VML mirror {field_name} disagrees with effective x14 {property}"
            ),
            _ => {},
        }
    }
    Ok(())
}

fn valid_vml_token(field_name: &str, value: &str) -> bool {
    match field_name {
        "DropStyle" => matches!(value, "Combo" | "ComboEdit" | "Simple"),
        "SelType" => matches!(value, "Single" | "Multi" | "Extend"),
        "TextHAlign" => matches!(
            value,
            "Left" | "Justify" | "Center" | "Right" | "Distributed"
        ),
        "TextVAlign" => matches!(
            value,
            "Top" | "Justify" | "Center" | "Bottom" | "Distributed"
        ),
        _ => true,
    }
}

fn check_vml_items(
    diagnostics: &mut DiagnosticBuffer<'_>,
    shape: &VmlShape,
    list: Option<&super::ItemList>,
    fmla_range: Option<&super::FormControlFormula>,
    _limits: &OwnerLimits,
) -> OwnerResult<()> {
    let Some(list) = list else {
        // No proven x14 itemLst value means that VML ListItem nodes remain
        // authored opaque data; retain them, but make the unresolved host
        // pairing visible to a read caller.
        if fmla_range.is_none()
            && shape
                .client_fields
                .iter()
                .any(|field| field.name == "ListItem")
        {
            push_diagnostic!(
                diagnostics,
                FormControlDiagnosticCode::UnprovenMirror,
                "VML ListItem sequence has no x14 itemLst to pair"
            );
        }
        return Ok(());
    };
    let mut count = 0usize;
    let mut mismatch = false;
    for field in shape
        .client_fields
        .iter()
        .filter(|field| field.name == "ListItem")
    {
        if list
            .items()
            .get(count)
            .is_none_or(|item| item.value() != field.value)
        {
            mismatch = true;
        }
        count = count
            .checked_add(1)
            .ok_or_else(|| owner_invalid("VML ListItem count overflow"))?;
    }
    if mismatch || count != list.items().len() {
        push_diagnostic!(
            diagnostics,
            FormControlDiagnosticCode::UnprovenMirror,
            "VML ListItem sequence disagrees with x14 itemLst"
        );
    }
    Ok(())
}

fn vml_unsigned(field: &VmlField) -> Option<u32> {
    if field.empty || field.value.is_empty() {
        return None;
    }
    if !field.value.bytes().all(|byte| byte.is_ascii_digit()) {
        return None;
    }
    field.value.parse().ok()
}

fn vml_boolean(field: &VmlField) -> Option<bool> {
    if field.empty || field.value.is_empty() {
        return Some(true);
    }
    match field.value.as_str() {
        "t" | "true" | "True" => Some(true),
        "f" | "false" | "False" => Some(false),
        _ => None,
    }
}

fn vml_drop_style(value: &super::KnownOrUnknown<super::DropStyle>) -> Option<&'static str> {
    match value {
        super::KnownOrUnknown::Known(super::DropStyle::Combo) => Some("Combo"),
        super::KnownOrUnknown::Known(super::DropStyle::ComboEdit) => Some("ComboEdit"),
        super::KnownOrUnknown::Known(super::DropStyle::Simple) => Some("Simple"),
        super::KnownOrUnknown::Unknown(_) => None,
    }
}

fn vml_selection_type(value: &super::KnownOrUnknown<super::SelectionType>) -> Option<&'static str> {
    match value {
        super::KnownOrUnknown::Known(super::SelectionType::Single) => Some("Single"),
        super::KnownOrUnknown::Known(super::SelectionType::Multi) => Some("Multi"),
        super::KnownOrUnknown::Known(super::SelectionType::Extended) => Some("Extend"),
        super::KnownOrUnknown::Unknown(_) => None,
    }
}

fn vml_horizontal_alignment(
    value: &super::KnownOrUnknown<super::TextHAlign>,
) -> Option<&'static str> {
    match value {
        super::KnownOrUnknown::Known(super::TextHAlign::Left) => Some("Left"),
        super::KnownOrUnknown::Known(super::TextHAlign::Justify) => Some("Justify"),
        super::KnownOrUnknown::Known(super::TextHAlign::Center) => Some("Center"),
        super::KnownOrUnknown::Known(super::TextHAlign::Right) => Some("Right"),
        super::KnownOrUnknown::Known(super::TextHAlign::Distributed) => Some("Distributed"),
        super::KnownOrUnknown::Unknown(_) => None,
    }
}

fn vml_vertical_alignment(
    value: &super::KnownOrUnknown<super::TextVAlign>,
) -> Option<&'static str> {
    match value {
        super::KnownOrUnknown::Known(super::TextVAlign::Top) => Some("Top"),
        super::KnownOrUnknown::Known(super::TextVAlign::Justify) => Some("Justify"),
        super::KnownOrUnknown::Known(super::TextVAlign::Center) => Some("Center"),
        super::KnownOrUnknown::Known(super::TextVAlign::Bottom) => Some("Bottom"),
        super::KnownOrUnknown::Known(super::TextVAlign::Distributed) => Some("Distributed"),
        super::KnownOrUnknown::Unknown(_) => None,
    }
}

fn check_vml_edit_val(
    diagnostics: &mut DiagnosticBuffer<'_>,
    shape: &VmlShape,
    properties: &Properties,
    _limits: &OwnerLimits,
) -> OwnerResult<()> {
    let expected = match properties.effective_edit_val() {
        super::KnownOrUnknown::Known(super::EditValidation::Text) => Some(0),
        super::KnownOrUnknown::Known(super::EditValidation::Integer) => Some(1),
        super::KnownOrUnknown::Known(super::EditValidation::Number) => Some(2),
        super::KnownOrUnknown::Known(super::EditValidation::Reference) => Some(3),
        super::KnownOrUnknown::Known(super::EditValidation::Formula) => Some(4),
        super::KnownOrUnknown::Unknown(_) => None,
    };
    let (field, duplicate) = vml_field(shape, "VTEdit");
    if duplicate {
        push_diagnostic!(
            diagnostics,
            FormControlDiagnosticCode::UnprovenMirror,
            "VML mirror VTEdit occurs more than once for editVal"
        );
    } else {
        match (field, expected) {
            // VTEdit omission is the proven text (0) default.
            (None, Some(0)) | (None, None) => {},
            (None, Some(_)) => push_diagnostic!(
                diagnostics,
                FormControlDiagnosticCode::UnprovenMirror,
                "VML mirror VTEdit is absent for editVal"
            ),
            (Some(_), None) => push_diagnostic!(
                diagnostics,
                FormControlDiagnosticCode::UnprovenMirror,
                "VML mirror VTEdit cannot be compared with an unknown x14 editVal token"
            ),
            (Some(field), Some(expected)) => match vml_unsigned(field) {
                Some(actual) if actual == expected => {},
                Some(_) => push_diagnostic!(
                    diagnostics,
                    FormControlDiagnosticCode::UnprovenMirror,
                    "VML mirror VTEdit disagrees with x14 editVal"
                ),
                None => push_diagnostic!(
                    diagnostics,
                    FormControlDiagnosticCode::UnprovenMirror,
                    "VML mirror VTEdit has an unknown numeric token"
                ),
            },
        }
    }
    Ok(())
}

fn match_vml<'a>(
    shapes: &'a [VmlShape],
    index: &VmlIdentityIndex,
    drawing: &DrawingShape,
    _expected_type: Option<&str>,
    control_name: Option<&str>,
    limits: &OwnerLimits,
) -> OwnerResult<&'a VmlShape> {
    if index.mismatched.contains(&drawing.id) {
        return Err(owner_invalid("VML id and o:spid disagree"));
    }
    let shape = match index.by_numeric.get(&drawing.id) {
        Some(Some(position)) => shapes
            .get(*position)
            .ok_or_else(|| owner_invalid("VML identity index is inconsistent"))?,
        Some(None) => {
            return Err(owner_invalid(
                "VML control shape identity is missing or ambiguous",
            ));
        },
        None => {
            return Err(owner_invalid(
                "VML control shape identity is missing or ambiguous",
            ));
        },
    };
    if numeric_spid(&shape.id).is_none() {
        let Some(control_name) = control_name else {
            return Err(owner_invalid("nonnumeric VML identity has no control name"));
        };
        if decode_object_name(&shape.id, limits)? != control_name {
            return Err(owner_invalid(
                "nonnumeric VML identity disagrees with control name",
            ));
        }
    }
    Ok(shape)
}

fn numeric_spid(value: &str) -> Option<u64> {
    let suffix = value.strip_prefix("_x0000_s")?;
    if suffix.is_empty() || !suffix.bytes().all(|byte| byte.is_ascii_digit()) {
        return None;
    }
    let number = suffix.parse().ok()?;
    (MIN_SP_ID..=MAX_SP_ID).contains(&number).then_some(number)
}

fn decode_object_name(value: &str, limits: &OwnerLimits) -> OwnerResult<String> {
    if value.len() > limits.max_name_bytes {
        return Err(owner_limit(
            "VML object identity",
            value.len(),
            limits.max_name_bytes,
        ));
    }
    let mut output = String::new();
    output
        .try_reserve(value.len().min(limits.max_name_bytes))
        .map_err(|source| owner_alloc("VML object identity", source))?;
    let mut characters = 0usize;
    let bytes = value.as_bytes();
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index..].starts_with(b"_x005F_x") {
            if index + 13 > bytes.len() {
                return Err(owner_invalid(
                    "VML object name has malformed protected escape",
                ));
            }
            let digits = &bytes[index + 8..index + 12];
            if !digits.iter().all(|byte| byte.is_ascii_hexdigit()) || bytes[index + 12] != b'_' {
                return Err(owner_invalid(
                    "VML object name has malformed protected escape",
                ));
            }
            let escaped = std::str::from_utf8(digits)
                .map_err(|_| owner_invalid("VML object name escape is not ASCII"))?;
            push_bounded_object_name(&mut output, &mut characters, "_x", limits)?;
            push_bounded_object_name(&mut output, &mut characters, escaped, limits)?;
            push_bounded_object_name(&mut output, &mut characters, "_", limits)?;
            index += 13;
            continue;
        }
        if bytes[index..].starts_with(b"_x") {
            if index + 7 > bytes.len() || bytes[index + 6] != b'_' {
                return Err(owner_invalid("VML object name has malformed escape"));
            }
            let digits = &bytes[index + 2..index + 6];
            if !digits.iter().all(|byte| byte.is_ascii_hexdigit()) {
                return Err(owner_invalid("VML object name has malformed escape"));
            }
            let value = u16::from_str_radix(
                std::str::from_utf8(digits)
                    .map_err(|_| owner_invalid("VML object name escape is not ASCII"))?,
                16,
            )
            .map_err(|_| owner_invalid("VML object name escape is invalid"))?;
            if (0xD800..=0xDFFF).contains(&value) {
                return Err(owner_invalid(
                    "VML object name escape is an unpaired surrogate",
                ));
            }
            let character = char::from_u32(value as u32)
                .ok_or_else(|| owner_invalid("VML object name escape is invalid Unicode"))?;
            let mut encoded = [0u8; 4];
            push_bounded_object_name(
                &mut output,
                &mut characters,
                character.encode_utf8(&mut encoded),
                limits,
            )?;
            index += 7;
            continue;
        }
        let character = value[index..]
            .chars()
            .next()
            .ok_or_else(|| owner_invalid("VML object name is not UTF-8"))?;
        let mut encoded = [0u8; 4];
        push_bounded_object_name(
            &mut output,
            &mut characters,
            character.encode_utf8(&mut encoded),
            limits,
        )?;
        index += character.len_utf8();
    }
    Ok(output)
}

fn push_bounded_object_name(
    output: &mut String,
    characters: &mut usize,
    value: &str,
    limits: &OwnerLimits,
) -> OwnerResult<()> {
    let bytes = output
        .len()
        .checked_add(value.len())
        .ok_or_else(|| owner_invalid("VML object identity length overflow"))?;
    if bytes > limits.max_name_bytes {
        return Err(owner_limit(
            "VML object identity",
            bytes,
            limits.max_name_bytes,
        ));
    }
    let count = value.chars().count();
    let total = characters
        .checked_add(count)
        .ok_or_else(|| owner_invalid("VML object identity character count overflow"))?;
    if total > MAX_AUTHORED_NAME_CHARS {
        return Err(owner_limit(
            "VML object identity characters",
            total,
            MAX_AUTHORED_NAME_CHARS,
        ));
    }
    output.push_str(value);
    *characters = total;
    Ok(())
}

fn attr_unqualified(
    element: &BytesStart<'_>,
    wanted: &[u8],
    resolver: &quick_xml::name::NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    limits: &OwnerLimits,
    resource: &'static str,
) -> OwnerResult<Option<String>> {
    for attribute in element.unchecked_attributes() {
        let attribute = attribute.map_err(|error| owner_invalid(error.to_string()))?;
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Unbound) && local.as_ref() == wanted {
            return Ok(Some(decode_attribute_bounded(
                attribute.value.as_ref(),
                decoder,
                limits,
                resource,
            )?));
        }
    }
    Ok(None)
}

fn attr_qualified(
    element: &BytesStart<'_>,
    wanted_ns: &[u8],
    wanted_local: &[u8],
    resolver: &quick_xml::name::NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    limits: &OwnerLimits,
    resource: &'static str,
) -> OwnerResult<Option<String>> {
    for attribute in element.unchecked_attributes() {
        let attribute = attribute.map_err(|error| owner_invalid(error.to_string()))?;
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if is_ns(&namespace, wanted_ns) && local.as_ref() == wanted_local {
            return Ok(Some(decode_attribute_bounded(
                attribute.value.as_ref(),
                decoder,
                limits,
                resource,
            )?));
        }
    }
    Ok(None)
}

fn pop_checked_end(
    stack: &mut Vec<(Vec<u8>, Vec<u8>)>,
    end: &BytesEnd<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
) -> OwnerResult<(Vec<u8>, Vec<u8>)> {
    let (namespace, local) = resolver.resolve_element(end.name());
    let actual = (namespace_bytes(namespace), local.as_ref().to_vec());
    let expected = stack
        .pop()
        .ok_or_else(|| owner_invalid("XML closing tag is unmatched"))?;
    if expected != actual {
        return Err(owner_invalid(format!(
            "XML closing QName does not match its opening QName (expected {}:{}, got {}:{})",
            String::from_utf8_lossy(&expected.0),
            String::from_utf8_lossy(&expected.1),
            String::from_utf8_lossy(&actual.0),
            String::from_utf8_lossy(&actual.1),
        )));
    }
    Ok(expected)
}

fn namespace_bytes(namespace: ResolveResult<'_>) -> Vec<u8> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            // quick-xml intentionally retains namespace declaration values in
            // their lexical form.  XML attribute references are nevertheless
            // expanded before namespace identity is compared, so canonicalize
            // the small set of namespaces admitted by this owner while
            // retaining unknown URIs byte-for-byte for diagnostics/QName
            // closure checks.
            for wanted in [
                SML,
                STRICT_SML,
                REL,
                STRICT_REL,
                XDR,
                STRICT_XDR,
                A14,
                VML,
                OFFICE,
                EXCEL,
                MCE,
                DRAWING,
                STRICT_DRAWING,
                MATH,
                STRICT_MATH,
                XML,
                XMLNS,
            ] {
                if namespace_uri_matches(value, wanted) {
                    return wanted.to_vec();
                }
            }
            value.as_ref().to_vec()
        },
        _ => Vec::new(),
    }
}

fn is_ns(namespace: &ResolveResult<'_>, wanted: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if namespace_uri_matches(value, wanted))
}

/// Compare a namespace declaration after XML attribute-value normalization.
///
/// `quick_xml::NamespaceResolver` deliberately exposes the lexical namespace
/// value, so `xmlns:x="http://example/&#x2F;name"` resolves to the bytes that
/// still contain `&#x2F;`.  Namespace identity is defined after expanding XML
/// character/entity references.  The admitted namespace URIs are ASCII; this
/// bounded comparator therefore decodes one scalar at a time without creating
/// a copy proportional to an attacker-controlled declaration.
fn namespace_uri_matches(raw: &[u8], wanted: &[u8]) -> bool {
    if raw == wanted {
        return true;
    }
    let mut raw_index = 0usize;
    let mut wanted_index = 0usize;
    while raw_index < raw.len() && wanted_index < wanted.len() {
        let value = if raw[raw_index] == b'&' {
            let entity_start = raw_index + 1;
            let Some(relative_end) = raw[entity_start..].iter().position(|byte| *byte == b';')
            else {
                return false;
            };
            let entity_end = entity_start + relative_end;
            let Some(scalar) = namespace_entity_scalar(&raw[entity_start..entity_end]) else {
                return false;
            };
            raw_index = entity_end + 1;
            scalar
        } else {
            let value = raw[raw_index] as u32;
            raw_index += 1;
            value
        };
        if value > u32::from(u8::MAX) || wanted[wanted_index] != value as u8 {
            return false;
        }
        wanted_index += 1;
    }
    raw_index == raw.len() && wanted_index == wanted.len()
}

fn namespace_entity_scalar(entity: &[u8]) -> Option<u32> {
    match entity {
        b"amp" => Some(u32::from(b'&')),
        b"lt" => Some(u32::from(b'<')),
        b"gt" => Some(u32::from(b'>')),
        b"apos" => Some(u32::from(b'\'')),
        b"quot" => Some(u32::from(b'"')),
        _ if entity.first() == Some(&b'#') => {
            let (radix, digits) = match entity.get(1) {
                Some(b'x' | b'X') => (16, &entity[2..]),
                Some(_) => (10, &entity[1..]),
                None => return None,
            };
            if digits.is_empty() {
                return None;
            }
            let mut value = 0u32;
            for digit in digits {
                let digit = match radix {
                    16 => match digit {
                        b'0'..=b'9' => u32::from(*digit - b'0'),
                        b'a'..=b'f' => u32::from(*digit - b'a' + 10),
                        b'A'..=b'F' => u32::from(*digit - b'A' + 10),
                        _ => return None,
                    },
                    _ => match digit {
                        b'0'..=b'9' => u32::from(*digit - b'0'),
                        _ => return None,
                    },
                };
                value = value.checked_mul(radix)?.checked_add(digit)?;
            }
            (value <= 0x10_FFFF && !(0xD800..=0xDFFF).contains(&value) && value != 0)
                .then_some(value)
        },
        _ => None,
    }
}

fn decode_attribute(value: &[u8], decoder: quick_xml::encoding::Decoder) -> OwnerResult<String> {
    let decoded = decoder
        .decode(value)
        .map_err(|error| owner_invalid(error.to_string()))?;
    quick_xml::escape::unescape(&decoded)
        .map(|value| value.into_owned())
        .map_err(|error| owner_invalid(error.to_string()))
}

fn decode_attribute_bounded(
    value: &[u8],
    decoder: quick_xml::encoding::Decoder,
    limits: &OwnerLimits,
    resource: &'static str,
) -> OwnerResult<String> {
    if value.len() > limits.max_name_bytes {
        return Err(owner_limit(resource, value.len(), limits.max_name_bytes));
    }
    let decoded = decode_attribute(value, decoder)?;
    if decoded.len() > limits.max_name_bytes {
        return Err(owner_limit(resource, decoded.len(), limits.max_name_bytes));
    }
    Ok(decoded)
}

fn bounded_name(
    value: String,
    limits: &OwnerLimits,
    resource: &'static str,
) -> OwnerResult<String> {
    if value.len() > limits.max_name_bytes {
        return Err(owner_limit(resource, value.len(), limits.max_name_bytes));
    }
    Ok(value)
}

fn bounded_local_name(
    value: &[u8],
    limits: &OwnerLimits,
    resource: &'static str,
) -> OwnerResult<String> {
    if value.len() > limits.max_name_bytes {
        return Err(owner_limit(resource, value.len(), limits.max_name_bytes));
    }
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(value.len())
        .map_err(|source| owner_alloc(resource, source))?;
    bytes.extend_from_slice(value);
    bounded_name(
        String::from_utf8(bytes)
            .map_err(|_| owner_invalid("VML ClientData field name is not UTF-8"))?,
        limits,
        resource,
    )
}

fn bounded_display_name(
    value: String,
    limits: &OwnerLimits,
    resource: &'static str,
) -> OwnerResult<String> {
    let value = bounded_name(value, limits, resource)?;
    if value.chars().count() > MAX_AUTHORED_NAME_CHARS {
        return Err(owner_limit(
            resource,
            value.chars().count(),
            MAX_AUTHORED_NAME_CHARS,
        ));
    }
    Ok(value)
}

pub(crate) fn leaf_limits(limits: &OwnerLimits) -> LeafLimits {
    let part_bytes = limits
        .max_mce_bytes
        .min(limits.max_semantic_bytes)
        .min(super::MAX_PART_BYTES);
    LeafLimits::new()
        .with_max_part_bytes(part_bytes)
        .with_max_output_bytes(part_bytes)
        .with_max_depth(limits.max_mce_depth)
        .with_max_events(limits.max_mce_events)
        .with_max_items(limits.max_mirror_nodes)
        .with_max_item_value_bytes(part_bytes)
        .with_max_attributes(super::MAX_ATTRIBUTES)
        .with_max_opaque_bytes(
            limits
                .max_mce_bytes
                .min(limits.max_semantic_bytes)
                .min(super::MAX_OPAQUE_BYTES),
        )
        .with_max_retained_bytes(
            limits
                .max_mce_bytes
                .min(limits.max_semantic_bytes)
                .min(super::MAX_RETAINED_BYTES),
        )
}

fn bounded_owned(source: &[u8], maximum: usize, resource: &'static str) -> OwnerResult<Vec<u8>> {
    if source.len() > maximum {
        return Err(owner_limit(resource, source.len(), maximum));
    }
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(source.len())
        .map_err(|source| owner_alloc(resource, source))?;
    owned.extend_from_slice(source);
    Ok(owned)
}

fn owner_invalid(message: impl Into<String>) -> FormControlOwnerError {
    FormControlOwnerError::Invalid(message.into())
}
fn owner_limit(resource: &'static str, observed: usize, maximum: usize) -> FormControlOwnerError {
    FormControlOwnerError::Limit {
        resource,
        observed,
        maximum,
    }
}
fn owner_alloc(
    resource: &'static str,
    source: std::collections::TryReserveError,
) -> FormControlOwnerError {
    FormControlOwnerError::Allocation { resource, source }
}

#[cfg(test)]
mod tests {
    use litchi_core::{
        Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits,
        Resource,
    };
    use litchi_opc::BlobPart;
    use std::num::{NonZeroU64, NonZeroUsize};

    use super::*;

    fn eager_inventory_package(count: usize) -> OpcPackage {
        let mut package = OpcPackage::new();
        for index in 0..count {
            let partname = PackURI::new(format!("/xl/ctrlProps{index}.xml"))
                .expect("test control-properties part name");
            package.add_part(Box::new(BlobPart::new(
                partname,
                CONTROL_PROPERTIES_CONTENT_TYPE.to_owned(),
                b"<formControlPr xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"/>"
                    .to_vec(),
            )));
        }
        package
    }

    fn managed_test_context(work: u64) -> (Budget, ExecutionContext) {
        let memory = 16 * 1024 * 1024;
        let budget = Budget::root(
            "xlsx-form-control-owner-test",
            CoreLimits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, work),
        );
        let (_, cancellation) = CancellationSource::pair();
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::new(1).expect("one test worker"),
            NonZeroUsize::new(1).expect("one in-flight task"),
            NonZeroU64::new(memory).expect("positive test byte ceiling"),
            0,
        )
        .expect("valid test execution policy");
        let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
        (budget, context)
    }

    #[test]
    fn nested_alternate_content_in_an_ignored_choice_stays_ignored() {
        let xml = br#"<root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="u"><mc:AlternateContent><mc:Choice Requires="u"><inner/></mc:Choice><mc:Fallback><nested-fallback/></mc:Fallback></mc:AlternateContent></mc:Choice><mc:Fallback><outer-fallback/></mc:Fallback></mc:AlternateContent></root>"#;
        let ranges = mce_branch_ranges(xml, FORM_CONTROL_NAMESPACE, &OwnerLimits::default(), None)
            .expect("nested MCE provenance should be bounded");
        assert!(ranges.selected.iter().all(|range| {
            !xml[range.clone()]
                .windows(b"inner".len())
                .any(|window| window == b"inner")
        }));
        assert!(ranges.ignored.iter().any(|range| {
            xml[range.clone()]
                .windows(b"inner".len())
                .any(|window| window == b"inner")
        }));
        let nested_wrapper = ranges
            .wrappers
            .iter()
            .min_by_key(|range| range.end.saturating_sub(range.start))
            .expect("nested wrapper range");
        assert!(!ranges.selected_parents.contains(nested_wrapper));
    }

    #[test]
    fn eager_inventory_budget_is_held_through_scan_and_released_after_drop() {
        let package = eager_inventory_package(1);
        let (budget, context) = managed_test_context(100);
        let inventory =
            eager_unreferenced_properties(&package, &OwnerLimits::default(), Some(&context))
                .expect("eager inventory should be admitted");
        let eager_memory = budget.used(Resource::Memory);
        let eager_objects = budget.used(Resource::Objects);
        assert!(eager_memory > 0);
        assert!(eager_objects > 0);

        let relationships = Relationships::new("/xl/worksheets".to_owned());
        let worksheet = eager_loaded_part(
            None,
            ct::SML_WORKSHEET,
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#
                .to_vec(),
            0,
            false,
            &relationships,
            OwnerLimits::default().max_sidecar_relationship_bytes,
            "worksheet bytes",
            None,
        )
        .expect("test worksheet part");
        let load = |_uri: &PackURI, _maximum: usize, _resource: &'static str| {
            Err(owner_invalid("test worksheet has no related parts"))
        };
        let collection = scan_owner(
            worksheet,
            &relationships,
            &load,
            &OwnerLimits::default(),
            None,
            Some(context),
            inventory.as_slice(),
            inventory.projection_bytes(),
        )
        .expect("empty worksheet scan should complete");
        assert_eq!(collection.diagnostics().len(), 1);
        assert!(budget.used(Resource::Memory) >= eager_memory);
        assert!(budget.used(Resource::Objects) >= eager_objects);

        drop(collection);
        assert_eq!(budget.used(Resource::Memory), eager_memory);
        assert_eq!(budget.used(Resource::Objects), eager_objects);
        drop(inventory);
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(budget.used(Resource::Objects), 0);
    }

    #[test]
    fn eager_inventory_projection_is_part_of_no_context_scan_cap() {
        let package = eager_inventory_package(1);
        let mut limits = OwnerLimits::default();
        let inventory = eager_unreferenced_properties(&package, &limits, None)
            .expect("eager inventory should fit its exact projection cap");
        let scan_projection_floor = size_of::<SourceControl>()
            .checked_mul(8.min(limits.max_controls))
            .expect("test projection floor");
        limits.max_projection_bytes = inventory
            .projection_bytes()
            .checked_add(scan_projection_floor)
            .expect("test aggregate projection cap");

        let relationships = Relationships::new("/xl/worksheets".to_owned());
        let worksheet = eager_loaded_part(
            None,
            ct::SML_WORKSHEET,
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#
                .to_vec(),
            0,
            false,
            &relationships,
            limits.max_sidecar_relationship_bytes,
            "worksheet bytes",
            None,
        )
        .expect("test worksheet part");
        let load = |_uri: &PackURI, _maximum: usize, _resource: &'static str| {
            Err(owner_invalid("test worksheet has no related parts"))
        };
        let error = scan_owner(
            worksheet,
            &relationships,
            &load,
            &limits,
            None,
            None,
            inventory.as_slice(),
            inventory.projection_bytes(),
        )
        .expect_err("inventory plus collection diagnostics must exceed the cap");
        assert!(matches!(
            error,
            FormControlOwnerError::Limit {
                resource: "form-control collection projection bytes",
                observed,
                maximum,
            } if observed > maximum && maximum == limits.max_projection_bytes
        ));

        let no_op_limits = OwnerLimits::default().with_max_projection_bytes(scan_projection_floor);
        let no_op_relationships = Relationships::new("/xl/worksheets".to_owned());
        let no_op_worksheet = eager_loaded_part(
            None,
            ct::SML_WORKSHEET,
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#
                .to_vec(),
            0,
            false,
            &no_op_relationships,
            no_op_limits.max_sidecar_relationship_bytes,
            "worksheet bytes",
            None,
        )
        .expect("test worksheet part");
        let no_op = scan_owner(
            no_op_worksheet,
            &no_op_relationships,
            &load,
            &no_op_limits,
            None,
            None,
            &[],
            0,
        )
        .expect("an empty owner has no projection allocation");
        assert!(no_op.controls.is_empty());
        assert!(no_op.diagnostics().is_empty());
    }

    #[test]
    fn eager_inventory_failure_releases_reservations() {
        let package = eager_inventory_package(2);
        let (budget, context) = managed_test_context(2);
        let error =
            eager_unreferenced_properties(&package, &OwnerLimits::default(), Some(&context))
                .expect_err("the second census pass should exceed the work budget");
        assert!(matches!(
            error,
            FormControlOwnerError::Execution(
                litchi_core::ExecutionError::ResourceLimit(limit)
            ) if limit.resource == Resource::Work
        ));
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(budget.used(Resource::Objects), 0);
    }
}
