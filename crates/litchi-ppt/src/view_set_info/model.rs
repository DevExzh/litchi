use crate::non_zoom_view::NoZoomViewInfo;
use crate::view_info::Ratio;

/// State of one pane splitter bar (`NormalViewSetBarStates`, MS-PPT 2.13.16).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ViewBarState {
    /// The region occupies a minimal area of the view.
    Minimized,
    /// The region has an intermediate size.
    Restored,
    /// The region occupies a maximal area of the view.
    Maximized,
}

/// Pane splitter state of the normal three-pane view (`NormalViewSetInfo9Atom`,
/// MS-PPT 2.4.21.3).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct NormalViewSetInfo {
    pub(super) left_portion: Ratio,
    pub(super) top_portion: Ratio,
    pub(super) vert_bar_state: ViewBarState,
    pub(super) horiz_bar_state: ViewBarState,
    pub(super) prefer_single_set: bool,
    pub(super) hide_thumbnails: bool,
    pub(super) bar_snapped: bool,
}

impl NormalViewSetInfo {
    /// Width of the side content pane as a fraction of the view width.
    pub const fn left_portion(&self) -> Ratio {
        self.left_portion
    }

    /// Height of the slide pane as a fraction of the view height.
    pub const fn top_portion(&self) -> Ratio {
        self.top_portion
    }

    /// State of the vertical splitter bar.
    pub const fn vert_bar_state(&self) -> ViewBarState {
        self.vert_bar_state
    }

    /// State of the horizontal splitter bar.
    pub const fn horiz_bar_state(&self) -> ViewBarState {
        self.horiz_bar_state
    }

    /// Whether the view consists of only the slide pane.
    pub const fn prefer_single_set(&self) -> bool {
        self.prefer_single_set
    }

    /// Whether the side content pane shows comments instead of thumbnails.
    pub const fn hide_thumbnails(&self) -> bool {
        self.hide_thumbnails
    }

    /// Whether the vertical bar snaps to specific positions when resized.
    pub const fn bar_snapped(&self) -> bool {
        self.bar_snapped
    }
}

/// The payload of a `NormalViewSetInfo9Atom`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum NormalViewSetPayload {
    /// The specification pane layout (MS-PPT 2.4.21.3).
    Layout(NormalViewSetInfo),
    /// An opaque payload preserved verbatim. POI's undocumented
    /// `SheetPropertiesAtom` (document timestamps) occupies the same record
    /// type in many files and falls into this variant.
    Other(Vec<u8>),
}

/// A `NormalViewSetInfo9Container` (MS-PPT 2.4.21.2).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NormalViewSet {
    pub(super) payload: NormalViewSetPayload,
}

impl NormalViewSet {
    /// The pane-layout payload.
    pub const fn payload(&self) -> &NormalViewSetPayload {
        &self.payload
    }

    /// The pane layout, when the atom carries the specification payload.
    pub const fn layout(&self) -> Option<&NormalViewSetInfo> {
        match &self.payload {
            NormalViewSetPayload::Layout(layout) => Some(layout),
            NormalViewSetPayload::Other(_) => None,
        }
    }
}

/// A `NotesTextViewInfo9Container` (MS-PPT 2.4.21.4): scaling of the
/// notes-text view.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NotesTextViewInfo {
    pub(super) view_info: NoZoomViewInfo,
}

impl NotesTextViewInfo {
    /// The notes-text view scaling and origin.
    pub const fn view_info(&self) -> &NoZoomViewInfo {
        &self.view_info
    }
}
