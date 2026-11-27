use super::super::model::{Frame, Kind, Mode};

/// Typed, inert specification for an OLE object to add or replace.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Definition {
    pub(crate) kind: Kind,
    pub(crate) mode: Mode,
    pub(crate) anchor: Frame,
    pub(crate) name: Option<String>,
    pub(crate) program_id: Option<String>,
    pub(crate) show_as_icon: Option<bool>,
    pub(crate) preview_width: Option<u32>,
    pub(crate) preview_height: Option<u32>,
    pub(crate) payload: Option<Vec<u8>>,
    pub(crate) target: Option<String>,
}

impl Definition {
    /// Create an embedded OLE/package object whose payload stays opaque.
    pub fn embedded(kind: Kind, anchor: Frame, payload: impl Into<Vec<u8>>) -> Self {
        Self {
            kind,
            mode: Mode::Embedded,
            anchor,
            name: None,
            program_id: None,
            show_as_icon: None,
            preview_width: None,
            preview_height: None,
            payload: Some(payload.into()),
            target: None,
        }
    }

    /// Create a linked object; the target is retained as an inert external URI.
    pub fn linked(kind: Kind, anchor: Frame, target: impl Into<String>) -> Self {
        Self {
            kind,
            mode: Mode::Linked,
            anchor,
            name: None,
            program_id: None,
            show_as_icon: None,
            preview_width: None,
            preview_height: None,
            payload: None,
            target: Some(target.into()),
        }
    }

    #[must_use]
    pub fn kind(&self) -> Kind {
        self.kind
    }

    #[must_use]
    pub fn mode(&self) -> Mode {
        self.mode
    }

    #[must_use]
    pub fn anchor(&self) -> Frame {
        self.anchor
    }

    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    #[must_use]
    pub fn program_id(&self) -> Option<&str> {
        self.program_id.as_deref()
    }

    #[must_use]
    pub fn show_as_icon(&self) -> Option<bool> {
        self.show_as_icon
    }

    #[must_use]
    pub fn preview_size(&self) -> Option<(u32, u32)> {
        self.preview_width.zip(self.preview_height)
    }

    #[must_use]
    pub fn payload(&self) -> Option<&[u8]> {
        self.payload.as_deref()
    }

    #[must_use]
    pub fn target(&self) -> Option<&str> {
        self.target.as_deref()
    }

    #[must_use]
    pub fn set_name(mut self, value: impl Into<String>) -> Self {
        self.name = Some(value.into());
        self
    }

    #[must_use]
    pub fn clear_name(mut self) -> Self {
        self.name = None;
        self
    }

    #[must_use]
    pub fn set_program_id(mut self, value: impl Into<String>) -> Self {
        self.program_id = Some(value.into());
        self
    }

    #[must_use]
    pub fn set_show_as_icon(mut self, value: Option<bool>) -> Self {
        self.show_as_icon = value;
        self
    }

    #[must_use]
    pub fn set_preview_size(mut self, value: Option<(u32, u32)>) -> Self {
        self.preview_width = value.map(|value| value.0);
        self.preview_height = value.map(|value| value.1);
        self
    }
}
