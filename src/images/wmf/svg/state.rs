//! Graphics state management for WMF rendering
//!
//! Tracks the current drawing state (device context) including:
//! - Selected GDI objects (pens, brushes, fonts)
//! - Current position
//! - Text and background colors

/// WMF pen (for stroking)
#[derive(Debug, Clone, Copy, Default)]
pub struct Pen {
    pub style: u16,
    pub width: u16,
    pub color: u32,
}

/// WMF brush (for filling)
#[derive(Debug, Clone, Copy)]
pub struct Brush {
    pub style: u16,
    pub color: u32,
}

impl Default for Brush {
    fn default() -> Self {
        Self {
            style: 1, // BS_NULL (no fill) is the default
            color: 0xFFFFFF,
        }
    }
}

/// WMF font
#[derive(Debug, Clone, Default)]
pub struct Font {
    pub height: i16,
    pub escapement: u16, // rotation in tenths of degrees
    pub weight: u16,
    pub italic: bool,
    pub underline: bool,
    pub strike_out: bool,
    pub name: String,
}

/// GDI object stored in object table
#[derive(Debug, Clone)]
pub enum GdiObject {
    Pen(Pen),
    Brush(Brush),
    Font(Font),
}

/// Graphics state (device context)
#[derive(Debug, Clone, Default)]
pub struct GraphicsState {
    pub position: (i16, i16),
    pub objects: Vec<Option<GdiObject>>,
    pub pen: Pen,
    pub brush: Brush,
    pub font: Font,
    pub text_color: u32,
    pub bk_color: u32,
    pub poly_fill_mode: u16, // 1=ALTERNATE, 2=WINDING
}

impl GraphicsState {
    pub fn new() -> Self {
        Self {
            font: Font {
                height: 12,
                name: "serif".to_string(),
                ..Default::default()
            },
            poly_fill_mode: 1, // ALTERNATE (evenodd) is default
            ..Default::default()
        }
    }
}
