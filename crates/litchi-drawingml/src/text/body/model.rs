//! Typed, host-neutral DrawingML text-body values.

use super::super::{Anchor, Autofit, Columns, Coordinate32, Direction, TextSize, Underline, Wrap};

/// Default horizontal body inset in EMUs (0.1 inch).
const DEFAULT_HORIZONTAL_INSET_EMU: i32 = 91_440;
/// Default vertical body inset in EMUs (0.05 inch).
const DEFAULT_VERTICAL_INSET_EMU: i32 = 45_720;

/// Text insets from `a:bodyPr`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Insets {
    /// Left inset.
    pub left: Coordinate32,
    /// Top inset.
    pub top: Coordinate32,
    /// Right inset.
    pub right: Coordinate32,
    /// Bottom inset.
    pub bottom: Coordinate32,
}

impl Default for Insets {
    fn default() -> Self {
        Self {
            left: Coordinate32::from(DEFAULT_HORIZONTAL_INSET_EMU),
            top: Coordinate32::from(DEFAULT_VERTICAL_INSET_EMU),
            right: Coordinate32::from(DEFAULT_HORIZONTAL_INSET_EMU),
            bottom: Coordinate32::from(DEFAULT_VERTICAL_INSET_EMU),
        }
    }
}

/// Text-body properties from `a:bodyPr`, with schema defaults applied.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Properties {
    /// Text insets.
    pub insets: Insets,
    /// Vertical anchoring of the text.
    pub vertical_anchor: Anchor,
    /// Whether the anchor point is horizontally centered (`anchorCtr`).
    pub anchor_center: bool,
    /// Text direction.
    pub direction: Direction,
    /// Text wrap behavior.
    pub wrap: Wrap,
    /// Autofit behavior.
    pub autofit: Autofit,
    /// Number of text columns (`numCol`; one when absent).
    pub column_count: Columns,
    /// Whether paragraph spacing is ignored at the first and last paragraphs.
    pub space_first_last_paragraph: bool,
}

impl Default for Properties {
    fn default() -> Self {
        Self {
            insets: Insets::default(),
            vertical_anchor: Anchor::default(),
            anchor_center: false,
            direction: Direction::default(),
            wrap: Wrap::default(),
            autofit: Autofit::default(),
            column_count: Columns::ONE,
            space_first_last_paragraph: false,
        }
    }
}

/// One DrawingML run inside a text-body paragraph.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Run {
    /// Run text.
    pub text: String,
    /// Explicit bold toggle (`a:rPr@b`), when declared.
    pub bold: Option<bool>,
    /// Explicit italic toggle (`a:rPr@i`), when declared.
    pub italic: Option<bool>,
    /// Exact underline style (`a:rPr@u`), when declared.
    pub underline: Option<Underline>,
    /// Font size in hundredths of a point (`a:rPr@sz`), when declared.
    pub font_size: Option<TextSize>,
}

/// One paragraph inside a DrawingML text body.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Paragraph {
    /// Runs in document order.
    pub runs: Vec<Run>,
}

impl Paragraph {
    /// Concatenate the paragraph's run text without evaluating fields.
    pub fn text(&self) -> String {
        self.runs.iter().map(|run| run.text.as_str()).collect()
    }
}

/// One host-neutral DrawingML text story.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Body {
    /// Body properties with DrawingML defaults applied.
    pub properties: Properties,
    /// Paragraphs in document order.
    pub paragraphs: Vec<Paragraph>,
}

impl Body {
    /// Return all paragraph text separated by newline characters.
    pub fn text(&self) -> String {
        self.paragraphs
            .iter()
            .map(Paragraph::text)
            .collect::<Vec<_>>()
            .join("\n")
    }
}

#[cfg(test)]
mod tests {
    use super::{Body, Insets, Paragraph, Properties, Run};

    #[test]
    fn applies_drawingml_defaults_and_joins_text() {
        let body = Body {
            properties: Properties::default(),
            paragraphs: vec![
                Paragraph {
                    runs: vec![Run {
                        text: "first".to_owned(),
                        ..Run::default()
                    }],
                },
                Paragraph {
                    runs: vec![Run {
                        text: "second".to_owned(),
                        ..Run::default()
                    }],
                },
            ],
        };
        assert_eq!(body.text(), "first\nsecond");
        assert_eq!(body.properties.insets, Insets::default());
    }
}
