//! Typed paragraph properties supported by the native table-cell style graph.

use crate::text::paragraph_alignment::native::{
    ParagraphStyleOverrides, inherited_alignment, inherited_text_style,
};
use crate::text::{TextAlignment, TextStyle};
use crate::{Error, IWorkPackage, Result};

#[derive(Debug, Clone, Copy, PartialEq)]
pub(super) enum CellParagraphProperty {
    Alignment(TextAlignment),
    TextStyle(TextStyle),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum CellParagraphPropertyKind {
    Alignment,
    TextStyle,
}

impl CellParagraphProperty {
    pub(super) const fn kind(self) -> CellParagraphPropertyKind {
        match self {
            Self::Alignment(_) => CellParagraphPropertyKind::Alignment,
            Self::TextStyle(_) => CellParagraphPropertyKind::TextStyle,
        }
    }

    pub(super) fn inherited(
        package: &IWorkPackage,
        style_id: u64,
        kind: CellParagraphPropertyKind,
    ) -> Result<Self> {
        match kind {
            CellParagraphPropertyKind::Alignment => {
                inherited_alignment(package, style_id).map(Self::Alignment)
            },
            CellParagraphPropertyKind::TextStyle => {
                inherited_text_style(package, style_id).map(Self::TextStyle)
            },
        }
    }

    pub(super) fn apply_to(
        self,
        overrides: &mut ParagraphStyleOverrides,
        inherited: Self,
    ) -> Result<()> {
        match (self, inherited) {
            (Self::Alignment(value), Self::Alignment(_)) => {
                overrides.alignment = Some(value);
            },
            (Self::TextStyle(value), Self::TextStyle(inherited)) => {
                overrides.point_size =
                    (value.point_size != inherited.point_size).then_some(value.point_size);
                overrides.bold = (value.bold != inherited.bold).then_some(value.bold);
                overrides.italic = (value.italic != inherited.italic).then_some(value.italic);
            },
            _ => {
                return Err(Error::InvalidFormat(
                    "iWork table-cell paragraph property kind mismatch".to_owned(),
                ));
            },
        }
        Ok(())
    }
}

impl CellParagraphPropertyKind {
    pub(super) const fn name(self) -> &'static str {
        match self {
            Self::Alignment => "text alignment",
            Self::TextStyle => "character formatting",
        }
    }

    pub(super) fn has_direct(self, overrides: &ParagraphStyleOverrides) -> bool {
        match self {
            Self::Alignment => overrides.alignment.is_some(),
            Self::TextStyle => {
                overrides.point_size.is_some()
                    || overrides.bold.is_some()
                    || overrides.italic.is_some()
            },
        }
    }

    pub(super) fn clear(self, overrides: &mut ParagraphStyleOverrides) {
        match self {
            Self::Alignment => overrides.alignment = None,
            Self::TextStyle => {
                overrides.point_size = None;
                overrides.bold = None;
                overrides.italic = None;
            },
        }
    }
}
