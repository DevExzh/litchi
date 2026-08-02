//! Document-level defaults for RTF mathematics.

use crate::{RtfError, RtfResult};

macro_rules! numeric_enum {
    ($name:ident { $($value:literal => $variant:ident),+ $(,)? }) => {
        #[derive(Debug, Clone, Copy, PartialEq, Eq)]
        pub enum $name {
            $($variant,)+
            Unknown(i32),
        }

        impl $name {
            pub fn from_rtf(value: i32) -> Self {
                match value {
                    $($value => Self::$variant,)+
                    value => Self::Unknown(value),
                }
            }

            pub fn rtf_value(self) -> i32 {
                match self {
                    $(Self::$variant => $value,)+
                    Self::Unknown(value) => value,
                }
            }
        }
    };
}

numeric_enum!(MathBinaryOperatorBreak {
    0 => Before,
    1 => After,
    2 => Duplicate,
});

numeric_enum!(MathBinarySubtractionBreak {
    0 => MinusMinus,
    1 => PlusMinus,
    2 => MinusPlus,
});

numeric_enum!(MathJustification {
    1 => CenteredAsGroup,
    2 => Centered,
    3 => Left,
    4 => Right,
});

numeric_enum!(MathLimitPlacement {
    0 => SubscriptSuperscript,
    1 => UnderOver,
});

numeric_enum!(MathFlag {
    0 => Off,
    1 => On,
});

/// The optional members of the `\mmathPr` document-properties destination.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DocumentMathProperties {
    pub binary_operator_break: Option<MathBinaryOperatorBreak>,
    pub binary_subtraction_break: Option<MathBinarySubtractionBreak>,
    pub default_justification: Option<MathJustification>,
    pub display_defaults: Option<MathFlag>,
    pub inter_equation_spacing: Option<i32>,
    pub integral_limit_placement: Option<MathLimitPlacement>,
    pub intra_equation_spacing: Option<i32>,
    pub left_margin: Option<i32>,
    pub math_font: Option<u32>,
    pub nary_limit_placement: Option<MathLimitPlacement>,
    pub post_spacing: Option<i32>,
    pub pre_spacing: Option<i32>,
    pub right_margin: Option<i32>,
    pub small_fractions: Option<MathFlag>,
    pub wrap_indent: Option<i32>,
    pub wrap_right: Option<MathFlag>,
}

impl DocumentMathProperties {
    pub fn new() -> Self {
        Self::default()
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        if self.math_font.is_some_and(|font| font > i32::MAX as u32) {
            return Err(RtfError::MalformedDocument(
                "RTF math font index exceeds the signed control-word range".to_string(),
            ));
        }
        Ok(())
    }
}
