// Character and value mappings from AST types to OMML representations
//
// These are the inverses of the lookup tables used by the OMML parser
// (`crate::omml::lookup`), so values emitted here survive a parse round-trip.

use crate::ast::{
    AccentType, Alignment, Fence, FractionType, FunctionName, LargeOperator, MatrixFence, Operator,
    Position, PredefinedSymbol, SpaceType, StyleType, VerticalAlignment,
};

/// Opening character for a fence type (empty string means "no fence")
pub fn fence_open_char(fence: Fence) -> &'static str {
    match fence {
        Fence::Paren => "(",
        Fence::Bracket => "[",
        Fence::Brace => "{",
        Fence::Angle => "⟨",
        Fence::Pipe => "|",
        Fence::DoublePipe => "‖",
        Fence::Floor => "⌊",
        Fence::Ceiling => "⌈",
        Fence::AngleBracket => "⟪",
        Fence::SquareBracket => "⟦",
        Fence::CurlyBrace => "⦃",
        Fence::None => "",
    }
}

/// Closing character for a fence type (empty string means "no fence")
pub fn fence_close_char(fence: Fence) -> &'static str {
    match fence {
        Fence::Paren => ")",
        Fence::Bracket => "]",
        Fence::Brace => "}",
        Fence::Angle => "⟩",
        Fence::Pipe => "|",
        Fence::DoublePipe => "‖",
        Fence::Floor => "⌋",
        Fence::Ceiling => "⌉",
        Fence::AngleBracket => "⟫",
        Fence::SquareBracket => "⟧",
        Fence::CurlyBrace => "⦄",
        Fence::None => "",
    }
}

/// Map a matrix fence to the equivalent delimiter fence pair
pub fn matrix_fence_pair(fence: MatrixFence) -> (Fence, Fence) {
    match fence {
        MatrixFence::None => (Fence::None, Fence::None),
        MatrixFence::Paren => (Fence::Paren, Fence::Paren),
        MatrixFence::Bracket => (Fence::Bracket, Fence::Bracket),
        MatrixFence::Brace => (Fence::Brace, Fence::Brace),
        MatrixFence::Pipe => (Fence::Pipe, Fence::Pipe),
        MatrixFence::DoublePipe => (Fence::DoublePipe, Fence::DoublePipe),
    }
}

/// Character (or word) used in `m:naryPr/m:chr` for a large operator
pub fn large_operator_char(op: LargeOperator) -> &'static str {
    match op {
        LargeOperator::Sum => "∑",
        LargeOperator::Product => "∏",
        LargeOperator::Coproduct => "∐",
        LargeOperator::Integral => "∫",
        LargeOperator::DoubleIntegral => "∬",
        LargeOperator::TripleIntegral => "∭",
        LargeOperator::ContourIntegral => "∮",
        LargeOperator::SurfaceIntegral => "∯",
        LargeOperator::VolumeIntegral => "∰",
        LargeOperator::Union => "⋃",
        LargeOperator::Intersection => "⋂",
        LargeOperator::BigUnion => "⨄",
        LargeOperator::BigIntersection => "⨅",
        LargeOperator::Limit => "lim",
        LargeOperator::Max => "max",
        LargeOperator::Min => "min",
        LargeOperator::Supremum => "sup",
        LargeOperator::Infimum => "inf",
        LargeOperator::ArgMax => "argmax",
        LargeOperator::ArgMin => "argmin",
    }
}

/// Combining character used in `m:accPr/m:chr` for an accent
pub fn accent_char(accent: AccentType) -> &'static str {
    match accent {
        AccentType::Hat => "\u{0302}",
        AccentType::Check => "\u{030C}",
        AccentType::Tilde => "\u{0303}",
        AccentType::Acute => "\u{0301}",
        AccentType::Grave => "\u{0300}",
        AccentType::Dot => "\u{0307}",
        AccentType::DoubleDot => "\u{0308}",
        AccentType::TripleDot => "\u{20DB}",
        AccentType::Bar => "\u{0305}",
        AccentType::Breve => "\u{0306}",
        AccentType::Vec => "\u{20D7}",
    }
}

/// Unicode text for a mathematical operator
pub fn operator_char(op: Operator) -> &'static str {
    match op {
        Operator::Plus => "+",
        Operator::Minus => "−",
        Operator::Multiply => "×",
        Operator::Divide => "÷",
        Operator::PlusMinus => "±",
        Operator::MinusPlus => "∓",
        Operator::Equals => "=",
        Operator::NotEquals => "≠",
        Operator::LessThan => "<",
        Operator::GreaterThan => ">",
        Operator::LessThanOrEqual => "≤",
        Operator::GreaterThanOrEqual => "≥",
        Operator::Times => "×",
        Operator::Dot => "⋅",
        Operator::Cross => "×",
        Operator::Star => "∗",
        Operator::Circle => "∘",
        Operator::Circ => "∘",
        Operator::Bullet => "∙",
        Operator::Wedge => "∧",
        Operator::Vee => "∨",
        Operator::Cap => "∩",
        Operator::Cup => "∪",
        Operator::In => "∈",
        Operator::NotIn => "∉",
        Operator::Subset => "⊂",
        Operator::Superset => "⊃",
        Operator::SubsetEq => "⊆",
        Operator::SupersetEq => "⊇",
        Operator::EmptySet => "∅",
        Operator::Union => "∪",
        Operator::Intersection => "∩",
        Operator::Approx => "≈",
        Operator::Cong => "≅",
        Operator::Equiv => "≡",
        Operator::Propto => "∝",
        Operator::Sim => "∼",
        Operator::Simeq => "≃",
        Operator::Asymp => "≍",
        Operator::Parallel => "∥",
        Operator::Perpendicular => "⊥",
        Operator::Angle => "∠",
        Operator::Nabla => "∇",
        Operator::Partial => "∂",
        Operator::Differential => "ⅆ",
        Operator::Infinity => "∞",
        Operator::Aleph => "ℵ",
        Operator::Prime => "′",
        Operator::DoublePrime => "″",
        Operator::TriplePrime => "‴",
        Operator::Ellipsis => "…",
        Operator::CDots => "⋯",
        Operator::VDots => "⋮",
        Operator::DDots => "⋱",
        Operator::Ldots => "…",
        Operator::LeftArrow => "←",
        Operator::RightArrow => "→",
        Operator::UpArrow => "↑",
        Operator::DownArrow => "↓",
        Operator::LeftRightArrow => "↔",
        Operator::UpDownArrow => "↕",
        Operator::ForAll => "∀",
        Operator::Exists => "∃",
        Operator::Not => "¬",
        Operator::And => "∧",
        Operator::Or => "∨",
        Operator::Implies => "⇒",
        Operator::Iff => "⇔",
        Operator::Therefore => "∴",
        Operator::Because => "∵",
        Operator::Box => "□",
        Operator::Diamond => "◆",
        Operator::Square => "■",
    }
}

/// Unicode text for a predefined symbol
pub fn predefined_symbol_char(symbol: PredefinedSymbol) -> &'static str {
    match symbol {
        PredefinedSymbol::Alpha => "α",
        PredefinedSymbol::Beta => "β",
        PredefinedSymbol::Gamma => "γ",
        PredefinedSymbol::Delta => "δ",
        PredefinedSymbol::Epsilon => "ε",
        PredefinedSymbol::Zeta => "ζ",
        PredefinedSymbol::Eta => "η",
        PredefinedSymbol::Theta => "θ",
        PredefinedSymbol::Iota => "ι",
        PredefinedSymbol::Kappa => "κ",
        PredefinedSymbol::Lambda => "λ",
        PredefinedSymbol::Mu => "μ",
        PredefinedSymbol::Nu => "ν",
        PredefinedSymbol::Xi => "ξ",
        PredefinedSymbol::Omicron => "ο",
        PredefinedSymbol::Pi => "π",
        PredefinedSymbol::Rho => "ρ",
        PredefinedSymbol::Sigma => "σ",
        PredefinedSymbol::Tau => "τ",
        PredefinedSymbol::Upsilon => "υ",
        PredefinedSymbol::Phi => "φ",
        PredefinedSymbol::Chi => "χ",
        PredefinedSymbol::Psi => "ψ",
        PredefinedSymbol::Omega => "ω",
        PredefinedSymbol::AlphaCap => "Α",
        PredefinedSymbol::BetaCap => "Β",
        PredefinedSymbol::GammaCap => "Γ",
        PredefinedSymbol::DeltaCap => "Δ",
        PredefinedSymbol::EpsilonCap => "Ε",
        PredefinedSymbol::ZetaCap => "Ζ",
        PredefinedSymbol::EtaCap => "Η",
        PredefinedSymbol::ThetaCap => "Θ",
        PredefinedSymbol::IotaCap => "Ι",
        PredefinedSymbol::KappaCap => "Κ",
        PredefinedSymbol::LambdaCap => "Λ",
        PredefinedSymbol::MuCap => "Μ",
        PredefinedSymbol::NuCap => "Ν",
        PredefinedSymbol::XiCap => "Ξ",
        PredefinedSymbol::OmicronCap => "Ο",
        PredefinedSymbol::PiCap => "Π",
        PredefinedSymbol::RhoCap => "Ρ",
        PredefinedSymbol::SigmaCap => "Σ",
        PredefinedSymbol::TauCap => "Τ",
        PredefinedSymbol::UpsilonCap => "Υ",
        PredefinedSymbol::PhiCap => "Φ",
        PredefinedSymbol::ChiCap => "Χ",
        PredefinedSymbol::PsiCap => "Ψ",
        PredefinedSymbol::OmegaCap => "Ω",
        PredefinedSymbol::Aleph => "ℵ",
        PredefinedSymbol::EulerGamma => "γ",
        PredefinedSymbol::ExponentialE => "e",
        PredefinedSymbol::ImaginaryI => "i",
        PredefinedSymbol::Infinity => "∞",
    }
}

/// Canonical name for a predefined function
pub fn function_name_str(function: FunctionName) -> &'static str {
    match function {
        FunctionName::Sin => "sin",
        FunctionName::Cos => "cos",
        FunctionName::Tan => "tan",
        FunctionName::Sec => "sec",
        FunctionName::Csc => "csc",
        FunctionName::Cot => "cot",
        FunctionName::ArcSin => "arcsin",
        FunctionName::ArcCos => "arccos",
        FunctionName::ArcTan => "arctan",
        FunctionName::ArcSec => "arcsec",
        FunctionName::ArcCsc => "arccsc",
        FunctionName::ArcCot => "arccot",
        FunctionName::Sinh => "sinh",
        FunctionName::Cosh => "cosh",
        FunctionName::Tanh => "tanh",
        FunctionName::Sech => "sech",
        FunctionName::Csch => "csch",
        FunctionName::Coth => "coth",
        FunctionName::Log => "log",
        FunctionName::Ln => "ln",
        FunctionName::Exp => "exp",
        FunctionName::Sqrt => "sqrt",
        FunctionName::Min => "min",
        FunctionName::Max => "max",
        FunctionName::Sup => "sup",
        FunctionName::Inf => "inf",
        FunctionName::Lim => "lim",
        FunctionName::Det => "det",
        FunctionName::Trace => "trace",
        FunctionName::Dim => "dim",
        FunctionName::Ker => "ker",
        FunctionName::Im => "Im",
        FunctionName::Re => "Re",
        FunctionName::Arg => "arg",
        FunctionName::Mod => "mod",
        FunctionName::Gcd => "gcd",
        FunctionName::Lcm => "lcm",
    }
}

/// Which run-property element a style is expressed with
pub enum StyleElement {
    /// `m:scr` (math script/alphabet selection)
    Script,
    /// `m:sty` (bold/italic style)
    Style,
}

/// Map a style to the OMML run-property element and value expressing it
pub fn style_value(style: StyleType) -> (StyleElement, &'static str) {
    match style {
        StyleType::Normal => (StyleElement::Script, "roman"),
        StyleType::Bold => (StyleElement::Style, "b"),
        StyleType::Italic => (StyleElement::Style, "i"),
        StyleType::BoldItalic => (StyleElement::Style, "bi"),
        StyleType::SansSerif => (StyleElement::Script, "sans-serif"),
        StyleType::SansSerifBold => (StyleElement::Script, "sans-serif-bold"),
        StyleType::SansSerifItalic => (StyleElement::Script, "sans-serif-italic"),
        StyleType::SansSerifBoldItalic => (StyleElement::Script, "sans-serif-bold-italic"),
        StyleType::Monospace => (StyleElement::Script, "monospace"),
        StyleType::Script => (StyleElement::Script, "script"),
        StyleType::BoldScript => (StyleElement::Script, "bold-script"),
        StyleType::Fraktur => (StyleElement::Script, "fraktur"),
        StyleType::BoldFraktur => (StyleElement::Script, "bold-fraktur"),
        StyleType::DoubleStruck => (StyleElement::Script, "double-struck"),
    }
}

/// Value for `m:fPr/m:type`
pub fn fraction_type_value(frac_type: FractionType) -> &'static str {
    match frac_type {
        FractionType::Bar => "bar",
        FractionType::NoBar => "noBar",
        FractionType::Skewed => "skw",
    }
}

/// Value for `m:pos` position properties
pub fn position_value(position: Position) -> &'static str {
    match position {
        Position::Prefix => "pre",
        Position::Postfix => "post",
        Position::Infix => "in",
        Position::Top => "top",
        Position::Bottom => "bot",
    }
}

/// Value for `m:vertJc`
pub fn vertical_alignment_value(alignment: VerticalAlignment) -> &'static str {
    match alignment {
        VerticalAlignment::Top => "top",
        VerticalAlignment::Bottom => "bot",
        VerticalAlignment::Center => "center",
        VerticalAlignment::Baseline => "baseline",
        VerticalAlignment::Axis => "axis",
    }
}

/// Value for `m:baseJc` (only vertical alignments are representable)
pub fn base_alignment_value(alignment: Alignment) -> Option<&'static str> {
    match alignment {
        Alignment::Top => Some("top"),
        Alignment::Center | Alignment::Centered => Some("center"),
        Alignment::Bottom => Some("bottom"),
        Alignment::Baseline => Some("baseline"),
        _ => None,
    }
}

/// Unicode space character approximating an AST space type
pub fn space_char(space: SpaceType) -> &'static str {
    match space {
        SpaceType::Thin => "\u{2009}",
        SpaceType::Medium => "\u{2005}",
        SpaceType::Thick => "\u{2004}",
        SpaceType::Quad => "\u{2001}",
        SpaceType::QQuad => "\u{2001}\u{2001}",
        SpaceType::Negative => "\u{200B}",
    }
}
