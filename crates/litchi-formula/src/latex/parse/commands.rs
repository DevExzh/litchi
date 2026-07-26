// Static command tables mapping LaTeX control sequences onto AST vocabulary
//
// These tables are the inverse of the AST-to-LaTeX tables in
// `latex::operators`, `latex::symbols` and `latex::conv::node`. Where the
// forward direction maps several AST variants onto the same LaTeX string, the
// inverse picks one representative so that a value survives a round trip
// unchanged.

use crate::ast::{
    AccentType, FunctionName, LargeOperator, Operator, PredefinedSymbol, SpaceType, StyleType,
};

/// Command name of the line break control sequence `\\`.
pub(crate) const CMD_LINE_BREAK: &str = "\\";
/// Command name of `\begin`.
pub(crate) const CMD_BEGIN: &str = "begin";
/// Command name of `\end`.
pub(crate) const CMD_END: &str = "end";
/// Command name of `\left`.
pub(crate) const CMD_LEFT: &str = "left";
/// Command name of `\right`.
pub(crate) const CMD_RIGHT: &str = "right";
/// Command name of `\limits`.
pub(crate) const CMD_LIMITS: &str = "limits";
/// Command name of `\nolimits`.
pub(crate) const CMD_NOLIMITS: &str = "nolimits";

/// Greek letters and named constants that have a dedicated AST variant.
///
/// The `var` spellings share the variant of their base letter: the AST has no
/// separate glyph for them, so `\varepsilon` renders as `\epsilon`.
pub(crate) static PREDEFINED_SYMBOLS: phf::Map<&'static str, PredefinedSymbol> = phf::phf_map! {
    "alpha"      => PredefinedSymbol::Alpha,
    "beta"       => PredefinedSymbol::Beta,
    "gamma"      => PredefinedSymbol::Gamma,
    "delta"      => PredefinedSymbol::Delta,
    "epsilon"    => PredefinedSymbol::Epsilon,
    "varepsilon" => PredefinedSymbol::Epsilon,
    "zeta"       => PredefinedSymbol::Zeta,
    "eta"        => PredefinedSymbol::Eta,
    "theta"      => PredefinedSymbol::Theta,
    "vartheta"   => PredefinedSymbol::Theta,
    "iota"       => PredefinedSymbol::Iota,
    "kappa"      => PredefinedSymbol::Kappa,
    "lambda"     => PredefinedSymbol::Lambda,
    "mu"         => PredefinedSymbol::Mu,
    "nu"         => PredefinedSymbol::Nu,
    "xi"         => PredefinedSymbol::Xi,
    "omicron"    => PredefinedSymbol::Omicron,
    "pi"         => PredefinedSymbol::Pi,
    "varpi"      => PredefinedSymbol::Pi,
    "rho"        => PredefinedSymbol::Rho,
    "varrho"     => PredefinedSymbol::Rho,
    "sigma"      => PredefinedSymbol::Sigma,
    "varsigma"   => PredefinedSymbol::Sigma,
    "tau"        => PredefinedSymbol::Tau,
    "upsilon"    => PredefinedSymbol::Upsilon,
    "phi"        => PredefinedSymbol::Phi,
    "varphi"     => PredefinedSymbol::Phi,
    "chi"        => PredefinedSymbol::Chi,
    "psi"        => PredefinedSymbol::Psi,
    "omega"      => PredefinedSymbol::Omega,

    "Alpha"      => PredefinedSymbol::AlphaCap,
    "Beta"       => PredefinedSymbol::BetaCap,
    "Gamma"      => PredefinedSymbol::GammaCap,
    "Delta"      => PredefinedSymbol::DeltaCap,
    "Epsilon"    => PredefinedSymbol::EpsilonCap,
    "Zeta"       => PredefinedSymbol::ZetaCap,
    "Eta"        => PredefinedSymbol::EtaCap,
    "Theta"      => PredefinedSymbol::ThetaCap,
    "Iota"       => PredefinedSymbol::IotaCap,
    "Kappa"      => PredefinedSymbol::KappaCap,
    "Lambda"     => PredefinedSymbol::LambdaCap,
    "Mu"         => PredefinedSymbol::MuCap,
    "Nu"         => PredefinedSymbol::NuCap,
    "Xi"         => PredefinedSymbol::XiCap,
    "Omicron"    => PredefinedSymbol::OmicronCap,
    "Pi"         => PredefinedSymbol::PiCap,
    "Rho"        => PredefinedSymbol::RhoCap,
    "Sigma"      => PredefinedSymbol::SigmaCap,
    "Tau"        => PredefinedSymbol::TauCap,
    "Upsilon"    => PredefinedSymbol::UpsilonCap,
    "Phi"        => PredefinedSymbol::PhiCap,
    "Chi"        => PredefinedSymbol::ChiCap,
    "Psi"        => PredefinedSymbol::PsiCap,
    "Omega"      => PredefinedSymbol::OmegaCap,

    "aleph"      => PredefinedSymbol::Aleph,
    "infty"      => PredefinedSymbol::Infinity,
};

/// Control sequences that map onto an [`Operator`].
pub(crate) static OPERATORS: phf::Map<&'static str, Operator> = phf::phf_map! {
    "pm"        => Operator::PlusMinus,
    "mp"        => Operator::MinusPlus,
    "cdot"      => Operator::Dot,
    "div"       => Operator::Divide,
    "times"     => Operator::Times,
    "ast"       => Operator::Star,
    "star"      => Operator::Star,
    "circ"      => Operator::Circ,
    "bullet"    => Operator::Bullet,
    "wedge"     => Operator::Wedge,
    "vee"       => Operator::Vee,
    "cap"       => Operator::Cap,
    "cup"       => Operator::Cup,

    "neq"       => Operator::NotEquals,
    "ne"        => Operator::NotEquals,
    "leq"       => Operator::LessThanOrEqual,
    "le"        => Operator::LessThanOrEqual,
    "geq"       => Operator::GreaterThanOrEqual,
    "ge"        => Operator::GreaterThanOrEqual,

    "in"        => Operator::In,
    "notin"     => Operator::NotIn,
    "subset"    => Operator::Subset,
    "supset"    => Operator::Superset,
    "subseteq"  => Operator::SubsetEq,
    "supseteq"  => Operator::SupersetEq,
    "emptyset"  => Operator::EmptySet,

    "approx"    => Operator::Approx,
    "cong"      => Operator::Cong,
    "equiv"     => Operator::Equiv,
    "propto"    => Operator::Propto,
    "sim"       => Operator::Sim,
    "simeq"     => Operator::Simeq,
    "asymp"     => Operator::Asymp,

    "parallel"  => Operator::Parallel,
    "perp"      => Operator::Perpendicular,
    "angle"     => Operator::Angle,

    "nabla"     => Operator::Nabla,
    "partial"   => Operator::Partial,

    "ldots"     => Operator::Ldots,
    "dots"      => Operator::Ldots,
    "cdots"     => Operator::CDots,
    "vdots"     => Operator::VDots,
    "ddots"     => Operator::DDots,

    "leftarrow"      => Operator::LeftArrow,
    "rightarrow"     => Operator::RightArrow,
    "uparrow"        => Operator::UpArrow,
    "downarrow"      => Operator::DownArrow,
    "leftrightarrow" => Operator::LeftRightArrow,
    "updownarrow"    => Operator::UpDownArrow,

    "forall"    => Operator::ForAll,
    "exists"    => Operator::Exists,
    "neg"       => Operator::Not,
    "lnot"      => Operator::Not,
    "land"      => Operator::And,
    "lor"       => Operator::Or,
    "implies"   => Operator::Implies,
    "iff"       => Operator::Iff,

    "therefore" => Operator::Therefore,
    "because"   => Operator::Because,
    "Box"       => Operator::Box,
    "Diamond"   => Operator::Diamond,
    "square"    => Operator::Square,
};

/// Control sequences that map onto an [`AccentType`].
pub(crate) static ACCENTS: phf::Map<&'static str, AccentType> = phf::phf_map! {
    "hat"        => AccentType::Hat,
    "widehat"    => AccentType::Hat,
    "check"      => AccentType::Check,
    "tilde"      => AccentType::Tilde,
    "widetilde"  => AccentType::Tilde,
    "acute"      => AccentType::Acute,
    "grave"      => AccentType::Grave,
    "dot"        => AccentType::Dot,
    "ddot"       => AccentType::DoubleDot,
    "dddot"      => AccentType::TripleDot,
    "bar"        => AccentType::Bar,
    "breve"      => AccentType::Breve,
    "vec"        => AccentType::Vec,
};

/// Control sequences that map onto a [`LargeOperator`].
///
/// `\min`, `\max`, `\sup` and `\inf` are deliberately absent: they parse as
/// named functions, matching how the AST-to-LaTeX direction renders them.
pub(crate) static LARGE_OPERATORS: phf::Map<&'static str, LargeOperator> = phf::phf_map! {
    "sum"    => LargeOperator::Sum,
    "prod"   => LargeOperator::Product,
    "coprod" => LargeOperator::Coproduct,
    "int"    => LargeOperator::Integral,
    "iint"   => LargeOperator::DoubleIntegral,
    "iiint"  => LargeOperator::TripleIntegral,
    "oint"   => LargeOperator::ContourIntegral,
    "oiint"  => LargeOperator::SurfaceIntegral,
    "oiiint" => LargeOperator::VolumeIntegral,
    "bigcup" => LargeOperator::Union,
    "bigcap" => LargeOperator::Intersection,
    "lim"    => LargeOperator::Limit,
};

/// Control sequences that map onto a [`FunctionName`].
pub(crate) static FUNCTIONS: phf::Map<&'static str, FunctionName> = phf::phf_map! {
    "sin"    => FunctionName::Sin,
    "cos"    => FunctionName::Cos,
    "tan"    => FunctionName::Tan,
    "sec"    => FunctionName::Sec,
    "csc"    => FunctionName::Csc,
    "cot"    => FunctionName::Cot,
    "arcsin" => FunctionName::ArcSin,
    "arccos" => FunctionName::ArcCos,
    "arctan" => FunctionName::ArcTan,
    "arcsec" => FunctionName::ArcSec,
    "arccsc" => FunctionName::ArcCsc,
    "arccot" => FunctionName::ArcCot,
    "sinh"   => FunctionName::Sinh,
    "cosh"   => FunctionName::Cosh,
    "tanh"   => FunctionName::Tanh,
    "sech"   => FunctionName::Sech,
    "csch"   => FunctionName::Csch,
    "coth"   => FunctionName::Coth,
    "log"    => FunctionName::Log,
    "ln"     => FunctionName::Ln,
    "exp"    => FunctionName::Exp,
    "min"    => FunctionName::Min,
    "max"    => FunctionName::Max,
    "sup"    => FunctionName::Sup,
    "inf"    => FunctionName::Inf,
    "det"    => FunctionName::Det,
    "trace"  => FunctionName::Trace,
    "dim"    => FunctionName::Dim,
    "ker"    => FunctionName::Ker,
    "arg"    => FunctionName::Arg,
    "mod"    => FunctionName::Mod,
    "bmod"   => FunctionName::Mod,
    "gcd"    => FunctionName::Gcd,
    "lcm"    => FunctionName::Lcm,
};

/// Named functions with no [`FunctionName`] variant, kept as free-form names.
pub(crate) static FUNCTION_WORDS: phf::Set<&'static str> = phf::phf_set! {
    "deg",
    "hom",
    "lg",
    "limsup",
    "liminf",
    "Pr",
};

/// Control sequences that select a math alphabet.
pub(crate) static MATH_STYLES: phf::Map<&'static str, StyleType> = phf::phf_map! {
    "mathrm"     => StyleType::Normal,
    "mathnormal" => StyleType::Italic,
    "mathbf"     => StyleType::Bold,
    "mathit"     => StyleType::Italic,
    "mathbb"     => StyleType::DoubleStruck,
    "mathcal"    => StyleType::Script,
    "mathscr"    => StyleType::Script,
    "mathfrak"   => StyleType::Fraktur,
    "mathsf"     => StyleType::SansSerif,
    "mathtt"     => StyleType::Monospace,
    "boldsymbol" => StyleType::BoldItalic,
    "bm"         => StyleType::BoldItalic,
};

/// Control sequences whose argument is verbatim text rather than math.
pub(crate) static TEXT_STYLES: phf::Map<&'static str, StyleType> = phf::phf_map! {
    "text"       => StyleType::Normal,
    "textrm"     => StyleType::Normal,
    "textnormal" => StyleType::Normal,
    "mbox"       => StyleType::Normal,
    "textbf"     => StyleType::Bold,
    "textit"     => StyleType::Italic,
    "texttt"     => StyleType::Monospace,
    "textsf"     => StyleType::SansSerif,
};

/// Control sequences that map onto a [`SpaceType`].
///
/// `\ ` (backslash space) has no exact AST counterpart and is treated as a
/// medium space.
pub(crate) static SPACES: phf::Map<&'static str, SpaceType> = phf::phf_map! {
    ","              => SpaceType::Thin,
    ":"              => SpaceType::Medium,
    ";"              => SpaceType::Thick,
    "!"              => SpaceType::Negative,
    " "              => SpaceType::Medium,
    "quad"           => SpaceType::Quad,
    "qquad"          => SpaceType::QQuad,
    "thinspace"      => SpaceType::Thin,
    "medspace"       => SpaceType::Medium,
    "thickspace"     => SpaceType::Thick,
    "enspace"        => SpaceType::Medium,
    "negthinspace"   => SpaceType::Negative,
    "negmedspace"    => SpaceType::Negative,
    "negthickspace"  => SpaceType::Negative,
};

/// Control sequences that carry only presentation intent the AST cannot model.
///
/// They are consumed and produce no node. `\(`, `\)`, `\[` and `\]` are listed
/// so that a complete `LatexConverter` output can be fed straight back in.
pub(crate) static IGNORED_COMMANDS: phf::Set<&'static str> = phf::phf_set! {
    "(",
    ")",
    "[",
    "]",
    "displaystyle",
    "textstyle",
    "scriptstyle",
    "scriptscriptstyle",
    "nonumber",
    "notag",
    "hline",
    "hfill",
    "noalign",
    "limits",
    "nolimits",
};

/// Unicode glyphs for delimiter control sequences.
///
/// Used when a delimiter appears without a partner, so that it still renders as
/// the intended bracket instead of degrading to `\text{...}`.
pub(crate) static DELIMITER_GLYPHS: phf::Map<&'static str, char> = phf::phf_map! {
    "{"      => '{',
    "}"      => '}',
    "|"      => '‖',
    "vert"   => '|',
    "Vert"   => '‖',
    "lvert"  => '|',
    "rvert"  => '|',
    "lVert"  => '‖',
    "rVert"  => '‖',
    "langle" => '⟨',
    "rangle" => '⟩',
    "lfloor" => '⌊',
    "rfloor" => '⌋',
    "lceil"  => '⌈',
    "rceil"  => '⌉',
    "lbrack" => '[',
    "rbrack" => ']',
    "lbrace" => '{',
    "rbrace" => '}',
};

/// Map a bare ASCII character onto an [`Operator`].
///
/// Only characters whose AST-to-LaTeX rendering is the identical glyph are
/// listed. `*`, `/`, `!` and `,` have no exact counterpart (`Star` renders as
/// `\ast`, `Divide` as `\div`), so the parser keeps them as literal text and
/// the round trip stays lossless.
#[inline]
pub(crate) const fn ascii_operator(ch: char) -> Option<Operator> {
    match ch {
        '+' => Some(Operator::Plus),
        '-' => Some(Operator::Minus),
        '=' => Some(Operator::Equals),
        '<' => Some(Operator::LessThan),
        '>' => Some(Operator::GreaterThan),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn greek_letters_resolve_to_predefined_symbols() {
        assert_eq!(
            PREDEFINED_SYMBOLS.get("alpha"),
            Some(&PredefinedSymbol::Alpha)
        );
        assert_eq!(
            PREDEFINED_SYMBOLS.get("Omega"),
            Some(&PredefinedSymbol::OmegaCap)
        );
        assert_eq!(
            PREDEFINED_SYMBOLS.get("varepsilon"),
            Some(&PredefinedSymbol::Epsilon)
        );
    }

    #[test]
    fn operator_table_inverts_the_rendering_table() {
        use crate::latex::operators::operator_to_latex;

        for (name, operator) in OPERATORS.entries() {
            let rendered = operator_to_latex(*operator);
            // Every rendered command must itself be a key of the table so that
            // a second parse yields the same operator.
            let stripped = rendered.strip_prefix('\\').unwrap_or(rendered);
            if stripped.chars().all(|ch| ch.is_ascii_alphabetic()) {
                assert!(
                    OPERATORS.contains_key(stripped),
                    "`{name}` renders as `{rendered}` which does not parse back"
                );
            }
        }
    }

    #[test]
    fn large_operator_table_inverts_the_rendering_table() {
        use crate::latex::operators::large_operator_to_latex;

        for (name, operator) in LARGE_OPERATORS.entries() {
            let rendered = large_operator_to_latex(*operator);
            assert_eq!(
                rendered.strip_prefix('\\'),
                Some(*name),
                "`{name}` does not round trip"
            );
        }
    }

    #[test]
    fn accent_table_inverts_the_rendering_table() {
        use crate::latex::operators::accent_to_latex;

        for accent in ACCENTS.values() {
            let rendered = accent_to_latex(*accent);
            let stripped = rendered.strip_prefix('\\').unwrap_or(rendered);
            assert_eq!(ACCENTS.get(stripped), Some(accent));
        }
    }

    #[test]
    fn space_table_inverts_the_rendering_table() {
        use crate::latex::operators::space_to_latex;

        for space in SPACES.values() {
            let rendered = space_to_latex(*space);
            let stripped = rendered.strip_prefix('\\').unwrap_or(rendered);
            assert_eq!(SPACES.get(stripped), Some(space));
        }
    }

    #[test]
    fn ascii_operators_cover_the_identity_glyphs() {
        assert_eq!(ascii_operator('+'), Some(Operator::Plus));
        assert_eq!(ascii_operator('='), Some(Operator::Equals));
        assert_eq!(ascii_operator('/'), None);
        assert_eq!(ascii_operator('*'), None);
    }
}
