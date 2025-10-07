// Character set handling and lookup tables
//
// Based on rtf2latex2e character mapping and lookup logic

use crate::formula::mtef::constants::*;

/// Character set attributes for typeface handling
#[derive(Debug, Clone)]
pub struct CharsetAttributes {
    pub math_attr: i32,        // Math attribute (0=text, 1=math, 2=force math, 3=force text)
    pub do_lookup: bool,       // Whether to do character lookup
    pub use_codepoint: bool,   // Whether to use codepoint as fallback
}

/// Character set information for each typeface slot
#[derive(Debug, Clone)]
pub struct CharsetInfo {
    pub attributes: CharsetAttributes,
    pub name: &'static str,
}

/// Typeface names (index 0-31, corresponding to typeface 129-160)
/// Based on rtf2latex2e typeFaceName array
#[allow(dead_code)]
const TYPEFACE_NAMES: &[&str] = &[
    "ZERO", "TEXT", "FUNCTION", "VARIABLE", "LCGREEK", "UCGREEK", "SYMBOL",
    "VECTOR", "NUMBER", "USER1", "USER2", "MTEXTRA", "UNKNOWN", "UNKNOWN",
    "UNKNOWN", "UNKNOWN", "UNKNOWN", "UNKNOWN", "UNKNOWN", "UNKNOWN",
    "UNKNOWN", "TEXT_FE", "EXPAND", "MARKER", "SPACE", "UNKNOWN", "UNKNOWN",
    "UNKNOWN", "UNKNOWN", "UNKNOWN", "UNKNOWN", "UNKNOWN"
];

/// Default character set attributes for each typeface slot
const DEFAULT_CHARSET_ATTRIBUTES: &[CharsetAttributes] = &[
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // ZERO
    CharsetAttributes { math_attr: 2, do_lookup: true, use_codepoint: true }, // TEXT
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // FUNCTION
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // VARIABLE
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // LCGREEK
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // UCGREEK
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SYMBOL
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SYMBOL2
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // MTEXTRA
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SPECIAL
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SPECIAL2
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SPECIAL3
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SPECIAL4
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SPECIAL5
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SPECIAL6
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SPECIAL7
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SPECIAL8
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SPECIAL9
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SPECIAL10
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SPECIAL11
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SPECIAL12
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SPECIAL13
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SPECIAL14
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SPECIAL15
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SPECIAL16
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SPECIAL17
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SPECIAL18
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SPECIAL19
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SPECIAL20
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SPECIAL21
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SPECIAL22
    CharsetAttributes { math_attr: 1, do_lookup: true, use_codepoint: true }, // SPECIAL23
];

/// Embellishment templates (math template,text template)
/// Each embellishment type has a template string with %1 as placeholder for the base character
pub const EMBELLISHMENT_TEMPLATES: &[&str] = &[
    "", // 0 - None
    "", // 1 - None
    "\\dot{%1} ,\\.%1 ", // 2 - embDOT
    "\\ddot{%1} ,\\\"%1 ", // 3 - embDDOT
    "\\dddot{%1} ,%1 ", // 4 - embTDOT
    "%1' ,%1 ", // 5 - embPRIME
    "%1'' ,%1 ", // 6 - embDPRIME
    "\\backprime %1 , %1", // 7 - embBPRIME
    "\\tilde{%1} ,\\~%1 ", // 8 - embTILDE
    "\\hat{%1} ,\\^%1 ", // 9 - embHAT
    "", // 10 - embNOT (empty in original)
    "\\vec{%1} ,%1 ", // 11 - embRARROW
    "\\overleftarrow1{%1} ,%1 ", // 12 - embLARROW
    "\\overleftrightarrow{%1} ,%1 ", // 13 - embBARROW
    "\\overrightarrow{%1} ,%1 ", // 14 - embR1ARROW
    "\\overleftarrow{%1} ,%1 ", // 15 - embL1ARROW
    "\\underline{%1} ,%1 ", // 16 - embMBAR
    "\\bar{%1} ,\\=%1 ", // 17 - embOBAR
    "%1''' ,", // 18 - embTPRIME
    "\\widehat{%1} ,%1 ", // 19 - embFROWN
    "\\breve{%1} ,%1 ", // 20 - embSMILE
    "{%1} ,%1 ", // 21 - embX_BARS
    "{%1} ,%1 ", // 22 - embUP_BAR
    "{%1} ,%1 ", // 23 - embDOWN_BAR
    "{%1} ,%1 ", // 24 - emb4DOT
    "\\d{%1} ,\\d{%1} ", // 25 - embU_1DOT
    "{%1} ,%1 ", // 26 - embU_2DOT
    "{%1} ,%1 ", // 27 - embU_3DOT
    "{%1} ,%1 ", // 28 - embU_4DOT
    "{%1} ,%1 ", // 29 - embU_BAR
    "{%1} ,%1 ", // 30 - embU_TILDE
    "{%1} ,%1 ", // 31 - embU_FROWN
    "{%1} ,%1 ", // 32 - embU_SMILE
    "{%1} ,%1 ", // 33 - embU_RARROW
    "{%1} ,%1 ", // 34 - embU_LARROW
    "{%1} ,%1 ", // 35 - embU_BARROW
    "{%1} ,%1 ", // 36 - embU_R1ARROW
    "{%1} ,%1 ", // 37 - embU_L1ARROW
];

/// Function name lookup table (based on rtf2latex2e Profile_FUNCTIONS)
static FUNCTION_LOOKUP_TABLE: phf::Map<&'static str, &'static str> = phf::phf_map! {
    "Pr" => "\\Pr ",
    "arccos" => "\\arccos ",
    "arcsin" => "\\arcsin ",
    "arctan" => "\\arctan ",
    "arg" => "\\arg ",
    "cos" => "\\cos ",
    "cosh" => "\\cosh ",
    "cot" => "\\cot ",
    "coth" => "\\coth ",
    "csc" => "\\csc ",
    "deg" => "\\deg ",
    "det" => "\\det ",
    "dim" => "\\dim ",
    "exp" => "\\exp ",
    "gcd" => "\\gcd ",
    "hom" => "\\hom ",
    "inf" => "\\inf ",
    "ker" => "\\ker ",
    "lim" => "\\lim ",
    "liminf" => "\\liminf ",
    "limsup" => "\\limsup ",
    "ln" => "\\ln ",
    "log" => "\\log ",
    "max" => "\\max ",
    "min" => "\\min ",
    "sec" => "\\sec ",
    "sin" => "\\sin ",
    "sinh" => "\\sinh ",
    "sup" => "\\sup ",
    "tan" => "\\tan ",
    "tanh" => "\\tanh ",
    "mod" => "\\mathop{\\rm mod} ",
    "glb" => "\\mathop{\\rm glb} ",
    "lub" => "\\mathop{\\rm lub} ",
    "int" => "\\mathop{\\rm int} ",
    "Im" => "\\mathop{\\rm Im} ",
    "Re" => "\\mathop{\\rm Re} ",
    "var" => "\\mathop{\\rm var} ",
    "cov" => "\\mathop{\\rm cov} ",
};

/// Perfect Hash Function map for character lookup (compile-time generated)
static CHAR_LOOKUP_TABLE: phf::Map<&'static str, &'static str> = phf::phf_map! {
    "130.91" => "\\lbrack ",
    "130.95" => "\\_ ",
    "131.95" => "\\_ ",
    "132.74" => "\\vartheta ",
    "132.86" => "\\varsigma ",
    "132.97" => "\\alpha ",
    "132.98" => "\\beta ",
    "132.99" => "\\chi ",
    "132.100" => "\\delta ",
    "132.101" => "\\epsilon ",
    "132.102" => "\\phi ",
    "132.103" => "\\gamma ",
    "132.104" => "\\eta ",
    "132.105" => "\\iota ",
    "132.106" => "\\varphi ",
    "132.107" => "\\kappa ",
    "132.108" => "\\lambda ",
    "132.109" => "\\mu ",
    "132.110" => "\\nu ",
    "132.111" => "\\o ",
    "132.112" => "\\pi ",
    "132.113" => "\\theta ",
    "132.114" => "\\rho ",
    "132.115" => "\\sigma ",
    "132.116" => "\\tau ",
    "132.117" => "\\upsilon ",
    "132.118" => "\\varpi ",
    "132.119" => "\\omega ",
    "132.120" => "\\xi ",
    "132.121" => "\\psi ",
    "132.122" => "\\zeta ",
    "132.182" => "\\partial ",
    "132.945" => "\\alpha ",
    "132.946" => "\\beta ",
    "132.967" => "\\chi ",
    "132.948" => "\\delta ",
    "132.949" => "\\epsilon ",
    "132.966" => "\\phi ",
    "132.947" => "\\gamma ",
    "132.951" => "\\eta ",
    "132.953" => "\\iota ",
    "132.981" => "\\varphi ",
    "132.954" => "\\kappa ",
    "132.955" => "\\lambda ",
    "132.956" => "\\mu ",
    "132.957" => "\\nu ",
    "132.959" => "\\o ",
    "132.960" => "\\pi ",
    "132.952" => "\\theta ",
    "132.961" => "\\rho ",
    "132.963" => "\\sigma ",
    "132.964" => "\\tau ",
    "132.965" => "\\upsilon ",
    "132.969" => "\\omega ",
    "132.958" => "\\xi ",
    "132.968" => "\\psi ",
    "132.950" => "\\zeta ",
    "132.977" => "\\vartheta ",
    "132.962" => "\\varsigma ",
    "132.982" => "\\varpi ",
    "133.65" => "A",
    "133.66" => "B",
    "133.67" => "X",
    "133.68" => "\\Delta ",
    "133.69" => "E",
    "133.70" => "\\Phi ",
    "133.71" => "\\Gamma ",
    "133.72" => "H",
    "133.73" => "I",
    "133.75" => "K",
    "133.76" => "\\Lambda ",
    "133.77" => "M",
    "133.78" => "N",
    "133.79" => "O",
    "133.80" => "\\Pi ",
    "133.81" => "\\Theta ",
    "133.82" => "P",
    "133.83" => "\\Sigma ",
    "133.84" => "T",
    "133.85" => "Y",
    "133.87" => "\\Omega ",
    "133.88" => "\\Xi ",
    "133.89" => "\\Psi ",
    "133.90" => "Z",
    "133.913" => "A",
    "133.914" => "B",
    "133.935" => "X",
    "133.916" => "\\Delta ",
    "133.917" => "E",
    "133.934" => "\\Phi ",
    "133.915" => "\\Gamma ",
    "133.919" => "H",
    "133.921" => "I",
    "133.922" => "K",
    "133.923" => "\\Lambda ",
    "133.924" => "M",
    "133.925" => "N",
    "133.927" => "O",
    "133.928" => "\\Pi ",
    "133.920" => "\\Theta ",
    "133.929" => "P",
    "133.931" => "\\Sigma ",
    "133.932" => "T",
    "133.933" => "Y",
    "133.937" => "\\Omega ",
    "133.926" => "\\Xi ",
    "133.936" => "\\Psi ",
    "133.918" => "Z",
    "134.34" => "\\forall ",
    "134.36" => "\\exists ",
    "134.39" => "\\ni ",
    "134.42" => "*",
    "134.43" => "+",
    "134.45" => "-",
    "134.61" => "=",
    "134.64" => "\\cong ",
    "134.92" => "\\therefore ",
    "134.94" => "\\bot ",
    "134.97" => "\\alpha ",
    "134.98" => "\\beta ",
    "134.99" => "\\chi ",
    "134.100" => "\\delta ",
    "134.101" => "\\epsilon ",
    "134.102" => "\\phi ",
    "134.103" => "\\gamma ",
    "134.104" => "\\eta ",
    "134.105" => "\\iota ",
    "134.106" => "\\varphi ",
    "134.107" => "\\kappa ",
    "134.108" => "\\lambda ",
    "134.109" => "\\mu ",
    "134.110" => "\\nu ",
    "134.112" => "\\pi ",
    "134.113" => "\\theta ",
    "134.114" => "\\rho ",
    "134.115" => "\\sigma ",
    "134.116" => "\\tau ",
    "134.117" => "\\upsilon ",
    "134.118" => "\\varpi ",
    "134.119" => "\\omega ",
    "134.120" => "\\xi ",
    "134.121" => "\\psi ",
    "134.122" => "\\zeta ",
    "134.163" => "\\leq ",
    "134.165" => "\\infty ",
    "134.171" => "\\leftrightarrow ",
    "134.172" => "\\leftarrow ",
    "134.173" => "\\uparrow ",
    "134.174" => "\\rightarrow ",
    "134.175" => "\\downarrow ",
    "134.176" => "^\\circ ",
    "134.177" => "\\pm ",
    "134.179" => "\\geq ",
    "134.180" => "\\times ",
    "134.181" => "\\propto ",
    "134.182" => "\\partial ",
    "134.183" => "\\bullet ",
    "134.184" => "\\div ",
    "134.185" => "\\neq ",
    "134.186" => "\\equiv ",
    "134.187" => "\\approx ",
    "134.191" => "\\hookleftarrow ",
    "134.192" => "\\aleph ",
    "134.193" => "\\Im ",
    "134.194" => "\\Re ",
    "134.195" => "\\wp ",
    "134.196" => "\\otimes ",
    "134.197" => "\\oplus ",
    "134.198" => "\\emptyset ",
    "134.199" => "\\cap ",
    "134.200" => "\\cup ",
    "134.201" => "\\supset ",
    "134.202" => "\\supseteq ",
    "134.203" => "\\nsubset ",
    "134.204" => "\\subset ",
    "134.205" => "\\subseteq ",
    "134.206" => "\\in ",
    "134.207" => "\\notin ",
    "134.208" => "\\angle ",
    "134.209" => "\\nabla ",
    "134.213" => "\\prod ",
    "134.215" => "\\cdot ",
    "134.216" => "\\neg ",
    "134.217" => "\\wedge ",
    "134.218" => "\\vee ",
    "134.219" => "\\Leftrightarrow ",
    "134.220" => "\\Leftarrow ",
    "134.221" => "\\Uparrow ",
    "134.222" => "\\Rightarrow ",
    "134.223" => "\\Downarrow ",
    "134.224" => "\\Diamond ",
    "134.225" => "\\langle ",
    "134.229" => "\\Sigma ",
    "134.241" => "\\rangle ",
    "134.242" => "\\smallint ",
    "134.247" => "\\div ",
    "134.8722" => "-",
    "134.8804" => "\\leq ",
    "134.8805" => "\\geq ",
    "134.8800" => "\\neq ",
    "134.8801" => "\\equiv ",
    "134.8776" => "\\approx ",
    "134.8773" => "\\cong ",
    "134.8733" => "\\propto ",
    "134.8727" => "\\ast ",
    "134.8901" => "\\cdot ",
    "134.8226" => "\\bullet ",
    "134.8855" => "\\otimes ",
    "134.8853" => "\\oplus ",
    "134.9001" => "\\langle ",
    "134.9002" => "\\rangle ",
    "134.8594" => "\\rightarrow ",
    "134.8592" => "\\leftarrow ",
    "134.8596" => "\\leftrightarrow ",
    "134.8593" => "\\uparrow ",
    "134.8595" => "\\downarrow ",
    "134.8658" => "\\Rightarrow ",
    "134.8656" => "\\Leftarrow ",
    "134.8660" => "\\Leftrightarrow ",
    "134.8657" => "\\Uparrow ",
    "134.8659" => "\\Downarrow ",
    "134.8629" => "\\hookleftarrow ",
    "134.8756" => "\\therefore ",
    "134.8717" => "\\backepsilon ",
    "134.8707" => "\\exists ",
    "134.8704" => "\\forall ",
    "134.8743" => "\\wedge ",
    "134.8744" => "\\vee ",
    "134.8712" => "\\in ",
    "134.8713" => "\\notin ",
    "134.8746" => "\\cup ",
    "134.8745" => "\\cap ",
    "134.8834" => "\\subset ",
    "134.8835" => "\\supset ",
    "134.8838" => "\\subseteq ",
    "134.8839" => "\\supseteq ",
    "134.8836" => "\\not\\subset ",
    "134.8709" => "\\emptyset ",
    "134.8706" => "\\partial ",
    "134.8711" => "\\nabla ",
    "134.8465" => "\\Im ",
    "134.8476" => "\\Re ",
    "134.8501" => "\\aleph ",
    "134.8736" => "\\angle ",
    "134.8869" => "\\bot ",
    "134.8900" => "\\lozenge ",
    "134.8734" => "\\infty ",
    "134.8472" => "\\wp ",
    "134.8747" => "\\smallint",
    "134.8721" => "\\sum ",
    "134.8719" => "\\prod ",
    "139.58" => "\\sim ",
    "139.59" => "\\simeq ",
    "139.60" => "\\vartriangleleft ",
    "139.61" => "\\ll ",
    "139.62" => "\\vartriangleright ",
    "139.63" => "\\gg ",
    "139.66" => "\\doteq ",
    "139.67" => "\\coprod ",
    "139.68" => "\\lambdabar ",
    "139.73" => "\\bigcap ",
    "139.75" => "\\ldots ",
    "139.76" => "\\cdots ",
    "139.77" => "\\vdots ",
    "139.78" => "\\ddots ",
    "139.79" => "\\ddots ",
    "139.81" => "\\because ",
    "139.85" => "\\bigcup ",
    "139.97" => "\\mapsto ",
    "139.98" => "\\updownarrow ",
    "139.99" => "\\Updownarrow ",
    "139.102" => "\\succ ",
    "139.104" => "\\hbar ",
    "139.108" => "\\ell ",
    "139.109" => "\\mp ",
    "139.111" => "\\circ ",
    "139.112" => "\\prec ",
    "139.8230" => "\\ldots ",
    "139.8943" => "\\cdots ",
    "139.8942" => "\\vdots ",
    "139.8944" => "\\ddots ",
    "139.8945" => "\\ddots ",
    "139.8826" => "\\prec ",
    "139.8827" => "\\succ ",
    "139.8882" => "\\vartriangleleft ",
    "139.8883" => "\\vartriangleright ",
    "139.8723" => "\\mp ",
    "139.8728" => "\\circ ",
    "139.8614" => "\\longmapsto ",
    "139.8597" => "\\updownarrow ",
    "139.8661" => "\\Updownarrow ",
    "139.4746" => "\\bigcup ",
    "139.4745" => "\\bigcap ",
    "139.8757" => "\\because ",
    "139.8467" => "\\ell ",
    "139.8463" => "\\hbar ",
    "139.411" => "\\lambdabar ",
    "139.8720" => "\\coprod ",
    "151.60160" => "{}",
    "152.1" => "{}",
    "152.8" => "\\/",
    "152.2" => "\\,",
    "152.4" => "\\;",
    "152.5" => "\\quad ",
    "152.60161" => "\\/",
    "152.61168" => "@,",
    "152.60162" => "\\,",
    "152.60164" => "\\;",
    "152.60165" => "\\quad ",
    "152.61186" => "\\, ",
};

/// Get character set attributes for a given typeface index
pub fn get_charset_attributes(charset_index: usize) -> CharsetAttributes {
    if charset_index < DEFAULT_CHARSET_ATTRIBUTES.len() {
        DEFAULT_CHARSET_ATTRIBUTES[charset_index].clone()
    } else {
        // Fallback for unknown typefaces
        CharsetAttributes {
            math_attr: MA_FORCE_MATH, // Default to math mode
            do_lookup: true,          // Enable character lookup
            use_codepoint: true,      // Enable codepoint fallback
        }
    }
}

/// Lookup a character in the character mapping table
pub fn lookup_character(typeface: usize, character: u16, math_attr: i32) -> Option<&'static str> {
    // Try primary lookup first
    let key = format!("{}.{}", typeface, character);
    if let Some(result) = CHAR_LOOKUP_TABLE.get(&key) {
        return Some(*result);
    }

    // For special cases with math attribute variations (like spaces)
    // Following rtf2latex2e logic for mode-dependent character lookup
    if typeface == 152 && math_attr == crate::formula::mtef::constants::MA_TEXT {
        let text_key = format!("{}.{}{}", typeface, character, 't');
        if let Some(result) = CHAR_LOOKUP_TABLE.get(&text_key) {
            return Some(*result);
        }
    } else if typeface == 152 && math_attr == crate::formula::mtef::constants::MA_MATH {
        let math_key = format!("{}.{}{}", typeface, character, 'm');
        if let Some(result) = CHAR_LOOKUP_TABLE.get(&math_key) {
            return Some(*result);
        }
    }

    None
}

/// Lookup a function name in the function mapping table
pub fn lookup_function(function_name: &str) -> Option<&'static str> {
    FUNCTION_LOOKUP_TABLE.get(function_name).map(|v| *v)
}

/// Get embellishment template for a given embellishment type
pub fn get_embellishment_template(embell: u8) -> &'static str {
    if (embell as usize) < EMBELLISHMENT_TEMPLATES.len() {
        EMBELLISHMENT_TEMPLATES[embell as usize]
    } else {
        ""
    }
}
