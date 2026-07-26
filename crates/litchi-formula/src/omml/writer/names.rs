// OMML element and attribute names (ECMA-376 Part 1, §22.1)
//
// All names carry the conventional `m:` prefix used by the OOXML math
// namespace. The namespace URI is declared on the root `m:oMath` element so
// serialized fragments are self-contained.

/// OOXML math namespace URI
pub const OMML_NAMESPACE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";
/// Namespace declaration attribute name
pub const ATTR_XMLNS_M: &str = "xmlns:m";
/// Value attribute used by OMML property elements
pub const ATTR_VAL: &str = "m:val";
/// Run color attribute (parser-compatible shorthand on `m:rPr`)
pub const ATTR_COLOR: &str = "m:color";

// Root and runs
pub const EL_OMATH: &str = "m:oMath";
pub const EL_RUN: &str = "m:r";
pub const EL_TEXT: &str = "m:t";
pub const EL_RUN_PROPS: &str = "m:rPr";
pub const EL_LITERAL: &str = "m:lit";
pub const EL_SCRIPT_STYLE: &str = "m:scr";
pub const EL_STYLE: &str = "m:sty";
pub const EL_NORMAL_TEXT: &str = "m:nor";

// Fractions
pub const EL_FRACTION: &str = "m:f";
pub const EL_FRACTION_PROPS: &str = "m:fPr";
pub const EL_FRACTION_TYPE: &str = "m:type";
pub const EL_NUMERATOR: &str = "m:num";
pub const EL_DENOMINATOR: &str = "m:den";

// Radicals
pub const EL_RADICAL: &str = "m:rad";
pub const EL_RADICAL_PROPS: &str = "m:radPr";
pub const EL_DEGREE_HIDE: &str = "m:degHide";
pub const EL_DEGREE: &str = "m:deg";

// Scripts
pub const EL_SUPERSCRIPT: &str = "m:sSup";
pub const EL_SUBSCRIPT: &str = "m:sSub";
pub const EL_SUB_SUP: &str = "m:sSubSup";
pub const EL_PRE_SCRIPT: &str = "m:sPre";
pub const EL_SUB: &str = "m:sub";
pub const EL_SUP: &str = "m:sup";

// Delimiters
pub const EL_DELIMITER: &str = "m:d";
pub const EL_DELIMITER_PROPS: &str = "m:dPr";
pub const EL_BEGIN_CHAR: &str = "m:begChr";
pub const EL_END_CHAR: &str = "m:endChr";
pub const EL_SEPARATOR_CHAR: &str = "m:sepChr";

// N-ary operators
pub const EL_NARY: &str = "m:nary";
pub const EL_NARY_PROPS: &str = "m:naryPr";
pub const EL_CHAR: &str = "m:chr";
pub const EL_SUB_HIDE: &str = "m:subHide";
pub const EL_SUP_HIDE: &str = "m:supHide";

// Functions
pub const EL_FUNCTION: &str = "m:func";
pub const EL_FUNCTION_NAME: &str = "m:fName";

// Element (base/argument) container
pub const EL_ELEMENT: &str = "m:e";

// Matrices
pub const EL_MATRIX: &str = "m:m";
pub const EL_MATRIX_PROPS: &str = "m:mPr";
pub const EL_MATRIX_ROW: &str = "m:mr";
pub const EL_BASE_JC: &str = "m:baseJc";
pub const EL_ROW_SPACING: &str = "m:rSp";

// Equation arrays
pub const EL_EQ_ARRAY: &str = "m:eqArr";
pub const EL_EQ_ARRAY_PROPS: &str = "m:eqArrPr";
pub const EL_MAX_DIST: &str = "m:maxDist";
pub const EL_OBJ_DIST: &str = "m:objDist";
pub const EL_ROW_SPACING_RULE: &str = "m:rSpRule";

// Accents and decorations
pub const EL_ACCENT: &str = "m:acc";
pub const EL_ACCENT_PROPS: &str = "m:accPr";
pub const EL_BAR: &str = "m:bar";
pub const EL_BAR_PROPS: &str = "m:barPr";
pub const EL_POSITION: &str = "m:pos";
pub const EL_BOX: &str = "m:box";
pub const EL_BORDER_BOX: &str = "m:borderBox";
pub const EL_PHANTOM: &str = "m:phant";
pub const EL_GROUP_CHAR: &str = "m:groupChr";
pub const EL_GROUP_CHAR_PROPS: &str = "m:groupChrPr";
pub const EL_VERT_JC: &str = "m:vertJc";

// Border box properties
pub const EL_BORDER_BOX_PROPS: &str = "m:borderBoxPr";
pub const EL_HIDE_TOP: &str = "m:hideTop";
pub const EL_HIDE_BOT: &str = "m:hideBot";
pub const EL_HIDE_LEFT: &str = "m:hideLeft";
pub const EL_HIDE_RIGHT: &str = "m:hideRight";
pub const EL_STRIKE_H: &str = "m:strikeH";
pub const EL_STRIKE_V: &str = "m:strikeV";
pub const EL_STRIKE_BLTR: &str = "m:strikeBLTR";
pub const EL_STRIKE_TLBR: &str = "m:strikeTLBR";

// Limits
pub const EL_LIM_LOW: &str = "m:limLow";
pub const EL_LIM_UPP: &str = "m:limUpp";
pub const EL_LIMIT: &str = "m:lim";

/// Boolean "on" value used by OMML on/off properties
pub const VAL_ON: &str = "1";
/// Boolean "off" value used by OMML on/off properties
pub const VAL_OFF: &str = "0";
