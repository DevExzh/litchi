//! Element names and their rules: the static alias, case-folding and
//! predicate tables of record 0010 §4.2–§4.3, harvested from OfficeCLI.
//!
//! One [`Vocabulary`] per format family names the elements of every format in
//! it: [`WORD`] serves DOCX, DOC, RTF and ODT; [`PRESENTATION`] serves PPTX,
//! PPT and ODP; [`WORKBOOK`] serves XLSX, XLSB, XLS and ODS (record 0010 §4.8:
//! "tier 2 uses the same paths and kinds as its OOXML sibling"). A format
//! adapter maps the canonical names it models to its own element kinds
//! (`ElementKind::from_name`); the names it does not model stay known
//! concepts, which is what lets the resolver tell `unsupported_capability`
//! from an unknown name.
//!
//! Every table is a `static` sorted array, checked at compile time. Lookups
//! binary-search with an inline ASCII case-insensitive comparison; they never
//! allocate.
//!
//! # Sources
//!
//! Path segment aliases come from OfficeCLI's `Core/PathAliases.cs`, the Word
//! query dispatch (`WordHandler.Query.cs`, element-type tests), the
//! PowerPoint path helpers and query dispatch
//! (`PowerPointHandler.Helpers.Path.cs`, `PowerPointHandler.Query.cs`) and the
//! Excel query dispatch (`ExcelHandler.Query.cs`, `KnownElementTypes` and the
//! `merge` rewrite). Canonical names are the segments OfficeCLI's node
//! builders emit (`p`/`r` in Word, `paragraph`/`run` in presentations,
//! `col[B]` in workbooks). Element names that OfficeCLI's handlers or
//! `schemas/help` accept beyond the plan's §4.2 table (Word `tab`, `ptab`,
//! `fieldChar`; presentation `zoom`, `model3d`; workbook `slicer`,
//! `sparkline`, ...) are included as canonical names so that they classify as
//! known concepts.

use caseless::Caseless as _;
use std::cmp::Ordering;
use unicode_normalization::UnicodeNormalization;

use crate::Family;

/// How a value in a path is compared with the value the document holds.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(crate) enum Case {
    /// Byte for byte: Word bookmark and form-field names, style ids, numeric
    /// ids.
    Exact,
    /// ASCII letters fold: hexadecimal `w14:paraId`, placeholder and
    /// header/footer type tokens, chart axis roles.
    Ascii,
    /// Canonical Unicode caseless comparison (NFD, full fold, NFD): worksheet names,
    /// defined names and shape names (the adopted canonical identity policy).
    ///
    /// Workbook adapters whose leaf has its own canonical name key (XLSX
    /// `sheet::key`, NFD plus default case folding) use that key instead, as
    /// record 0010 §4.2 requires for sheet names.
    Caseless,
}

impl Case {
    /// Whether `held` (from the document) matches `wanted` (from the path).
    /// ASCII matching is allocation-free; normalization may buffer combining marks.
    pub(crate) fn matches(self, held: &str, wanted: &str) -> bool {
        match self {
            Self::Exact => held == wanted,
            Self::Ascii => held.eq_ignore_ascii_case(wanted),
            Self::Caseless => held
                .nfd()
                .default_case_fold()
                .nfd()
                .eq(wanted.nfd().default_case_fold().nfd()),
        }
    }
}

/// A stable-identity key an element accepts as `[@key=value]` (§4.3), with
/// the rule its values compare by. Keys themselves match ASCII
/// case-insensitively (`@PARAID` is `@paraId`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(crate) struct Identity {
    /// The key as OfficeCLI spells it, without the `@`.
    pub(crate) key: &'static str,
    /// How the key's values compare.
    pub(crate) case: Case,
}

impl Identity {
    pub(crate) fn is_native(self) -> bool {
        !matches!(self.key, "name" | "role" | "type")
    }

    const fn new(key: &'static str, case: Case) -> Self {
        Self { key, case }
    }
}

/// What a bare token in brackets means for an element (§4.3): `col[B]`,
/// `namedrange[Total]`, `placeholder[ctrTitle]`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(crate) enum Bare {
    /// Column letters (`col[B]`), read as the 1-based column number.
    Column,
    /// The element's name, compared by the given rule (`namedrange[NAME]`,
    /// `formfield[name]`).
    Name(Case),
    /// A type token (`placeholder[ctrTitle]`).
    Token,
}

/// What a 1-based index `[N]` means for an element (§4.3).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(crate) enum IndexRule {
    /// The Nth child of that kind under the parent.
    Ordinal,
    /// A number the element carries: numbering `level[L]` is `w:ilvl` = L.
    Number,
    /// A number the element carries when the parent is the document root,
    /// an ordinal otherwise: Word `/footnote[N]` is note id N, while
    /// `/footnotes/footnote[N]` is the Nth note.
    NumberUnderRoot,
    /// A worksheet row number, `1..=1_048_576`.
    Row,
    /// A worksheet column number, `1..=16_384`.
    Column,
}

/// Children that a step names by free text rather than by element kind
/// (§4.2).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(crate) enum Named {
    /// Steps under this element name element kinds.
    No,
    /// A step without predicates is first the name of a `child` element (a
    /// worksheet at a workbook's root); only when no child has that name is
    /// it read as an element kind, so a sheet named `Chart` stays reachable
    /// as `/Chart/A1`.
    First {
        /// The canonical name of the named child's element.
        child: &'static str,
        /// How names compare.
        case: Case,
    },
    /// Every step under this element is the name of a `child` element (a
    /// style id under Word's `/styles`); it is never read as an element kind.
    Only {
        /// The canonical name of the named child's element.
        child: &'static str,
        /// How names compare.
        case: Case,
    },
}

/// How paths address one element: its identity keys, bare tokens, index
/// meaning, and named or cell-addressed children.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(crate) struct ElementRule {
    /// The canonical name: the path segment OfficeCLI emits.
    pub(crate) name: &'static str,
    /// The `[@key=value]` identities the element accepts.
    pub(crate) identities: &'static [Identity],
    /// What a bare token means, if anything.
    pub(crate) bare: Option<Bare>,
    /// What `[N]` means.
    pub(crate) index: IndexRule,
    /// Children addressed by name.
    pub(crate) named: Named,
    /// Whether a step under this element may be an A1 cell or range
    /// (`/Sheet1/A1`, `/Sheet1/A1:C5`).
    pub(crate) cells: bool,
}

impl ElementRule {
    const fn new(name: &'static str) -> Self {
        Self {
            name,
            identities: &[],
            bare: None,
            index: IndexRule::Ordinal,
            named: Named::No,
            cells: false,
        }
    }

    const fn ids(self, identities: &'static [Identity]) -> Self {
        Self { identities, ..self }
    }

    const fn bare(self, bare: Bare) -> Self {
        Self {
            bare: Some(bare),
            ..self
        }
    }

    const fn index(self, index: IndexRule) -> Self {
        Self { index, ..self }
    }

    const fn named(self, named: Named) -> Self {
        Self { named, ..self }
    }

    const fn cells(self) -> Self {
        Self {
            cells: true,
            ..self
        }
    }

    /// The identity this element accepts for `key` (ASCII
    /// case-insensitive).
    pub(crate) fn identity(&self, key: &str) -> Option<Identity> {
        self.identities
            .iter()
            .find(|identity| identity.key.eq_ignore_ascii_case(key))
            .copied()
    }
}

/// The element vocabulary of one format family.
#[derive(Debug)]
pub(crate) struct Vocabulary {
    /// The family whose formats this vocabulary names.
    pub(crate) family: Family,
    /// The canonical name of the document root, which no step names.
    pub(crate) root: &'static str,
    /// `(alias, canonical name)`, sorted by alias. Aliases are lowercase
    /// ASCII; every canonical name except the root is its own alias.
    pub(crate) aliases: &'static [(&'static str, &'static str)],
    /// Element rules, sorted ASCII case-insensitively by name.
    pub(crate) elements: &'static [ElementRule],
    /// The family's top-level path forms, for `unsupported_path` suggestions.
    pub(crate) roots: &'static str,
}

impl Vocabulary {
    /// The canonical name `name` stands for, matched ASCII
    /// case-insensitively. Allocation-free.
    pub(crate) fn canonical(&self, name: &str) -> Option<&'static str> {
        // Aliases are lowercase, so folding both sides orders them as stored.
        self.aliases
            .binary_search_by(|(alias, _)| cmp_ascii_ci(alias.as_bytes(), name.as_bytes()))
            .ok()
            .map(|index| self.aliases[index].1)
    }

    /// The rule of a canonical element name (ASCII case-insensitive).
    pub(crate) fn rule(&self, canonical: &str) -> Option<&'static ElementRule> {
        self.index(canonical).map(|index| &self.elements[index])
    }

    /// The position of a canonical element name in [`Self::elements`]
    /// (ASCII case-insensitive): a dense index an adapter may use for its
    /// kinds. Allocation-free.
    pub(crate) fn index(&self, canonical: &str) -> Option<usize> {
        self.elements
            .binary_search_by(|rule| cmp_ascii_ci(rule.name.as_bytes(), canonical.as_bytes()))
            .ok()
    }

    /// Every spelling a step may use, aliases included.
    pub(crate) fn spellings(&self) -> impl Iterator<Item = (&'static str, &'static str)> {
        self.aliases.iter().copied()
    }
}

/// Orders byte strings with ASCII letters folded to lowercase. Inline
/// byte comparison: for these short names it is several times faster than a
/// `memcmp` call per probe.
const fn cmp_ascii_ci(left: &[u8], right: &[u8]) -> Ordering {
    let mut index = 0;
    while index < left.len() && index < right.len() {
        let (a, b) = (
            left[index].to_ascii_lowercase(),
            right[index].to_ascii_lowercase(),
        );
        if a < b {
            return Ordering::Less;
        }
        if a > b {
            return Ordering::Greater;
        }
        index += 1;
    }
    if left.len() < right.len() {
        Ordering::Less
    } else if left.len() > right.len() {
        Ordering::Greater
    } else {
        Ordering::Equal
    }
}

/// The vocabulary of another family that knows `name`: the concept exists in
/// Office documents, just not in `own`'s family (`/slide[1]` on a Word
/// document).
pub(crate) fn elsewhere(own: &Vocabulary, name: &str) -> Option<&'static Vocabulary> {
    FAMILIES
        .into_iter()
        .find(|vocabulary| !std::ptr::eq(*vocabulary, own) && vocabulary.canonical(name).is_some())
}

/// The three family vocabularies.
pub(crate) const FAMILIES: [&Vocabulary; 3] = [&WORD, &PRESENTATION, &WORKBOOK];

/// Compile-time checks of one vocabulary: aliases sorted, unique and
/// lowercase ASCII; elements sorted and unique;
/// every alias names an element; the root is an element no alias names;
/// named children are elements.
const fn well_formed(vocabulary: &Vocabulary) -> bool {
    let aliases = vocabulary.aliases;
    let mut index = 0;
    while index < aliases.len() {
        let alias = aliases[index].0.as_bytes();
        if alias.is_empty() {
            return false;
        }
        let mut byte = 0;
        while byte < alias.len() {
            if !alias[byte].is_ascii() || alias[byte].is_ascii_uppercase() {
                return false;
            }
            byte += 1;
        }
        if index > 0
            && !matches!(
                cmp_ascii_ci(aliases[index - 1].0.as_bytes(), alias),
                Ordering::Less
            )
        {
            return false;
        }
        let target = aliases[index].1;
        if !has_element(vocabulary, target) || equal(target, vocabulary.root) {
            return false;
        }
        index += 1;
    }
    let elements = vocabulary.elements;
    let mut index = 0;
    while index < elements.len() {
        if index > 0
            && !matches!(
                cmp_ascii_ci(
                    elements[index - 1].name.as_bytes(),
                    elements[index].name.as_bytes()
                ),
                Ordering::Less
            )
        {
            return false;
        }
        match elements[index].named {
            Named::No => {},
            Named::First { child, .. } | Named::Only { child, .. } => {
                if !has_element(vocabulary, child) {
                    return false;
                }
            },
        }
        index += 1;
    }
    has_element(vocabulary, vocabulary.root)
}

const fn has_element(vocabulary: &Vocabulary, name: &str) -> bool {
    let elements = vocabulary.elements;
    let mut index = 0;
    while index < elements.len() {
        if equal(elements[index].name, name) {
            return true;
        }
        index += 1;
    }
    false
}

const fn equal(left: &str, right: &str) -> bool {
    matches!(
        cmp_ascii_ci(left.as_bytes(), right.as_bytes()),
        Ordering::Equal
    )
}

/// Whether every element except the root has its canonical name as an alias.
const fn names_are_aliases(vocabulary: &Vocabulary) -> bool {
    let elements = vocabulary.elements;
    let mut index = 0;
    while index < elements.len() {
        let name = elements[index].name;
        if !equal(name, vocabulary.root) {
            let aliases = vocabulary.aliases;
            let mut found = false;
            let mut alias = 0;
            while alias < aliases.len() {
                if equal(aliases[alias].0, name) && equal(aliases[alias].1, name) {
                    found = true;
                }
                alias += 1;
            }
            if !found {
                return false;
            }
        }
        index += 1;
    }
    true
}

const _: () = {
    assert!(well_formed(&WORD) && names_are_aliases(&WORD));
    assert!(well_formed(&PRESENTATION) && names_are_aliases(&PRESENTATION));
    assert!(well_formed(&WORKBOOK) && names_are_aliases(&WORKBOOK));
    // `KindSet` is a 128-bit set over a format's kinds.
    assert!(WORD.elements.len() <= 128);
    assert!(PRESENTATION.elements.len() <= 128);
    assert!(WORKBOOK.elements.len() <= 128);
};

use Case::{Ascii, Caseless, Exact};

/// Word-processing documents: DOCX, DOC, RTF, ODT.
pub(crate) static WORD: Vocabulary = Vocabulary {
    family: Family::Document,
    root: "document",
    aliases: &[
        ("abstractnum", "abstractNum"),
        ("axis", "axis"),
        ("body", "body"),
        ("bookmark", "bookmark"),
        ("bookmarkend", "bookmarkEnd"),
        ("br", "break"),
        ("break", "break"),
        ("cell", "tc"),
        ("chart", "chart"),
        ("col", "col"),
        ("column", "col"),
        ("comment", "comment"),
        ("comments", "comments"),
        ("contentcontrol", "sdt"),
        ("customxml", "customXml"),
        ("embed", "ole"),
        ("endnote", "endnote"),
        ("endnotes", "endnotes"),
        ("equation", "equation"),
        ("field", "field"),
        ("fieldchar", "fieldChar"),
        ("fldchar", "fieldChar"),
        ("footer", "footer"),
        ("footnote", "footnote"),
        ("footnotes", "footnotes"),
        ("formfield", "formfield"),
        ("formula", "equation"),
        ("group", "group"),
        ("header", "header"),
        ("hyperlink", "hyperlink"),
        ("image", "picture"),
        ("img", "picture"),
        ("instrtext", "instrText"),
        ("level", "level"),
        ("lvl", "level"),
        ("math", "equation"),
        ("num", "num"),
        ("numbering", "numbering"),
        ("object", "ole"),
        ("ole", "ole"),
        ("oleobject", "ole"),
        ("omath", "equation"),
        ("omathpara", "equation"),
        ("p", "p"),
        ("pagebreak", "break"),
        ("paragraph", "p"),
        ("permstart", "permStart"),
        ("picture", "picture"),
        ("positionaltab", "ptab"),
        ("ptab", "ptab"),
        ("r", "r"),
        ("revision", "revision"),
        ("row", "tr"),
        ("run", "r"),
        ("sdt", "sdt"),
        ("section", "section"),
        ("series", "series"),
        ("settings", "settings"),
        ("shape", "shape"),
        ("style", "style"),
        ("styles", "styles"),
        ("tab", "tab"),
        ("table", "tbl"),
        ("tableofcontents", "toc"),
        ("tbl", "tbl"),
        ("tc", "tc"),
        ("textbox", "textbox"),
        ("theme", "theme"),
        ("toc", "toc"),
        ("tr", "tr"),
        ("watermark", "watermark"),
    ],
    elements: &[
        ElementRule::new("abstractNum").ids(&[Identity::new("id", Exact)]),
        ElementRule::new("axis").ids(&[Identity::new("role", Ascii)]),
        ElementRule::new("body"),
        ElementRule::new("bookmark").ids(&[Identity::new("name", Exact)]),
        ElementRule::new("bookmarkEnd").ids(&[Identity::new("id", Exact)]),
        ElementRule::new("break"),
        ElementRule::new("chart"),
        ElementRule::new("col"),
        ElementRule::new("comment").ids(&[Identity::new("commentId", Exact)]),
        ElementRule::new("comments"),
        ElementRule::new("customXml"),
        ElementRule::new("document"),
        ElementRule::new("endnote")
            .ids(&[Identity::new("endnoteId", Exact)])
            .index(IndexRule::NumberUnderRoot),
        ElementRule::new("endnotes"),
        ElementRule::new("equation"),
        ElementRule::new("field"),
        ElementRule::new("fieldChar"),
        ElementRule::new("footer").ids(&[Identity::new("type", Ascii)]),
        ElementRule::new("footnote")
            .ids(&[Identity::new("footnoteId", Exact)])
            .index(IndexRule::NumberUnderRoot),
        ElementRule::new("footnotes"),
        ElementRule::new("formfield")
            .ids(&[Identity::new("name", Exact)])
            .bare(Bare::Name(Exact)),
        ElementRule::new("group"),
        ElementRule::new("header").ids(&[Identity::new("type", Ascii)]),
        ElementRule::new("hyperlink"),
        ElementRule::new("instrText"),
        ElementRule::new("level")
            .ids(&[Identity::new("ilvl", Exact)])
            .index(IndexRule::Number),
        ElementRule::new("num").ids(&[Identity::new("id", Exact)]),
        ElementRule::new("numbering"),
        ElementRule::new("ole"),
        ElementRule::new("p").ids(&[
            Identity::new("paraId", Ascii),
            Identity::new("textId", Ascii),
        ]),
        ElementRule::new("permStart").ids(&[Identity::new("id", Exact)]),
        ElementRule::new("picture"),
        ElementRule::new("ptab"),
        ElementRule::new("r"),
        ElementRule::new("revision")
            .ids(&[Identity::new("id", Exact), Identity::new("type", Ascii)]),
        ElementRule::new("sdt").ids(&[Identity::new("sdtId", Exact)]),
        ElementRule::new("section"),
        ElementRule::new("series"),
        ElementRule::new("settings"),
        ElementRule::new("shape"),
        ElementRule::new("style"),
        ElementRule::new("styles").named(Named::Only {
            child: "style",
            case: Exact,
        }),
        ElementRule::new("tab"),
        ElementRule::new("tbl"),
        ElementRule::new("tc"),
        ElementRule::new("textbox"),
        ElementRule::new("theme"),
        ElementRule::new("toc"),
        ElementRule::new("tr"),
        ElementRule::new("watermark"),
    ],
    roots: "/body, /header[N], /footer[N], /footnote[N], /endnote[N], \
            /comments/comment[N], /styles/<style id>, /numbering or /section[N]",
};

/// OpenDocument text: the Word family names plus the elements ODF has as
/// their own concepts (`h` headings, `list`/`listItem`, `frame`, the
/// `/metadata` node), with `annotation` and `span` as aliases of `comment`
/// and `r`. ODF paragraphs carry no OOXML wire identity.
#[cfg(feature = "odt")]
pub(crate) static ODT: Vocabulary = Vocabulary {
    family: WORD.family,
    root: WORD.root,
    aliases: &[
        ("abstractnum", "abstractNum"),
        ("annotation", "comment"),
        ("axis", "axis"),
        ("body", "body"),
        ("bookmark", "bookmark"),
        ("bookmarkend", "bookmarkEnd"),
        ("br", "break"),
        ("break", "break"),
        ("cell", "tc"),
        ("chart", "chart"),
        ("col", "col"),
        ("column", "col"),
        ("comment", "comment"),
        ("comments", "comments"),
        ("contentcontrol", "sdt"),
        ("customxml", "customXml"),
        ("embed", "ole"),
        ("endnote", "endnote"),
        ("endnotes", "endnotes"),
        ("equation", "equation"),
        ("field", "field"),
        ("fieldchar", "fieldChar"),
        ("fldchar", "fieldChar"),
        ("footer", "footer"),
        ("footnote", "footnote"),
        ("footnotes", "footnotes"),
        ("formfield", "formfield"),
        ("formula", "equation"),
        ("frame", "frame"),
        ("group", "group"),
        ("h", "h"),
        ("header", "header"),
        ("heading", "h"),
        ("hyperlink", "hyperlink"),
        ("image", "picture"),
        ("img", "picture"),
        ("instrtext", "instrText"),
        ("level", "level"),
        ("li", "listItem"),
        ("list", "list"),
        ("listitem", "listItem"),
        ("lvl", "level"),
        ("math", "equation"),
        ("metadata", "metadata"),
        ("num", "num"),
        ("numbering", "numbering"),
        ("object", "ole"),
        ("ole", "ole"),
        ("oleobject", "ole"),
        ("omath", "equation"),
        ("omathpara", "equation"),
        ("p", "p"),
        ("pagebreak", "break"),
        ("paragraph", "p"),
        ("permstart", "permStart"),
        ("picture", "picture"),
        ("positionaltab", "ptab"),
        ("ptab", "ptab"),
        ("r", "r"),
        ("revision", "revision"),
        ("row", "tr"),
        ("run", "r"),
        ("sdt", "sdt"),
        ("section", "section"),
        ("series", "series"),
        ("settings", "settings"),
        ("shape", "shape"),
        ("span", "r"),
        ("style", "style"),
        ("styles", "styles"),
        ("tab", "tab"),
        ("table", "tbl"),
        ("tableofcontents", "toc"),
        ("tbl", "tbl"),
        ("tc", "tc"),
        ("textbox", "textbox"),
        ("theme", "theme"),
        ("toc", "toc"),
        ("tr", "tr"),
        ("watermark", "watermark"),
    ],
    elements: &[
        ElementRule::new("abstractNum").ids(&[Identity::new("id", Exact)]),
        ElementRule::new("axis").ids(&[Identity::new("role", Ascii)]),
        ElementRule::new("body"),
        ElementRule::new("bookmark").ids(&[Identity::new("name", Exact)]),
        ElementRule::new("bookmarkEnd").ids(&[Identity::new("id", Exact)]),
        ElementRule::new("break"),
        ElementRule::new("chart"),
        ElementRule::new("col"),
        ElementRule::new("comment"),
        ElementRule::new("comments"),
        ElementRule::new("customXml"),
        ElementRule::new("document"),
        ElementRule::new("endnote"),
        ElementRule::new("endnotes"),
        ElementRule::new("equation"),
        ElementRule::new("field"),
        ElementRule::new("fieldChar"),
        ElementRule::new("footer").ids(&[Identity::new("type", Ascii)]),
        ElementRule::new("footnote"),
        ElementRule::new("footnotes"),
        ElementRule::new("formfield")
            .ids(&[Identity::new("name", Exact)])
            .bare(Bare::Name(Exact)),
        ElementRule::new("frame"),
        ElementRule::new("group"),
        ElementRule::new("h"),
        ElementRule::new("header").ids(&[Identity::new("type", Ascii)]),
        ElementRule::new("hyperlink"),
        ElementRule::new("instrText"),
        ElementRule::new("level")
            .ids(&[Identity::new("ilvl", Exact)])
            .index(IndexRule::Number),
        ElementRule::new("list"),
        ElementRule::new("listItem"),
        ElementRule::new("metadata"),
        ElementRule::new("num").ids(&[Identity::new("id", Exact)]),
        ElementRule::new("numbering"),
        ElementRule::new("ole"),
        ElementRule::new("p"),
        ElementRule::new("permStart").ids(&[Identity::new("id", Exact)]),
        ElementRule::new("picture"),
        ElementRule::new("ptab"),
        ElementRule::new("r"),
        ElementRule::new("revision")
            .ids(&[Identity::new("id", Exact), Identity::new("type", Ascii)]),
        ElementRule::new("sdt").ids(&[Identity::new("sdtId", Exact)]),
        ElementRule::new("section"),
        ElementRule::new("series"),
        ElementRule::new("settings"),
        ElementRule::new("shape"),
        ElementRule::new("style"),
        ElementRule::new("styles").named(Named::Only {
            child: "style",
            case: Exact,
        }),
        ElementRule::new("tab"),
        ElementRule::new("tbl"),
        ElementRule::new("tc"),
        ElementRule::new("textbox"),
        ElementRule::new("theme"),
        ElementRule::new("toc"),
        ElementRule::new("tr"),
        ElementRule::new("watermark"),
    ],
    roots: "/body, /styles/<style name> or /metadata",
};

#[cfg(feature = "odt")]
const _: () = assert!(well_formed(&ODT) && names_are_aliases(&ODT) && ODT.elements.len() <= 128);

/// Legacy Word and RTF share the semantic path spelling but have no OOXML
/// paragraph wire identities. Keep the known unsupported concepts for typed
/// refusals.
#[cfg(feature = "doc")]
pub(crate) static DOC: Vocabulary = Vocabulary {
    family: WORD.family,
    root: WORD.root,
    aliases: WORD.aliases,
    elements: &word_elements_without_paragraph_identity(),
    roots: WORD.roots,
};

/// RTF uses the same Word family aliases and the same absence of OOXML
/// paragraph identities as legacy DOC. Keeping a distinct static lets the
/// adapter's vocabulary remain an explicit format-level registry entry.
#[cfg(feature = "rtf")]
pub(crate) static RTF: Vocabulary = Vocabulary {
    family: WORD.family,
    root: WORD.root,
    aliases: WORD.aliases,
    elements: &word_elements_without_paragraph_identity(),
    roots: WORD.roots,
};

#[cfg(any(feature = "doc", feature = "rtf"))]
const fn word_elements_without_paragraph_identity() -> [ElementRule; WORD.elements.len()] {
    let mut rules = [ElementRule::new(""); WORD.elements.len()];
    let mut index = 0;
    while index < rules.len() {
        rules[index] = WORD.elements[index];
        let name = rules[index].name.as_bytes();
        if name.len() == 1 && name[0] == b'p' {
            rules[index].identities = &[];
        }
        index += 1;
    }
    rules
}

/// Presentations: PPTX, PPT, ODP.
///
/// In paths, `textbox` and `title` are plain aliases of `shape`, counted with
/// every other shape (`/slide[1]/title[2]` is the second shape). Query
/// selectors narrow them, as OfficeCLI's do: `title` selects title and
/// centred-title placeholders, `textbox` every other shape (query
/// compilation, M1 step R3).
pub(crate) static PRESENTATION: Vocabulary = Vocabulary {
    family: Family::Presentation,
    root: "presentation",
    aliases: &[
        ("3dmodel", "model3d"),
        ("animate", "animation"),
        ("animation", "animation"),
        ("audio", "audio"),
        ("axis", "axis"),
        ("br", "linebreak"),
        ("cell", "tc"),
        ("chart", "chart"),
        ("col", "col"),
        ("column", "col"),
        ("comment", "comment"),
        ("connection", "connector"),
        ("connector", "connector"),
        ("embed", "ole"),
        ("equation", "equation"),
        ("formula", "equation"),
        ("group", "group"),
        ("hlink", "hyperlink"),
        ("hyperlink", "hyperlink"),
        ("image", "picture"),
        ("img", "picture"),
        ("layout", "slideLayout"),
        ("linebreak", "linebreak"),
        ("math", "equation"),
        ("media", "media"),
        ("model3d", "model3d"),
        ("modern-comment", "modernComment"),
        ("moderncomment", "modernComment"),
        ("node", "node"),
        ("notes", "notes"),
        ("noteslide", "noteSlide"),
        ("notesmaster", "notesMaster"),
        ("object", "ole"),
        ("ole", "ole"),
        ("oleobject", "ole"),
        ("p", "paragraph"),
        ("paragraph", "paragraph"),
        ("ph", "placeholder"),
        ("pic", "picture"),
        ("picture", "picture"),
        ("placeholder", "placeholder"),
        ("r", "run"),
        ("reply", "reply"),
        ("row", "tr"),
        ("run", "run"),
        ("series", "series"),
        ("shape", "shape"),
        ("slide", "slide"),
        ("slidelayout", "slideLayout"),
        ("slidemaster", "slideMaster"),
        ("slidezoom", "zoom"),
        ("smartart", "smartart"),
        ("table", "table"),
        ("tc", "tc"),
        ("textbox", "shape"),
        ("theme", "theme"),
        ("thread", "modernComment"),
        ("threadedcomment", "modernComment"),
        ("title", "shape"),
        ("tr", "tr"),
        ("video", "video"),
        ("zoom", "zoom"),
    ],
    elements: &[
        ElementRule::new("animation"),
        ElementRule::new("audio").ids(SHAPE_IDENTITIES),
        ElementRule::new("axis").ids(&[Identity::new("role", Ascii)]),
        ElementRule::new("chart").ids(SHAPE_IDENTITIES),
        ElementRule::new("col"),
        ElementRule::new("comment"),
        ElementRule::new("connector").ids(SHAPE_IDENTITIES),
        ElementRule::new("equation"),
        ElementRule::new("group").ids(SHAPE_IDENTITIES),
        ElementRule::new("hyperlink"),
        ElementRule::new("linebreak"),
        ElementRule::new("media"),
        ElementRule::new("model3d"),
        ElementRule::new("modernComment"),
        ElementRule::new("node"),
        ElementRule::new("notes"),
        ElementRule::new("noteSlide"),
        ElementRule::new("notesMaster"),
        ElementRule::new("ole").ids(SHAPE_IDENTITIES),
        ElementRule::new("paragraph"),
        ElementRule::new("picture").ids(SHAPE_IDENTITIES),
        ElementRule::new("placeholder")
            .ids(&[
                Identity::new("type", Ascii),
                Identity::new("idx", Exact),
                Identity::new("id", Exact),
                Identity::new("name", Caseless),
            ])
            .bare(Bare::Token),
        ElementRule::new("presentation"),
        ElementRule::new("reply"),
        ElementRule::new("run"),
        ElementRule::new("series").ids(&[Identity::new("id", Exact)]),
        ElementRule::new("shape").ids(SHAPE_IDENTITIES),
        ElementRule::new("slide"),
        ElementRule::new("slideLayout"),
        ElementRule::new("slideMaster"),
        ElementRule::new("smartart").ids(SHAPE_IDENTITIES),
        ElementRule::new("table").ids(SHAPE_IDENTITIES),
        ElementRule::new("tc"),
        ElementRule::new("theme"),
        ElementRule::new("tr"),
        ElementRule::new("video").ids(SHAPE_IDENTITIES),
        ElementRule::new("zoom"),
    ],
    roots: "/slide[N], /slidemaster[N], /slidelayout[N], /notesMaster or /theme",
};

/// The identities of a presentation shape-tree element: the `cNvPr` id, and
/// its name compared as OfficeCLI's `MatchesShapeName` does (case-insensitive;
/// the adapter also ignores a `!!` morph prefix on either side).
const SHAPE_IDENTITIES: &[Identity] =
    &[Identity::new("id", Exact), Identity::new("name", Caseless)];

/// Workbooks: XLSX, XLSB, XLS, ODS.
pub(crate) static WORKBOOK: Vocabulary = Vocabulary {
    family: Family::Workbook,
    root: "workbook",
    aliases: &[
        ("autofilter", "autofilter"),
        ("axis", "axis"),
        ("cell", "cell"),
        ("cf", "cf"),
        ("chart", "chart"),
        ("col", "col"),
        ("colbreak", "colbreak"),
        ("column", "col"),
        ("comment", "comment"),
        ("conditionalformatting", "cf"),
        ("datavalidation", "dataValidation"),
        ("definedname", "namedrange"),
        ("embed", "ole"),
        ("hyperlink", "hyperlink"),
        ("image", "picture"),
        ("listobject", "table"),
        ("media", "picture"),
        ("merge", "mergeCell"),
        ("mergecell", "mergeCell"),
        ("mergedrange", "mergeCell"),
        ("namedrange", "namedrange"),
        ("note", "comment"),
        ("object", "ole"),
        ("ole", "ole"),
        ("oleobject", "ole"),
        ("picture", "picture"),
        ("pivot", "pivottable"),
        ("pivottable", "pivottable"),
        ("range", "range"),
        ("row", "row"),
        ("rowbreak", "rowbreak"),
        ("run", "run"),
        ("series", "series"),
        ("shape", "shape"),
        ("sheet", "sheet"),
        ("slicer", "slicer"),
        ("sparkline", "sparkline"),
        ("table", "table"),
        ("validation", "dataValidation"),
    ],
    elements: &[
        ElementRule::new("autofilter"),
        ElementRule::new("axis").ids(&[Identity::new("role", Ascii)]),
        ElementRule::new("cell"),
        ElementRule::new("cf"),
        ElementRule::new("chart"),
        ElementRule::new("col")
            .index(IndexRule::Column)
            .bare(Bare::Column),
        ElementRule::new("colbreak"),
        ElementRule::new("comment"),
        ElementRule::new("dataValidation"),
        ElementRule::new("hyperlink"),
        ElementRule::new("mergeCell"),
        ElementRule::new("namedrange")
            .ids(&[Identity::new("name", Caseless)])
            .bare(Bare::Name(Caseless)),
        ElementRule::new("ole"),
        ElementRule::new("picture"),
        ElementRule::new("pivottable"),
        ElementRule::new("range"),
        ElementRule::new("row").index(IndexRule::Row),
        ElementRule::new("rowbreak"),
        ElementRule::new("run"),
        ElementRule::new("series"),
        ElementRule::new("shape"),
        ElementRule::new("sheet").cells(),
        ElementRule::new("slicer"),
        ElementRule::new("sparkline"),
        ElementRule::new("table").ids(&[Identity::new("name", Caseless)]),
        ElementRule::new("workbook").named(Named::First {
            child: "sheet",
            case: Caseless,
        }),
    ],
    roots: "/<sheet name>, /sheet[N] or /namedrange[NAME] (cells as /Sheet1/A1 or Sheet1!A1)",
};
