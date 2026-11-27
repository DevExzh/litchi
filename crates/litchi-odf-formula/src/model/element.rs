//! Inert, namespace-aware `MathML` data model.

macro_rules! content_symbols {
    ($( $constant:ident => $name:literal ),+ $(,)?) => {
        impl ContentSymbol {
            $(pub const $constant: Self = Self($name);)+

            /// Every accepted named Content `MathML` symbol.
            pub const ALL: &'static [Self] = &[$(Self::$constant),+];

            /// Parse an exact `MathML` local name.
            #[must_use]
            pub fn from_local_name(name: &str) -> Option<Self> {
                match name {
                    $($name => Some(Self::$constant),)+
                    _ => None,
                }
            }
        }
    };
}

/// The `MathML` namespace used by Formula documents.
pub(crate) const MATHML_NAMESPACE: &str = "http://www.w3.org/1998/Math/MathML";

/// A checked Content `MathML` structural kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum ContentKind {
    Application,
    BoundVariable,
    Condition,
    Declaration,
    Degree,
    DomainOfApplication,
    Function,
    Identifier,
    Interval,
    Lambda,
    List,
    LogBase,
    LowLimit,
    Matrix,
    MatrixRow,
    MomentAbout,
    Number,
    Otherwise,
    Piece,
    Piecewise,
    Relation,
    Separator,
    Set,
    Symbol,
    SymbolToken,
    UpLimit,
    Vector,
}

/// A named empty Content `MathML` symbol.
///
/// The constants cover the complete named-symbol corpus accepted by the
/// crate's `MathML` 2 validator. The private representation prevents callers
/// from constructing an unknown symbol.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ContentSymbol(&'static str);

content_symbols! {
    ABS => "abs", AND => "and", APPROX => "approx", ARCCOS => "arccos",
    ARCCOSH => "arccosh", ARCCOT => "arccot", ARCCOTH => "arccoth",
    ARCCSC => "arccsc", ARCCSCH => "arccsch", ARCSEC => "arcsec",
    ARCSECH => "arcsech", ARCSIN => "arcsin", ARCSINH => "arcsinh",
    ARCTAN => "arctan", ARCTANH => "arctanh", ARG => "arg", CARD => "card",
    CARTESIAN_PRODUCT => "cartesianproduct", CEILING => "ceiling",
    COMPLEXES => "complexes", COMPOSE => "compose", CONJUGATE => "conjugate",
    CODOMAIN => "codomain", COS => "cos", COSH => "cosh", COT => "cot",
    COTH => "coth", CSC => "csc", CSCH => "csch", CURL => "curl",
    DETERMINANT => "determinant", DIFF => "diff", DIVERGENCE => "divergence",
    DIVIDE => "divide", DOMAIN => "domain", EMPTY_SET => "emptyset", EQ => "eq",
    EQUIVALENT => "equivalent", EULER_GAMMA => "eulergamma", EXISTS => "exists",
    EXP => "exp", EXPONENTIAL_E => "exponentiale", FACTORIAL => "factorial",
    FACTOR_OF => "factorof", FALSE => "false", FLOOR => "floor", FORALL => "forall",
    GCD => "gcd", GEQ => "geq", GRAD => "grad", GT => "gt", IDENT => "ident",
    IMAGE => "image", IMAGINARY => "imaginary", IMAGINARY_I => "imaginaryi",
    IMPLIES => "implies", IN => "in", INFINITY => "infinity", INTEGERS => "integers",
    INTERSECT => "intersect", INT => "int", INVERSE => "inverse",
    LAPLACIAN => "laplacian", LCM => "lcm", LEQ => "leq", LIMIT => "limit",
    LN => "ln", LOG => "log", LT => "lt", MAX => "max", MEAN => "mean",
    MEDIAN => "median", MIN => "min", MINUS => "minus", MODE => "mode",
    MOMENT => "moment", NATURAL_NUMBERS => "naturalnumbers", NEQ => "neq",
    NOT => "not", NOT_A_NUMBER => "notanumber", NOT_IN => "notin",
    NOT_PR_SUBSET => "notprsubset", NOT_SUBSET => "notsubset", OR => "or",
    OUTER_PRODUCT => "outerproduct", PARTIAL_DIFF => "partialdiff", PI => "pi",
    PLUS => "plus", POWER => "power", PRIMES => "primes", PRODUCT => "product",
    PR_SUBSET => "prsubset", QUOTIENT => "quotient", RATIONALS => "rationals",
    REAL => "real", REALS => "reals", REM => "rem", ROOT => "root",
    SCALAR_PRODUCT => "scalarproduct", SDEV => "sdev", SEC => "sec", SECH => "sech",
    SELECTOR => "selector", SET_DIFF => "setdiff", SIN => "sin", SINH => "sinh",
    SUBSET => "subset", SUM => "sum", TAN => "tan", TANH => "tanh",
    TENDS_TO => "tendsto", TIMES => "times", TRANSPOSE => "transpose", TRUE => "true",
    UNION => "union", VARIANCE => "variance", VECTOR_PRODUCT => "vectorproduct", XOR => "xor",
}

impl ContentSymbol {
    /// Exact `MathML` local name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        self.0
    }
}

/// A commonly used `MathML` element kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Kind {
    Math,
    Semantics,
    Annotation,
    AnnotationXml,
    Row,
    Identifier,
    Number,
    Operator,
    Text,
    Space,
    StringLiteral,
    Glyph,
    Fraction,
    SquareRoot,
    Root,
    Style,
    Error,
    Padded,
    Phantom,
    Fenced,
    Enclose,
    Subscript,
    Superscript,
    SubSuperscript,
    Under,
    Over,
    UnderOver,
    MultiScripts,
    Table,
    TableRow,
    TableCell,
    AlignGroup,
    AlignMark,
    Action,
    None,
    PreScripts,
    /// A Content `MathML` structural element.
    Content(ContentKind),
    /// A named empty Content `MathML` symbol.
    ContentSymbol(ContentSymbol),
    /// A future `MathML` element or a vendor element in another namespace.
    Other,
}

/// One decoded attribute with its expanded namespace name.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Attribute {
    namespace_uri: Option<String>,
    local_name: String,
    value: String,
}

impl Attribute {
    pub(crate) fn from_parts(
        namespace_uri: Option<String>,
        local_name: String,
        value: String,
    ) -> Self {
        Self {
            namespace_uri,
            local_name,
            value,
        }
    }

    /// Return the expanded namespace URI, or `None` for an unqualified attribute.
    #[must_use]
    pub fn namespace_uri(&self) -> Option<&str> {
        self.namespace_uri.as_deref()
    }

    /// Return the XML local name.
    #[must_use]
    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    /// Return the decoded and normalized XML attribute value.
    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
    }
}

/// Ordered mixed content within a `MathML` element.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Content {
    /// Decoded character content, including CDATA and character references.
    ///
    /// Named references other than `XML`'s five predefined entities are retained
    /// in `&name;` notation because `MathML` 2 documents may declare them in a
    /// document type definition that is intentionally not evaluated here.
    Text(String),
    /// A child element.
    Element(Element),
}

/// A complete element in the formula's `MathML` subtree.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Element {
    namespace_uri: Option<String>,
    local_name: String,
    attributes: Vec<Attribute>,
    content: Vec<Content>,
}

impl Element {
    pub(crate) fn fixed_mathml(local_name: &'static str) -> Self {
        Self::from_parts(
            Some(MATHML_NAMESPACE.to_string()),
            local_name.to_string(),
            Vec::new(),
            Vec::new(),
        )
    }

    pub(crate) fn set_fixed_attribute(&mut self, local_name: &'static str, value: &str) {
        if let Some(existing) = self.attributes.iter_mut().find(|attribute| {
            attribute.namespace_uri().is_none() && attribute.local_name == local_name
        }) {
            *existing = Attribute::from_parts(None, local_name.to_string(), value.to_string());
        } else {
            self.attributes.push(Attribute::from_parts(
                None,
                local_name.to_string(),
                value.to_string(),
            ));
        }
    }

    pub(crate) fn from_parts(
        namespace_uri: Option<String>,
        local_name: String,
        attributes: Vec<Attribute>,
        content: Vec<Content>,
    ) -> Self {
        Self {
            namespace_uri,
            local_name,
            attributes,
            content,
        }
    }

    pub(crate) fn attributes_mut(&mut self) -> &mut Vec<Attribute> {
        &mut self.attributes
    }

    pub(crate) fn content_mut(&mut self) -> &mut Vec<Content> {
        &mut self.content
    }

    /// Return the element's expanded namespace URI.
    #[must_use]
    pub fn namespace_uri(&self) -> Option<&str> {
        self.namespace_uri.as_deref()
    }

    /// Return the element's XML local name.
    #[must_use]
    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    /// Classify common `MathML` elements without discarding unknown ones.
    #[must_use]
    pub fn kind(&self) -> Kind {
        if self.namespace_uri() != Some(MATHML_NAMESPACE) {
            return Kind::Other;
        }
        match self.local_name.as_str() {
            "math" => Kind::Math,
            "semantics" => Kind::Semantics,
            "annotation" => Kind::Annotation,
            "annotation-xml" => Kind::AnnotationXml,
            "mrow" => Kind::Row,
            "mi" => Kind::Identifier,
            "mn" => Kind::Number,
            "mo" => Kind::Operator,
            "mtext" => Kind::Text,
            "mspace" => Kind::Space,
            "ms" => Kind::StringLiteral,
            "mglyph" => Kind::Glyph,
            "mfrac" => Kind::Fraction,
            "msqrt" => Kind::SquareRoot,
            "mroot" => Kind::Root,
            "mstyle" => Kind::Style,
            "merror" => Kind::Error,
            "mpadded" => Kind::Padded,
            "mphantom" => Kind::Phantom,
            "mfenced" => Kind::Fenced,
            "menclose" => Kind::Enclose,
            "msub" => Kind::Subscript,
            "msup" => Kind::Superscript,
            "msubsup" => Kind::SubSuperscript,
            "munder" => Kind::Under,
            "mover" => Kind::Over,
            "munderover" => Kind::UnderOver,
            "mmultiscripts" => Kind::MultiScripts,
            "mtable" => Kind::Table,
            "mtr" | "mlabeledtr" => Kind::TableRow,
            "mtd" => Kind::TableCell,
            "maligngroup" => Kind::AlignGroup,
            "malignmark" => Kind::AlignMark,
            "maction" => Kind::Action,
            "none" => Kind::None,
            "mprescripts" => Kind::PreScripts,
            "apply" => Kind::Content(ContentKind::Application),
            "bvar" => Kind::Content(ContentKind::BoundVariable),
            "condition" => Kind::Content(ContentKind::Condition),
            "declare" => Kind::Content(ContentKind::Declaration),
            "degree" => Kind::Content(ContentKind::Degree),
            "domainofapplication" => Kind::Content(ContentKind::DomainOfApplication),
            "fn" => Kind::Content(ContentKind::Function),
            "ci" => Kind::Content(ContentKind::Identifier),
            "interval" => Kind::Content(ContentKind::Interval),
            "lambda" => Kind::Content(ContentKind::Lambda),
            "list" => Kind::Content(ContentKind::List),
            "logbase" => Kind::Content(ContentKind::LogBase),
            "lowlimit" => Kind::Content(ContentKind::LowLimit),
            "matrix" => Kind::Content(ContentKind::Matrix),
            "matrixrow" => Kind::Content(ContentKind::MatrixRow),
            "momentabout" => Kind::Content(ContentKind::MomentAbout),
            "cn" => Kind::Content(ContentKind::Number),
            "otherwise" => Kind::Content(ContentKind::Otherwise),
            "piece" => Kind::Content(ContentKind::Piece),
            "piecewise" => Kind::Content(ContentKind::Piecewise),
            "reln" => Kind::Content(ContentKind::Relation),
            "sep" => Kind::Content(ContentKind::Separator),
            "set" => Kind::Content(ContentKind::Set),
            "csymbol" => Kind::Content(ContentKind::SymbolToken),
            "uplimit" => Kind::Content(ContentKind::UpLimit),
            "vector" => Kind::Content(ContentKind::Vector),
            name => ContentSymbol::from_local_name(name).map_or(Kind::Other, Kind::ContentSymbol),
        }
    }

    /// Return all decoded attributes in document order.
    #[must_use]
    pub fn attributes(&self) -> &[Attribute] {
        &self.attributes
    }

    /// Find an attribute by expanded name.
    pub fn attribute(&self, namespace_uri: Option<&str>, local_name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| {
                attribute.namespace_uri() == namespace_uri && attribute.local_name == local_name
            })
            .map(Attribute::value)
    }

    /// Return ordered mixed content.
    #[must_use]
    pub fn content(&self) -> &[Content] {
        &self.content
    }

    /// Iterate direct child elements.
    pub fn children(&self) -> impl Iterator<Item = &Element> {
        self.content.iter().filter_map(|content| match content {
            Content::Element(element) => Some(element),
            Content::Text(_) => None,
        })
    }

    /// Compose all descendant character content in exact element/text order.
    #[must_use]
    pub fn all_text(&self) -> String {
        let mut output = String::new();
        let mut pending: Vec<_> = self.content.iter().rev().collect();
        while let Some(content) = pending.pop() {
            match content {
                Content::Text(text) => output.push_str(text),
                Content::Element(child) => pending.extend(child.content.iter().rev()),
            }
        }
        output
    }

    pub(crate) fn collect_annotations<'a>(&'a self, output: &mut Vec<&'a Element>) {
        let mut pending = vec![self];
        while let Some(element) = pending.pop() {
            if matches!(element.kind(), Kind::Annotation | Kind::AnnotationXml) {
                output.push(element);
            }
            let children: Vec<_> = element.children().collect();
            pending.extend(children.into_iter().rev());
        }
    }
}
