//! Read-only projection of generated ODF text indexes.

/// Scope used by an index source.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum IndexScope {
    /// Include the complete document.
    Document,
    /// Include the current chapter.
    Chapter,
}

/// Caption formatting policy used by illustration and table indexes.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CaptionSequenceFormat {
    /// Use the sequence text.
    Text,
    /// Use category and value.
    CategoryAndValue,
    /// Use the complete caption.
    Caption,
}

/// Typed scalar source configuration for one generated index family.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum IndexSourceOptions {
    /// Table-of-contents source attributes.
    TableOfContents {
        outline_level: Option<usize>,
        use_outline_level: Option<bool>,
        use_index_marks: Option<bool>,
        use_index_source_styles: Option<bool>,
    },
    /// Illustration or table caption source attributes.
    Illustration {
        use_caption: Option<bool>,
        caption_sequence_name: Option<String>,
        caption_sequence_format: Option<CaptionSequenceFormat>,
    },
    /// Object-index source attributes.
    Object {
        use_spreadsheet_objects: Option<bool>,
        use_math_objects: Option<bool>,
        use_draw_objects: Option<bool>,
        use_chart_objects: Option<bool>,
        use_other_objects: Option<bool>,
    },
    /// User-index source attributes.
    User {
        use_index_marks: Option<bool>,
        use_index_source_styles: Option<bool>,
        use_graphics: Option<bool>,
        use_tables: Option<bool>,
        use_floating_frames: Option<bool>,
        use_objects: Option<bool>,
        copy_outline_levels: Option<bool>,
        index_name: String,
    },
    /// Alphabetical-index source attributes.
    Alphabetical {
        ignore_case: Option<bool>,
        main_entry_style_name: Option<String>,
        alphabetical_separators: Option<bool>,
        combine_entries: Option<bool>,
        combine_entries_with_dash: Option<bool>,
        combine_entries_with_pp: Option<bool>,
        use_keys_as_entries: Option<bool>,
        capitalize_entries: Option<bool>,
        comma_separated: Option<bool>,
        language: Option<String>,
        country: Option<String>,
        script: Option<String>,
        rfc_language_tag: Option<String>,
        sort_algorithm: Option<String>,
    },
    /// Bibliography source. Bibliography has no scalar source attributes in
    /// the first read-only metadata slice.
    Bibliography,
}

/// Common scope and scalar options of an index source.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct IndexSource {
    scope: Option<IndexScope>,
    relative_tab_stop_position: Option<bool>,
    options: IndexSourceOptions,
}

impl IndexSource {
    pub(crate) const fn projected(
        scope: Option<IndexScope>,
        relative_tab_stop_position: Option<bool>,
        options: IndexSourceOptions,
    ) -> Self {
        Self {
            scope,
            relative_tab_stop_position,
            options,
        }
    }

    /// Optional document/chapter scope.
    #[must_use]
    pub const fn scope(&self) -> Option<IndexScope> {
        self.scope
    }
    /// Optional relative tab-stop flag.
    #[must_use]
    pub const fn relative_tab_stop_position(&self) -> Option<bool> {
        self.relative_tab_stop_position
    }
    /// Family-specific source options.
    #[must_use]
    pub const fn options(&self) -> &IndexSourceOptions {
        &self.options
    }
}

/// Generated text-index family.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Kind {
    TableOfContents,
    Illustration,
    Table,
    Object,
    User,
    Alphabetical,
    Bibliography,
    Other(String),
}

impl Kind {
    /// Lexical ODF index element name represented by this kind.
    #[must_use]
    pub fn as_str(&self) -> &str {
        match self {
            Self::TableOfContents => "table-of-content",
            Self::Illustration => "illustration-index",
            Self::Table => "table-index",
            Self::Object => "object-index",
            Self::User => "user-index",
            Self::Alphabetical => "alphabetical-index",
            Self::Bibliography => "bibliography",
            Self::Other(value) => value,
        }
    }
}

/// One generated index declaration with its inert cached body text.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Index {
    body: String,
    kind: Kind,
    name: Option<String>,
    protected: bool,
    protected_present: Option<bool>,
    protection_key: Option<String>,
    protection_key_digest_algorithm: Option<String>,
    source: Option<String>,
    source_options: Option<IndexSource>,
    style_name: Option<String>,
    xml_id: Option<String>,
}

impl Index {
    pub(crate) const fn projected(
        kind: Kind,
        name: Option<String>,
        protected: bool,
        protected_present: Option<bool>,
        protection_key: Option<String>,
        protection_key_digest_algorithm: Option<String>,
        source: Option<String>,
        source_options: Option<IndexSource>,
        style_name: Option<String>,
        xml_id: Option<String>,
        body: String,
    ) -> Self {
        Self {
            body,
            kind,
            name,
            protected,
            protected_present,
            protection_key,
            protection_key_digest_algorithm,
            source,
            source_options,
            style_name,
            xml_id,
        }
    }

    /// Index family.
    #[must_use]
    pub const fn kind(&self) -> &Kind {
        &self.kind
    }

    /// Optional producer-visible index name.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Whether the index is marked protected.
    #[must_use]
    pub const fn is_protected(&self) -> bool {
        self.protected
    }

    /// Optional source protection flag, preserving absent-versus-false.
    #[must_use]
    pub const fn protected(&self) -> Option<bool> {
        self.protected_present
    }

    /// Optional protection key.
    #[must_use]
    pub fn protection_key(&self) -> Option<&str> {
        self.protection_key.as_deref()
    }

    /// Optional protection-key digest algorithm IRI.
    #[must_use]
    pub fn protection_key_digest_algorithm(&self) -> Option<&str> {
        self.protection_key_digest_algorithm.as_deref()
    }

    /// Optional index style reference.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// Optional XML identity.
    #[must_use]
    pub fn xml_id(&self) -> Option<&str> {
        self.xml_id.as_deref()
    }

    /// Inert source/template text, if a source element was present.
    #[must_use]
    pub fn source(&self) -> Option<&str> {
        self.source.as_deref()
    }

    /// Typed scalar source configuration, if the matching source element was
    /// present and valid.
    #[must_use]
    pub const fn source_options(&self) -> Option<&IndexSource> {
        self.source_options.as_ref()
    }

    /// Stored cached index body text. Entries are never regenerated.
    #[must_use]
    pub fn body(&self) -> &str {
        &self.body
    }
}
