use super::super::{TextIndex, TextIndexContent, TextIndexKind};
use super::semantic::*;
use super::xml::{
    LocaleLexical, attr, bibliography_token_element, caption_index, common_source_attributes,
    element, element_content, finish_typed_index, invalid, optional_bool,
    optional_locale_attribute, optional_name, optional_positive, positive, required, set_attr,
    single_template_content, source_styles_element, title_template_element, token_element,
    validate_index, validate_links,
};
use super::{FO, MAX_BODY_PARAGRAPHS, MAX_TEMPLATES, MAX_TOKENS, STYLE, TEXT};
use litchi_core::{Error, Result};

impl TextIndex {
    /// Create ODF index markup from caller-provided templates and cached body.
    ///
    /// This does not scan headings, evaluate fields, paginate, or refresh the cache.
    pub fn table_of_contents(
        name: impl Into<String>,
        source: TableOfContentsSource,
        body: TextIndexBody,
    ) -> Result<Self> {
        let name = name.into();
        required(&name, "text index name")?;
        if source.entry_templates.len() > MAX_TEMPLATES
            || source.source_styles.len() > MAX_TEMPLATES
            || body.paragraphs.len() > MAX_BODY_PARAGRAPHS
        {
            return invalid("text index exceeds configured resource limits");
        }

        let mut source_attributes = Vec::new();
        optional_positive(
            &mut source_attributes,
            "outline-level",
            source.outline_level,
        )?;
        optional_bool(
            &mut source_attributes,
            "use-outline-level",
            source.use_outline_level,
        );
        optional_bool(
            &mut source_attributes,
            "use-index-marks",
            source.use_index_marks,
        );
        optional_bool(
            &mut source_attributes,
            "use-index-source-styles",
            source.use_index_source_styles,
        );
        if let Some(scope) = source.scope {
            source_attributes.push(attr(
                TEXT,
                "index-scope",
                match scope {
                    TextIndexScope::Document => "document",
                    TextIndexScope::Chapter => "chapter",
                },
            ));
        }
        optional_bool(
            &mut source_attributes,
            "relative-tab-stop-position",
            source.relative_tab_stop_position,
        );

        let mut source_content = Vec::new();
        if let Some(title) = source.title_template {
            let mut attributes = Vec::new();
            optional_name(&mut attributes, "style-name", title.style_name)?;
            source_content.push(element_content(element(
                TEXT,
                "index-title-template",
                attributes,
                vec![TextIndexContent::Text(title.text)],
            )));
        }

        let mut token_count = 0usize;
        for template in source.entry_templates {
            positive(template.outline_level, "entry template outline level")?;
            required(&template.style_name, "entry template style name")?;
            token_count = token_count
                .checked_add(template.tokens.len())
                .ok_or_else(|| {
                    Error::InvalidFormat("text index token count overflow".to_string())
                })?;
            if token_count > MAX_TOKENS {
                return invalid("text index contains too many entry tokens");
            }
            validate_links(&template.tokens)?;
            let mut tokens = Vec::with_capacity(template.tokens.len());
            for token in template.tokens {
                tokens.push(element_content(token_element(token)?));
            }
            source_content.push(element_content(element(
                TEXT,
                "table-of-content-entry-template",
                vec![
                    attr(TEXT, "outline-level", template.outline_level.to_string()),
                    attr(TEXT, "style-name", template.style_name),
                ],
                tokens,
            )));
        }

        for styles in source.source_styles {
            positive(styles.outline_level, "source styles outline level")?;
            let mut children = Vec::with_capacity(styles.style_names.len());
            for style_name in styles.style_names {
                required(&style_name, "source style name")?;
                children.push(element_content(element(
                    TEXT,
                    "index-source-style",
                    vec![attr(TEXT, "style-name", style_name)],
                    Vec::new(),
                )));
            }
            source_content.push(element_content(element(
                TEXT,
                "index-source-styles",
                vec![attr(
                    TEXT,
                    "outline-level",
                    styles.outline_level.to_string(),
                )],
                children,
            )));
        }

        let mut body_content = Vec::new();
        if let Some(title) = body.title {
            required(&title.name, "index body title name")?;
            let mut title_attrs = vec![attr(TEXT, "name", title.name)];
            optional_name(&mut title_attrs, "style-name", title.section_style_name)?;
            let mut paragraph_attrs = Vec::new();
            optional_name(
                &mut paragraph_attrs,
                "style-name",
                title.paragraph_style_name,
            )?;
            body_content.push(element_content(element(
                TEXT,
                "index-title",
                title_attrs,
                vec![element_content(element(
                    TEXT,
                    "p",
                    paragraph_attrs,
                    vec![TextIndexContent::Text(title.text)],
                ))],
            )));
        }
        for paragraph in body.paragraphs {
            let mut attributes = Vec::new();
            optional_name(&mut attributes, "style-name", paragraph.style_name)?;
            body_content.push(element_content(element(
                TEXT,
                "p",
                attributes,
                vec![TextIndexContent::Text(paragraph.text)],
            )));
        }

        let index = Self {
            kind: TextIndexKind::TableOfContents,
            root: element(
                TEXT,
                "table-of-content",
                vec![attr(TEXT, "name", name)],
                vec![
                    element_content(element(
                        TEXT,
                        "table-of-content-source",
                        source_attributes,
                        source_content,
                    )),
                    element_content(element(TEXT, "index-body", Vec::new(), body_content)),
                ],
            ),
        };
        validate_index(&index)?;
        Ok(index)
    }

    pub fn illustration_index(
        name: impl Into<String>,
        source: IllustrationIndexSource,
        body: TextIndexBody,
    ) -> Result<Self> {
        caption_index(
            TextIndexKind::Illustration,
            "illustration-index",
            "illustration-index-source",
            "illustration-index-entry-template",
            name.into(),
            source,
            body,
        )
    }

    pub fn table_index(
        name: impl Into<String>,
        source: IllustrationIndexSource,
        body: TextIndexBody,
    ) -> Result<Self> {
        caption_index(
            TextIndexKind::Table,
            "table-index",
            "table-index-source",
            "table-index-entry-template",
            name.into(),
            source,
            body,
        )
    }

    pub fn object_index(
        name: impl Into<String>,
        source: ObjectIndexSource,
        body: TextIndexBody,
    ) -> Result<Self> {
        let name = name.into();
        required(&name, "text index name")?;
        let mut attributes =
            common_source_attributes(source.scope, source.relative_tab_stop_position);
        optional_bool(
            &mut attributes,
            "use-spreadsheet-objects",
            source.use_spreadsheet_objects,
        );
        optional_bool(&mut attributes, "use-math-objects", source.use_math_objects);
        optional_bool(&mut attributes, "use-draw-objects", source.use_draw_objects);
        optional_bool(
            &mut attributes,
            "use-chart-objects",
            source.use_chart_objects,
        );
        optional_bool(
            &mut attributes,
            "use-other-objects",
            source.use_other_objects,
        );
        let content = single_template_content(
            source.title_template,
            source.entry_template,
            "object-index-entry-template",
        )?;
        finish_typed_index(
            TextIndexKind::Object,
            "object-index",
            "object-index-source",
            name,
            attributes,
            content,
            body,
        )
    }

    pub fn user_index(
        name: impl Into<String>,
        source: UserIndexSource,
        body: TextIndexBody,
    ) -> Result<Self> {
        let name = name.into();
        required(&name, "text index name")?;
        if source.entry_templates.len() > MAX_TEMPLATES
            || source.source_styles.len() > MAX_TEMPLATES
        {
            return invalid("user index exceeds configured template limits");
        }
        let mut attributes =
            common_source_attributes(source.scope, source.relative_tab_stop_position);
        attributes.push(attr(TEXT, "index-name", source.index_name));
        optional_bool(&mut attributes, "use-index-marks", source.use_index_marks);
        optional_bool(
            &mut attributes,
            "use-index-source-styles",
            source.use_index_source_styles,
        );
        optional_bool(&mut attributes, "use-graphics", source.use_graphics);
        optional_bool(&mut attributes, "use-tables", source.use_tables);
        optional_bool(
            &mut attributes,
            "use-floating-frames",
            source.use_floating_frames,
        );
        optional_bool(&mut attributes, "use-objects", source.use_objects);
        optional_bool(
            &mut attributes,
            "copy-outline-levels",
            source.copy_outline_levels,
        );
        let mut content = Vec::new();
        if let Some(title) = source.title_template {
            content.push(element_content(title_template_element(title)?));
        }
        let mut token_count = 0usize;
        for template in source.entry_templates {
            positive(template.outline_level, "user index entry outline level")?;
            required(&template.style_name, "user index entry style name")?;
            token_count = token_count
                .checked_add(template.tokens.len())
                .ok_or_else(|| {
                    Error::InvalidFormat("user index token count overflow".to_string())
                })?;
            if token_count > MAX_TOKENS {
                return invalid("user index contains too many entry tokens");
            }
            validate_links(&template.tokens)?;
            content.push(element_content(element(
                TEXT,
                "user-index-entry-template",
                vec![
                    attr(TEXT, "outline-level", template.outline_level.to_string()),
                    attr(TEXT, "style-name", template.style_name),
                ],
                template
                    .tokens
                    .into_iter()
                    .map(token_element)
                    .collect::<Result<Vec<_>>>()?
                    .into_iter()
                    .map(element_content)
                    .collect(),
            )));
        }
        for styles in source.source_styles {
            content.push(element_content(source_styles_element(styles)?));
        }
        finish_typed_index(
            TextIndexKind::User,
            "user-index",
            "user-index-source",
            name,
            attributes,
            content,
            body,
        )
    }

    pub fn alphabetical_index(
        name: impl Into<String>,
        source: AlphabeticalIndexSource,
        body: TextIndexBody,
    ) -> Result<Self> {
        let name = name.into();
        required(&name, "text index name")?;
        if source.entry_templates.len() > MAX_TEMPLATES {
            return invalid("alphabetical index has too many templates");
        }
        let mut attributes =
            common_source_attributes(source.scope, source.relative_tab_stop_position);
        optional_bool(&mut attributes, "ignore-case", source.ignore_case);
        optional_name(
            &mut attributes,
            "main-entry-style-name",
            source.main_entry_style_name,
        )?;
        optional_bool(
            &mut attributes,
            "alphabetical-separators",
            source.alphabetical_separators,
        );
        optional_bool(&mut attributes, "combine-entries", source.combine_entries);
        optional_bool(
            &mut attributes,
            "combine-entries-with-dash",
            source.combine_entries_with_dash,
        );
        optional_bool(
            &mut attributes,
            "combine-entries-with-pp",
            source.combine_entries_with_pp,
        );
        optional_bool(
            &mut attributes,
            "use-keys-as-entries",
            source.use_keys_as_entries,
        );
        optional_bool(
            &mut attributes,
            "capitalize-entries",
            source.capitalize_entries,
        );
        optional_bool(&mut attributes, "comma-separated", source.comma_separated);
        optional_locale_attribute(
            &mut attributes,
            FO,
            "language",
            source.language,
            LocaleLexical::LanguageCode,
        )?;
        optional_locale_attribute(
            &mut attributes,
            FO,
            "country",
            source.country,
            LocaleLexical::CountryOrScript,
        )?;
        optional_locale_attribute(
            &mut attributes,
            FO,
            "script",
            source.script,
            LocaleLexical::CountryOrScript,
        )?;
        optional_locale_attribute(
            &mut attributes,
            STYLE,
            "rfc-language-tag",
            source.rfc_language_tag,
            LocaleLexical::LanguageTag,
        )?;
        if let Some(sort_algorithm) = source.sort_algorithm {
            attributes.push(attr(TEXT, "sort-algorithm", sort_algorithm));
        }
        let mut content = Vec::new();
        if let Some(title) = source.title_template {
            content.push(element_content(title_template_element(title)?));
        }
        let mut token_count = 0usize;
        for template in source.entry_templates {
            required(&template.style_name, "alphabetical index entry style name")?;
            token_count = token_count
                .checked_add(template.tokens.len())
                .ok_or_else(|| {
                    Error::InvalidFormat("alphabetical index token count overflow".to_string())
                })?;
            if token_count > MAX_TOKENS {
                return invalid("alphabetical index contains too many tokens");
            }
            if template.tokens.iter().any(|token| {
                matches!(
                    token,
                    TextIndexEntryToken::LinkStart { .. } | TextIndexEntryToken::LinkEnd { .. }
                )
            }) {
                return invalid("alphabetical index templates do not permit link tokens");
            }
            content.push(element_content(element(
                TEXT,
                "alphabetical-index-entry-template",
                vec![
                    attr(TEXT, "outline-level", template.level.as_str()),
                    attr(TEXT, "style-name", template.style_name),
                ],
                template
                    .tokens
                    .into_iter()
                    .map(token_element)
                    .collect::<Result<Vec<_>>>()?
                    .into_iter()
                    .map(element_content)
                    .collect(),
            )));
        }
        finish_typed_index(
            TextIndexKind::Alphabetical,
            "alphabetical-index",
            "alphabetical-index-source",
            name,
            attributes,
            content,
            body,
        )
    }

    pub fn bibliography(
        name: impl Into<String>,
        source: BibliographyIndexSource,
        body: TextIndexBody,
    ) -> Result<Self> {
        let name = name.into();
        required(&name, "text index name")?;
        if source.entry_templates.len() > MAX_TEMPLATES {
            return invalid("bibliography has too many templates");
        }
        let mut content = Vec::new();
        if let Some(title) = source.title_template {
            content.push(element_content(title_template_element(title)?));
        }
        let mut token_count = 0usize;
        for template in source.entry_templates {
            required(&template.style_name, "bibliography entry style name")?;
            token_count = token_count
                .checked_add(template.tokens.len())
                .ok_or_else(|| {
                    Error::InvalidFormat("bibliography token count overflow".to_string())
                })?;
            if token_count > MAX_TOKENS {
                return invalid("bibliography contains too many tokens");
            }
            let tokens = template
                .tokens
                .into_iter()
                .map(bibliography_token_element)
                .collect::<Result<Vec<_>>>()?
                .into_iter()
                .map(element_content)
                .collect();
            content.push(element_content(element(
                TEXT,
                "bibliography-entry-template",
                vec![
                    attr(
                        TEXT,
                        "bibliography-type",
                        template.bibliography_type.as_str(),
                    ),
                    attr(TEXT, "style-name", template.style_name),
                ],
                tokens,
            )));
        }
        finish_typed_index(
            TextIndexKind::Bibliography,
            "bibliography",
            "bibliography-source",
            name,
            Vec::new(),
            content,
            body,
        )
    }

    pub fn with_protected(mut self, protected: bool) -> Self {
        set_attr(
            &mut self.root.attributes,
            TEXT,
            "protected",
            protected.to_string(),
        );
        self
    }

    pub fn with_style_name(mut self, style_name: impl Into<String>) -> Result<Self> {
        let style_name = style_name.into();
        required(&style_name, "text index style name")?;
        set_attr(&mut self.root.attributes, TEXT, "style-name", style_name);
        Ok(self)
    }
}
