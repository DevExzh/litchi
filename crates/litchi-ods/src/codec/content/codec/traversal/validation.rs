use super::{
    BTreeMap, BytesStart, CALCEXT_NAMESPACE_URI, CellBuilder, CellDetective, CellMatrixSpan,
    CellMerge, CellRangeSource, ColorTransformationType, Column, ConditionalColorScaleEntry,
    ConditionalCustomIcon, ConditionalDataBar, ConditionalDataBarEntry, ConditionalDateIs,
    ConditionalDateType, ConditionalFormatCondition, ConditionalFormatEntryType,
    ConditionalIconSet, ConditionalIconSetEntry, DataBarAxisPosition, Decoder, DetectiveDirection,
    DetectiveHighlightedRange, DetectiveOperation, DetectiveOperationKind, Error, Event,
    IconSetType, LOEXT_NAMESPACE_URI, Link, NamedDefinition, NamedDefinitionScope, NamedExpression,
    NamedRange, NamedRangeUsage, Namespace, NamespaceResolver, NonZeroUsize, NsReader,
    OFFICE_NAMESPACE, Parser, PrefixDeclaration, ResolveResult, Result, RowBuilder, Sheet,
    SheetPrintSettings, SheetScenario, SheetStyle, SheetTableSource, Sparkline, SparklineAxisType,
    SparklineColorTransformation, SparklineComplexColor, SparklineEmptyCells, SparklineGroup,
    SparklineType, TABLE_NAMESPACE, TABLE_NAMESPACE_URI, TEXT_NAMESPACE, TableSourceMode,
    TableVisibility, TextHyperlinkActuate, TextHyperlinkShow, ThemeColorType, XLINK_NAMESPACE,
    XmlVersion, split_cell_range_addresses, validate_color_scale_entry, validate_condition,
    validate_data_bar_attributes, validate_data_bar_entry, validate_date_is,
    validate_icon_set_entry, validate_scenario, validate_sparkline,
    validate_sparkline_group_attributes, validate_table_source,
};

impl Parser {
    pub(super) fn is_office_annotation(
        element: &BytesStart<'_>,
        namespaces: &BTreeMap<String, String>,
    ) -> bool {
        Self::element_name_is(
            element.name().as_ref(),
            namespaces,
            OFFICE_NAMESPACE,
            "annotation",
        )
    }

    pub(super) fn element_name_is(
        qualified_name: &[u8],
        namespaces: &BTreeMap<String, String>,
        namespace: &str,
        local_name: &str,
    ) -> bool {
        let Ok(name) = std::str::from_utf8(qualified_name) else {
            return false;
        };
        let (prefix, local) = match name.split_once(':') {
            Some(parts) => parts,
            None => ("", name),
        };
        local == local_name
            && namespaces
                .get(prefix)
                .is_some_and(|candidate| candidate == namespace)
    }

    pub(super) fn attribute_name_is(
        qualified_name: &[u8],
        namespaces: &BTreeMap<String, String>,
        namespace: &str,
        local_name: &str,
    ) -> bool {
        let Ok(name) = std::str::from_utf8(qualified_name) else {
            return false;
        };
        let Some((prefix, local)) = name.split_once(':') else {
            return false;
        };
        local == local_name
            && namespaces
                .get(prefix)
                .is_some_and(|candidate| candidate == namespace)
    }

    pub(super) fn push_namespace_scope(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &mut BTreeMap<String, String>,
    ) -> Result<Vec<(String, Option<String>)>> {
        let mut previous_bindings = Vec::new();
        for attribute in element.attributes() {
            let attribute = attribute.map_err(|error| {
                Error::InvalidFormat(format!("invalid XML namespace declaration: {error}"))
            })?;
            let name = std::str::from_utf8(attribute.key.as_ref()).map_err(|_| {
                Error::InvalidFormat("invalid UTF-8 in XML namespace declaration".to_string())
            })?;
            let prefix = if name == "xmlns" {
                ""
            } else if let Some(prefix) = name.strip_prefix("xmlns:") {
                prefix
            } else {
                continue;
            };
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid XML namespace URI: {error}"))
                })?
                .into_owned();
            let prefix = prefix.to_string();
            let previous = namespaces.insert(prefix.clone(), value);
            previous_bindings.push((prefix, previous));
        }
        Ok(previous_bindings)
    }

    pub(super) fn pop_namespace_scope(
        namespaces: &mut BTreeMap<String, String>,
        scope: Option<Vec<(String, Option<String>)>>,
    ) {
        let Some(scope) = scope else { return };
        for (prefix, previous) in scope.into_iter().rev() {
            if let Some(previous) = previous {
                namespaces.insert(prefix, previous);
            } else {
                namespaces.remove(&prefix);
            }
        }
    }

    pub(super) fn push_text_empty_element(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
        text: &mut String,
    ) -> Result<()> {
        if Self::element_name_is(
            element.name().as_ref(),
            namespaces,
            TEXT_NAMESPACE,
            "line-break",
        ) {
            text.push('\n');
        } else if Self::element_name_is(element.name().as_ref(), namespaces, TEXT_NAMESPACE, "tab")
        {
            text.push('\t');
        } else if Self::element_name_is(element.name().as_ref(), namespaces, TEXT_NAMESPACE, "s") {
            let mut count = 1usize;
            for attribute in element.attributes() {
                let attribute = attribute.map_err(|error| {
                    Error::InvalidFormat(format!("invalid text:s attribute: {error}"))
                })?;
                if Self::attribute_name_is(attribute.key.as_ref(), namespaces, TEXT_NAMESPACE, "c")
                {
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                        .map_err(|error| {
                            Error::InvalidFormat(format!("invalid text:s count: {error}"))
                        })?;
                    count = value.parse::<usize>().map_err(|_| {
                        Error::InvalidFormat(format!("invalid text:s count '{value}'"))
                    })?;
                }
            }
            if count > 1_000_000 {
                return Err(Error::InvalidFormat(
                    "text:s count exceeds the supported safety limit".to_string(),
                ));
            }
            text.extend(std::iter::repeat_n(' ', count));
        }
        Ok(())
    }

    /// Parse document-global and sheet-local named ranges and expressions.
    ///
    /// Namespace URIs are resolved rather than assuming the conventional
    /// `table` prefix, because XML namespace prefixes are freely replaceable.
    pub fn parse_named_definitions(xml_content: &str) -> Result<Vec<NamedDefinition>> {
        let mut reader = NsReader::from_str(xml_content);
        let mut buf = Vec::new();
        let mut definitions = Vec::new();
        let mut table_stack: Vec<Option<String>> = Vec::new();
        let mut active_scope: Option<NamedDefinitionScope> = None;

        loop {
            let event = reader
                .read_resolved_event_into(&mut buf)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;

            match event {
                (namespace, Event::Start(element)) if Self::is_table_namespace(&namespace) => {
                    match element.local_name().as_ref() {
                        b"table" => {
                            let name = Self::table_attribute(
                                reader.resolver(),
                                reader.decoder(),
                                &element,
                                b"name",
                            )?;
                            table_stack.push(name);
                        },
                        b"named-expressions" => {
                            if active_scope.is_some() {
                                return Err(Error::InvalidFormat(
                                    "nested table:named-expressions element".to_string(),
                                ));
                            }
                            active_scope = Some(match table_stack.last() {
                                Some(Some(sheet)) => NamedDefinitionScope::Sheet(sheet.clone()),
                                Some(None) => {
                                    return Err(Error::InvalidFormat(
                                        "sheet-local named definitions require table:name"
                                            .to_string(),
                                    ));
                                },
                                None => NamedDefinitionScope::Global,
                            });
                        },
                        b"named-range" | b"named-expression" => {
                            if let Some(scope) = &active_scope {
                                definitions.push(Self::parse_named_definition(
                                    reader.resolver(),
                                    reader.decoder(),
                                    &element,
                                    scope.clone(),
                                )?);
                            }
                        },
                        _ => {},
                    }
                },
                (namespace, Event::Empty(element)) if Self::is_table_namespace(&namespace) => {
                    match element.local_name().as_ref() {
                        b"named-range" | b"named-expression" => {
                            if let Some(scope) = &active_scope {
                                definitions.push(Self::parse_named_definition(
                                    reader.resolver(),
                                    reader.decoder(),
                                    &element,
                                    scope.clone(),
                                )?);
                            }
                        },
                        _ => {},
                    }
                },
                (namespace, Event::End(element)) if Self::is_table_namespace(&namespace) => {
                    match element.local_name().as_ref() {
                        b"named-expressions" => active_scope = None,
                        b"table" => {
                            table_stack.pop();
                        },
                        _ => {},
                    }
                },
                (_, Event::Eof) => break,
                _ => {},
            }
            buf.clear();
        }

        Ok(definitions)
    }

    pub(super) fn is_table_namespace(namespace: &ResolveResult<'_>) -> bool {
        matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == TABLE_NAMESPACE)
    }

    pub(super) fn parse_named_definition(
        resolver: &NamespaceResolver,
        decoder: Decoder,
        element: &BytesStart<'_>,
        scope: NamedDefinitionScope,
    ) -> Result<NamedDefinition> {
        let name = Self::required_table_attribute(resolver, decoder, element, b"name")?;
        let base_cell_address =
            Self::table_attribute(resolver, decoder, element, b"base-cell-address")?;

        match element.local_name().as_ref() {
            b"named-range" => {
                let cell_range_address = Self::required_table_attribute(
                    resolver,
                    decoder,
                    element,
                    b"cell-range-address",
                )?;
                let mut range = NamedRange::new(name, cell_range_address, scope)?;
                range.base_cell_address = base_cell_address;
                if let Some(usable_as) =
                    Self::table_attribute(resolver, decoder, element, b"range-usable-as")?
                {
                    if usable_as.is_empty() {
                        return Err(Error::InvalidFormat(
                            "table:range-usable-as must not be empty".to_string(),
                        ));
                    }
                    if usable_as != "none" {
                        for token in usable_as.split_whitespace() {
                            let usage = NamedRangeUsage::parse(token)?;
                            if !range.usable_as.contains(&usage) {
                                range.usable_as.push(usage);
                            }
                        }
                    }
                }
                range.validate()?;
                Ok(range.into())
            },
            b"named-expression" => {
                let expression =
                    Self::required_table_attribute(resolver, decoder, element, b"expression")?;
                let formula_namespace_uri = Self::formula_namespace_uri(resolver, &expression)?;
                let mut expression = if let Some(uri) = formula_namespace_uri {
                    NamedExpression::new_with_namespace(name, expression, uri, scope)?
                } else {
                    NamedExpression::new(name, expression, scope)?
                };
                expression.base_cell_address = base_cell_address;
                expression.validate()?;
                Ok(expression.into())
            },
            _ => Err(Error::InvalidFormat(
                "unexpected named definition element".to_string(),
            )),
        }
    }

    pub(super) fn required_table_attribute(
        resolver: &NamespaceResolver,
        decoder: Decoder,
        element: &BytesStart<'_>,
        local_name: &[u8],
    ) -> Result<String> {
        Self::table_attribute(resolver, decoder, element, local_name)?.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{} is missing required table:{} attribute",
                String::from_utf8_lossy(element.local_name().as_ref()),
                String::from_utf8_lossy(local_name)
            ))
        })
    }

    pub(super) fn table_attribute(
        resolver: &NamespaceResolver,
        decoder: Decoder,
        element: &BytesStart<'_>,
        local_name: &[u8],
    ) -> Result<Option<String>> {
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let (namespace, local) = resolver.resolve_attribute(attribute.key);
            if Self::is_table_namespace(&namespace) && local.as_ref() == local_name {
                let value = attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid XML attribute value: {error}"))
                    })?;
                return Ok(Some(value.into_owned()));
            }
        }
        Ok(None)
    }

    pub(super) fn formula_namespace_uri(
        resolver: &NamespaceResolver,
        expression: &str,
    ) -> Result<Option<String>> {
        let Some((prefix, remainder)) = expression.split_once(':') else {
            return Ok(None);
        };
        if prefix.is_empty() || !remainder.starts_with('=') {
            return Ok(None);
        }

        for (declaration, namespace) in resolver.bindings() {
            if let PrefixDeclaration::Named(candidate) = declaration
                && candidate == prefix.as_bytes()
            {
                return String::from_utf8(namespace.as_ref().to_vec())
                    .map(Some)
                    .map_err(|_| {
                        Error::InvalidFormat(format!(
                            "formula namespace for prefix '{prefix}' is not UTF-8"
                        ))
                    });
            }
        }

        Err(Error::InvalidFormat(format!(
            "formula prefix '{prefix}' is not bound to a namespace"
        )))
    }

    /// Extract table name from table:table element
    pub(super) fn extract_table_name(
        e: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<String> {
        for attr_result in e.attributes() {
            let attr = attr_result
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            if Self::attribute_name_is(attr.key.as_ref(), namespaces, TABLE_NAMESPACE_URI, "name") {
                return attr
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .map(|value| value.into_owned())
                    .map_err(|error| Error::InvalidFormat(format!("invalid table name: {error}")));
            }
        }
        Ok("Sheet1".to_string()) // Default name
    }

    pub(super) fn parse_repeated(
        e: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
        local_name: &str,
    ) -> Result<usize> {
        for attribute in e.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                TABLE_NAMESPACE_URI,
                local_name,
            ) {
                return Self::parse_positive_usize(&attribute, decoder, local_name);
            }
        }
        Ok(1)
    }

    pub(super) fn parse_structural_attributes(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<(Option<String>, Option<String>, TableVisibility)> {
        let mut style_name = None;
        let mut default_cell_style_name = None;
        let mut visibility = TableVisibility::Visible;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_table = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    TABLE_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_table("style-name") {
                style_name = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "table:style-name",
                )?);
            } else if is_table("default-cell-style-name") {
                default_cell_style_name = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "table:default-cell-style-name",
                )?);
            } else if is_table("visibility") {
                visibility = TableVisibility::parse(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "table:visibility",
                )?)?;
            }
        }
        Ok((style_name, default_cell_style_name, visibility))
    }

    pub(super) fn parse_group_display(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<bool> {
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                TABLE_NAMESPACE_URI,
                "display",
            ) {
                return Self::parse_bool_attribute(&attribute, decoder);
            }
        }
        Ok(true)
    }

    pub(super) fn parse_sheet_formatting(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<(SheetStyle, SheetPrintSettings)> {
        let mut style = SheetStyle::default();
        let mut print = SheetPrintSettings::default();
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_table = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    TABLE_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_table("style-name") {
                style.style_name = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "table:style-name",
                )?);
            } else if is_table("template-name") {
                style.template_name = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "table:template-name",
                )?);
            } else if is_table("use-first-row-styles") {
                style.usage.use_first_row_styles =
                    Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_table("use-last-row-styles") {
                style.usage.use_last_row_styles =
                    Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_table("use-first-column-styles") {
                style.usage.use_first_column_styles =
                    Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_table("use-last-column-styles") {
                style.usage.use_last_column_styles =
                    Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_table("use-banding-rows-styles") {
                style.usage.use_banding_row_styles =
                    Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_table("use-banding-columns-styles") {
                style.usage.use_banding_column_styles =
                    Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_table("print") {
                print.printable = Self::parse_bool_attribute(&attribute, decoder)?;
            } else if is_table("print-ranges") {
                print.ranges = split_cell_range_addresses(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "table:print-ranges",
                )?)?;
            }
        }
        Ok((style, print))
    }

    pub(super) fn parse_scenario(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<SheetScenario> {
        let mut ranges = None;
        let mut is_active = None;
        let mut display_border = None;
        let mut border_color = None;
        let mut copy_back = None;
        let mut copy_styles = None;
        let mut copy_formulas = None;
        let mut comment = None;
        let mut protected = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_table = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    TABLE_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_table("scenario-ranges") {
                ranges = Some(split_cell_range_addresses(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "table:scenario-ranges",
                )?)?);
            } else if is_table("is-active") {
                is_active = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_table("display-border") {
                display_border = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_table("border-color") {
                border_color = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "table:border-color",
                )?);
            } else if is_table("copy-back") {
                copy_back = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_table("copy-styles") {
                copy_styles = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_table("copy-formulas") {
                copy_formulas = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_table("comment") {
                comment = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "table:comment",
                )?);
            } else if is_table("protected") {
                protected = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            }
        }
        let scenario = SheetScenario {
            ranges: ranges.ok_or_else(|| {
                Error::InvalidFormat("table:scenario requires table:scenario-ranges".to_string())
            })?,
            is_active: is_active.ok_or_else(|| {
                Error::InvalidFormat("table:scenario requires table:is-active".to_string())
            })?,
            display_border,
            border_color,
            copy_back,
            copy_styles,
            copy_formulas,
            comment,
            protected,
        };
        validate_scenario(&scenario)?;
        Ok(scenario)
    }

    /// Parse the attributes of a `text:a` hyperlink inside cell content.
    ///
    /// The visible link text is collected separately while the element's
    /// subtree is read, so the returned hyperlink has empty text.
    pub(super) fn parse_hyperlink(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<Link> {
        let mut href = None;
        let mut name = None;
        let mut title = None;
        let mut target_frame_name = None;
        let mut show = None;
        let mut actuate = None;
        let mut style_name = None;
        let mut visited_style_name = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let decode = |name| Self::decode_attribute(&attribute, decoder, name);
            if Self::attribute_name_is(attribute.key.as_ref(), namespaces, XLINK_NAMESPACE, "href")
            {
                href = Some(decode("xlink:href")?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                XLINK_NAMESPACE,
                "type",
            ) {
                let value = decode("xlink:type")?;
                if value != "simple" {
                    return Err(Error::InvalidFormat(format!(
                        "invalid hyperlink xlink:type '{value}'"
                    )));
                }
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                OFFICE_NAMESPACE,
                "name",
            ) {
                name = Some(decode("office:name")?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                OFFICE_NAMESPACE,
                "title",
            ) {
                title = Some(decode("office:title")?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                OFFICE_NAMESPACE,
                "target-frame-name",
            ) {
                target_frame_name = Some(decode("office:target-frame-name")?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                XLINK_NAMESPACE,
                "show",
            ) {
                let value = decode("xlink:show")?;
                show = Some(TextHyperlinkShow::parse(&value).ok_or_else(|| {
                    Error::InvalidFormat(format!("invalid hyperlink xlink:show '{value}'"))
                })?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                XLINK_NAMESPACE,
                "actuate",
            ) {
                let value = decode("xlink:actuate")?;
                actuate = Some(TextHyperlinkActuate::parse(&value).ok_or_else(|| {
                    Error::InvalidFormat(format!("invalid hyperlink xlink:actuate '{value}'"))
                })?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                TEXT_NAMESPACE,
                "style-name",
            ) {
                style_name = Some(decode("text:style-name")?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                TEXT_NAMESPACE,
                "visited-style-name",
            ) {
                visited_style_name = Some(decode("text:visited-style-name")?);
            }
        }
        let href = href.ok_or_else(|| {
            Error::InvalidFormat("text:a hyperlink requires xlink:href".to_string())
        })?;
        Ok(Link {
            href,
            text: String::new(),
            range: 0..0,
            name,
            title,
            target_frame_name,
            show,
            actuate,
            style_name,
            visited_style_name,
        })
    }

    pub(super) fn parse_table_source(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<SheetTableSource> {
        let mut link_type = None;
        let mut href = None;
        let mut actuate_on_request = false;
        let mut mode = None;
        let mut table_name = None;
        let mut filter_name = None;
        let mut filter_options = None;
        let mut refresh_delay = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let decode = |name| Self::decode_attribute(&attribute, decoder, name);
            if Self::attribute_name_is(attribute.key.as_ref(), namespaces, XLINK_NAMESPACE, "type")
            {
                link_type = Some(decode("xlink:type")?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                XLINK_NAMESPACE,
                "href",
            ) {
                href = Some(decode("xlink:href")?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                XLINK_NAMESPACE,
                "actuate",
            ) {
                let value = decode("xlink:actuate")?;
                if value != "onRequest" {
                    return Err(Error::InvalidFormat(format!(
                        "invalid table source xlink:actuate '{value}'"
                    )));
                }
                actuate_on_request = true;
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                TABLE_NAMESPACE_URI,
                "mode",
            ) {
                mode = Some(TableSourceMode::parse(&decode("table:mode")?)?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                TABLE_NAMESPACE_URI,
                "table-name",
            ) {
                table_name = Some(decode("table:table-name")?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                TABLE_NAMESPACE_URI,
                "filter-name",
            ) {
                filter_name = Some(decode("table:filter-name")?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                TABLE_NAMESPACE_URI,
                "filter-options",
            ) {
                filter_options = Some(decode("table:filter-options")?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                TABLE_NAMESPACE_URI,
                "refresh-delay",
            ) {
                refresh_delay = Some(decode("table:refresh-delay")?);
            }
        }
        let link_type = link_type.ok_or_else(|| {
            Error::InvalidFormat("table:table-source requires xlink:type".to_string())
        })?;
        if link_type != "simple" {
            return Err(Error::InvalidFormat(format!(
                "invalid table source xlink:type '{link_type}'"
            )));
        }
        let source = SheetTableSource {
            href: href.ok_or_else(|| {
                Error::InvalidFormat("table:table-source requires xlink:href".to_string())
            })?,
            mode,
            table_name,
            actuate_on_request,
            filter_name,
            filter_options,
            refresh_delay,
        };
        validate_table_source(&source)?;
        Ok(source)
    }

    pub(super) fn parse_cell_range_source(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<CellRangeSource> {
        let mut name = None;
        let mut href = None;
        let mut link_type = None;
        let mut rows = None;
        let mut columns = None;
        let mut actuate_on_request = false;
        let mut filter_name = None;
        let mut filter_options = None;
        let mut refresh_delay = None;

        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let decode =
                |qualified_name| Self::decode_attribute(&attribute, decoder, qualified_name);
            let is_table = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    TABLE_NAMESPACE_URI,
                    local_name,
                )
            };
            let is_xlink = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    XLINK_NAMESPACE,
                    local_name,
                )
            };
            if is_table("name") {
                name = Some(decode("table:name")?);
            } else if is_table("last-row-spanned") {
                rows = Some(Self::parse_positive_usize(
                    &attribute,
                    decoder,
                    "last-row-spanned",
                )?);
            } else if is_table("last-column-spanned") {
                columns = Some(Self::parse_positive_usize(
                    &attribute,
                    decoder,
                    "last-column-spanned",
                )?);
            } else if is_table("filter-name") {
                filter_name = Some(decode("table:filter-name")?);
            } else if is_table("filter-options") {
                filter_options = Some(decode("table:filter-options")?);
            } else if is_table("refresh-delay") {
                refresh_delay = Some(decode("table:refresh-delay")?);
            } else if is_xlink("type") {
                link_type = Some(decode("xlink:type")?);
            } else if is_xlink("href") {
                href = Some(decode("xlink:href")?);
            } else if is_xlink("actuate") {
                let value = decode("xlink:actuate")?;
                if value != "onRequest" {
                    return Err(Error::InvalidFormat(format!(
                        "invalid cell range source xlink:actuate '{value}'"
                    )));
                }
                actuate_on_request = true;
            }
        }

        let link_type = link_type.ok_or_else(|| {
            Error::InvalidFormat("table:cell-range-source requires xlink:type".to_string())
        })?;
        if link_type != "simple" {
            return Err(Error::InvalidFormat(format!(
                "invalid cell range source xlink:type '{link_type}'"
            )));
        }
        let mut source = CellRangeSource::new(
            name.ok_or_else(|| {
                Error::InvalidFormat("table:cell-range-source requires table:name".to_string())
            })?,
            href.ok_or_else(|| {
                Error::InvalidFormat("table:cell-range-source requires xlink:href".to_string())
            })?,
            rows.ok_or_else(|| {
                Error::InvalidFormat(
                    "table:cell-range-source requires table:last-row-spanned".to_string(),
                )
            })?,
            columns.ok_or_else(|| {
                Error::InvalidFormat(
                    "table:cell-range-source requires table:last-column-spanned".to_string(),
                )
            })?,
        )?;
        source.set_actuate_on_request(actuate_on_request);
        source.set_filter_name(filter_name);
        source.set_filter_options(filter_options);
        source.set_refresh_delay(refresh_delay)?;
        Ok(source)
    }

    pub(super) fn parse_detective_child(
        detective: &mut CellDetective,
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<()> {
        if Self::element_name_is(
            element.name().as_ref(),
            namespaces,
            TABLE_NAMESPACE_URI,
            "highlighted-range",
        ) {
            if !detective.operations().is_empty() {
                return Err(Error::InvalidFormat(
                    "table:highlighted-range must precede table:operation".to_string(),
                ));
            }
            detective
                .add_highlighted_range(Self::parse_detective_range(element, decoder, namespaces)?);
        } else if Self::element_name_is(
            element.name().as_ref(),
            namespaces,
            TABLE_NAMESPACE_URI,
            "operation",
        ) {
            detective.add_operation(Self::parse_detective_operation(
                element, decoder, namespaces,
            )?);
        } else {
            return Err(Error::InvalidFormat(
                "table:detective contains an unsupported child element".to_string(),
            ));
        }
        Ok(())
    }

    pub(super) fn parse_detective_range(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<DetectiveHighlightedRange> {
        let mut address = None;
        let mut direction = None;
        let mut contains_error = None;
        let mut marked_invalid = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_table = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    TABLE_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_table("cell-range-address") {
                address = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "table:cell-range-address",
                )?);
            } else if is_table("direction") {
                direction = Some(DetectiveDirection::parse(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "table:direction",
                )?)?);
            } else if is_table("contains-error") {
                contains_error = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_table("marked-invalid") {
                marked_invalid = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            }
        }

        if let Some(marked_invalid) = marked_invalid {
            if address.is_some() || direction.is_some() || contains_error.is_some() {
                return Err(Error::InvalidFormat(
                    "invalid detective ranges cannot contain directional range attributes"
                        .to_string(),
                ));
            }
            Ok(DetectiveHighlightedRange::invalid(marked_invalid))
        } else {
            DetectiveHighlightedRange::valid(
                address,
                direction.ok_or_else(|| {
                    Error::InvalidFormat(
                        "table:highlighted-range requires table:direction".to_string(),
                    )
                })?,
                contains_error,
            )
        }
    }

    pub(super) fn parse_detective_operation(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<DetectiveOperation> {
        let mut kind = None;
        let mut index = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                TABLE_NAMESPACE_URI,
                "name",
            ) {
                kind = Some(DetectiveOperationKind::parse(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "table:name",
                )?)?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                TABLE_NAMESPACE_URI,
                "index",
            ) {
                let value = Self::decode_attribute(&attribute, decoder, "table:index")?;
                index = Some(value.parse::<usize>().map_err(|_| {
                    Error::InvalidFormat(format!(
                        "invalid non-negative detective operation index '{value}'"
                    ))
                })?);
            }
        }
        Ok(DetectiveOperation::new(
            kind.ok_or_else(|| {
                Error::InvalidFormat("table:operation requires table:name".to_string())
            })?,
            index.ok_or_else(|| {
                Error::InvalidFormat("table:operation requires table:index".to_string())
            })?,
        ))
    }

    pub(super) fn parse_column(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<(Column, usize)> {
        let repeated =
            Self::parse_repeated(element, decoder, namespaces, "number-columns-repeated")?;
        let (style_name, default_cell_style_name, visibility) =
            Self::parse_structural_attributes(element, decoder, namespaces)?;
        Ok((
            Column {
                index: 0,
                style_name,
                default_cell_style_name,
                visibility,
            },
            repeated,
        ))
    }

    pub(super) fn parse_positive_usize(
        attribute: &quick_xml::events::attributes::Attribute<'_>,
        decoder: Decoder,
        attribute_name: &str,
    ) -> Result<usize> {
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid table:{attribute_name}: {error}"))
            })?;
        let parsed = value.parse::<usize>().map_err(|_| {
            Error::InvalidFormat(format!("invalid table:{attribute_name} value '{value}'"))
        })?;
        if parsed == 0 {
            return Err(Error::InvalidFormat(format!(
                "table:{attribute_name} must be positive"
            )));
        }
        Ok(parsed)
    }

    /// Parse cell attributes and create a CellBuilder
    pub(super) fn parse_cell_attributes(
        e: &quick_xml::events::BytesStart,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<CellBuilder> {
        let mut value_type = None;
        let mut value_str = None;
        let mut currency = None;
        let mut formula = None;
        let mut validation_name = None;
        let mut style_name = None;
        let mut protect = None;
        let mut protected = None;
        let mut repeated = 1;
        let mut row_span = 1;
        let mut column_span = 1;
        let mut matrix_row_span = None;
        let mut matrix_column_span = None;

        for attr_result in e.attributes() {
            let attr = attr_result
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_office = |local_name| {
                Self::attribute_name_is(attr.key.as_ref(), namespaces, OFFICE_NAMESPACE, local_name)
            };
            let is_table = |local_name| {
                Self::attribute_name_is(
                    attr.key.as_ref(),
                    namespaces,
                    TABLE_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_office("value-type") {
                value_type = Some(Self::decode_attribute(&attr, decoder, "office:value-type")?);
            } else if is_office("value") {
                value_str = Some(Self::decode_attribute(&attr, decoder, "office:value")?);
            } else if is_office("currency") {
                currency = Some(Self::decode_attribute(&attr, decoder, "office:currency")?);
            } else if is_table("formula") {
                formula = Some(Self::decode_attribute(&attr, decoder, "table:formula")?);
            } else if is_table("content-validation-name") {
                validation_name = Some(Self::decode_attribute(
                    &attr,
                    decoder,
                    "table:content-validation-name",
                )?);
            } else if is_table("style-name") {
                style_name = Some(Self::decode_attribute(&attr, decoder, "table:style-name")?);
            } else if is_table("protect") {
                protect = Some(Self::parse_bool_attribute(&attr, decoder)?);
            } else if is_table("protected") {
                protected = Some(Self::parse_bool_attribute(&attr, decoder)?);
            } else if is_table("number-columns-repeated") {
                repeated = Self::parse_positive_usize(&attr, decoder, "number-columns-repeated")?;
            } else if is_table("number-rows-spanned") {
                row_span = Self::parse_positive_usize(&attr, decoder, "number-rows-spanned")?;
            } else if is_table("number-columns-spanned") {
                column_span = Self::parse_positive_usize(&attr, decoder, "number-columns-spanned")?;
            } else if is_table("number-matrix-rows-spanned") {
                matrix_row_span = Some(Self::parse_positive_usize(
                    &attr,
                    decoder,
                    "number-matrix-rows-spanned",
                )?);
            } else if is_table("number-matrix-columns-spanned") {
                matrix_column_span = Some(Self::parse_positive_usize(
                    &attr,
                    decoder,
                    "number-matrix-columns-spanned",
                )?);
            }
        }

        Ok(CellBuilder::from_parts(
            value_type,
            value_str,
            currency,
            formula,
            validation_name,
            style_name,
            if matrix_row_span.is_some() || matrix_column_span.is_some() {
                Some(CellMatrixSpan::new(
                    matrix_row_span.unwrap_or(1),
                    matrix_column_span.unwrap_or(1),
                )?)
            } else {
                None
            },
            protect,
            protected,
            repeated,
            if row_span == 1 && column_span == 1 {
                CellMerge::None
            } else {
                CellMerge::Span {
                    rows: NonZeroUsize::new(row_span).expect("positive row span was checked"),
                    columns: NonZeroUsize::new(column_span)
                        .expect("positive column span was checked"),
                }
            },
        ))
    }

    pub(super) fn parse_bool_attribute(
        attribute: &quick_xml::events::attributes::Attribute<'_>,
        decoder: Decoder,
    ) -> Result<bool> {
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::InvalidFormat(format!("invalid Boolean attribute: {error}")))?;
        match value.as_ref() {
            "true" | "1" => Ok(true),
            "false" | "0" => Ok(false),
            _ => Err(Error::InvalidFormat(format!(
                "invalid Boolean attribute value '{value}'"
            ))),
        }
    }

    /// Parse the `calcext:target-range-address` of a `calcext:conditional-format`.
    pub(super) fn parse_conditional_format_ranges(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<Vec<String>> {
        let mut ranges = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                CALCEXT_NAMESPACE_URI,
                "target-range-address",
            ) {
                if ranges.is_some() {
                    return Err(Error::InvalidFormat(
                        "duplicate calcext:target-range-address attribute".to_string(),
                    ));
                }
                ranges = Some(split_cell_range_addresses(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:target-range-address",
                )?)?);
            }
        }
        let ranges = ranges.ok_or_else(|| {
            Error::InvalidFormat(
                "calcext:conditional-format requires calcext:target-range-address".to_string(),
            )
        })?;
        if ranges.is_empty() {
            return Err(Error::InvalidFormat(
                "calcext:target-range-address requires at least one range".to_string(),
            ));
        }
        Ok(ranges)
    }

    /// Parse one inert `calcext:condition` rule from its attributes.
    pub(super) fn parse_calcext_condition(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<ConditionalFormatCondition> {
        let mut condition = None;
        let mut apply_style_name = None;
        let mut base_cell_address = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_calcext = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    CALCEXT_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_calcext("value") {
                condition = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:value",
                )?);
            } else if is_calcext("apply-style-name") {
                apply_style_name = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:apply-style-name",
                )?);
            } else if is_calcext("base-cell-address") {
                base_cell_address = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:base-cell-address",
                )?);
            }
        }
        let rule = ConditionalFormatCondition {
            condition: condition.ok_or_else(|| {
                Error::InvalidFormat("calcext:condition requires calcext:value".to_string())
            })?,
            apply_style_name: apply_style_name.ok_or_else(|| {
                Error::InvalidFormat(
                    "calcext:condition requires calcext:apply-style-name".to_string(),
                )
            })?,
            base_cell_address,
        };
        validate_condition(&rule)?;
        Ok(rule)
    }

    /// Parse one inert `calcext:color-scale-entry` from its attributes.
    pub(super) fn parse_color_scale_entry(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<ConditionalColorScaleEntry> {
        let mut entry_type = None;
        let mut value = None;
        let mut color = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_calcext = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    CALCEXT_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_calcext("type") {
                entry_type = Some(ConditionalFormatEntryType::parse(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:type",
                )?)?);
            } else if is_calcext("value") {
                value = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:value",
                )?);
            } else if is_calcext("color") {
                color = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:color",
                )?);
            }
        }
        let entry = ConditionalColorScaleEntry {
            entry_type: entry_type.ok_or_else(|| {
                Error::InvalidFormat("calcext:color-scale-entry requires calcext:type".to_string())
            })?,
            value: value.ok_or_else(|| {
                Error::InvalidFormat("calcext:color-scale-entry requires calcext:value".to_string())
            })?,
            color: color.ok_or_else(|| {
                Error::InvalidFormat("calcext:color-scale-entry requires calcext:color".to_string())
            })?,
        };
        validate_color_scale_entry(&entry)?;
        Ok(entry)
    }

    /// Parse one inert data-bar `calcext:formatting-entry` from its attributes.
    pub(super) fn parse_data_bar_entry(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<ConditionalDataBarEntry> {
        let (entry_type, value) = Self::parse_formatting_entry(element, decoder, namespaces)?;
        let entry = ConditionalDataBarEntry { entry_type, value };
        validate_data_bar_entry(&entry)?;
        Ok(entry)
    }

    /// Parse one inert icon-set `calcext:formatting-entry` from its attributes.
    pub(super) fn parse_icon_set_entry(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<ConditionalIconSetEntry> {
        let mut greater_equal = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                CALCEXT_NAMESPACE_URI,
                "greater-equal",
            ) {
                greater_equal = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            }
        }
        let (entry_type, value) = Self::parse_formatting_entry(element, decoder, namespaces)?;
        let entry = ConditionalIconSetEntry {
            entry_type,
            value,
            greater_equal,
        };
        validate_icon_set_entry(&entry)?;
        Ok(entry)
    }

    /// Parse the shared `calcext:type`/`calcext:value` pair of a formatting entry.
    pub(super) fn parse_formatting_entry(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<(ConditionalFormatEntryType, String)> {
        let mut entry_type = None;
        let mut value = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_calcext = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    CALCEXT_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_calcext("type") {
                entry_type = Some(ConditionalFormatEntryType::parse(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:type",
                )?)?);
            } else if is_calcext("value") {
                value = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:value",
                )?);
            }
        }
        Ok((
            entry_type.ok_or_else(|| {
                Error::InvalidFormat("calcext:formatting-entry requires calcext:type".to_string())
            })?,
            value.ok_or_else(|| {
                Error::InvalidFormat("calcext:formatting-entry requires calcext:value".to_string())
            })?,
        ))
    }

    /// Parse the attributes of an inert `calcext:data-bar` element.
    pub(super) fn parse_data_bar_attributes(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<ConditionalDataBar> {
        let mut data_bar = ConditionalDataBar::new(Vec::new());
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_calcext = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    CALCEXT_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_calcext("positive-color") {
                data_bar.positive_color = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:positive-color",
                )?);
            } else if is_calcext("negative-color") {
                data_bar.negative_color = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:negative-color",
                )?);
            } else if is_calcext("gradient") {
                data_bar.gradient = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("axis-position") {
                data_bar.axis_position = Some(DataBarAxisPosition::parse(
                    &Self::decode_attribute(&attribute, decoder, "calcext:axis-position")?,
                )?);
            } else if is_calcext("show-value") {
                data_bar.show_value = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("axis-color") {
                data_bar.axis_color = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:axis-color",
                )?);
            } else if is_calcext("min-length") {
                data_bar.min_length = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:min-length",
                )?);
            } else if is_calcext("max-length") {
                data_bar.max_length = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:max-length",
                )?);
            }
        }
        validate_data_bar_attributes(&data_bar)?;
        Ok(data_bar)
    }

    /// Parse one inert `calcext:custom-iconset` assignment from its attributes.
    pub(super) fn parse_custom_icon(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<ConditionalCustomIcon> {
        let mut icon_set_type = None;
        let mut index = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_calcext = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    CALCEXT_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_calcext("custom-iconset-name") {
                icon_set_type = Some(IconSetType::parse(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:custom-iconset-name",
                )?)?);
            } else if is_calcext("custom-iconset-index") {
                let lexical =
                    Self::decode_attribute(&attribute, decoder, "calcext:custom-iconset-index")?;
                index = Some(lexical.parse::<u32>().map_err(|_| {
                    Error::InvalidFormat(format!(
                        "calcext:custom-iconset-index requires a non-negative integer, found '{lexical}'"
                    ))
                })?);
            }
        }
        Ok(ConditionalCustomIcon {
            icon_set_type: icon_set_type.ok_or_else(|| {
                Error::InvalidFormat(
                    "calcext:custom-iconset requires calcext:custom-iconset-name".to_string(),
                )
            })?,
            index: index.ok_or_else(|| {
                Error::InvalidFormat(
                    "calcext:custom-iconset requires calcext:custom-iconset-index".to_string(),
                )
            })?,
        })
    }

    /// Parse the attributes of an inert `calcext:icon-set` element.
    pub(super) fn parse_icon_set_attributes(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<ConditionalIconSet> {
        let mut icon_set_type = None;
        let mut show_value = None;
        let mut custom = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_calcext = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    CALCEXT_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_calcext("icon-set-type") {
                icon_set_type = Some(IconSetType::parse(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:icon-set-type",
                )?)?);
            } else if is_calcext("show-value") {
                show_value = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("custom") {
                custom = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            }
        }
        let icon_set = ConditionalIconSet {
            icon_set_type: icon_set_type.ok_or_else(|| {
                Error::InvalidFormat("calcext:icon-set requires calcext:icon-set-type".to_string())
            })?,
            show_value,
            custom,
            custom_icons: Vec::new(),
            entries: Vec::new(),
        };
        Ok(icon_set)
    }

    /// Parse one inert `calcext:date-is` rule from its attributes.
    pub(super) fn parse_date_is(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<ConditionalDateIs> {
        let mut date = None;
        let mut style = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_calcext = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    CALCEXT_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_calcext("date") {
                date = Some(ConditionalDateType::parse(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:date",
                )?)?);
            } else if is_calcext("style") {
                style = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:style",
                )?);
            }
        }
        let date_is = ConditionalDateIs {
            date: date.ok_or_else(|| {
                Error::InvalidFormat("calcext:date-is requires calcext:date".to_string())
            })?,
            style: style.ok_or_else(|| {
                Error::InvalidFormat("calcext:date-is requires calcext:style".to_string())
            })?,
        };
        validate_date_is(&date_is)?;
        Ok(date_is)
    }

    /// Parse one inert `calcext:sparkline` cell assignment from its attributes.
    pub(super) fn parse_sparkline(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<Sparkline> {
        let mut cell_address = None;
        let mut data_ranges = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_calcext = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    CALCEXT_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_calcext("cell-address") {
                cell_address = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:cell-address",
                )?);
            } else if is_calcext("data-range") {
                data_ranges = Some(split_cell_range_addresses(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:data-range",
                )?)?);
            }
        }
        let sparkline = Sparkline {
            cell_address: cell_address.ok_or_else(|| {
                Error::InvalidFormat("calcext:sparkline requires calcext:cell-address".to_string())
            })?,
            data_ranges: data_ranges.ok_or_else(|| {
                Error::InvalidFormat("calcext:sparkline requires calcext:data-range".to_string())
            })?,
        };
        validate_sparkline(&sparkline)?;
        Ok(sparkline)
    }

    /// Parse the attributes of an inert `calcext:sparkline-group` element.
    pub(super) fn parse_sparkline_group_attributes(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<SparklineGroup> {
        let mut group = SparklineGroup::new(Vec::new());
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_calcext = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    CALCEXT_NAMESPACE_URI,
                    local_name,
                )
            };
            let decode = |name| Self::decode_attribute(&attribute, decoder, name);
            if is_calcext("id") {
                group.id = Some(decode("calcext:id")?);
            } else if is_calcext("type") {
                group.sparkline_type = Some(SparklineType::parse(&decode("calcext:type")?)?);
            } else if is_calcext("line-width") {
                group.line_width = Some(decode("calcext:line-width")?);
            } else if is_calcext("date-axis") {
                group.flags.date_axis = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("display-empty-cells-as") {
                group.display_empty_cells_as = Some(SparklineEmptyCells::parse(&decode(
                    "calcext:display-empty-cells-as",
                )?)?);
            } else if is_calcext("markers") {
                group.flags.markers = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("high") {
                group.flags.high = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("low") {
                group.flags.low = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("first") {
                group.flags.first = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("last") {
                group.flags.last = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("negative") {
                group.flags.negative = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("display-x-axis") {
                group.flags.display_x_axis = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("display-hidden") {
                group.flags.display_hidden = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("min-axis-type") {
                group.min_axis_type =
                    Some(SparklineAxisType::parse(&decode("calcext:min-axis-type")?)?);
            } else if is_calcext("max-axis-type") {
                group.max_axis_type =
                    Some(SparklineAxisType::parse(&decode("calcext:max-axis-type")?)?);
            } else if is_calcext("right-to-left") {
                group.flags.right_to_left = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("manual-max") {
                group.manual_max = Some(decode("calcext:manual-max")?);
            } else if is_calcext("manual-min") {
                group.manual_min = Some(decode("calcext:manual-min")?);
            } else if is_calcext("color-series") {
                group.colors.series = Some(decode("calcext:color-series")?);
            } else if is_calcext("color-negative") {
                group.colors.negative = Some(decode("calcext:color-negative")?);
            } else if is_calcext("color-axis") {
                group.colors.axis = Some(decode("calcext:color-axis")?);
            } else if is_calcext("color-markers") {
                group.colors.markers = Some(decode("calcext:color-markers")?);
            } else if is_calcext("color-first") {
                group.colors.first = Some(decode("calcext:color-first")?);
            } else if is_calcext("color-last") {
                group.colors.last = Some(decode("calcext:color-last")?);
            } else if is_calcext("color-high") {
                group.colors.high = Some(decode("calcext:color-high")?);
            } else if is_calcext("color-low") {
                group.colors.low = Some(decode("calcext:color-low")?);
            }
        }
        validate_sparkline_group_attributes(&group)?;
        Ok(group)
    }

    /// Parse one inert `calcext:sparkline-*-complex-color` from its attributes.
    pub(super) fn parse_complex_color(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<SparklineComplexColor> {
        let mut theme_type = None;
        let mut color_type = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_loext = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    LOEXT_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_loext("theme-type") {
                theme_type = Some(ThemeColorType::parse(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "loext:theme-type",
                )?)?);
            } else if is_loext("color-type") {
                color_type = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "loext:color-type",
                )?);
            }
        }
        if let Some(color_type) = &color_type
            && color_type != "theme"
        {
            return Err(Error::InvalidFormat(format!(
                "unsupported loext:color-type value '{color_type}'"
            )));
        }
        Ok(SparklineComplexColor {
            theme_type: theme_type.ok_or_else(|| {
                Error::InvalidFormat("calcext complex color requires loext:theme-type".to_string())
            })?,
            transformations: Vec::new(),
        })
    }

    /// Parse one inert `loext:transformation` from its attributes.
    pub(super) fn parse_color_transformation(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<SparklineColorTransformation> {
        let mut transformation_type = None;
        let mut value = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_loext = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    LOEXT_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_loext("type") {
                transformation_type = Some(ColorTransformationType::parse(
                    &Self::decode_attribute(&attribute, decoder, "loext:type")?,
                )?);
            } else if is_loext("value") {
                let lexical = Self::decode_attribute(&attribute, decoder, "loext:value")?;
                let parsed: i32 = lexical.parse().map_err(|_| {
                    Error::InvalidFormat(format!(
                        "loext:value requires an integer value, found '{lexical}'"
                    ))
                })?;
                value = Some(i16::try_from(parsed).map_err(|_| {
                    Error::InvalidFormat(format!(
                        "loext:value '{lexical}' is outside the supported range"
                    ))
                })?);
            }
        }
        Ok(SparklineColorTransformation {
            transformation_type: transformation_type.ok_or_else(|| {
                Error::InvalidFormat("loext:transformation requires loext:type".to_string())
            })?,
            value: value.ok_or_else(|| {
                Error::InvalidFormat("loext:transformation requires loext:value".to_string())
            })?,
        })
    }

    pub(super) fn decode_attribute(
        attribute: &quick_xml::events::attributes::Attribute<'_>,
        decoder: Decoder,
        name: &str,
    ) -> Result<String> {
        attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map(|value| value.into_owned())
            .map_err(|error| Error::InvalidFormat(format!("invalid {name}: {error}")))
    }
}

impl RowBuilder {
    pub(super) fn from_element(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<Self> {
        let repeated =
            Parser::parse_repeated(element, decoder, namespaces, "number-rows-repeated")?;
        let (style_name, default_cell_style_name, visibility) =
            Parser::parse_structural_attributes(element, decoder, namespaces)?;
        Ok(Self::from_parts(
            repeated,
            style_name,
            default_cell_style_name,
            visibility,
        ))
    }
}
