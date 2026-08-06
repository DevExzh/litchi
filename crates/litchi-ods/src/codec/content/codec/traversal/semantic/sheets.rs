use super::super::{
    BTreeMap, Builder, CALCEXT_NAMESPACE_URI, COMPLEX_COLOR_SLOTS, CellBuilder, CellDetective,
    CellTextContentBuilder, ConditionalColorScale, ConditionalFormat, ConditionalFormatRule,
    DATA_BAR_ENTRY_COUNT, Error, Event, LOEXT_NAMESPACE_URI, MAX_COLOR_TRANSFORMATIONS,
    MAX_ENTRIES_PER_RULE, MAX_RULES_PER_FORMAT, MAX_SPARKLINES_PER_GROUP, OFFICE_NAMESPACE, Parser,
    Reader, Result, RowBuilder, Sheet, SheetBuilder, TABLE_NAMESPACE, TABLE_NAMESPACE_URI,
    TEXT_NAMESPACE, XLINK_NAMESPACE, XmlVersion, decode_reference,
    model::{
        PendingCalcextRule, PendingConditionalFormat, PendingHyperlink,
        PendingSparklineComplexColor, PendingSparklineGroup, SheetTextField,
    },
    parse_dde_source, validate_conditional_format, validate_rule, validate_sparkline_group,
};

pub(super) trait SheetTraversal {
    fn parse_sheets(xml_content: &str) -> Result<Vec<Sheet>>;
}

impl SheetTraversal for Parser {
    /// Parse all sheets from ODS content.xml
    // quick-xml exposes a streaming event source, so the format's nested parser
    // state is intentionally coordinated here without constructing a DOM.
    #[allow(clippy::cognitive_complexity)]
    fn parse_sheets(xml_content: &str) -> Result<Vec<Sheet>> {
        let mut reader = Reader::from_str(xml_content);
        let mut buf = Vec::new();
        let mut sheets = Vec::new();

        // Parser state
        let mut current_sheet: Option<SheetBuilder> = None;
        let mut current_row: Option<RowBuilder> = None;
        let mut current_cell: Option<CellBuilder> = None;
        let mut text_element_depth = 0usize;
        let mut text_content = String::new();
        let mut rich_text_builder: Option<CellTextContentBuilder> = None;
        let mut annotation_builder: Option<Builder> = None;
        let mut annotation_depth = 0usize;
        let mut pending_hyperlink: Option<PendingHyperlink> = None;
        let mut detective_builder: Option<CellDetective> = None;
        let mut detective_child_open = false;
        let mut sheet_text_field = None;
        let mut sheet_text = String::new();
        let mut document_namespaces = BTreeMap::new();
        let mut namespace_scopes = Vec::new();
        let mut element_depth = 0usize;
        let mut spreadsheet_depth = None;
        let mut current_sheet_depth = None;
        let mut sheet_dde_source_depth = None;
        let mut conditional_formats_depth = None;
        let mut pending_conditional_format: Option<PendingConditionalFormat> = None;
        let mut pending_calcext_rule: Option<PendingCalcextRule> = None;
        let mut calcext_leaf_open_depth = None;
        let mut calcext_skip_depth = None;
        let mut sparkline_groups_depth = None;
        let mut pending_sparkline_group: Option<PendingSparklineGroup> = None;
        let mut pending_sparkline_complex_color: Option<PendingSparklineComplexColor> = None;
        let mut sparkline_list_depth = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let namespace_scope =
                        Self::push_namespace_scope(e, reader.decoder(), &mut document_namespaces)?;
                    namespace_scopes.push(namespace_scope);
                    element_depth += 1;
                    if Self::element_name_is(
                        e.name().as_ref(),
                        &document_namespaces,
                        OFFICE_NAMESPACE,
                        "spreadsheet",
                    ) {
                        spreadsheet_depth = Some(element_depth);
                    }

                    if sheet_dde_source_depth.is_some() {
                        return Err(Error::InvalidFormat(
                            "office:dde-source must not contain child elements".to_string(),
                        ));
                    } else if let Some(sheet) = current_sheet.as_mut()
                        && current_row.is_none()
                        && current_sheet_depth.is_some_and(|depth| element_depth == depth + 1)
                        && e.local_name().as_ref() == b"dde-source"
                    {
                        if !Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            OFFICE_NAMESPACE,
                            "dde-source",
                        ) {
                            return Err(Error::InvalidFormat(
                                "spoofed office:dde-source namespace".to_string(),
                            ));
                        }
                        let source = parse_dde_source(e, reader.decoder(), &document_namespaces)?;
                        sheet.set_dde_source(source)?;
                        sheet_dde_source_depth = Some(element_depth);
                    } else if calcext_skip_depth.is_some() {
                        // Unknown extension content inside
                        // `calcext:conditional-formats` is skipped entirely.
                    } else if calcext_leaf_open_depth.is_some() {
                        return Err(Error::InvalidFormat(
                            "calcext rule and entry elements must not contain child elements"
                                .to_string(),
                        ));
                    } else if let Some(rule) = pending_calcext_rule.as_mut() {
                        let is_calcext = |local_name: &str| {
                            Self::element_name_is(
                                e.name().as_ref(),
                                &document_namespaces,
                                CALCEXT_NAMESPACE_URI,
                                local_name,
                            )
                        };
                        match rule {
                            PendingCalcextRule::ColorScale { entries, .. }
                                if is_calcext("color-scale-entry") =>
                            {
                                if entries.len() >= MAX_ENTRIES_PER_RULE {
                                    return Err(Error::InvalidFormat(format!(
                                        "calcext:color-scale exceeds the {MAX_ENTRIES_PER_RULE} entry safety limit"
                                    )));
                                }
                                entries.push(Self::parse_color_scale_entry(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?);
                                calcext_leaf_open_depth = Some(element_depth);
                            },
                            PendingCalcextRule::DataBar { data_bar, .. }
                                if is_calcext("formatting-entry")
                                    || is_calcext("data-bar-entry") =>
                            {
                                if data_bar.entries.len() >= DATA_BAR_ENTRY_COUNT {
                                    return Err(Error::InvalidFormat(format!(
                                        "calcext:data-bar supports exactly {DATA_BAR_ENTRY_COUNT} calcext:formatting-entry elements"
                                    )));
                                }
                                data_bar.entries.push(Self::parse_data_bar_entry(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?);
                                calcext_leaf_open_depth = Some(element_depth);
                            },
                            PendingCalcextRule::IconSet { icon_set, .. }
                                if is_calcext("formatting-entry") =>
                            {
                                if icon_set.entries.len() >= MAX_ENTRIES_PER_RULE {
                                    return Err(Error::InvalidFormat(format!(
                                        "calcext:icon-set exceeds the {MAX_ENTRIES_PER_RULE} entry safety limit"
                                    )));
                                }
                                icon_set.entries.push(Self::parse_icon_set_entry(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?);
                                calcext_leaf_open_depth = Some(element_depth);
                            },
                            PendingCalcextRule::IconSet { icon_set, .. }
                                if is_calcext("custom-iconset") =>
                            {
                                if icon_set.custom_icons.len() >= MAX_ENTRIES_PER_RULE {
                                    return Err(Error::InvalidFormat(format!(
                                        "calcext:icon-set exceeds the {MAX_ENTRIES_PER_RULE} custom icon safety limit"
                                    )));
                                }
                                icon_set.custom_icons.push(Self::parse_custom_icon(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?);
                                calcext_leaf_open_depth = Some(element_depth);
                            },
                            _ => {
                                // Unmodeled content is skipped.
                                calcext_skip_depth = Some(element_depth);
                            },
                        }
                    } else if let Some(pending) = pending_conditional_format.as_mut() {
                        let is_calcext = |local_name: &str| {
                            Self::element_name_is(
                                e.name().as_ref(),
                                &document_namespaces,
                                CALCEXT_NAMESPACE_URI,
                                local_name,
                            )
                        };
                        if is_calcext("condition") {
                            if pending.rules.len() >= MAX_RULES_PER_FORMAT {
                                return Err(Error::InvalidFormat(format!(
                                    "conditional format exceeds the {MAX_RULES_PER_FORMAT} rule safety limit"
                                )));
                            }
                            let condition = Self::parse_calcext_condition(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                            )?;
                            pending.rules.push(condition.into());
                            calcext_leaf_open_depth = Some(element_depth);
                        } else if is_calcext("color-scale") {
                            pending_calcext_rule = Some(PendingCalcextRule::ColorScale {
                                entries: Vec::new(),
                                depth: element_depth,
                            });
                        } else if is_calcext("data-bar") {
                            pending_calcext_rule = Some(PendingCalcextRule::DataBar {
                                data_bar: Self::parse_data_bar_attributes(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?,
                                depth: element_depth,
                            });
                        } else if is_calcext("icon-set") {
                            pending_calcext_rule = Some(PendingCalcextRule::IconSet {
                                icon_set: Self::parse_icon_set_attributes(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?,
                                depth: element_depth,
                            });
                        } else if is_calcext("date-is") {
                            if pending.rules.len() >= MAX_RULES_PER_FORMAT {
                                return Err(Error::InvalidFormat(format!(
                                    "conditional format exceeds the {MAX_RULES_PER_FORMAT} rule safety limit"
                                )));
                            }
                            let date_is =
                                Self::parse_date_is(e, reader.decoder(), &document_namespaces)?;
                            pending.rules.push(date_is.into());
                            calcext_leaf_open_depth = Some(element_depth);
                        } else {
                            // Unmodeled rule content is skipped.
                            calcext_skip_depth = Some(element_depth);
                        }
                    } else if conditional_formats_depth.is_some() {
                        if Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "conditional-format",
                        ) {
                            let target_range_addresses = Self::parse_conditional_format_ranges(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                            )?;
                            pending_conditional_format = Some(PendingConditionalFormat {
                                target_range_addresses,
                                rules: Vec::new(),
                                depth: element_depth,
                            });
                        } else {
                            calcext_skip_depth = Some(element_depth);
                        }
                    } else if sparkline_list_depth.is_some() {
                        if Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "sparkline",
                        ) {
                            let sparkline =
                                Self::parse_sparkline(e, reader.decoder(), &document_namespaces)?;
                            let pending = pending_sparkline_group.as_mut().ok_or_else(|| {
                                Error::InvalidFormat(
                                    "calcext:sparklines must be inside calcext:sparkline-group"
                                        .to_string(),
                                )
                            })?;
                            if pending.group.sparklines.len() >= MAX_SPARKLINES_PER_GROUP {
                                return Err(Error::InvalidFormat(format!(
                                    "sparkline group exceeds the {MAX_SPARKLINES_PER_GROUP} sparkline safety limit"
                                )));
                            }
                            pending.group.sparklines.push(sparkline);
                            calcext_leaf_open_depth = Some(element_depth);
                        } else {
                            calcext_skip_depth = Some(element_depth);
                        }
                    } else if let Some(pending) = pending_sparkline_complex_color.as_mut() {
                        if Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            LOEXT_NAMESPACE_URI,
                            "transformation",
                        ) {
                            if pending.color.transformations.len() >= MAX_COLOR_TRANSFORMATIONS {
                                return Err(Error::InvalidFormat(format!(
                                    "complex color exceeds the {MAX_COLOR_TRANSFORMATIONS} transformation safety limit"
                                )));
                            }
                            pending
                                .color
                                .transformations
                                .push(Self::parse_color_transformation(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?);
                            calcext_leaf_open_depth = Some(element_depth);
                        } else {
                            calcext_skip_depth = Some(element_depth);
                        }
                    } else if pending_sparkline_group.is_some() {
                        if Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "sparklines",
                        ) {
                            sparkline_list_depth = Some(element_depth);
                        } else if let Some(slot) = COMPLEX_COLOR_SLOTS.iter().find(|slot| {
                            Self::element_name_is(
                                e.name().as_ref(),
                                &document_namespaces,
                                CALCEXT_NAMESPACE_URI,
                                slot,
                            )
                        }) {
                            let color = Self::parse_complex_color(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                            )?;
                            pending_sparkline_complex_color = Some(PendingSparklineComplexColor {
                                slot,
                                color,
                                depth: element_depth,
                            });
                        } else {
                            // Unmodeled children are skipped.
                            calcext_skip_depth = Some(element_depth);
                        }
                    } else if sparkline_groups_depth.is_some() {
                        if Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "sparkline-group",
                        ) {
                            let group = Self::parse_sparkline_group_attributes(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                            )?;
                            pending_sparkline_group = Some(PendingSparklineGroup {
                                group,
                                depth: element_depth,
                            });
                        } else {
                            calcext_skip_depth = Some(element_depth);
                        }
                    } else if current_sheet.is_some()
                        && current_row.is_none()
                        && current_sheet_depth.is_some_and(|depth| element_depth == depth + 1)
                        && e.local_name().as_ref() == b"sparkline-groups"
                    {
                        if !Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "sparkline-groups",
                        ) {
                            return Err(Error::InvalidFormat(
                                "spoofed calcext:sparkline-groups namespace".to_string(),
                            ));
                        }
                        sparkline_groups_depth = Some(element_depth);
                    } else if current_sheet.is_some()
                        && current_row.is_none()
                        && current_sheet_depth.is_some_and(|depth| element_depth == depth + 1)
                        && e.local_name().as_ref() == b"conditional-formats"
                    {
                        if !Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "conditional-formats",
                        ) {
                            return Err(Error::InvalidFormat(
                                "spoofed calcext:conditional-formats namespace".to_string(),
                            ));
                        }
                        conditional_formats_depth = Some(element_depth);
                    } else if let Some(builder) = detective_builder.as_mut() {
                        if detective_child_open {
                            return Err(Error::InvalidFormat(
                                "table:detective child elements must be empty".to_string(),
                            ));
                        }
                        Self::parse_detective_child(
                            builder,
                            e,
                            reader.decoder(),
                            &document_namespaces,
                        )?;
                        detective_child_open = true;
                    } else if let Some(builder) = annotation_builder.as_mut() {
                        builder.start(e, reader.decoder())?;
                        annotation_depth += 1;
                    } else if current_cell.is_some()
                        && Self::is_office_annotation(e, &document_namespaces)
                    {
                        annotation_builder = Some(Builder::new(
                            e,
                            reader.decoder(),
                            document_namespaces.clone(),
                        )?);
                    } else if text_element_depth > 0 {
                        rich_text_builder
                            .as_mut()
                            .expect("rich-text builder exists inside text:p")
                            .start(e, reader.decoder())?;
                        if Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TEXT_NAMESPACE,
                            "a",
                        ) {
                            if pending_hyperlink.is_some() {
                                return Err(Error::InvalidFormat(
                                    "text:a hyperlinks must not be nested".to_string(),
                                ));
                            }
                            pending_hyperlink = Some(PendingHyperlink {
                                link: Self::parse_hyperlink(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?,
                                text_start: text_content.len(),
                                depth: text_element_depth + 1,
                            });
                        }
                        text_element_depth += 1;
                    } else if Self::element_name_is(
                        e.name().as_ref(),
                        &document_namespaces,
                        TABLE_NAMESPACE_URI,
                        "table",
                    ) && spreadsheet_depth.is_some_and(|depth| element_depth == depth + 1)
                    {
                        let name =
                            Self::extract_table_name(e, reader.decoder(), &document_namespaces)?;
                        let (style, print_settings) = Self::parse_sheet_formatting(
                            e,
                            reader.decoder(),
                            &document_namespaces,
                        )?;
                        current_sheet =
                            Some(SheetBuilder::with_formatting(name, style, print_settings));
                        current_sheet_depth = Some(element_depth);
                    } else if current_sheet.is_some()
                        && current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table-source",
                        )
                    {
                        let source =
                            Self::parse_table_source(e, reader.decoder(), &document_namespaces)?;
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.set_table_source(source)?;
                        }
                    } else if current_sheet.is_some()
                        && current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "title",
                        )
                    {
                        if sheet_text_field.replace(SheetTextField::Title).is_some() {
                            return Err(Error::InvalidFormat(
                                "table title or description elements must not be nested"
                                    .to_string(),
                            ));
                        }
                        sheet_text.clear();
                    } else if current_sheet.is_some()
                        && current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "desc",
                        )
                    {
                        if sheet_text_field
                            .replace(SheetTextField::Description)
                            .is_some()
                        {
                            return Err(Error::InvalidFormat(
                                "table title or description elements must not be nested"
                                    .to_string(),
                            ));
                        }
                        sheet_text.clear();
                    } else if current_sheet.is_some()
                        && current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "scenario",
                        )
                    {
                        let scenario =
                            Self::parse_scenario(e, reader.decoder(), &document_namespaces)?;
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.set_scenario(scenario)?;
                        }
                    } else if current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table-column-group",
                        )
                    {
                        if let Some(sheet) = current_sheet.as_mut() {
                            let display = Self::parse_group_display(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                            )?;
                            sheet.begin_column_group(display)?;
                        }
                    } else if current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table-header-columns",
                        )
                    {
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.begin_column_header()?;
                        }
                    } else if current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table-row-group",
                        )
                    {
                        if let Some(sheet) = current_sheet.as_mut() {
                            let display = Self::parse_group_display(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                            )?;
                            sheet.begin_row_group(display)?;
                        }
                    } else if current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table-header-rows",
                        )
                    {
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.begin_row_header()?;
                        }
                    } else if current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table-column",
                        )
                    {
                        if let Some(sheet) = current_sheet.as_mut() {
                            let (column, repeated) =
                                Self::parse_column(e, reader.decoder(), &document_namespaces)?;
                            sheet.add_repeated_column(column, repeated)?;
                        }
                    } else if current_sheet.is_some()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table-row",
                        )
                    {
                        current_row = Some(RowBuilder::from_element(
                            e,
                            reader.decoder(),
                            &document_namespaces,
                        )?);
                    } else if current_row.is_some()
                        && (Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table-cell",
                        ) || Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "covered-table-cell",
                        ))
                    {
                        let covered = Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "covered-table-cell",
                        );
                        let mut cell_builder =
                            Self::parse_cell_attributes(e, reader.decoder(), &document_namespaces)?;
                        if covered {
                            cell_builder.mark_covered();
                        }
                        current_cell = Some(cell_builder);
                        text_content.clear();
                        rich_text_builder = None;
                        pending_hyperlink = None;
                        text_element_depth = 0;
                    } else if let Some(cell) = current_cell.as_mut()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "cell-range-source",
                        )
                    {
                        let source = Self::parse_cell_range_source(
                            e,
                            reader.decoder(),
                            &document_namespaces,
                        )?;
                        cell.set_range_source(source)?;
                    } else if let Some(cell) = current_cell.as_mut()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "detective",
                        )
                    {
                        cell.begin_detective()?;
                        detective_builder = Some(CellDetective::new());
                    } else if current_cell.is_some()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TEXT_NAMESPACE,
                            "p",
                        )
                    {
                        if !text_content.is_empty() {
                            text_content.push('\n');
                        }
                        let builder = rich_text_builder.get_or_insert_with(|| {
                            CellTextContentBuilder::new(document_namespaces.clone())
                        });
                        builder.start(e, reader.decoder())?;
                        text_element_depth = 1;
                    }
                },
                Ok(Event::Empty(ref e)) => {
                    let empty_scope =
                        Self::push_namespace_scope(e, reader.decoder(), &mut document_namespaces)?;
                    if sheet_dde_source_depth.is_some() {
                        return Err(Error::InvalidFormat(
                            "office:dde-source must not contain child elements".to_string(),
                        ));
                    }
                    if let Some(sheet) = current_sheet.as_mut()
                        && current_row.is_none()
                        && current_sheet_depth.is_some_and(|depth| element_depth == depth)
                        && e.local_name().as_ref() == b"dde-source"
                    {
                        if !Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            OFFICE_NAMESPACE,
                            "dde-source",
                        ) {
                            return Err(Error::InvalidFormat(
                                "spoofed office:dde-source namespace".to_string(),
                            ));
                        }
                        let source = parse_dde_source(e, reader.decoder(), &document_namespaces)?;
                        sheet.set_dde_source(source)?;
                        Self::pop_namespace_scope(&mut document_namespaces, Some(empty_scope));
                        buf.clear();
                        continue;
                    }
                    if calcext_skip_depth.is_some() {
                        Self::pop_namespace_scope(&mut document_namespaces, Some(empty_scope));
                        buf.clear();
                        continue;
                    }
                    if calcext_leaf_open_depth.is_some() {
                        return Err(Error::InvalidFormat(
                            "calcext rule and entry elements must not contain child elements"
                                .to_string(),
                        ));
                    }
                    if let Some(rule) = pending_calcext_rule.as_mut() {
                        let is_calcext = |local_name: &str| {
                            Self::element_name_is(
                                e.name().as_ref(),
                                &document_namespaces,
                                CALCEXT_NAMESPACE_URI,
                                local_name,
                            )
                        };
                        match rule {
                            PendingCalcextRule::ColorScale { entries, .. }
                                if is_calcext("color-scale-entry") =>
                            {
                                if entries.len() >= MAX_ENTRIES_PER_RULE {
                                    return Err(Error::InvalidFormat(format!(
                                        "calcext:color-scale exceeds the {MAX_ENTRIES_PER_RULE} entry safety limit"
                                    )));
                                }
                                entries.push(Self::parse_color_scale_entry(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?);
                            },
                            PendingCalcextRule::DataBar { data_bar, .. }
                                if is_calcext("formatting-entry")
                                    || is_calcext("data-bar-entry") =>
                            {
                                if data_bar.entries.len() >= DATA_BAR_ENTRY_COUNT {
                                    return Err(Error::InvalidFormat(format!(
                                        "calcext:data-bar supports exactly {DATA_BAR_ENTRY_COUNT} calcext:formatting-entry elements"
                                    )));
                                }
                                data_bar.entries.push(Self::parse_data_bar_entry(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?);
                            },
                            PendingCalcextRule::IconSet { icon_set, .. }
                                if is_calcext("formatting-entry") =>
                            {
                                if icon_set.entries.len() >= MAX_ENTRIES_PER_RULE {
                                    return Err(Error::InvalidFormat(format!(
                                        "calcext:icon-set exceeds the {MAX_ENTRIES_PER_RULE} entry safety limit"
                                    )));
                                }
                                icon_set.entries.push(Self::parse_icon_set_entry(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?);
                            },
                            PendingCalcextRule::IconSet { icon_set, .. }
                                if is_calcext("custom-iconset") =>
                            {
                                if icon_set.custom_icons.len() >= MAX_ENTRIES_PER_RULE {
                                    return Err(Error::InvalidFormat(format!(
                                        "calcext:icon-set exceeds the {MAX_ENTRIES_PER_RULE} custom icon safety limit"
                                    )));
                                }
                                icon_set.custom_icons.push(Self::parse_custom_icon(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?);
                            },
                            _ => {},
                        }
                        Self::pop_namespace_scope(&mut document_namespaces, Some(empty_scope));
                        buf.clear();
                        continue;
                    }
                    if let Some(pending) = pending_conditional_format.as_mut() {
                        let is_calcext = |local_name: &str| {
                            Self::element_name_is(
                                e.name().as_ref(),
                                &document_namespaces,
                                CALCEXT_NAMESPACE_URI,
                                local_name,
                            )
                        };
                        if is_calcext("condition") {
                            if pending.rules.len() >= MAX_RULES_PER_FORMAT {
                                return Err(Error::InvalidFormat(format!(
                                    "conditional format exceeds the {MAX_RULES_PER_FORMAT} rule safety limit"
                                )));
                            }
                            let condition = Self::parse_calcext_condition(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                            )?;
                            pending.rules.push(condition.into());
                        } else if is_calcext("date-is") {
                            if pending.rules.len() >= MAX_RULES_PER_FORMAT {
                                return Err(Error::InvalidFormat(format!(
                                    "conditional format exceeds the {MAX_RULES_PER_FORMAT} rule safety limit"
                                )));
                            }
                            let date_is =
                                Self::parse_date_is(e, reader.decoder(), &document_namespaces)?;
                            pending.rules.push(date_is.into());
                        } else if is_calcext("color-scale") {
                            let rule: ConditionalFormatRule =
                                ConditionalColorScale::new(Vec::new()).into();
                            validate_rule(&rule)?;
                        } else if is_calcext("data-bar") {
                            let data_bar = Self::parse_data_bar_attributes(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                            )?;
                            let rule: ConditionalFormatRule = data_bar.into();
                            validate_rule(&rule)?;
                        } else if is_calcext("icon-set") {
                            let icon_set = Self::parse_icon_set_attributes(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                            )?;
                            let rule: ConditionalFormatRule = icon_set.into();
                            validate_rule(&rule)?;
                        }
                        Self::pop_namespace_scope(&mut document_namespaces, Some(empty_scope));
                        buf.clear();
                        continue;
                    }
                    if conditional_formats_depth.is_some() {
                        if Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "conditional-format",
                        ) {
                            let format = ConditionalFormat {
                                target_range_addresses: Self::parse_conditional_format_ranges(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?,
                                rules: Vec::new(),
                            };
                            validate_conditional_format(&format)?;
                            if let Some(sheet) = current_sheet.as_mut() {
                                sheet.add_conditional_format(format)?;
                            }
                        }
                        Self::pop_namespace_scope(&mut document_namespaces, Some(empty_scope));
                        buf.clear();
                        continue;
                    }
                    if sparkline_list_depth.is_some() {
                        if Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "sparkline",
                        ) {
                            let sparkline =
                                Self::parse_sparkline(e, reader.decoder(), &document_namespaces)?;
                            let pending = pending_sparkline_group.as_mut().ok_or_else(|| {
                                Error::InvalidFormat(
                                    "calcext:sparklines must be inside calcext:sparkline-group"
                                        .to_string(),
                                )
                            })?;
                            if pending.group.sparklines.len() >= MAX_SPARKLINES_PER_GROUP {
                                return Err(Error::InvalidFormat(format!(
                                    "sparkline group exceeds the {MAX_SPARKLINES_PER_GROUP} sparkline safety limit"
                                )));
                            }
                            pending.group.sparklines.push(sparkline);
                        }
                        Self::pop_namespace_scope(&mut document_namespaces, Some(empty_scope));
                        buf.clear();
                        continue;
                    }
                    if let Some(pending) = pending_sparkline_complex_color.as_mut() {
                        if Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            LOEXT_NAMESPACE_URI,
                            "transformation",
                        ) {
                            if pending.color.transformations.len() >= MAX_COLOR_TRANSFORMATIONS {
                                return Err(Error::InvalidFormat(format!(
                                    "complex color exceeds the {MAX_COLOR_TRANSFORMATIONS} transformation safety limit"
                                )));
                            }
                            pending
                                .color
                                .transformations
                                .push(Self::parse_color_transformation(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?);
                        }
                        Self::pop_namespace_scope(&mut document_namespaces, Some(empty_scope));
                        buf.clear();
                        continue;
                    }
                    if let Some(pending) = pending_sparkline_group.as_mut() {
                        // An empty complex-color element carries no
                        // transformations; an empty `calcext:sparklines`
                        // container and unknown children carry nothing.
                        if let Some(slot) = COMPLEX_COLOR_SLOTS.iter().find(|slot| {
                            Self::element_name_is(
                                e.name().as_ref(),
                                &document_namespaces,
                                CALCEXT_NAMESPACE_URI,
                                slot,
                            )
                        }) {
                            let color = Self::parse_complex_color(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                            )?;
                            pending.group.complex_colors.assign_slot(slot, color)?;
                        }
                        Self::pop_namespace_scope(&mut document_namespaces, Some(empty_scope));
                        buf.clear();
                        continue;
                    }
                    if sparkline_groups_depth.is_some() {
                        if Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "sparkline-group",
                        ) {
                            let group = Self::parse_sparkline_group_attributes(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                            )?;
                            // An empty group has no sparklines and is invalid.
                            validate_sparkline_group(&group)?;
                        }
                        Self::pop_namespace_scope(&mut document_namespaces, Some(empty_scope));
                        buf.clear();
                        continue;
                    }
                    if current_sheet.is_some()
                        && current_row.is_none()
                        && current_sheet_depth.is_some_and(|depth| element_depth == depth)
                        && e.local_name().as_ref() == b"sparkline-groups"
                    {
                        if !Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "sparkline-groups",
                        ) {
                            return Err(Error::InvalidFormat(
                                "spoofed calcext:sparkline-groups namespace".to_string(),
                            ));
                        }
                        // An empty container declares no sparkline groups.
                        Self::pop_namespace_scope(&mut document_namespaces, Some(empty_scope));
                        buf.clear();
                        continue;
                    }
                    if current_sheet.is_some()
                        && current_row.is_none()
                        && current_sheet_depth.is_some_and(|depth| element_depth == depth)
                        && e.local_name().as_ref() == b"conditional-formats"
                    {
                        if !Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "conditional-formats",
                        ) {
                            return Err(Error::InvalidFormat(
                                "spoofed calcext:conditional-formats namespace".to_string(),
                            ));
                        }
                        // An empty container declares no conditional formats.
                        Self::pop_namespace_scope(&mut document_namespaces, Some(empty_scope));
                        buf.clear();
                        continue;
                    }
                    if spreadsheet_depth == Some(element_depth)
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table",
                        )
                    {
                        let name =
                            Self::extract_table_name(e, reader.decoder(), &document_namespaces)?;
                        let (style, print_settings) = Self::parse_sheet_formatting(
                            e,
                            reader.decoder(),
                            &document_namespaces,
                        )?;
                        sheets.push(
                            SheetBuilder::with_formatting(name, style, print_settings).build()?,
                        );
                    } else if let Some(builder) = detective_builder.as_mut() {
                        if detective_child_open {
                            return Err(Error::InvalidFormat(
                                "table:detective child elements must be empty".to_string(),
                            ));
                        }
                        Self::parse_detective_child(
                            builder,
                            e,
                            reader.decoder(),
                            &document_namespaces,
                        )?;
                    } else if let Some(builder) = annotation_builder.as_mut() {
                        builder.empty(e, reader.decoder())?;
                    } else if current_cell.is_some()
                        && Self::is_office_annotation(e, &document_namespaces)
                    {
                        let annotation =
                            Builder::new(e, reader.decoder(), document_namespaces.clone())?
                                .finish()?;
                        if let Some(cell) = current_cell.as_mut() {
                            cell.set_annotation(annotation);
                        }
                    } else if text_element_depth > 0 {
                        rich_text_builder
                            .as_mut()
                            .expect("rich-text builder exists inside text:p")
                            .empty(e, reader.decoder())?;
                        if Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TEXT_NAMESPACE,
                            "a",
                        ) {
                            if pending_hyperlink.is_some() {
                                return Err(Error::InvalidFormat(
                                    "text:a hyperlinks must not be nested".to_string(),
                                ));
                            }
                            let link =
                                Self::parse_hyperlink(e, reader.decoder(), &document_namespaces)?;
                            let mut link = link;
                            let position = text_content.len();
                            link.set_range(position..position);
                            if let Some(cell) = current_cell.as_mut() {
                                cell.push_hyperlink(link);
                            }
                        } else {
                            Self::push_text_empty_element(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                                &mut text_content,
                            )?;
                        }
                    } else if current_cell.is_some()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TEXT_NAMESPACE,
                            "p",
                        )
                    {
                        if !text_content.is_empty() {
                            text_content.push('\n');
                        }
                        let builder = rich_text_builder.get_or_insert_with(|| {
                            CellTextContentBuilder::new(document_namespaces.clone())
                        });
                        builder.empty(e, reader.decoder())?;
                    } else if let Some(cell) = current_cell.as_mut()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "cell-range-source",
                        )
                    {
                        let source = Self::parse_cell_range_source(
                            e,
                            reader.decoder(),
                            &document_namespaces,
                        )?;
                        cell.set_range_source(source)?;
                    } else if let Some(cell) = current_cell.as_mut()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "detective",
                        )
                    {
                        cell.set_detective(CellDetective::new())?;
                    } else if current_sheet.is_some()
                        && current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "title",
                        )
                    {
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.set_title(String::new())?;
                        }
                    } else if current_sheet.is_some()
                        && current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "desc",
                        )
                    {
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.set_description(String::new())?;
                        }
                    } else if current_sheet.is_some()
                        && current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table-source",
                        )
                    {
                        let source =
                            Self::parse_table_source(e, reader.decoder(), &document_namespaces)?;
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.set_table_source(source)?;
                        }
                    } else if current_sheet.is_some()
                        && current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "scenario",
                        )
                    {
                        let scenario =
                            Self::parse_scenario(e, reader.decoder(), &document_namespaces)?;
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.set_scenario(scenario)?;
                        }
                    } else if current_sheet.is_some()
                        && current_row.is_none()
                        && [
                            "table-column-group",
                            "table-header-columns",
                            "table-row-group",
                            "table-header-rows",
                        ]
                        .iter()
                        .any(|local_name| {
                            Self::element_name_is(
                                e.name().as_ref(),
                                &document_namespaces,
                                TABLE_NAMESPACE_URI,
                                local_name,
                            )
                        })
                    {
                        return Err(Error::InvalidFormat(
                            "empty table groups and header containers are not valid".to_string(),
                        ));
                    } else if current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table-column",
                        )
                    {
                        if let Some(sheet) = current_sheet.as_mut() {
                            let (column, repeated) =
                                Self::parse_column(e, reader.decoder(), &document_namespaces)?;
                            sheet.add_repeated_column(column, repeated)?;
                        }
                    } else if current_row.is_some()
                        && (Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table-cell",
                        ) || Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "covered-table-cell",
                        ))
                    {
                        let covered = Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "covered-table-cell",
                        );
                        let mut cell_builder =
                            Self::parse_cell_attributes(e, reader.decoder(), &document_namespaces)?;
                        if covered {
                            cell_builder.mark_covered();
                        }
                        if let Some(row) = current_row.as_mut() {
                            row.add_repeated_cells(&cell_builder, "", None)?;
                        }
                    } else if current_sheet.is_some()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table-row",
                        )
                    {
                        let row_builder =
                            RowBuilder::from_element(e, reader.decoder(), &document_namespaces)?;
                        let repeated = row_builder.repeated();
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.add_repeated_row(row_builder.build(), repeated)?;
                        }
                    }
                    Self::pop_namespace_scope(&mut document_namespaces, Some(empty_scope));
                },
                Ok(Event::Text(ref t)) if detective_builder.is_some() => {
                    let decoded = t.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid detective text: {error}"))
                    })?;
                    if !decoded.trim().is_empty() {
                        return Err(Error::InvalidFormat(
                            "table:detective cannot contain text".to_string(),
                        ));
                    }
                },
                Ok(Event::CData(ref t)) if detective_builder.is_some() => {
                    let decoded = t.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid detective CDATA: {error}"))
                    })?;
                    if !decoded.trim().is_empty() {
                        return Err(Error::InvalidFormat(
                            "table:detective cannot contain CDATA".to_string(),
                        ));
                    }
                },
                Ok(Event::GeneralRef(_)) if detective_builder.is_some() => {
                    return Err(Error::InvalidFormat(
                        "table:detective cannot contain entity references".to_string(),
                    ));
                },
                Ok(Event::Text(_)) | Ok(Event::CData(_)) | Ok(Event::GeneralRef(_))
                    if sheet_dde_source_depth.is_some() =>
                {
                    return Err(Error::InvalidFormat(
                        "office:dde-source must be empty".to_string(),
                    ));
                },
                Ok(Event::Text(ref t)) if annotation_builder.is_some() => {
                    if let Some(builder) = annotation_builder.as_mut() {
                        builder.text(t)?;
                    }
                },
                Ok(Event::CData(ref t)) if annotation_builder.is_some() => {
                    if let Some(builder) = annotation_builder.as_mut() {
                        builder.cdata(t)?;
                    }
                },
                Ok(Event::GeneralRef(ref reference)) if annotation_builder.is_some() => {
                    if let Some(builder) = annotation_builder.as_mut() {
                        builder.reference(reference)?;
                    }
                },
                Ok(Event::Text(ref t)) if sheet_text_field.is_some() => {
                    let decoded = t.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid table text: {error}"))
                    })?;
                    let decoded = quick_xml::escape::unescape(&decoded).map_err(|error| {
                        Error::InvalidFormat(format!("invalid table character reference: {error}"))
                    })?;
                    sheet_text.push_str(&decoded);
                },
                Ok(Event::CData(ref t)) if sheet_text_field.is_some() => {
                    let decoded = t.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid table CDATA: {error}"))
                    })?;
                    sheet_text.push_str(&decoded);
                },
                Ok(Event::GeneralRef(ref reference)) if sheet_text_field.is_some() => {
                    sheet_text.push_str(&decode_reference(reference)?);
                },
                Ok(Event::Text(ref t)) if text_element_depth > 0 && current_cell.is_some() => {
                    rich_text_builder
                        .as_mut()
                        .expect("rich-text builder exists inside text:p")
                        .text(t)?;
                    let decoded = t.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid cell text: {error}"))
                    })?;
                    let decoded = quick_xml::escape::unescape(&decoded).map_err(|error| {
                        Error::InvalidFormat(format!("invalid cell character reference: {error}"))
                    })?;
                    text_content.push_str(&decoded);
                },
                Ok(Event::CData(ref t)) if text_element_depth > 0 && current_cell.is_some() => {
                    rich_text_builder
                        .as_mut()
                        .expect("rich-text builder exists inside text:p")
                        .cdata(t)?;
                    let decoded = t.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid cell CDATA: {error}"))
                    })?;
                    text_content.push_str(&decoded);
                },
                Ok(Event::GeneralRef(ref reference))
                    if text_element_depth > 0 && current_cell.is_some() =>
                {
                    rich_text_builder
                        .as_mut()
                        .expect("rich-text builder exists inside text:p")
                        .reference(reference)?;
                    text_content.push_str(&decode_reference(reference)?);
                },
                Ok(Event::End(ref e)) => {
                    let closes_sheet_dde_source = sheet_dde_source_depth == Some(element_depth)
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            OFFICE_NAMESPACE,
                            "dde-source",
                        );
                    let closes_current_sheet = current_sheet_depth == Some(element_depth)
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table",
                        );
                    let closes_spreadsheet = spreadsheet_depth == Some(element_depth)
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            OFFICE_NAMESPACE,
                            "spreadsheet",
                        );
                    let closes_calcext_skip = calcext_skip_depth == Some(element_depth);
                    let closes_calcext_leaf = calcext_leaf_open_depth == Some(element_depth);
                    let closes_calcext_rule = pending_calcext_rule
                        .as_ref()
                        .is_some_and(|rule| rule.depth() == element_depth)
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            pending_calcext_rule
                                .as_ref()
                                .expect("pending calcext rule was checked")
                                .element_name(),
                        );
                    let closes_conditional_format = pending_conditional_format
                        .as_ref()
                        .is_some_and(|pending| pending.depth == element_depth)
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "conditional-format",
                        );
                    let closes_conditional_formats = conditional_formats_depth
                        == Some(element_depth)
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "conditional-formats",
                        );
                    let closes_sparkline_list = sparkline_list_depth == Some(element_depth)
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "sparklines",
                        );
                    let closes_sparkline_complex_color = pending_sparkline_complex_color
                        .as_ref()
                        .is_some_and(|pending| {
                            pending.depth == element_depth
                                && Self::element_name_is(
                                    e.name().as_ref(),
                                    &document_namespaces,
                                    CALCEXT_NAMESPACE_URI,
                                    pending.slot,
                                )
                        });
                    let closes_sparkline_group = pending_sparkline_group
                        .as_ref()
                        .is_some_and(|pending| pending.depth == element_depth)
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "sparkline-group",
                        );
                    let closes_sparkline_groups = sparkline_groups_depth == Some(element_depth)
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "sparkline-groups",
                        );
                    element_depth = element_depth.saturating_sub(1);
                    if closes_sheet_dde_source {
                        sheet_dde_source_depth = None;
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }
                    if closes_calcext_skip {
                        calcext_skip_depth = None;
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }
                    if closes_calcext_leaf {
                        calcext_leaf_open_depth = None;
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }
                    if closes_calcext_rule {
                        let rule = pending_calcext_rule
                            .take()
                            .expect("pending calcext rule was checked")
                            .finish();
                        validate_rule(&rule)?;
                        let pending = pending_conditional_format.as_mut().ok_or_else(|| {
                            Error::InvalidFormat(
                                "calcext rule closed outside calcext:conditional-format"
                                    .to_string(),
                            )
                        })?;
                        if pending.rules.len() >= MAX_RULES_PER_FORMAT {
                            return Err(Error::InvalidFormat(format!(
                                "conditional format exceeds the {MAX_RULES_PER_FORMAT} rule safety limit"
                            )));
                        }
                        pending.rules.push(rule);
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }
                    if closes_conditional_format {
                        let pending = pending_conditional_format
                            .take()
                            .expect("pending conditional format was checked");
                        let format = ConditionalFormat {
                            target_range_addresses: pending.target_range_addresses,
                            rules: pending.rules,
                        };
                        validate_conditional_format(&format)?;
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.add_conditional_format(format)?;
                        }
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }
                    if closes_conditional_formats {
                        conditional_formats_depth = None;
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }
                    if closes_sparkline_list {
                        sparkline_list_depth = None;
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }
                    if closes_sparkline_complex_color {
                        let pending = pending_sparkline_complex_color
                            .take()
                            .expect("pending complex color was checked");
                        let group = pending_sparkline_group.as_mut().ok_or_else(|| {
                            Error::InvalidFormat(
                                "complex color closed outside calcext:sparkline-group".to_string(),
                            )
                        })?;
                        group
                            .group
                            .complex_colors
                            .assign_slot(pending.slot, pending.color)?;
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }
                    if closes_sparkline_group {
                        let pending = pending_sparkline_group
                            .take()
                            .expect("pending sparkline group was checked");
                        validate_sparkline_group(&pending.group)?;
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.add_sparkline_group(pending.group)?;
                        }
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }
                    if closes_sparkline_groups {
                        sparkline_groups_depth = None;
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }
                    if detective_builder.is_some() {
                        if detective_child_open {
                            detective_child_open = false;
                        } else if Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "detective",
                        ) {
                            let detective = detective_builder
                                .take()
                                .expect("detective builder was checked");
                            let cell = current_cell.as_mut().ok_or_else(|| {
                                Error::InvalidFormat(
                                    "table:detective must be contained in a table cell".to_string(),
                                )
                            })?;
                            cell.set_detective(detective)?;
                        }
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }

                    if annotation_builder.is_some() {
                        if annotation_depth == 0 {
                            let annotation = annotation_builder
                                .take()
                                .expect("annotation builder was checked")
                                .finish()?;
                            if let Some(cell) = current_cell.as_mut() {
                                cell.set_annotation(annotation);
                            }
                        } else {
                            annotation_builder
                                .as_mut()
                                .expect("annotation builder was checked")
                                .end_element()?;
                            annotation_depth -= 1;
                        }
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }

                    if text_element_depth > 0 {
                        rich_text_builder
                            .as_mut()
                            .expect("rich-text builder exists inside text:p")
                            .end()?;
                        if pending_hyperlink
                            .as_ref()
                            .is_some_and(|pending| pending.depth == text_element_depth)
                        {
                            let pending = pending_hyperlink
                                .take()
                                .expect("pending hyperlink was checked");
                            let mut link = pending.link;
                            let range = pending.text_start..text_content.len();
                            link.text = text_content
                                .get(range.clone())
                                .unwrap_or_default()
                                .to_string();
                            link.set_range(range);
                            if let Some(cell) = current_cell.as_mut() {
                                cell.push_hyperlink(link);
                            }
                        }
                        text_element_depth -= 1;
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }

                    if let Some(field) = sheet_text_field
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            field.local_name(),
                        )
                    {
                        let value = std::mem::take(&mut sheet_text);
                        if let Some(sheet) = current_sheet.as_mut() {
                            match field {
                                SheetTextField::Title => sheet.set_title(value)?,
                                SheetTextField::Description => sheet.set_description(value)?,
                            }
                        }
                        sheet_text_field = None;
                    } else if Self::element_name_is(
                        e.name().as_ref(),
                        &document_namespaces,
                        TABLE_NAMESPACE_URI,
                        "table-cell",
                    ) || Self::element_name_is(
                        e.name().as_ref(),
                        &document_namespaces,
                        TABLE_NAMESPACE_URI,
                        "covered-table-cell",
                    ) {
                        if let Some(cell_builder) = current_cell.take()
                            && let Some(ref mut row_builder) = current_row
                        {
                            let rich_text = rich_text_builder
                                .take()
                                .map(CellTextContentBuilder::finish)
                                .transpose()?;
                            row_builder.add_repeated_cells(
                                &cell_builder,
                                &text_content,
                                rich_text.as_ref(),
                            )?;
                        }
                    } else if Self::element_name_is(
                        e.name().as_ref(),
                        &document_namespaces,
                        TABLE_NAMESPACE_URI,
                        "table-row",
                    ) {
                        if let Some(row_builder) = current_row.take() {
                            let repeated = row_builder.repeated();
                            let row = row_builder.build();
                            if let Some(ref mut sheet_builder) = current_sheet {
                                sheet_builder.add_repeated_row(row, repeated)?;
                            }
                        }
                    } else if Self::element_name_is(
                        e.name().as_ref(),
                        &document_namespaces,
                        TABLE_NAMESPACE_URI,
                        "table-column-group",
                    ) {
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.end_column_group()?;
                        }
                    } else if Self::element_name_is(
                        e.name().as_ref(),
                        &document_namespaces,
                        TABLE_NAMESPACE_URI,
                        "table-header-columns",
                    ) {
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.end_column_header()?;
                        }
                    } else if Self::element_name_is(
                        e.name().as_ref(),
                        &document_namespaces,
                        TABLE_NAMESPACE_URI,
                        "table-row-group",
                    ) {
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.end_row_group()?;
                        }
                    } else if Self::element_name_is(
                        e.name().as_ref(),
                        &document_namespaces,
                        TABLE_NAMESPACE_URI,
                        "table-header-rows",
                    ) {
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.end_row_header()?;
                        }
                    } else if closes_current_sheet && let Some(sheet_builder) = current_sheet.take()
                    {
                        sheets.push(sheet_builder.build()?);
                        current_sheet_depth = None;
                    }
                    if closes_spreadsheet {
                        spreadsheet_depth = None;
                    }
                    Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                },
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(Error::InvalidFormat(format!("XML parsing error: {}", e)));
                },
                _ => {},
            }
            buf.clear();
        }

        if sheet_dde_source_depth.is_some() {
            return Err(Error::InvalidFormat(
                "unterminated office:dde-source".to_string(),
            ));
        }
        if conditional_formats_depth.is_some() {
            return Err(Error::InvalidFormat(
                "unterminated calcext:conditional-formats".to_string(),
            ));
        }
        if sparkline_groups_depth.is_some() {
            return Err(Error::InvalidFormat(
                "unterminated calcext:sparkline-groups".to_string(),
            ));
        }

        super::super::super::super::package::attach_content_assets(xml_content, &mut sheets)?;
        Ok(sheets)
    }
}
