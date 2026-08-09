use super::{
    BufRead, BytesStart, ChartXmlReader, ColorMapOverride, ColorMapping, ColorSchemeIndex, Error,
    Event, ExtensionList, ExternalData, HeaderFooter, IGNORED_NAMESPACE_ELEMENT, Layout, Marker,
    PageMargins, PageOrientation, PageSetup, PivotFormat, PivotSource, PrintSettings, Protection,
    Result, ShapeProperties, TextProperties, TitleText, consume_empty_chart_element, get_attr,
    invalid_attribute, missing_attribute, optional_bool_attr, optional_i32_attr, optional_u32_attr,
    parse_bool_attr, parse_data_label, parse_series_marker, parse_text_element,
    required_named_f64_attr, required_u32_attr,
};

pub(crate) fn parse_pivot_source<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<PivotSource> {
    let mut name = None;
    let mut format_id = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"name" => {
                if name.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot source contains duplicate names".into(),
                    ));
                }
                name = Some(parse_text_element(reader, b"name")?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"name" => {
                if name.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot source contains duplicate names".into(),
                    ));
                }
                name = Some(String::new());
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"fmtId" => {
                if format_id.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot source contains duplicate format IDs".into(),
                    ));
                }
                format_id = Some(required_u32_attr(element, "chart pivot-source format ID")?);
                consume_empty_chart_element(reader, b"fmtId", "chart pivot-source format ID")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"fmtId" => {
                if format_id.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot source contains duplicate format IDs".into(),
                    ));
                }
                format_id = Some(required_u32_attr(element, "chart pivot-source format ID")?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"pivotSource" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart pivot source".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(PivotSource::new(
        name.ok_or_else(|| missing_attribute("chart pivot-source name"))?,
        format_id.ok_or_else(|| missing_attribute("chart pivot-source format ID"))?,
    ))
}

pub(crate) fn parse_chart_protection<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Protection> {
    let mut protection = Protection::default();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) => {
                let field = match element.local_name().as_ref() {
                    b"chartObject" => &mut protection.chart_object,
                    b"data" => &mut protection.data,
                    b"formatting" => &mut protection.formatting,
                    b"selection" => &mut protection.selection,
                    b"userInterface" => &mut protection.user_interface,
                    _ => {
                        buf.clear();
                        continue;
                    },
                };
                if field.is_some() {
                    return Err(Error::Invalid(format!(
                        "chart protection contains duplicate {} settings",
                        String::from_utf8_lossy(element.local_name().as_ref())
                    )));
                }
                *field = Some(parse_bool_attr(element)?);
                let end_name = element.local_name().as_ref().to_vec();
                consume_empty_chart_element(reader, &end_name, "chart protection switch")?;
            },
            Ok(Event::Empty(ref element)) => {
                let field = match element.local_name().as_ref() {
                    b"chartObject" => &mut protection.chart_object,
                    b"data" => &mut protection.data,
                    b"formatting" => &mut protection.formatting,
                    b"selection" => &mut protection.selection,
                    b"userInterface" => &mut protection.user_interface,
                    _ => {
                        buf.clear();
                        continue;
                    },
                };
                if field.is_some() {
                    return Err(Error::Invalid(format!(
                        "chart protection contains duplicate {} settings",
                        String::from_utf8_lossy(element.local_name().as_ref())
                    )));
                }
                *field = Some(parse_bool_attr(element)?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"protection" => {
                return Ok(protection);
            },
            Ok(Event::Eof) => {
                return Err(Error::Invalid("chart protection is not closed".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
}

pub(crate) fn parse_color_map_override<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<ColorMapOverride> {
    let mut mapping = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element))
                if element.local_name().as_ref() == b"masterClrMapping" =>
            {
                if mapping.is_some() {
                    return Err(Error::Invalid(
                        "chart color-map override contains multiple choices".into(),
                    ));
                }
                consume_empty_chart_element(
                    reader,
                    b"masterClrMapping",
                    "chart master color mapping",
                )?;
                mapping = Some(ColorMapOverride::Master);
            },
            Ok(Event::Empty(ref element))
                if element.local_name().as_ref() == b"masterClrMapping" =>
            {
                set_color_map_override_choice(&mut mapping, ColorMapOverride::Master)?;
            },
            Ok(Event::Start(ref element))
                if element.local_name().as_ref() == b"overrideClrMapping" =>
            {
                if mapping.is_some() {
                    return Err(Error::Invalid(
                        "chart color-map override contains multiple choices".into(),
                    ));
                }
                let value = parse_color_mapping(element)?;
                consume_empty_chart_element(
                    reader,
                    b"overrideClrMapping",
                    "chart override color mapping",
                )?;
                mapping = Some(ColorMapOverride::Override(value));
            },
            Ok(Event::Empty(ref element))
                if element.local_name().as_ref() == b"overrideClrMapping" =>
            {
                let value = ColorMapOverride::Override(parse_color_mapping(element)?);
                set_color_map_override_choice(&mut mapping, value)?;
            },
            Ok(Event::Start(ref element) | Event::Empty(ref element))
                if element.local_name().as_ref() != IGNORED_NAMESPACE_ELEMENT.as_bytes() =>
            {
                return Err(Error::Invalid(
                    "chart color-map override contains an unexpected choice".into(),
                ));
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"clrMapOvr" => {
                return mapping.ok_or_else(|| {
                    Error::Invalid("chart color-map override requires a mapping choice".into())
                });
            },
            Ok(Event::Eof) => {
                return Err(Error::Invalid(
                    "chart color-map override is not closed".into(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
}

pub(crate) fn set_color_map_override_choice(
    target: &mut Option<ColorMapOverride>,
    value: ColorMapOverride,
) -> Result<()> {
    if target.is_some() {
        return Err(Error::Invalid(
            "chart color-map override contains multiple choices".into(),
        ));
    }
    *target = Some(value);
    Ok(())
}

pub(crate) fn parse_color_mapping(element: &BytesStart<'_>) -> Result<ColorMapping> {
    Ok(ColorMapping {
        background1: required_color_scheme_index(element, b"bg1")?,
        text1: required_color_scheme_index(element, b"tx1")?,
        background2: required_color_scheme_index(element, b"bg2")?,
        text2: required_color_scheme_index(element, b"tx2")?,
        accent1: required_color_scheme_index(element, b"accent1")?,
        accent2: required_color_scheme_index(element, b"accent2")?,
        accent3: required_color_scheme_index(element, b"accent3")?,
        accent4: required_color_scheme_index(element, b"accent4")?,
        accent5: required_color_scheme_index(element, b"accent5")?,
        accent6: required_color_scheme_index(element, b"accent6")?,
        hyperlink: required_color_scheme_index(element, b"hlink")?,
        followed_hyperlink: required_color_scheme_index(element, b"folHlink")?,
    })
}

pub(crate) fn required_chart_relationship_id<R: BufRead>(
    reader: &ChartXmlReader<R>,
    element: &BytesStart<'_>,
) -> Result<String> {
    reader
        .relationship_attribute_value(element, b"id")?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| Error::Invalid("chart relationship ID is required".into()))
}

pub(crate) fn parse_external_data<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
    relationship_id: String,
) -> Result<ExternalData> {
    let mut external_data = ExternalData::new(relationship_id);
    let mut saw_auto_update = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"autoUpdate" => {
                if saw_auto_update {
                    return Err(Error::Invalid(
                        "chart external data contains duplicate auto-update settings".into(),
                    ));
                }
                saw_auto_update = true;
                external_data.auto_update = Some(parse_bool_attr(element)?);
                consume_empty_chart_element(
                    reader,
                    b"autoUpdate",
                    "chart external-data auto-update setting",
                )?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"autoUpdate" => {
                if saw_auto_update {
                    return Err(Error::Invalid(
                        "chart external data contains duplicate auto-update settings".into(),
                    ));
                }
                saw_auto_update = true;
                external_data.auto_update = Some(parse_bool_attr(element)?);
            },
            Ok(Event::Start(ref element) | Event::Empty(ref element))
                if element.local_name().as_ref() != IGNORED_NAMESPACE_ELEMENT.as_bytes() =>
            {
                return Err(Error::Invalid(
                    "chart external data contains an unexpected child".into(),
                ));
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"externalData" => {
                return Ok(external_data);
            },
            Ok(Event::Eof) => {
                return Err(Error::Invalid("chart external data is not closed".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
}

pub(crate) fn required_color_scheme_index(
    element: &BytesStart<'_>,
    attribute: &[u8],
) -> Result<ColorSchemeIndex> {
    let value = get_attr(element, attribute).ok_or_else(|| {
        Error::Invalid(format!(
            "chart color mapping is missing its {} assignment",
            String::from_utf8_lossy(attribute)
        ))
    })?;
    match value.as_slice() {
        b"dk1" => Ok(ColorSchemeIndex::Dark1),
        b"lt1" => Ok(ColorSchemeIndex::Light1),
        b"dk2" => Ok(ColorSchemeIndex::Dark2),
        b"lt2" => Ok(ColorSchemeIndex::Light2),
        b"accent1" => Ok(ColorSchemeIndex::Accent1),
        b"accent2" => Ok(ColorSchemeIndex::Accent2),
        b"accent3" => Ok(ColorSchemeIndex::Accent3),
        b"accent4" => Ok(ColorSchemeIndex::Accent4),
        b"accent5" => Ok(ColorSchemeIndex::Accent5),
        b"accent6" => Ok(ColorSchemeIndex::Accent6),
        b"hlink" => Ok(ColorSchemeIndex::Hyperlink),
        b"folHlink" => Ok(ColorSchemeIndex::FollowedHyperlink),
        _ => Err(invalid_attribute("chart color-scheme index", &value)),
    }
}

pub(crate) fn parse_pivot_formats<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Vec<PivotFormat>> {
    let mut formats = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"pivotFmt" => {
                let format = parse_pivot_format(reader)?;
                if formats
                    .iter()
                    .any(|existing: &PivotFormat| existing.index == format.index)
                {
                    return Err(Error::Invalid(format!(
                        "chart contains duplicate pivot-format index {}",
                        format.index
                    )));
                }
                formats.push(format);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"pivotFmt" => {
                return Err(Error::Invalid(
                    "chart pivot format is missing its index".into(),
                ));
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"pivotFmts" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart pivot formats".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(formats)
}

pub(crate) fn parse_pivot_format<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<PivotFormat> {
    let mut index = None;
    let mut shape_properties = None;
    let mut text_properties = None;
    let mut marker = None;
    let mut data_label = None;
    let mut extension_list = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot format contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart pivot-format shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot format contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot format contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(TextProperties::from_xml(
                    reader.capture_fragment(element, "chart pivot-format text properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot format contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(TextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"marker" => {
                if marker.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot format contains duplicate markers".into(),
                    ));
                }
                marker = Some(parse_series_marker(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"marker" => {
                if marker.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot format contains duplicate markers".into(),
                    ));
                }
                marker = Some(Marker::new());
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"dLbl" => {
                if data_label.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot format contains duplicate data labels".into(),
                    ));
                }
                data_label = Some(parse_data_label(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"dLbl" => {
                return Err(Error::Invalid(
                    "chart pivot-format data label is missing its index".into(),
                ));
            },
            Ok(Event::Start(ref element) | Event::Empty(ref element))
                if element.local_name().as_ref() == b"idx" =>
            {
                if index.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot format contains duplicate indexes".into(),
                    ));
                }
                index = Some(required_u32_attr(element, "chart pivot-format index")?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot format contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(element, "chart pivot-format extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot format contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"pivotFmt" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart pivot format".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let mut format =
        PivotFormat::new(index.ok_or_else(|| missing_attribute("chart pivot-format index"))?);
    format.shape_properties = shape_properties;
    format.text_properties = text_properties;
    format.marker = marker;
    format.data_label = data_label;
    format.extension_list = extension_list;
    Ok(format)
}

pub(crate) fn parse_print_settings<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<PrintSettings> {
    let mut settings = PrintSettings::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"headerFooter" => {
                if settings.header_footer.is_some() {
                    return Err(Error::Invalid(
                        "chart print settings contain duplicate headers and footers".into(),
                    ));
                }
                settings.header_footer = Some(parse_chart_header_footer(reader, element)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"headerFooter" => {
                if settings.header_footer.is_some() {
                    return Err(Error::Invalid(
                        "chart print settings contain duplicate headers and footers".into(),
                    ));
                }
                settings.header_footer = Some(parse_chart_header_footer_attributes(element)?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"pageMargins" => {
                if settings.page_margins.is_some() {
                    return Err(Error::Invalid(
                        "chart print settings contain duplicate page margins".into(),
                    ));
                }
                settings.page_margins = Some(parse_chart_page_margins(element)?);
                consume_empty_chart_element(reader, b"pageMargins", "chart page margins")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"pageMargins" => {
                if settings.page_margins.is_some() {
                    return Err(Error::Invalid(
                        "chart print settings contain duplicate page margins".into(),
                    ));
                }
                settings.page_margins = Some(parse_chart_page_margins(element)?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"pageSetup" => {
                if settings.page_setup.is_some() {
                    return Err(Error::Invalid(
                        "chart print settings contain duplicate page setup".into(),
                    ));
                }
                settings.page_setup = Some(parse_chart_page_setup(element)?);
                consume_empty_chart_element(reader, b"pageSetup", "chart page setup")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"pageSetup" => {
                if settings.page_setup.is_some() {
                    return Err(Error::Invalid(
                        "chart print settings contain duplicate page setup".into(),
                    ));
                }
                settings.page_setup = Some(parse_chart_page_setup(element)?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"printSettings" => {
                break;
            },
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart print settings".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(settings)
}

pub(crate) fn parse_chart_header_footer_attributes(
    element: &BytesStart<'_>,
) -> Result<HeaderFooter> {
    let mut header_footer = HeaderFooter::new();
    header_footer.align_with_margins = optional_bool_attr(
        element,
        b"alignWithMargins",
        true,
        "header/footer alignment",
    )?;
    header_footer.different_odd_even = optional_bool_attr(
        element,
        b"differentOddEven",
        false,
        "odd/even header/footer selection",
    )?;
    header_footer.different_first = optional_bool_attr(
        element,
        b"differentFirst",
        false,
        "first-page header/footer selection",
    )?;
    Ok(header_footer)
}

pub(crate) fn parse_chart_header_footer<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
    element: &BytesStart<'_>,
) -> Result<HeaderFooter> {
    let mut header_footer = parse_chart_header_footer_attributes(element)?;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref child)) => {
                let value = match child.local_name().as_ref() {
                    b"oddHeader" => &mut header_footer.odd_header,
                    b"oddFooter" => &mut header_footer.odd_footer,
                    b"evenHeader" => &mut header_footer.even_header,
                    b"evenFooter" => &mut header_footer.even_footer,
                    b"firstHeader" => &mut header_footer.first_header,
                    b"firstFooter" => &mut header_footer.first_footer,
                    name if name == IGNORED_NAMESPACE_ELEMENT.as_bytes() => {
                        buf.clear();
                        continue;
                    },
                    _ => {
                        return Err(Error::Invalid(
                            "chart header/footer contains an unexpected child".into(),
                        ));
                    },
                };
                if value.is_some() {
                    return Err(Error::Invalid(
                        "chart header/footer contains a duplicate string".into(),
                    ));
                }
                *value = Some(parse_text_element(reader, child.local_name().as_ref())?);
            },
            Ok(Event::Empty(ref child)) => {
                let value = match child.local_name().as_ref() {
                    b"oddHeader" => &mut header_footer.odd_header,
                    b"oddFooter" => &mut header_footer.odd_footer,
                    b"evenHeader" => &mut header_footer.even_header,
                    b"evenFooter" => &mut header_footer.even_footer,
                    b"firstHeader" => &mut header_footer.first_header,
                    b"firstFooter" => &mut header_footer.first_footer,
                    name if name == IGNORED_NAMESPACE_ELEMENT.as_bytes() => {
                        buf.clear();
                        continue;
                    },
                    _ => {
                        return Err(Error::Invalid(
                            "chart header/footer contains an unexpected child".into(),
                        ));
                    },
                };
                if value.replace(String::new()).is_some() {
                    return Err(Error::Invalid(
                        "chart header/footer contains a duplicate string".into(),
                    ));
                }
            },
            Ok(Event::End(ref child)) if child.local_name().as_ref() == b"headerFooter" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart header/footer".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(header_footer)
}

pub(crate) fn parse_chart_page_margins(element: &BytesStart<'_>) -> Result<PageMargins> {
    Ok(PageMargins::new(
        required_named_f64_attr(element, b"l", "chart left page margin")?,
        required_named_f64_attr(element, b"r", "chart right page margin")?,
        required_named_f64_attr(element, b"t", "chart top page margin")?,
        required_named_f64_attr(element, b"b", "chart bottom page margin")?,
        required_named_f64_attr(element, b"header", "chart header page margin")?,
        required_named_f64_attr(element, b"footer", "chart footer page margin")?,
    ))
}

pub(crate) fn parse_chart_page_setup(element: &BytesStart<'_>) -> Result<PageSetup> {
    let mut setup = PageSetup::new();
    setup.paper_size = optional_u32_attr(element, b"paperSize", 1, "chart printer paper size")?;
    setup.first_page_number =
        optional_u32_attr(element, b"firstPageNumber", 1, "chart first page number")?;
    setup.orientation = match get_attr(element, b"orientation").as_deref() {
        None | Some(b"default") => PageOrientation::Default,
        Some(b"portrait") => PageOrientation::Portrait,
        Some(b"landscape") => PageOrientation::Landscape,
        Some(value) => return Err(invalid_attribute("chart page orientation", value)),
    };
    setup.black_and_white = optional_bool_attr(
        element,
        b"blackAndWhite",
        false,
        "chart black-and-white printing",
    )?;
    setup.draft = optional_bool_attr(element, b"draft", false, "chart draft printing")?;
    setup.use_first_page_number = optional_bool_attr(
        element,
        b"useFirstPageNumber",
        false,
        "chart first-page-number usage",
    )?;
    setup.horizontal_dpi = optional_i32_attr(
        element,
        b"horizontalDpi",
        600,
        "chart horizontal printer resolution",
    )?;
    setup.vertical_dpi = optional_i32_attr(
        element,
        b"verticalDpi",
        600,
        "chart vertical printer resolution",
    )?;
    setup.copies = optional_u32_attr(element, b"copies", 1, "chart print copies")?;
    Ok(setup)
}

pub(crate) struct ParsedTitle {
    pub(crate) text: TitleText,
    pub(crate) layout: Option<Layout>,
    pub(crate) overlay: bool,
    pub(crate) shape_properties: Option<ShapeProperties>,
    pub(crate) text_properties: Option<TextProperties>,
    pub(crate) extension_list: Option<ExtensionList>,
}
