use super::*;

#[derive(Default)]
struct ParsedAxisScaling {
    min: Option<f64>,
    max: Option<f64>,
    log_base: Option<f64>,
}

impl ParsedAxisScaling {
    fn parse_element(&mut self, element: &BytesStart<'_>) -> Result<bool> {
        let (slot, description) = match element.local_name().as_ref() {
            b"min" => (&mut self.min, "chart axis minimum"),
            b"max" => (&mut self.max, "chart axis maximum"),
            b"logBase" => (&mut self.log_base, "chart logarithmic base"),
            _ => return Ok(false),
        };
        if slot.is_some() {
            return Err(Error::Invalid(format!(
                "{description} is specified more than once"
            )));
        }
        let value = required_f64_attr(element, description)?;
        if element.local_name().as_ref() == b"logBase" && !(2.0..=1000.0).contains(&value) {
            return Err(Error::Invalid(
                "chart logarithmic base must be between 2 and 1000".into(),
            ));
        }
        *slot = Some(value);
        Ok(true)
    }

    fn validate(&self) -> Result<()> {
        if self.min.zip(self.max).is_some_and(|(min, max)| min > max) {
            return Err(Error::Invalid("chart axis minimum exceeds maximum".into()));
        }
        Ok(())
    }
}

pub(crate) struct ParsedAxisCommon {
    axis_id: Option<u32>,
    cross_axis_id: Option<u32>,
    position: Option<AxisPosition>,
    title: Option<TitleText>,
    title_layout: Option<Layout>,
    title_overlay: bool,
    title_shape_properties: Option<ShapeProperties>,
    title_text_properties: Option<TextProperties>,
    title_extension_list: Option<ExtensionList>,
    number_format: Option<NumberFormat>,
    orientation: AxisOrientation,
    major_tick_mark: TickMark,
    minor_tick_mark: TickMark,
    tick_label_position: TickLabelPosition,
    deleted: bool,
    cross_mode: AxisCrossMode,
    crosses_at: Option<f64>,
    show_major_gridlines: bool,
    show_minor_gridlines: bool,
    major_gridlines: Option<Lines>,
    minor_gridlines: Option<Lines>,
    pub(super) shape_properties: Option<ShapeProperties>,
    pub(super) text_properties: Option<TextProperties>,
    scaling_extension_list: Option<ExtensionList>,
    pub(super) extension_list: Option<ExtensionList>,
}

impl ParsedAxisCommon {
    fn new() -> Self {
        Self {
            axis_id: None,
            cross_axis_id: None,
            position: None,
            title: None,
            title_layout: None,
            title_overlay: false,
            title_shape_properties: None,
            title_text_properties: None,
            title_extension_list: None,
            number_format: None,
            orientation: AxisOrientation::MinMax,
            major_tick_mark: TickMark::Out,
            minor_tick_mark: TickMark::None,
            tick_label_position: TickLabelPosition::NextTo,
            deleted: false,
            cross_mode: AxisCrossMode::AutoZero,
            crosses_at: None,
            show_major_gridlines: false,
            show_minor_gridlines: false,
            major_gridlines: None,
            minor_gridlines: None,
            shape_properties: None,
            text_properties: None,
            scaling_extension_list: None,
            extension_list: None,
        }
    }

    fn finish(self) -> Result<AxisCommon> {
        let axis_id = self
            .axis_id
            .ok_or_else(|| missing_attribute("chart axis ID"))?;
        let position = self
            .position
            .ok_or_else(|| missing_attribute("chart axis position"))?;
        let cross_axis_id = self
            .cross_axis_id
            .ok_or_else(|| missing_attribute("chart crossing-axis ID"))?;
        let mut common = AxisCommon::new(axis_id, position, cross_axis_id);
        common.title = self.title;
        common.layout = self.title_layout;
        common.title_overlay = self.title_overlay;
        common.title_shape_properties = self.title_shape_properties;
        common.title_text_properties = self.title_text_properties;
        common.title_extension_list = self.title_extension_list;
        common.number_format = self.number_format;
        common.orientation = self.orientation;
        common.major_tick_mark = self.major_tick_mark;
        common.minor_tick_mark = self.minor_tick_mark;
        common.tick_label_position = self.tick_label_position;
        common.deleted = self.deleted;
        common.cross_mode = self.cross_mode;
        common.crosses_at = self.crosses_at;
        common.show_major_gridlines = self.show_major_gridlines;
        common.show_minor_gridlines = self.show_minor_gridlines;
        common.major_gridlines = self.major_gridlines;
        common.minor_gridlines = self.minor_gridlines;
        common.shape_properties = self.shape_properties;
        common.text_properties = self.text_properties;
        common.scaling_extension_list = self.scaling_extension_list;
        common.extension_list = self.extension_list;
        Ok(common)
    }
}

pub(crate) fn parse_axis_title<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
    common: &mut ParsedAxisCommon,
) -> Result<()> {
    if common.title.is_some() {
        return Err(Error::Invalid(
            "chart axis contains duplicate titles".into(),
        ));
    }
    let title = parse_title(reader)?;
    common.title = Some(title.text);
    common.title_layout = title.layout;
    common.title_overlay = title.overlay;
    common.title_shape_properties = title.shape_properties;
    common.title_text_properties = title.text_properties;
    common.title_extension_list = title.extension_list;
    Ok(())
}

pub(crate) fn parse_axis_common_element(
    common: &mut ParsedAxisCommon,
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<bool> {
    match element.local_name().as_ref() {
        b"axId" => common.axis_id = Some(required_u32_attr(element, "chart axis ID")?),
        b"orientation" => common.orientation = parse_axis_orientation(element)?,
        b"delete" => common.deleted = parse_bool_attr(element)?,
        b"axPos" => common.position = Some(parse_axis_position(element)?),
        b"numFmt" => {
            common.number_format = Some(parse_number_format(element, decoder, "chart axis")?)
        },
        b"majorTickMark" => common.major_tick_mark = parse_tick_mark(element)?,
        b"minorTickMark" => common.minor_tick_mark = parse_tick_mark(element)?,
        b"tickLblPos" => common.tick_label_position = parse_tick_label_position(element)?,
        b"crossAx" => {
            common.cross_axis_id = Some(required_u32_attr(element, "chart crossing-axis ID")?);
        },
        b"crosses" => common.cross_mode = parse_axis_cross_mode(element)?,
        b"crossesAt" => {
            common.crosses_at = Some(required_f64_attr(element, "chart axis crossing value")?);
        },
        _ => return Ok(false),
    }
    Ok(true)
}

pub(crate) fn is_axis_common_fragment(element: &BytesStart<'_>) -> bool {
    matches!(
        element.local_name().as_ref(),
        b"majorGridlines" | b"minorGridlines" | b"spPr" | b"txPr"
    )
}

pub(crate) fn parse_axis_common_fragment<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
    common: &mut ParsedAxisCommon,
    element: &BytesStart<'_>,
    empty: bool,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"majorGridlines" => {
            let lines = if empty {
                Lines::new()
            } else {
                parse_chart_lines(reader, b"majorGridlines")?
            };
            set_chart_lines(
                &mut common.major_gridlines,
                lines,
                "chart axis major gridlines",
            )?;
            common.show_major_gridlines = true;
        },
        b"minorGridlines" => {
            let lines = if empty {
                Lines::new()
            } else {
                parse_chart_lines(reader, b"minorGridlines")?
            };
            set_chart_lines(
                &mut common.minor_gridlines,
                lines,
                "chart axis minor gridlines",
            )?;
            common.show_minor_gridlines = true;
        },
        b"spPr" => {
            if common.shape_properties.is_some() {
                return Err(Error::Invalid(
                    "chart axis contains duplicate shape properties".into(),
                ));
            }
            let xml = if empty {
                reader.capture_empty_fragment(element)?
            } else {
                reader.capture_fragment(element, "chart axis shape properties")?
            };
            common.shape_properties = Some(ShapeProperties::from_xml(xml)?);
        },
        b"txPr" => {
            if common.text_properties.is_some() {
                return Err(Error::Invalid(
                    "chart axis contains duplicate text properties".into(),
                ));
            }
            let xml = if empty {
                reader.capture_empty_fragment(element)?
            } else {
                reader.capture_fragment(element, "chart axis text properties")?
            };
            common.text_properties = Some(TextProperties::from_xml(xml)?);
        },
        _ => {},
    }
    Ok(())
}

pub(crate) fn parse_axis_extension<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
    common: &mut ParsedAxisCommon,
    element: &BytesStart<'_>,
    empty: bool,
    scaling: bool,
) -> Result<()> {
    let target = if scaling {
        &mut common.scaling_extension_list
    } else {
        &mut common.extension_list
    };
    if target.is_some() {
        return Err(Error::Invalid(format!(
            "chart axis contains duplicate {} extension lists",
            if scaling { "scaling" } else { "axis" }
        )));
    }
    let xml = if empty {
        reader.capture_empty_fragment(element)?
    } else {
        reader.capture_fragment(
            element,
            if scaling {
                "chart axis scaling extension list"
            } else {
                "chart axis extension list"
            },
        )?
    };
    *target = Some(ExtensionList::from_xml(xml)?);
    Ok(())
}

pub(crate) fn parse_category_axis<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Option<CategoryAxis>> {
    let mut common = ParsedAxisCommon::new();
    let mut scaling = ParsedAxisScaling::default();
    let mut auto = true;
    let mut label_align = None;
    let mut label_offset = None;
    let mut tick_label_skip = None;
    let mut tick_mark_skip = None;
    let mut no_multi_level = false;
    let mut in_scaling = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"scaling" => {
                in_scaling = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_axis_extension(reader, &mut common, element, false, in_scaling)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_axis_extension(reader, &mut common, element, true, in_scaling)?;
            },
            Ok(Event::Start(ref element)) if is_axis_common_fragment(element) => {
                parse_axis_common_fragment(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if is_axis_common_fragment(element) => {
                parse_axis_common_fragment(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"title" => {
                parse_axis_title(reader, &mut common)?;
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if !parse_axis_common_element(&mut common, element, reader.decoder())? =>
            {
                if in_scaling && scaling.parse_element(element)? {
                    buf.clear();
                    continue;
                }
                match element.local_name().as_ref() {
                    b"auto" => auto = parse_bool_attr(element)?,
                    b"lblAlgn" => label_align = Some(parse_axis_label_align(element)?),
                    b"lblOffset" => {
                        label_offset = Some(bounded_u32_attr(
                            element,
                            "chart axis label offset",
                            0,
                            1000,
                        )?);
                    },
                    b"tickLblSkip" => {
                        tick_label_skip = Some(required_positive_u32_attr(
                            element,
                            "chart tick-label skip",
                        )?);
                    },
                    b"tickMarkSkip" => {
                        tick_mark_skip =
                            Some(required_positive_u32_attr(element, "chart tick-mark skip")?);
                    },
                    b"noMultiLvlLbl" => no_multi_level = parse_bool_attr(element)?,
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"catAx" => break,
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"scaling" => {
                in_scaling = false;
            },
            Ok(Event::Eof) => return Err(unterminated_axis("category")),
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    scaling.validate()?;
    let common = common.finish()?;
    let mut axis = CategoryAxis::new(common.axis_id, common.position, common.cross_axis_id);
    axis.common = common;
    axis.min = scaling.min;
    axis.max = scaling.max;
    axis.log_base = scaling.log_base;
    axis.auto = auto;
    axis.label_align = label_align;
    axis.label_offset = label_offset;
    axis.tick_label_skip = tick_label_skip;
    axis.tick_mark_skip = tick_mark_skip;
    axis.no_multi_level = no_multi_level;
    Ok(Some(axis))
}

pub(crate) fn parse_value_axis<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Option<ValueAxis>> {
    let mut common = ParsedAxisCommon::new();
    let mut scaling = ParsedAxisScaling::default();
    let mut major_unit = None;
    let mut minor_unit = None;
    let mut display_units = None;
    let mut cross_between = AxisCrossBetween::Between;
    let mut in_scaling = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"scaling" => {
                in_scaling = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_axis_extension(reader, &mut common, element, false, in_scaling)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_axis_extension(reader, &mut common, element, true, in_scaling)?;
            },
            Ok(Event::Start(ref element)) if is_axis_common_fragment(element) => {
                parse_axis_common_fragment(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if is_axis_common_fragment(element) => {
                parse_axis_common_fragment(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"title" => {
                parse_axis_title(reader, &mut common)?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"dispUnits" => {
                if display_units.is_some() {
                    return Err(Error::Invalid(
                        "chart value axis contains duplicate display units".into(),
                    ));
                }
                display_units = Some(parse_display_units(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"dispUnits" => {
                return Err(Error::Invalid(
                    "chart display units are missing their unit".into(),
                ));
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if !parse_axis_common_element(&mut common, element, reader.decoder())? =>
            {
                if in_scaling && scaling.parse_element(element)? {
                    buf.clear();
                    continue;
                }
                match element.local_name().as_ref() {
                    b"majorUnit" => {
                        major_unit = Some(required_positive_f64_attr(
                            element,
                            "chart axis major unit",
                        )?);
                    },
                    b"minorUnit" => {
                        minor_unit = Some(required_positive_f64_attr(
                            element,
                            "chart axis minor unit",
                        )?);
                    },
                    b"crossBetween" => cross_between = parse_axis_cross_between(element)?,
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"valAx" => break,
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"scaling" => {
                in_scaling = false;
            },
            Ok(Event::Eof) => return Err(unterminated_axis("value")),
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    scaling.validate()?;
    let common = common.finish()?;
    let mut axis = ValueAxis::new(common.axis_id, common.position, common.cross_axis_id);
    axis.common = common;
    axis.min = scaling.min;
    axis.max = scaling.max;
    axis.major_unit = major_unit;
    axis.minor_unit = minor_unit;
    axis.log_base = scaling.log_base;
    axis.cross_between = cross_between;
    axis.display_units = display_units.map(Box::new);
    Ok(Some(axis))
}

pub(crate) fn parse_display_units<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<DisplayUnits> {
    let mut built_in_unit = None;
    let mut custom_unit = None;
    let mut parsed_label = ParsedDisplayUnitsLabel::default();
    let mut extension_list = None;
    let mut saw_label = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"dispUnitsLbl" => {
                if saw_label {
                    return Err(Error::Invalid(
                        "chart display units contain duplicate labels".into(),
                    ));
                }
                saw_label = true;
                parsed_label = parse_display_units_label(reader)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"dispUnitsLbl" => {
                if saw_label {
                    return Err(Error::Invalid(
                        "chart display units contain duplicate labels".into(),
                    ));
                }
                saw_label = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart display units contain duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(element, "chart display-units extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart display units contain duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"builtInUnit" => {
                        if built_in_unit.is_some() {
                            return Err(Error::Invalid(
                                "chart display units contain duplicate built-in units".into(),
                            ));
                        }
                        built_in_unit = Some(parse_built_in_unit(element)?);
                    },
                    b"custUnit" => {
                        if custom_unit.is_some() {
                            return Err(Error::Invalid(
                                "chart display units contain duplicate custom units".into(),
                            ));
                        }
                        custom_unit = Some(required_positive_f64_attr(
                            element,
                            "chart custom display unit",
                        )?);
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"dispUnits" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart display units".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    if built_in_unit.is_some() == custom_unit.is_some() {
        return Err(Error::Invalid(
            "chart display units require exactly one built-in or custom unit".into(),
        ));
    }
    Ok(DisplayUnits {
        built_in_unit,
        custom_unit,
        show_label: saw_label,
        label: parsed_label.label,
        layout: parsed_label.layout,
        label_shape_properties: parsed_label.shape_properties,
        label_text_properties: parsed_label.text_properties,
        extension_list,
    })
}

#[derive(Default)]
pub(crate) struct ParsedDisplayUnitsLabel {
    label: Option<TitleText>,
    pub(super) layout: Option<Layout>,
    pub(super) shape_properties: Option<ShapeProperties>,
    pub(super) text_properties: Option<TextProperties>,
}

pub(crate) fn parse_display_units_label<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<ParsedDisplayUnitsLabel> {
    let mut text = String::new();
    let mut formula = None;
    let mut layout = None;
    let mut in_text = false;
    let mut saw_text = false;
    let mut shape_properties = None;
    let mut text_properties = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"layout" => {
                if layout.is_some() {
                    return Err(Error::Invalid(
                        "chart display-units label contains duplicate layouts".into(),
                    ));
                }
                layout = Some(parse_layout(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"layout" => {
                layout = Some(match layout {
                    None => Layout::new(),
                    Some(_) => {
                        return Err(Error::Invalid(
                            "chart display-units label contains duplicate layouts".into(),
                        ));
                    },
                });
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart display-units label contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader
                        .capture_fragment(element, "chart display-units label shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart display-units label contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart display-units label contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(TextProperties::from_xml(
                    reader
                        .capture_fragment(element, "chart display-units label text properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart display-units label contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(TextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"f" => {
                if formula.is_some() {
                    return Err(Error::Invalid(
                        "chart display-units label contains duplicate formula references".into(),
                    ));
                }
                formula = Some(parse_text_element(reader, b"f")?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"t" => {
                saw_text = true;
                in_text = true;
            },
            Ok(Event::Text(value)) if in_text => {
                text.push_str(
                    &value
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?,
                );
            },
            Ok(Event::CData(value)) if in_text => {
                text.push_str(
                    &value
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?,
                );
            },
            Ok(Event::GeneralRef(reference)) if in_text => {
                text.push_str(&decode_xml_reference(&reference)?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"t" => {
                in_text = false;
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"dispUnitsLbl" => {
                break;
            },
            Ok(Event::Eof) => {
                return Err(Error::Invalid(
                    "unterminated chart display-units label".into(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let label = if let Some(formula) = formula {
        if saw_text {
            return Err(Error::Invalid(
                "chart display-units label mixes a formula reference with literal text".into(),
            ));
        }
        Some(TitleText::Reference(DataSourceRef::new(formula)))
    } else if saw_text {
        Some(TitleText::Literal(RichText::new(text)))
    } else {
        None
    };
    Ok(ParsedDisplayUnitsLabel {
        label,
        layout,
        shape_properties,
        text_properties,
    })
}

pub(crate) fn parse_date_axis<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Option<DateAxis>> {
    let mut common = ParsedAxisCommon::new();
    let mut scaling = ParsedAxisScaling::default();
    let mut major_unit = None;
    let mut minor_unit = None;
    let mut major_time_unit = None;
    let mut minor_time_unit = None;
    let mut base_time_unit = None;
    let mut auto = true;
    let mut label_offset = None;
    let mut in_scaling = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"scaling" => {
                in_scaling = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_axis_extension(reader, &mut common, element, false, in_scaling)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_axis_extension(reader, &mut common, element, true, in_scaling)?;
            },
            Ok(Event::Start(ref element)) if is_axis_common_fragment(element) => {
                parse_axis_common_fragment(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if is_axis_common_fragment(element) => {
                parse_axis_common_fragment(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"title" => {
                parse_axis_title(reader, &mut common)?;
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if !parse_axis_common_element(&mut common, element, reader.decoder())? =>
            {
                if in_scaling && scaling.parse_element(element)? {
                    buf.clear();
                    continue;
                }
                match element.local_name().as_ref() {
                    b"majorUnit" => {
                        major_unit = Some(required_positive_f64_attr(
                            element,
                            "chart date-axis major unit",
                        )?);
                    },
                    b"minorUnit" => {
                        minor_unit = Some(required_positive_f64_attr(
                            element,
                            "chart date-axis minor unit",
                        )?);
                    },
                    b"majorTimeUnit" => major_time_unit = Some(parse_time_unit(element)?),
                    b"minorTimeUnit" => minor_time_unit = Some(parse_time_unit(element)?),
                    b"baseTimeUnit" => base_time_unit = Some(parse_time_unit(element)?),
                    b"auto" => auto = parse_bool_attr(element)?,
                    b"lblOffset" => {
                        label_offset = Some(bounded_u32_attr(
                            element,
                            "chart date-axis label offset",
                            0,
                            1000,
                        )?);
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"dateAx" => break,
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"scaling" => {
                in_scaling = false;
            },
            Ok(Event::Eof) => return Err(unterminated_axis("date")),
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    scaling.validate()?;
    let common = common.finish()?;
    let mut axis = DateAxis::new(common.axis_id, common.position, common.cross_axis_id);
    axis.common = common;
    axis.min = scaling.min;
    axis.max = scaling.max;
    axis.log_base = scaling.log_base;
    axis.major_unit = major_unit;
    axis.minor_unit = minor_unit;
    axis.major_time_unit = major_time_unit;
    axis.minor_time_unit = minor_time_unit;
    axis.base_time_unit = base_time_unit;
    axis.auto = auto;
    axis.label_offset = label_offset;
    Ok(Some(axis))
}

pub(crate) fn parse_series_axis<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Option<SeriesAxis>> {
    let mut common = ParsedAxisCommon::new();
    let mut scaling = ParsedAxisScaling::default();
    let mut tick_label_skip = None;
    let mut tick_mark_skip = None;
    let mut in_scaling = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"scaling" => {
                in_scaling = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_axis_extension(reader, &mut common, element, false, in_scaling)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_axis_extension(reader, &mut common, element, true, in_scaling)?;
            },
            Ok(Event::Start(ref element)) if is_axis_common_fragment(element) => {
                parse_axis_common_fragment(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if is_axis_common_fragment(element) => {
                parse_axis_common_fragment(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"title" => {
                parse_axis_title(reader, &mut common)?;
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if !parse_axis_common_element(&mut common, element, reader.decoder())? =>
            {
                if in_scaling && scaling.parse_element(element)? {
                    buf.clear();
                    continue;
                }
                match element.local_name().as_ref() {
                    b"tickLblSkip" => {
                        tick_label_skip = Some(required_positive_u32_attr(
                            element,
                            "chart series-axis tick-label skip",
                        )?);
                    },
                    b"tickMarkSkip" => {
                        tick_mark_skip = Some(required_positive_u32_attr(
                            element,
                            "chart series-axis tick-mark skip",
                        )?);
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"serAx" => break,
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"scaling" => {
                in_scaling = false;
            },
            Ok(Event::Eof) => return Err(unterminated_axis("series")),
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    scaling.validate()?;
    let common = common.finish()?;
    let mut axis = SeriesAxis::new(common.axis_id, common.position, common.cross_axis_id);
    axis.common = common;
    axis.min = scaling.min;
    axis.max = scaling.max;
    axis.log_base = scaling.log_base;
    axis.tick_label_skip = tick_label_skip;
    axis.tick_mark_skip = tick_mark_skip;
    Ok(Some(axis))
}

pub(crate) fn unterminated_axis(kind: &str) -> Error {
    Error::Invalid(format!("unterminated chart {kind} axis"))
}
