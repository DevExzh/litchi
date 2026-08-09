use super::{
    Area3DTypeGroup, AreaTypeGroup, Axis, BandFormat, Bar3DTypeGroup, BarDirection, BarGrouping,
    BarTypeGroup, BubbleScale, BubbleSize, BubbleTypeGroup, BufRead, BytesStart, ChartXmlReader,
    DataLabels, DataSourceRef, DataTable, Decoder, DoughnutTypeGroup, Error, Event, ExtensionList,
    Layout, LayoutMode, LayoutTarget, Line3DTypeGroup, LineTypeGroup, Lines, OfPieSplitType,
    OfPieType, OfPieTypeGroup, ParsedTitle, PictureFormat, PictureOptions, Pie3DTypeGroup,
    PieTypeGroup, PlotArea, RadarStyle, RadarTypeGroup, Result, RichText, ScatterStyle,
    ScatterTypeGroup, ShapeProperties, StockTypeGroup, Surface3DTypeGroup, SurfaceTypeGroup,
    TextProperties, TitleText, TypeGroup, TypeGroupCommon, UpDownBars, View3D, WallFloor,
    bounded_percentage_i32_attr, bounded_percentage_u32_attr, bounded_u32_attr,
    consume_empty_chart_element, decode_xml_reference, get_attr, invalid_attribute,
    missing_attribute, parse_bar_shape, parse_bool_attr, parse_category_axis, parse_data_labels,
    parse_date_axis, parse_grouping, parse_series, parse_series_axis, parse_text_element,
    parse_value_axis, required_enum_attr, required_f64_attr, required_positive_f64_attr,
    required_string_attr, required_u32_attr,
};

pub(crate) fn parse_title<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<ParsedTitle> {
    let mut text = String::new();
    let mut formula = None;
    let mut layout = None;
    let mut overlay = false;
    let mut saw_overlay = false;
    let mut shape_properties = None;
    let mut text_properties = None;
    let mut extension_list = None;
    let mut buf = Vec::new();
    let mut in_text = false;
    let mut saw_text = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"layout" => {
                if layout.is_some() {
                    return Err(Error::Invalid(
                        "chart title contains duplicate layouts".into(),
                    ));
                }
                layout = Some(parse_layout(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"layout" => {
                layout = Some(match layout {
                    None => Layout::new(),
                    Some(_) => {
                        return Err(Error::Invalid(
                            "chart title contains duplicate layouts".into(),
                        ));
                    },
                });
            },
            Ok(Event::Start(ref element) | Event::Empty(ref element))
                if element.local_name().as_ref() == b"overlay" =>
            {
                if saw_overlay {
                    return Err(Error::Invalid(
                        "chart title contains duplicate overlay flags".into(),
                    ));
                }
                overlay = parse_bool_attr(element)?;
                saw_overlay = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart title contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart title shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart title contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart title contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(TextProperties::from_xml(
                    reader.capture_fragment(element, "chart title text properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart title contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(TextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart title contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(element, "chart title extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart title contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"f" => {
                if formula.is_some() {
                    return Err(Error::Invalid(
                        "chart title contains duplicate formula references".to_string(),
                    ));
                }
                formula = Some(parse_text_element(reader, b"f")?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"t" => {
                saw_text = true;
                in_text = true;
            },
            Ok(Event::Text(e)) if in_text => {
                text.push_str(&e.decode().map_err(|e| Error::Xml(e.to_string()))?);
            },
            Ok(Event::CData(e)) if in_text => {
                text.push_str(&e.decode().map_err(|e| Error::Xml(e.to_string()))?);
            },
            Ok(Event::GeneralRef(reference)) if in_text => {
                text.push_str(&decode_xml_reference(&reference)?);
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"t" => {
                in_text = false;
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"title" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart element".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let text = if let Some(formula) = formula {
        if saw_text {
            return Err(Error::Invalid(
                "chart title mixes a formula reference with literal text".to_string(),
            ));
        }
        TitleText::Reference(DataSourceRef::new(formula))
    } else {
        TitleText::Literal(RichText::new(text))
    };
    Ok(ParsedTitle {
        text,
        layout,
        overlay,
        shape_properties,
        text_properties,
        extension_list,
    })
}

pub(crate) fn parse_view_3d<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<View3D> {
    let mut view = View3D::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"rotX" => {
                        view.rot_x = Some(bounded_u32_attr(e, "chart X rotation", 0, 90)?);
                    },
                    b"rotY" => {
                        view.rot_y = Some(bounded_u32_attr(e, "chart Y rotation", 0, 360)?);
                    },
                    b"perspective" => {
                        view.perspective = Some(bounded_u32_attr(e, "chart perspective", 0, 240)?);
                    },
                    b"hPercent" => {
                        view.height_percent =
                            Some(bounded_u32_attr(e, "chart height percentage", 5, 500)?);
                    },
                    b"depthPercent" => {
                        view.depth_percent =
                            Some(bounded_u32_attr(e, "chart depth percentage", 20, 2000)?);
                    },
                    b"rAngAx" => view.right_angle_axes = parse_bool_attr(e)?,
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"view3D" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart element".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    Ok(view)
}

pub(crate) fn parse_wall_floor<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<WallFloor> {
    let mut wall_floor = WallFloor::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"thickness" => {
                if wall_floor.thickness.is_some() {
                    return Err(Error::Invalid(
                        "chart surface contains duplicate thickness values".into(),
                    ));
                }
                wall_floor.thickness = Some(bounded_u32_attr(e, "chart wall thickness", 0, 4096)?);
                consume_empty_chart_element(reader, b"thickness", "chart surface thickness")?;
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"thickness" => {
                if wall_floor.thickness.is_some() {
                    return Err(Error::Invalid(
                        "chart surface contains duplicate thickness values".into(),
                    ));
                }
                wall_floor.thickness = Some(bounded_u32_attr(e, "chart wall thickness", 0, 4096)?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if wall_floor.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart surface contains duplicate shape properties".into(),
                    ));
                }
                wall_floor.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(e, "chart surface shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if wall_floor.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart surface contains duplicate shape properties".into(),
                    ));
                }
                wall_floor.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(e)?,
                )?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pictureOptions" => {
                if wall_floor.picture_options.is_some() {
                    return Err(Error::Invalid(
                        "chart surface contains duplicate picture options".into(),
                    ));
                }
                wall_floor.picture_options = Some(parse_picture_options(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"pictureOptions" => {
                if wall_floor.picture_options.is_some() {
                    return Err(Error::Invalid(
                        "chart surface contains duplicate picture options".into(),
                    ));
                }
                wall_floor.picture_options = Some(PictureOptions::default());
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if wall_floor.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart surface contains duplicate extension lists".into(),
                    ));
                }
                wall_floor.extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(e, "chart surface extension list")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if wall_floor.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart surface contains duplicate extension lists".into(),
                    ));
                }
                wall_floor.extension_list =
                    Some(ExtensionList::from_xml(reader.capture_empty_fragment(e)?)?);
            },
            Ok(Event::End(ref e)) => {
                let tag_name = e.local_name();
                if tag_name.as_ref() == b"floor"
                    || tag_name.as_ref() == b"backWall"
                    || tag_name.as_ref() == b"sideWall"
                {
                    break;
                }
            },
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart element".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    Ok(wall_floor)
}

pub(crate) fn parse_picture_options<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<PictureOptions> {
    let mut options = PictureOptions::default();
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(ref element))
                if is_picture_option_child(element.local_name().as_ref()) =>
            {
                parse_picture_option_child(&mut options, element, reader.decoder())?;
                let name = element.local_name().as_ref().to_vec();
                consume_empty_chart_element(reader, &name, "chart picture option")?;
            },
            Ok(Event::Empty(ref element))
                if is_picture_option_child(element.local_name().as_ref()) =>
            {
                parse_picture_option_child(&mut options, element, reader.decoder())?;
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"pictureOptions" => {
                break;
            },
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart picture options".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buffer.clear();
    }
    Ok(options)
}

pub(crate) fn is_picture_option_child(name: &[u8]) -> bool {
    matches!(
        name,
        b"applyToFront" | b"applyToSides" | b"applyToEnd" | b"pictureFormat" | b"pictureStackUnit"
    )
}

pub(crate) fn parse_picture_option_child(
    options: &mut PictureOptions,
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<()> {
    let target = match element.local_name().as_ref() {
        b"applyToFront" => Some(&mut options.apply_to_front),
        b"applyToSides" => Some(&mut options.apply_to_sides),
        b"applyToEnd" => Some(&mut options.apply_to_end),
        b"pictureFormat" => {
            if options.picture_format.is_some() {
                return Err(Error::Invalid(
                    "chart picture options contain duplicate formats".into(),
                ));
            }
            let value = required_string_attr(element, b"val", decoder, "chart picture format")?;
            options.picture_format = Some(match value.as_str() {
                "stretch" => PictureFormat::Stretch,
                "stack" => PictureFormat::Stack,
                "stackScale" => PictureFormat::StackScale,
                _ => {
                    return Err(Error::Invalid(format!(
                        "invalid chart picture format '{value}'"
                    )));
                },
            });
            None
        },
        b"pictureStackUnit" => {
            if options.picture_stack_unit.is_some() {
                return Err(Error::Invalid(
                    "chart picture options contain duplicate stack units".into(),
                ));
            }
            options.picture_stack_unit = Some(required_positive_f64_attr(
                element,
                "chart picture stack unit",
            )?);
            None
        },
        _ => {
            return Err(Error::Invalid(format!(
                "invalid chart picture option '{}'",
                String::from_utf8_lossy(element.local_name().as_ref())
            )));
        },
    };
    if let Some(target) = target {
        if target.is_some() {
            return Err(Error::Invalid(
                "chart picture options contain a duplicate apply flag".into(),
            ));
        }
        *target = Some(parse_bool_attr(element)?);
    }
    Ok(())
}

pub(crate) fn parse_plot_area<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<PlotArea> {
    let mut plot_area = PlotArea::new();
    let mut saw_data_table = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"layout" => {
                plot_area.layout = Some(parse_layout(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"layout" => {
                plot_area.layout = Some(Layout::new());
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dTable" => {
                if saw_data_table {
                    return Err(Error::Invalid(
                        "chart plot area contains duplicate data tables".into(),
                    ));
                }
                saw_data_table = true;
                plot_area.data_table = Some(parse_data_table(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"dTable" => {
                if saw_data_table {
                    return Err(Error::Invalid(
                        "chart plot area contains duplicate data tables".into(),
                    ));
                }
                saw_data_table = true;
                plot_area.data_table = Some(DataTable::default());
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if plot_area.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart plot area contains duplicate shape properties".into(),
                    ));
                }
                plot_area.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(e, "chart plot-area shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if plot_area.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart plot area contains duplicate shape properties".into(),
                    ));
                }
                plot_area.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(e)?,
                )?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if plot_area.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart plot area contains duplicate extension lists".into(),
                    ));
                }
                plot_area.extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(e, "chart plot-area extension list")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if plot_area.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart plot area contains duplicate extension lists".into(),
                    ));
                }
                plot_area.extension_list =
                    Some(ExtensionList::from_xml(reader.capture_empty_fragment(e)?)?);
            },
            Ok(Event::Empty(ref e))
                if is_chart_type_group_name(e.local_name().as_ref())
                    || matches!(
                        e.local_name().as_ref(),
                        b"catAx" | b"valAx" | b"dateAx" | b"serAx"
                    ) =>
            {
                return Err(Error::Invalid(format!(
                    "chart plot-area element {} cannot be empty",
                    String::from_utf8_lossy(e.local_name().as_ref())
                )));
            },
            Ok(Event::Start(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"barChart" => {
                        if let Some(group) = parse_bar_chart(reader)? {
                            plot_area.type_groups.push(TypeGroup::Bar(group));
                        }
                    },
                    b"bar3DChart" => {
                        if let Some(group) = parse_bar_3d_chart(reader)? {
                            plot_area.type_groups.push(TypeGroup::Bar3D(group));
                        }
                    },
                    b"lineChart" => {
                        if let Some(group) = parse_line_chart(reader)? {
                            plot_area.type_groups.push(TypeGroup::Line(group));
                        }
                    },
                    b"pieChart" => {
                        if let Some(group) = parse_pie_chart(reader)? {
                            plot_area.type_groups.push(TypeGroup::Pie(group));
                        }
                    },
                    b"areaChart" => {
                        if let Some(group) = parse_area_chart(reader)? {
                            plot_area.type_groups.push(TypeGroup::Area(group));
                        }
                    },
                    b"area3DChart" => {
                        plot_area
                            .type_groups
                            .push(TypeGroup::Area3D(parse_area_3d_chart(reader)?));
                    },
                    b"bubbleChart" => {
                        plot_area
                            .type_groups
                            .push(TypeGroup::Bubble(parse_bubble_chart(reader)?));
                    },
                    b"doughnutChart" => {
                        plot_area
                            .type_groups
                            .push(TypeGroup::Doughnut(parse_doughnut_chart(reader)?));
                    },
                    b"line3DChart" => {
                        plot_area
                            .type_groups
                            .push(TypeGroup::Line3D(parse_line_3d_chart(reader)?));
                    },
                    b"ofPieChart" => {
                        plot_area
                            .type_groups
                            .push(TypeGroup::OfPie(parse_of_pie_chart(reader)?));
                    },
                    b"pie3DChart" => {
                        plot_area
                            .type_groups
                            .push(TypeGroup::Pie3D(parse_pie_3d_chart(reader)?));
                    },
                    b"radarChart" => {
                        plot_area
                            .type_groups
                            .push(TypeGroup::Radar(parse_radar_chart(reader)?));
                    },
                    b"scatterChart" => {
                        if let Some(group) = parse_scatter_chart(reader)? {
                            plot_area.type_groups.push(TypeGroup::Scatter(group));
                        }
                    },
                    b"stockChart" => {
                        plot_area
                            .type_groups
                            .push(TypeGroup::Stock(parse_stock_chart(reader)?));
                    },
                    b"surfaceChart" => {
                        plot_area
                            .type_groups
                            .push(TypeGroup::Surface(parse_surface_chart(reader)?));
                    },
                    b"surface3DChart" => {
                        plot_area
                            .type_groups
                            .push(TypeGroup::Surface3D(parse_surface_3d_chart(reader)?));
                    },
                    b"catAx" => {
                        if let Some(axis) = parse_category_axis(reader)? {
                            plot_area.axes.push(Axis::Category(axis));
                        }
                    },
                    b"valAx" => {
                        if let Some(axis) = parse_value_axis(reader)? {
                            plot_area.axes.push(Axis::Value(axis));
                        }
                    },
                    b"dateAx" => {
                        if let Some(axis) = parse_date_axis(reader)? {
                            plot_area.axes.push(Axis::Date(axis));
                        }
                    },
                    b"serAx" => {
                        if let Some(axis) = parse_series_axis(reader)? {
                            plot_area.axes.push(Axis::Series(axis));
                        }
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"plotArea" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart element".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    Ok(plot_area)
}

pub(crate) fn is_chart_type_group_name(name: &[u8]) -> bool {
    matches!(
        name,
        b"areaChart"
            | b"area3DChart"
            | b"barChart"
            | b"bar3DChart"
            | b"bubbleChart"
            | b"doughnutChart"
            | b"lineChart"
            | b"line3DChart"
            | b"ofPieChart"
            | b"pieChart"
            | b"pie3DChart"
            | b"radarChart"
            | b"scatterChart"
            | b"stockChart"
            | b"surfaceChart"
            | b"surface3DChart"
    )
}

pub(crate) fn parse_data_table<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<DataTable> {
    let mut data_table = DataTable::default();
    let mut seen = [false; 4];
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if data_table.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart data table contains duplicate shape properties".into(),
                    ));
                }
                data_table.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart data-table shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if data_table.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart data table contains duplicate shape properties".into(),
                    ));
                }
                data_table.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if data_table.text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart data table contains duplicate text properties".into(),
                    ));
                }
                data_table.text_properties = Some(TextProperties::from_xml(
                    reader.capture_fragment(element, "chart data-table text properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if data_table.text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart data table contains duplicate text properties".into(),
                    ));
                }
                data_table.text_properties = Some(TextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if data_table.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart data table contains duplicate extension lists".into(),
                    ));
                }
                data_table.extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(element, "chart data-table extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if data_table.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart data table contains duplicate extension lists".into(),
                    ));
                }
                data_table.extension_list = Some(ExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element) | Event::Empty(ref element)) => {
                let field = match element.local_name().as_ref() {
                    b"showHorzBorder" => Some((0, &mut data_table.show_horizontal_border)),
                    b"showVertBorder" => Some((1, &mut data_table.show_vertical_border)),
                    b"showOutline" => Some((2, &mut data_table.show_outline)),
                    b"showKeys" => Some((3, &mut data_table.show_legend_keys)),
                    _ => None,
                };
                if let Some((index, field)) = field {
                    if seen[index] {
                        return Err(Error::Invalid(
                            "chart data table contains a duplicate visibility setting".into(),
                        ));
                    }
                    seen[index] = true;
                    *field = parse_bool_attr(element)?;
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"dTable" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart data table".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(data_table)
}

pub(crate) fn parse_layout<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Layout> {
    let mut layout = Layout::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"layoutTarget" => {
                    layout.target = Some(
                        match required_enum_attr(e, "chart layout target")?.as_str() {
                            "inner" => LayoutTarget::Inner,
                            "outer" => LayoutTarget::Outer,
                            value => {
                                return Err(invalid_attribute(
                                    "chart layout target",
                                    value.as_bytes(),
                                ));
                            },
                        },
                    );
                },
                b"xMode" => layout.x_mode = Some(parse_layout_mode(e)?),
                b"yMode" => layout.y_mode = Some(parse_layout_mode(e)?),
                b"wMode" => layout.width_mode = Some(parse_layout_mode(e)?),
                b"hMode" => layout.height_mode = Some(parse_layout_mode(e)?),
                b"x" => layout.x = Some(required_f64_attr(e, "chart layout X position")?),
                b"y" => layout.y = Some(required_f64_attr(e, "chart layout Y position")?),
                b"w" => layout.width = Some(required_f64_attr(e, "chart layout width")?),
                b"h" => layout.height = Some(required_f64_attr(e, "chart layout height")?),
                _ => {},
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"layout" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart layout".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(layout)
}

pub(crate) fn parse_layout_mode(element: &BytesStart<'_>) -> Result<LayoutMode> {
    let value = get_attr(element, b"val").ok_or_else(|| missing_attribute("chart layout mode"))?;
    match value.as_slice() {
        b"edge" => Ok(LayoutMode::Edge),
        b"factor" => Ok(LayoutMode::Factor),
        _ => Err(invalid_attribute("chart layout mode", &value)),
    }
}

pub(crate) fn parse_common_type_group<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
    end_name: &[u8],
    supports_data_labels: bool,
    supports_axes: bool,
    mut drop_lines: Option<&mut Option<Lines>>,
    mut extra: impl FnMut(&BytesStart<'_>) -> Result<()>,
) -> Result<TypeGroupCommon> {
    let mut common = TypeGroupCommon::new();
    let mut saw_data_labels = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element))
                if drop_lines.is_some() && element.local_name().as_ref() == b"dropLines" =>
            {
                let lines = parse_chart_lines(reader, b"dropLines")?;
                if let Some(target) = drop_lines.as_deref_mut() {
                    set_chart_lines(target, lines, "chart drop lines")?;
                }
            },
            Ok(Event::Empty(ref element))
                if drop_lines.is_some() && element.local_name().as_ref() == b"dropLines" =>
            {
                if let Some(target) = drop_lines.as_deref_mut() {
                    set_chart_lines(target, Lines::new(), "chart drop lines")?;
                }
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref element))
                if supports_data_labels && element.local_name().as_ref() == b"dLbls" =>
            {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(parse_data_labels(reader)?);
            },
            Ok(Event::Empty(ref element))
                if supports_data_labels && element.local_name().as_ref() == b"dLbls" =>
            {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(DataLabels::default());
            },
            Ok(Event::Start(ref element) | Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"varyColors" => common.vary_colors = parse_bool_attr(element)?,
                    b"axId" if supports_axes => common
                        .axis_ids
                        .push(required_u32_attr(element, "chart type-group axis ID")?),
                    b"ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    _ => extra(element)?,
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == end_name => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart type group".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(common)
}

pub(crate) fn parse_type_group_extension<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
    common: &mut TypeGroupCommon,
    element: &BytesStart<'_>,
    empty: bool,
) -> Result<()> {
    if common.extension_list.is_some() {
        return Err(Error::Invalid(
            "chart type group contains duplicate extension lists".into(),
        ));
    }
    let xml = if empty {
        reader.capture_empty_fragment(element)?
    } else {
        reader.capture_fragment(element, "chart type-group extension list")?
    };
    common.extension_list = Some(ExtensionList::from_xml(xml)?);
    Ok(())
}

pub(crate) fn begin_group_data_labels(seen: &mut bool) -> Result<()> {
    if *seen {
        return Err(Error::Invalid(
            "chart type group contains duplicate data-label settings".into(),
        ));
    }
    *seen = true;
    Ok(())
}

pub(crate) fn set_chart_lines(
    target: &mut Option<Lines>,
    lines: Lines,
    description: &str,
) -> Result<()> {
    if target.replace(lines).is_some() {
        return Err(Error::Invalid(format!("{description} are duplicated")));
    }
    Ok(())
}

pub(crate) fn parse_chart_lines<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
    end_name: &[u8],
) -> Result<Lines> {
    let mut lines = Lines::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if lines.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart lines contain duplicate shape properties".into(),
                    ));
                }
                lines.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart-line shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if lines.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart lines contain duplicate shape properties".into(),
                    ));
                }
                lines.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == end_name => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart-line formatting".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(lines)
}

pub(crate) fn set_empty_up_down_bars(
    target: &mut Option<UpDownBars>,
    description: &str,
) -> Result<()> {
    if target.replace(UpDownBars::default()).is_some() {
        return Err(Error::Invalid(format!(
            "{description} contains duplicate up/down bars"
        )));
    }
    Ok(())
}

pub(crate) fn parse_area_3d_chart<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Area3DTypeGroup> {
    let mut grouping = BarGrouping::Standard;
    let mut gap_depth = None;
    let mut drop_lines = None;
    let common = parse_common_type_group(
        reader,
        b"area3DChart",
        true,
        true,
        Some(&mut drop_lines),
        |element| {
            match element.local_name().as_ref() {
                b"grouping" => grouping = parse_grouping(element)?,
                b"gapDepth" => {
                    gap_depth = Some(match get_attr(element, b"val") {
                        Some(_) => {
                            bounded_percentage_u32_attr(element, "area 3D gap depth", 0, 500)?
                        },
                        None => 150,
                    });
                },
                _ => {},
            }
            Ok(())
        },
    )?;
    let mut group = Area3DTypeGroup::new(grouping);
    group.common = common;
    group.gap_depth = gap_depth;
    group.drop_lines = drop_lines;
    Ok(group)
}

pub(crate) fn parse_bubble_chart<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<BubbleTypeGroup> {
    let mut group_bubble_3d = false;
    let mut scale = BubbleScale::default();
    let mut show_negative_bubbles = true;
    let mut size = BubbleSize::default();
    let common = parse_common_type_group(reader, b"bubbleChart", true, true, None, |element| {
        match element.local_name().as_ref() {
            b"bubble3D" => group_bubble_3d = parse_bool_attr(element)?,
            b"bubbleScale" => {
                let value = match get_attr(element, b"val") {
                    Some(_) => bounded_percentage_u32_attr(element, "bubble scale", 0, 300)?,
                    None => u32::from(BubbleScale::DEFAULT),
                };
                scale = BubbleScale::try_from(value)
                    .map_err(|error| Error::Invalid(error.to_string()))?;
            },
            b"showNegBubbles" => show_negative_bubbles = parse_bool_attr(element)?,
            b"sizeRepresents" => {
                let value = get_attr(element, b"val").unwrap_or_else(|| b"area".to_vec());
                size = BubbleSize::from_xml(&value).map_err(|_error| {
                    invalid_attribute("chart bubble size representation", &value)
                })?;
            },
            _ => {},
        }
        Ok(())
    })?;
    let mut group = BubbleTypeGroup::new();
    group.common = common;
    // ECMA-376 permits bubble3D at group level, but Microsoft Office rejects
    // that placement. Canonicalize its semantic effect onto each series,
    // which is the Office-compatible representation.
    if group_bubble_3d {
        for series in &mut group.common.series {
            series.bubble_3d = true;
        }
    }
    group.set_scale(scale);
    group.show_negative_bubbles = show_negative_bubbles;
    group.set_size(size);
    Ok(group)
}

pub(crate) fn parse_doughnut_chart<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<DoughnutTypeGroup> {
    let mut first_slice_angle = 0;
    let mut hole_size = 50;
    let common = parse_common_type_group(reader, b"doughnutChart", true, false, None, |element| {
        match element.local_name().as_ref() {
            b"firstSliceAng" => {
                first_slice_angle = match get_attr(element, b"val") {
                    Some(_) => required_u32_attr(element, "first-slice angle")?,
                    None => 0,
                };
                if first_slice_angle > 360 {
                    return Err(Error::Invalid(
                        "chart first-slice angle exceeds 360".to_string(),
                    ));
                }
            },
            b"holeSize" => {
                hole_size = match get_attr(element, b"val") {
                    Some(_) => bounded_percentage_u32_attr(element, "doughnut hole size", 1, 90)?,
                    None => 10,
                };
            },
            _ => {},
        }
        Ok(())
    })?;
    let mut group = DoughnutTypeGroup::new();
    group.common = common;
    group.first_slice_angle = first_slice_angle;
    group.hole_size = hole_size;
    Ok(group)
}

pub(crate) fn parse_line_3d_chart<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Line3DTypeGroup> {
    let mut grouping = BarGrouping::Standard;
    let mut gap_depth = None;
    let mut drop_lines = None;
    let common = parse_common_type_group(
        reader,
        b"line3DChart",
        true,
        true,
        Some(&mut drop_lines),
        |element| {
            match element.local_name().as_ref() {
                b"grouping" => grouping = parse_grouping(element)?,
                b"gapDepth" => {
                    gap_depth = Some(match get_attr(element, b"val") {
                        Some(_) => {
                            bounded_percentage_u32_attr(element, "line 3D gap depth", 0, 500)?
                        },
                        None => 150,
                    });
                },
                _ => {},
            }
            Ok(())
        },
    )?;
    let mut group = Line3DTypeGroup::new(grouping);
    group.common = common;
    group.gap_depth = gap_depth;
    group.drop_lines = drop_lines;
    Ok(group)
}

pub(crate) fn parse_pie_3d_chart<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Pie3DTypeGroup> {
    let mut group = Pie3DTypeGroup::new();
    group.common = parse_common_type_group(reader, b"pie3DChart", true, false, None, |_| Ok(()))?;
    Ok(group)
}

pub(crate) fn parse_of_pie_chart<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<OfPieTypeGroup> {
    let mut group = OfPieTypeGroup::new(OfPieType::Pie);
    let mut saw_of_pie_type = false;
    let mut saw_gap_width = false;
    let mut saw_split_type = false;
    let mut saw_split_position = false;
    let mut saw_custom_split = false;
    let mut saw_second_pie_size = false;
    let mut saw_data_labels = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"serLines" => {
                group
                    .series_lines
                    .push(parse_chart_lines(reader, b"serLines")?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"serLines" => {
                group.series_lines.push(Lines::new());
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut group.common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut group.common, element, true)?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                group.common.data_labels = Some(parse_data_labels(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                group.common.data_labels = Some(DataLabels::default());
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"custSplit" => {
                if saw_custom_split {
                    return Err(Error::Invalid(
                        "of-pie chart contains duplicate custom splits".into(),
                    ));
                }
                saw_custom_split = true;
                group.custom_split_points = Some(parse_custom_pie_split(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"custSplit" => {
                if saw_custom_split {
                    return Err(Error::Invalid(
                        "of-pie chart contains duplicate custom splits".into(),
                    ));
                }
                saw_custom_split = true;
                group.custom_split_points = Some(Vec::new());
            },
            Ok(Event::Start(ref element) | Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"ofPieType" => {
                        if saw_of_pie_type {
                            return Err(Error::Invalid(
                                "of-pie chart contains duplicate plot types".into(),
                            ));
                        }
                        saw_of_pie_type = true;
                        group.of_pie_type = match get_attr(element, b"val").as_deref() {
                            None | Some(b"pie") => OfPieType::Pie,
                            Some(b"bar") => OfPieType::Bar,
                            Some(value) => return Err(invalid_attribute("of-pie type", value)),
                        };
                    },
                    b"varyColors" => group.common.vary_colors = parse_bool_attr(element)?,
                    b"ser" => {
                        if let Some(series) = parse_series(reader)? {
                            group.common.series.push(series);
                        }
                    },
                    b"gapWidth" => {
                        if saw_gap_width {
                            return Err(Error::Invalid(
                                "of-pie chart contains duplicate gap widths".into(),
                            ));
                        }
                        saw_gap_width = true;
                        group.gap_width = Some(match get_attr(element, b"val") {
                            Some(_) => {
                                bounded_percentage_u32_attr(element, "of-pie gap width", 0, 500)?
                            },
                            None => 150,
                        });
                    },
                    b"splitType" => {
                        if saw_split_type {
                            return Err(Error::Invalid(
                                "of-pie chart contains duplicate split types".into(),
                            ));
                        }
                        saw_split_type = true;
                        group.split_type = Some(match get_attr(element, b"val").as_deref() {
                            None | Some(b"auto") => OfPieSplitType::Automatic,
                            Some(b"cust") => OfPieSplitType::Custom,
                            Some(b"percent") => OfPieSplitType::Percent,
                            Some(b"pos") => OfPieSplitType::Position,
                            Some(b"val") => OfPieSplitType::Value,
                            Some(value) => {
                                return Err(invalid_attribute("of-pie split type", value));
                            },
                        });
                    },
                    b"splitPos" => {
                        if saw_split_position {
                            return Err(Error::Invalid(
                                "of-pie chart contains duplicate split positions".into(),
                            ));
                        }
                        saw_split_position = true;
                        group.split_position =
                            Some(required_f64_attr(element, "of-pie split position")?);
                    },
                    b"secondPieSize" => {
                        if saw_second_pie_size {
                            return Err(Error::Invalid(
                                "of-pie chart contains duplicate secondary sizes".into(),
                            ));
                        }
                        saw_second_pie_size = true;
                        group.second_pie_size = Some(match get_attr(element, b"val") {
                            Some(_) => bounded_percentage_u32_attr(
                                element,
                                "of-pie secondary size",
                                5,
                                200,
                            )?,
                            None => 75,
                        });
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"ofPieChart" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated of-pie chart".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    if !saw_of_pie_type {
        return Err(Error::Invalid(
            "of-pie chart is missing its plot type".into(),
        ));
    }
    Ok(group)
}

pub(crate) fn parse_custom_pie_split<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Vec<u32>> {
    let mut points = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element) | Event::Empty(ref element))
                if element.local_name().as_ref() == b"secondPiePt" =>
            {
                points.push(required_u32_attr(element, "of-pie secondary point")?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"custSplit" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated of-pie custom split".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(points)
}

pub(crate) fn parse_up_down_bars<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<UpDownBars> {
    let mut bars = UpDownBars::default();
    let mut saw_gap_width = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"upBars" => {
                let lines = parse_chart_lines(reader, b"upBars")?;
                set_chart_lines(&mut bars.up_bars, lines, "chart up bars")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"upBars" => {
                set_chart_lines(&mut bars.up_bars, Lines::new(), "chart up bars")?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"downBars" => {
                let lines = parse_chart_lines(reader, b"downBars")?;
                set_chart_lines(&mut bars.down_bars, lines, "chart down bars")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"downBars" => {
                set_chart_lines(&mut bars.down_bars, Lines::new(), "chart down bars")?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if bars.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart up/down bars contain duplicate extension lists".into(),
                    ));
                }
                bars.extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(element, "chart up/down-bar extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if bars.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart up/down bars contain duplicate extension lists".into(),
                    ));
                }
                bars.extension_list = Some(ExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element) | Event::Empty(ref element))
                if element.local_name().as_ref() == b"gapWidth" =>
            {
                if saw_gap_width {
                    return Err(Error::Invalid(
                        "chart up/down bars contain duplicate gap widths".into(),
                    ));
                }
                saw_gap_width = true;
                bars.gap_width = Some(match get_attr(element, b"val") {
                    Some(_) => {
                        bounded_percentage_u32_attr(element, "chart up/down-bar gap width", 0, 500)?
                    },
                    None => 150,
                });
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"upDownBars" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart up/down bars".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(bars)
}

pub(crate) fn parse_radar_chart<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<RadarTypeGroup> {
    let mut style = RadarStyle::Standard;
    let common = parse_common_type_group(reader, b"radarChart", true, true, None, |element| {
        if element.local_name().as_ref() == b"radarStyle" {
            let value =
                get_attr(element, b"val").ok_or_else(|| missing_attribute("chart radar style"))?;
            style = match value.as_slice() {
                b"standard" => RadarStyle::Standard,
                b"filled" => RadarStyle::Filled,
                b"marker" => RadarStyle::Marker,
                _ => return Err(invalid_attribute("chart radar style", &value)),
            };
        }
        Ok(())
    })?;
    let mut group = RadarTypeGroup::new(style);
    group.common = common;
    Ok(group)
}

pub(crate) fn parse_stock_chart<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<StockTypeGroup> {
    let mut common = TypeGroupCommon::new();
    let mut saw_data_labels = false;
    let mut drop_lines = None;
    let mut high_low_lines = None;
    let mut up_down_bars = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"dropLines" => {
                let lines = parse_chart_lines(reader, b"dropLines")?;
                set_chart_lines(&mut drop_lines, lines, "stock chart drop lines")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"dropLines" => {
                set_chart_lines(&mut drop_lines, Lines::new(), "stock chart drop lines")?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"hiLowLines" => {
                let lines = parse_chart_lines(reader, b"hiLowLines")?;
                set_chart_lines(&mut high_low_lines, lines, "stock chart high/low lines")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"hiLowLines" => {
                set_chart_lines(
                    &mut high_low_lines,
                    Lines::new(),
                    "stock chart high/low lines",
                )?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"upDownBars" => {
                if up_down_bars.is_some() {
                    return Err(Error::Invalid(
                        "stock chart contains duplicate up/down bars".into(),
                    ));
                }
                up_down_bars = Some(parse_up_down_bars(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"upDownBars" => {
                set_empty_up_down_bars(&mut up_down_bars, "stock chart")?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(parse_data_labels(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(DataLabels::default());
            },
            Ok(Event::Start(ref element) | Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    b"axId" => common
                        .axis_ids
                        .push(required_u32_attr(element, "stock chart axis ID")?),
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"stockChart" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated stock chart".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let mut group = StockTypeGroup::new();
    group.common = common;
    group.drop_lines = drop_lines;
    group.high_low_lines = high_low_lines;
    group.up_down_bars = up_down_bars;
    Ok(group)
}

pub(crate) fn parse_surface_chart<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<SurfaceTypeGroup> {
    let (common, wireframe, band_formats) = parse_surface_type_group(reader, b"surfaceChart")?;
    let mut group = SurfaceTypeGroup::new();
    group.common = common;
    group.wireframe = wireframe;
    group.band_formats = band_formats;
    Ok(group)
}

pub(crate) fn parse_surface_3d_chart<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Surface3DTypeGroup> {
    let (common, wireframe, band_formats) = parse_surface_type_group(reader, b"surface3DChart")?;
    let mut group = Surface3DTypeGroup::new();
    group.common = common;
    group.wireframe = wireframe;
    group.band_formats = band_formats;
    Ok(group)
}

pub(crate) fn parse_surface_type_group<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
    end_name: &[u8],
) -> Result<(TypeGroupCommon, bool, Option<Vec<BandFormat>>)> {
    let mut common = TypeGroupCommon::new();
    let mut wireframe = false;
    let mut band_formats = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"bandFmts" => {
                if band_formats.is_some() {
                    return Err(Error::Invalid(
                        "surface chart contains duplicate band-format collections".into(),
                    ));
                }
                band_formats = Some(parse_surface_band_formats(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"bandFmts" => {
                if band_formats.is_some() {
                    return Err(Error::Invalid(
                        "surface chart contains duplicate band-format collections".into(),
                    ));
                }
                band_formats = Some(Vec::new());
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"ser" => {
                if let Some(series) = parse_series(reader)? {
                    common.series.push(series);
                }
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"ser" => {
                return Err(Error::Invalid(
                    "surface chart contains an empty series".into(),
                ));
            },
            Ok(Event::Start(ref element) | Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"wireframe" => wireframe = parse_bool_attr(element)?,
                    b"axId" => common
                        .axis_ids
                        .push(required_u32_attr(element, "surface chart axis ID")?),
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == end_name => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated surface chart".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok((common, wireframe, band_formats))
}

pub(crate) fn parse_surface_band_formats<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Vec<BandFormat>> {
    let mut formats = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"bandFmt" => {
                let format = parse_surface_band_format(reader)?;
                if formats
                    .iter()
                    .any(|existing: &BandFormat| existing.index == format.index)
                {
                    return Err(Error::Invalid(format!(
                        "surface chart contains duplicate band index {}",
                        format.index
                    )));
                }
                formats.push(format);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"bandFmt" => {
                return Err(Error::Invalid(
                    "surface chart band format is missing its index".into(),
                ));
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"bandFmts" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid(
                    "unterminated surface chart band formats".into(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(formats)
}

pub(crate) fn parse_surface_band_format<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<BandFormat> {
    let mut index = None;
    let mut shape_properties = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "surface chart band format contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(element, "surface chart band shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "surface chart band format contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element) | Event::Empty(ref element))
                if element.local_name().as_ref() == b"idx" =>
            {
                if index.is_some() {
                    return Err(Error::Invalid(
                        "surface chart band format contains duplicate indexes".into(),
                    ));
                }
                index = Some(required_u32_attr(element, "surface chart band index")?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"bandFmt" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid(
                    "unterminated surface chart band format".into(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    let mut format =
        BandFormat::new(index.ok_or_else(|| missing_attribute("surface chart band index"))?);
    format.shape_properties = shape_properties;
    Ok(format)
}

pub(crate) fn parse_bar_chart<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Option<BarTypeGroup>> {
    let mut direction = BarDirection::Column;
    let mut grouping = BarGrouping::Clustered;
    let mut common = TypeGroupCommon::new();
    let mut saw_data_labels = false;
    let mut gap_width = None;
    let mut overlap = None;
    let mut series_lines = Vec::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"serLines" => {
                series_lines.push(parse_chart_lines(reader, b"serLines")?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"serLines" => {
                series_lines.push(Lines::new());
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(parse_data_labels(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(DataLabels::default());
            },
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"barDir" => {
                        if let Some(val) = get_attr(e, b"val") {
                            direction = if val.as_slice() == b"bar" {
                                BarDirection::Bar
                            } else if val.as_slice() == b"col" {
                                BarDirection::Column
                            } else {
                                return Err(invalid_attribute("chart bar direction", &val));
                            };
                        }
                    },
                    b"grouping" => {
                        grouping = parse_grouping(e)?;
                    },
                    b"varyColors" => {
                        common.vary_colors = parse_bool_attr(e)?;
                    },
                    b"axId" => common
                        .axis_ids
                        .push(required_u32_attr(e, "bar chart axis ID")?),
                    b"ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    b"gapWidth" => {
                        gap_width = Some(match get_attr(e, b"val") {
                            Some(_) => bounded_percentage_u32_attr(e, "chart gap width", 0, 500)?,
                            None => 150,
                        });
                    },
                    b"overlap" => {
                        overlap = Some(match get_attr(e, b"val") {
                            Some(_) => bounded_percentage_i32_attr(e, "chart overlap", -100, 100)?,
                            None => 0,
                        });
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"barChart" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart element".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let mut group = BarTypeGroup::new(direction, grouping);
    group.common = common;
    group.gap_width = gap_width;
    group.overlap = overlap;
    group.series_lines = series_lines;
    Ok(Some(group))
}

pub(crate) fn parse_bar_3d_chart<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Option<Bar3DTypeGroup>> {
    let mut direction = BarDirection::Column;
    let mut grouping = BarGrouping::Clustered;
    let mut common = TypeGroupCommon::new();
    let mut saw_data_labels = false;
    let mut gap_width = None;
    let mut gap_depth = None;
    let mut shape = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(parse_data_labels(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(DataLabels::default());
            },
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"barDir" => {
                        if let Some(val) = get_attr(e, b"val") {
                            direction = if val.as_slice() == b"bar" {
                                BarDirection::Bar
                            } else if val.as_slice() == b"col" {
                                BarDirection::Column
                            } else {
                                return Err(invalid_attribute("chart bar direction", &val));
                            };
                        }
                    },
                    b"grouping" => {
                        grouping = parse_grouping(e)?;
                    },
                    b"varyColors" => {
                        common.vary_colors = parse_bool_attr(e)?;
                    },
                    b"axId" => common
                        .axis_ids
                        .push(required_u32_attr(e, "3D bar chart axis ID")?),
                    b"ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    b"gapWidth" => {
                        gap_width = Some(match get_attr(e, b"val") {
                            Some(_) => bounded_percentage_u32_attr(e, "chart gap width", 0, 500)?,
                            None => 150,
                        });
                    },
                    b"gapDepth" => {
                        gap_depth = Some(match get_attr(e, b"val") {
                            Some(_) => bounded_percentage_u32_attr(e, "chart gap depth", 0, 500)?,
                            None => 150,
                        });
                    },
                    b"shape" => {
                        shape = Some(parse_bar_shape(e, "chart bar shape")?);
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"bar3DChart" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart element".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let mut group = Bar3DTypeGroup::new(direction, grouping);
    group.common = common;
    group.gap_width = gap_width;
    group.gap_depth = gap_depth;
    group.shape = shape;
    Ok(Some(group))
}

pub(crate) fn parse_line_chart<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Option<LineTypeGroup>> {
    let mut grouping = BarGrouping::Standard;
    let mut common = TypeGroupCommon::new();
    let mut saw_data_labels = false;
    let mut marker = true;
    let mut smooth = false;
    let mut drop_lines = None;
    let mut high_low_lines = None;
    let mut up_down_bars = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"dropLines" => {
                let lines = parse_chart_lines(reader, b"dropLines")?;
                set_chart_lines(&mut drop_lines, lines, "line chart drop lines")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"dropLines" => {
                set_chart_lines(&mut drop_lines, Lines::new(), "line chart drop lines")?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"hiLowLines" => {
                let lines = parse_chart_lines(reader, b"hiLowLines")?;
                set_chart_lines(&mut high_low_lines, lines, "line chart high/low lines")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"hiLowLines" => {
                set_chart_lines(
                    &mut high_low_lines,
                    Lines::new(),
                    "line chart high/low lines",
                )?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"upDownBars" => {
                if up_down_bars.is_some() {
                    return Err(Error::Invalid(
                        "line chart contains duplicate up/down bars".into(),
                    ));
                }
                up_down_bars = Some(parse_up_down_bars(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"upDownBars" => {
                set_empty_up_down_bars(&mut up_down_bars, "line chart")?;
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(parse_data_labels(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(DataLabels::default());
            },
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"grouping" => {
                        grouping = parse_grouping(e)?;
                    },
                    b"varyColors" => {
                        common.vary_colors = parse_bool_attr(e)?;
                    },
                    b"axId" => common
                        .axis_ids
                        .push(required_u32_attr(e, "line chart axis ID")?),
                    b"ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    b"marker" => marker = parse_bool_attr(e)?,
                    b"smooth" => smooth = parse_bool_attr(e)?,
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"lineChart" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart element".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let mut group = LineTypeGroup::new(grouping);
    group.common = common;
    group.marker = marker;
    group.smooth = smooth;
    group.drop_lines = drop_lines;
    group.high_low_lines = high_low_lines;
    group.up_down_bars = up_down_bars;
    Ok(Some(group))
}

pub(crate) fn parse_pie_chart<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Option<PieTypeGroup>> {
    let mut common = TypeGroupCommon::new();
    let mut saw_data_labels = false;
    let mut first_slice_angle = 0;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(parse_data_labels(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(DataLabels::default());
            },
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"varyColors" => {
                        common.vary_colors = parse_bool_attr(e)?;
                    },
                    b"ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    b"firstSliceAng" => {
                        first_slice_angle = match get_attr(e, b"val") {
                            Some(_) => bounded_u32_attr(e, "chart first-slice angle", 0, 360)?,
                            None => 0,
                        };
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"pieChart" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart element".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let mut group = PieTypeGroup::new();
    group.common = common;
    group.first_slice_angle = first_slice_angle;
    Ok(Some(group))
}

pub(crate) fn parse_area_chart<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Option<AreaTypeGroup>> {
    let mut grouping = BarGrouping::Standard;
    let mut common = TypeGroupCommon::new();
    let mut saw_data_labels = false;
    let mut drop_lines = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"dropLines" => {
                let lines = parse_chart_lines(reader, b"dropLines")?;
                set_chart_lines(&mut drop_lines, lines, "area chart drop lines")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"dropLines" => {
                set_chart_lines(&mut drop_lines, Lines::new(), "area chart drop lines")?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(parse_data_labels(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(DataLabels::default());
            },
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"grouping" => {
                        grouping = parse_grouping(e)?;
                    },
                    b"varyColors" => {
                        common.vary_colors = parse_bool_attr(e)?;
                    },
                    b"axId" => common
                        .axis_ids
                        .push(required_u32_attr(e, "area chart axis ID")?),
                    b"ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"areaChart" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart element".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let mut group = AreaTypeGroup::new(grouping);
    group.common = common;
    group.drop_lines = drop_lines;
    Ok(Some(group))
}

pub(crate) fn parse_scatter_chart<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Option<ScatterTypeGroup>> {
    let mut style = ScatterStyle::LineMarker;
    let mut common = TypeGroupCommon::new();
    let mut saw_data_labels = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(parse_data_labels(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(DataLabels::default());
            },
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"scatterStyle" => {
                        if let Some(val) = get_attr(e, b"val") {
                            style = match val.as_slice() {
                                b"line" => ScatterStyle::Line,
                                b"marker" => ScatterStyle::Marker,
                                b"none" => ScatterStyle::None,
                                b"smooth" => ScatterStyle::Smooth,
                                b"smoothMarker" => ScatterStyle::SmoothMarker,
                                b"lineMarker" => ScatterStyle::LineMarker,
                                _ => {
                                    return Err(invalid_attribute("chart scatter style", &val));
                                },
                            };
                        }
                    },
                    b"varyColors" => {
                        common.vary_colors = parse_bool_attr(e)?;
                    },
                    b"axId" => common
                        .axis_ids
                        .push(required_u32_attr(e, "scatter chart axis ID")?),
                    b"ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"scatterChart" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart element".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let mut group = ScatterTypeGroup::new(style);
    group.common = common;
    Ok(Some(group))
}
