use super::{
    BarShape, BufRead, BytesStart, ChartXmlReader, DataLabel, DataLabels, DataPoint, DataSourceRef,
    Error, ErrorBar, ErrorBarDirection, ErrorBarType, ErrorBarValueType, Event, ExtensionList,
    IGNORED_NAMESPACE_ELEMENT, Layout, Lines, Marker, NumberFormat, NumericData, PictureOptions,
    Result, RichText, Series, ShapeProperties, StringData, TextProperties, TitleText, Trendline,
    TrendlineType, bounded_u32_attr, decode_xml_reference, get_attr, invalid_attribute,
    missing_attribute, parse_bool_attr, parse_chart_lines, parse_data_label_position, parse_layout,
    parse_marker_style, parse_number_format, parse_picture_options, required_enum_attr,
    required_f64_attr, required_nonnegative_f64_attr, required_u32_attr, set_chart_lines,
};

pub(crate) fn parse_series<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Option<Series>> {
    let mut series = Series::new(0);
    let mut saw_index = false;
    let mut saw_order = false;
    let mut saw_marker = false;
    let mut saw_data_labels = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if series.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart series contains duplicate shape properties".into(),
                    ));
                }
                series.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(e, "chart series shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if series.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart series contains duplicate shape properties".into(),
                    ));
                }
                series.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(e)?,
                )?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pictureOptions" => {
                if series.picture_options.is_some() {
                    return Err(Error::Invalid(
                        "chart series contains duplicate picture options".into(),
                    ));
                }
                series.picture_options = Some(parse_picture_options(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"pictureOptions" => {
                if series.picture_options.is_some() {
                    return Err(Error::Invalid(
                        "chart series contains duplicate picture options".into(),
                    ));
                }
                series.picture_options = Some(PictureOptions::default());
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if series.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart series contains duplicate extension lists".into(),
                    ));
                }
                series.extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(e, "chart series extension list")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if series.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart series contains duplicate extension lists".into(),
                    ));
                }
                series.extension_list =
                    Some(ExtensionList::from_xml(reader.capture_empty_fragment(e)?)?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"marker" => {
                if saw_marker {
                    return Err(Error::Invalid(
                        "chart series has duplicate marker".to_string(),
                    ));
                }
                saw_marker = true;
                series.marker_present = true;
                let marker = parse_series_marker(reader)?;
                series.marker_symbol = marker.symbol;
                series.marker_size = marker.size;
                series.marker_shape_properties = marker.shape_properties;
                series.marker_extension_list = marker.extension_list;
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"marker" => {
                if saw_marker {
                    return Err(Error::Invalid(
                        "chart series has duplicate marker".to_string(),
                    ));
                }
                saw_marker = true;
                series.marker_present = true;
            },
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"idx" => {
                        if saw_index {
                            return Err(Error::Invalid(
                                "chart series has duplicate index".to_string(),
                            ));
                        }
                        saw_index = true;
                        series.index = required_u32_attr(e, "chart series index")?;
                    },
                    b"order" => {
                        if saw_order {
                            return Err(Error::Invalid(
                                "chart series has duplicate order".to_string(),
                            ));
                        }
                        saw_order = true;
                        series.order = required_u32_attr(e, "chart series order")?;
                    },
                    b"tx" => {
                        series.title = parse_series_title(reader)?;
                    },
                    b"invertIfNegative" => {
                        series.invert_if_negative = parse_bool_attr(e)?;
                    },
                    b"dPt" => {
                        let point = parse_data_point(reader)?;
                        if series
                            .data_points
                            .iter()
                            .any(|existing| existing.index == point.index)
                        {
                            return Err(Error::Invalid(format!(
                                "chart series has duplicate data-point index {}",
                                point.index
                            )));
                        }
                        series.data_points.push(point);
                    },
                    b"dLbls" => {
                        if saw_data_labels {
                            return Err(Error::Invalid(
                                "chart series has duplicate data-label settings".to_string(),
                            ));
                        }
                        saw_data_labels = true;
                        series.data_labels = Some(parse_data_labels(reader)?);
                    },
                    b"trendline" => series.trendlines.push(parse_trendline(reader)?),
                    b"errBars" => {
                        let error_bar = parse_error_bar(reader)?;
                        if series
                            .error_bars
                            .iter()
                            .any(|existing| existing.direction == error_bar.direction)
                        {
                            return Err(Error::Invalid(
                                "chart series has duplicate error-bar direction".to_string(),
                            ));
                        }
                        series.error_bars.push(error_bar);
                    },
                    b"cat" => {
                        series.categories = parse_string_data(reader)?;
                    },
                    b"val" => {
                        series.values = parse_numeric_data(reader)?;
                    },
                    b"xVal" => {
                        series.x_values = parse_numeric_data(reader)?;
                    },
                    b"yVal" => {
                        series.y_values = parse_numeric_data(reader)?;
                    },
                    b"bubbleSize" => {
                        series.bubble_sizes = parse_numeric_data(reader)?;
                    },
                    b"explosion" => {
                        series.explosion = Some(required_u32_attr(e, "series explosion")?);
                    },
                    b"bubble3D" => {
                        series.bubble_3d = parse_bool_attr(e)?;
                    },
                    b"smooth" => {
                        series.smooth = parse_bool_attr(e)?;
                    },
                    b"shape" => {
                        if series.bar_shape.is_some() {
                            return Err(Error::Invalid(
                                "chart series contains duplicate bar shapes".into(),
                            ));
                        }
                        series.bar_shape = Some(parse_bar_shape(e, "chart series bar shape")?);
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"ser" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart element".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    if !saw_index || !saw_order {
        return Err(Error::Invalid(
            "chart series requires both index and order".to_string(),
        ));
    }
    Ok(Some(series))
}

pub(crate) fn parse_bar_shape(element: &BytesStart<'_>, description: &str) -> Result<BarShape> {
    let value = get_attr(element, b"val").unwrap_or_else(|| b"box".to_vec());
    match value.as_slice() {
        b"box" => Ok(BarShape::Box),
        b"cone" => Ok(BarShape::Cone),
        b"coneToMax" => Ok(BarShape::ConeToMax),
        b"cylinder" => Ok(BarShape::Cylinder),
        b"pyramid" => Ok(BarShape::Pyramid),
        b"pyramidToMax" => Ok(BarShape::PyramidToMax),
        _ => Err(invalid_attribute(description, &value)),
    }
}

pub(crate) fn parse_series_marker<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Marker> {
    let mut symbol = None;
    let mut size = None;
    let mut shape_properties = None;
    let mut extension_list = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart marker has duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart marker shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart marker has duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart marker has duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(element, "chart marker extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart marker has duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element) | Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"symbol" => {
                        if symbol.is_some() {
                            return Err(Error::Invalid(
                                "chart marker has duplicate symbol".to_string(),
                            ));
                        }
                        symbol = Some(parse_marker_style(element)?);
                    },
                    b"size" => {
                        if size.is_some() {
                            return Err(Error::Invalid(
                                "chart marker has duplicate size".to_string(),
                            ));
                        }
                        size = Some(bounded_u32_attr(element, "chart marker size", 2, 72)?);
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"marker" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid(
                    "unterminated chart series marker".to_string(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(Marker {
        symbol,
        size,
        shape_properties,
        extension_list,
    })
}

pub(crate) fn parse_data_point<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<DataPoint> {
    let mut index = None;
    let mut explosion = None;
    let mut marker = None;
    let mut invert_if_negative = false;
    let mut bubble_3d = None;
    let mut shape_properties = None;
    let mut picture_options = None;
    let mut extension_list = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"marker" => {
                if marker.is_some() {
                    return Err(Error::Invalid(
                        "chart data point contains duplicate markers".into(),
                    ));
                }
                marker = Some(parse_series_marker(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"marker" => {
                if marker.is_some() {
                    return Err(Error::Invalid(
                        "chart data point contains duplicate markers".into(),
                    ));
                }
                marker = Some(Marker::new());
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart data point contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart data-point shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart data point contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"pictureOptions" => {
                if picture_options.is_some() {
                    return Err(Error::Invalid(
                        "chart data point contains duplicate picture options".into(),
                    ));
                }
                picture_options = Some(parse_picture_options(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"pictureOptions" => {
                if picture_options.is_some() {
                    return Err(Error::Invalid(
                        "chart data point contains duplicate picture options".into(),
                    ));
                }
                picture_options = Some(PictureOptions::default());
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart data point contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(element, "chart data-point extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart data point contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element) | Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"idx" => {
                        index = Some(required_u32_attr(element, "chart data-point index")?);
                    },
                    b"explosion" => {
                        explosion = Some(required_u32_attr(element, "chart data-point explosion")?);
                    },
                    b"invertIfNegative" => invert_if_negative = parse_bool_attr(element)?,
                    b"bubble3D" => bubble_3d = Some(parse_bool_attr(element)?),
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"dPt" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart data point".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    let mut point =
        DataPoint::new(index.ok_or_else(|| missing_attribute("chart data-point index"))?);
    point.explosion = explosion;
    if let Some(marker) = marker {
        point.marker_present = true;
        point.marker_size = marker.size;
        point.marker_symbol = marker.symbol;
        point.marker_shape_properties = marker.shape_properties;
        point.marker_extension_list = marker.extension_list;
    }
    point.invert_if_negative = invert_if_negative;
    point.bubble_3d = bubble_3d;
    point.shape_properties = shape_properties;
    point.picture_options = picture_options;
    point.extension_list = extension_list;
    Ok(point)
}

pub(crate) fn parse_data_labels<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<DataLabels> {
    let mut labels = DataLabels::new();
    let mut saw_number_format = false;
    let mut saw_delete = false;
    let mut saw_shared_settings = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"dLbl" => {
                let label = parse_data_label(reader)?;
                if labels
                    .labels
                    .iter()
                    .any(|existing| existing.index == label.index)
                {
                    return Err(Error::Invalid(format!(
                        "chart data labels contain duplicate point index {}",
                        label.index
                    )));
                }
                labels.labels.push(label);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"dLbl" => {
                return Err(Error::Invalid(
                    "chart point data label is missing its index".into(),
                ));
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if labels.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart data labels contain duplicate shape properties".into(),
                    ));
                }
                labels.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart data-label shape properties")?,
                )?);
                saw_shared_settings = true;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if labels.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart data labels contain duplicate shape properties".into(),
                    ));
                }
                labels.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
                saw_shared_settings = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if labels.text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart data labels contain duplicate text properties".into(),
                    ));
                }
                labels.text_properties = Some(TextProperties::from_xml(
                    reader.capture_fragment(element, "chart data-label text properties")?,
                )?);
                saw_shared_settings = true;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if labels.text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart data labels contain duplicate text properties".into(),
                    ));
                }
                labels.text_properties = Some(TextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
                saw_shared_settings = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"leaderLines" => {
                set_chart_lines(
                    &mut labels.leader_lines,
                    parse_chart_lines(reader, b"leaderLines")?,
                    "chart data-label leader lines",
                )?;
                saw_shared_settings = true;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"leaderLines" => {
                set_chart_lines(
                    &mut labels.leader_lines,
                    Lines::new(),
                    "chart data-label leader lines",
                )?;
                saw_shared_settings = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if labels.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart data labels contain duplicate extension lists".into(),
                    ));
                }
                labels.extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(element, "chart data-label extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if labels.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart data labels contain duplicate extension lists".into(),
                    ));
                }
                labels.extension_list = Some(ExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"separator" => {
                saw_shared_settings = true;
                labels.separator = Some(parse_text_element(reader, b"separator")?);
            },
            Ok(Event::Start(ref element) | Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"delete" => {
                        if saw_delete {
                            return Err(Error::Invalid(
                                "chart data labels contain duplicate delete flags".into(),
                            ));
                        }
                        labels.deleted = parse_bool_attr(element)?;
                        saw_delete = true;
                    },
                    b"numFmt" => {
                        if saw_number_format {
                            return Err(Error::Invalid(
                                "chart data labels contain duplicate number formats".into(),
                            ));
                        }
                        labels.number_format = Some(parse_number_format(
                            element,
                            reader.decoder(),
                            "chart data-label",
                        )?);
                        saw_number_format = true;
                        saw_shared_settings = true;
                    },
                    b"dLblPos" => {
                        labels.position = Some(parse_data_label_position(element)?);
                        saw_shared_settings = true;
                    },
                    b"showLegendKey" => {
                        labels.show_legend_key = parse_bool_attr(element)?;
                        saw_shared_settings = true;
                    },
                    b"showVal" => {
                        labels.show_value = parse_bool_attr(element)?;
                        saw_shared_settings = true;
                    },
                    b"showCatName" => {
                        labels.show_category_name = parse_bool_attr(element)?;
                        saw_shared_settings = true;
                    },
                    b"showSerName" => {
                        labels.show_series_name = parse_bool_attr(element)?;
                        saw_shared_settings = true;
                    },
                    b"showPercent" => {
                        labels.show_percent = parse_bool_attr(element)?;
                        saw_shared_settings = true;
                    },
                    b"showBubbleSize" => {
                        labels.show_bubble_size = parse_bool_attr(element)?;
                        saw_shared_settings = true;
                    },
                    b"showLeaderLines" => {
                        labels.show_leader_lines = parse_bool_attr(element)?;
                        saw_shared_settings = true;
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"dLbls" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart data labels".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    if saw_delete && saw_shared_settings {
        return Err(Error::Invalid(
            "chart data labels mix deletion with shared settings".into(),
        ));
    }
    Ok(labels)
}

pub(crate) fn parse_data_label<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<DataLabel> {
    let mut label = DataLabel::new(0);
    let mut saw_index = false;
    let mut saw_delete = false;
    let mut saw_settings = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"layout" => {
                if label.layout.is_some() {
                    return Err(Error::Invalid(
                        "chart point data label contains duplicate layouts".into(),
                    ));
                }
                label.layout = Some(parse_layout(reader)?);
                saw_settings = true;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"layout" => {
                label.layout = Some(match label.layout {
                    None => Layout::new(),
                    Some(_) => {
                        return Err(Error::Invalid(
                            "chart point data label contains duplicate layouts".into(),
                        ));
                    },
                });
                saw_settings = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"tx" => {
                if label.text.is_some() {
                    return Err(Error::Invalid(
                        "chart point data label contains duplicate text".into(),
                    ));
                }
                label.text = parse_label_text(reader)?;
                saw_settings = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if label.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart point data label contains duplicate shape properties".into(),
                    ));
                }
                label.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart point data-label shape properties")?,
                )?);
                saw_settings = true;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if label.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart point data label contains duplicate shape properties".into(),
                    ));
                }
                label.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
                saw_settings = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if label.text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart point data label contains duplicate text properties".into(),
                    ));
                }
                label.text_properties = Some(TextProperties::from_xml(
                    reader.capture_fragment(element, "chart point data-label text properties")?,
                )?);
                saw_settings = true;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if label.text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart point data label contains duplicate text properties".into(),
                    ));
                }
                label.text_properties = Some(TextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
                saw_settings = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if label.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart point data label contains duplicate extension lists".into(),
                    ));
                }
                label.extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(element, "chart point data-label extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if label.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart point data label contains duplicate extension lists".into(),
                    ));
                }
                label.extension_list = Some(ExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"separator" => {
                label.separator = Some(parse_text_element(reader, b"separator")?);
                saw_settings = true;
            },
            Ok(Event::Start(ref element) | Event::Empty(ref element)) => {
                let is_setting = match element.local_name().as_ref() {
                    b"idx" => {
                        if saw_index {
                            return Err(Error::Invalid(
                                "chart point data label contains duplicate indexes".into(),
                            ));
                        }
                        label.index = required_u32_attr(element, "chart point data-label index")?;
                        saw_index = true;
                        false
                    },
                    b"delete" => {
                        if saw_delete {
                            return Err(Error::Invalid(
                                "chart point data label contains duplicate delete flags".into(),
                            ));
                        }
                        label.deleted = parse_bool_attr(element)?;
                        saw_delete = true;
                        false
                    },
                    b"numFmt" => {
                        if label.number_format.is_some() {
                            return Err(Error::Invalid(
                                "chart point data label contains duplicate number formats".into(),
                            ));
                        }
                        label.number_format = Some(parse_number_format(
                            element,
                            reader.decoder(),
                            "chart point data-label",
                        )?);
                        true
                    },
                    b"dLblPos" => {
                        label.position = Some(parse_data_label_position(element)?);
                        true
                    },
                    b"showLegendKey" => {
                        label.show_legend_key = parse_bool_attr(element)?;
                        true
                    },
                    b"showVal" => {
                        label.show_value = parse_bool_attr(element)?;
                        true
                    },
                    b"showCatName" => {
                        label.show_category_name = parse_bool_attr(element)?;
                        true
                    },
                    b"showSerName" => {
                        label.show_series_name = parse_bool_attr(element)?;
                        true
                    },
                    b"showPercent" => {
                        label.show_percent = parse_bool_attr(element)?;
                        true
                    },
                    b"showBubbleSize" => {
                        label.show_bubble_size = parse_bool_attr(element)?;
                        true
                    },
                    _ => false,
                };
                saw_settings |= is_setting;
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"dLbl" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart point data label".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    if !saw_index {
        return Err(missing_attribute("chart point data-label index"));
    }
    if saw_delete && saw_settings {
        return Err(Error::Invalid(
            "chart point data label mixes deletion with label settings".into(),
        ));
    }
    Ok(label)
}

pub(crate) fn parse_label_text<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Option<TitleText>> {
    let mut text = String::new();
    let mut formula = None;
    let mut saw_text = false;
    let mut in_text = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"f" => {
                if formula.is_some() {
                    return Err(Error::Invalid(
                        "chart data label contains duplicate formula references".into(),
                    ));
                }
                formula = Some(parse_text_element(reader, b"f")?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"t" => {
                saw_text = true;
                in_text = true;
            },
            Ok(Event::Text(value)) if in_text => text.push_str(
                &value
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?,
            ),
            Ok(Event::CData(value)) if in_text => text.push_str(
                &value
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?,
            ),
            Ok(Event::GeneralRef(reference)) if in_text => {
                text.push_str(&decode_xml_reference(&reference)?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"t" => {
                in_text = false;
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"tx" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart data-label text".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    if let Some(formula) = formula {
        if saw_text {
            return Err(Error::Invalid(
                "chart data label mixes a formula reference with literal text".into(),
            ));
        }
        Ok(Some(TitleText::Reference(DataSourceRef::new(formula))))
    } else if saw_text {
        Ok(Some(TitleText::Literal(RichText::new(text))))
    } else {
        Ok(None)
    }
}

pub(crate) fn parse_trendline<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Trendline> {
    let mut trendline = Trendline::linear();
    let mut saw_type = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e))
                if matches!(e.local_name().as_ref(), b"name" | b"trendlineName") =>
            {
                if trendline.name.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline contains duplicate names".into(),
                    ));
                }
                trendline.name = Some(parse_text_element(reader, e.local_name().as_ref())?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if trendline.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline contains duplicate shape properties".into(),
                    ));
                }
                trendline.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(e, "chart trendline shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if trendline.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline contains duplicate shape properties".into(),
                    ));
                }
                trendline.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(e)?,
                )?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if trendline.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline contains duplicate extension lists".into(),
                    ));
                }
                trendline.extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(e, "chart trendline extension list")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if trendline.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline contains duplicate extension lists".into(),
                    ));
                }
                trendline.extension_list =
                    Some(ExtensionList::from_xml(reader.capture_empty_fragment(e)?)?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"trendlineLbl" => {
                if trendline.show_label {
                    return Err(Error::Invalid(
                        "chart trendline contains duplicate labels".into(),
                    ));
                }
                trendline.show_label = true;
                let label = parse_trendline_label(reader)?;
                trendline.label = label.text;
                trendline.label_layout = label.layout;
                trendline.label_number_format = label.number_format;
                trendline.label_shape_properties = label.shape_properties;
                trendline.label_text_properties = label.text_properties;
                trendline.label_extension_list = label.extension_list;
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"trendlineLbl" => {
                if trendline.show_label {
                    return Err(Error::Invalid(
                        "chart trendline contains duplicate labels".into(),
                    ));
                }
                trendline.show_label = true;
            },
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"trendlineType" => {
                    let value = get_attr(e, b"val")
                        .ok_or_else(|| missing_attribute("chart trendline type"))?;
                    trendline.trendline_type = match value.as_slice() {
                        b"exp" => TrendlineType::Exponential,
                        b"linear" => TrendlineType::Linear,
                        b"log" => TrendlineType::Logarithmic,
                        b"movingAvg" => TrendlineType::MovingAverage,
                        b"poly" => TrendlineType::Polynomial,
                        b"power" => TrendlineType::Power,
                        _ => return Err(invalid_attribute("chart trendline type", &value)),
                    };
                    saw_type = true;
                },
                b"order" => trendline.order = Some(bounded_u32_attr(e, "trendline order", 2, 6)?),
                b"period" => {
                    trendline.period = Some(bounded_u32_attr(e, "trendline period", 2, 255)?);
                },
                b"forward" => {
                    trendline.forward =
                        Some(required_nonnegative_f64_attr(e, "trendline forward")?);
                },
                b"backward" => {
                    trendline.backward =
                        Some(required_nonnegative_f64_attr(e, "trendline backward")?);
                },
                b"intercept" => {
                    trendline.intercept = Some(required_f64_attr(e, "trendline intercept")?);
                },
                b"dispEq" => trendline.display_equation = parse_bool_attr(e)?,
                b"dispRSqr" => trendline.display_r_squared = parse_bool_attr(e)?,
                _ => {},
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"trendline" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart trendline".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    if !saw_type {
        return Err(missing_attribute("chart trendline type"));
    }
    if matches!(trendline.trendline_type, TrendlineType::Polynomial) && trendline.order.is_none() {
        return Err(missing_attribute("polynomial trendline order"));
    }
    if matches!(trendline.trendline_type, TrendlineType::MovingAverage)
        && trendline.period.is_none()
    {
        return Err(missing_attribute("moving-average trendline period"));
    }
    if !matches!(trendline.trendline_type, TrendlineType::Polynomial) && trendline.order.is_some() {
        return Err(Error::Invalid(
            "only polynomial trendlines can specify an order".to_string(),
        ));
    }
    if !matches!(trendline.trendline_type, TrendlineType::MovingAverage)
        && trendline.period.is_some()
    {
        return Err(Error::Invalid(
            "only moving-average trendlines can specify a period".to_string(),
        ));
    }
    Ok(trendline)
}

#[derive(Default)]
pub(crate) struct ParsedTrendlineLabel {
    text: Option<TitleText>,
    pub(super) layout: Option<Layout>,
    number_format: Option<NumberFormat>,
    pub(super) shape_properties: Option<ShapeProperties>,
    pub(super) text_properties: Option<TextProperties>,
    pub(super) extension_list: Option<ExtensionList>,
}

pub(crate) fn parse_trendline_label<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<ParsedTrendlineLabel> {
    let mut text = None;
    let mut saw_text = false;
    let mut layout = None;
    let mut number_format = None;
    let mut shape_properties = None;
    let mut text_properties = None;
    let mut extension_list = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"layout" => {
                if layout.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline label contains duplicate layouts".into(),
                    ));
                }
                layout = Some(parse_layout(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"layout" => {
                layout = Some(match layout {
                    None => Layout::new(),
                    Some(_) => {
                        return Err(Error::Invalid(
                            "chart trendline label contains duplicate layouts".into(),
                        ));
                    },
                });
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"tx" => {
                if saw_text {
                    return Err(Error::Invalid(
                        "chart trendline label contains duplicate text".into(),
                    ));
                }
                saw_text = true;
                text = parse_label_text(reader)?;
            },
            Ok(Event::Start(ref element) | Event::Empty(ref element))
                if element.local_name().as_ref() == b"numFmt" =>
            {
                if number_format.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline label contains duplicate number formats".into(),
                    ));
                }
                number_format = Some(parse_number_format(
                    element,
                    reader.decoder(),
                    "chart trendline-label",
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline label contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart trendline-label shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline label contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline label contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(TextProperties::from_xml(
                    reader.capture_fragment(element, "chart trendline-label text properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline label contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(TextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline label contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(element, "chart trendline-label extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline label contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"trendlineLbl" => {
                break;
            },
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart trendline label".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(ParsedTrendlineLabel {
        text,
        layout,
        number_format,
        shape_properties,
        text_properties,
        extension_list,
    })
}

pub(crate) fn parse_error_bar<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<ErrorBar> {
    let mut direction = None;
    let mut error_type = None;
    let mut value_type = None;
    let mut value = None;
    let mut plus_values = None;
    let mut minus_values = None;
    let mut no_end_cap = false;
    let mut shape_properties = None;
    let mut extension_list = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart error bars contain duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(e, "chart error-bar shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart error bars contain duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(e)?,
                )?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart error bars contain duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(e, "chart error-bar extension list")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart error bars contain duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(reader.capture_empty_fragment(e)?)?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"plus" => {
                plus_values = parse_numeric_data(reader)?;
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"minus" => {
                minus_values = parse_numeric_data(reader)?;
            },
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"errDir" => {
                    direction = Some(
                        match required_enum_attr(e, "error-bar direction")?.as_str() {
                            "x" => ErrorBarDirection::X,
                            "y" => ErrorBarDirection::Y,
                            value => {
                                return Err(invalid_attribute(
                                    "error-bar direction",
                                    value.as_bytes(),
                                ));
                            },
                        },
                    );
                },
                b"errBarType" => {
                    error_type = Some(match required_enum_attr(e, "error-bar type")?.as_str() {
                        "both" => ErrorBarType::Both,
                        "plus" => ErrorBarType::Plus,
                        "minus" => ErrorBarType::Minus,
                        value => return Err(invalid_attribute("error-bar type", value.as_bytes())),
                    });
                },
                b"errValType" => {
                    value_type = Some(
                        match required_enum_attr(e, "error-bar value type")?.as_str() {
                            "fixedVal" => ErrorBarValueType::Fixed,
                            "percentage" => ErrorBarValueType::Percentage,
                            "stdDev" => ErrorBarValueType::StdDev,
                            "stdErr" => ErrorBarValueType::StdErr,
                            "cust" => ErrorBarValueType::Custom,
                            value => {
                                return Err(invalid_attribute(
                                    "error-bar value type",
                                    value.as_bytes(),
                                ));
                            },
                        },
                    );
                },
                b"noEndCap" => no_end_cap = parse_bool_attr(e)?,
                b"val" => value = Some(required_nonnegative_f64_attr(e, "error-bar value")?),
                _ => {},
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"errBars" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart error bars".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    let error_bar = ErrorBar {
        direction: direction.ok_or_else(|| missing_attribute("error-bar direction"))?,
        error_type: error_type.ok_or_else(|| missing_attribute("error-bar type"))?,
        value_type: value_type.ok_or_else(|| missing_attribute("error-bar value type"))?,
        value,
        plus_values,
        minus_values,
        no_end_cap,
        shape_properties,
        extension_list,
    };
    match error_bar.value_type {
        ErrorBarValueType::Fixed | ErrorBarValueType::Percentage | ErrorBarValueType::StdDev
            if error_bar.value.is_none() =>
        {
            return Err(missing_attribute("error-bar scalar value"));
        },
        ErrorBarValueType::Custom
            if error_bar.plus_values.is_none() && error_bar.minus_values.is_none() =>
        {
            return Err(missing_attribute("custom error-bar values"));
        },
        ErrorBarValueType::StdErr | ErrorBarValueType::Custom if error_bar.value.is_some() => {
            return Err(Error::Invalid(
                "standard-error and custom error bars cannot have a scalar value".to_string(),
            ));
        },
        _ => {},
    }
    Ok(error_bar)
}

pub(crate) fn parse_string_data<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Option<StringData>> {
    let mut data = StringData::from_values(Vec::new());
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"f" => {
                let formula = parse_text_element(reader, b"f")?;
                data.source_ref = Some(DataSourceRef::new(formula));
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pt" => {
                if let Some(text) = parse_point_text(reader)? {
                    data.values.push(text);
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"cat" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart element".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    Ok(Some(data))
}

pub(crate) fn parse_numeric_data<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Option<NumericData>> {
    let mut data = NumericData::from_values(Vec::new());
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"f" => {
                let formula = parse_text_element(reader, b"f")?;
                data.source_ref = Some(DataSourceRef::new(formula));
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"formatCode" => {
                data.format_code = Some(parse_text_element(reader, b"formatCode")?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pt" => {
                if let Some(val) = parse_point_value(reader)? {
                    data.values.push(val);
                }
            },
            Ok(Event::End(ref e))
                if matches!(
                    e.local_name().as_ref(),
                    b"val" | b"xVal" | b"yVal" | b"bubbleSize" | b"plus" | b"minus"
                ) =>
            {
                break;
            },
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart element".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    Ok(Some(data))
}

pub(crate) fn parse_series_title<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Option<TitleText>> {
    let mut title = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"strRef" => {
                let formula = parse_series_title_reference(reader)?;
                set_title(
                    &mut title,
                    TitleText::Reference(DataSourceRef::new(formula)),
                )?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"rich" => {
                let text = parse_series_title_rich_text(reader)?;
                set_title(&mut title, TitleText::Literal(RichText::new(text)))?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"f" => {
                let formula = parse_text_element(reader, b"f")?;
                set_title(
                    &mut title,
                    TitleText::Reference(DataSourceRef::new(formula)),
                )?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"v" => {
                let value = parse_text_element(reader, b"v")?;
                set_title(&mut title, TitleText::Literal(RichText::new(value)))?;
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"tx" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid(
                    "unterminated chart series title".to_string(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(title)
}

/// Read a series title reference and ignore its optional cached string data.
///
/// Spreadsheet applications commonly emit both `<c:f>` and a `<c:strCache>`
/// under `<c:strRef>`. The cache is a derived value, not a second title source;
/// the typed model retains the authoritative reference and avoids rejecting a
/// valid Office-produced chart as a duplicate title.
pub(crate) fn parse_series_title_reference<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<String> {
    let mut formula = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"f" => {
                if formula.is_some() {
                    return Err(Error::Invalid(
                        "chart series title reference contains duplicate formulas".into(),
                    ));
                }
                formula = Some(parse_text_element(reader, b"f")?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"strRef" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid(
                    "unterminated chart series title reference".into(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    formula.ok_or_else(|| Error::Invalid("chart series title reference requires a formula".into()))
}

pub(crate) fn parse_series_title_rich_text<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<String> {
    let mut text = String::new();
    let mut in_text = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"t" => {
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
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"rich" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart series rich text".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(text)
}

pub(crate) fn set_title(target: &mut Option<TitleText>, title: TitleText) -> Result<()> {
    if target.replace(title).is_some() {
        return Err(Error::Invalid(
            "chart series has duplicate title values".to_string(),
        ));
    }
    Ok(())
}

pub(crate) fn parse_text_element<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
    end_name: &[u8],
) -> Result<String> {
    let mut text = String::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Text(value)) => {
                text.push_str(
                    &value
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?,
                );
            },
            Ok(Event::CData(value)) => {
                text.push_str(
                    &value
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?,
                );
            },
            Ok(Event::GeneralRef(reference)) => {
                text.push_str(&decode_xml_reference(&reference)?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == end_name => break,
            Ok(Event::Start(ref element) | Event::Empty(ref element))
                if element.local_name().as_ref() == IGNORED_NAMESPACE_ELEMENT.as_bytes() => {},
            Ok(Event::Start(_) | Event::Empty(_)) => {
                return Err(Error::Invalid(
                    "chart text element contains nested markup".to_string(),
                ));
            },
            Ok(Event::Eof) => {
                return Err(Error::Invalid(
                    "unterminated chart text element".to_string(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(text)
}

pub(crate) fn parse_point_text<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Option<String>> {
    let mut text = String::new();
    let mut buf = Vec::new();
    let mut in_v = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"v" => {
                in_v = true;
            },
            Ok(Event::Text(e)) if in_v => {
                text.push_str(&e.decode().map_err(|e| Error::Xml(e.to_string()))?);
            },
            Ok(Event::CData(e)) if in_v => {
                text.push_str(&e.decode().map_err(|e| Error::Xml(e.to_string()))?);
            },
            Ok(Event::GeneralRef(reference)) if in_v => {
                text.push_str(&decode_xml_reference(&reference)?);
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"v" => {
                in_v = false;
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"pt" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart element".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    Ok(Some(text))
}

pub(crate) fn parse_point_value<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Option<f64>> {
    if let Some(text) = parse_point_text(reader)? {
        Ok(Some(text.trim().parse::<f64>().map_err(|_| {
            Error::Invalid(format!("invalid chart numeric point '{text}'"))
        })?))
    } else {
        Ok(None)
    }
}
