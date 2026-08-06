use super::*;

pub(crate) fn parse_legend<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Legend> {
    let mut position = LegendPosition::Right;
    let mut overlay = false;
    let mut layout = None;
    let mut entries = Vec::new();
    let mut shape_properties = None;
    let mut text_properties = None;
    let mut extension_list = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"legendEntry" => {
                let entry = parse_legend_entry(reader)?;
                if entries
                    .iter()
                    .any(|existing: &LegendEntry| existing.index == entry.index)
                {
                    return Err(Error::Invalid(format!(
                        "chart legend contains duplicate entry index {}",
                        entry.index
                    )));
                }
                entries.push(entry);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"legendEntry" => {
                return Err(Error::Invalid(
                    "chart legend entry is missing its index".into(),
                ));
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"layout" => {
                if layout.is_some() {
                    return Err(Error::Invalid(
                        "chart legend contains duplicate layouts".into(),
                    ));
                }
                layout = Some(parse_layout(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"layout" => {
                layout = Some(match layout {
                    None => Layout::new(),
                    Some(_) => {
                        return Err(Error::Invalid(
                            "chart legend contains duplicate layouts".into(),
                        ));
                    },
                });
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart legend contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart legend shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart legend contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart legend contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(TextProperties::from_xml(
                    reader.capture_fragment(element, "chart legend text properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart legend contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(TextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart legend contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(element, "chart legend extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart legend contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"legendPos" => {
                        let value = get_attr(e, b"val")
                            .ok_or_else(|| missing_attribute("chart legend position"))?;
                        position = match value.as_slice() {
                            b"b" => LegendPosition::Bottom,
                            b"l" => LegendPosition::Left,
                            b"r" => LegendPosition::Right,
                            b"t" => LegendPosition::Top,
                            b"tr" => LegendPosition::TopRight,
                            _ => return Err(invalid_attribute("chart legend position", &value)),
                        };
                    },
                    b"overlay" => {
                        overlay = parse_bool_attr(e)?;
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"legend" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart element".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let mut legend = Legend::new(position).with_overlay(overlay);
    legend.layout = layout;
    legend.entries = entries;
    legend.shape_properties = shape_properties;
    legend.text_properties = text_properties;
    legend.extension_list = extension_list;
    Ok(legend)
}

pub(crate) fn parse_legend_entry<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<LegendEntry> {
    let mut index = None;
    let mut deleted = false;
    let mut saw_delete = false;
    let mut text_properties = None;
    let mut extension_list = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if saw_delete || text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart legend entry contains multiple choice values".into(),
                    ));
                }
                text_properties = Some(TextProperties::from_xml(
                    reader.capture_fragment(element, "chart legend-entry text properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if saw_delete || text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart legend entry contains multiple choice values".into(),
                    ));
                }
                text_properties = Some(TextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart legend entry contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(element, "chart legend-entry extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart legend entry contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"idx" => {
                        if index.is_some() {
                            return Err(Error::Invalid(
                                "chart legend entry contains duplicate indexes".into(),
                            ));
                        }
                        index = Some(required_u32_attr(element, "chart legend entry index")?);
                    },
                    b"delete" => {
                        if saw_delete || text_properties.is_some() {
                            return Err(Error::Invalid(
                                "chart legend entry contains multiple choice values".into(),
                            ));
                        }
                        deleted = parse_bool_attr(element)?;
                        saw_delete = true;
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"legendEntry" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart legend entry".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let index = index.ok_or_else(|| missing_attribute("chart legend entry index"))?;
    if !saw_delete && text_properties.is_none() {
        return Err(Error::Invalid(
            "chart legend entry is missing its delete or text-properties choice".into(),
        ));
    }
    let mut entry = LegendEntry::new(index);
    entry.deleted = deleted;
    entry.text_properties = text_properties;
    entry.extension_list = extension_list;
    Ok(entry)
}
