use super::super::model::*;
use super::super::package::Part;
use super::CONTENT_TYPE;
use super::limits::{
    A, A_STRICT, CX, IMAGE_REL, MAX_FORMAT_OVERRIDES, MAX_GEO_BINARY_BYTES, MAX_LABEL_TEXT_BYTES,
    MAX_PRINT_TEXT_BYTES, MAX_XML_BYTES, PACKAGE_REL, R, WORKBOOK_CONTENT_TYPES,
};
use litchi_opc::OpcPackage;
use litchi_opc::PackURI;
use litchi_opc::part::BlobPart;
use litchi_opc::part::Part as _;

fn producer_xml(auto_update: &str) -> String {
    format!(
        r#"<?xml version="1.0"?><cx:chartSpace xmlns:cx="{CX}" xmlns:a="{A}" xmlns:r="{R}" version="1.0" featureList="sunburst" fallbackImg="rIdImage"><cx:chartData><cx:externalData r:id="rIdWorkbook" cx:autoUpdate="{auto_update}"/><cx:data id="0"><cx:strDim type="cat"><cx:f dir="row">Sheet1!$A$2:$A$3</cx:f><cx:lvl ptCount="2" name="Category"><cx:pt idx="0">A</cx:pt><cx:pt idx="1">B</cx:pt></cx:lvl></cx:strDim><cx:numDim type="size"><cx:lvl ptCount="2" formatCode="General"><cx:pt idx="0">3</cx:pt><cx:pt idx="1">4</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:title pos="t" align="ctr" overlay="1"><cx:tx><cx:txData><cx:v>Quarterly</cx:v></cx:txData></cx:tx><cx:spPr><a:noFill/></cx:spPr><cx:txPr><a:bodyPr/><a:lstStyle/><a:p/></cx:txPr><cx:offset top="0.1" left="-0.2"/></cx:title><cx:plotArea><cx:plotAreaRegion><cx:plotSurface><cx:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></cx:spPr><cx:extLst/></cx:plotSurface><cx:series layoutId="sunburst" uniqueId="series-0"><cx:tx><cx:txData><cx:f>Sheet1!$B$1</cx:f><cx:v>Revenue</cx:v></cx:txData></cx:tx><cx:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></cx:spPr><cx:valueColors><cx:minColor><a:srgbClr val="0000FF"/></cx:minColor><cx:midColor><a:schemeClr val="accent1"/></cx:midColor><cx:maxColor><a:srgbClr val="FF0000"/></cx:maxColor></cx:valueColors><cx:valueColorPositions count="3"><cx:min><cx:extreme/></cx:min><cx:mid><cx:percent val="50"/></cx:mid><cx:max><cx:number val="10"/></cx:max></cx:valueColorPositions><cx:dataPt idx="1"><cx:spPr><a:ln/></cx:spPr></cx:dataPt><cx:dataLabels pos="bestFit"><cx:numFmt formatCode="0.0" sourceLinked="0"/><cx:spPr/><cx:txPr><a:bodyPr/><a:lstStyle/><a:p/></cx:txPr><cx:visibility seriesName="1" categoryName="true" value="0"/><cx:separator> | </cx:separator><cx:dataLabel idx="0" pos="t"><cx:spPr/><cx:visibility value="1"/><cx:separator>:</cx:separator></cx:dataLabel><cx:dataLabelHidden idx="1"/></cx:dataLabels><cx:dataId val="0"/><cx:layoutPr><cx:parentLabelLayout val="banner"/><cx:regionLabelLayout val="bestFitOnly"/><cx:visibility connectorLines="1" meanLine="false" outliers="true"/><cx:binning intervalClosed="l" underflow="auto" overflow="10"><cx:binCount>4</cx:binCount></cx:binning><cx:geography projectionType="mercator" viewedRegionType="countryRegion" cultureLanguage="en-US" cultureRegion="US" attribution="Map data"><cx:geoCache provider="Microsoft"><cx:binary>AQID</cx:binary><cx:clear><cx:geoLocationQueryResults><cx:geoLocationQueryResult><cx:geoLocationQuery countryRegion="US" entityType="CountryRegion"/><cx:geoLocations><cx:geoLocation latitude="47.6" longitude="-122.3" entityName="United States" entityType="CountryRegion"><cx:address countryRegion="United States" isoCountryCode="US"/></cx:geoLocation></cx:geoLocations></cx:geoLocationQueryResult></cx:geoLocationQueryResults><cx:geoDataEntityQueryResults><cx:geoDataEntityQueryResult><cx:geoDataEntityQuery entityType="CountryRegion" entityId="US"/><cx:geoData entityName="United States" entityId="US" east="-66" west="-125" north="49" south="24"><cx:geoPolygons><cx:geoPolygon polygonId="p1" numPoints="4" pcaRings="0,0 1,0 1,1 0,0"/></cx:geoPolygons><cx:copyrights><cx:copyright>Map provider</cx:copyright></cx:copyrights></cx:geoData></cx:geoDataEntityQueryResult></cx:geoDataEntityQueryResults><cx:geoDataPointToEntityQueryResults><cx:geoDataPointToEntityQueryResult><cx:geoDataPointQuery entityType="CountryRegion" latitude="47.6" longitude="-122.3"/><cx:geoDataPointToEntityQuery entityType="CountryRegion" entityId="US"/></cx:geoDataPointToEntityQueryResult></cx:geoDataPointToEntityQueryResults><cx:geoChildEntitiesQueryResults><cx:geoChildEntitiesQueryResult><cx:geoChildEntitiesQuery entityId="US"><cx:geoChildTypes><cx:entityType>AdminDistrict</cx:entityType></cx:geoChildTypes></cx:geoChildEntitiesQuery><cx:geoChildEntities><cx:geoHierarchyEntity entityName="Washington" entityId="WA" entityType="AdminDistrict"/></cx:geoChildEntities></cx:geoChildEntitiesQueryResult></cx:geoChildEntitiesQueryResults><cx:geoParentEntitiesQueryResults><cx:geoParentEntitiesQueryResult><cx:geoParentEntitiesQuery entityId="WA"/><cx:geoEntity entityName="Washington" entityType="AdminDistrict"/><cx:geoParentEntity entityId="US"/></cx:geoParentEntitiesQueryResult></cx:geoParentEntitiesQueryResults></cx:clear></cx:geoCache></cx:geography><cx:statistics quartileMethod="inclusive"/><cx:subtotals><cx:idx>1</cx:idx><cx:idx>3</cx:idx></cx:subtotals></cx:layoutPr><cx:axisId>7</cx:axisId><cx:axisId>8</cx:axisId></cx:series><cx:extLst/></cx:plotAreaRegion><cx:axis id="7"><cx:catScaling gapWidth="0.5"/></cx:axis><cx:axis id="8" hidden="0"><cx:valScaling min="auto" max="10" majorUnit="2" minorUnit="auto"/><cx:title><cx:tx><cx:txData><cx:v>Value Axis</cx:v></cx:txData></cx:tx><cx:spPr><a:noFill/></cx:spPr><cx:txPr><a:bodyPr/><a:lstStyle/><a:p/></cx:txPr><cx:offset top="0.25" left="0.5"/><cx:extLst/></cx:title><cx:units unit="millions"><cx:unitsLabel><cx:tx><cx:txData><cx:v>Millions</cx:v></cx:txData></cx:tx><cx:spPr><a:noFill/></cx:spPr><cx:txPr><a:bodyPr/><a:lstStyle/><a:p/></cx:txPr><cx:extLst/></cx:unitsLabel><cx:extLst/></cx:units><cx:majorGridlines><cx:spPr><a:ln/></cx:spPr><cx:extLst/></cx:majorGridlines><cx:minorGridlines/><cx:majorTickMarks type="in"><cx:extLst/></cx:majorTickMarks><cx:minorTickMarks type="none"/><cx:tickLabels><cx:extLst/></cx:tickLabels><cx:numFmt formatCode="0.00" sourceLinked="true"/><cx:spPr><a:ln w="12700"/></cx:spPr><cx:txPr><a:bodyPr/><a:lstStyle/><a:p/></cx:txPr><cx:extLst/></cx:axis><cx:spPr><a:noFill/></cx:spPr><cx:extLst/></cx:plotArea><cx:legend pos="r" align="max" overlay="false"><cx:spPr><a:noFill/></cx:spPr><cx:txPr><a:bodyPr/><a:lstStyle/><a:p/></cx:txPr><cx:offset top="INF" left="0"/></cx:legend><cx:extLst/></cx:chart><cx:spPr><a:noFill/></cx:spPr><cx:txPr><a:bodyPr/><a:lstStyle/><a:p/></cx:txPr><cx:clrMapOvr bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><cx:fmtOvrs><cx:fmtOvr idx="2"><cx:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></cx:spPr><cx:extLst/></cx:fmtOvr></cx:fmtOvrs><cx:printSettings><cx:headerFooter alignWithMargins="0" differentOddEven="1" differentFirst="true"><cx:oddHeader>Quarterly Report</cx:oddHeader><cx:oddFooter>Confidential</cx:oddFooter><cx:evenHeader>Even Page</cx:evenHeader><cx:evenFooter>2</cx:evenFooter><cx:firstHeader>Cover</cx:firstHeader><cx:firstFooter>First</cx:firstFooter></cx:headerFooter><cx:pageMargins l="0.7" r="0.7" t="0.75" b="0.75" header="0.3" footer="0.3"/><cx:pageSetup paperSize="9" firstPageNumber="2" orientation="landscape" blackAndWhite="1" draft="false" useFirstPageNumber="true" horizontalDpi="600" verticalDpi="300" copies="2"/></cx:printSettings><cx:extLst/></cx:chartSpace>"#
    )
}

fn chart_part(xml: String) -> BlobPart {
    BlobPart::new(
        PackURI::new("/ppt/charts/chartEx1.xml").unwrap(),
        CONTENT_TYPE.into(),
        xml.into_bytes(),
    )
}

#[test]
fn parses_producer_core_and_preserves_round_trip() {
    let part = chart_part(producer_xml("0"));
    let document = Part::from_part(&part).unwrap().parse().unwrap();
    assert_eq!(document.info().version, "1.0");
    assert_eq!(document.info().features, ["sunburst"]);
    let data = &document.info().data_sets[0];
    assert_eq!(
        (data.id, data.string_dimensions, data.numeric_dimensions),
        (0, 1, 1)
    );
    assert!(
        matches!(&data.dimensions[0], Dimension::String { kind: StringDimensionType::Category, formula: Some(Formula { direction: FormulaDirection::Row, .. }), levels, .. } if levels[0].points[0].value == "A")
    );
    assert!(
        matches!(&data.dimensions[1], Dimension::Numeric { kind: NumericDimensionType::Size, levels, .. } if levels[0].points[0].value == "3")
    );
    let series = &document.info().series[0];
    assert_eq!(series.layout, SeriesLayout::Sunburst);
    assert!(
        matches!(&series.text, Some(Text::Data { formula: Some(Formula { direction: FormulaDirection::Column, .. }), value: Some(value) }) if value == "Revenue")
    );
    assert_eq!(series.shape_properties.as_ref().unwrap().child_elements, 1);
    assert_eq!(
        series
            .value_colors
            .as_ref()
            .unwrap()
            .middle
            .as_ref()
            .unwrap()
            .kind,
        ColorKind::Scheme
    );
    assert!(
        matches!(series.value_color_positions.as_ref().unwrap().middle, Some(ColorPosition::Percent(ref value)) if value == "50")
    );
    assert_eq!(series.data_points[0].index, 1);
    let labels = series.data_labels.as_ref().unwrap();
    assert_eq!(labels.position, Some(DataLabelPosition::BestFit));
    assert_eq!(labels.number_format.as_ref().unwrap().format_code, "0.0");
    assert_eq!(
        labels.visibility.as_ref().unwrap().category_name,
        Some(true)
    );
    assert_eq!(labels.labels[0].index, 0);
    assert_eq!(labels.hidden_indices, [1]);
    assert_eq!(series.data_id, Some(0));
    assert_eq!(series.axis_ids, [7, 8]);
    let layout = series.layout_properties.as_ref().unwrap();
    assert_eq!(layout.parent_label, Some(ParentLabelLayout::Banner));
    assert_eq!(layout.region_label, Some(RegionLabelLayout::BestFitOnly));
    assert_eq!(
        layout.visibility.as_ref().unwrap().connector_lines,
        Some(true)
    );
    assert!(matches!(
        layout.binning.as_ref().unwrap().choice,
        Some(BinningChoice::Count(4))
    ));
    assert!(matches!(
        layout.geography.as_ref().unwrap(),
        Geography {
            projection: Some(GeoProjection::Mercator),
            viewed_region: Some(GeoMappingLevel::CountryRegion),
            has_cache: true,
            ..
        }
    ));
    let cache = layout.geography.as_ref().unwrap().cache.as_ref().unwrap();
    assert_eq!(cache.provider, "Microsoft");
    assert!(matches!(
        cache.entries[0],
        GeoCacheEntry::Binary {
            encoded_characters: 4,
            decoded_bytes: 3
        }
    ));
    let GeoCacheEntry::Clear(clear) = &cache.entries[1] else {
        panic!("expected clear geography cache")
    };
    assert_eq!(
        clear.location_query_results.as_ref().unwrap()[0]
            .location
            .as_ref()
            .unwrap()
            .entity_name,
        "United States"
    );
    assert_eq!(
        clear.data_entity_query_results.as_ref().unwrap()[0]
            .data
            .as_ref()
            .unwrap()
            .polygons
            .as_ref()
            .unwrap()[0]
            .num_points,
        "4"
    );
    assert_eq!(
        clear.data_point_to_entity_query_results.as_ref().unwrap()[0]
            .entity_query
            .as_ref()
            .unwrap()
            .entity_id,
        "US"
    );
    assert_eq!(
        clear.child_entities_query_results.as_ref().unwrap()[0]
            .children
            .as_ref()
            .unwrap()[0]
            .entity_id,
        "WA"
    );
    assert_eq!(
        clear.parent_entities_query_results.as_ref().unwrap()[0]
            .parent_entity_id
            .as_deref(),
        Some("US")
    );
    assert_eq!(layout.quartile_method, Some(QuartileMethod::Inclusive));
    assert_eq!(layout.subtotals, [1, 3]);
    assert!(document.info().has_plot_surface);
    assert_eq!(document.info().axes.len(), 2);
    assert!(matches!(
        document.info().axes[0].scaling,
        AxisScaling::Category { .. }
    ));
    assert!(matches!(
        document.info().axes[1].scaling,
        AxisScaling::Value { .. }
    ));
    let value_axis = &document.info().axes[1];
    assert!(
        matches!(&value_axis.title.as_ref().unwrap().text, Some(Text::Data { value: Some(value), .. }) if value == "Value Axis")
    );
    assert_eq!(
        value_axis
            .title
            .as_ref()
            .unwrap()
            .offset
            .as_ref()
            .unwrap()
            .left,
        "0.5"
    );
    assert!(value_axis.title.as_ref().unwrap().has_extension_list);
    let units = value_axis.units.as_ref().unwrap();
    assert_eq!(units.unit, Some(AxisUnit::Millions));
    assert!(
        matches!(&units.label.as_ref().unwrap().text, Some(Text::Data { value: Some(value), .. }) if value == "Millions")
    );
    assert!(units.has_extension_list && units.label.as_ref().unwrap().has_extension_list);
    assert_eq!(
        value_axis
            .major_gridlines
            .as_ref()
            .unwrap()
            .shape_properties
            .as_ref()
            .unwrap()
            .child_elements,
        1
    );
    assert!(value_axis.minor_gridlines.is_some());
    assert_eq!(
        value_axis.major_tick_marks.as_ref().unwrap().kind,
        Some(TickMarkType::Inside)
    );
    assert_eq!(
        value_axis.minor_tick_marks.as_ref().unwrap().kind,
        Some(TickMarkType::None)
    );
    assert!(value_axis.tick_labels.as_ref().unwrap().has_extension_list);
    assert_eq!(
        value_axis.number_format.as_ref().unwrap(),
        &NumberFormat {
            format_code: "0.00".into(),
            source_linked: Some(true)
        }
    );
    assert!(
        value_axis.shape_properties.is_some()
            && value_axis.text_properties.is_some()
            && value_axis.has_extension_list
    );
    let plot_area = &document.info().plot_area;
    assert!(plot_area.shape_properties.is_some() && plot_area.has_extension_list);
    let surface = plot_area.region.plot_surface.as_ref().unwrap();
    assert_eq!(surface.shape_properties.as_ref().unwrap().child_elements, 1);
    assert!(surface.has_extension_list && plot_area.region.has_extension_list);
    let chart_space = &document.info().chart_space_formatting;
    assert!(chart_space.shape_properties.is_some() && chart_space.text_properties.is_some());
    assert_eq!(
        chart_space
            .color_mapping_override
            .as_ref()
            .unwrap()
            .attributes,
        12
    );
    let overrides = chart_space.format_overrides.as_ref().unwrap();
    assert_eq!(overrides.len(), 1);
    assert_eq!(overrides[0].index, 2);
    assert!(overrides[0].shape_properties.is_some() && overrides[0].has_extension_list);
    let print = chart_space.print_settings.as_ref().unwrap();
    let header_footer = print.header_footer.as_ref().unwrap();
    assert_eq!(
        (
            header_footer.align_with_margins,
            header_footer.different_odd_even,
            header_footer.different_first
        ),
        (false, true, true)
    );
    assert_eq!(
        header_footer.odd_header.as_deref(),
        Some("Quarterly Report")
    );
    assert_eq!(print.page_margins.as_ref().unwrap().left, "0.7");
    let page_setup = print.page_setup.as_ref().unwrap();
    assert_eq!(
        (
            page_setup.paper_size,
            page_setup.first_page_number,
            page_setup.orientation
        ),
        (9, 2, PageOrientation::Landscape)
    );
    assert_eq!(
        (
            page_setup.black_and_white,
            page_setup.draft,
            page_setup.use_first_page_number
        ),
        (true, false, true)
    );
    assert_eq!(
        (
            page_setup.horizontal_dpi,
            page_setup.vertical_dpi,
            page_setup.copies
        ),
        (600, 300, 2)
    );
    assert!(chart_space.has_extension_list);
    let title = document.info().chart.title.as_ref().unwrap();
    assert_eq!(
        (title.position, title.alignment, title.overlay),
        (SidePosition::Top, PositionAlignment::Center, true)
    );
    assert!(
        matches!(&title.text, Some(Text::Data { formula: None, value: Some(value) }) if value == "Quarterly")
    );
    assert_eq!(
        title.offset.as_ref().unwrap(),
        &Offset {
            top: "0.1".into(),
            left: "-0.2".into()
        }
    );
    let legend = document.info().chart.legend.as_ref().unwrap();
    assert_eq!(
        (legend.position, legend.alignment, legend.overlay),
        (SidePosition::Right, PositionAlignment::Maximum, false)
    );
    assert_eq!(legend.offset.as_ref().unwrap().top, "INF");
    assert!(document.info().chart.has_extension_list);
    assert!(document.info().has_title && document.info().has_legend);
    assert_eq!(document.to_xml(), part.blob());
    let reparsed = chart_part(String::from_utf8(document.to_xml()).unwrap());
    assert_eq!(
        Part::from_part(&reparsed).unwrap().parse().unwrap().info(),
        document.info()
    );
}

#[test]
fn validates_inert_package_relationships_without_opening_targets() {
    let mut chart = chart_part(producer_xml("0"));
    chart.rels_mut().add_relationship(
        PACKAGE_REL[0].into(),
        "../embeddings/data.xlsx".into(),
        "rIdWorkbook".into(),
        false,
    );
    chart.rels_mut().add_relationship(
        IMAGE_REL[1].into(),
        "../media/fallback.png".into(),
        "rIdImage".into(),
        false,
    );
    let mut package = OpcPackage::new();
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/embeddings/data.xlsx").unwrap(),
        WORKBOOK_CONTENT_TYPES[0].into(),
        b"not opened".to_vec(),
    )));
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/media/fallback.png").unwrap(),
        "image/png".into(),
        b"not opened".to_vec(),
    )));
    let document = Part::from_part(&chart)
        .unwrap()
        .parse_in_package(&package)
        .unwrap();
    assert!(
        matches!(document.external_data_target(), Some(ExternalDataTarget::EmbeddedPackage { part_name, .. }) if part_name == "/ppt/embeddings/data.xlsx")
    );
    assert_eq!(
        document.fallback_image_part_name(),
        Some("/ppt/media/fallback.png")
    );
}

#[test]
fn rejects_hostile_schema_and_resource_cases() {
    let cases = [
        producer_xml("0").replace(CX, "urn:vendor:chartex"),
        producer_xml("0").replace(
            "<cx:data id=\"0\">",
            "<cx:data id=\"0\"></cx:data><cx:data id=\"0\">",
        ),
        producer_xml("0").replace("<cx:chartData>", "<cx:chart/><cx:chartData>"),
        producer_xml("0").replace("<cx:strDim type=\"cat\">", "<cx:strDim>"),
        format!(
            "<!DOCTYPE x [<!ENTITY e SYSTEM 'file:///etc/passwd'>]>{}",
            producer_xml("0")
        ),
    ];
    for xml in cases {
        assert!(Part::from_part(&chart_part(xml)).unwrap().parse().is_err());
    }
    let oversized = chart_part(" ".repeat(MAX_XML_BYTES + 1));
    assert!(Part::from_part(&oversized).unwrap().parse().is_err());
}

#[test]
fn rejects_invalid_dimension_choices_points_and_series_references() {
    let base = producer_xml("0");
    let cases = [
        base.replace("type=\"cat\"", "type=\"vendor\""),
        base.replace("type=\"size\"", "type=\"z\""),
        base.replace("<cx:f dir=\"row\">", "<cx:nf dir=\"row\">")
            .replace("</cx:f>", "</cx:nf>"),
        base.replace("</cx:strDim>", "<cx:f>late</cx:f></cx:strDim>"),
        base.replace(
            "</cx:lvl></cx:strDim>",
            "<cx:pt idx=\"0\">B</cx:pt></cx:lvl></cx:strDim>",
        ),
        base.replace("<cx:pt idx=\"0\">3</cx:pt>", "<cx:pt idx=\"9\">3</cx:pt>"),
        base.replace(">3</cx:pt>", ">not-a-number</cx:pt>"),
        base.replace("<cx:dataId val=\"0\"/>", "<cx:dataId val=\"99\"/>"),
        base.replace(
            "<cx:dataId val=\"0\"/>",
            "<cx:dataId val=\"0\"/><cx:dataId val=\"0\"/>",
        ),
        base.replace(
            "<cx:dataId val=\"0\"/>",
            "<cx:layoutPr/><cx:dataId val=\"0\"/>",
        ),
    ];
    for xml in cases {
        assert!(Part::from_part(&chart_part(xml)).unwrap().parse().is_err());
    }
}

#[test]
fn rejects_invalid_layout_axes_scaling_and_plot_surface_grammar() {
    let base = producer_xml("0");
    let cases = [
            base.replace("<cx:binning intervalClosed=\"l\"", "<cx:aggregation/><cx:binning intervalClosed=\"l\""),
            base.replace("<cx:binCount>4</cx:binCount>", "<cx:binSize>1</cx:binSize><cx:binCount>4</cx:binCount>"),
            base.replace("intervalClosed=\"l\"", "intervalClosed=\"center\""),
            base.replace("quartileMethod=\"inclusive\"", "quartileMethod=\"median\""),
            base.replace("<cx:idx>3</cx:idx>", "<cx:idx>1</cx:idx>"),
            base.replace(" cultureRegion=\"US\"", ""),
            base.replace("<cx:axis id=\"8\" hidden=\"0\">", "<cx:axis id=\"7\" hidden=\"0\">"),
            base.replace("<cx:axisId>8</cx:axisId>", "<cx:axisId>99</cx:axisId>"),
            base.replace("<cx:axisId>8</cx:axisId>", "<cx:axisId>7</cx:axisId>"),
            base.replace("majorUnit=\"2\"", "majorUnit=\"0\""),
            base.replace("gapWidth=\"0.5\"", "gapWidth=\"-1\""),
            base.replace("<cx:catScaling gapWidth=\"0.5\"/>", "<cx:catScaling gapWidth=\"0.5\"/><cx:valScaling/>"),
            base.replace(r#"<cx:plotSurface><cx:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></cx:spPr><cx:extLst/></cx:plotSurface>"#, "").replace("</cx:series><cx:extLst/></cx:plotAreaRegion>", r#"</cx:series><cx:plotSurface><cx:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></cx:spPr><cx:extLst/></cx:plotSurface></cx:plotAreaRegion>"#),
            base.replace(r#"<cx:plotSurface><cx:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></cx:spPr><cx:extLst/></cx:plotSurface>"#, r#"<cx:plotSurface><cx:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></cx:spPr><cx:extLst/></cx:plotSurface><cx:plotSurface><cx:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></cx:spPr><cx:extLst/></cx:plotSurface>"#),
            base.replace("<cx:parentLabelLayout val=\"banner\"/>", "").replace("<cx:regionLabelLayout val=\"bestFitOnly\"/>", "<cx:regionLabelLayout val=\"bestFitOnly\"/><cx:parentLabelLayout val=\"banner\"/>"),
        ];
    for xml in cases {
        assert!(Part::from_part(&chart_part(xml)).unwrap().parse().is_err());
    }
}

#[test]
fn rejects_invalid_axis_title_units_gridlines_ticks_and_formatting() {
    let base = producer_xml("0");
    let cases = [
        base.replace(
            "<cx:title><cx:tx><cx:txData><cx:v>Value Axis",
            "<cx:title vendor=\"1\"><cx:tx><cx:txData><cx:v>Value Axis",
        ),
        base.replace(
            "<cx:title><cx:tx><cx:txData><cx:v>Value Axis",
            "<cx:units/><cx:title><cx:tx><cx:txData><cx:v>Value Axis",
        ),
        base.replace(
            "</cx:title><cx:units unit=\"millions\">",
            "</cx:title><cx:title/><cx:units unit=\"millions\">",
        ),
        base.replace(
            "<cx:v>Value Axis</cx:v></cx:txData></cx:tx><cx:spPr>",
            "<cx:v>Value Axis</cx:v></cx:txData></cx:tx><cx:tx/><cx:spPr>",
        ),
        base.replace("unit=\"millions\"", "unit=\"vendor\""),
        base.replace(
            "</cx:unitsLabel><cx:extLst/></cx:units>",
            "</cx:unitsLabel><cx:unitsLabel/><cx:extLst/></cx:units>",
        ),
        base.replace(
            "<cx:units unit=\"millions\"><cx:unitsLabel>",
            "<cx:units unit=\"millions\"><cx:extLst/><cx:unitsLabel>",
        ),
        base.replace(
            "<cx:v>Millions</cx:v></cx:txData></cx:tx><cx:spPr>",
            "<cx:v>Millions</cx:v></cx:txData></cx:tx><cx:tx/><cx:spPr>",
        ),
        base.replace("<cx:majorGridlines>", "<cx:majorGridlines vendor=\"1\">"),
        base.replace(
            "<cx:majorGridlines><cx:spPr><a:ln/></cx:spPr>",
            "<cx:majorGridlines><cx:spPr/><cx:spPr><a:ln/></cx:spPr>",
        ),
        base.replace(
            "<cx:majorGridlines>",
            "<cx:minorGridlines/><cx:majorGridlines>",
        ),
        base.replace("type=\"in\"", "type=\"vendor\""),
        base.replace(
            "<cx:majorTickMarks type=\"in\"><cx:extLst/>",
            "<cx:majorTickMarks type=\"in\"><a:extLst/>",
        ),
        base.replace("<cx:tickLabels>", "<cx:tickLabels vendor=\"1\">"),
        base.replace(
            "<cx:tickLabels><cx:extLst/></cx:tickLabels>",
            "<cx:tickLabels><cx:vendor/></cx:tickLabels>",
        ),
        base.replace(
            "<cx:numFmt formatCode=\"0.00\" sourceLinked=\"true\"/>",
            "<cx:numFmt sourceLinked=\"true\"/>",
        ),
        base.replace(
            "<cx:spPr><a:ln w=\"12700\"/></cx:spPr><cx:txPr>",
            "<cx:spPr><cx:ln w=\"12700\"/></cx:spPr><cx:txPr>",
        ),
        base.replace(
            "</cx:txPr><cx:extLst/></cx:axis>",
            "</cx:txPr><cx:txPr/><cx:extLst/></cx:axis>",
        ),
        base.replace(
            "<cx:extLst/></cx:axis>",
            "<cx:extLst/><cx:extLst/></cx:axis>",
        ),
        base.replace(
            "<cx:offset top=\"0.25\" left=\"0.5\"/>",
            "<cx:offset top=\"0.25\"/>",
        ),
        base.replace(" version=\"1.0\"", " version=\"0.0\"")
            .replace("<cx:offset top=\"0.1\" left=\"-0.2\"/>", "")
            .replace("<cx:offset top=\"INF\" left=\"0\"/>", ""),
    ];
    for xml in cases {
        assert!(Part::from_part(&chart_part(xml)).unwrap().parse().is_err());
    }
}

#[test]
fn rejects_hostile_geography_cache_grammar_and_bounds() {
    let base = producer_xml("0");
    let cases = [
            base.replace("<cx:geoCache provider=\"Microsoft\">", "<cx:geoCache>"),
            base.replace("<cx:binary>AQID</cx:binary><cx:clear>", ""),
            base.replace("<cx:binary>AQID</cx:binary>", "<cx:binary>AQI!</cx:binary>"),
            base.replace("<cx:geoLocationQueryResults>", "<cx:geoDataEntityQueryResults/><cx:geoLocationQueryResults>"),
            base.replace("<cx:geoDataEntityQueryResults>", "<cx:geoLocationQueryResults/><cx:geoDataEntityQueryResults>"),
            base.replacen("entityType=\"CountryRegion\"", "entityType=\"countryRegion\"", 1),
            base.replace("<cx:geoParentEntitiesQuery entityId=\"WA\"/>", ""),
            base.replace("<cx:geoLocations><cx:geoLocation", "<cx:geoLocations><cx:geoLocation entityName=\"duplicate\" entityType=\"Region\"/><cx:geoLocation"),
            base.replace("<cx:geoCache provider=\"Microsoft\">", "<cx:geoCache xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" provider=\"Microsoft\" r:id=\"rIdMap\">"),
            base.replace("<cx:geoPolygon polygonId=\"p1\" numPoints=\"4\"", "<cx:geoPolygon polygonId=\"p1\" numPoints=\"four\""),
        ];
    for xml in cases {
        assert!(Part::from_part(&chart_part(xml)).unwrap().parse().is_err());
    }
    let oversized = "A".repeat((MAX_GEO_BINARY_BYTES + 1).div_ceil(3) * 4);
    assert!(
        Part::from_part(&chart_part(base.replace("AQID", &oversized)))
            .unwrap()
            .parse()
            .is_err()
    );
}

#[test]
fn applies_chart_print_defaults() {
    let xml = producer_xml("0")
            .replace(" alignWithMargins=\"0\" differentOddEven=\"1\" differentFirst=\"true\"", "")
            .replace("<cx:pageSetup paperSize=\"9\" firstPageNumber=\"2\" orientation=\"landscape\" blackAndWhite=\"1\" draft=\"false\" useFirstPageNumber=\"true\" horizontalDpi=\"600\" verticalDpi=\"300\" copies=\"2\"/>", "<cx:pageSetup/>");
    let document = Part::from_part(&chart_part(xml)).unwrap().parse().unwrap();
    let print = document
        .info()
        .chart_space_formatting
        .print_settings
        .as_ref()
        .unwrap();
    let header_footer = print.header_footer.as_ref().unwrap();
    assert_eq!(
        (
            header_footer.align_with_margins,
            header_footer.different_odd_even,
            header_footer.different_first
        ),
        (true, false, false)
    );
    let setup = print.page_setup.as_ref().unwrap();
    assert_eq!(
        (setup.paper_size, setup.first_page_number, setup.orientation),
        (1, 1, PageOrientation::Default)
    );
    assert_eq!(
        (
            setup.black_and_white,
            setup.draft,
            setup.use_first_page_number
        ),
        (false, false, false)
    );
    assert_eq!(
        (setup.horizontal_dpi, setup.vertical_dpi, setup.copies),
        (600, 600, 1)
    );
}

#[test]
fn rejects_invalid_plot_and_chart_space_formatting_and_print_settings() {
    let base = producer_xml("0");
    let cases = [
            base.replace("<cx:plotSurface>", "<cx:plotSurface vendor=\"1\">"),
            base.replace("<cx:plotSurface><cx:spPr><a:solidFill><a:schemeClr val=\"accent2\"/></a:solidFill>", "<cx:plotSurface><cx:spPr><cx:solidFill><a:schemeClr val=\"accent2\"/></cx:solidFill>"),
            base.replace("<cx:plotSurface><cx:spPr>", "<cx:plotSurface><cx:extLst/><cx:spPr>"),
            base.replace("</cx:axis><cx:spPr><a:noFill/></cx:spPr>", "</cx:axis><cx:spPr/><cx:spPr><a:noFill/></cx:spPr>"),
            base.replace("</cx:series><cx:extLst/></cx:plotAreaRegion>", "</cx:series><cx:extLst/><cx:series layoutId=\"funnel\"/></cx:plotAreaRegion>"),
            base.replace("</cx:chart><cx:spPr><a:noFill/></cx:spPr><cx:txPr>", "</cx:chart><cx:txPr/><cx:spPr><a:noFill/></cx:spPr><cx:txPr>"),
            base.replace("</cx:chart><cx:spPr><a:noFill/></cx:spPr>", "</cx:chart><cx:spPr><cx:noFill/></cx:spPr>"),
            base.replace("folHlink=\"folHlink\"/>", "folHlink=\"folHlink\"><cx:vendor/></cx:clrMapOvr>"),
            base.replace("<cx:fmtOvrs>", "<cx:fmtOvrs vendor=\"1\">"),
            base.replace("<cx:fmtOvr idx=\"2\">", "<cx:fmtOvr>"),
            base.replace("</cx:fmtOvr></cx:fmtOvrs>", "</cx:fmtOvr><cx:fmtOvr idx=\"2\"/></cx:fmtOvrs>"),
            base.replace("<cx:fmtOvr idx=\"2\"><cx:spPr><a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill>", "<cx:fmtOvr idx=\"2\"><cx:spPr><cx:solidFill><a:srgbClr val=\"00FF00\"/></cx:solidFill>"),
            base.replace("<cx:fmtOvr idx=\"2\"><cx:spPr>", "<cx:fmtOvr idx=\"2\"><cx:extLst/><cx:spPr>"),
            base.replace("<cx:printSettings>", "<cx:printSettings vendor=\"1\">"),
            base.replace("</cx:headerFooter><cx:pageMargins", "</cx:headerFooter><cx:pageSetup/><cx:pageMargins"),
            base.replace("</cx:headerFooter><cx:pageMargins", "</cx:headerFooter><cx:headerFooter/><cx:pageMargins"),
            base.replace("differentFirst=\"true\"", "differentFirst=\"yes\""),
            base.replace("<cx:oddHeader>Quarterly Report</cx:oddHeader><cx:oddFooter>", "<cx:oddFooter/><cx:oddHeader>Quarterly Report</cx:oddHeader><cx:oddFooter>"),
            base.replace("<cx:oddHeader>Quarterly Report</cx:oddHeader>", "<cx:oddHeader><cx:v>Quarterly Report</cx:v></cx:oddHeader>"),
            base.replace(" l=\"0.7\"", ""),
            base.replace("l=\"0.7\"", "l=\"invalid\""),
            base.replace("orientation=\"landscape\"", "orientation=\"sideways\""),
            base.replace("blackAndWhite=\"1\"", "blackAndWhite=\"yes\""),
            base.replace("paperSize=\"9\"", "paperSize=\"-1\""),
            base.replace("horizontalDpi=\"600\"", "horizontalDpi=\"999999999999\""),
            base.replace("copies=\"2\"/>", "copies=\"2\"><cx:vendor/></cx:pageSetup>"),
        ];
    for (index, xml) in cases.into_iter().enumerate() {
        assert!(
            Part::from_part(&chart_part(xml)).unwrap().parse().is_err(),
            "hostile chart-space case {index}"
        );
    }
    let oversized = base.replace("Quarterly Report", &"x".repeat(MAX_PRINT_TEXT_BYTES + 1));
    assert!(
        Part::from_part(&chart_part(oversized))
            .unwrap()
            .parse()
            .is_err()
    );

    let many = (0..=MAX_FORMAT_OVERRIDES)
        .map(|index| format!("<cx:fmtOvr idx=\"{index}\"/>"))
        .collect::<String>();
    let excessive = base.replace("<cx:fmtOvr idx=\"2\"><cx:spPr><a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill></cx:spPr><cx:extLst/></cx:fmtOvr>", &many);
    assert!(
        Part::from_part(&chart_part(excessive))
            .unwrap()
            .parse()
            .is_err()
    );
}

#[test]
fn parses_strict_rich_series_text_as_bounded_inert_drawingml() {
    let xml = producer_xml("0").replace(A, A_STRICT).replace(
        "<cx:txData><cx:f>Sheet1!$B$1</cx:f><cx:v>Revenue</cx:v></cx:txData>",
        "<cx:rich><a:bodyPr/><a:lstStyle/><a:p/></cx:rich>",
    );
    let document = Part::from_part(&chart_part(xml)).unwrap().parse().unwrap();
    assert!(matches!(
        document.info().series[0].text,
        Some(Text::Rich(DrawingPayload {
            child_elements: 3,
            ..
        }))
    ));
}

#[test]
fn rejects_invalid_series_formatting_colors_points_and_labels() {
    let base = producer_xml("0");
    let cases = [
            base.replace("</cx:txData></cx:tx>", "</cx:txData><cx:rich/></cx:tx>"),
            base.replace("<cx:f>Sheet1!$B$1</cx:f><cx:v>Revenue</cx:v>", "<cx:v>Revenue</cx:v><cx:f>Sheet1!$B$1</cx:f>"),
            base.replacen("<cx:spPr>", "<a:spPr>", 1).replacen("</cx:spPr>", "</a:spPr>", 1),
            base.replacen("<a:solidFill>", "<cx:solidFill>", 1).replacen("</a:solidFill>", "</cx:solidFill>", 1),
            base.replace("val=\"0000FF\"", "val=\"00GGFF\""),
            base.replace("<cx:minColor>", "<cx:maxColor>").replace("</cx:minColor>", "</cx:maxColor>"),
            base.replace("count=\"3\"", "count=\"4\""),
            base.replace("<cx:mid><cx:percent val=\"50\"/></cx:mid>", "<cx:mid><cx:extreme/></cx:mid>"),
            base.replace("<cx:percent val=\"50\"/>", "<cx:percent val=\"101\"/>"),
            base.replace("<cx:number val=\"10\"/>", "<cx:number val=\"not-number\"/>"),
            base.replace("</cx:dataPt>", "</cx:dataPt><cx:dataPt idx=\"1\"/>"),
            base.replace("<cx:dataPt idx=\"1\">", "<cx:dataPt idx=\"99\">"),
            base.replace("<cx:dataLabels pos=\"bestFit\">", "<cx:dataLabels pos=\"vendor\">"),
            base.replace("<cx:numFmt formatCode=\"0.0\" sourceLinked=\"0\"/>", "<cx:numFmt sourceLinked=\"0\"/>"),
            base.replace("categoryName=\"true\"", "categoryName=\"yes\""),
            base.replace("<cx:dataLabelHidden idx=\"1\"/>", "<cx:dataLabelHidden idx=\"0\"/>"),
            base.replace("<cx:dataLabel idx=\"0\"", "<cx:dataLabel idx=\"99\""),
            base.replace("<cx:dataLabelHidden idx=\"1\"/>", "<cx:dataLabelHidden idx=\"99\"/>"),
            base.replace("<cx:visibility seriesName=\"1\" categoryName=\"true\" value=\"0\"/><cx:separator> | </cx:separator>", "<cx:separator> | </cx:separator><cx:visibility seriesName=\"1\" categoryName=\"true\" value=\"0\"/>"),
            base.replace("</cx:dataLabels><cx:dataId", "</cx:dataLabels><cx:dataLabels/><cx:dataId"),
            base.replace("<cx:dataLabel idx=\"0\" pos=\"t\"><cx:spPr/>", "<cx:dataLabel idx=\"0\" pos=\"t\"><cx:spPr/><cx:spPr/>"),
        ];
    for xml in cases {
        assert!(Part::from_part(&chart_part(xml)).unwrap().parse().is_err());
    }
    let oversized = base.replace(" | ", &"x".repeat(MAX_LABEL_TEXT_BYTES + 1));
    assert!(
        Part::from_part(&chart_part(oversized))
            .unwrap()
            .parse()
            .is_err()
    );
}

#[test]
fn applies_title_legend_defaults_and_accepts_mp_feature_offsets() {
    let xml = producer_xml("0")
        .replace(
            " version=\"1.0\" featureList=\"sunburst\"",
            " version=\"0.0\" featureList=\"mp\"",
        )
        .replace(" pos=\"t\" align=\"ctr\" overlay=\"1\"", "")
        .replace(" pos=\"r\" align=\"max\" overlay=\"false\"", "");
    let document = Part::from_part(&chart_part(xml)).unwrap().parse().unwrap();
    let title = document.info().chart.title.as_ref().unwrap();
    assert_eq!(
        (title.position, title.alignment, title.overlay),
        (SidePosition::Top, PositionAlignment::Center, false)
    );
    let legend = document.info().chart.legend.as_ref().unwrap();
    assert_eq!(
        (legend.position, legend.alignment, legend.overlay),
        (SidePosition::Right, PositionAlignment::Center, false)
    );
}

#[test]
fn rejects_invalid_chart_title_legend_and_offset_grammar() {
    let base = producer_xml("0");
    let cases = [
            base.replace("</cx:title><cx:plotArea>", "</cx:title><cx:title/><cx:plotArea>"),
            base.replace("<cx:plotArea>", "<cx:legend/><cx:plotArea>"),
            base.replace("<cx:extLst/></cx:chart>", "<cx:extLst/><cx:extLst/></cx:chart>"),
            base.replace("<cx:chart>", "<cx:chart vendor=\"1\">"),
            base.replace("<cx:title pos=\"t\"", "<cx:title pos=\"vendor\""),
            base.replace("align=\"ctr\" overlay=\"1\"", "align=\"middle\" overlay=\"1\""),
            base.replace("overlay=\"1\"", "overlay=\"yes\""),
            base.replace("<cx:tx><cx:txData><cx:v>Quarterly</cx:v></cx:txData></cx:tx><cx:spPr>", "<cx:spPr/><cx:tx><cx:txData><cx:v>Quarterly</cx:v></cx:txData></cx:tx><cx:spPr>"),
            base.replace("<cx:tx><cx:txData><cx:v>Quarterly</cx:v></cx:txData></cx:tx>", "<cx:tx/>"),
            base.replacen("<a:noFill/>", "<cx:noFill/>", 1),
            base.replace("<cx:offset top=\"0.1\" left=\"-0.2\"/>", "<cx:offset left=\"-0.2\"/>"),
            base.replace("top=\"0.1\" left=\"-0.2\"", "top=\"bad\" left=\"-0.2\""),
            base.replace("<cx:offset top=\"0.1\" left=\"-0.2\"/>", "<cx:offset top=\"0.1\" left=\"-0.2\"><cx:x/></cx:offset>"),
            base.replace(" version=\"1.0\"", " version=\"0.0\""),
            base.replace("<cx:legend pos=\"r\"", "<cx:legend pos=\"side\""),
            base.replace("align=\"max\" overlay=\"false\"", "align=\"middle\" overlay=\"false\""),
            base.replace("overlay=\"false\"", "overlay=\"off\""),
            base.replace("<cx:txPr><a:bodyPr/><a:lstStyle/><a:p/></cx:txPr><cx:offset top=\"INF\"", "<cx:offset top=\"INF\" left=\"0\"/><cx:txPr><a:bodyPr/><a:lstStyle/><a:p/></cx:txPr><cx:offset top=\"INF\""),
            base.replace("<cx:legend pos=\"r\" align=\"max\" overlay=\"false\"><cx:spPr>", "<cx:legend pos=\"r\" align=\"max\" overlay=\"false\"><cx:spPr/><cx:spPr>"),
            base.replace("<cx:plotArea>", "<cx:vendorPlotArea>").replace("</cx:plotArea>", "</cx:vendorPlotArea>"),
        ];
    for xml in cases {
        assert!(Part::from_part(&chart_part(xml)).unwrap().parse().is_err());
    }
}

#[test]
fn rejects_auto_update_external_missing_wrong_type_and_ambiguous_targets() {
    for (auto, rel_type, target, external) in [
        ("1", PACKAGE_REL[0], "../embeddings/data.xlsx", false),
        ("0", IMAGE_REL[0], "../embeddings/data.xlsx", false),
        ("0", PACKAGE_REL[0], "https://example.test/data.xlsx", true),
        ("0", PACKAGE_REL[0], "../embeddings/data.xlsx#x", false),
    ] {
        let mut chart = chart_part(producer_xml(auto).replace(" fallbackImg=\"rIdImage\"", ""));
        chart.rels_mut().add_relationship(
            rel_type.into(),
            target.into(),
            "rIdWorkbook".into(),
            external,
        );
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/embeddings/data.xlsx").unwrap(),
            WORKBOOK_CONTENT_TYPES[0].into(),
            Vec::new(),
        )));
        assert!(
            Part::from_part(&chart)
                .unwrap()
                .parse_in_package(&package)
                .is_err()
        );
    }
}
