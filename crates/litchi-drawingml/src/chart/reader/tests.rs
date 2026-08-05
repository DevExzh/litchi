//! Focused DrawingML chart reader and writer conformance tests.

use super::read;

use crate::chart::axis::{
    Axis, AxisCrossBetween, AxisCrossMode, AxisLabelAlign, BuiltInUnit, CategoryAxis, DateAxis,
    DisplayUnits, SeriesAxis, TimeUnit, ValueAxis,
};
use crate::chart::bubble::{Scale as BubbleScale, Size as BubbleSize};
use crate::chart::data::{Layout, NumberFormat, NumericData, TitleText};
use crate::chart::legend::{Legend, LegendEntry};
use crate::chart::model::{
    Chart, ColorMapOverride, ColorSchemeIndex, ExtensionList, ExternalData, PageMargins,
    PageOrientation, PictureFormat, PictureOptions, PivotFormat, PrintSettings, ShapeProperties,
    TextProperties,
};
use crate::chart::plot_area::{
    Area3DTypeGroup, AreaTypeGroup, BandFormat, Bar3DTypeGroup, BarShape, BarTypeGroup,
    BubbleTypeGroup, DataTable, DoughnutTypeGroup, Line3DTypeGroup, LineTypeGroup, Lines,
    OfPieTypeGroup, Pie3DTypeGroup, PieTypeGroup, RadarTypeGroup, ScatterTypeGroup, StockTypeGroup,
    Surface3DTypeGroup, SurfaceTypeGroup, TypeGroup, UpDownBars,
};
use crate::chart::series::{
    DataLabel, DataLabels, DataPoint, ErrorBar, ErrorBarDirection, ErrorBarType, ErrorBarValueType,
    Marker, Series, Trendline,
};
use crate::chart::types::{
    AxisOrientation, AxisPosition, BarDirection, BarGrouping, DataLabelPosition, LayoutMode,
    LayoutTarget, LegendPosition, MarkerStyle, OfPieSplitType, OfPieType, RadarStyle, ScatterStyle,
    TickLabelPosition, TickMark,
};

#[test]
fn round_trips_and_validates_chart_language_and_pivot_source() {
    let xml = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:lang val="zh-Hant"/><c:pivotSource><c:name>Pivot &amp; One</c:name>
            <c:fmtId val="42"/></c:pivotSource>
        <c:chart><c:plotArea/></c:chart></c:chartSpace>"#;
    let chart = read(xml.as_slice()).unwrap();
    assert_eq!(chart.language.as_deref(), Some("zh-Hant"));
    let source = chart.pivot_source.as_ref().unwrap();
    assert_eq!(source.name, "Pivot & One");
    assert_eq!(source.format_id, 42);

    let mut output = Vec::new();
    crate::chart::writer::write(&mut output, &chart).unwrap();
    let reparsed = read(output.as_slice()).unwrap();
    assert_eq!(reparsed.language.as_deref(), Some("zh-Hant"));
    assert_eq!(reparsed.pivot_source.as_ref().unwrap().name, "Pivot & One");

    for invalid in [
        br#"<c:lang val="en-US"/><c:lang val="fr-FR"/>"#.as_slice(),
        br#"<c:pivotSource><c:fmtId val="1"/></c:pivotSource>"#.as_slice(),
        br#"<c:pivotSource><c:name>Pivot</c:name></c:pivotSource>"#.as_slice(),
        br#"<c:pivotSource/>"#.as_slice(),
    ] {
        let mut document =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">"#
                .to_vec();
        document.extend_from_slice(invalid);
        document.extend_from_slice(b"<c:chart><c:plotArea/></c:chart></c:chartSpace>");
        assert!(read(document.as_slice()).is_err());
    }
}

#[test]
fn round_trips_and_validates_chart_protection() {
    let xml = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:protection><c:chartObject/><c:data val="0"/><c:formatting val="true"/>
            <c:selection val="false"/><c:userInterface/></c:protection>
        <c:chart><c:plotArea/></c:chart></c:chartSpace>"#;
    let chart = read(xml.as_slice()).unwrap();
    let protection = chart.protection.as_ref().unwrap();
    assert_eq!(protection.chart_object, Some(true));
    assert_eq!(protection.data, Some(false));
    assert_eq!(protection.formatting, Some(true));
    assert_eq!(protection.selection, Some(false));
    assert_eq!(protection.user_interface, Some(true));

    let mut output = Vec::new();
    crate::chart::writer::write(&mut output, &chart).unwrap();
    let reparsed = read(output.as_slice()).unwrap();
    let protection = reparsed.protection.as_ref().unwrap();
    assert_eq!(protection.chart_object, Some(true));
    assert_eq!(protection.selection, Some(false));

    let empty = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:protection/><c:chart><c:plotArea/></c:chart></c:chartSpace>"#;
    let empty = read(empty.as_slice()).unwrap();
    let empty = empty.protection.as_ref().unwrap();
    assert_eq!(empty.chart_object, None);
    assert_eq!(empty.user_interface, None);

    for invalid in [
        br#"<c:protection><c:data/><c:data val="0"/></c:protection>"#.as_slice(),
        br#"<c:protection><c:selection val="maybe"/></c:protection>"#.as_slice(),
        br#"<c:protection/><c:protection/>"#.as_slice(),
        br#"<c:protection><c:data><c:data/></c:data></c:protection>"#.as_slice(),
    ] {
        let mut document =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">"#
                .to_vec();
        document.extend_from_slice(invalid);
        document.extend_from_slice(b"<c:chart><c:plotArea/></c:chart></c:chartSpace>");
        assert!(read(document.as_slice()).is_err());
    }
}

#[test]
fn round_trips_and_validates_chart_color_map_overrides() {
    let xml = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <c:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="accent1"
            tx2="accent2" accent1="accent3" accent2="accent4" accent3="accent5"
            accent4="accent6" accent5="hlink" accent6="folHlink" hlink="dk2"
            folHlink="lt2"/></c:clrMapOvr>
        <c:chart><c:plotArea/></c:chart></c:chartSpace>"#;
    let chart = read(xml.as_slice()).unwrap();
    let ColorMapOverride::Override(mapping) = chart.color_map_override.as_ref().unwrap() else {
        panic!("expected explicit chart color mapping");
    };
    assert_eq!(mapping.background1, ColorSchemeIndex::Dark1);
    assert_eq!(mapping.background2, ColorSchemeIndex::Accent1);
    assert_eq!(mapping.accent5, ColorSchemeIndex::Hyperlink);
    assert_eq!(mapping.followed_hyperlink, ColorSchemeIndex::Light2);

    let mut output = Vec::new();
    crate::chart::writer::write(&mut output, &chart).unwrap();
    let reparsed = read(output.as_slice()).unwrap();
    assert_eq!(reparsed.color_map_override, chart.color_map_override);

    let master = br#"<c:chartSpace xmlns:c="http://purl.oclc.org/ooxml/drawingml/chart"
            xmlns:d="http://purl.oclc.org/ooxml/drawingml/main">
        <c:clrMapOvr><d:masterClrMapping></d:masterClrMapping></c:clrMapOvr>
        <c:chart><c:plotArea/></c:chart></c:chartSpace>"#;
    assert_eq!(
        read(master.as_slice()).unwrap().color_map_override,
        Some(ColorMapOverride::Master)
    );

    for invalid in [
        br#"<c:clrMapOvr/>"#.as_slice(),
        br#"<c:clrMapOvr><a:masterClrMapping/><a:masterClrMapping/></c:clrMapOvr>"#.as_slice(),
        br#"<c:clrMapOvr><a:overrideClrMapping bg1="lt1"/></c:clrMapOvr>"#.as_slice(),
        br#"<c:clrMapOvr><a:overrideClrMapping bg1="none" tx1="lt1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></c:clrMapOvr>"#.as_slice(),
        br#"<c:clrMapOvr><c:masterClrMapping/></c:clrMapOvr>"#.as_slice(),
    ] {
        let mut document = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#.to_vec();
        document.extend_from_slice(invalid);
        document.extend_from_slice(b"<c:chart><c:plotArea/></c:chart></c:chartSpace>");
        assert!(read(document.as_slice()).is_err());
    }
}

#[test]
fn round_trips_and_validates_chart_external_data() {
    let xml = br#"<c:chartSpace
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <c:chart><c:plotArea/></c:chart>
        <c:externalData rel:id="rId7"><c:autoUpdate val="0"/></c:externalData>
    </c:chartSpace>"#;
    let chart = read(xml.as_slice()).unwrap();
    let external_data = chart.external_data.as_ref().unwrap();
    assert_eq!(external_data.relationship_id.as_deref(), Some("rId7"));
    assert_eq!(external_data.auto_update, Some(false));

    let mut output = Vec::new();
    crate::chart::writer::write(&mut output, &chart).unwrap();
    let reparsed = read(output.as_slice()).unwrap();
    assert_eq!(reparsed.external_data, chart.external_data);

    let implicit_true = br#"<c:chartSpace
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <c:chart></c:chart><c:externalData r:id="rId1"><c:autoUpdate/></c:externalData>
    </c:chartSpace>"#;
    assert_eq!(
        read(implicit_true.as_slice())
            .unwrap()
            .external_data
            .unwrap()
            .auto_update,
        Some(true)
    );

    for invalid in [
        br#"<c:externalData/>"#.as_slice(),
        br#"<c:externalData id="rId1"/>"#.as_slice(),
        br#"<c:externalData r:id="rId1"><c:autoUpdate/><c:autoUpdate/></c:externalData>"#
            .as_slice(),
        br#"<c:externalData r:id="rId1"><c:autoUpdate val="maybe"/></c:externalData>"#.as_slice(),
        br#"<c:externalData r:id="rId1"/><c:externalData r:id="rId2"/>"#.as_slice(),
    ] {
        let mut document = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart></c:chart>"#.to_vec();
        document.extend_from_slice(invalid);
        document.extend_from_slice(b"</c:chartSpace>");
        assert!(read(document.as_slice()).is_err());
    }

    let mut pending = Chart::new();
    pending.external_data = Some(ExternalData::pending());
    assert!(crate::chart::writer::write(&mut Vec::new(), &pending).is_err());
}

#[test]
fn round_trips_and_validates_chart_user_shapes_relationships() {
    let xml = br#"<c:chartSpace
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <c:chart><c:plotArea/></c:chart><c:userShapes q:id="shapeRel"/>
    </c:chartSpace>"#;
    let chart = read(xml.as_slice()).unwrap();
    assert_eq!(
        chart
            .user_shapes
            .as_ref()
            .unwrap()
            .relationship_id
            .as_deref(),
        Some("shapeRel")
    );
    let mut output = Vec::new();
    crate::chart::writer::write(&mut output, &chart).unwrap();
    assert_eq!(
        read(output.as_slice()).unwrap().user_shapes,
        chart.user_shapes
    );

    for invalid in [
        br#"<c:userShapes/>"#.as_slice(),
        br#"<c:userShapes r:id="one"><c:autoUpdate/></c:userShapes>"#.as_slice(),
        br#"<c:userShapes r:id="one"/><c:userShapes r:id="two"/>"#.as_slice(),
    ] {
        let mut document = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart></c:chart>"#.to_vec();
        document.extend_from_slice(invalid);
        document.extend_from_slice(b"</c:chartSpace>");
        assert!(read(document.as_slice()).is_err());
    }
}

#[test]
fn preserves_chart_space_drawing_and_extension_fragments() {
    let xml = br#"<c:chartSpace
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:x="urn:example:chart-extension">
        <c:chart><c:plotArea>
            <c:spPr><a:solidFill><a:srgbClr val="654321"/></a:solidFill></c:spPr>
            <c:extLst><c:ext uri="plot"><x:plotPayload/></c:ext></c:extLst>
        </c:plotArea><c:extLst><c:ext uri="chart"><x:chartPayload/></c:ext></c:extLst></c:chart>
        <c:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill></c:spPr>
        <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Label</a:t></a:r></a:p></c:txPr>
        <c:extLst><c:ext uri="example"><x:payload enabled="1"/></c:ext></c:extLst>
    </c:chartSpace>"#;
    let chart = read(xml.as_slice()).unwrap();
    let shape_properties = chart.shape_properties.as_ref().unwrap();
    assert!(
        std::str::from_utf8(shape_properties.as_xml())
            .unwrap()
            .contains("123456")
    );
    let extension_list = chart.extension_list.as_ref().unwrap();
    assert!(
        std::str::from_utf8(extension_list.as_xml())
            .unwrap()
            .contains(r#"xmlns:x="urn:example:chart-extension""#)
    );
    assert!(
        std::str::from_utf8(chart.plot_area.shape_properties.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("654321")
    );
    assert!(
        std::str::from_utf8(chart.chart_extension_list.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("chartPayload")
    );

    let mut output = Vec::new();
    crate::chart::writer::write(&mut output, &chart).unwrap();
    let reparsed = read(output.as_slice()).unwrap();
    assert_eq!(reparsed.shape_properties, chart.shape_properties);
    assert_eq!(reparsed.text_properties, chart.text_properties);
    assert_eq!(reparsed.extension_list, chart.extension_list);
    assert_eq!(
        reparsed.plot_area.shape_properties,
        chart.plot_area.shape_properties
    );
    assert_eq!(
        reparsed.plot_area.extension_list,
        chart.plot_area.extension_list
    );
    assert_eq!(reparsed.chart_extension_list, chart.chart_extension_list);

    assert!(
        ShapeProperties::from_xml(
            br#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#
                .to_vec()
        )
        .is_err()
    );
    assert!(
        ExtensionList::from_xml(
            br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/><c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#
                .to_vec()
        )
        .is_err()
    );
}

#[test]
fn round_trips_chart_surface_shape_and_picture_options() {
    let xml = br#"<c:chartSpace
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:x="urn:example:surface">
        <c:chart><c:floor>
            <c:thickness val="64"/>
            <c:spPr><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></c:spPr>
            <c:pictureOptions>
                <c:applyToFront val="0"/><c:applyToSides/><c:applyToEnd val="1"/>
                <c:pictureFormat val="stackScale"/><c:pictureStackUnit val="2.5"/>
            </c:pictureOptions>
            <c:extLst><c:ext uri="surface"><x:payload/></c:ext></c:extLst>
        </c:floor><c:plotArea/></c:chart>
    </c:chartSpace>"#;
    let chart = read(xml.as_slice()).unwrap();
    let floor = chart.floor.as_ref().unwrap();
    assert_eq!(floor.thickness, Some(64));
    assert!(
        std::str::from_utf8(floor.shape_properties.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("ABCDEF")
    );
    let options = floor.picture_options.as_ref().unwrap();
    assert_eq!(options.apply_to_front, Some(false));
    assert_eq!(options.apply_to_sides, Some(true));
    assert_eq!(options.apply_to_end, Some(true));
    assert_eq!(options.picture_format, Some(PictureFormat::StackScale));
    assert_eq!(options.picture_stack_unit, Some(2.5));
    assert!(
        std::str::from_utf8(floor.extension_list.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("urn:example:surface")
    );

    let mut output = Vec::new();
    crate::chart::writer::write(&mut output, &chart).unwrap();
    let reparsed = read(output.as_slice()).unwrap();
    let reparsed_floor = reparsed.floor.unwrap();
    assert_eq!(reparsed_floor.thickness, floor.thickness);
    assert_eq!(reparsed_floor.shape_properties, floor.shape_properties);
    assert_eq!(reparsed_floor.picture_options, floor.picture_options);
    assert_eq!(reparsed_floor.extension_list, floor.extension_list);

    let invalid = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:floor><c:pictureOptions><c:pictureFormat val="tile"/></c:pictureOptions></c:floor><c:plotArea/></c:chart></c:chartSpace>"#;
    assert!(read(invalid.as_slice()).is_err());

    let empty = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:view3D/><c:floor/><c:backWall/><c:sideWall/><c:plotArea/></c:chart></c:chartSpace>"#;
    let empty_chart = read(empty.as_slice()).unwrap();
    assert!(empty_chart.view_3d.is_some());
    assert!(empty_chart.floor.is_some());
    assert!(empty_chart.back_wall.is_some());
    assert!(empty_chart.side_wall.is_some());
}

#[test]
fn round_trips_series_and_data_point_formatting_fragments() {
    let xml = br#"<c:chartSpace
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:x="urn:example:series">
        <c:chart><c:plotArea><c:barChart>
            <c:barDir val="col"/><c:grouping val="clustered"/><c:ser>
                <c:idx val="0"/><c:order val="0"/>
                <c:spPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></c:spPr>
                <c:invertIfNegative val="1"/>
                <c:pictureOptions><c:applyToFront val="0"/>
                    <c:pictureFormat val="stack"/><c:pictureStackUnit val="2"/>
                </c:pictureOptions>
                <c:dPt><c:idx val="2"/>
                    <c:spPr><a:solidFill><a:srgbClr val="AABBCC"/></a:solidFill></c:spPr>
                    <c:pictureOptions><c:applyToSides/></c:pictureOptions>
                    <c:extLst><c:ext uri="point"><x:pointPayload/></c:ext></c:extLst>
                </c:dPt>
                <c:shape val="cylinder"/>
                <c:extLst><c:ext uri="series"><x:seriesPayload/></c:ext></c:extLst>
            </c:ser><c:axId val="1"/><c:axId val="2"/>
        </c:barChart></c:plotArea></c:chart>
    </c:chartSpace>"#;
    let chart = read(xml.as_slice()).unwrap();
    let TypeGroup::Bar(group) = &chart.plot_area.type_groups[0] else {
        panic!("expected a bar chart");
    };
    let series = &group.common.series[0];
    assert!(
        std::str::from_utf8(series.shape_properties.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("112233")
    );
    assert_eq!(
        series.picture_options.as_ref().unwrap().picture_format,
        Some(PictureFormat::Stack)
    );
    assert_eq!(
        series.picture_options.as_ref().unwrap().picture_stack_unit,
        Some(2.0)
    );
    assert_eq!(series.bar_shape, Some(BarShape::Cylinder));
    assert!(
        std::str::from_utf8(series.extension_list.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("seriesPayload")
    );
    let point = &series.data_points[0];
    assert_eq!(
        point.picture_options.as_ref().unwrap().apply_to_sides,
        Some(true)
    );
    assert!(
        std::str::from_utf8(point.extension_list.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("pointPayload")
    );

    let mut output = Vec::new();
    crate::chart::writer::write(&mut output, &chart).unwrap();
    let reparsed = read(output.as_slice()).unwrap();
    let TypeGroup::Bar(reparsed_group) = &reparsed.plot_area.type_groups[0] else {
        panic!("expected a bar chart");
    };
    let reparsed_series = &reparsed_group.common.series[0];
    assert_eq!(reparsed_series.shape_properties, series.shape_properties);
    assert_eq!(reparsed_series.picture_options, series.picture_options);
    assert_eq!(reparsed_series.bar_shape, series.bar_shape);
    assert_eq!(reparsed_series.extension_list, series.extension_list);
    assert_eq!(
        reparsed_series.data_points[0].shape_properties,
        point.shape_properties
    );
    assert_eq!(
        reparsed_series.data_points[0].picture_options,
        point.picture_options
    );
    assert_eq!(
        reparsed_series.data_points[0].extension_list,
        point.extension_list
    );

    let unsupported =
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:plotArea><c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/>
            <c:pictureOptions/></c:ser></c:lineChart></c:plotArea></c:chart>
        </c:chartSpace>"#;
    let unsupported = read(unsupported.as_slice()).unwrap();
    assert!(crate::chart::writer::write(&mut Vec::new(), &unsupported).is_err());

    let unsupported_shape =
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:plotArea><c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/>
            <c:shape val="cone"/></c:ser></c:lineChart></c:plotArea></c:chart>
        </c:chartSpace>"#;
    let unsupported_shape = read(unsupported_shape.as_slice()).unwrap();
    assert!(crate::chart::writer::write(&mut Vec::new(), &unsupported_shape).is_err());

    for invalid_shape in [
        br#"<c:shape val="sphere"/>"#.as_slice(),
        br#"<c:shape/><c:shape/>"#.as_slice(),
    ] {
        let mut document = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/>"#.to_vec();
        document.extend_from_slice(invalid_shape);
        document.extend_from_slice(b"</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>");
        assert!(read(document.as_slice()).is_err());
    }

    for invalid_unit in ["0", "-1"] {
        let document = format!(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/><c:pictureOptions><c:pictureStackUnit val="{invalid_unit}"/></c:pictureOptions></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#
        );
        assert!(read(document.as_bytes()).is_err());
    }

    let mut invalid = Chart::new();
    let mut bar = BarTypeGroup::new(BarDirection::Column, BarGrouping::Clustered);
    let mut series = Series::new(0);
    series.picture_options = Some(PictureOptions {
        picture_stack_unit: Some(0.0),
        ..PictureOptions::default()
    });
    bar.common.series.push(series);
    invalid.plot_area.type_groups.push(TypeGroup::Bar(bar));
    assert!(crate::chart::writer::write(&mut Vec::new(), &invalid).is_err());

    for supported in [
        br#"<c:areaChart><c:grouping val="standard"/><c:ser><c:idx val="0"/><c:order val="0"/><c:pictureOptions/></c:ser></c:areaChart>"#.as_slice(),
        br#"<c:area3DChart><c:grouping val="standard"/><c:ser><c:idx val="0"/><c:order val="0"/><c:pictureOptions/></c:ser></c:area3DChart>"#.as_slice(),
        br#"<c:bubbleChart><c:ser><c:idx val="0"/><c:order val="0"/><c:invertIfNegative/></c:ser></c:bubbleChart>"#.as_slice(),
    ] {
        let mut document = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea>"#.to_vec();
        document.extend_from_slice(supported);
        document.extend_from_slice(b"</c:plotArea></c:chart></c:chartSpace>");
        let supported = read(document.as_slice()).unwrap();
        crate::chart::writer::write(&mut Vec::new(), &supported).unwrap();
    }
}

#[test]
fn round_trips_and_validates_chart_print_settings() {
    let xml = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:plotArea/></c:chart><c:printSettings>
            <c:headerFooter alignWithMargins="0" differentOddEven="1" differentFirst="true">
                <c:oddHeader>&amp;LRevenue</c:oddHeader><c:oddFooter>&amp;P / &amp;N</c:oddFooter>
                <c:evenHeader/><c:firstFooter><![CDATA[First & last]]></c:firstFooter>
            </c:headerFooter>
            <c:pageMargins l="0.2" r="0.3" t="0.4" b="0.5" header="0.1" footer="0.15"/>
            <c:pageSetup paperSize="9" firstPageNumber="4" orientation="landscape"
                blackAndWhite="1" draft="true" useFirstPageNumber="1"
                horizontalDpi="300" verticalDpi="1200" copies="2"/>
        </c:printSettings></c:chartSpace>"#;
    let chart = read(xml.as_slice()).unwrap();
    let settings = chart.print_settings.as_ref().unwrap();
    let header_footer = settings.header_footer.as_ref().unwrap();
    assert!(!header_footer.align_with_margins);
    assert!(header_footer.different_odd_even);
    assert!(header_footer.different_first);
    assert_eq!(header_footer.odd_header.as_deref(), Some("&LRevenue"));
    assert_eq!(header_footer.even_header.as_deref(), Some(""));
    assert_eq!(header_footer.first_footer.as_deref(), Some("First & last"));
    let margins = settings.page_margins.as_ref().unwrap();
    assert_eq!(margins.left, 0.2);
    assert_eq!(margins.footer, 0.15);
    let setup = settings.page_setup.as_ref().unwrap();
    assert_eq!(setup.paper_size, 9);
    assert_eq!(setup.first_page_number, 4);
    assert_eq!(setup.orientation, PageOrientation::Landscape);
    assert!(setup.black_and_white);
    assert!(setup.draft);
    assert!(setup.use_first_page_number);
    assert_eq!(setup.horizontal_dpi, 300);
    assert_eq!(setup.vertical_dpi, 1200);
    assert_eq!(setup.copies, 2);

    let mut output = Vec::new();
    crate::chart::writer::write(&mut output, &chart).unwrap();
    let reparsed = read(output.as_slice()).unwrap();
    assert_eq!(
        reparsed
            .print_settings
            .as_ref()
            .unwrap()
            .header_footer
            .as_ref()
            .unwrap()
            .odd_footer
            .as_deref(),
        Some("&P / &N")
    );

    let empty = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:plotArea/></c:chart><c:printSettings/></c:chartSpace>"#;
    let empty = read(empty.as_slice()).unwrap();
    let empty = empty.print_settings.unwrap();
    assert!(empty.header_footer.is_none());
    assert!(empty.page_margins.is_none());
    assert!(empty.page_setup.is_none());

    for invalid in [
        br#"<c:pageMargins l="0.2" r="0.3" t="0.4" b="0.5" header="0.1"/>"#.as_slice(),
        br#"<c:pageSetup orientation="diagonal"/>"#.as_slice(),
        br#"<c:pageSetup/><c:pageSetup/>"#.as_slice(),
        br#"<c:headerFooter><c:bogus/></c:headerFooter>"#.as_slice(),
    ] {
        let mut document = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea/></c:chart><c:printSettings>"#.to_vec();
        document.extend_from_slice(invalid);
        document.extend_from_slice(b"</c:printSettings></c:chartSpace>");
        assert!(read(document.as_slice()).is_err());
    }

    let mut invalid = Chart::new();
    let mut settings = PrintSettings::new();
    settings.page_margins = Some(PageMargins::new(f64::NAN, 0.3, 0.4, 0.5, 0.1, 0.15));
    invalid.print_settings = Some(settings);
    assert!(crate::chart::writer::write(&mut Vec::new(), &invalid).is_err());
}

#[test]
fn round_trips_and_validates_pivot_chart_formats() {
    let xml =
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:x="urn:example:pivot-format">
        <c:chart><c:pivotFmts>
            <c:pivotFmt><c:idx val="2"/><c:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill></c:spPr>
                <c:txPr><a:bodyPr rot="600000"/><a:lstStyle/><a:p/></c:txPr><c:marker>
                <c:symbol val="diamond"/><c:size val="8"/><c:spPr><a:ln w="25400"/></c:spPr>
                <c:extLst><c:ext uri="marker"><x:markerPayload/></c:ext></c:extLst>
            </c:marker><c:dLbl><c:idx val="2"/><c:showVal val="1"/></c:dLbl>
                <c:extLst><c:ext uri="pivot"><x:payload/></c:ext></c:extLst></c:pivotFmt>
            <c:pivotFmt><c:idx val="7"/><c:marker/></c:pivotFmt>
        </c:pivotFmts><c:plotArea/></c:chart></c:chartSpace>"#;
    let chart = read(xml.as_slice()).unwrap();
    let formats = chart.pivot_formats.as_ref().unwrap();
    assert_eq!(formats.len(), 2);
    assert_eq!(formats[0].index, 2);
    let marker = formats[0].marker.as_ref().unwrap();
    assert_eq!(marker.symbol, Some(MarkerStyle::Diamond));
    assert_eq!(marker.size, Some(8));
    assert!(
        std::str::from_utf8(marker.shape_properties.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("25400")
    );
    assert!(marker.extension_list.is_some());
    assert!(formats[0].data_label.as_ref().unwrap().show_value);
    assert!(
        std::str::from_utf8(formats[0].shape_properties.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("123456")
    );
    assert!(
        std::str::from_utf8(formats[0].text_properties.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("600000")
    );
    assert!(
        std::str::from_utf8(formats[0].extension_list.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("payload")
    );
    assert!(formats[1].marker.is_some());

    let mut output = Vec::new();
    crate::chart::writer::write(&mut output, &chart).unwrap();
    let reparsed = read(output.as_slice()).unwrap();
    assert_eq!(reparsed.pivot_formats.as_ref().unwrap().len(), 2);
    let reparsed_format = &reparsed.pivot_formats.as_ref().unwrap()[0];
    assert_eq!(
        reparsed_format.shape_properties,
        formats[0].shape_properties
    );
    assert_eq!(reparsed_format.text_properties, formats[0].text_properties);
    assert_eq!(reparsed_format.extension_list, formats[0].extension_list);

    let empty = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:pivotFmts/><c:plotArea/></c:chart></c:chartSpace>"#;
    assert!(
        read(empty.as_slice())
            .unwrap()
            .pivot_formats
            .unwrap()
            .is_empty()
    );

    let duplicate =
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:pivotFmts><c:pivotFmt><c:idx val="1"/></c:pivotFmt>
            <c:pivotFmt><c:idx val="1"/></c:pivotFmt></c:pivotFmts>
            <c:plotArea/></c:chart></c:chartSpace>"#;
    assert!(read(duplicate.as_slice()).is_err());

    let duplicate_shape = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:pivotFmts><c:pivotFmt><c:idx val="1"/><c:spPr/><c:spPr/></c:pivotFmt></c:pivotFmts><c:plotArea/></c:chart></c:chartSpace>"#;
    assert!(read(duplicate_shape.as_slice()).is_err());

    let mut invalid = Chart::new();
    invalid.pivot_formats = Some(vec![PivotFormat::new(3), PivotFormat::new(3)]);
    assert!(crate::chart::writer::write(&mut Vec::new(), &invalid).is_err());

    let mut invalid = Chart::new();
    let mut format = PivotFormat::new(3);
    format.marker = Some(Marker {
        symbol: None,
        size: Some(73),
        ..Marker::default()
    });
    invalid.pivot_formats = Some(vec![format]);
    assert!(crate::chart::writer::write(&mut Vec::new(), &invalid).is_err());
}

#[test]
fn preserves_explicit_empty_series_and_point_markers() {
    let xml = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:plotArea><c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/>
            <c:marker/><c:dPt><c:idx val="0"/><c:marker/></c:dPt>
        </c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>"#;
    let chart = read(xml.as_slice()).unwrap();
    let TypeGroup::Line(group) = &chart.plot_area.type_groups[0] else {
        panic!("expected line chart");
    };
    let series = &group.common.series[0];
    assert!(series.marker_present);
    assert!(series.data_points[0].marker_present);

    let mut output = Vec::new();
    crate::chart::writer::write(&mut output, &chart).unwrap();
    let reparsed = read(output.as_slice()).unwrap();
    let TypeGroup::Line(group) = &reparsed.plot_area.type_groups[0] else {
        panic!("expected line chart");
    };
    assert!(group.common.series[0].marker_present);
    assert!(group.common.series[0].data_points[0].marker_present);
}

#[test]
fn parses_prefixed_chart_content() {
    let xml =
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue &amp; <![CDATA[Growth]]></a:t></a:r></a:p>
            </c:rich></c:tx></c:title><c:plotArea><c:barChart>
            <c:barDir val="bar"/><c:grouping val="stacked"/><c:ser>
                <c:idx val="2"/><c:order val="1"/>
                <c:cat><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache><c:pt idx="0"><c:v>East</c:v></c:pt>
                    </c:strCache></c:strRef></c:cat>
                <c:val><c:numRef><c:f>Sheet1!$B$1</c:f><c:numCache><c:formatCode>0.0</c:formatCode><c:pt idx="0"><c:v>42.5</c:v></c:pt>
                    </c:numCache></c:numRef></c:val>
                <c:dLbls><c:dLbl><c:idx val="0"/><c:delete val="1"/></c:dLbl>
                    <c:showVal val="1"/></c:dLbls>
            </c:ser></c:barChart></c:plotArea>
            <c:legend><c:legendPos val="b"/><c:overlay val="1"/></c:legend>
            <c:showDLblsOverMax val="1"/>
        </c:chart><c:style val="12"/></c:chartSpace>"#;

    let chart = read(xml.as_slice()).unwrap();
    let Some(TitleText::Literal(title)) = chart.title.as_ref() else {
        panic!("expected a literal chart title");
    };
    assert_eq!(title.text, "Revenue & Growth");
    assert_eq!(chart.plot_area.type_groups.len(), 1);
    let TypeGroup::Bar(group) = &chart.plot_area.type_groups[0] else {
        panic!("expected a bar chart");
    };
    assert_eq!(group.direction, BarDirection::Bar);
    assert_eq!(group.grouping, BarGrouping::Stacked);
    assert_eq!(group.common.series.len(), 1);
    assert_eq!(group.common.series[0].index, 2);
    assert_eq!(
        group.common.series[0].categories.as_ref().unwrap().values,
        ["East"]
    );
    assert_eq!(
        group.common.series[0]
            .categories
            .as_ref()
            .unwrap()
            .source_ref
            .as_ref()
            .unwrap()
            .formula,
        "Sheet1!$A$1"
    );
    assert_eq!(
        group.common.series[0].values.as_ref().unwrap().values,
        [42.5]
    );
    let values = group.common.series[0].values.as_ref().unwrap();
    assert_eq!(values.source_ref.as_ref().unwrap().formula, "Sheet1!$B$1");
    assert_eq!(values.format_code.as_deref(), Some("0.0"));
    let labels = group.common.series[0].data_labels.as_ref().unwrap();
    assert!(labels.show_value);
    assert!(!labels.deleted);
    assert_eq!(labels.labels.len(), 1);
    assert_eq!(labels.labels[0].index, 0);
    assert!(labels.labels[0].deleted);
    assert_eq!(chart.legend.unwrap().position, LegendPosition::Bottom);
    assert!(chart.show_data_labels_over_max);
    assert_eq!(chart.style, Some(12));
}

#[test]
fn parses_strict_chart_and_drawingml_namespaces() {
    let xml = br#"<c:chartSpace xmlns:c="http://purl.oclc.org/ooxml/drawingml/chart"
            xmlns:a="http://purl.oclc.org/ooxml/drawingml/main">
        <c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Strict title</a:t></a:r></a:p>
            </c:rich></c:tx></c:title><c:plotArea><c:pieChart></c:pieChart></c:plotArea></c:chart>
        <c:style val="7"/>
    </c:chartSpace>"#;

    let chart = read(xml.as_slice()).unwrap();
    let Some(TitleText::Literal(title)) = chart.title else {
        panic!("expected a literal chart title");
    };
    assert_eq!(title.text, "Strict title");
    assert_eq!(chart.style, Some(7));
    assert!(matches!(
        chart.plot_area.type_groups.as_slice(),
        [TypeGroup::Pie(_)]
    ));
}

#[test]
fn ignores_foreign_namespace_lookalikes_and_their_descendants() {
    let xml = br#"<c:chartSpace
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:x="urn:example:chart-extension">
        <x:style val="48"/>
        <c:chart><c:title><c:tx><c:strRef>
            <c:f>Sheet<x:payload><c:style val="47"/><c:masterClrMapping/>ignored</x:payload>1!$A$1</c:f>
        </c:strRef></c:tx></c:title><c:plotArea>
            <x:barChart><c:barChart><c:ser><c:idx val="9"/></c:ser></c:barChart></x:barChart>
            <c:lineChart></c:lineChart>
        </c:plotArea></c:chart>
        <c:style val="4"/>
    </c:chartSpace>"#;

    let chart = read(xml.as_slice()).unwrap();
    let Some(TitleText::Reference(title)) = chart.title else {
        panic!("expected a chart title reference");
    };
    assert_eq!(title.formula, "Sheet1!$A$1");
    assert_eq!(chart.style, Some(4));
    assert!(matches!(
        chart.plot_area.type_groups.as_slice(),
        [TypeGroup::Line(_)]
    ));
}

#[test]
fn rejects_non_chart_roots_and_trailing_roots() {
    let chart_namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    let foreign_root = br#"<x:chartSpace xmlns:x="urn:example"><x:chart/></x:chartSpace>"#;
    let trailing_root = format!(
        r#"<c:chartSpace xmlns:c="{chart_namespace}"><c:chart/></c:chartSpace><c:chartSpace xmlns:c="{chart_namespace}"/>"#
    );

    assert!(read(foreign_root.as_slice()).is_err());
    assert!(read(trailing_root.as_bytes()).is_err());
}

#[test]
fn preserves_automatic_display_units_labels() {
    let xml = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:plotArea><c:valAx><c:axId val="1"/><c:scaling/>
            <c:axPos val="l"/><c:crossAx val="2"/><c:dispUnits>
                <c:builtInUnit val="thousands"/><c:dispUnitsLbl/>
            </c:dispUnits>
        </c:valAx></c:plotArea></c:chart>
    </c:chartSpace>"#;

    let chart = read(xml.as_slice()).unwrap();
    let Axis::Value(axis) = &chart.plot_area.axes[0] else {
        panic!("expected value axis");
    };
    let display_units = axis.display_units.as_ref().unwrap();
    assert_eq!(display_units.built_in_unit, Some(BuiltInUnit::Thousands));
    assert!(display_units.show_label);
    assert!(display_units.label.is_none());
    assert!(display_units.layout.is_none());

    let mut output = Vec::new();
    crate::chart::writer::write(&mut output, &chart).unwrap();
    assert!(
        std::str::from_utf8(&output)
            .unwrap()
            .contains("<c:dispUnitsLbl>")
    );
}

#[test]
fn round_trips_display_units_label_formatting_and_extensions() {
    let xml = br#"<c:chartSpace
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:x="urn:example:display-units">
        <c:chart><c:plotArea><c:valAx><c:axId val="1"/><c:scaling/>
            <c:axPos val="l"/><c:crossAx val="2"/><c:dispUnits>
                <c:builtInUnit val="millions"/><c:dispUnitsLbl>
                    <c:layout/><c:spPr><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></c:spPr>
                    <c:txPr><a:bodyPr rot="600000"/><a:lstStyle/><a:p/></c:txPr>
                </c:dispUnitsLbl>
                <c:extLst><c:ext uri="display-units"><x:payload/></c:ext></c:extLst>
            </c:dispUnits>
        </c:valAx></c:plotArea></c:chart>
    </c:chartSpace>"#;

    let chart = read(xml.as_slice()).unwrap();
    let Axis::Value(axis) = &chart.plot_area.axes[0] else {
        panic!("expected value axis");
    };
    let display_units = axis.display_units.as_ref().unwrap();
    assert!(display_units.show_label);
    assert!(
        std::str::from_utf8(
            display_units
                .label_shape_properties
                .as_ref()
                .unwrap()
                .as_xml()
        )
        .unwrap()
        .contains("ABCDEF")
    );
    assert!(
        std::str::from_utf8(
            display_units
                .label_text_properties
                .as_ref()
                .unwrap()
                .as_xml()
        )
        .unwrap()
        .contains("600000")
    );
    assert!(
        std::str::from_utf8(display_units.extension_list.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("payload")
    );

    let mut output = Vec::new();
    crate::chart::writer::write(&mut output, &chart).unwrap();
    let reparsed = read(output.as_slice()).unwrap();
    let Axis::Value(reparsed_axis) = &reparsed.plot_area.axes[0] else {
        panic!("expected value axis");
    };
    let reparsed_units = reparsed_axis.display_units.as_ref().unwrap();
    assert_eq!(
        reparsed_units.label_shape_properties,
        display_units.label_shape_properties
    );
    assert_eq!(
        reparsed_units.label_text_properties,
        display_units.label_text_properties
    );
    assert_eq!(reparsed_units.extension_list, display_units.extension_list);

    let duplicate = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:valAx><c:axId val="1"/><c:scaling/><c:axPos val="l"/><c:crossAx val="2"/><c:dispUnits><c:builtInUnit val="millions"/><c:dispUnitsLbl><c:spPr/><c:spPr/></c:dispUnitsLbl></c:dispUnits></c:valAx></c:plotArea></c:chart></c:chartSpace>"#;
    assert!(read(duplicate.as_slice()).is_err());
}

#[test]
fn parses_and_validates_chart_data_tables() {
    let xml = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:x="urn:example:data-table">
        <c:chart><c:plotArea><c:lineChart></c:lineChart>
            <c:dTable><c:showHorzBorder/><c:showVertBorder val="0"/>
                <c:showOutline val="true"/><c:showKeys val="false"/>
                <c:spPr><a:solidFill><a:srgbClr val="F0E0D0"/></a:solidFill></c:spPr>
                <c:txPr><a:bodyPr/><a:lstStyle/><a:p/></c:txPr>
                <c:extLst><c:ext uri="table"><x:payload/></c:ext></c:extLst>
            </c:dTable>
        </c:plotArea></c:chart>
    </c:chartSpace>"#;
    let chart = read(xml.as_slice()).unwrap();
    let table = chart.plot_area.data_table.as_ref().unwrap();
    assert!(table.show_horizontal_border);
    assert!(!table.show_vertical_border);
    assert!(table.show_outline);
    assert!(!table.show_legend_keys);
    assert!(
        std::str::from_utf8(table.shape_properties.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("F0E0D0")
    );
    assert!(table.text_properties.is_some());
    assert!(
        std::str::from_utf8(table.extension_list.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("urn:example:data-table")
    );

    let mut output = Vec::new();
    crate::chart::writer::write(&mut output, &chart).unwrap();
    let reparsed = read(output.as_slice()).unwrap();
    let reparsed = reparsed.plot_area.data_table.unwrap();
    assert_eq!(reparsed.shape_properties, table.shape_properties);
    assert_eq!(reparsed.text_properties, table.text_properties);
    assert_eq!(reparsed.extension_list, table.extension_list);

    let duplicate =
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:plotArea><c:lineChart></c:lineChart><c:dTable>
            <c:showKeys/><c:showKeys val="0"/>
        </c:dTable></c:plotArea></c:chart>
    </c:chartSpace>"#;
    assert!(read(duplicate.as_slice()).is_err());
}

#[test]
fn parses_and_validates_chart_group_data_labels() {
    let xml = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:plotArea>
            <c:lineChart><c:dLbls><c:numFmt formatCode="0.0%" sourceLinked="0"/>
                <c:dLblPos val="r"/><c:showVal/><c:showCatName val="true"/>
                <c:separator> / </c:separator><c:showLeaderLines/>
            </c:dLbls></c:lineChart>
            <c:area3DChart><c:dLbls/></c:area3DChart>
        </c:plotArea></c:chart>
    </c:chartSpace>"#;
    let chart = read(xml.as_slice()).unwrap();
    let TypeGroup::Line(line) = &chart.plot_area.type_groups[0] else {
        panic!("expected line chart");
    };
    let labels = line.common.data_labels.as_ref().unwrap();
    assert_eq!(labels.position, Some(DataLabelPosition::Right));
    assert!(labels.show_value);
    assert!(labels.show_category_name);
    assert!(labels.show_leader_lines);
    assert_eq!(labels.separator.as_deref(), Some(" / "));
    assert_eq!(labels.number_format.as_ref().unwrap().format_code, "0.0%");
    assert!(!labels.number_format.as_ref().unwrap().source_linked);
    let TypeGroup::Area3D(area) = &chart.plot_area.type_groups[1] else {
        panic!("expected 3D area chart");
    };
    assert!(area.common.data_labels.is_some());

    let duplicate =
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:plotArea><c:barChart><c:dLbls/><c:dLbls/>
        </c:barChart></c:plotArea></c:chart></c:chartSpace>"#;
    assert!(read(duplicate.as_slice()).is_err());
}

#[test]
fn preserves_and_validates_chart_group_axis_bindings() {
    let xml = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:plotArea>
            <c:lineChart><c:axId val="17"/><c:axId val="29"/></c:lineChart>
            <c:area3DChart><c:axId val="31"/><c:axId val="37"/><c:axId val="41"/></c:area3DChart>
        </c:plotArea></c:chart>
    </c:chartSpace>"#;
    let chart = read(xml.as_slice()).unwrap();
    let TypeGroup::Line(line) = &chart.plot_area.type_groups[0] else {
        panic!("expected line chart");
    };
    assert_eq!(line.common.axis_ids, [17, 29]);
    let TypeGroup::Area3D(area) = &chart.plot_area.type_groups[1] else {
        panic!("expected 3D area chart");
    };
    assert_eq!(area.common.axis_ids, [31, 37, 41]);

    for axis_ids in [vec![7], vec![7, 7], vec![1, 2, 3]] {
        let mut chart = Chart::new();
        let mut line = LineTypeGroup::new(BarGrouping::Standard);
        line.common.axis_ids = axis_ids;
        chart.plot_area.type_groups.push(TypeGroup::Line(line));
        assert!(crate::chart::writer::write(&mut Vec::new(), &chart).is_err());
    }
}

#[test]
fn round_trips_chart_lines_and_up_down_bars() {
    let xml =
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:x="urn:example:up-down-bars">
        <c:chart><c:plotArea>
            <c:areaChart><c:dropLines/></c:areaChart>
            <c:area3DChart><c:dropLines/></c:area3DChart>
            <c:barChart><c:serLines><c:spPr><a:solidFill><a:srgbClr val="333333"/></a:solidFill></c:spPr></c:serLines><c:serLines/></c:barChart>
            <c:lineChart><c:dropLines><c:spPr><a:solidFill><a:srgbClr val="111111"/></a:solidFill></c:spPr></c:dropLines><c:hiLowLines/><c:upDownBars>
                <c:gapWidth val="225%"/><c:upBars><c:spPr><a:solidFill><a:srgbClr val="222222"/></a:solidFill></c:spPr></c:upBars><c:downBars/>
                <c:extLst><c:ext uri="bars"><x:payload/></c:ext></c:extLst>
            </c:upDownBars></c:lineChart>
            <c:line3DChart><c:dropLines/></c:line3DChart>
            <c:ofPieChart><c:ofPieType val="bar"/><c:serLines/></c:ofPieChart>
            <c:stockChart><c:dropLines/><c:hiLowLines/><c:upDownBars/></c:stockChart>
        </c:plotArea></c:chart>
    </c:chartSpace>"#;
    let chart = read(xml.as_slice()).unwrap();
    let TypeGroup::Area(area) = &chart.plot_area.type_groups[0] else {
        panic!("expected area chart");
    };
    assert!(area.drop_lines.is_some());
    let TypeGroup::Area3D(area) = &chart.plot_area.type_groups[1] else {
        panic!("expected 3D area chart");
    };
    assert!(area.drop_lines.is_some());
    let TypeGroup::Bar(bar) = &chart.plot_area.type_groups[2] else {
        panic!("expected bar chart");
    };
    assert_eq!(bar.series_lines.len(), 2);
    assert!(
        std::str::from_utf8(
            bar.series_lines[0]
                .shape_properties
                .as_ref()
                .unwrap()
                .as_xml()
        )
        .unwrap()
        .contains("333333")
    );
    let TypeGroup::Line(line) = &chart.plot_area.type_groups[3] else {
        panic!("expected line chart");
    };
    assert!(line.drop_lines.is_some());
    assert!(line.high_low_lines.is_some());
    assert!(
        std::str::from_utf8(
            line.drop_lines
                .as_ref()
                .unwrap()
                .shape_properties
                .as_ref()
                .unwrap()
                .as_xml()
        )
        .unwrap()
        .contains("111111")
    );
    let bars = line.up_down_bars.as_ref().unwrap();
    assert_eq!(bars.gap_width, Some(225));
    assert!(bars.up_bars.is_some());
    assert!(bars.down_bars.is_some());
    assert!(
        std::str::from_utf8(
            bars.up_bars
                .as_ref()
                .unwrap()
                .shape_properties
                .as_ref()
                .unwrap()
                .as_xml()
        )
        .unwrap()
        .contains("222222")
    );
    assert!(
        std::str::from_utf8(bars.extension_list.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("urn:example:up-down-bars")
    );
    let TypeGroup::Line3D(line) = &chart.plot_area.type_groups[4] else {
        panic!("expected 3D line chart");
    };
    assert!(line.drop_lines.is_some());
    let TypeGroup::OfPie(of_pie) = &chart.plot_area.type_groups[5] else {
        panic!("expected of-pie chart");
    };
    assert_eq!(of_pie.series_lines.len(), 1);
    let TypeGroup::Stock(stock) = &chart.plot_area.type_groups[6] else {
        panic!("expected stock chart");
    };
    assert!(stock.drop_lines.is_some());
    assert!(stock.high_low_lines.is_some());
    assert!(stock.up_down_bars.is_some());

    let mut output = Vec::new();
    crate::chart::writer::write(&mut output, &chart).unwrap();
    let reparsed = read(output.as_slice()).unwrap();
    let TypeGroup::Line(line) = &reparsed.plot_area.type_groups[3] else {
        panic!("expected line chart");
    };
    assert_eq!(line.up_down_bars.as_ref().unwrap().gap_width, Some(225));
    assert_eq!(
        line.up_down_bars.as_ref().unwrap().extension_list,
        bars.extension_list
    );
    let TypeGroup::Bar(bar) = &reparsed.plot_area.type_groups[2] else {
        panic!("expected bar chart");
    };
    assert_eq!(bar.series_lines.len(), 2);
    assert!(
        std::str::from_utf8(
            bar.series_lines[0]
                .shape_properties
                .as_ref()
                .unwrap()
                .as_xml()
        )
        .unwrap()
        .contains("333333")
    );

    let duplicate =
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:plotArea><c:lineChart><c:dropLines/><c:dropLines/>
        </c:lineChart></c:plotArea></c:chart></c:chartSpace>"#;
    assert!(read(duplicate.as_slice()).is_err());

    let duplicate_formatting =
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:plotArea><c:lineChart><c:dropLines><c:spPr/><c:spPr/></c:dropLines>
        </c:lineChart></c:plotArea></c:chart></c:chartSpace>"#;
    assert!(read(duplicate_formatting.as_slice()).is_err());

    let mut invalid = Chart::new();
    let mut line = LineTypeGroup::new(BarGrouping::Standard);
    line.up_down_bars = Some(UpDownBars {
        gap_width: Some(501),
        ..UpDownBars::default()
    });
    invalid.plot_area.type_groups.push(TypeGroup::Line(line));
    assert!(crate::chart::writer::write(&mut Vec::new(), &invalid).is_err());
}

#[test]
fn round_trips_and_validates_surface_band_formats() {
    let xml =
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <c:chart><c:plotArea>
            <c:surfaceChart><c:wireframe/><c:bandFmts>
                <c:bandFmt><c:idx val="2"/><c:spPr><a:solidFill><a:srgbClr val="102030"/></a:solidFill></c:spPr></c:bandFmt>
                <c:bandFmt><c:idx val="7"/><c:spPr/></c:bandFmt>
            </c:bandFmts></c:surfaceChart>
            <c:surface3DChart><c:bandFmts/></c:surface3DChart>
        </c:plotArea></c:chart>
    </c:chartSpace>"#;
    let chart = read(xml.as_slice()).unwrap();
    let TypeGroup::Surface(surface) = &chart.plot_area.type_groups[0] else {
        panic!("expected surface chart");
    };
    assert!(surface.wireframe);
    assert_eq!(
        surface
            .band_formats
            .as_ref()
            .unwrap()
            .iter()
            .map(|format| format.index)
            .collect::<Vec<_>>(),
        [2, 7]
    );
    assert!(
        std::str::from_utf8(
            surface.band_formats.as_ref().unwrap()[0]
                .shape_properties
                .as_ref()
                .unwrap()
                .as_xml()
        )
        .unwrap()
        .contains("102030")
    );
    let expected_shape_properties = surface.band_formats.as_ref().unwrap()[0]
        .shape_properties
        .clone();
    let TypeGroup::Surface3D(surface) = &chart.plot_area.type_groups[1] else {
        panic!("expected 3D surface chart");
    };
    assert!(surface.band_formats.as_ref().unwrap().is_empty());

    let mut output = Vec::new();
    crate::chart::writer::write(&mut output, &chart).unwrap();
    let reparsed = read(output.as_slice()).unwrap();
    let TypeGroup::Surface(surface) = &reparsed.plot_area.type_groups[0] else {
        panic!("expected surface chart");
    };
    assert_eq!(surface.band_formats.as_ref().unwrap().len(), 2);
    assert_eq!(
        surface.band_formats.as_ref().unwrap()[0].shape_properties,
        expected_shape_properties
    );

    let duplicate =
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:plotArea><c:surfaceChart><c:bandFmts>
            <c:bandFmt><c:idx val="1"/></c:bandFmt>
            <c:bandFmt><c:idx val="1"/></c:bandFmt>
        </c:bandFmts></c:surfaceChart></c:plotArea></c:chart></c:chartSpace>"#;
    assert!(read(duplicate.as_slice()).is_err());

    let mut invalid = Chart::new();
    let mut surface = SurfaceTypeGroup::new();
    surface.band_formats = Some(vec![BandFormat::new(3), BandFormat::new(3)]);
    invalid
        .plot_area
        .type_groups
        .push(TypeGroup::Surface(surface));
    assert!(crate::chart::writer::write(&mut Vec::new(), &invalid).is_err());
}

#[test]
fn parses_of_pie_schema_defaults_and_empty_custom_split() {
    let xml = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:plotArea><c:ofPieChart><c:ofPieType/>
            <c:gapWidth/><c:splitType/><c:custSplit/><c:secondPieSize/>
        </c:ofPieChart></c:plotArea></c:chart>
    </c:chartSpace>"#;

    let chart = read(xml.as_slice()).unwrap();
    let [TypeGroup::OfPie(group)] = chart.plot_area.type_groups.as_slice() else {
        panic!("expected an of-pie chart");
    };
    assert_eq!(group.of_pie_type, OfPieType::Pie);
    assert_eq!(group.gap_width, Some(150));
    assert_eq!(group.split_type, Some(OfPieSplitType::Automatic));
    assert_eq!(group.custom_split_points.as_deref(), Some([].as_slice()));
    assert_eq!(group.second_pie_size, Some(75));

    let percent_xml =
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:plotArea><c:ofPieChart><c:ofPieType val="bar"/>
            <c:gapWidth val="225%"/><c:secondPieSize val="80%"/>
        </c:ofPieChart></c:plotArea></c:chart>
    </c:chartSpace>"#;
    let chart = read(percent_xml.as_slice()).unwrap();
    let [TypeGroup::OfPie(group)] = chart.plot_area.type_groups.as_slice() else {
        panic!("expected an of-pie chart");
    };
    assert_eq!(group.gap_width, Some(225));
    assert_eq!(group.second_pie_size, Some(80));
}

#[test]
fn rejects_invalid_of_pie_input_and_output() {
    for xml in [
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:ofPieChart/></c:plotArea></c:chart></c:chartSpace>"#.as_slice(),
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:ofPieChart><c:ofPieType val="ring"/></c:ofPieChart></c:plotArea></c:chart></c:chartSpace>"#.as_slice(),
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:ofPieChart><c:ofPieType/><c:gapWidth val="501"/></c:ofPieChart></c:plotArea></c:chart></c:chartSpace>"#.as_slice(),
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:ofPieChart><c:ofPieType/><c:secondPieSize val="4"/></c:ofPieChart></c:plotArea></c:chart></c:chartSpace>"#.as_slice(),
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:ofPieChart><c:ofPieType/><c:splitPos val="NaN"/></c:ofPieChart></c:plotArea></c:chart></c:chartSpace>"#.as_slice(),
    ] {
        assert!(read(xml).is_err());
    }

    for invalid_group in [
        OfPieTypeGroup {
            gap_width: Some(501),
            ..OfPieTypeGroup::new(OfPieType::Pie)
        },
        OfPieTypeGroup {
            split_position: Some(f64::INFINITY),
            ..OfPieTypeGroup::new(OfPieType::Pie)
        },
        OfPieTypeGroup {
            second_pie_size: Some(4),
            ..OfPieTypeGroup::new(OfPieType::Pie)
        },
    ] {
        let mut chart = Chart::new();
        chart
            .plot_area
            .type_groups
            .push(TypeGroup::OfPie(invalid_group));
        assert!(crate::chart::writer::write(&mut Vec::new(), &chart).is_err());
    }
}

#[test]
fn parses_chart_group_percentage_union_values() {
    let xml =
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:plotArea>
            <c:barChart><c:barDir val="col"/><c:gapWidth val="225%"/><c:overlap val="-25%"/></c:barChart>
            <c:bar3DChart><c:barDir val="bar"/><c:gapWidth/><c:gapDepth val="175%"/><c:shape/></c:bar3DChart>
            <c:bubbleChart><c:bubbleScale val="125%"/><c:sizeRepresents/></c:bubbleChart>
            <c:doughnutChart><c:firstSliceAng/><c:holeSize val="5%"/></c:doughnutChart>
            <c:area3DChart><c:gapDepth val="500%"/></c:area3DChart>
            <c:lineChart><c:smooth/></c:lineChart>
            <c:line3DChart><c:gapDepth/></c:line3DChart>
        </c:plotArea></c:chart>
    </c:chartSpace>"#;

    let chart = read(xml.as_slice()).unwrap();
    let TypeGroup::Bar(bar) = &chart.plot_area.type_groups[0] else {
        panic!("expected bar chart");
    };
    assert_eq!(bar.gap_width, Some(225));
    assert_eq!(bar.overlap, Some(-25));
    let TypeGroup::Bar3D(bar) = &chart.plot_area.type_groups[1] else {
        panic!("expected 3D bar chart");
    };
    assert_eq!(bar.gap_width, Some(150));
    assert_eq!(bar.gap_depth, Some(175));
    assert_eq!(bar.shape, Some(BarShape::Box));
    let TypeGroup::Bubble(bubble) = &chart.plot_area.type_groups[2] else {
        panic!("expected bubble chart");
    };
    assert_eq!(bubble.scale().get(), 125);
    assert_eq!(bubble.size(), BubbleSize::Area);
    let TypeGroup::Doughnut(doughnut) = &chart.plot_area.type_groups[3] else {
        panic!("expected doughnut chart");
    };
    assert_eq!(doughnut.first_slice_angle, 0);
    assert_eq!(doughnut.hole_size, 5);
    let TypeGroup::Area3D(area) = &chart.plot_area.type_groups[4] else {
        panic!("expected 3D area chart");
    };
    assert_eq!(area.gap_depth, Some(500));
    let TypeGroup::Line(line) = &chart.plot_area.type_groups[5] else {
        panic!("expected line chart");
    };
    assert!(line.smooth);
    let TypeGroup::Line3D(line) = &chart.plot_area.type_groups[6] else {
        panic!("expected 3D line chart");
    };
    assert_eq!(line.gap_depth, Some(150));
}

#[test]
fn writer_rejects_invalid_chart_group_ranges() {
    let mut bar = BarTypeGroup::new(BarDirection::Column, BarGrouping::Clustered);
    bar.gap_width = Some(501);
    let mut bar_3d = Bar3DTypeGroup::new(BarDirection::Column, BarGrouping::Clustered);
    bar_3d.gap_depth = Some(501);
    let mut doughnut = DoughnutTypeGroup::new();
    doughnut.hole_size = 0;
    let mut line_3d = Line3DTypeGroup::new(BarGrouping::Standard);
    line_3d.gap_depth = Some(501);
    let mut pie = PieTypeGroup::new();
    pie.first_slice_angle = 361;

    for group in [
        TypeGroup::Bar(bar),
        TypeGroup::Bar3D(bar_3d),
        TypeGroup::Doughnut(doughnut),
        TypeGroup::Line3D(line_3d),
        TypeGroup::Pie(pie),
    ] {
        let mut chart = Chart::new();
        chart.plot_area.type_groups.push(group);
        assert!(crate::chart::writer::write(&mut Vec::new(), &chart).is_err());
    }
}

#[test]
fn writer_rejects_invalid_axes_and_duplicate_legend_entries() {
    let mut chart = Chart::new();
    let mut axis = ValueAxis::new(1, AxisPosition::Left, 2);
    let mut units = DisplayUnits::custom(1_000.0);
    units.built_in_unit = Some(BuiltInUnit::Thousands);
    axis.display_units = Some(Box::new(units));
    chart.plot_area.axes.push(Axis::Value(axis));
    assert!(crate::chart::writer::write(&mut Vec::new(), &chart).is_err());

    let mut chart = Chart::new();
    let mut axis = CategoryAxis::new(1, AxisPosition::Bottom, 2);
    axis.min = Some(2.0);
    axis.max = Some(1.0);
    chart.plot_area.axes.push(Axis::Category(axis));
    assert!(crate::chart::writer::write(&mut Vec::new(), &chart).is_err());

    let mut chart = Chart::new();
    let mut axis = ValueAxis::new(1, AxisPosition::Left, 2);
    axis.display_units = Some(Box::new(DisplayUnits::custom(f64::NAN)));
    chart.plot_area.axes.push(Axis::Value(axis));
    assert!(crate::chart::writer::write(&mut Vec::new(), &chart).is_err());

    let mut chart = Chart::new();
    let legend = Legend {
        entries: vec![LegendEntry::new(4), LegendEntry::new(4)],
        ..Legend::default()
    };
    chart.legend = Some(legend);
    assert!(crate::chart::writer::write(&mut Vec::new(), &chart).is_err());
}

#[test]
fn round_trips_legend_and_entry_formatting_fragments() {
    let xml = br#"<c:chartSpace
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:x="urn:example:legend">
        <c:chart><c:plotArea/><c:legend><c:legendPos val="b"/>
            <c:legendEntry><c:idx val="2"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p/></c:txPr>
                <c:extLst><c:ext uri="entry"><x:entryPayload/></c:ext></c:extLst>
            </c:legendEntry>
            <c:legendEntry><c:idx val="3"/><c:delete val="1"/></c:legendEntry>
            <c:overlay val="1"/>
            <c:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill></c:spPr>
            <c:txPr><a:bodyPr rot="1200000"/><a:lstStyle/><a:p/></c:txPr>
            <c:extLst><c:ext uri="legend"><x:legendPayload/></c:ext></c:extLst>
        </c:legend></c:chart>
    </c:chartSpace>"#;
    let chart = read(xml.as_slice()).unwrap();
    let legend = chart.legend.as_ref().unwrap();
    assert_eq!(legend.position, LegendPosition::Bottom);
    assert!(legend.overlay);
    assert_eq!(legend.entries.len(), 2);
    assert!(legend.entries[0].text_properties.is_some());
    assert!(legend.entries[0].extension_list.is_some());
    assert!(legend.entries[1].deleted);
    assert!(
        std::str::from_utf8(legend.shape_properties.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("123456")
    );
    assert!(
        std::str::from_utf8(legend.text_properties.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("1200000")
    );
    assert!(
        std::str::from_utf8(legend.extension_list.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("legendPayload")
    );

    let mut output = Vec::new();
    crate::chart::writer::write(&mut output, &chart).unwrap();
    let reparsed = read(output.as_slice()).unwrap();
    let reparsed = reparsed.legend.as_ref().unwrap();
    assert_eq!(reparsed.shape_properties, legend.shape_properties);
    assert_eq!(reparsed.text_properties, legend.text_properties);
    assert_eq!(reparsed.extension_list, legend.extension_list);
    assert_eq!(
        reparsed.entries[0].text_properties,
        legend.entries[0].text_properties
    );
    assert_eq!(
        reparsed.entries[0].extension_list,
        legend.entries[0].extension_list
    );

    for invalid_entry in [
        br#"<c:legendEntry><c:idx val="1"/></c:legendEntry>"#.as_slice(),
        br#"<c:legendEntry><c:idx val="1"/><c:delete/><c:txPr/></c:legendEntry>"#.as_slice(),
    ] {
        let mut document = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea/><c:legend>"#.to_vec();
        document.extend_from_slice(invalid_entry);
        document.extend_from_slice(b"</c:legend></c:chart></c:chartSpace>");
        assert!(read(document.as_slice()).is_err());
    }

    let mut invalid = Chart::new();
    let mut legend = Legend::default();
    let mut entry = LegendEntry::new(1);
    entry.deleted = true;
    entry.text_properties = Some(
        TextProperties::from_xml(
            br#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#
                .to_vec(),
        )
        .unwrap(),
    );
    legend.entries.push(entry);
    invalid.legend = Some(legend);
    assert!(crate::chart::writer::write(&mut Vec::new(), &invalid).is_err());
}

#[test]
fn round_trips_chart_and_axis_title_formatting_fragments() {
    let xml = br#"<c:chartSpace
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:x="urn:example:title">
        <c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx>
            <c:spPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></c:spPr>
            <c:txPr><a:bodyPr rot="600000"/><a:lstStyle/><a:p/></c:txPr>
            <c:extLst><c:ext uri="chart-title"><x:chartPayload/></c:ext></c:extLst>
        </c:title><c:plotArea><c:catAx><c:axId val="1"/><c:scaling>
            <c:extLst><c:ext uri="scaling"><x:scalingPayload/></c:ext></c:extLst>
        </c:scaling>
            <c:axPos val="b"/><c:majorGridlines><c:spPr><a:solidFill><a:srgbClr val="445566"/></a:solidFill></c:spPr></c:majorGridlines>
            <c:minorGridlines/><c:title><c:tx><c:rich><a:p><a:r><a:t>Quarter</a:t></a:r></a:p></c:rich></c:tx>
                <c:spPr><a:noFill/></c:spPr><c:txPr><a:bodyPr vert="vert"/><a:lstStyle/><a:p/></c:txPr>
                <c:extLst><c:ext uri="axis-title"><x:axisPayload/></c:ext></c:extLst>
            </c:title><c:spPr><a:ln w="12700"/></c:spPr>
            <c:txPr><a:bodyPr rot="-600000"/><a:lstStyle/><a:p/></c:txPr>
            <c:crossAx val="2"/><c:extLst><c:ext uri="axis"><x:axisBodyPayload/></c:ext></c:extLst>
            </c:catAx></c:plotArea></c:chart>
    </c:chartSpace>"#;
    let chart = read(xml.as_slice()).unwrap();
    assert!(
        std::str::from_utf8(chart.title_shape_properties.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("112233")
    );
    assert!(
        std::str::from_utf8(chart.title_text_properties.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("600000")
    );
    assert!(
        std::str::from_utf8(chart.title_extension_list.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("chartPayload")
    );
    let common = chart.plot_area.axes[0].common();
    assert!(common.show_major_gridlines);
    assert!(common.show_minor_gridlines);
    assert!(
        std::str::from_utf8(
            common
                .major_gridlines
                .as_ref()
                .unwrap()
                .shape_properties
                .as_ref()
                .unwrap()
                .as_xml()
        )
        .unwrap()
        .contains("445566")
    );
    assert!(common.minor_gridlines.is_some());
    assert!(
        std::str::from_utf8(common.shape_properties.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("12700")
    );
    assert!(
        std::str::from_utf8(common.text_properties.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("-600000")
    );
    assert!(
        std::str::from_utf8(common.scaling_extension_list.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("scalingPayload")
    );
    assert!(
        std::str::from_utf8(common.extension_list.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("axisBodyPayload")
    );
    assert!(common.title_shape_properties.is_some());
    assert!(
        std::str::from_utf8(common.title_text_properties.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("vert")
    );
    assert!(
        std::str::from_utf8(common.title_extension_list.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("axisPayload")
    );

    let mut output = Vec::new();
    crate::chart::writer::write(&mut output, &chart).unwrap();
    let reparsed = read(output.as_slice()).unwrap();
    assert_eq!(
        reparsed.title_shape_properties,
        chart.title_shape_properties
    );
    assert_eq!(reparsed.title_text_properties, chart.title_text_properties);
    assert_eq!(reparsed.title_extension_list, chart.title_extension_list);
    let reparsed_common = reparsed.plot_area.axes[0].common();
    assert_eq!(
        reparsed_common.title_shape_properties,
        common.title_shape_properties
    );
    assert_eq!(
        reparsed_common.title_text_properties,
        common.title_text_properties
    );
    assert_eq!(
        reparsed_common.title_extension_list,
        common.title_extension_list
    );
    assert_eq!(reparsed_common.major_gridlines, common.major_gridlines);
    assert_eq!(reparsed_common.minor_gridlines, common.minor_gridlines);
    assert_eq!(reparsed_common.shape_properties, common.shape_properties);
    assert_eq!(reparsed_common.text_properties, common.text_properties);
    assert_eq!(
        reparsed_common.scaling_extension_list,
        common.scaling_extension_list
    );
    assert_eq!(reparsed_common.extension_list, common.extension_list);

    let duplicate = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:title><c:spPr/><c:spPr/></c:title><c:plotArea/></c:chart></c:chartSpace>"#;
    assert!(read(duplicate.as_slice()).is_err());

    let duplicate_gridlines = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:catAx><c:axId val="1"/><c:scaling/><c:axPos val="b"/><c:majorGridlines/><c:majorGridlines/><c:crossAx val="2"/></c:catAx></c:plotArea></c:chart></c:chartSpace>"#;
    assert!(read(duplicate_gridlines.as_slice()).is_err());
}

#[test]
fn rejects_truncated_and_invalid_chart_values() {
    for xml in [
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea>"#.as_slice(),
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotVisOnly val="yes"/></c:chart></c:chartSpace>"#.as_slice(),
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:dispBlanksAs val="empty"/></c:chart></c:chartSpace>"#.as_slice(),
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:style val="0"/><c:chart></c:chart></c:chartSpace>"#.as_slice(),
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:view3D><c:perspective val="241"/></c:view3D></c:chart></c:chartSpace>"#.as_slice(),
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:legend><c:legendPos val="center"/></c:legend></c:chart></c:chartSpace>"#.as_slice(),
    ] {
        assert!(read(xml).is_err());
    }
}

#[test]
fn rejects_empty_plot_containers_without_consuming_following_content() {
    for xml in [
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea><c:lineChart/><c:dTable/></c:plotArea></c:chart>
        </c:chartSpace>"#
            .as_slice(),
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea><c:catAx/><c:dTable/></c:plotArea></c:chart>
        </c:chartSpace>"#
            .as_slice(),
    ] {
        assert!(read(xml).is_err());
    }

    let empty_layout =
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea><c:layout/><c:lineChart></c:lineChart>
            </c:plotArea></c:chart>
        </c:chartSpace>"#;
    let chart = read(empty_layout.as_slice()).unwrap();
    assert!(chart.plot_area.layout.is_some());
}

// This intentionally exercises every chart-group variant in one shared
// round trip so additions cannot silently escape the exhaustive matrix.
#[allow(clippy::cognitive_complexity)]
#[test]
fn writer_round_trips_every_modeled_chart_group() {
    let mut area_3d = Area3DTypeGroup::new(BarGrouping::Stacked);
    area_3d.gap_depth = Some(175);
    area_3d.common.axis_ids = vec![10, 20, 30];
    let mut doughnut = DoughnutTypeGroup::new();
    doughnut.first_slice_angle = 45;
    doughnut.hole_size = 60;
    let mut surface = SurfaceTypeGroup::new();
    surface.wireframe = true;
    let mut surface_3d = Surface3DTypeGroup::new();
    surface_3d.wireframe = true;
    let mut bubble = BubbleTypeGroup::new()
        .with_scale(BubbleScale::new(125).unwrap())
        .with_size(BubbleSize::Width);
    bubble.show_negative_bubbles = false;
    let mut bubble_series = Series::new(0);
    bubble_series.x_values = Some(NumericData::from_values(vec![1.0]));
    bubble_series.y_values = Some(NumericData::from_values(vec![2.0]));
    bubble_series.bubble_sizes = Some(NumericData::from_values(vec![3.0]));
    bubble_series.bubble_3d = true;
    bubble.common.series.push(bubble_series);
    let mut scatter = ScatterTypeGroup::new(ScatterStyle::SmoothMarker);
    let mut scatter_series = Series::new(4);
    scatter_series.marker_symbol = Some(MarkerStyle::Star);
    scatter_series.marker_size = Some(9);
    scatter_series.marker_shape_properties = Some(
        ShapeProperties::from_xml(
            br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:solidFill><a:srgbClr val="AABBCC"/></a:solidFill></c:spPr>"#.to_vec(),
        )
        .unwrap(),
    );
    scatter_series.marker_extension_list = Some(
        ExtensionList::from_xml(
            br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:x="urn:example:marker"><c:ext uri="series-marker"><x:payload/></c:ext></c:extLst>"#.to_vec(),
        )
        .unwrap(),
    );
    scatter_series.smooth = true;
    let mut point = DataPoint::new(2).with_marker(7, MarkerStyle::Diamond);
    point.marker_shape_properties = Some(
        ShapeProperties::from_xml(
            br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:ln w="12700"/></c:spPr>"#.to_vec(),
        )
        .unwrap(),
    );
    point.marker_extension_list = Some(
        ExtensionList::from_xml(
            br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:x="urn:example:marker"><c:ext uri="point-marker"><x:payload/></c:ext></c:extLst>"#.to_vec(),
        )
        .unwrap(),
    );
    point.invert_if_negative = true;
    point.bubble_3d = Some(false);
    point.explosion = Some(15);
    scatter_series.data_points.push(point);
    let mut labels = DataLabels::new()
        .with_position(DataLabelPosition::Top)
        .with_show_value(true);
    labels.number_format = Some(NumberFormat::new("0.0%").with_source_linked(false));
    labels.show_series_name = true;
    labels.show_leader_lines = true;
    labels.shape_properties = Some(
        ShapeProperties::from_xml(
            br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:solidFill><a:srgbClr val="DDEEFF"/></a:solidFill></c:spPr>"#.to_vec(),
        )
        .unwrap(),
    );
    labels.text_properties = Some(
        TextProperties::from_xml(
            br#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr rot="1200000"/><a:lstStyle/><a:p/></c:txPr>"#.to_vec(),
        )
        .unwrap(),
    );
    labels.leader_lines = Some(Lines {
        shape_properties: Some(
            ShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:ln w="38100"/></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        ),
    });
    labels.extension_list = Some(
        ExtensionList::from_xml(
            br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:x="urn:example:labels"><c:ext uri="labels"><x:payload/></c:ext></c:extLst>"#.to_vec(),
        )
        .unwrap(),
    );
    labels.separator = Some(" & ".to_string());
    let mut point_label = DataLabel::new(2);
    point_label.layout = Some(Layout::new().with_position(0.6, 0.7));
    point_label.text = Some(TitleText::from_ref("Sheet1!$E$2"));
    point_label.number_format = Some(NumberFormat::new("$0.00"));
    point_label.position = Some(DataLabelPosition::Left);
    point_label.show_category_name = true;
    point_label.separator = Some(" / ".to_string());
    point_label.shape_properties = Some(
        ShapeProperties::from_xml(
            br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:noFill/></c:spPr>"#.to_vec(),
        )
        .unwrap(),
    );
    point_label.text_properties = Some(
        TextProperties::from_xml(
            br#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr vert="vert"/><a:lstStyle/><a:p/></c:txPr>"#.to_vec(),
        )
        .unwrap(),
    );
    point_label.extension_list = Some(
        ExtensionList::from_xml(
            br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:x="urn:example:labels"><c:ext uri="point-label"><x:payload/></c:ext></c:extLst>"#.to_vec(),
        )
        .unwrap(),
    );
    labels.labels.push(point_label);
    scatter_series.data_labels = Some(labels);
    let mut trendline = Trendline::linear();
    trendline.name = Some("Forecast & fit".to_string());
    trendline.forward = Some(2.5);
    trendline.intercept = Some(-1.0);
    trendline.display_equation = true;
    trendline.display_r_squared = true;
    trendline.show_label = true;
    trendline.label = Some(TitleText::from_ref("Sheet1!$F$2"));
    trendline.label_layout = Some(Layout::new().with_size(0.3, 0.2));
    trendline.label_number_format = Some(NumberFormat::new("0.000").with_source_linked(false));
    trendline.shape_properties = Some(
        ShapeProperties::from_xml(
            br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:ln w="63500"/></c:spPr>"#.to_vec(),
        )
        .unwrap(),
    );
    trendline.label_shape_properties = Some(
        ShapeProperties::from_xml(
            br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:solidFill><a:srgbClr val="FEDCBA"/></a:solidFill></c:spPr>"#.to_vec(),
        )
        .unwrap(),
    );
    trendline.label_text_properties = Some(
        TextProperties::from_xml(
            br#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr rot="1800000"/><a:lstStyle/><a:p/></c:txPr>"#.to_vec(),
        )
        .unwrap(),
    );
    trendline.label_extension_list = Some(
        ExtensionList::from_xml(
            br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:x="urn:example:trendline"><c:ext uri="label"><x:payload/></c:ext></c:extLst>"#.to_vec(),
        )
        .unwrap(),
    );
    trendline.extension_list = Some(
        ExtensionList::from_xml(
            br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:x="urn:example:trendline"><c:ext uri="trendline"><x:payload/></c:ext></c:extLst>"#.to_vec(),
        )
        .unwrap(),
    );
    scatter_series.trendlines.push(trendline);
    scatter_series.error_bars.push(ErrorBar {
        direction: ErrorBarDirection::Y,
        error_type: ErrorBarType::Both,
        value_type: ErrorBarValueType::Fixed,
        value: Some(1.5),
        plus_values: None,
        minus_values: None,
        no_end_cap: true,
        shape_properties: Some(
            ShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:ln w="50800"/></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        ),
        extension_list: Some(
            ExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:x="urn:example:error-bars"><c:ext uri="error-bars"><x:payload/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        ),
    });
    scatter_series.error_bars.push(ErrorBar {
        direction: ErrorBarDirection::X,
        error_type: ErrorBarType::Plus,
        value_type: ErrorBarValueType::Custom,
        value: None,
        plus_values: Some(NumericData::from_ref("Sheet1!$D$2:$D$4")),
        minus_values: None,
        no_end_cap: false,
        shape_properties: None,
        extension_list: None,
    });
    scatter.common.series.push(scatter_series);
    let mut of_pie = OfPieTypeGroup::new(OfPieType::Bar);
    of_pie.common.vary_colors = true;
    of_pie.gap_width = Some(225);
    of_pie.split_type = Some(OfPieSplitType::Custom);
    of_pie.split_position = Some(3.5);
    of_pie.custom_split_points = Some(vec![1, 4]);
    of_pie.second_pie_size = Some(80);
    let mut line = LineTypeGroup::new(BarGrouping::Standard);
    line.smooth = true;
    line.common.axis_ids = vec![41, 42];
    let mut group_labels = DataLabels::new()
        .with_position(DataLabelPosition::Right)
        .with_show_value(true);
    group_labels.show_category_name = true;
    group_labels.separator = Some(" | ".to_string());
    line.common.data_labels = Some(group_labels);
    let mut line_3d = Line3DTypeGroup::new(BarGrouping::PercentStacked);
    line_3d.gap_depth = Some(210);
    line_3d.common.axis_ids = vec![50, 51, 52];

    let mut chart = Chart::new();
    chart.plot_area.data_table = Some(DataTable {
        show_horizontal_border: true,
        show_vertical_border: false,
        show_outline: true,
        show_legend_keys: true,
        ..DataTable::default()
    });
    chart.plot_area.type_groups = vec![
        TypeGroup::Area(AreaTypeGroup::new(BarGrouping::Standard)),
        TypeGroup::Area3D(area_3d),
        TypeGroup::Bar(BarTypeGroup::new(
            BarDirection::Column,
            BarGrouping::Clustered,
        )),
        TypeGroup::Bar3D(Bar3DTypeGroup::new(BarDirection::Bar, BarGrouping::Stacked)),
        TypeGroup::Bubble(bubble),
        TypeGroup::Doughnut(doughnut),
        TypeGroup::Line(line),
        TypeGroup::Line3D(line_3d),
        TypeGroup::OfPie(of_pie),
        TypeGroup::Pie(PieTypeGroup::new()),
        TypeGroup::Pie3D(Pie3DTypeGroup::new()),
        TypeGroup::Radar(RadarTypeGroup::new(RadarStyle::Filled)),
        TypeGroup::Scatter(scatter),
        TypeGroup::Stock(StockTypeGroup::new()),
        TypeGroup::Surface(surface),
        TypeGroup::Surface3D(surface_3d),
    ];
    for (index, group) in chart.plot_area.type_groups.iter_mut().enumerate() {
        group.common_mut().extension_list = Some(
            ExtensionList::from_xml(
                format!(
                    r#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:x="urn:example:group"><c:ext uri="group-{index}"><x:payload/></c:ext></c:extLst>"#
                )
                .into_bytes(),
            )
            .unwrap(),
        );
    }

    let mut xml = Vec::new();
    crate::chart::writer::write(&mut xml, &chart).unwrap();
    assert!(
        std::str::from_utf8(&xml)
            .unwrap()
            .contains("<c:name>Forecast &amp; fit</c:name>")
    );
    let parsed = read(xml.as_slice()).unwrap();

    assert_eq!(parsed.plot_area.type_groups.len(), 16);
    for (index, group) in parsed.plot_area.type_groups.iter().enumerate() {
        assert!(
            std::str::from_utf8(group.common().extension_list.as_ref().unwrap().as_xml())
                .unwrap()
                .contains(&format!(r#"uri="group-{index}""#))
        );
    }
    let data_table = parsed.plot_area.data_table.as_ref().unwrap();
    assert!(data_table.show_horizontal_border);
    assert!(!data_table.show_vertical_border);
    assert!(data_table.show_outline);
    assert!(data_table.show_legend_keys);
    assert!(matches!(
        parsed.plot_area.type_groups[0],
        TypeGroup::Area(_)
    ));
    assert!(matches!(
        parsed.plot_area.type_groups[1],
        TypeGroup::Area3D(_)
    ));
    let TypeGroup::Area3D(group) = &parsed.plot_area.type_groups[1] else {
        unreachable!();
    };
    assert_eq!(group.gap_depth, Some(175));
    assert_eq!(group.common.axis_ids, [10, 20, 30]);
    assert!(matches!(
        parsed.plot_area.type_groups[4],
        TypeGroup::Bubble(_)
    ));
    assert!(matches!(
        parsed.plot_area.type_groups[5],
        TypeGroup::Doughnut(_)
    ));
    assert!(matches!(
        parsed.plot_area.type_groups[7],
        TypeGroup::Line3D(_)
    ));
    let TypeGroup::Line(group) = &parsed.plot_area.type_groups[6] else {
        unreachable!();
    };
    assert!(group.smooth);
    assert_eq!(group.common.axis_ids, [41, 42]);
    let labels = group.common.data_labels.as_ref().unwrap();
    assert_eq!(labels.position, Some(DataLabelPosition::Right));
    assert!(labels.show_value);
    assert!(labels.show_category_name);
    assert_eq!(labels.separator.as_deref(), Some(" | "));
    let TypeGroup::Line3D(group) = &parsed.plot_area.type_groups[7] else {
        unreachable!();
    };
    assert_eq!(group.gap_depth, Some(210));
    assert_eq!(group.common.axis_ids, [50, 51, 52]);
    assert!(matches!(parsed.plot_area.type_groups[9], TypeGroup::Pie(_)));
    assert!(matches!(
        parsed.plot_area.type_groups[10],
        TypeGroup::Pie3D(_)
    ));
    assert!(matches!(
        parsed.plot_area.type_groups[11],
        TypeGroup::Radar(_)
    ));
    assert!(matches!(
        parsed.plot_area.type_groups[13],
        TypeGroup::Stock(_)
    ));
    assert!(matches!(
        parsed.plot_area.type_groups[14],
        TypeGroup::Surface(_)
    ));
    assert!(matches!(
        parsed.plot_area.type_groups[15],
        TypeGroup::Surface3D(_)
    ));
    let TypeGroup::Doughnut(group) = &parsed.plot_area.type_groups[5] else {
        unreachable!();
    };
    assert_eq!(group.first_slice_angle, 45);
    assert_eq!(group.hole_size, 60);
    let TypeGroup::Bubble(group) = &parsed.plot_area.type_groups[4] else {
        unreachable!();
    };
    assert_eq!(group.scale().get(), 125);
    assert!(!group.show_negative_bubbles);
    assert_eq!(group.size(), BubbleSize::Width);
    assert_eq!(
        group.common.series[0].x_values.as_ref().unwrap().values,
        [1.0]
    );
    assert_eq!(
        group.common.series[0].y_values.as_ref().unwrap().values,
        [2.0]
    );
    assert_eq!(
        group.common.series[0].bubble_sizes.as_ref().unwrap().values,
        [3.0]
    );
    assert!(group.common.series[0].bubble_3d);
    let TypeGroup::OfPie(group) = &parsed.plot_area.type_groups[8] else {
        unreachable!();
    };
    assert_eq!(group.of_pie_type, OfPieType::Bar);
    assert!(group.common.vary_colors);
    assert_eq!(group.gap_width, Some(225));
    assert_eq!(group.split_type, Some(OfPieSplitType::Custom));
    assert_eq!(group.split_position, Some(3.5));
    assert_eq!(group.custom_split_points.as_deref(), Some(&[1, 4][..]));
    assert_eq!(group.second_pie_size, Some(80));
    let TypeGroup::Scatter(group) = &parsed.plot_area.type_groups[12] else {
        unreachable!();
    };
    let series = &group.common.series[0];
    assert_eq!(series.marker_symbol, Some(MarkerStyle::Star));
    assert_eq!(series.marker_size, Some(9));
    assert!(
        std::str::from_utf8(series.marker_shape_properties.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("AABBCC")
    );
    assert!(series.marker_extension_list.is_some());
    assert!(series.smooth);
    assert_eq!(series.data_points.len(), 1);
    assert_eq!(series.data_points[0].index, 2);
    assert_eq!(series.data_points[0].marker_size, Some(7));
    assert_eq!(
        series.data_points[0].marker_symbol,
        Some(MarkerStyle::Diamond)
    );
    assert!(
        std::str::from_utf8(
            series.data_points[0]
                .marker_shape_properties
                .as_ref()
                .unwrap()
                .as_xml()
        )
        .unwrap()
        .contains("12700")
    );
    assert!(series.data_points[0].marker_extension_list.is_some());
    assert!(series.data_points[0].invert_if_negative);
    assert_eq!(series.data_points[0].bubble_3d, Some(false));
    assert_eq!(series.data_points[0].explosion, Some(15));
    let labels = series.data_labels.as_ref().unwrap();
    assert_eq!(labels.position, Some(DataLabelPosition::Top));
    assert!(labels.show_value);
    assert!(labels.show_series_name);
    assert!(labels.show_leader_lines);
    assert!(
        std::str::from_utf8(labels.shape_properties.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("DDEEFF")
    );
    assert!(
        std::str::from_utf8(labels.text_properties.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("1200000")
    );
    assert!(
        std::str::from_utf8(
            labels
                .leader_lines
                .as_ref()
                .unwrap()
                .shape_properties
                .as_ref()
                .unwrap()
                .as_xml()
        )
        .unwrap()
        .contains("38100")
    );
    assert!(labels.extension_list.is_some());
    let number_format = labels.number_format.as_ref().unwrap();
    assert_eq!(number_format.format_code, "0.0%");
    assert!(!number_format.source_linked);
    assert!(
        std::str::from_utf8(
            series.trendlines[0]
                .shape_properties
                .as_ref()
                .unwrap()
                .as_xml()
        )
        .unwrap()
        .contains("63500")
    );
    assert!(
        std::str::from_utf8(
            series.trendlines[0]
                .label_shape_properties
                .as_ref()
                .unwrap()
                .as_xml()
        )
        .unwrap()
        .contains("FEDCBA")
    );
    assert!(
        std::str::from_utf8(
            series.trendlines[0]
                .label_text_properties
                .as_ref()
                .unwrap()
                .as_xml()
        )
        .unwrap()
        .contains("1800000")
    );
    assert!(series.trendlines[0].label_extension_list.is_some());
    assert!(series.trendlines[0].extension_list.is_some());
    assert_eq!(labels.separator.as_deref(), Some(" & "));
    assert_eq!(labels.labels.len(), 1);
    let point_label = &labels.labels[0];
    assert_eq!(point_label.index, 2);
    assert_eq!(point_label.layout.as_ref().unwrap().x, Some(0.6));
    assert_eq!(point_label.layout.as_ref().unwrap().y, Some(0.7));
    let Some(TitleText::Reference(text)) = point_label.text.as_ref() else {
        panic!("expected point data-label formula");
    };
    assert_eq!(text.formula, "Sheet1!$E$2");
    assert_eq!(
        point_label.number_format.as_ref().unwrap().format_code,
        "$0.00"
    );
    assert_eq!(point_label.position, Some(DataLabelPosition::Left));
    assert!(point_label.show_category_name);
    assert_eq!(point_label.separator.as_deref(), Some(" / "));
    assert!(point_label.shape_properties.is_some());
    assert!(
        std::str::from_utf8(point_label.text_properties.as_ref().unwrap().as_xml())
            .unwrap()
            .contains("vert")
    );
    assert!(point_label.extension_list.is_some());
    assert_eq!(series.trendlines.len(), 1);
    assert_eq!(series.trendlines[0].name.as_deref(), Some("Forecast & fit"));
    assert_eq!(series.trendlines[0].forward, Some(2.5));
    assert_eq!(series.trendlines[0].intercept, Some(-1.0));
    assert!(series.trendlines[0].display_equation);
    assert!(series.trendlines[0].display_r_squared);
    assert!(series.trendlines[0].show_label);
    let Some(TitleText::Reference(label)) = series.trendlines[0].label.as_ref() else {
        panic!("expected trendline-label formula");
    };
    assert_eq!(label.formula, "Sheet1!$F$2");
    assert_eq!(
        series.trendlines[0].label_layout.as_ref().unwrap().width,
        Some(0.3)
    );
    assert_eq!(
        series.trendlines[0].label_layout.as_ref().unwrap().height,
        Some(0.2)
    );
    let number_format = series.trendlines[0].label_number_format.as_ref().unwrap();
    assert_eq!(number_format.format_code, "0.000");
    assert!(!number_format.source_linked);
    assert_eq!(series.error_bars.len(), 2);
    assert_eq!(series.error_bars[0].direction, ErrorBarDirection::Y);
    assert_eq!(series.error_bars[0].value, Some(1.5));
    assert!(series.error_bars[0].no_end_cap);
    assert!(
        std::str::from_utf8(
            series.error_bars[0]
                .shape_properties
                .as_ref()
                .unwrap()
                .as_xml()
        )
        .unwrap()
        .contains("50800")
    );
    assert!(series.error_bars[0].extension_list.is_some());
    assert_eq!(series.error_bars[1].value_type, ErrorBarValueType::Custom);
    assert_eq!(
        series.error_bars[1]
            .plus_values
            .as_ref()
            .unwrap()
            .source_ref
            .as_ref()
            .unwrap()
            .formula,
        "Sheet1!$D$2:$D$4"
    );
}

#[test]
fn writer_round_trips_modeled_axis_properties_in_one_scaling_block() {
    let mut category = CategoryAxis::new(10, AxisPosition::Bottom, 20);
    category.common.orientation = AxisOrientation::MaxMin;
    category.common.title = Some(TitleText::from_string("Quarter"));
    category.common.title_overlay = true;
    category.common.layout = Some(Layout::new().with_position(0.3, 0.4));
    category.common.number_format =
        Some(NumberFormat::new("mmm-yy \"fiscal\"").with_source_linked(false));
    category.common.major_tick_mark = TickMark::Cross;
    category.common.minor_tick_mark = TickMark::In;
    category.common.tick_label_position = TickLabelPosition::Low;
    category.common.deleted = true;
    category.common.cross_mode = AxisCrossMode::Max;
    category.common.show_major_gridlines = true;
    category.min = Some(1.0);
    category.max = Some(12.0);
    category.log_base = Some(2.0);
    category.auto = false;
    category.label_align = Some(AxisLabelAlign::Right);
    category.label_offset = Some(250);
    category.tick_label_skip = Some(2);
    category.tick_mark_skip = Some(3);
    category.no_multi_level = true;

    let mut value = ValueAxis::new(20, AxisPosition::Left, 10);
    value.common.crosses_at = Some(0.5);
    value.common.show_minor_gridlines = true;
    value.min = Some(-5.0);
    value.max = Some(100.0);
    value.major_unit = Some(10.0);
    value.minor_unit = Some(2.0);
    value.log_base = Some(10.0);
    value.cross_between = AxisCrossBetween::MidCategory;
    let mut display_units = DisplayUnits::built_in(BuiltInUnit::Millions);
    display_units.show_label = true;
    display_units.label = Some(TitleText::from_string("Millions sold"));
    display_units.layout = Some(Layout::new().with_position(0.15, 0.25));
    value.display_units = Some(Box::new(display_units));

    let mut date = DateAxis::new(30, AxisPosition::Top, 40);
    date.min = Some(45_000.0);
    date.max = Some(46_000.0);
    date.log_base = Some(10.0);
    date.major_unit = Some(2.0);
    date.minor_unit = Some(1.0);
    date.major_time_unit = Some(TimeUnit::Months);
    date.minor_time_unit = Some(TimeUnit::Days);
    date.base_time_unit = Some(TimeUnit::Years);
    date.auto = false;
    date.label_offset = Some(175);

    let mut series = SeriesAxis::new(40, AxisPosition::Right, 30);
    series.min = Some(1.0);
    series.max = Some(8.0);
    series.log_base = Some(2.0);
    series.tick_label_skip = Some(4);
    series.tick_mark_skip = Some(5);

    let mut chart = Chart::new();
    chart.title = Some(TitleText::from_ref("Sheet1!$C$1"));
    chart.title_layout = Some(Layout::new().with_size(0.5, 0.1));
    chart.title_overlay = true;
    let mut layout = Layout::new().with_position(0.1, 0.2).with_size(0.7, 0.6);
    layout.target = Some(LayoutTarget::Inner);
    layout.x_mode = Some(LayoutMode::Factor);
    layout.y_mode = Some(LayoutMode::Edge);
    chart.plot_area.layout = Some(layout);
    let mut legend = Legend::new(LegendPosition::Top).with_overlay(true);
    legend.layout = Some(Layout::new().with_size(0.4, 0.2));
    let mut deleted_entry = LegendEntry::new(2);
    deleted_entry.deleted = true;
    legend.entries = vec![deleted_entry, LegendEntry::new(3)];
    chart.legend = Some(legend);
    chart.plot_area.axes = vec![
        Axis::Category(category),
        Axis::Value(value),
        Axis::Date(date),
        Axis::Series(series),
    ];

    let mut xml = Vec::new();
    crate::chart::writer::write(&mut xml, &chart).unwrap();
    let xml_text = std::str::from_utf8(&xml).unwrap();
    assert_eq!(xml_text.matches("<c:scaling>").count(), 4);
    let parsed = read(xml.as_slice()).unwrap();
    let Some(TitleText::Reference(title)) = parsed.title.as_ref() else {
        panic!("expected chart title reference");
    };
    assert_eq!(title.formula, "Sheet1!$C$1");
    assert!(parsed.title_overlay);
    assert_eq!(parsed.title_layout.as_ref().unwrap().width, Some(0.5));
    assert_eq!(parsed.title_layout.as_ref().unwrap().height, Some(0.1));
    let layout = parsed.plot_area.layout.as_ref().unwrap();
    assert_eq!(layout.x, Some(0.1));
    assert_eq!(layout.y, Some(0.2));
    assert_eq!(layout.width, Some(0.7));
    assert_eq!(layout.height, Some(0.6));
    assert_eq!(layout.target, Some(LayoutTarget::Inner));
    assert_eq!(layout.x_mode, Some(LayoutMode::Factor));
    assert_eq!(layout.y_mode, Some(LayoutMode::Edge));
    assert_eq!(parsed.plot_area.axes.len(), 4);

    let Axis::Category(category) = &parsed.plot_area.axes[0] else {
        unreachable!();
    };
    assert_eq!(category.common.orientation, AxisOrientation::MaxMin);
    assert_eq!(category.min, Some(1.0));
    assert_eq!(category.max, Some(12.0));
    assert_eq!(category.log_base, Some(2.0));
    assert!(category.common.deleted);
    assert_eq!(category.common.major_tick_mark, TickMark::Cross);
    assert_eq!(category.common.minor_tick_mark, TickMark::In);
    assert_eq!(category.common.tick_label_position, TickLabelPosition::Low);
    assert_eq!(category.common.cross_mode, AxisCrossMode::Max);
    assert!(category.common.show_major_gridlines);
    let Some(TitleText::Literal(title)) = category.common.title.as_ref() else {
        panic!("expected literal category-axis title");
    };
    assert_eq!(title.text, "Quarter");
    assert!(category.common.title_overlay);
    assert_eq!(category.common.layout.as_ref().unwrap().x, Some(0.3));
    assert_eq!(category.common.layout.as_ref().unwrap().y, Some(0.4));
    assert_eq!(category.label_align, Some(AxisLabelAlign::Right));
    assert_eq!(category.label_offset, Some(250));
    assert_eq!(category.tick_label_skip, Some(2));
    assert_eq!(category.tick_mark_skip, Some(3));
    assert!(category.no_multi_level);
    let number_format = category.common.number_format.as_ref().unwrap();
    assert_eq!(number_format.format_code, "mmm-yy \"fiscal\"");
    assert!(!number_format.source_linked);

    let Axis::Value(value) = &parsed.plot_area.axes[1] else {
        unreachable!();
    };
    assert_eq!(value.min, Some(-5.0));
    assert_eq!(value.max, Some(100.0));
    assert_eq!(value.log_base, Some(10.0));
    assert_eq!(value.major_unit, Some(10.0));
    assert_eq!(value.minor_unit, Some(2.0));
    assert_eq!(value.common.crosses_at, Some(0.5));
    assert!(value.common.show_minor_gridlines);
    assert_eq!(value.cross_between, AxisCrossBetween::MidCategory);
    assert_eq!(
        value.display_units.as_ref().unwrap().built_in_unit,
        Some(BuiltInUnit::Millions)
    );
    let display_units = value.display_units.as_ref().unwrap();
    assert!(display_units.show_label);
    let Some(TitleText::Literal(label)) = display_units.label.as_ref() else {
        panic!("expected literal display-units label");
    };
    assert_eq!(label.text, "Millions sold");
    assert_eq!(display_units.layout.as_ref().unwrap().x, Some(0.15));
    assert_eq!(display_units.layout.as_ref().unwrap().y, Some(0.25));

    let legend = parsed.legend.as_ref().unwrap();
    assert_eq!(legend.position, LegendPosition::Top);
    assert!(legend.overlay);
    assert_eq!(legend.layout.as_ref().unwrap().width, Some(0.4));
    assert_eq!(legend.layout.as_ref().unwrap().height, Some(0.2));
    assert_eq!(legend.entries.len(), 2);
    assert_eq!(legend.entries[0].index, 2);
    assert!(legend.entries[0].deleted);
    assert_eq!(legend.entries[1].index, 3);
    assert!(!legend.entries[1].deleted);

    let Axis::Date(date) = &parsed.plot_area.axes[2] else {
        unreachable!();
    };
    assert_eq!(date.min, Some(45_000.0));
    assert_eq!(date.max, Some(46_000.0));
    assert_eq!(date.log_base, Some(10.0));
    assert_eq!(date.major_time_unit, Some(TimeUnit::Months));
    assert_eq!(date.minor_time_unit, Some(TimeUnit::Days));
    assert_eq!(date.base_time_unit, Some(TimeUnit::Years));
    assert!(!date.auto);
    assert_eq!(date.label_offset, Some(175));

    let Axis::Series(series) = &parsed.plot_area.axes[3] else {
        unreachable!();
    };
    assert_eq!(series.tick_label_skip, Some(4));
    assert_eq!(series.tick_mark_skip, Some(5));
    assert_eq!(series.min, Some(1.0));
    assert_eq!(series.max, Some(8.0));
    assert_eq!(series.log_base, Some(2.0));
}

#[test]
fn rejects_invalid_scaling_on_every_axis_kind() {
    for axis in [
        r#"<c:catAx><c:axId val="1"/><c:scaling><c:min val="1"/><c:min val="2"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>"#,
        r#"<c:valAx><c:axId val="1"/><c:scaling><c:max val="1"/><c:min val="2"/></c:scaling><c:axPos val="l"/><c:crossAx val="2"/></c:valAx>"#,
        r#"<c:dateAx><c:axId val="1"/><c:scaling><c:logBase val="1"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:dateAx>"#,
        r#"<c:serAx><c:axId val="1"/><c:scaling><c:logBase val="1001"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:serAx>"#,
    ] {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea>{axis}</c:plotArea></c:chart></c:chartSpace>"#
        );
        assert!(read(xml.as_bytes()).is_err());
    }
}
