//! Regression tests for the ODP XML parser owner.

use super::{Parser, model::ShapeBuilder};
use crate::model::legacy_animation::Kind as AnimationKind;
use crate::model::{
    Action, Actuate, Direction, DrawingAttributeNamespace, DrawingHyperlink, DrawingShapeKind,
    EnhancedGeometryChildKind, HyperlinkShow, Kind, Namespace, Shape, ShapeEventListener, Show,
    Slide, SoundShow, Speed, Type,
};
use litchi_core::ShapeType;

const TEST_PRESENTATION_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0">
<office:body>
    <office:presentation>
        <draw:page draw:name="Slide1">
            <draw:frame draw:name="Title" presentation:class="title" svg:x="1cm" svg:y="1cm" svg:width="18cm" svg:height="3cm">
                <draw:text-box>
                    <text:p>Welcome</text:p>
                </draw:text-box>
            </draw:frame>
            <draw:rect draw:name="Box1" svg:x="2cm" svg:y="5cm" svg:width="5cm" svg:height="3cm">
                <draw:text-box>
                    <text:p>Rectangle content</text:p>
                </draw:text-box>
            </draw:rect>
        </draw:page>
        <draw:page draw:name="Slide2">
            <draw:frame draw:name="Content" presentation:class="object" svg:x="1cm" svg:y="4cm">
                <draw:text-box>
                    <text:p>Bullet 1</text:p>
                    <text:p>Bullet 2</text:p>
                </draw:text-box>
            </draw:frame>
        </draw:page>
    </office:presentation>
</office:body>
</office:document-content>"#;

#[test]
fn preserves_drawing_element_kinds_and_unmodeled_geometry_attributes() {
    let xml = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
        xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
        xmlns:r="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0">
        <o:body><o:presentation><d:page>
          <d:rect d:name="rect" d:corner-radius="2mm"/>
          <d:ellipse d:name="ellipse" d:kind="section" d:start-angle="30" s:rx="2cm" s:ry="1cm"/>
          <d:circle d:name="circle" s:cx="3cm" s:cy="4cm" s:r="2cm"/>
          <d:path d:name="path" s:viewBox="0 0 100 100" s:d="M 0 0 L 100 100"/>
          <d:polygon d:name="polygon" s:viewBox="0 0 10 10" d:points="0,0 10,0 5,10"/>
          <d:polyline d:name="polyline" d:points="0,0 5,5 10,0"/>
          <d:regular-polygon d:name="regular" d:corners="7" d:concave="true" d:sharpness="25%"/>
          <d:page-thumbnail d:name="thumb" d:page-number="2"/>
          <d:measure d:name="measure" s:x1="1cm" s:y1="2cm" s:x2="3cm" s:y2="4cm"/>
          <d:caption d:name="caption" d:caption-point-x="1cm" d:caption-point-y="2cm"/>
          <d:connector d:name="connector" d:type="curve" d:start-shape="a" d:end-shape="b" s:x1="0cm" s:y1="0cm" s:x2="1cm" s:y2="1cm"/>
          <d:control d:name="control" d:control="control1"/>
          <d:custom-shape d:name="custom" d:engine="vendor" d:data="opaque" r:transform="rotatex(0.5)">
            <d:enhanced-geometry d:type="non-primitive" s:viewBox="0 0 21600 21600"
              d:modifiers="10800" d:enhanced-path="M 0 0 L ?f0 21600 Z" r:projection="perspective">
              <d:equation d:name="f0" d:formula="$0 * 2 &amp; 21600"/>
              <d:handle d:handle-position="$0 10800" d:handle-range-x-minimum="0" d:handle-range-x-maximum="21600"/>
            </d:enhanced-geometry>
          </d:custom-shape>
        </d:page></o:presentation></o:body>
    </o:document-content>"#;
    let slides = Parser::parse_slides(xml).unwrap();
    let shapes = &slides[0].shapes;
    let expected = [
        DrawingShapeKind::Rectangle,
        DrawingShapeKind::Ellipse,
        DrawingShapeKind::Circle,
        DrawingShapeKind::Path,
        DrawingShapeKind::Polygon,
        DrawingShapeKind::Polyline,
        DrawingShapeKind::RegularPolygon,
        DrawingShapeKind::PageThumbnail,
        DrawingShapeKind::Measure,
        DrawingShapeKind::Caption,
        DrawingShapeKind::Connector,
        DrawingShapeKind::Control,
        DrawingShapeKind::CustomShape,
    ];
    assert_eq!(shapes.len(), expected.len());
    for (index, (shape, expected_kind)) in shapes.iter().zip(expected).enumerate() {
        assert_eq!(shape.drawing_kind(), Some(expected_kind));
        let regenerated = crate::Builder::generate_shape_xml(shape, index).unwrap();
        assert!(regenerated.starts_with(&format!("<{}", expected_kind.element_name())));
        assert!(!regenerated.contains("draw:layer="));
    }
    let ellipse = crate::Builder::generate_shape_xml(&shapes[1], 1).unwrap();
    assert!(ellipse.contains(r#"draw:kind="section""#));
    assert!(ellipse.contains(r#"draw:start-angle="30""#));
    assert!(ellipse.contains(r#"svg:rx="2cm""#));
    let path = crate::Builder::generate_shape_xml(&shapes[3], 3).unwrap();
    assert!(path.contains(r#"svg:viewBox="0 0 100 100""#));
    assert!(path.contains(r#"svg:d="M 0 0 L 100 100""#));
    let connector = crate::Builder::generate_shape_xml(&shapes[10], 10).unwrap();
    assert!(connector.contains(r#"draw:type="curve""#));
    assert!(connector.contains(r#"draw:start-shape="a""#));
    let custom = &shapes[12];
    assert!(custom.drawing_attributes().iter().any(|attribute| {
        attribute.namespace() == DrawingAttributeNamespace::Dr3d
            && attribute.local_name() == "transform"
            && attribute.value() == "rotatex(0.5)"
    }));
    let geometry = custom.enhanced_geometry().unwrap();
    assert_eq!(geometry.children().len(), 2);
    assert_eq!(
        geometry.children()[0].kind(),
        EnhancedGeometryChildKind::Equation
    );
    assert_eq!(
        geometry.children()[1].kind(),
        EnhancedGeometryChildKind::Handle
    );
    let regenerated = crate::Builder::generate_shape_xml(custom, 12).unwrap();
    assert!(regenerated.contains("<draw:enhanced-geometry"));
    assert!(regenerated.contains(r#"dr3d:projection="perspective""#));
    assert!(regenerated.contains(r#"draw:formula="$0 * 2 &amp; 21600""#));
    assert!(regenerated.contains(r#"draw:handle-position="$0 10800""#));
}

#[test]
fn preserves_recursive_inert_three_dimensional_scenes() {
    let xml = r##"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
        xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
        xmlns:r="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0">
        <o:body><o:presentation><d:page>
          <r:scene s:x="1cm" s:y="2cm" s:width="8cm" s:height="6cm"
            r:vrp="(0 0 10)" r:projection="perspective" r:ambient-color="#112233">
            <r:light r:direction="(0 0 -1)" r:diffuse-color="#ffffff" r:enabled="true"/>
            <r:cube r:min-edge="(-1 -1 -1)" r:max-edge="(1 1 1)" r:transform="rotatex(0.5)"/>
            <r:scene r:shade-mode="phong">
              <r:sphere r:center="(0 0 0)" r:size="(2 2 2)"/>
              <r:extrude s:viewBox="0 0 10 10" s:d="M 0 0 L 10 10"/>
              <r:rotate s:viewBox="0 0 20 20" s:d="M 1 1 L 19 19"/>
            </r:scene>
          </r:scene>
        </d:page></o:presentation></o:body>
    </o:document-content>"##;
    let slides = Parser::parse_slides(xml).unwrap();
    let scene = &slides[0].shapes[0];
    assert_eq!(
        scene.drawing_kind(),
        Some(DrawingShapeKind::ThreeDimensionalScene)
    );
    assert_eq!(scene.x.as_deref(), Some("1cm"));
    assert_eq!(scene.children.len(), 3);
    assert_eq!(
        scene.children[0].drawing_kind(),
        Some(DrawingShapeKind::ThreeDimensionalLight)
    );
    assert_eq!(
        scene.children[1].drawing_kind(),
        Some(DrawingShapeKind::ThreeDimensionalCube)
    );
    let nested = &scene.children[2];
    assert_eq!(
        nested.drawing_kind(),
        Some(DrawingShapeKind::ThreeDimensionalScene)
    );
    assert_eq!(nested.children.len(), 3);
    assert_eq!(
        nested.children[0].drawing_kind(),
        Some(DrawingShapeKind::ThreeDimensionalSphere)
    );
    assert_eq!(
        nested.children[1].drawing_kind(),
        Some(DrawingShapeKind::ThreeDimensionalExtrude)
    );
    assert_eq!(
        nested.children[2].drawing_kind(),
        Some(DrawingShapeKind::ThreeDimensionalRotate)
    );
    let regenerated = crate::Builder::generate_shape_xml(scene, 0).unwrap();
    assert!(regenerated.starts_with("<dr3d:scene"));
    assert!(regenerated.contains(r#"dr3d:projection="perspective""#));
    assert!(regenerated.contains(r#"dr3d:direction="(0 0 -1)""#));
    assert!(regenerated.contains(r#"<dr3d:cube"#));
    assert!(regenerated.contains(r#"<dr3d:sphere"#));
    assert!(regenerated.contains(r#"svg:d="M 0 0 L 10 10""#));
}

#[test]
fn rejects_invalid_three_dimensional_shape_hierarchies() {
    for body in [
        r#"<r:cube/>"#,
        r#"<d:g><r:sphere/></d:g>"#,
        r#"<r:scene><d:rect/></r:scene>"#,
        r#"<r:scene><r:cube/><r:light r:direction="(0 0 -1)"/></r:scene>"#,
        r#"<r:scene><r:cube>not empty</r:cube></r:scene>"#,
        r#"<r:scene><r:cube><d:glue-point/></r:cube></r:scene>"#,
        r#"<r:scene><r:light/></r:scene>"#,
        r#"<r:scene><r:extrude s:d="M 0 0"/></r:scene>"#,
        r#"<r:scene><r:rotate s:viewBox="0 0 10 10"/></r:scene>"#,
    ] {
        let xml = format!(
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:r="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"><o:body><o:presentation><d:page>{body}</d:page></o:presentation></o:body></o:document-content>"#
        );
        assert!(Parser::parse_slides(&xml).is_err(), "accepted {body}");
    }
}

#[test]
fn rejects_misplaced_or_invalid_enhanced_geometry() {
    for shape in [
        "<d:rect><d:enhanced-geometry/></d:rect>",
        "<d:custom-shape><d:enhanced-geometry/><d:enhanced-geometry/></d:custom-shape>",
        "<d:custom-shape><d:enhanced-geometry><d:handle/><d:equation/></d:enhanced-geometry></d:custom-shape>",
        "<d:custom-shape><d:equation/></d:custom-shape>",
    ] {
        let xml = format!(
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><o:body><o:presentation><d:page>{shape}</d:page></o:presentation></o:body></o:document-content>"#
        );
        assert!(Parser::parse_slides(&xml).is_err(), "accepted {shape}");
    }
}

const TEST_EMPTY_PRESENTATION: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0">
<office:body>
    <office:presentation>
    </office:presentation>
</office:body>
</office:document-content>"#;

const TEST_SHAPES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0">
<office:body>
    <office:presentation>
        <draw:page draw:name="Shapes">
            <draw:ellipse draw:name="Circle1" svg:x="1cm" svg:y="1cm" svg:width="3cm" svg:height="3cm">
                <draw:text-box>
                    <text:p>Circle</text:p>
                </draw:text-box>
            </draw:ellipse>
            <draw:line draw:name="Line1" svg:x1="0cm" svg:y1="0cm" svg:x2="10cm" svg:y2="10cm"/>
            <draw:connector draw:name="Connector1" svg:x1="1cm" svg:y1="2cm" svg:x2="3cm" svg:y2="4cm"/>
            <draw:custom-shape draw:name="Custom1" svg:x="5cm" svg:y="5cm"/>
            <presentation:notes><draw:frame><draw:text-box><text:p>Speaker note</text:p></draw:text-box></draw:frame></presentation:notes>
        </draw:page>
    </office:presentation>
</office:body>
</office:document-content>"#;

#[test]
fn test_parse_slides() {
    let slides = Parser::parse_slides(TEST_PRESENTATION_XML).unwrap();
    assert_eq!(slides.len(), 2);

    // First slide
    assert_eq!(slides[0].title, Some("Welcome".to_string()));
    assert_eq!(slides[0].index, 0);
    assert!(slides[0].text.is_empty());
    assert_eq!(slides[0].shapes.len(), 1);
    assert_eq!(slides[0].shapes[0].text, "Rectangle content");
    assert_eq!(slides[0].all_text(), "Welcome\nRectangle content");

    // Second slide
    assert_eq!(slides[1].title, None);
    assert_eq!(slides[1].index, 1);
    assert_eq!(slides[1].text, "Bullet 1\nBullet 2");
    assert!(slides[1].shapes.is_empty());
}

#[test]
fn test_parse_empty_presentation() {
    let slides = Parser::parse_slides(TEST_EMPTY_PRESENTATION).unwrap();
    assert!(slides.is_empty());
}

#[test]
fn test_parse_shapes() {
    let slides = Parser::parse_slides(TEST_SHAPES_XML).unwrap();
    assert_eq!(slides.len(), 1);

    let slide = &slides[0];
    assert_eq!(slide.shapes.len(), 4);
    assert!(
        slide
            .shapes
            .iter()
            .any(|shape| shape.shape_type == ShapeType::Connector)
    );
    assert_eq!(slide.notes.as_deref(), Some("Speaker note"));
    assert!(!slide.all_text().contains("Speaker note"));
}

#[test]
fn test_slide_debug() {
    let slide = Slide {
        title: Some("Test".to_string()),
        text: "Content".to_string(),
        index: 0,
        notes: None,
        transition: None,
        animations: vec![],
        legacy_animation: None,
        shapes: vec![],
    };
    let debug_str = format!("{:?}", slide);
    assert!(debug_str.contains("Slide"));
    assert!(debug_str.contains("Test"));
}

#[test]
fn test_slide_clone() {
    let slide = Slide {
        title: Some("Test".to_string()),
        text: "Content".to_string(),
        index: 0,
        notes: None,
        transition: None,
        animations: vec![],
        legacy_animation: None,
        shapes: vec![],
    };
    let cloned = slide.clone();
    assert_eq!(slide.title, cloned.title);
    assert_eq!(slide.text, cloned.text);
}

#[test]
fn test_shape_debug() {
    let shape = Shape {
        shape_type: ShapeType::TextBox,
        text: "Shape text".to_string(),
        name: Some("Shape1".to_string()),
        x: Some("1cm".to_string()),
        y: Some("2cm".to_string()),
        width: Some("10cm".to_string()),
        height: Some("5cm".to_string()),
        style_name: Some("Style1".to_string()),
        ..Shape::new()
    };
    let debug_str = format!("{:?}", shape);
    assert!(debug_str.contains("Shape"));
    assert!(debug_str.contains("TextBox"));
}

#[test]
fn test_shape_clone() {
    let shape = Shape {
        shape_type: ShapeType::AutoShape,
        text: "Text".to_string(),
        name: Some("Name".to_string()),
        x: Some("0cm".to_string()),
        y: Some("0cm".to_string()),
        width: Some("5cm".to_string()),
        height: Some("3cm".to_string()),
        style_name: None,
        ..Shape::new()
    };
    let cloned = shape.clone();
    assert_eq!(shape.shape_type, cloned.shape_type);
    assert_eq!(shape.name, cloned.name);
}

#[test]
fn test_shape_type_variants() {
    // Test all shape type variants
    let types = vec![
        ShapeType::TextBox,
        ShapeType::AutoShape,
        ShapeType::Line,
        ShapeType::Placeholder,
        ShapeType::Picture,
        ShapeType::Group,
        ShapeType::Connector,
        ShapeType::Table,
        ShapeType::GraphicFrame,
        ShapeType::Unknown,
    ];

    for shape_type in types {
        let shape = Shape {
            shape_type,
            text: String::new(),
            name: None,
            x: None,
            y: None,
            width: None,
            height: None,
            style_name: None,
            ..Shape::new()
        };
        let _ = format!("{:?}", shape);
    }
}

#[test]
fn test_shape_type_equality() {
    assert_eq!(ShapeType::TextBox, ShapeType::TextBox);
    assert_ne!(ShapeType::TextBox, ShapeType::Line);
    assert_ne!(ShapeType::AutoShape, ShapeType::Picture);
}

#[test]
fn test_shape_type_clone() {
    let t1 = ShapeType::Placeholder;
    let t2 = t1;
    assert_eq!(t1, t2);
}

#[test]
fn test_shape_type_copy() {
    let t1 = ShapeType::Line;
    let t2 = t1;
    assert_eq!(t1, t2); // Copy trait allows this
}

#[test]
fn test_shape_builder() {
    let builder = ShapeBuilder::new();
    let shape = builder.build();
    assert_eq!(shape.shape_type, ShapeType::AutoShape);
    assert!(shape.text.is_empty());
}

#[test]
fn test_shape_builder_with_data() {
    let mut builder = ShapeBuilder::new();
    builder.name = Some("TestShape".to_string());
    builder.x = Some("1cm".to_string());
    builder.y = Some("2cm".to_string());
    builder.width = Some("10cm".to_string());
    builder.height = Some("5cm".to_string());
    builder.text = "Hello".to_string();
    builder.shape_type = ShapeType::TextBox;

    let shape = builder.build();
    assert_eq!(shape.name, Some("TestShape".to_string()));
    assert_eq!(shape.x, Some("1cm".to_string()));
    assert_eq!(shape.text, "Hello");
    assert_eq!(shape.shape_type, ShapeType::TextBox);
}

#[test]
fn test_shape_builder_clone() {
    let builder = ShapeBuilder::new();
    let cloned = builder.build().clone();
    assert_eq!(cloned.shape_type, ShapeType::AutoShape);
}

#[test]
fn parses_picture_shape_and_unescapes_href() {
    let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:presentation><draw:page draw:name="Images"><draw:frame draw:name="Picture"><draw:image xlink:href="Pictures/a&amp;b.png"/></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#;

    let slides = Parser::parse_slides(xml).unwrap();
    let shape = &slides[0].shapes[0];
    assert_eq!(shape.shape_type, ShapeType::Picture);
    assert_eq!(shape.image_href(), Some("Pictures/a&b.png"));
}

#[test]
fn preserves_shape_stacking_transform_and_presentation_role() {
    let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"><office:body><office:presentation><draw:page><draw:frame draw:name="Chart" draw:layer="controls" draw:z-index="184467440737095516160" draw:transform="rotate (0.5) translate (1cm 2cm)" presentation:class="chart" presentation:placeholder="true" presentation:user-transformed="false"/></draw:page></office:presentation></office:body></office:document-content>"#;

    let slides = Parser::parse_slides(xml).unwrap();
    let shape = &slides[0].shapes[0];
    assert_eq!(shape.shape_type, ShapeType::Placeholder);
    assert_eq!(shape.layer(), Some("controls"));
    assert_eq!(shape.z_index(), Some("184467440737095516160"));
    assert_eq!(shape.transform(), Some("rotate (0.5) translate (1cm 2cm)"));
    assert_eq!(shape.presentation_class(), Some("chart"));
    assert_eq!(shape.presentation_placeholder, Some(true));
    assert_eq!(shape.presentation_user_transformed, Some(false));
}

#[test]
fn rejects_invalid_shape_stacking_and_boolean_values() {
    for attribute in [
        r#"draw:z-index="-1""#,
        r#"draw:z-index="1.5""#,
        r#"presentation:placeholder="yes""#,
        r#"presentation:user-transformed="no""#,
    ] {
        let xml = format!(
            r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"><office:body><office:presentation><draw:page><draw:frame {attribute}/></draw:page></office:presentation></office:body></office:document-content>"#
        );
        assert!(Parser::parse_slides(&xml).is_err(), "accepted {attribute}");
    }
}

#[test]
fn preserves_shape_groups_and_identifies_opaque_frames() {
    let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:presentation><draw:page><draw:g draw:name="Group"><draw:a xlink:href="https://example.test/group-child" xlink:type="simple"><draw:rect/></draw:a></draw:g><draw:frame draw:name="Table"><table:table/></draw:frame><draw:frame draw:name="Object"><draw:object/></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#;

    let slides = Parser::parse_slides(xml).unwrap();
    let types: Vec<_> = slides[0]
        .shapes
        .iter()
        .map(|shape| shape.shape_type)
        .collect();
    assert_eq!(
        types,
        [ShapeType::Group, ShapeType::Table, ShapeType::GraphicFrame]
    );
    let group = &slides[0].shapes[0];
    assert_eq!(group.children().len(), 1);
    assert_eq!(
        group.children()[0].drawing_kind(),
        Some(DrawingShapeKind::Rectangle)
    );
    assert_eq!(
        group.children()[0].hyperlink().map(DrawingHyperlink::href),
        Some("https://example.test/group-child")
    );
    let regenerated = crate::Builder::generate_shape_xml(group, 0).unwrap();
    assert!(regenerated.starts_with(r#"<draw:g draw:name="Group">"#));
    assert!(regenerated.contains("<draw:rect"));
    assert!(regenerated.contains("<draw:a"));
}

#[test]
fn bounds_nested_shape_group_depth() {
    let mut xml = String::from(
        r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><o:body><o:presentation><d:page>"#,
    );
    for _ in 0..65 {
        xml.push_str("<d:g>");
    }
    for _ in 0..65 {
        xml.push_str("</d:g>");
    }
    xml.push_str("</d:page></o:presentation></o:body></o:document-content>");
    let error = Parser::parse_slides(&xml).unwrap_err();
    assert!(error.to_string().contains("64 levels"));
}

#[test]
fn preserves_text_across_spans_and_odf_whitespace_elements() {
    let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:presentation><draw:page><draw:frame presentation:class="object"><draw:text-box><text:p><text:s/>Hel<text:span>lo</text:span> <text:span>world</text:span><text:s text:c="2"/>again<text:tab/>tab<text:line-break/>line &amp; more</text:p><text:p/><text:p>second paragraph<text:s/></text:p></draw:text-box></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#;

    let slides = Parser::parse_slides(xml).unwrap();
    assert_eq!(
        slides[0].text,
        " Hello world  again\ttab\nline & more\n\nsecond paragraph "
    );
}

#[test]
fn rejects_excessive_explicit_space_expansion() {
    let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:presentation><draw:page><draw:frame><draw:text-box><text:p>x<text:s text:c="1000001"/></text:p></draw:text-box></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#;

    let error = Parser::parse_slides(xml).unwrap_err();
    assert!(error.to_string().contains("safety limit"));
}

#[test]
fn parses_arbitrary_odf_namespace_prefixes() {
    let xml = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:l="http://www.w3.org/1999/xlink" xmlns:tb="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:f="urn:example:wrong"><o:body><o:presentation><f:page><t:p>ignored</t:p></f:page><d:page><d:frame d:name="Aliased Title" p:class="title" s:x="1cm"><d:text-box><t:p>Aliased<t:s/>title</t:p></d:text-box></d:frame><d:frame d:name="Picture"><d:image l:href="Pictures/a&amp;b.png"/></d:frame><d:connector d:name="Link" s:x1="1cm" s:y1="2cm" s:x2="3cm" s:y2="4cm"/><d:frame d:name="Table"><tb:table/></d:frame><p:notes><d:frame><d:text-box><t:p>Aliased note</t:p></d:text-box></d:frame></p:notes></d:page></o:presentation></o:body></o:document-content>"#;

    let slides = Parser::parse_slides(xml).unwrap();
    assert_eq!(slides.len(), 1);
    assert_eq!(slides[0].title.as_deref(), Some("Aliased title"));
    assert_eq!(slides[0].notes.as_deref(), Some("Aliased note"));
    let picture = &slides[0].shapes[0];
    assert_eq!(picture.name(), Some("Picture"));
    assert_eq!(picture.image_href(), Some("Pictures/a&b.png"));
    let connector = &slides[0].shapes[1];
    assert_eq!(connector.shape_type, ShapeType::Connector);
    assert_eq!(connector.position(), (Some("1cm"), Some("2cm")));
    assert_eq!(connector.dimensions(), (Some("3cm"), Some("4cm")));
    assert_eq!(slides[0].shapes[2].shape_type, ShapeType::Table);
}

#[test]
fn resolves_transition_styles_across_package_parts_and_inheritance() {
    let styles = r##"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:m="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" xmlns:l="http://www.w3.org/1999/xlink"><o:styles><s:default-style s:family="drawing-page"><s:drawing-page-properties p:transition-speed="slow"/></s:default-style><s:style s:name="Base" s:family="drawing-page"><s:drawing-page-properties p:transition-type="automatic" p:duration="PT8S"><p:sound l:type="simple" l:href="Sounds/a&amp;b.ogg" l:actuate="onRequest" l:show="replace" p:play-full="true"/></s:drawing-page-properties></s:style></o:styles></o:document-styles>"##;
    let content = r##"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:m="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0"><o:automatic-styles><s:style s:name="Child" s:family="drawing-page" s:parent-style-name="Base"><s:drawing-page-properties p:transition-style="fade-from-left" p:transition-speed="fast" m:type="fade" m:subtype="crossfade" m:direction="reverse" m:fadeColor="#aB09fF"/></s:style></o:automatic-styles><o:body><o:presentation><d:page d:style-name="Child"/></o:presentation></o:body></o:document-content>"##;

    let slides = Parser::parse_slides_with_styles(content, Some(styles)).unwrap();
    let transition = slides[0].transition().unwrap();
    assert_eq!(transition.transition_type(), Some(Type::Automatic));
    assert_eq!(transition.style().unwrap().as_str(), "fade-from-left");
    assert_eq!(transition.speed(), Some(Speed::Fast));
    assert_eq!(transition.smil_type(), Some("fade"));
    assert_eq!(transition.smil_subtype(), Some("crossfade"));
    assert_eq!(transition.direction(), Some(Direction::Reverse));
    assert_eq!(transition.fade_color(), Some("#aB09fF"));
    assert_eq!(transition.duration(), Some("PT8S"));
    let sound = transition.sound().unwrap();
    assert_eq!(sound.href, "Sounds/a&b.ogg");
    assert_eq!(sound.play_full, Some(true));
    assert!(sound.actuate_on_request);
    assert_eq!(sound.show, Some(SoundShow::Replace));
}

#[test]
fn rejects_cyclic_transition_style_inheritance() {
    let content = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><o:automatic-styles><s:style s:name="A" s:family="drawing-page" s:parent-style-name="B"/><s:style s:name="B" s:family="drawing-page" s:parent-style-name="A"/></o:automatic-styles><o:body><o:presentation><d:page d:style-name="A"/></o:presentation></o:body></o:document-content>"#;
    let error = Parser::parse_slides_with_styles(content, None).unwrap_err();
    assert!(error.to_string().contains("cyclic"));
}

#[test]
fn parses_complete_namespace_aware_animation_trees() {
    let xml = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:a="urn:oasis:names:tc:opendocument:xmlns:animation:1.0" xmlns:m="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:l="http://www.w3.org/1999/xlink" xmlns:e="urn:example:animation-extension" xmlns:f="urn:example:not-animation"><o:body><o:presentation><d:page><f:par/><a:par m:begin="slide.begin+1s" p:node-type="timing-root" e:flag="keep &amp; roundtrip"><a:animate a:formula="x+1" m:targetElement="shape1"/><a:animateColor a:color-interpolation="rgb"/><a:animateMotion s:path="M 0 0 L 1 1"/><a:animateTransform s:type="rotate"/><a:audio l:href="Sounds/chime.ogg" xml:id="audio1"/><a:command a:command="show"><a:param a:name="page" a:value="2"/></a:command><a:iterate a:iterate-type="by-paragraph"><a:set m:to="visible"/></a:iterate><a:par/><a:seq><a:transitionFilter m:type="fade"/></a:seq><a:set m:attributeName="visibility"/><a:transitionFilter m:subtype="crossfade"/></a:par></d:page></o:presentation></o:body></o:document-content>"#;

    let slides = Parser::parse_slides(xml).unwrap();
    assert_eq!(slides.len(), 1);
    assert_eq!(slides[0].animations.len(), 1);
    let root = &slides[0].animations[0];
    assert_eq!(root.kind(), Kind::Parallel);
    assert_eq!(root.children().len(), 11);
    assert_eq!(
        root.attribute(&Namespace::Smil, "begin"),
        Some("slide.begin+1s")
    );
    assert_eq!(
        root.attribute(
            &Namespace::Other("urn:example:animation-extension".to_string()),
            "flag"
        ),
        Some("keep & roundtrip")
    );
    assert_eq!(root.children()[0].kind(), Kind::Animate);
    assert_eq!(root.children()[4].kind(), Kind::Audio);
    let command = &root.children()[5];
    assert_eq!(command.kind(), Kind::Command);
    assert_eq!(command.children()[0].kind(), Kind::Parameter);
    assert_eq!(
        command.children()[0].attribute(&Namespace::Animation, "value"),
        Some("2")
    );
    assert_eq!(root.children()[6].children()[0].kind(), Kind::Set);
    assert_eq!(
        root.children()[8].children()[0].kind(),
        Kind::TransitionFilter
    );
}

#[test]
fn rejects_invalid_animation_structure() {
    let invalid_trees = [
        "<a:animate><a:set/></a:animate>",
        "<a:command><a:animate/></a:command>",
        "<a:param a:name=\"orphan\"/>",
        "<a:notInOdf/>",
        "<a:par>not whitespace</a:par>",
    ];
    for tree in invalid_trees {
        let xml = format!(
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:a="urn:oasis:names:tc:opendocument:xmlns:animation:1.0"><o:body><o:presentation><d:page>{tree}</d:page></o:presentation></o:body></o:document-content>"#
        );
        assert!(Parser::parse_slides(&xml).is_err(), "accepted {tree}");
    }
}

#[test]
fn bounds_animation_nesting() {
    let mut tree = "<a:par>".repeat(129);
    tree.push_str("<a:set/>");
    tree.push_str(&"</a:par>".repeat(129));
    let xml = format!(
        r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:a="urn:oasis:names:tc:opendocument:xmlns:animation:1.0"><o:body><o:presentation><d:page>{tree}</d:page></o:presentation></o:body></o:document-content>"#
    );

    let error = Parser::parse_slides(&xml).unwrap_err();
    assert!(error.to_string().contains("128 levels"));
}

#[test]
fn parses_inert_media_plugins_and_parameters() {
    let xml = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:x="http://www.w3.org/1999/xlink" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"><o:body><o:presentation><d:page><d:frame d:name="Movie" s:x="1cm" s:y="2cm"><d:plugin x:href="Media/a&amp;b.mp4" x:type="simple" d:mime-type="video/mp4" x:show="embed" x:actuate="onRequest" xml:id="movie1"><d:param d:name="autoplay" d:value="false"/><d:param d:name="caption" d:value="A &amp; B"> </d:param></d:plugin></d:frame></d:page></o:presentation></o:body></o:document-content>"#;

    let slides = Parser::parse_slides(xml).unwrap();
    let shape = &slides[0].shapes[0];
    assert_eq!(shape.shape_type, ShapeType::GraphicFrame);
    let media = shape.media().unwrap();
    assert_eq!(media.href(), "Media/a&b.mp4");
    assert_eq!(media.mime_type(), Some("video/mp4"));
    assert_eq!(media.show(), Some(Show::Embed));
    assert_eq!(media.actuate(), Some(Actuate::OnRequest));
    assert_eq!(media.xml_id(), Some("movie1"));
    assert_eq!(media.parameters().len(), 2);
    assert_eq!(media.parameters()[0].name(), "autoplay");
    assert_eq!(media.parameters()[1].value(), "A & B");
}

#[test]
fn rejects_schema_invalid_media_plugins() {
    let invalid_plugins = [
        r#"<d:frame><d:plugin x:type="simple"/></d:frame>"#,
        r#"<d:frame><d:plugin x:href="a.mp4"/></d:frame>"#,
        r#"<d:frame><d:plugin x:href="a.mp4" x:type="extended"/></d:frame>"#,
        r#"<d:frame><d:plugin x:href="a.mp4" x:type="simple" x:show="invalid"/></d:frame>"#,
        r#"<d:plugin x:href="a.mp4" x:type="simple"/>"#,
        r#"<d:rect><d:plugin x:href="a.mp4" x:type="simple"/></d:rect>"#,
        r#"<d:frame><d:plugin x:href="a.mp4" x:type="simple">text</d:plugin></d:frame>"#,
        r#"<d:frame><d:plugin x:href="a.mp4" x:type="simple"><d:param d:name="x"/></d:plugin></d:frame>"#,
        r#"<d:frame><d:plugin x:href="a.mp4" x:type="simple"><d:param d:name="x" d:value="y"><d:param d:name="nested" d:value="z"/></d:param></d:plugin></d:frame>"#,
    ];
    for plugin in invalid_plugins {
        let xml = format!(
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:x="http://www.w3.org/1999/xlink"><o:body><o:presentation><d:page>{plugin}</d:page></o:presentation></o:body></o:document-content>"#
        );
        assert!(Parser::parse_slides(&xml).is_err(), "accepted {plugin}");
    }
}

#[test]
fn parses_shape_hyperlinks_and_inert_event_bindings() {
    let xml = r##"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:sc="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:x="http://www.w3.org/1999/xlink" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"><o:body><o:presentation><d:page><d:a x:type="simple" x:href="#page2" x:actuate="onRequest" x:show="replace" o:target-frame-name="_self" o:name="jump" o:title="Jump &amp; return" o:server-map="true" xml:id="link1"><d:rect d:name="Action" s:x="1cm"><o:event-listeners><sc:event-listener sc:event-name="dom:click" sc:language="ooo:script" sc:macro-name="Standard.Module1.Main"/><sc:event-listener sc:event-name="dom:mouseover" sc:language="javascript" x:type="simple" x:href="Scripts/hover.js" x:actuate="onRequest"/><p:event-listener sc:event-name="dom:click" p:action="show" p:effect="fade" p:direction="from-left" p:speed="fast" p:start-scale="50%" x:type="simple" x:href="#page3" x:show="embed" x:actuate="onRequest" p:verb="2"><p:sound x:type="simple" x:href="Sounds/click.ogg" x:actuate="onRequest" x:show="replace" p:play-full="true" xml:id="sound1"/></p:event-listener></o:event-listeners></d:rect></d:a></d:page></o:presentation></o:body></o:document-content>"##;

    let slides = Parser::parse_slides(xml).unwrap();
    let shape = &slides[0].shapes[0];
    let hyperlink = shape.hyperlink().unwrap();
    assert_eq!(hyperlink.href(), "#page2");
    assert!(hyperlink.actuate_on_request());
    assert_eq!(hyperlink.show(), Some(HyperlinkShow::Replace));
    assert_eq!(hyperlink.target_frame_name(), Some("_self"));
    assert_eq!(hyperlink.title(), Some("Jump & return"));
    assert_eq!(hyperlink.server_map(), Some(true));
    assert_eq!(hyperlink.xml_id(), Some("link1"));

    assert_eq!(shape.event_listeners().len(), 3);
    let ShapeEventListener::Script(macro_listener) = &shape.event_listeners()[0] else {
        panic!("expected script listener");
    };
    assert_eq!(
        macro_listener.macro_name.as_deref(),
        Some("Standard.Module1.Main")
    );
    let ShapeEventListener::Action(action) = &shape.event_listeners()[2] else {
        panic!("expected presentation listener");
    };
    assert_eq!(action.action, Action::Show);
    assert_eq!(action.effect.as_ref().unwrap().as_str(), "fade");
    assert_eq!(action.direction.as_ref().unwrap().as_str(), "from-left");
    assert_eq!(action.speed, Some(Speed::Fast));
    assert_eq!(action.start_scale.as_deref(), Some("50%"));
    assert_eq!(action.verb, Some(2));
    assert_eq!(action.sound.as_ref().unwrap().href, "Sounds/click.ogg");
}

#[test]
fn rejects_invalid_shape_hyperlinks_and_event_bindings() {
    let invalid = [
        r##"<d:a x:href="#p"><d:rect/></d:a>"##,
        r##"<d:a x:type="simple" x:href="#p"/>"##,
        r##"<d:a x:type="simple" x:href="#p"><d:rect/><d:rect/></d:a>"##,
        r#"<p:event-listener sc:event-name="dom:click" p:action="next-page"/>"#,
        r#"<d:rect><o:event-listeners><sc:event-listener sc:event-name="dom:click" sc:language="ooo:script" sc:macro-name="M" x:type="simple" x:href="S"/></o:event-listeners></d:rect>"#,
        r#"<d:rect><o:event-listeners><p:event-listener sc:event-name="dom:click" p:action="invalid"/></o:event-listeners></d:rect>"#,
        r#"<d:rect><o:event-listeners><p:event-listener sc:event-name="dom:click" p:action="sound"><p:sound x:href="a" x:type="extended"/></p:event-listener></o:event-listeners></d:rect>"#,
        r#"<d:rect><o:event-listeners/><o:event-listeners/></d:rect>"#,
    ];
    for fragment in invalid {
        let xml = format!(
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:sc="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:x="http://www.w3.org/1999/xlink"><o:body><o:presentation><d:page>{fragment}</d:page></o:presentation></o:body></o:document-content>"#
        );
        assert!(Parser::parse_slides(&xml).is_err(), "accepted {fragment}");
    }
}

#[test]
fn parses_legacy_presentation_effect_trees() {
    let xml = r##"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:x="http://www.w3.org/1999/xlink" xmlns:e="urn:example:effects"><o:body><o:presentation><d:page><p:animations e:mode="legacy"><p:animation-group><p:show-shape d:shape-id="shape1" p:effect="fade" p:speed="fast"><p:sound x:href="Sounds/a&amp;b.ogg" x:type="simple" p:play-full="true"/></p:show-shape><p:dim d:shape-id="shape1" d:color="#808080"/><p:hide-text d:shape-id="shape2"/><p:play d:shape-id="movie1"/></p:animation-group></p:animations></d:page></o:presentation></o:body></o:document-content>"##;

    let slides = Parser::parse_slides(xml).unwrap();
    let root = slides[0].legacy_animation().unwrap();
    assert_eq!(root.kind(), AnimationKind::Animations);
    assert_eq!(
        root.attribute(&Namespace::Other("urn:example:effects".to_string()), "mode"),
        Some("legacy")
    );
    let group = &root.children()[0];
    assert_eq!(group.kind(), AnimationKind::Group);
    assert_eq!(group.children().len(), 4);
    let show = &group.children()[0];
    assert_eq!(show.kind(), AnimationKind::ShowShape);
    assert_eq!(show.attribute(&Namespace::Draw, "shape-id"), Some("shape1"));
    assert_eq!(show.children()[0].kind(), AnimationKind::Sound);
    assert_eq!(
        show.children()[0].attribute(&Namespace::Xlink, "href"),
        Some("Sounds/a&b.ogg")
    );
}

#[test]
fn rejects_invalid_legacy_presentation_effects() {
    let invalid = [
        r#"<p:show-shape d:shape-id="orphan"/>"#,
        r#"<p:animations><p:show-shape/></p:animations>"#,
        r#"<p:animations><p:dim d:shape-id="s"/></p:animations>"#,
        r#"<p:animations><p:play d:shape-id="s"><p:sound x:href="a" x:type="simple"/></p:play></p:animations>"#,
        r#"<p:animations><p:show-shape d:shape-id="s"><p:sound x:href="a" x:type="extended"/></p:show-shape></p:animations>"#,
        r#"<p:animations>text</p:animations>"#,
    ];
    for effects in invalid {
        let xml = format!(
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:x="http://www.w3.org/1999/xlink"><o:body><o:presentation><d:page>{effects}</d:page></o:presentation></o:body></o:document-content>"#
        );
        assert!(Parser::parse_slides(&xml).is_err(), "accepted {effects}");
    }
}
