use super::codec::{opaque, read, write};
use super::{
    Error, Id, Metadata, Raster, RasterChild, Reference, Relationship, Resolver, Target,
    validate_relationships,
};

const MODEL_XML: &[u8] = br#"<m3d:model3d xmlns:m3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:f="urn:future" r:embed="rIdModel" r:link="rIdLinked"><m3d:spPr><a:futureShape f:value="keep"/></m3d:spPr><m3d:camera/><m3d:trans/><m3d:attrSrcUrl f:source="keep"/><m3d:raster rName="Office Renderer" rVer="17.0"><m3d:blip r:embed="rIdRaster"><m3d:future/></m3d:blip><m3d:future f:payload="yes"/></m3d:raster><m3d:extLst><a:ext uri="urn:future"><f:payload/></a:ext></m3d:extLst><m3d:objViewport/><m3d:ptLight/></m3d:model3d>"#;

#[test]
fn reads_typed_relationships_and_keeps_the_scene_inert() {
    let metadata = read(MODEL_XML).expect("model3d fixture must parse");
    assert_eq!(
        metadata.reference.embedded.as_ref().unwrap().as_str(),
        "rIdModel"
    );
    assert_eq!(
        metadata.reference.linked.as_ref().unwrap().as_str(),
        "rIdLinked"
    );
    let raster = metadata.raster().expect("raster metadata");
    assert_eq!(raster.renderer_name.as_ref(), "Office Renderer");
    let RasterChild::Blip(blip) = &raster.children[0] else {
        panic!("expected typed raster blip");
    };
    assert_eq!(
        blip.reference.embedded.as_ref().unwrap().as_str(),
        "rIdRaster"
    );
    assert_eq!(blip.children[0].local_name(), "future");
    assert_eq!(metadata.children.len(), 8);
}

#[test]
fn preserves_unknown_children_and_namespace_declarations_on_round_trip() {
    let metadata = read(MODEL_XML).expect("model3d fixture must parse");
    let encoded = write(&metadata).expect("model3d fixture must serialize");
    assert!(
        encoded
            .windows(b"f:payload".len())
            .any(|window| window == b"f:payload")
    );
    assert!(
        encoded
            .windows(b"f:value".len())
            .any(|window| window == b"f:value")
    );
    let reparsed = read(&encoded).expect("serialized model3d must parse");
    assert_eq!(reparsed, metadata);
}

#[test]
fn rejects_invalid_sequence_and_duplicate_raster_blip() {
    let missing = br#"<m3d:model3d xmlns:m3d="http://schemas.microsoft.com/office/drawing/2017/model3d"><m3d:spPr/><m3d:camera/></m3d:model3d>"#;
    assert!(
        matches!(read(missing), Err(crate::Error::Invalid(message)) if message.contains("trans"))
    );

    let duplicate = br#"<m3d:model3d xmlns:m3d="http://schemas.microsoft.com/office/drawing/2017/model3d"><m3d:spPr/><m3d:camera/><m3d:trans/><m3d:raster rName="x" rVer="y"><m3d:blip/><m3d:blip/></m3d:raster><m3d:objViewport/></m3d:model3d>"#;
    assert!(
        matches!(read(duplicate), Err(crate::Error::Invalid(message)) if message.contains("blip"))
    );
}

#[test]
fn relationship_graph_distinguishes_embedded_linked_and_missing_targets() {
    let metadata = read(MODEL_XML).expect("model3d fixture must parse");
    let graph = FixtureGraph;
    validate_relationships(&metadata, &graph).expect("fixture graph is valid");

    let mut missing = metadata.clone();
    missing.reference.embedded = Some(Id::new("rIdMissing").unwrap());
    assert!(matches!(
        validate_relationships(&missing, &graph),
        Err(Error::MissingRelationship { field: "model", .. })
    ));

    let mut wrong_mode = metadata;
    wrong_mode.reference.linked = Some(Id::new("rIdInternal").unwrap());
    assert!(matches!(
        validate_relationships(&wrong_mode, &graph),
        Err(Error::LinkedTargetIsInternal { field: "model", .. })
    ));
}

#[test]
fn opaque_constructor_accepts_one_element_and_rejects_multiple_roots() {
    let value = opaque(br#"<f:future xmlns:f="urn:future"><f:nested/></f:future>"#)
        .expect("self-contained inert element");
    assert_eq!(value.local_name(), "future");
    assert_eq!(value.namespace(), "urn:future");
    assert!(opaque(b"<f:one xmlns:f=\"urn:f\"/><f:two xmlns:f=\"urn:f\"/>").is_err());
}

#[test]
fn authoring_facade_can_replace_raster_without_losing_scene_children() {
    let mut metadata = read(MODEL_XML).expect("model3d fixture must parse");
    let mut raster = Raster::new("Renderer", "1").unwrap();
    raster.children.push(RasterChild::Blip(super::Blip {
        reference: Reference::embedded("rIdRaster").unwrap(),
        children: Vec::new(),
        namespaces: Vec::new(),
    }));
    metadata.set_raster(raster);
    let encoded = write(&metadata).expect("edited model3d must serialize");
    assert!(
        encoded
            .windows(b"f:payload".len())
            .any(|window| window == b"f:payload")
    );
    assert!(
        encoded
            .windows(b"Renderer".len())
            .any(|window| window == b"Renderer")
    );
}

struct FixtureGraph;

impl Resolver for FixtureGraph {
    fn relationship<'a>(&'a self, id: &Id) -> Option<Relationship<'a>> {
        match id.as_str() {
            "rIdModel" | "rIdRaster" | "rIdInternal" => Some(Relationship {
                relationship_type: "urn:test:internal",
                target: Target::Internal("/word/media/model.glb"),
            }),
            "rIdLinked" => Some(Relationship {
                relationship_type: "urn:test:external",
                target: Target::External("https://example.test/model.glb"),
            }),
            _ => None,
        }
    }
}

#[allow(dead_code)]
fn _metadata_is_constructible() -> Metadata {
    Metadata::new(Reference::none())
}
