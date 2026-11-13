use super::codec::{Limits, opaque, read, write};
use super::{
    Error, Id, MAX_NAMESPACE_TEXT_BYTES, MAX_RELATIONSHIP_ID_BYTES, Metadata, NAMESPACE, Namespace,
    Raster, RasterChild, Reference, Relationship, Resolver, Target, ValueError, validate,
    validate_relationships,
};

const MODEL_XML: &[u8] = br#"<m3d:model3d xmlns:m3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:f="urn:future" r:embed="rIdModel" r:link="rIdLinked"><m3d:spPr><a:futureShape f:value="keep"/></m3d:spPr><m3d:camera/><m3d:trans/><m3d:attrSrcUrl f:source="keep"/><m3d:raster rName="Office Renderer" rVer="17.0"><m3d:blip r:embed="rIdRaster"><m3d:future/></m3d:blip><m3d:future f:payload="yes"/></m3d:raster><m3d:extLst><a:ext uri="urn:future"><f:payload/></a:ext></m3d:extLst><m3d:objViewport/><m3d:ptLight/></m3d:model3d>"#;

#[test]
fn reads_typed_relationships_and_keeps_the_scene_inert() {
    let metadata = read(MODEL_XML, Limits::DEFAULT).expect("model3d fixture must parse");
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
    let metadata = read(MODEL_XML, Limits::DEFAULT).expect("model3d fixture must parse");
    let encoded = write(&metadata, Limits::DEFAULT).expect("model3d fixture must serialize");
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
    let reparsed = read(&encoded, Limits::DEFAULT).expect("serialized model3d must parse");
    assert_eq!(reparsed, metadata);
}

#[test]
fn writer_picks_fresh_prefixes_around_foreign_bindings() {
    let xml = br#"<model3d xmlns="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:m3d="urn:foreign-model" xmlns:r="urn:foreign-relationship" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships" rel:embed="rIdModel"><spPr/><camera/><trans/><objViewport/><m3d:foreign/><r:foreign/></model3d>"#;
    let metadata = read(xml, Limits::DEFAULT).expect("default-prefix scene must parse");
    let unknown = metadata
        .children
        .iter()
        .filter_map(|child| match child {
            super::Child::Opaque(value) if value.namespace().starts_with("urn:") => {
                Some((value.local_name(), value.namespace()))
            },
            super::Child::Raster(_) => None,
            super::Child::Opaque(_) => None,
        })
        .collect::<Vec<_>>();
    assert_eq!(
        unknown,
        vec![
            ("foreign", "urn:foreign-model"),
            ("foreign", "urn:foreign-relationship"),
        ]
    );

    let encoded = write(&metadata, Limits::DEFAULT).expect("fresh-prefix scene must serialize");
    assert!(encoded.starts_with(b"<m3d1:model3d"));
    assert!(
        encoded
            .windows(b"xmlns:m3d1=".len())
            .any(|window| window == b"xmlns:m3d1=")
    );
    assert_eq!(
        encoded
            .windows(b"xmlns:m3d=".len())
            .filter(|window| *window == b"xmlns:m3d=")
            .count(),
        1
    );
    assert_eq!(
        encoded
            .windows(b"xmlns:r=".len())
            .filter(|window| *window == b"xmlns:r=")
            .count(),
        1
    );
    let reparsed = read(&encoded, Limits::DEFAULT).expect("fresh-prefix output must parse");
    let reparsed_unknown = reparsed
        .children
        .iter()
        .filter_map(|child| match child {
            super::Child::Opaque(value) if value.namespace().starts_with("urn:") => {
                Some((value.local_name(), value.namespace()))
            },
            super::Child::Raster(_) => None,
            super::Child::Opaque(_) => None,
        })
        .collect::<Vec<_>>();
    assert_eq!(reparsed_unknown, unknown);

    let mut narrow = Limits::DEFAULT;
    narrow.max_output_bytes = encoded.len() - 1;
    assert!(matches!(
        write(&metadata, narrow),
        Err(crate::Error::Limit {
            kind: crate::LimitKind::OutputBytes,
            observed,
            maximum,
        }) if observed == encoded.len() && maximum + 1 == observed
    ));
}

#[test]
fn rejects_invalid_sequence_and_duplicate_raster_blip() {
    let missing = br#"<m3d:model3d xmlns:m3d="http://schemas.microsoft.com/office/drawing/2017/model3d"><m3d:spPr/><m3d:camera/></m3d:model3d>"#;
    assert!(
        matches!(read(missing, Limits::DEFAULT), Err(crate::Error::Invalid(message)) if message.contains("trans"))
    );

    let duplicate = br#"<m3d:model3d xmlns:m3d="http://schemas.microsoft.com/office/drawing/2017/model3d"><m3d:spPr/><m3d:camera/><m3d:trans/><m3d:raster rName="x" rVer="y"><m3d:blip/><m3d:blip/></m3d:raster><m3d:objViewport/></m3d:model3d>"#;
    assert!(
        matches!(read(duplicate, Limits::DEFAULT), Err(crate::Error::Invalid(message)) if message.contains("blip"))
    );
}

#[test]
fn read_counts_root_and_direct_empty_children_once_operation_wide() {
    let xml = br#"<m3d:model3d xmlns:m3d="http://schemas.microsoft.com/office/drawing/2017/model3d"><m3d:spPr/><m3d:camera/><m3d:trans/><m3d:objViewport/></m3d:model3d>"#;
    let mut limits = Limits::DEFAULT;
    limits.max_nodes = 4;
    assert!(matches!(
        read(xml, limits),
        Err(crate::Error::Limit {
            kind: crate::LimitKind::XmlNodes,
            ..
        })
    ));
}

#[test]
fn read_admits_an_opaque_child_before_copying_it() {
    let xml = br#"<m3d:model3d xmlns:m3d="http://schemas.microsoft.com/office/drawing/2017/model3d"><m3d:spPr/><m3d:camera/><m3d:trans/><m3d:objViewport/><f:future xmlns:f="urn:future">retained</f:future></m3d:model3d>"#;
    let mut limits = Limits::DEFAULT;
    limits.max_fragment_bytes = 16;
    assert!(matches!(
        read(xml, limits),
        Err(crate::Error::Limit {
            kind: crate::LimitKind::InertFragmentBytes,
            ..
        })
    ));
}

#[test]
fn relationship_graph_distinguishes_embedded_linked_and_missing_targets() {
    let metadata = read(MODEL_XML, Limits::DEFAULT).expect("model3d fixture must parse");
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
    let value = opaque(
        br#"<f:future xmlns:f="urn:future"><f:nested/></f:future>"#,
        Limits::DEFAULT,
    )
    .expect("self-contained inert element");
    assert_eq!(value.local_name(), "future");
    assert_eq!(value.namespace(), "urn:future");
    assert!(
        opaque(
            b"<f:one xmlns:f=\"urn:f\"/><f:two xmlns:f=\"urn:f\"/>",
            Limits::DEFAULT
        )
        .is_err()
    );
}

#[test]
fn authoring_facade_can_replace_raster_without_losing_scene_children() {
    let mut metadata = read(MODEL_XML, Limits::DEFAULT).expect("model3d fixture must parse");
    let mut raster = Raster::new("Renderer", "1").unwrap();
    raster.children.push(RasterChild::Blip(super::Blip {
        reference: Reference::embedded("rIdRaster").unwrap(),
        children: Vec::new(),
        namespaces: Vec::new(),
    }));
    metadata.set_raster(raster);
    let encoded = write(&metadata, Limits::DEFAULT).expect("edited model3d must serialize");
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

#[test]
fn scalar_constructor_bounds_keep_typed_limit_dimensions() {
    let relationship_id = "r".repeat(MAX_RELATIONSHIP_ID_BYTES + 1);
    assert!(matches!(
        Id::new(&relationship_id),
        Err(ValueError::Limit {
            kind: crate::LimitKind::AttributeValueBytes,
            observed,
            maximum: MAX_RELATIONSHIP_ID_BYTES,
        }) if observed == relationship_id.len()
    ));

    let prefix = "p".repeat(MAX_NAMESPACE_TEXT_BYTES + 1);
    assert!(matches!(
        Namespace::new(Some(&prefix), "urn:future"),
        Err(ValueError::Limit {
            kind: crate::LimitKind::NamespacePrefixBytes,
            observed,
            maximum: MAX_NAMESPACE_TEXT_BYTES,
        }) if observed == prefix.len()
    ));
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

/// Scene children fed to the ordering matrices, in schema order. `raster` is
/// the typed child; `future` is an element in a foreign namespace.
const SCENE_CHILDREN: [&str; 11] = [
    "spPr",
    "camera",
    "trans",
    "attrSrcUrl",
    "raster",
    "extLst",
    "objViewport",
    "winViewport",
    "ambientLight",
    "ptLight",
    "future",
];

const THROUGH_TRANSFORM: [&str; 3] = ["spPr", "camera", "trans"];
const THROUGH_VIEWPORT: [&str; 4] = ["spPr", "camera", "trans", "objViewport"];

/// Build one scene child. `opaque:NAME` is an inert element in the model
/// namespace, `foreign:NAME` an inert element in a foreign namespace.
fn scene_child(name: &str) -> super::Child {
    if name == "raster" {
        return super::Child::Raster(Raster::new("Renderer", "1").unwrap());
    }
    let xml = if name == "future" {
        r#"<f:future xmlns:f="urn:future"/>"#.to_string()
    } else if let Some(local) = name.strip_prefix("foreign:") {
        format!(r#"<f:{local} xmlns:f="urn:future"/>"#)
    } else {
        let local = name.strip_prefix("opaque:").unwrap_or(name);
        format!(r#"<m:{local} xmlns:m="{NAMESPACE}"/>"#)
    };
    super::Child::Opaque(opaque(xml.as_bytes(), Limits::DEFAULT).unwrap())
}

/// Validate a scene and name the exact result: `ok`, or the error variant
/// with its payload.
fn scene_outcome(names: &[&str]) -> String {
    let mut metadata = Metadata::new(Reference::none());
    metadata.children = names.iter().map(|name| scene_child(name)).collect();
    match validate(&metadata) {
        Ok(()) => "ok".to_string(),
        Err(Error::ChildOrder { child }) => format!("order({child})"),
        Err(Error::MissingChild(child)) => format!("missing({child})"),
        Err(Error::DuplicateChild(child)) => format!("duplicate({child})"),
        Err(Error::MultipleViewports) => "viewports".to_string(),
        Err(other) => panic!("unexpected scene validation error {other:?}"),
    }
}

/// One row per first child: the outcomes of `prefix + [first, second]` for
/// every second child of [`SCENE_CHILDREN`], in that order.
fn scene_pair_grid(prefix: &[&str]) -> Vec<String> {
    SCENE_CHILDREN
        .iter()
        .map(|first| {
            let outcomes: Vec<String> = SCENE_CHILDREN
                .iter()
                .map(|second| {
                    let mut names = prefix.to_vec();
                    names.push(first);
                    names.push(second);
                    scene_outcome(&names)
                })
                .collect();
            format!("{first} => {}", outcomes.join(" | "))
        })
        .collect()
}

fn scene_single_row(prefix: &[&str]) -> String {
    let outcomes: Vec<String> = SCENE_CHILDREN
        .iter()
        .map(|child| {
            let mut names = prefix.to_vec();
            names.push(child);
            scene_outcome(&names)
        })
        .collect();
    outcomes.join(" | ")
}

fn assert_scene_grid(prefix: &[&str], expected: &[&str]) {
    let actual = scene_pair_grid(prefix);
    assert_eq!(actual, expected, "{actual:#?}");
}

#[test]
fn scene_order_of_every_child_pair_in_an_empty_scene() {
    assert_scene_grid(
        &[],
        &[
            "spPr => order(spPr) | missing(trans) | missing(camera) | order(attrSrcUrl) | order(raster) | order(extLst) | order(viewport) | order(viewport) | order(ambientLight) | order(light) | missing(camera)",
            "camera => order(spPr) | order(camera) | missing(spPr) | order(attrSrcUrl) | order(raster) | order(extLst) | order(viewport) | order(viewport) | order(ambientLight) | order(light) | missing(spPr)",
            "trans => order(spPr) | order(camera) | order(trans) | missing(spPr) | missing(spPr) | missing(spPr) | missing(spPr) | missing(spPr) | order(ambientLight) | order(light) | missing(spPr)",
            "attrSrcUrl => order(attrSrcUrl) | order(attrSrcUrl) | order(attrSrcUrl) | order(attrSrcUrl) | order(attrSrcUrl) | order(attrSrcUrl) | order(attrSrcUrl) | order(attrSrcUrl) | order(attrSrcUrl) | order(attrSrcUrl) | order(attrSrcUrl)",
            "raster => order(raster) | order(raster) | order(raster) | order(raster) | order(raster) | order(raster) | order(raster) | order(raster) | order(raster) | order(raster) | order(raster)",
            "extLst => order(extLst) | order(extLst) | order(extLst) | order(extLst) | order(extLst) | order(extLst) | order(extLst) | order(extLst) | order(extLst) | order(extLst) | order(extLst)",
            "objViewport => order(viewport) | order(viewport) | order(viewport) | order(viewport) | order(viewport) | order(viewport) | order(viewport) | order(viewport) | order(viewport) | order(viewport) | order(viewport)",
            "winViewport => order(viewport) | order(viewport) | order(viewport) | order(viewport) | order(viewport) | order(viewport) | order(viewport) | order(viewport) | order(viewport) | order(viewport) | order(viewport)",
            "ambientLight => order(ambientLight) | order(ambientLight) | order(ambientLight) | order(ambientLight) | order(ambientLight) | order(ambientLight) | order(ambientLight) | order(ambientLight) | order(ambientLight) | order(ambientLight) | order(ambientLight)",
            "ptLight => order(light) | order(light) | order(light) | order(light) | order(light) | order(light) | order(light) | order(light) | order(light) | order(light) | order(light)",
            "future => missing(camera) | missing(spPr) | missing(spPr) | order(attrSrcUrl) | order(raster) | order(extLst) | order(viewport) | order(viewport) | order(ambientLight) | order(light) | missing(spPr)",
        ],
    );
}

#[test]
fn scene_order_of_every_child_pair_after_the_transform() {
    assert_scene_grid(
        &THROUGH_TRANSFORM,
        &[
            "spPr => order(spPr) | order(spPr) | order(spPr) | order(spPr) | order(spPr) | order(spPr) | order(spPr) | order(spPr) | order(spPr) | order(spPr) | order(spPr)",
            "camera => order(camera) | order(camera) | order(camera) | order(camera) | order(camera) | order(camera) | order(camera) | order(camera) | order(camera) | order(camera) | order(camera)",
            "trans => order(trans) | order(trans) | order(trans) | order(trans) | order(trans) | order(trans) | order(trans) | order(trans) | order(trans) | order(trans) | order(trans)",
            "attrSrcUrl => order(spPr) | order(camera) | order(trans) | missing(objViewport or winViewport) | missing(objViewport or winViewport) | missing(objViewport or winViewport) | ok | ok | order(ambientLight) | order(light) | missing(objViewport or winViewport)",
            "raster => order(spPr) | order(camera) | order(trans) | order(attrSrcUrl) | duplicate(raster) | missing(objViewport or winViewport) | ok | ok | order(ambientLight) | order(light) | missing(objViewport or winViewport)",
            "extLst => order(spPr) | order(camera) | order(trans) | order(attrSrcUrl) | order(raster) | order(extLst) | ok | ok | order(ambientLight) | order(light) | missing(objViewport or winViewport)",
            "objViewport => order(spPr) | order(camera) | order(trans) | order(attrSrcUrl) | order(raster) | order(extLst) | viewports | viewports | ok | ok | ok",
            "winViewport => order(spPr) | order(camera) | order(trans) | order(attrSrcUrl) | order(raster) | order(extLst) | viewports | viewports | ok | ok | ok",
            "ambientLight => order(ambientLight) | order(ambientLight) | order(ambientLight) | order(ambientLight) | order(ambientLight) | order(ambientLight) | order(ambientLight) | order(ambientLight) | order(ambientLight) | order(ambientLight) | order(ambientLight)",
            "ptLight => order(light) | order(light) | order(light) | order(light) | order(light) | order(light) | order(light) | order(light) | order(light) | order(light) | order(light)",
            "future => order(spPr) | order(camera) | order(trans) | missing(objViewport or winViewport) | missing(objViewport or winViewport) | missing(objViewport or winViewport) | ok | ok | order(ambientLight) | order(light) | missing(objViewport or winViewport)",
        ],
    );
}

#[test]
fn scene_order_of_every_child_pair_after_the_viewport() {
    assert_scene_grid(
        &THROUGH_VIEWPORT,
        &[
            "spPr => order(spPr) | order(spPr) | order(spPr) | order(spPr) | order(spPr) | order(spPr) | order(spPr) | order(spPr) | order(spPr) | order(spPr) | order(spPr)",
            "camera => order(camera) | order(camera) | order(camera) | order(camera) | order(camera) | order(camera) | order(camera) | order(camera) | order(camera) | order(camera) | order(camera)",
            "trans => order(trans) | order(trans) | order(trans) | order(trans) | order(trans) | order(trans) | order(trans) | order(trans) | order(trans) | order(trans) | order(trans)",
            "attrSrcUrl => order(attrSrcUrl) | order(attrSrcUrl) | order(attrSrcUrl) | order(attrSrcUrl) | order(attrSrcUrl) | order(attrSrcUrl) | order(attrSrcUrl) | order(attrSrcUrl) | order(attrSrcUrl) | order(attrSrcUrl) | order(attrSrcUrl)",
            "raster => order(raster) | order(raster) | order(raster) | order(raster) | order(raster) | order(raster) | order(raster) | order(raster) | order(raster) | order(raster) | order(raster)",
            "extLst => order(extLst) | order(extLst) | order(extLst) | order(extLst) | order(extLst) | order(extLst) | order(extLst) | order(extLst) | order(extLst) | order(extLst) | order(extLst)",
            "objViewport => viewports | viewports | viewports | viewports | viewports | viewports | viewports | viewports | viewports | viewports | viewports",
            "winViewport => viewports | viewports | viewports | viewports | viewports | viewports | viewports | viewports | viewports | viewports | viewports",
            "ambientLight => order(spPr) | order(camera) | order(trans) | order(attrSrcUrl) | order(raster) | order(extLst) | viewports | viewports | order(ambientLight) | ok | ok",
            "ptLight => order(spPr) | order(camera) | order(trans) | order(attrSrcUrl) | order(raster) | order(extLst) | viewports | viewports | ok | ok | ok",
            "future => order(spPr) | order(camera) | order(trans) | order(attrSrcUrl) | order(raster) | order(extLst) | viewports | viewports | ok | ok | ok",
        ],
    );
}

#[test]
fn scene_order_of_each_child_alone_and_of_the_empty_scene() {
    // For each prefix: the prefix alone, then the prefix followed by each
    // child of `SCENE_CHILDREN` in order.
    let actual = [
        scene_outcome(&[]),
        scene_single_row(&[]),
        scene_outcome(&THROUGH_TRANSFORM),
        scene_single_row(&THROUGH_TRANSFORM),
        scene_outcome(&THROUGH_VIEWPORT),
        scene_single_row(&THROUGH_VIEWPORT),
    ];
    let expected: [&str; 6] = [
        "missing(spPr)",
        "missing(camera) | missing(spPr) | missing(spPr) | order(attrSrcUrl) | order(raster) | order(extLst) | order(viewport) | order(viewport) | order(ambientLight) | order(light) | missing(spPr)",
        "missing(objViewport or winViewport)",
        "order(spPr) | order(camera) | order(trans) | missing(objViewport or winViewport) | missing(objViewport or winViewport) | missing(objViewport or winViewport) | ok | ok | order(ambientLight) | order(light) | missing(objViewport or winViewport)",
        "ok",
        "order(spPr) | order(camera) | order(trans) | order(attrSrcUrl) | order(raster) | order(extLst) | viewports | viewports | ok | ok | ok",
    ];
    assert_eq!(actual, expected, "{actual:#?}");
}

#[test]
fn scene_order_of_longer_sequences() {
    let rows: &[(&[&str], &str)] = &[
        // The complete sequence, with every optional child and every light kind.
        (
            &[
                "spPr",
                "camera",
                "trans",
                "attrSrcUrl",
                "raster",
                "extLst",
                "winViewport",
                "ambientLight",
                "ptLight",
                "spotLight",
                "dirLight",
                "unkLight",
            ],
            "ok",
        ),
        // Each light kind after and before the viewport.
        (
            &["spPr", "camera", "trans", "objViewport", "spotLight"],
            "ok",
        ),
        (
            &["spPr", "camera", "trans", "objViewport", "dirLight"],
            "ok",
        ),
        (
            &["spPr", "camera", "trans", "objViewport", "unkLight"],
            "ok",
        ),
        (&["spPr", "camera", "trans", "spotLight"], "order(light)"),
        (
            &["spPr", "camera", "trans", "dirLight", "objViewport"],
            "order(light)",
        ),
        (
            &["spPr", "camera", "trans", "unkLight", "objViewport"],
            "order(light)",
        ),
        // The ambient light stays single across other lights.
        (
            &[
                "spPr",
                "camera",
                "trans",
                "objViewport",
                "ambientLight",
                "ptLight",
                "ambientLight",
            ],
            "order(ambientLight)",
        ),
        (
            &[
                "spPr",
                "camera",
                "trans",
                "objViewport",
                "ptLight",
                "ambientLight",
                "ptLight",
                "spotLight",
            ],
            "ok",
        ),
        // Earlier children after a light.
        (
            &[
                "spPr",
                "camera",
                "trans",
                "objViewport",
                "ptLight",
                "winViewport",
            ],
            "viewports",
        ),
        (
            &[
                "spPr",
                "camera",
                "trans",
                "objViewport",
                "ptLight",
                "extLst",
            ],
            "order(extLst)",
        ),
        (
            &[
                "spPr",
                "camera",
                "trans",
                "objViewport",
                "ptLight",
                "raster",
            ],
            "order(raster)",
        ),
        (
            &[
                "spPr",
                "camera",
                "trans",
                "objViewport",
                "ptLight",
                "attrSrcUrl",
            ],
            "order(attrSrcUrl)",
        ),
        (
            &["spPr", "camera", "trans", "objViewport", "ptLight", "trans"],
            "order(trans)",
        ),
        (
            &[
                "spPr",
                "camera",
                "trans",
                "objViewport",
                "ptLight",
                "camera",
            ],
            "order(camera)",
        ),
        (
            &["spPr", "camera", "trans", "objViewport", "ptLight", "spPr"],
            "order(spPr)",
        ),
        // Repeats and reversals among the optional children.
        (
            &[
                "spPr",
                "camera",
                "trans",
                "attrSrcUrl",
                "attrSrcUrl",
                "attrSrcUrl",
                "objViewport",
            ],
            "ok",
        ),
        (
            &[
                "spPr",
                "camera",
                "trans",
                "raster",
                "attrSrcUrl",
                "objViewport",
            ],
            "order(attrSrcUrl)",
        ),
        (
            &["spPr", "camera", "trans", "extLst", "raster", "objViewport"],
            "order(raster)",
        ),
        (
            &[
                "spPr",
                "camera",
                "trans",
                "extLst",
                "future",
                "extLst",
                "objViewport",
            ],
            "order(extLst)",
        ),
        (
            &["spPr", "camera", "trans", "raster", "future", "raster"],
            "duplicate(raster)",
        ),
        // An inert `raster` element in the model namespace.
        (
            &[
                "spPr",
                "camera",
                "trans",
                "opaque:raster",
                "opaque:raster",
                "objViewport",
            ],
            "ok",
        ),
        (
            &[
                "spPr",
                "camera",
                "trans",
                "raster",
                "opaque:raster",
                "objViewport",
            ],
            "ok",
        ),
        (
            &[
                "spPr",
                "camera",
                "trans",
                "opaque:raster",
                "raster",
                "objViewport",
            ],
            "ok",
        ),
        (
            &[
                "spPr",
                "camera",
                "trans",
                "extLst",
                "opaque:raster",
                "objViewport",
            ],
            "order(raster)",
        ),
        // Known local names in a foreign namespace are inert.
        (
            &["foreign:spPr", "camera", "trans", "objViewport"],
            "missing(spPr)",
        ),
        (
            &["foreign:spPr", "spPr", "camera", "trans", "objViewport"],
            "ok",
        ),
        (
            &[
                "future",
                "spPr",
                "future",
                "camera",
                "future",
                "trans",
                "future",
                "objViewport",
                "future",
            ],
            "ok",
        ),
        // Missing and misplaced required children.
        (&["spPr", "trans", "objViewport"], "missing(camera)"),
        (&["camera", "trans", "objViewport"], "missing(spPr)"),
        (&["trans", "objViewport"], "missing(spPr)"),
        (&["spPr", "camera", "objViewport"], "order(viewport)"),
        (&["camera", "spPr", "trans", "objViewport"], "order(spPr)"),
        (&["trans", "camera"], "order(camera)"),
        (&["spPr", "trans", "camera"], "order(camera)"),
        (&["spPr", "camera", "spPr"], "order(spPr)"),
        (&["spPr", "camera", "camera"], "order(camera)"),
        (&["spPr", "camera", "trans", "trans"], "order(trans)"),
        (&["spPr", "camera", "trans", "camera"], "order(camera)"),
        (&["spPr", "camera", "trans", "spPr"], "order(spPr)"),
        // The typed raster is checked for duplicates before its position.
        (&["raster", "raster"], "order(raster)"),
        (&["raster", "spPr", "raster"], "order(raster)"),
    ];
    let actual: Vec<String> = rows.iter().map(|(names, _)| scene_outcome(names)).collect();
    let expected: Vec<&str> = rows.iter().map(|(_, outcome)| *outcome).collect();
    assert_eq!(actual, expected, "{actual:#?}");
}

#[test]
fn scene_order_error_messages() {
    assert_eq!(
        Error::ChildOrder { child: "viewport" }.to_string(),
        "model3d child 'viewport' is out of sequence"
    );
    assert_eq!(
        Error::MissingChild("objViewport or winViewport").to_string(),
        "model3d required child 'objViewport or winViewport' is missing"
    );
    assert_eq!(
        Error::DuplicateChild("raster").to_string(),
        "model3d child 'raster' occurs more than once"
    );
    assert_eq!(
        Error::MultipleViewports.to_string(),
        "model3d object and window viewports cannot both be present"
    );
}
