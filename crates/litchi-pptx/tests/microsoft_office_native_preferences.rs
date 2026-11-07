//! Native-source presentation metadata operations; no playback or native UI validation.
use litchi_drawingml::coordinate::Coordinate;
use litchi_opc::PackURI;
use litchi_pptx::presentation_properties::metadata::events::{self, Draft, Trigger};
use litchi_pptx::{
    Error, Package, laser,
    tag::{List, Tag},
    time::Offset,
};
use std::{collections::BTreeMap, io::Cursor};

const FILES: [&str; 6] = [
    "powerpoint-visual-objects.pptx",
    "powerpoint-visual-objects.ppsx",
    "powerpoint-visual-objects.potx",
    "powerpoint-passive-macro.pptm",
    "powerpoint-passive-macro.ppsm",
    "powerpoint-passive-macro.potm",
];
const SLIDE: &str = "ppt/slides/slide1.xml";
const PROPS: &str = "ppt/presProps.xml";
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/powerpoint")
            .join(name),
    )
    .unwrap()
}
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let r = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    r.member_names()
        .unwrap()
        .into_iter()
        .map(|n| {
            let b = r.read_member(&n).unwrap();
            (n, b)
        })
        .collect()
}
fn changed(a: &[u8], b: &[u8]) -> Vec<String> {
    let a = members(a);
    let b = members(b);
    let names: std::collections::BTreeSet<_> = a.keys().chain(b.keys()).collect();
    names
        .into_iter()
        .filter(|n| a.get(*n) != b.get(*n))
        .cloned()
        .collect()
}
fn fragment(xml: &[u8], name: &[u8]) -> Vec<u8> {
    use quick_xml::events::Event;
    let mut r = quick_xml::Reader::from_reader(xml);
    let mut start = None;
    let mut depth = 0usize;
    loop {
        let at = r.buffer_position() as usize;
        let e = r.read_event().unwrap();
        let end = r.buffer_position() as usize;
        match e {
            Event::Start(e) => {
                if start.is_none() && e.name().as_ref() == name {
                    start = Some((at, depth));
                }
                depth += 1;
            },
            Event::Empty(e) if start.is_none() && e.name().as_ref() == name => {
                return xml[at..end].to_vec();
            },
            Event::End(_) => {
                depth -= 1;
                if start.is_some_and(|(_, d)| d == depth) {
                    return xml[start.unwrap().0..end].to_vec();
                }
            },
            Event::Eof => panic!("missing XML fragment"),
            _ => {},
        }
    }
}
fn graph_elements(xml: &[u8]) -> Vec<(String, BTreeMap<String, String>)> {
    use quick_xml::events::Event;
    let mut r = quick_xml::Reader::from_reader(xml);
    let mut out = Vec::new();
    loop {
        match r.read_event().unwrap() {
            Event::Start(e) | Event::Empty(e) => {
                let attrs = e
                    .attributes()
                    .map(|a| {
                        let a = a.unwrap();
                        (
                            String::from_utf8(a.key.as_ref().to_vec()).unwrap(),
                            a.decoded_and_normalized_value(
                                quick_xml::XmlVersion::Explicit1_0,
                                r.decoder(),
                            )
                            .unwrap()
                            .into_owned(),
                        )
                    })
                    .collect();
                out.push((
                    String::from_utf8(e.name().as_ref().to_vec()).unwrap(),
                    attrs,
                ));
            },
            Event::Eof => break,
            Event::Decl(_) | Event::End(_) => {},
            Event::Text(t) => assert!(t.as_ref().iter().all(u8::is_ascii_whitespace)),
            _ => panic!("unexpected graph markup"),
        }
    }
    out.sort();
    out
}
fn content_types(parts: &BTreeMap<String, Vec<u8>>) -> BTreeMap<String, String> {
    let mut defaults = BTreeMap::new();
    let mut overrides = BTreeMap::new();
    for (name, attrs) in graph_elements(&parts["[Content_Types].xml"]) {
        match name.as_str() {
            "Default" => {
                assert_eq!(attrs.len(), 2);
                assert!(
                    defaults
                        .insert(attrs["Extension"].clone(), attrs["ContentType"].clone())
                        .is_none()
                );
            },
            "Override" => {
                assert_eq!(attrs.len(), 2);
                assert!(
                    overrides
                        .insert(attrs["PartName"].clone(), attrs["ContentType"].clone())
                        .is_none()
                );
            },
            "Types" => assert_eq!(
                attrs,
                BTreeMap::from([(
                    "xmlns".into(),
                    "http://schemas.openxmlformats.org/package/2006/content-types".into()
                )])
            ),
            _ => panic!("unexpected content-type element: {name}"),
        }
    }
    parts
        .keys()
        .filter(|n| *n != "[Content_Types].xml")
        .map(|n| {
            let kind = overrides
                .get(&format!("/{n}"))
                .or_else(|| defaults.get(n.rsplit('.').next().unwrap()))
                .unwrap();
            (n.clone(), kind.clone())
        })
        .collect()
}
fn assert_same_payloads_and_graph(before: &[u8], after: &[u8]) {
    let a = members(before);
    let b = members(after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    // A default and per-part overrides can describe the same OPC typing.
    // Check every member's resolved type, not that serialization choice.
    assert_eq!(content_types(&a), content_types(&b));
    for (n, data) in a {
        if n == "[Content_Types].xml" {
            continue;
        }
        if n.ends_with(".rels") {
            assert_eq!(graph_elements(&data), graph_elements(&b[&n]), "{n}");
        } else {
            assert_eq!(data, b[&n], "{n}");
        }
    }
}

fn save(p: &Package) -> Package {
    let file = tempfile::NamedTempFile::with_suffix(".pptx").unwrap();
    p.save(file.path()).unwrap();
    Package::open(file.path()).unwrap()
}
fn tags(value: &str) -> List {
    let mut tags = List::new();
    tags.add(Tag::new("Review", value).unwrap()).unwrap();
    tags.add(Tag::new("Literal", "<inert & 日本語>").unwrap())
        .unwrap();
    tags
}
fn traces(p: &Package) -> Vec<laser::Trace> {
    traces_on(p, 0, SLIDE)
}
fn traces_on(p: &Package, index: usize, part: &str) -> Vec<laser::Trace> {
    laser::load_slide_traces(
        index,
        p.opc()
            .unwrap()
            .get_part(&PackURI::new(format!("/{part}")).unwrap())
            .unwrap(),
        &laser::Limits::default(),
        &mut laser::Usage::default(),
    )
    .unwrap()
}
fn points() -> Vec<laser::TracePoint> {
    [(0, 0, 0), (250, 914400, 457200), (500, -1, 1000)]
        .into_iter()
        .map(|(t, x, y)| {
            laser::TracePoint::new(
                Offset::ms(t),
                Coordinate::emu(x).unwrap(),
                Coordinate::emu(y).unwrap(),
            )
        })
        .collect()
}
fn event_values(object: u32) -> Vec<Draft> {
    vec![
        Draft::trigger(Trigger::OnClick, Offset::ms(100), object),
        Draft::play(Offset::ms(200), object),
        Draft::pause(Offset::ms(300), object),
        Draft::seek(Offset::ms(400), object, Offset::ms(50)),
        Draft::resume(Offset::ms(500), object),
        Draft::stop(Offset::ms(600), object),
        Draft::null(Offset::ms(700), object),
    ]
}

#[test]
fn native_default_text_style_bytes_survive_independent_property_edits() {
    for name in FILES {
        let bytes = native(name);
        let p = Package::from_reader(Cursor::new(&bytes)).unwrap();
        let parts = members(&bytes);
        // Borrowed-buffer ingress may re-encode the ZIP container; it still
        // retains every member payload. Reader ingress retains the container.
        assert!(members(&Package::from_bytes(&bytes).unwrap().to_bytes().unwrap()) == parts);
        // These are explicit raw-owner assertions, not a typed Kinsoku API.
        let style = fragment(&parts["ppt/presentation.xml"], b"p:defaultTextStyle");
        let nodes = graph_elements(&style);
        for level in 1..=9 {
            let attrs = &nodes
                .iter()
                .find(|(n, _)| n == &format!("a:lvl{level}pPr"))
                .unwrap()
                .1;
            assert_eq!(attrs["marL"], ((level - 1) * 457200).to_string());
            for (key, value) in [
                ("algn", "l"),
                ("defTabSz", "914400"),
                ("rtl", "0"),
                ("eaLnBrk", "1"),
                ("latinLnBrk", "0"),
                ("hangingPunct", "1"),
            ] {
                assert_eq!(attrs[key], value);
            }
        }
        assert_eq!(
            nodes
                .iter()
                .filter(|(n, a)| n == "a:defRPr"
                    && a.get("sz").is_some_and(|v| v == "1800")
                    && a.get("kern").is_some_and(|v| v == "1200"))
                .count(),
            9
        );
        assert_eq!(
            nodes
                .iter()
                .filter(|(n, a)| n == "a:latin" && a.get("typeface").is_some_and(|v| v == "+mn-lt"))
                .count(),
            9
        );
        let snapshot = p.read_only_recommended_snapshot().unwrap().unwrap();
        assert_eq!(snapshot.value(), None);
        assert_eq!(snapshot.source_xml(), parts[PROPS]);
        assert_eq!(
            p.edit_read_only_recommended()
                .unwrap()
                .commit()
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            bytes
        );
        let mut e = p.edit_read_only_recommended().unwrap();
        e.set(true).unwrap();
        let c = e.commit().unwrap();
        let after = save(c.package());
        let out = members(&after.to_bytes().unwrap());
        assert_eq!(out["ppt/presentation.xml"], parts["ppt/presentation.xml"]);
        for child in [
            b"p14:discardImageEditData".as_slice(),
            b"p14:defaultImageDpi",
            b"p15:chartTrackingRefBased",
        ] {
            assert_eq!(fragment(&out[PROPS], child), fragment(&parts[PROPS], child));
        }
        assert_eq!(changed(&bytes, &after.to_bytes().unwrap()), [PROPS]);
    }
}

#[test]
fn native_readonly_recommendation_crud_is_inert_reversible_and_source_checked() {
    for name in FILES {
        let bytes = native(name);
        let p = Package::from_reader(Cursor::new(&bytes)).unwrap();
        let mut e = p.edit_read_only_recommended().unwrap();
        assert_eq!(e.remove().unwrap(), None);
        assert_eq!(e.set(true).unwrap(), None);
        let created = e.commit().unwrap();
        let after = save(created.package());
        assert_eq!(after.read_only_recommended().unwrap(), Some(true));
        assert_same_payloads_and_graph(
            &after.to_bytes().unwrap(),
            &p.apply_read_only_recommended(created.patch())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
        );
        assert!(matches!(
            after.apply_read_only_recommended(created.patch()),
            Err(Error::PatchConflict)
        ));
        assert_same_payloads_and_graph(
            &bytes,
            &after
                .apply_read_only_recommended(&created.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
        );
        let mut e = after.edit_read_only_recommended().unwrap();
        assert_eq!(e.set(true).unwrap(), Some(true));
        assert_eq!(
            e.commit().unwrap().package().to_bytes().unwrap(),
            after.to_bytes().unwrap()
        );
        let mut e = after.edit_read_only_recommended().unwrap();
        assert_eq!(e.set(false).unwrap(), Some(true));
        let updated = e.commit().unwrap();
        let off = save(updated.package());
        assert_eq!(off.read_only_recommended().unwrap(), Some(false));
        let mut e = off.edit_read_only_recommended().unwrap();
        assert_eq!(e.remove().unwrap(), Some(false));
        assert_eq!(e.remove().unwrap(), None);
        let removed = e.commit().unwrap();
        let absent = save(removed.package());
        assert_eq!(absent.read_only_recommended().unwrap(), None);
        assert!(members(&absent.to_bytes().unwrap()).contains_key(PROPS));
        assert_same_payloads_and_graph(
            &off.to_bytes().unwrap(),
            &absent
                .apply_read_only_recommended(&removed.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
        );
        // A recommendation must not turn into a library access-control policy.
        let mut tag_edit = after.edit_tags().unwrap();
        tag_edit.set(0_usize, tags("Allowed")).unwrap();
        let allowed = tag_edit.commit().unwrap();
        assert_eq!(
            allowed.package().read_only_recommended().unwrap(),
            Some(true)
        );
        assert_eq!(
            allowed
                .package()
                .tags(0_usize)
                .unwrap()
                .unwrap()
                .get("review")
                .unwrap()
                .value(),
            "Allowed"
        );
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_slide_tag_crud_preserves_shapes_resources_and_prior_staged_changes() {
    for name in FILES {
        let bytes = native(name);
        let p = Package::from_reader(Cursor::new(&bytes)).unwrap();
        assert!(p.tags(0_usize).unwrap().is_none());
        let mut e = p.edit_tags().unwrap();
        assert_eq!(e.remove(0_usize).unwrap(), None);
        assert_eq!(e.commit().unwrap().package().to_bytes().unwrap(), bytes);
        let mut e = p.edit_tags().unwrap();
        assert_eq!(e.set(0_usize, tags("First")).unwrap(), None);
        assert!(matches!(
            e.set(usize::MAX, tags("Invalid")),
            Err(Error::SlideIndexOutOfBounds { .. })
        ));
        assert_eq!(e.tags(0_usize).unwrap(), Some(tags("First")));
        let created = e.commit().unwrap();
        let after = save(created.package());
        assert_eq!(after.tags(0_usize).unwrap(), Some(tags("First")));
        let out = after.to_bytes().unwrap();
        let differences = changed(&bytes, &out);
        assert_eq!(
            differences,
            [
                "[Content_Types].xml",
                "ppt/slides/_rels/slide1.xml.rels",
                SLIDE,
                "ppt/tags/tag1.xml"
            ]
        );
        let parts = members(&bytes);
        let changed_parts = members(&out);
        // The new direct tag anchor belongs inside cSld. Its shape tree and
        // existing slide creation ID remain byte-exact.
        for child in [b"p:spTree".as_slice(), b"p14:creationId"] {
            assert_eq!(
                fragment(&parts[SLIDE], child),
                fragment(&changed_parts[SLIDE], child)
            );
        }
        assert_same_payloads_and_graph(
            &out,
            &p.apply_tags(created.patch())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
        );
        assert!(matches!(
            after.apply_tags(created.patch()),
            Err(Error::PatchConflict)
        ));
        assert_same_payloads_and_graph(
            &bytes,
            &after
                .apply_tags(&created.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
        );
        let mut e = after.edit_tags().unwrap();
        assert_eq!(e.set(0_usize, tags("Second")).unwrap(), Some(tags("First")));
        let updated = e.commit().unwrap();
        let second = save(updated.package());
        assert_eq!(second.tags(0_usize).unwrap(), Some(tags("Second")));
        assert_eq!(
            changed(&out, &second.to_bytes().unwrap()),
            ["ppt/tags/tag1.xml"]
        );
        let mut e = second.edit_tags().unwrap();
        assert_eq!(e.remove(0_usize).unwrap(), Some(tags("Second")));
        let removed = e.commit().unwrap();
        let absent = save(removed.package());
        assert!(absent.tags(0_usize).unwrap().is_none());
        assert_same_payloads_and_graph(&bytes, &absent.to_bytes().unwrap());
        assert_same_payloads_and_graph(
            &second.to_bytes().unwrap(),
            &absent
                .apply_tags(&removed.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
        );
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_shape_tags_keep_direct_and_sibling_owners_distinct() {
    for name in FILES {
        let bytes = native(name);
        let p = Package::from_reader(Cursor::new(&bytes)).unwrap();
        let mut direct = p.edit_tags().unwrap();
        direct.set(0_usize, tags("Slide")).unwrap();
        let direct = direct.commit().unwrap();
        let base = direct.package();
        let before = base.to_bytes().unwrap();
        let mut e = base.edit_shape_tags().unwrap();
        assert_eq!(e.set(0_usize, "Title 1", tags("Shape")).unwrap(), None);
        assert!(matches!(
            e.set(0_usize, "Missing native shape", tags("Invalid")),
            Err(Error::ShapeLookup(
                litchi_pptx::shape::LookupError::NameNotFound { .. }
            ))
        ));
        let created = e.commit().unwrap();
        let after = save(created.package());
        let output = after.to_bytes().unwrap();
        assert_eq!(after.tags(0_usize).unwrap(), Some(tags("Slide")));
        assert_eq!(
            after.shape_tags(0_usize, "Title 1").unwrap(),
            Some(tags("Shape"))
        );
        assert_eq!(after.shape_tags(0_usize, "Subtitle 2").unwrap(), None);
        assert_eq!(
            changed(&before, &output),
            [
                "[Content_Types].xml",
                "ppt/slides/_rels/slide1.xml.rels",
                SLIDE,
                "ppt/tags/tag2.xml"
            ]
        );
        let source_parts = members(&before);
        let edited_parts = members(&output);
        for child in [b"p:txBody".as_slice(), b"p14:creationId"] {
            assert_eq!(
                fragment(&source_parts[SLIDE], child),
                fragment(&edited_parts[SLIDE], child)
            );
        }
        assert_same_payloads_and_graph(
            &before,
            &after
                .apply_shape_tags(&created.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
        );
        let mut e = after.edit_shape_tags().unwrap();
        assert_eq!(
            e.set(0_usize, "Title 1", tags("Updated shape")).unwrap(),
            Some(tags("Shape"))
        );
        let updated = e.commit().unwrap();
        let updated = save(updated.package());
        assert_eq!(
            changed(&output, &updated.to_bytes().unwrap()),
            ["ppt/tags/tag2.xml"]
        );
        let mut e = updated.edit_shape_tags().unwrap();
        assert_eq!(
            e.remove(0_usize, "Title 1").unwrap(),
            Some(tags("Updated shape"))
        );
        let removed = e.commit().unwrap();
        let absent = save(removed.package());
        assert_eq!(absent.tags(0_usize).unwrap(), Some(tags("Slide")));
        assert_eq!(absent.shape_tags(0_usize, "Title 1").unwrap(), None);
        assert_same_payloads_and_graph(&before, &absent.to_bytes().unwrap());
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_laser_creation_preserves_slide_content_and_refuses_replacement_atomically() {
    for name in [FILES[0], FILES[3], "powerpoint-master-motion.pptx"] {
        let bytes = native(name);
        let p = Package::from_reader(Cursor::new(&bytes)).unwrap();
        assert!(traces(&p).is_empty());
        let pts = points();
        let mut e = p.edit_laser_traces().unwrap();
        e.set(0_usize, &pts).unwrap();
        assert!(matches!(
            e.set(0_usize, &pts),
            Err(Error::Invalid(message))
                if message == "slide already contains a laser-trace extension; replacement is not supported"
        ));
        let created = e.commit().unwrap();
        let after = save(created.package());
        let output = after.to_bytes().unwrap();
        let ts = traces(&after);
        assert_eq!(ts.len(), 1);
        assert_eq!(ts[0].slide_index(), 0);
        assert_eq!(ts[0].trace_index(), 0);
        assert_eq!(ts[0].points(), pts);
        assert_eq!(changed(&bytes, &output), [SLIDE]);
        let old = members(&bytes);
        let new = members(&output);
        for child in [b"p:cSld".as_slice(), b"p14:creationId"] {
            assert_eq!(fragment(&old[SLIDE], child), fragment(&new[SLIDE], child));
        }
        assert_same_payloads_and_graph(
            &output,
            &p.apply_laser_traces(created.patch())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
        );
        assert!(matches!(
            after.apply_laser_traces(created.patch()),
            Err(Error::PatchConflict)
        ));
        assert_same_payloads_and_graph(
            &bytes,
            &after
                .apply_laser_traces(&created.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
        );
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_event_crud_preserves_existing_laser_creation_ids_and_media_payloads() {
    const EVENT_SLIDE: &str = "ppt/slides/slide2.xml";
    for name in [
        "powerpoint-media-objects.pptx",
        "powerpoint-modern-media.pptx",
    ] {
        let bytes = native(name);
        let p = Package::from_reader(Cursor::new(&bytes)).unwrap();
        let media = litchi_pptx::media_parts::load(
            p.opc().unwrap(),
            &PackURI::new(format!("/{EVENT_SLIDE}")).unwrap(),
        )
        .unwrap();
        assert_eq!(media.pictures.len(), 1);
        assert_eq!(media.pictures[0].shape_id, 4);
        assert_eq!(
            media.pictures[0].kind,
            litchi_pptx::media_parts::Kind::Video
        );
        let mut laser = p.edit_laser_traces().unwrap();
        laser.set(1_usize, &points()).unwrap();
        let laser = laser.commit().unwrap();
        let base = laser.package();
        let before = base.to_bytes().unwrap();
        let values = event_values(4);
        let mut e = base.edit_show_events().unwrap();
        e.set(1_usize, &values).unwrap();
        let created = e.commit().unwrap();
        let after = save(created.package());
        let output = after.to_bytes().unwrap();
        let loaded = events::load(
            after.opc().unwrap(),
            &PackURI::new(format!("/{EVENT_SLIDE}")).unwrap(),
            events::Limits::default(),
        )
        .unwrap()
        .unwrap();
        assert_eq!(loaded.len(), 7);
        for (index, (actual, expected)) in loaded.iter().zip(&values).enumerate() {
            assert_eq!(actual.event_index(), index);
            assert_eq!(actual.kind(), expected.kind());
            assert_eq!(actual.time(), expected.time());
            assert_eq!(actual.object_id(), expected.object_id());
            assert_eq!(actual.seek_time(), expected.seek_time());
        }
        assert_eq!(
            traces_on(&after, 1, EVENT_SLIDE),
            traces_on(base, 1, EVENT_SLIDE)
        );
        assert_eq!(changed(&before, &output), [EVENT_SLIDE]);
        let old = members(&before);
        let new = members(&output);
        for child in [
            b"p:cSld".as_slice(),
            b"p14:creationId",
            b"p14:laserTraceLst",
        ] {
            assert_eq!(
                fragment(&old[EVENT_SLIDE], child),
                fragment(&new[EVENT_SLIDE], child)
            );
        }
        assert_same_payloads_and_graph(
            &before,
            &after
                .apply_show_events(&created.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
        );
        assert!(matches!(
            after.apply_show_events(created.patch()),
            Err(Error::PatchConflict)
        ));
        let mut e = after.edit_show_events().unwrap();
        e.update(1_usize, |d| {
            d.replace(vec![Draft::null(Offset::ms(800), 4)])
        })
        .unwrap();
        assert!(matches!(
            e.update(1_usize, |d| d.replace(vec![])),
            Err(Error::Invalid(message))
                if message == "slide-show event storage requires at least one event"
        ));
        assert_eq!(e.show_events(1_usize).unwrap().unwrap().len(), 1);
        let updated = e.commit().unwrap();
        let updated = save(updated.package());
        assert_eq!(
            updated
                .edit_show_events()
                .unwrap()
                .show_events(1_usize)
                .unwrap()
                .unwrap()[0]
                .time(),
            &Offset::ms(800)
        );
        let mut e = updated.edit_show_events().unwrap();
        assert!(e.remove(1_usize).unwrap().is_some());
        let removed = e.commit().unwrap();
        let absent = save(removed.package());
        assert_eq!(
            absent
                .edit_show_events()
                .unwrap()
                .show_events(1_usize)
                .unwrap(),
            None
        );
        assert_same_payloads_and_graph(&before, &absent.to_bytes().unwrap());
        assert_same_payloads_and_graph(
            &updated.to_bytes().unwrap(),
            &absent
                .apply_show_events(&removed.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
        );
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}
