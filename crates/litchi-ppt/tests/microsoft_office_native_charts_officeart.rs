//! Bounded Graph publication and opaque OfficeArt-host metadata on native PowerPoint parents.
use litchi_core::{Position, edit::RefusalKind};
use litchi_ppt::builder::{Chart as Request, ChartKind};
use litchi_ppt::client_data::{ClientDataChild as Child, ClientDataChildKind as ChildKind};
use litchi_ppt::embedded::{object, storage::Storage};
use litchi_ppt::{Package, Record, RecordParser, RecordType, SourceLimits, text_edit::Target};
use std::{collections::BTreeMap, io::Cursor, path::Path};
#[path = "../../../tools/test_support/cfb_directory_metadata.rs"]
mod cfb_metadata;
const VISUAL: [&str; 3] = [
    "powerpoint-visual-objects.ppt",
    "powerpoint-visual-objects.pps",
    "powerpoint-visual-objects.pot",
];
const MACRO: &str = "powerpoint-passive-macro.ppt";
const OBJECT: u32 = 100;
const OPAQUE: u16 = 0x7ffe;
fn native(file: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/powerpoint")
            .join(file),
    )
    .unwrap()
}
fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    ole.list_streams()
        .into_iter()
        .map(|p| {
            let b = ole
                .open_stream(&p.iter().map(String::as_str).collect::<Vec<_>>())
                .unwrap();
            (p, b)
        })
        .collect()
}
fn wire(r: &Record) -> Vec<u8> {
    let mut b = ((r.instance << 4) | r.version).to_le_bytes().to_vec();
    b.extend(r.record_type_raw.to_le_bytes());
    b.extend(r.data_length.to_le_bytes());
    b.extend_from_slice(&r.data);
    b
}
fn persisted(bytes: &[u8]) -> BTreeMap<u32, Vec<u8>> {
    let streams = streams(bytes);
    let data = &streams[&vec!["PowerPoint Document".into()]];
    let limits = SourceLimits::default();
    let mut parser = RecordParser::new();
    parser.parse_document(data, &limits).unwrap();
    let map = litchi_ppt::PersistMapping::build_from_records_ref(&parser.find_records_ref());
    map.get_persist_ids()
        .into_iter()
        .map(|id| {
            let r = Record::parse(data, map.get_offset(id).unwrap() as usize, &limits)
                .unwrap()
                .0;
            (id, wire(&r))
        })
        .collect()
}
fn preserve(p: &Package, q: &Package, owned: &[u32]) {
    let before = p.to_bytes().unwrap();
    let after = q.to_bytes().unwrap();
    let a = streams(&before);
    let b = streams(&after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, data) in &a {
        if name != &["PowerPoint Document".to_string()] && name != &["Current User".to_string()] {
            assert!(b[name] == *data, "changed {name:?}");
        }
    }
    let doc = vec!["PowerPoint Document".into()];
    assert!(b[&doc].starts_with(&a[&doc]));
    let user = vec!["Current User".into()];
    assert_eq!(&a[&user][..16], &b[&user][..16]);
    assert_eq!(&a[&user][20..], &b[&user][20..]);
    assert_eq!(
        cfb_metadata::directory(&before),
        cfb_metadata::directory(&after)
    );
    let a = persisted(&before);
    let b = persisted(&after);
    for (id, bytes) in a {
        if !owned.contains(&id) {
            assert!(b.get(&id) == Some(&bytes), "changed persisted object {id}");
        }
    }
    let a = p.presentation().unwrap();
    let b = q.presentation().unwrap();
    assert_eq!(a.slide_count(), b.slide_count());
    for (a, b) in a.slides().unwrap().iter().zip(b.slides().unwrap()) {
        assert_eq!(a.text().unwrap(), b.text().unwrap());
    }
}
struct TempPpt(std::path::PathBuf);
impl TempPpt {
    fn new() -> Self {
        static NEXT: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);
        let count = NEXT.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
        let now = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        Self(std::env::temp_dir().join(format!(
            "litchi-native-chart-{}-{now}-{count}.ppt",
            std::process::id()
        )))
    }
}
impl Drop for TempPpt {
    fn drop(&mut self) {
        let _ = std::fs::remove_file(&self.0);
    }
}
fn save(p: &Package) -> Package {
    let f = TempPpt::new();
    p.save(&f.0).unwrap();
    let q = Package::open(&f.0).unwrap();
    assert!(q.to_bytes().unwrap() == p.to_bytes().unwrap());
    q
}
fn publish_object(p: &Package, c: &object::Commit, owned: &[u32]) -> Package {
    let q = save(c.package());
    preserve(p, &q, owned);
    assert!(
        p.apply_objects(c.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == q.to_bytes().unwrap()
    );
    assert!(
        q.apply_objects(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == p.to_bytes().unwrap()
    );
    q
}
fn publish_client(p: &Package, c: &litchi_ppt::client_data::Commit, slide_id: u32) -> Package {
    let q = save(c.package());
    preserve(p, &q, &[slide_id]);
    assert!(
        p.apply_client_data(c.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == q.to_bytes().unwrap()
    );
    assert!(
        q.apply_client_data(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == p.to_bytes().unwrap()
    );
    q
}
struct Fresh {
    definition: object::ExternalObject,
    storage: Storage,
    chart: Vec<u8>,
    title: String,
}
fn fresh(kind: ChartKind, offset: f64) -> Fresh {
    let mut builder = litchi_ppt::writer::Builder::new();
    let slide = builder.add_slide().unwrap();
    let mut request = Request::new(kind);
    let title = format!("{kind:?} stored 🍎");
    request.set_title(title.clone());
    request.set_categories(["Q1", "Q2", "Q3"]);
    request
        .add_series(
            Some("Stored series"),
            vec![10.0 + offset, 20.0 + offset, 30.0 + offset],
        )
        .unwrap();
    builder.add_chart(slide, 40, 50, 400, 240, request).unwrap();
    let mut bytes = Cursor::new(Vec::new());
    litchi_ppt::raw::encode(&mut builder, &mut bytes).unwrap();
    let p = Package::from_bytes(&bytes.into_inner()).unwrap();
    let d = p.presentation().unwrap();
    let inventory = d.charts().unwrap();
    assert_eq!(
        (
            inventory.seen(),
            inventory.len(),
            inventory.failures().count()
        ),
        (1, 1, 0)
    );
    let graph = inventory.get(0).unwrap().as_graph().unwrap();
    assert!(graph.info().frame().is_some());
    let id = graph.info().object_id();
    let mut definition = d.ole_objects().unwrap().unwrap().get(id).unwrap().clone();
    let object::ExternalObject::Object(ref mut object) = definition else {
        panic!("Graph object")
    };
    object.object.id = OBJECT;
    let storage = d.ole_storage(graph.info().persist_id()).unwrap().unwrap();
    let chart = graph
        .book()
        .charts()
        .unwrap()
        .next()
        .unwrap()
        .unwrap()
        .as_bytes()
        .to_vec();
    let model = graph.semantic_chart().unwrap();
    assert_eq!(model.title(), Some(title.as_str()));
    assert_eq!(model.series().len(), 1);
    assert_eq!(model.groups().len(), 1);
    assert!(matches!(
        (kind, &model.groups()[0].family),
        (ChartKind::Bar, litchi_ograph::chart::Family::Bar { .. })
            | (ChartKind::Line, litchi_ograph::chart::Family::Line { .. })
            | (ChartKind::Pie, litchi_ograph::chart::Family::Pie { .. })
    ));
    assert_eq!(model.caches().len(), 7);
    let numbers = model
        .caches()
        .iter()
        .filter_map(|v| match v.value() {
            litchi_ograph::chart::ValueRef::Number(n) => Some(n),
            _ => None,
        })
        .collect::<Vec<_>>();
    assert_eq!(numbers, [10.0 + offset, 20.0 + offset, 30.0 + offset]);
    Fresh {
        definition,
        storage,
        chart,
        title,
    }
}
fn add(p: &Package, fresh: &Fresh) -> (Package, u32) {
    let doc = p
        .presentation()
        .unwrap()
        .slide_directory()
        .document_persist_id();
    let mut edit = p.edit_objects().unwrap();
    let persist = edit
        .add(fresh.definition.clone(), fresh.storage.clone())
        .unwrap();
    let c = edit.commit().unwrap();
    let q = publish_object(p, &c, &[doc]);
    let d = q.presentation().unwrap();
    let inventory = d.charts().unwrap();
    let chart = inventory
        .charts()
        .find(|c| c.info().object_id() == OBJECT)
        .unwrap();
    assert_eq!(chart.kind(), litchi_ppt::chart::Kind::Graph);
    assert_eq!(chart.info().persist_id(), persist);
    assert!(chart.info().frame().is_none());
    assert_eq!(
        chart.semantic_chart(0).unwrap().unwrap().title(),
        Some(fresh.title.as_str())
    );
    (q, persist)
}
fn frame_target(p: &Package) -> (Target, u32) {
    let d = p.presentation().unwrap();
    let slide = d.slide_at(3).unwrap().unwrap();
    let shapes = slide.shapes().unwrap();
    let index = shapes.iter().position(|s| s.id() == 5122).unwrap();
    (
        Target::new(Position::new(3), Position::new(index)),
        slide.persist_id(),
    )
}
fn shape_records(p: &Package, skip: u16) -> Vec<(u16, u8, u16, Vec<u8>)> {
    fn find<'a, 'b>(
        shapes: &'a [litchi_ppt::odraw::Shape<'b>],
    ) -> Option<&'a litchi_ppt::odraw::Shape<'b>> {
        for s in shapes {
            if s.id() == 5122 {
                return Some(s);
            }
            if let Some(found) = find(s.children()) {
                return Some(found);
            }
        }
        None
    }
    let d = p.presentation().unwrap();
    let slide = d.slide_at(3).unwrap().unwrap();
    let drawing = slide
        .record()
        .children
        .iter()
        .find(|r| r.record_type == RecordType::PPDrawing)
        .unwrap();
    let parsed =
        litchi_ppt::odraw::parse_drawing(&drawing.data, &litchi_odraw::Limits::default()).unwrap();
    let shape = find(parsed.shapes()).unwrap();
    shape
        .meta()
        .children()
        .map(Result::unwrap)
        .filter(|r| r.raw_kind() != skip)
        .map(|r| (r.raw_kind(), r.version(), r.instance(), r.data().to_vec()))
        .collect()
}
fn opaque() -> Child {
    Child::opaque(6, 0x123, OPAQUE, b"opaque vendor bytes\0\x0f\xff".to_vec()).unwrap()
}
fn bind(p: &Package) -> Package {
    let (target, slide) = frame_target(p);
    let before = shape_records(p, 0xf011);
    let mut edit = p.edit_client_data(target).unwrap();
    let i = edit
        .children()
        .unwrap()
        .iter()
        .position(|c| c.kind() == ChildKind::ExternalObjectReference)
        .unwrap();
    let old = edit
        .replace(
            i,
            Child::new(
                ChildKind::ExternalObjectReference,
                OBJECT.to_le_bytes().to_vec(),
            )
            .unwrap(),
        )
        .unwrap();
    assert_eq!(old.external_object_id(), Some(1));
    edit.append(opaque()).unwrap();
    let c = edit.commit().unwrap();
    let q = publish_client(p, &c, slide);
    assert_eq!(shape_records(&q, 0xf011), before);
    let d = q.presentation().unwrap();
    let charts = d.charts().unwrap();
    let graph = charts
        .charts()
        .find(|c| c.info().object_id() == OBJECT)
        .unwrap();
    assert_eq!(
        graph.info().frame().map(|f| (f.slide(), f.shape())),
        Some((4, 5122))
    );
    assert!(
        charts
            .charts()
            .find(|c| c.info().object_id() == 1)
            .unwrap()
            .info()
            .frame()
            .is_none()
    );
    q
}

#[test]
fn native_ppt_excel_charts_replay_exactly_and_refuse_unproven_host_rewrites() {
    for file in VISUAL {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        let d = p.presentation().unwrap();
        let charts = d.charts().unwrap();
        assert_eq!(
            (charts.seen(), charts.len(), charts.failures().count()),
            (1, 1, 0)
        );
        let chart = charts.get(0).unwrap();
        assert_eq!(chart.info().program(), Some("Excel.Chart.8"));
        assert_eq!(
            chart.info().frame().map(|f| (f.slide(), f.shape())),
            Some((4, 5122))
        );
        assert_eq!(chart.kind(), litchi_ppt::chart::Kind::Excel);
        let model = chart.semantic_chart(0).unwrap().unwrap();
        assert_eq!(model.caches().len(), 24);
        assert_eq!(model.series().len(), 3);
        let raw = chart
            .charts()
            .unwrap()
            .next()
            .unwrap()
            .unwrap()
            .as_bytes()
            .to_vec();
        assert_eq!(model.encode().unwrap().as_bytes(), raw);
        assert!(
            matches!(chart.draft_package(),Err(litchi_ppt::package::Error::InvalidFormat(m)) if m=="only standalone Graph chart packages support a package draft")
        );
        let mut model = chart.semantic_chart(0).unwrap().unwrap();
        model.set_title(Some("Unproven parsed title edit".into()));
        assert!(
            matches!(model.encode(),Err(litchi_ograph::Error::Refused(r)) if r.kind()==RefusalKind::UnsupportedContent)
        );
        assert!(p.to_bytes().unwrap() == bytes);
    }
}
#[test]
fn native_ppt_graph_profile_storage_creation_replacement_and_removal_preserve_native_owners() {
    for file in VISUAL.into_iter().chain([MACRO]) {
        for kind in [ChartKind::Bar, ChartKind::Line, ChartKind::Pie] {
            let bytes = native(file);
            let p = Package::from_bytes(&bytes).unwrap();
            let authored = fresh(kind, 0.0);
            let (q, persist) = add(&p, &authored);
            let d = q.presentation().unwrap();
            let charts = d.charts().unwrap();
            let graph = charts
                .charts()
                .find(|c| c.info().object_id() == OBJECT)
                .unwrap()
                .as_graph()
                .unwrap();
            let draft = graph.draft_package().unwrap();
            let original_streams = streams(graph.package().as_bytes());
            let book = graph.package().workbook().unwrap();
            let prelude = book.as_bytes()[..book.chart().offset()].to_vec();
            assert!(!draft.is_dirty());
            assert_eq!(
                draft.into_bytes().unwrap(),
                authored
                    .storage
                    .decompressed_bytes(64 * 1024 * 1024)
                    .unwrap()
            );
            let mut edit = q.edit_objects().unwrap();
            graph
                .replace_package(&mut edit, graph.draft_package().unwrap())
                .unwrap();
            let noop = edit.commit().unwrap();
            assert!(noop.patch().is_noop());
            assert!(noop.package().to_bytes().unwrap() == q.to_bytes().unwrap());
            let replacement = fresh(ChartKind::Line, 7.0);
            let mut draft = graph.draft_package().unwrap();
            draft
                .replace_chart(
                    litchi_ograph::chart::Stream::open(
                        replacement.chart.clone(),
                        litchi_ograph::Limits::default(),
                    )
                    .unwrap(),
                )
                .unwrap();
            let mut edit = q.edit_objects().unwrap();
            graph.replace_package(&mut edit, draft).unwrap();
            let c = edit.commit().unwrap();
            let changed = publish_object(&q, &c, &[persist]);
            let d = changed.presentation().unwrap();
            let charts = d.charts().unwrap();
            let graph = charts
                .charts()
                .find(|c| c.info().object_id() == OBJECT)
                .unwrap();
            let model = graph.semantic_chart(0).unwrap().unwrap();
            assert_eq!(model.title(), Some(replacement.title.as_str()));
            assert!(matches!(
                model.groups()[0].family,
                litchi_ograph::chart::Family::Line { .. }
            ));
            assert_eq!(
                model
                    .caches()
                    .iter()
                    .filter_map(|value| match value.value() {
                        litchi_ograph::chart::ValueRef::Number(number) => Some(number),
                        _ => None,
                    })
                    .collect::<Vec<_>>(),
                [17.0, 27.0, 37.0]
            );
            let graph = graph.as_graph().unwrap();
            let changed_streams = streams(graph.package().as_bytes());
            assert_eq!(
                original_streams.keys().collect::<Vec<_>>(),
                changed_streams.keys().collect::<Vec<_>>()
            );
            for (name, data) in original_streams {
                if name != ["Workbook"] {
                    assert!(changed_streams[&name] == data);
                }
            }
            assert!(
                graph
                    .package()
                    .workbook()
                    .unwrap()
                    .as_bytes()
                    .starts_with(&prelude)
            );
            let doc = d.slide_directory().document_persist_id();
            let mut edit = changed.edit_objects().unwrap();
            assert_eq!(edit.remove(OBJECT).unwrap().id(), OBJECT);
            let c = edit.commit().unwrap();
            let removed = publish_object(&changed, &c, &[doc, persist]);
            assert!(
                removed
                    .presentation()
                    .unwrap()
                    .charts()
                    .unwrap()
                    .charts()
                    .all(|c| c.info().object_id() != OBJECT)
            );
            assert_eq!(native(file), bytes);
        }
    }
}
#[test]
fn native_ppt_chart_frame_retarget_and_anchor_edits_keep_opaque_client_data() {
    for file in VISUAL {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        let (p, _) = add(&p, &fresh(ChartKind::Bar, 0.0));
        let q = bind(&p);
        let (target, slide) = frame_target(&q);
        let fields = shape_records(&q, 0xf010);
        let mut edit = q.edit().unwrap();
        let expected = litchi_ppt::Anchor::full(120, 240, 3120, 2240).unwrap();
        edit.set_shape_anchor(target, expected).unwrap();
        let c = edit.commit().unwrap();
        let moved = save(c.package());
        preserve(&q, &moved, &[slide]);
        assert_eq!(shape_records(&moved, 0xf010), fields);
        let records = shape_records(&moved, 0);
        let anchor = records.iter().find(|r| r.0 == 0xf010).unwrap();
        assert_eq!(
            litchi_ppt::client_anchor::Data::parse(
                &anchor.3,
                &litchi_ppt::client_anchor::Limits::default()
            )
            .unwrap(),
            expected.data()
        );
        assert!(
            q.apply(c.patch()).unwrap().package().to_bytes().unwrap() == moved.to_bytes().unwrap()
        );
        assert!(
            moved
                .apply(&c.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
                == q.to_bytes().unwrap()
        );
        let edit = moved.edit_client_data(target).unwrap();
        let unknown = edit
            .client_data()
            .unwrap()
            .unknown_children()
            .collect::<Vec<_>>();
        assert_eq!(unknown.len(), 1);
        assert_eq!(unknown[0], &opaque());
        assert_eq!(unknown[0].to_bytes().unwrap(), opaque().to_bytes().unwrap());
        let wire = unknown[0].to_bytes().unwrap();
        let (record, used) =
            litchi_odraw::Record::parse(&wire, 0, litchi_odraw::Limits::default()).unwrap();
        assert_eq!(used, wire.len());
        assert_eq!(record.kind(), litchi_odraw::RecordKind::Unknown(OPAQUE));
        assert_eq!(
            (record.version(), record.instance(), record.data()),
            (6, 0x123, opaque().payload())
        );
        assert_eq!(native(file), bytes);
    }
}

#[test]
fn native_ppt_graph_replacement_checks_host_identity_before_touching_staged_objects() {
    for file in VISUAL {
        let p = Package::from_bytes(&native(file)).unwrap();
        let (p, _) = add(&p, &fresh(ChartKind::Bar, 0.0));
        let original = p.to_bytes().unwrap();
        let d = p.presentation().unwrap();
        let charts = d.charts().unwrap();
        let graph = charts
            .charts()
            .find(|c| c.info().object_id() == OBJECT)
            .unwrap()
            .as_graph()
            .unwrap();
        let mut edit = p.edit_objects().unwrap();
        edit.remove(OBJECT).unwrap();
        let staged = edit.objects().unwrap().clone();
        assert!(
            matches!(graph.replace_package(&mut edit,graph.draft_package().unwrap()),Err(object::Error::Refused(r)) if r.kind()==RefusalKind::NotFound)
        );
        assert_eq!(edit.objects().unwrap(), &staged);
        let native_chart = charts.charts().find(|c| c.info().object_id() == 1).unwrap();
        let mut excel = d.ole_objects().unwrap().unwrap().get(1).unwrap().clone();
        let object::ExternalObject::Object(ref mut definition) = excel else {
            panic!("Excel object")
        };
        definition.object.id = OBJECT;
        let storage = d
            .ole_storage(native_chart.info().persist_id())
            .unwrap()
            .unwrap();
        edit.add(excel, storage).unwrap();
        let staged = edit.objects().unwrap().clone();
        assert!(
            matches!(graph.replace_package(&mut edit,graph.draft_package().unwrap()),Err(object::Error::Refused(r)) if r.kind()==RefusalKind::UnsupportedTarget)
        );
        assert_eq!(edit.objects().unwrap(), &staged);
        assert!(p.to_bytes().unwrap() == original);
    }
}

#[test]
fn native_ppt_opaque_client_data_replacement_removal_and_failed_staging_are_exact() {
    for file in VISUAL {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        let (p, _) = add(&p, &fresh(ChartKind::Pie, 0.0));
        let p = bind(&p);
        let (target, slide) = frame_target(&p);
        let fields = shape_records(&p, 0xf011);
        let mut edit = p.edit_client_data(target).unwrap();
        let index = edit
            .children()
            .unwrap()
            .iter()
            .position(|c| c.is_unknown())
            .unwrap();
        let replacement = Child::opaque(
            6,
            0x123,
            OPAQUE,
            b"replacement opaque bytes with a longer payload".to_vec(),
        )
        .unwrap();
        assert_eq!(edit.replace(index, replacement.clone()).unwrap(), opaque());
        let c = edit.commit().unwrap();
        let q = publish_client(&p, &c, slide);
        assert_eq!(shape_records(&q, 0xf011), fields);
        let mut edit = q.edit_client_data(target).unwrap();
        assert_eq!(edit.remove(index).unwrap(), replacement);
        let staged = edit.children().unwrap().to_vec();
        assert!(
            matches!(edit.append(Child::new(ChildKind::ExternalObjectReference,OBJECT.to_le_bytes().to_vec()).unwrap()),Err(litchi_ppt::client_data::Error::Refused(r)) if r.kind()==RefusalKind::InvalidValue)
        );
        assert_eq!(edit.children().unwrap(), staged);
        let c = edit.commit().unwrap();
        let removed = publish_client(&q, &c, slide);
        assert!(
            removed
                .edit_client_data(target)
                .unwrap()
                .client_data()
                .unwrap()
                .unknown_children()
                .next()
                .is_none()
        );
        assert_eq!(shape_records(&removed, 0xf011), fields);
        assert!(matches!(
            removed.apply_client_data(c.patch()),
            Err(litchi_ppt::client_data::Error::PatchConflict)
        ));
        assert!(
            matches!(Child::opaque(0,0,0x0bc1,OBJECT.to_le_bytes().to_vec()),Err(litchi_ppt::package::Error::Corrupted(m)) if m=="opaque client-data record type 0x0BC1 is a defined child")
        );
        assert_eq!(native(file), bytes);
    }
}
