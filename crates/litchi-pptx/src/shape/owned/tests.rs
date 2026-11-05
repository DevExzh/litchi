#![allow(clippy::unwrap_used, clippy::expect_used, reason = "test assertions")]
use super::*;
const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
fn tree(body: &str) -> Vec<u8> {
    format!(r#"<p:spTree xmlns:p="{PML}">{body}</p:spTree>"#).into_bytes()
}
fn shape(name: &str) -> String {
    format!(r#"<p:sp><p:nvSpPr><p:cNvPr id="7" name="{name}"/></p:nvSpPr></p:sp>"#)
}
#[test]
fn selection_preserves_absence_ambiguity_and_redaction() {
    let scene = Scene::read(
        tree(&format!("{}{}", shape("secret"), shape("secret"))),
        Limits::DEFAULT,
    )
    .unwrap();
    let error = scene.shape("secret").unwrap_err();
    assert_eq!(error, Error::Ambiguous { matches: 2 });
    assert!(
        !format!("{error} {error:?} {scene:?} {:?}", Selector::from("secret")).contains("secret")
    );
    assert!(scene.shape("SECRET").unwrap().is_none());
    assert!(scene.shape(usize::MAX).unwrap().is_none());
    assert!(scene.shape(2usize).unwrap().is_none());
    let handle = scene.shape(Position::new(1)).unwrap().unwrap();
    assert!(!format!("{handle:?} {:?}", scene.iter()).contains("secret"));
    assert_eq!(handle.position(), Position::new(1));
    assert_eq!(handle.view().name(), Some("secret"));
}
#[test]
fn handles_and_iterators_outlive_the_scene_and_preserve_group_axes() {
    let scene = Scene::read(tree(&format!(r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="1" name="group"/></p:nvGrpSpPr>{}<p:grpSp><p:nvGrpSpPr><p:cNvPr id="2" name="nested"/></p:nvGrpSpPr>{}</p:grpSp></p:grpSp>{}"#, shape("first"), shape("deep"), shape("last"))), Limits::DEFAULT).unwrap();
    let group = scene.shape("group").unwrap().unwrap();
    let roots = scene.roots();
    let mut all = scene.iter();
    drop(scene);
    assert_eq!(
        group
            .children()
            .map(|h| h.position().get())
            .collect::<Vec<_>>(),
        [1, 2]
    );
    assert_eq!(
        roots.map(|h| h.position().get()).collect::<Vec<_>>(),
        [0, 4]
    );
    assert_eq!(all.size_hint(), (5, Some(5)));
    let handles = all.by_ref().collect::<Vec<_>>();
    assert_eq!(handles[3].parent().unwrap().position().get(), 2);
    assert!(handles[0].parent().is_none());
    assert!(handles[1].children().next().is_none());
    assert_eq!(
        handles
            .iter()
            .map(|h| h.position().get())
            .collect::<Vec<_>>(),
        [0, 1, 2, 3, 4]
    );
    assert!(all.next().is_none());
    assert!(all.next().is_none());
    assert_eq!(all.size_hint(), (0, Some(0)));
}
#[test]
fn unchanged_input_is_shared_including_bom_and_spans() {
    let mut bytes = vec![0xef, 0xbb, 0xbf];
    bytes.extend(tree(&shape("Title")));
    let input = Arc::new(bytes);
    let scene = Scene::read(Arc::clone(&input), Limits::DEFAULT).unwrap();
    assert_eq!(raw::xml(&scene).as_ptr(), input.as_ptr());
    assert!(!scene.is_rewritten());
    let clone = scene.clone();
    assert!(Arc::ptr_eq(&scene.data, &clone.data));
    let handle = scene.shape("Title").unwrap().unwrap();
    drop(scene);
    drop(input);
    let view = handle.view();
    let span = view.span().unwrap();
    assert_eq!(
        view.xml().unwrap(),
        &raw::xml(&clone)[span.start() as usize..span.end().unwrap() as usize]
    );
}
#[test]
fn mce_owns_processed_xml_and_releases_original_allocation() {
    let input = Arc::new(format!(r#"<p:spTree xmlns:p="{PML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="x">{}</mc:Choice><mc:Fallback>{}</mc:Fallback></mc:AlternateContent></p:spTree>"#, shape("discard"), shape("keep")).into_bytes());
    let scene = Scene::read(Arc::clone(&input), Limits::DEFAULT).unwrap();
    assert!(scene.is_rewritten());
    assert_eq!(Arc::strong_count(&input), 1);
    assert_eq!(scene.len(), 1);
    assert!(scene.shape("discard").unwrap().is_none());
    let retained = scene.shape("keep").unwrap().unwrap();
    let borrowed = super::super::Scene::read(input.as_slice(), Limits::DEFAULT).unwrap();
    assert_eq!(raw::xml(&scene), borrowed.xml());
    assert_eq!(
        retained.view().xml().unwrap(),
        borrowed.at(0).unwrap().xml().unwrap()
    );
}
#[test]
fn parsing_and_limits_match_the_borrowed_reader() {
    let bytes = tree(&shape("Title"));
    let limits = Limits::new(1, 1, 1, 1, 1, 1).unwrap();
    assert!(Scene::read(bytes.clone(), limits).is_err());
    assert!(super::super::Scene::read(&bytes, limits).is_err());
    assert!(Scene::read(b"<p:spTree>".to_vec(), Limits::DEFAULT).is_err());
    let empty = Scene::read(tree(""), Limits::DEFAULT).unwrap();
    assert!(empty.is_empty());
    assert!(empty.iter().next().is_none());
    assert!(empty.roots().next().is_none());
}

#[test]
fn table_projection_keeps_inherited_namespace_aliases() {
    const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const TABLE_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/table";
    let xml = format!(
        r#"<q:spTree xmlns:q="{PML}" xmlns:d="{DML}">
            <q:graphicFrame>
                <q:nvGraphicFramePr><q:cNvPr id="5" name="Grid"/></q:nvGraphicFramePr>
                <d:graphic><d:graphicData uri="{TABLE_URI}"><d:tbl>
                    <d:tr><d:tc><d:txBody><d:p><d:r><d:t>Cell</d:t></d:r></d:p></d:txBody></d:tc></d:tr>
                </d:tbl></d:graphicData></d:graphic>
            </q:graphicFrame>
        </q:spTree>"#
    );
    let scene = Scene::read(xml.into_bytes(), Limits::DEFAULT).unwrap();
    let handle = scene.shape("Grid").unwrap().unwrap();
    let table = handle.table().unwrap();
    assert_eq!(table.row_count().unwrap(), 1);
    assert_eq!(table.cell(0, 0).unwrap().unwrap().text().unwrap(), "Cell");
}

#[test]
fn table_projection_does_not_treat_unbound_or_rebound_prefixes_as_drawingml() {
    const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const TABLE_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/table";
    const FOREIGN: &str = "urn:foreign";
    let cases = [
        format!(
            r#"<q:spTree xmlns:q="{PML}" xmlns:d="{DML}">
                <q:graphicFrame><q:nvGraphicFramePr><q:cNvPr id="5" name="Grid"/></q:nvGraphicFramePr>
                    <a:graphic><a:graphicData uri="{TABLE_URI}"><a:tbl><a:tr><a:tc/></a:tr></a:tbl></a:graphicData></a:graphic>
                </q:graphicFrame>
            </q:spTree>"#
        ),
        format!(
            r#"<q:spTree xmlns:q="{PML}" xmlns:a="{DML}">
                <q:graphicFrame><q:nvGraphicFramePr><q:cNvPr id="5" name="Grid"/></q:nvGraphicFramePr>
                    <a:graphic xmlns:a="{FOREIGN}"><a:graphicData uri="{TABLE_URI}"><a:tbl><a:tr><a:tc/></a:tr></a:tbl></a:graphicData></a:graphic>
                </q:graphicFrame>
            </q:spTree>"#
        ),
    ];
    for xml in cases {
        let scene = Scene::read(xml.into_bytes(), Limits::DEFAULT).unwrap();
        let handle = scene.shape("Grid").unwrap().unwrap();
        assert!(handle.table().is_err());
    }
}

#[test]
fn retained_types_are_small_send_and_sync() {
    fn assert_traits<T: Send + Sync + Clone + 'static>() {}
    assert_traits::<Scene>();
    assert_traits::<Handle>();
    assert_traits::<Handles>();
    assert_eq!(size_of::<Scene>(), size_of::<usize>());
    assert_eq!(size_of::<Handle>(), 2 * size_of::<usize>());
}
