//! Explicit classic-comment derivatives of native modern-comment and macro parents.
use litchi_pptx::{
    Error, Package, SourceBackedPresentation,
    comments::{self, Author, Comment, Comments, Conformance, List},
};
use std::{collections::BTreeMap, io::Cursor, path::Path};
const AUTHORS: &str = "/ppt/commentAuthors.xml";
const COMMENTS: &str = "/ppt/comments/nativeClassic069.xml";
const CONF: Conformance = Conformance::Transitional;
fn native(file: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/powerpoint")
            .join(file),
    )
    .unwrap()
}
fn members(p: &Package) -> BTreeMap<String, Vec<u8>> {
    let r = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        p.to_bytes().unwrap(),
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
fn first() -> Comment {
    Comment::new(7, "Classic comment 🍎 & review", 100, 240).with_date_time("2026-11-07T12:34:56Z")
}
fn graph(part: &str) -> Comments {
    Comments {
        author_relationship_id: "rIdNativeClassicAuthors069".into(),
        author_part_name: AUTHORS.into(),
        authors: vec![Author {
            id: 7,
            name: "Ada & Co".into(),
            initials: "AC".into(),
            last_index: 1,
            color_index: 2,
        }],
        slides: vec![List {
            slide_part_name: part.into(),
            relationship_id: "rIdNativeClassic069".into(),
            part_name: COMMENTS.into(),
            comments: vec![first()],
        }],
    }
}
fn preserved(base: &Package, after: &Package, owner_rels: &str) {
    let a = members(base);
    let b = members(after);
    for (n, v) in a {
        if n != "[Content_Types].xml" && n != "ppt/_rels/presentation.xml.rels" && n != owner_rels {
            assert!(b.get(&n) == Some(&v), "changed {n}");
        }
    }
    let rels = |r: &litchi_opc::Relationships| {
        let mut v = r
            .iter()
            .filter(|r| {
                !r.reltype().ends_with("/comments") && !r.reltype().ends_with("/commentAuthors")
            })
            .map(|r| {
                (
                    r.r_id().to_owned(),
                    r.reltype().to_owned(),
                    r.target_ref().to_owned(),
                    r.is_external(),
                )
            })
            .collect::<Vec<_>>();
        v.sort();
        v
    };
    let a = base.opc().unwrap();
    let b = after.opc().unwrap();
    assert_eq!(rels(a.rels()), rels(b.rels()));
    for p in a.iter_parts() {
        let q = b.get_part(p.partname()).unwrap();
        assert_eq!(p.content_type(), q.content_type());
        assert_eq!(rels(p.rels()), rels(q.rels()));
    }
}
fn publish(p: &Package, c: &comments::Commit) -> Package {
    let f = tempfile::NamedTempFile::with_suffix(".pptx").unwrap();
    c.package().save(f.path()).unwrap();
    let q = Package::open(f.path()).unwrap();
    assert!(members(&q) == members(c.package()));
    assert!(members(p.apply_comments(c.patch()).unwrap().package()) == members(&q));
    assert!(members(q.apply_comments(&c.patch().inverse()).unwrap().package()) == members(p));
    q
}
fn read(p: &Package, slide: usize, expected: &Comment) {
    let view = p.presentation().unwrap();
    let s = view.slide(slide).unwrap().unwrap();
    let eager = s.comment_snapshot().unwrap().unwrap();
    let entry = eager
        .entries()
        .iter()
        .find(|e| {
            litchi_pptx::presentation::related::raw::comment_author_id(e) == expected.author_id
                && litchi_pptx::presentation::related::raw::comment_index(e) == expected.index
        })
        .unwrap();
    assert_eq!(entry.text(), expected.text);
    assert_eq!(entry.date_time(), expected.date_time.as_deref());
    assert_eq!(entry.location(), (expected.x, expected.y));
    let source = SourceBackedPresentation::from_reader(Cursor::new(p.to_bytes().unwrap())).unwrap();
    let data = source.slide(slide).unwrap().comments().unwrap().unwrap();
    let other = data
        .entries()
        .iter()
        .find(|e| {
            litchi_pptx::presentation::related::raw::comment_author_id(e) == expected.author_id
                && litchi_pptx::presentation::related::raw::comment_index(e) == expected.index
        })
        .unwrap();
    assert_eq!(other.text(), entry.text());
    assert_eq!(other.author_name(), entry.author_name());
    assert_eq!(other.initials(), entry.initials());
    assert_eq!(other.date_time(), entry.date_time());
    assert_eq!(other.location(), entry.location());
}
#[test]
fn native_pptx_classic_comment_lifecycle_preserves_modern_threads_and_macro_payloads() {
    for (file, slide, part, rels) in [
        (
            "powerpoint-navigation-review.pptx",
            3,
            "/ppt/slides/slide4.xml",
            "ppt/slides/_rels/slide4.xml.rels",
        ),
        (
            "powerpoint-passive-macro.pptm",
            0,
            "/ppt/slides/slide1.xml",
            "ppt/slides/_rels/slide1.xml.rels",
        ),
    ] {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        let source_bytes = p.to_bytes().unwrap();
        assert!(p.edit_comments().unwrap().comments().unwrap().is_none());
        let mut e = p.edit_comments().unwrap();
        e.set(&graph(part), CONF).unwrap();
        let c = e.commit().unwrap();
        let added = publish(&p, &c);
        assert_eq!(
            added.edit_comments().unwrap().comments().unwrap(),
            Some(graph(part))
        );
        read(&added, slide, &first());
        preserved(&p, &added, rels);
        assert!(matches!(
            added.apply_comments(c.patch()),
            Err(Error::PatchConflict)
        ));
        let mut e = added.edit_comments().unwrap();
        e.update(slide, 7, 1, first(), CONF).unwrap();
        assert!(e.commit().unwrap().patch().is_noop());
        let mut e = added.edit_comments().unwrap();
        e.add_author(Author::new(9, "Béatrice", "B"), CONF).unwrap();
        let second = Comment::new(9, "Second 日本語", -50, 300);
        e.add(slide, second.clone(), CONF).unwrap();
        let mut changed = first();
        changed.text = "Updated classic comment".into();
        e.update(slide, 7, 1, changed.clone(), CONF).unwrap();
        e.reorder(slide, &[(9, 1), (7, 1)], CONF).unwrap();
        e.reorder_authors(&[9, 7], CONF).unwrap();
        let before_error = e.comments().unwrap();
        assert!(
            matches!(e.remove_author(7,CONF),Err(Error::Invalid(m)) if m=="cannot remove an author referenced by a comment")
        );
        assert_eq!(e.comments().unwrap(), before_error);
        let c = e.commit().unwrap();
        let updated = publish(&added, &c);
        preserved(&p, &updated, rels);
        read(&updated, slide, &changed);
        read(&updated, slide, &second);
        let stored = updated
            .edit_comments()
            .unwrap()
            .comments()
            .unwrap()
            .unwrap();
        assert_eq!(stored.author_part_name, AUTHORS);
        assert_eq!(stored.author_relationship_id, "rIdNativeClassicAuthors069");
        assert_eq!(stored.slides[0].part_name, COMMENTS);
        assert_eq!(stored.slides[0].relationship_id, "rIdNativeClassic069");
        assert_eq!(
            stored.authors.iter().map(|a| a.id).collect::<Vec<_>>(),
            [9, 7]
        );
        assert_eq!(stored.slides[0].comments, [second, changed]);
        let mut second_author = Author::new(9, "Béatrice", "B");
        second_author.last_index = 1;
        assert_eq!(
            stored.authors,
            [second_author, graph(part).authors[0].clone()]
        );
        let a = members(&added);
        let b = members(&updated);
        assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
        assert_eq!(
            a.keys()
                .filter(|n| a.get(*n) != b.get(*n))
                .map(String::as_str)
                .collect::<Vec<_>>(),
            [
                AUTHORS.trim_start_matches('/'),
                COMMENTS.trim_start_matches('/')
            ]
        );
        let mut e = updated.edit_comments().unwrap();
        assert!(e.remove(slide, 7, 1, CONF).unwrap());
        assert!(e.remove_author(7, CONF).unwrap());
        assert!(e.remove(slide, 9, 1, CONF).unwrap());
        assert!(e.remove_author(9, CONF).unwrap());
        let c = e.commit().unwrap();
        let cleared = publish(&updated, &c);
        preserved(&p, &cleared, rels);
        let remaining = cleared
            .edit_comments()
            .unwrap()
            .comments()
            .unwrap()
            .unwrap();
        assert!(remaining.authors.is_empty() && remaining.slides.is_empty());
        assert!(!members(&cleared).contains_key(COMMENTS.trim_start_matches('/')));
        assert!(members(&cleared).contains_key(AUTHORS.trim_start_matches('/')));
        assert!(p.to_bytes().unwrap() == source_bytes);
        assert!(native(file) == bytes);
    }
}
#[test]
fn native_pptx_classic_invalid_date_and_index_preserve_earlier_staging() {
    let bytes = native("powerpoint-navigation-review.pptx");
    let p = Package::from_bytes(&bytes).unwrap();
    let mut e = p.edit_comments().unwrap();
    e.set(&graph("/ppt/slides/slide4.xml"), CONF).unwrap();
    let before = e.comments().unwrap();
    assert!(
        matches!(e.set(&graph("/ppt/slides/slide4.xml"),CONF),Err(Error::Invalid(m)) if m=="package already contains presentation comments")
    );
    assert_eq!(e.comments().unwrap(), before);
    for (value, reason) in [
        (
            first().with_index(0),
            "replacement comment identity must remain stable",
        ),
        (
            first().with_date_time("not-a-date"),
            "invalid XML dateTime 'not-a-date'",
        ),
    ] {
        assert!(matches!(e.update(3usize,7,1,value,CONF),Err(Error::Invalid(m)) if m==reason));
        assert_eq!(e.comments().unwrap(), before);
    }
    let c = e.commit().unwrap();
    let q = publish(&p, &c);
    read(&q, 3, &first());
    preserved(&p, &q, "ppt/slides/_rels/slide4.xml.rels");
    assert!(native("powerpoint-navigation-review.pptx") == bytes);
}
