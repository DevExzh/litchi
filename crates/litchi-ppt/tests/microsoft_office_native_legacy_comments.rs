//! Comment 2000 derivatives on native binary decks; no modern collaboration semantics.
use litchi_core::Position;
use litchi_ppt::{
    Package, ParsedComment, RecordType,
    comments::{CommentChange, DateTime, NewComment},
};
use std::{collections::BTreeMap, io::Cursor, path::Path};
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/powerpoint")
            .join(name),
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
fn comments(p: &Package) -> Vec<ParsedComment> {
    p.presentation()
        .unwrap()
        .slide_at(0)
        .unwrap()
        .unwrap()
        .comments()
        .unwrap()
}
struct TempPpt(std::path::PathBuf);
impl TempPpt {
    fn new() -> Self {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        Self(std::env::temp_dir().join(format!(
            "litchi-native-comment-{}-{nonce}.ppt",
            std::process::id()
        )))
    }
}
impl Drop for TempPpt {
    fn drop(&mut self) {
        let _ = std::fs::remove_file(&self.0);
    }
}
fn publish(source: &Package, commit: &litchi_ppt::package::edit::Commit) -> Package {
    let f = TempPpt::new();
    commit.package().save(&f.0).unwrap();
    let p = Package::open(&f.0).unwrap();
    let before = source.presentation().unwrap();
    let after = p.presentation().unwrap();
    assert_eq!(after.slide_count(), before.slide_count());
    let old_masters = before.masters().unwrap();
    let new_masters = after.masters().unwrap();
    assert_eq!(old_masters.main().len(), new_masters.main().len());
    for (a, b) in old_masters.main().iter().zip(new_masters.main()) {
        assert!(a.record().data == b.record().data);
    }
    assert_eq!(
        old_masters.notes().map(|m| m.record().data.as_ref()),
        new_masters.notes().map(|m| m.record().data.as_ref())
    );
    assert_eq!(
        old_masters.handout().map(|m| m.record().data.as_ref()),
        new_masters.handout().map(|m| m.record().data.as_ref())
    );

    let a = before.slides().unwrap();
    let b = after.slides().unwrap();
    for (i, (a, b)) in a.iter().zip(&b).enumerate() {
        assert_eq!(a.text().unwrap(), b.text().unwrap());
        if i > 0 {
            assert!(a.record().data == b.record().data);
        }
        let drawings = |r: &litchi_ppt::Record| {
            r.children
                .iter()
                .filter(|r| r.record_type == RecordType::PPDrawing)
                .map(|r| r.data.to_vec())
                .collect::<Vec<_>>()
        };
        assert!(drawings(a.record()) == drawings(b.record()));
    }
    let x = streams(&source.to_bytes().unwrap());
    let y = streams(&p.to_bytes().unwrap());
    assert_eq!(x.keys().collect::<Vec<_>>(), y.keys().collect::<Vec<_>>());
    for (name, data) in x {
        if name != ["PowerPoint Document"] && name != ["Current User"] {
            assert!(y.get(&name) == Some(&data), "changed {name:?}");
        }
    }
    p.presentation().unwrap().comment_catalog().unwrap();
    assert!(
        p.apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == source.to_bytes().unwrap()
    );
    assert!(
        source
            .apply(commit.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == p.to_bytes().unwrap()
    );
    p
}
#[test]
fn native_ppt_classic_comment_lifecycle_retains_slide_drawings_and_other_streams() {
    for name in [
        "powerpoint-master-motion.ppt",
        "powerpoint-visual-objects.ppt",
    ] {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        assert!(
            p.presentation()
                .unwrap()
                .comment_catalog()
                .unwrap()
                .comments()
                .next()
                .is_none()
        );
        let note = NewComment {
            author: "Native Reviewer".into(),
            initials: "NR".into(),
            text: "Legacy note 日本語 https://example.invalid/inert".into(),
            date: DateTime::parse("2026-11-07T10:20:30.125"),
            x: 120,
            y: 48,
        };
        let mut edit = p.edit().unwrap();
        assert_eq!(edit.add_comment(Position::new(0), &note).unwrap().get(), 0);
        let c = edit.commit().unwrap();
        let added = publish(&p, &c);
        let notes = comments(&added);
        assert_eq!(notes.len(), 1);
        let n = &notes[0];
        assert_eq!(
            (&n.author, &n.initials, &n.text),
            (&note.author, &note.initials, &note.text)
        );
        assert_eq!((n.index, n.x, n.y), (1, 120, 48));
        assert_eq!(
            (
                n.year,
                n.month,
                n.day,
                n.hour,
                n.minute,
                n.second,
                n.millisecond
            ),
            (2026, 11, 7, 10, 20, 30, 125)
        );
        let mut edit = added.edit().unwrap();
        edit.set_comment(
            Position::new(0),
            Position::new(0),
            &CommentChange {
                text: Some("Reviewed classic note".into()),
                x: Some(240),
                y: None,
            },
        )
        .unwrap();
        let c = edit.commit().unwrap();
        let updated = publish(&added, &c);
        let changed = comments(&updated);
        assert_eq!(changed.len(), 1);
        assert_eq!(changed[0].text, "Reviewed classic note");
        assert_eq!((changed[0].index, changed[0].x, changed[0].y), (1, 240, 48));
        assert_eq!(changed[0].author, n.author);
        let mut edit = updated.edit().unwrap();
        edit.remove_comment(Position::new(0), Position::new(0))
            .unwrap();
        let c = edit.commit().unwrap();
        let removed = publish(&updated, &c);
        assert!(comments(&removed).is_empty());
        assert!(p.to_bytes().unwrap() == bytes);
        assert!(native(name) == bytes);
    }
}
