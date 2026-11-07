//! Real comment threads and inert Litchi-authored task/reaction derivatives.
use litchi_pptx::modern_comments::semantic::{reactions, tasks};
use litchi_pptx::modern_comments::{self as modern, AnchorKind, Comment, Progress, Status};
use litchi_pptx::presentation::modern_comments::raw;
use litchi_pptx::{Error, Package, SourceBackedPresentation};
use std::{collections::BTreeMap, io::Cursor};
const FILE: &str = "powerpoint-navigation-review.pptx";
const PART: &str = "ppt/comments/modernComment_103_9D759682.xml";
const AUTHOR: &str = "{82E4CB02-5A72-9BFD-724D-897B10AA4E6A}";
const COMMENT: &str = "{D7A76B2D-DFBA-E946-BBF1-5135CC7B3682}";
const REPLY: &str = "{79313414-C4BC-F348-8E2A-3A08BC17B204}";
const NEW_REPLY: &str = "{11111111-2222-4333-8444-555555555555}";
const ROOT_TEXT: &str = "PPT_COMMENT_ROOT_003: verify the internal destination and hidden slide.";
const REPLY_TEXT: &str = "PPT_COMMENT_REPLY_003: confirmed section and custom-show membership.";
// Caller-selected extension URIs; no native application recognition is claimed.
const TASK_URI: &str = "urn:litchi:native-comment-test:task";
const REACTION_URI: &str = "urn:litchi:native-comment-test:reaction";
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

// Inverse replay after reopening restores exact OPC payloads. The edited ZIP
// member is recompressed; its producer-specific compression metadata is not owned.
#[track_caller]
fn assert_members_equal(a: &[u8], b: &[u8]) {
    assert_eq!(changed(a, b), Vec::<String>::new());
}

fn package() -> Package {
    Package::from_reader(Cursor::new(native(FILE))).unwrap()
}
fn save(p: &Package) -> Package {
    let f = tempfile::NamedTempFile::with_suffix(".pptx").unwrap();
    p.save(f.path()).unwrap();
    Package::open(f.path()).unwrap()
}
fn comment(p: &Package) -> Comment {
    p.edit_modern_comments()
        .unwrap()
        .comment(3, COMMENT)
        .unwrap()
        .unwrap()
}
fn assert_native_content(p: &Package, expected: &Comment) {
    let c = comment(p);
    assert_eq!(c.id, expected.id);
    assert_eq!(c.author_id, expected.author_id);
    assert_eq!(c.created, expected.created);
    assert_eq!(c.anchors, expected.anchors);
    assert_eq!(c.position, expected.position);
    assert_eq!(c.text_body_xml, expected.text_body_xml);
    assert_eq!(c.replies.len(), 1);
    assert_eq!(c.replies[0].id, expected.replies[0].id);
    assert_eq!(c.replies[0].author_id, expected.replies[0].author_id);
    assert_eq!(
        c.replies[0].text_body_xml,
        expected.replies[0].text_body_xml
    );
}
#[test]
fn native_text_anchored_threads_match_eager_and_source_semantics() {
    for (name, index) in [(FILE, 3), ("powerpoint-master-motion.pptx", 4)] {
        let bytes = native(name);
        let p = Package::from_reader(Cursor::new(&bytes)).unwrap();
        let e = p.edit_modern_comments().unwrap();
        assert!(e.changes().unwrap().is_empty());
        let g = e.graph().unwrap();
        let authors = g.authors.as_ref().unwrap();
        assert_eq!(authors.part_name, "/ppt/authors.xml");
        assert_eq!(authors.authors.authors.len(), 1);
        let a = &authors.authors.authors[0];
        assert_eq!(a.id, AUTHOR);
        assert_eq!(a.name, "Ryker Zhu");
        assert_eq!(a.initials.as_deref(), Some("RZ"));
        assert_eq!(a.user_id, "Ryker Zhu");
        assert_eq!(a.provider_id, "None");
        assert_eq!(g.comments.len(), 1);
        assert_eq!(g.comments[0].part_name, format!("/{PART}"));
        assert_eq!(
            g.comments[0].slide_part_name,
            format!("/ppt/slides/slide{}.xml", index + 1)
        );
        let c = e.comment(index, COMMENT).unwrap().unwrap();
        assert_eq!(c.created, "2026-10-15T07:23:09.649");
        assert_eq!(c.status, None);
        assert_eq!(c.complete, None);
        assert_eq!(c.assigned_to, None);
        assert_eq!(c.start_date, None);
        assert_eq!(c.due_date, None);
        assert_eq!(c.title, None);
        assert_eq!(
            c.position,
            Some(modern::Position {
                x: 5649097,
                y: 485089
            })
        );
        assert_eq!(c.anchors.len(), 1);
        assert_eq!(c.anchors[0].kind, AnchorKind::TextRangeMoniker);
        let anchor = std::str::from_utf8(&c.anchors[0].xml).unwrap();
        for token in [
            "sldId=\"259\"",
            "cId=\"2641729154\"",
            "id=\"3\"",
            "cp=\"0\"",
            "len=\"52\"",
            "hash=\"270430480\"",
            "{F67C6603-5192-0FA9-21C1-5021739A0453}",
        ] {
            assert!(anchor.contains(token), "{token}");
        }
        assert_eq!(c.task_details().unwrap(), None);
        assert_eq!(c.reactions().unwrap(), None);
        assert_eq!(c.replies.len(), 1);
        assert_eq!(c.replies[0].id, REPLY);
        assert_eq!(c.replies[0].created, "2026-10-15T07:23:18.033");
        assert_eq!(c.replies[0].reactions().unwrap(), None);
        let eager = p
            .presentation()
            .unwrap()
            .slide(index)
            .unwrap()
            .unwrap()
            .modern_comments()
            .unwrap()
            .unwrap();
        let source = SourceBackedPresentation::from_reader(Cursor::new(&bytes)).unwrap();
        let read = source
            .slide(index)
            .unwrap()
            .modern_comments()
            .unwrap()
            .unwrap();
        for list in [&eager, &read] {
            assert_eq!(list.entries().len(), 1);
            let c = &list.entries()[0];
            assert_eq!(c.text(), Some(ROOT_TEXT));
            assert_eq!(c.replies()[0].text(), Some(REPLY_TEXT));
            assert_eq!(c.author_name(), Some("Ryker Zhu"));
            assert_eq!(c.replies()[0].author_name(), Some("Ryker Zhu"));
            assert_eq!(c.status(), Status::Active);
            assert_eq!(c.progress(), Progress::ZERO);
            assert_eq!(raw::comment_id(c), COMMENT);
            assert_eq!(raw::reply_id(&c.replies()[0]), REPLY);
        }
        assert_eq!(raw::data(&eager), raw::data(&read));
        assert_eq!(raw::author_data(&eager), raw::author_data(&read));
        assert_eq!(e.commit().unwrap().package().to_bytes().unwrap(), bytes);
    }
}
#[test]
fn native_thread_task_fields_preserve_anchor_reply_and_failed_staging() {
    let p = package();
    let bytes = p.to_bytes().unwrap();
    let original = comment(&p);
    let mut e = p.edit_modern_comments().unwrap();
    assert!(
        e.update(3, COMMENT, |c| {
            c.title = Some("Task <review> & verify".into());
            c.complete = Some(Progress::new(25).unwrap());
            c.assigned_to = Some(vec![AUTHOR.into()]);
            c.start_date = Some("2026-11-07T10:00:00Z".into());
            c.due_date = Some("2026-11-08T18:00:00Z".into());
            c.status = Some(Status::Active);
        })
        .unwrap()
    );
    let staged = e.comment(3, COMMENT).unwrap().unwrap();
    assert!(matches!(
        e.update(3, COMMENT, |c| c.assigned_to = Some(vec![NEW_REPLY.into()])),
        Err(Error::Invalid(_))
    ));
    assert!(matches!(
        e.update(3, COMMENT, |c| c.id = NEW_REPLY.into()),
        Err(Error::Invalid(_))
    ));
    assert_eq!(e.comment(3, COMMENT).unwrap(), Some(staged.clone()));
    let c = e.commit().unwrap();
    let out = save(c.package());
    let output = out.to_bytes().unwrap();
    assert_eq!(changed(&bytes, &output), [PART]);
    assert_eq!(comment(&out), staged);
    assert!(
        std::str::from_utf8(&members(&output)[PART])
            .unwrap()
            .contains(r#"complete="25000""#)
    );
    assert_native_content(&out, &original);
    let view = out
        .presentation()
        .unwrap()
        .slide(3)
        .unwrap()
        .unwrap()
        .modern_comments()
        .unwrap()
        .unwrap();
    assert_eq!(view.entries()[0].title(), Some("Task <review> & verify"));
    assert_eq!(view.entries()[0].progress(), Progress::new(25).unwrap());
    let source = SourceBackedPresentation::from_reader(Cursor::new(&output)).unwrap();
    let source_view = source.slide(3).unwrap().modern_comments().unwrap().unwrap();
    let value = &source_view.entries()[0];
    assert_eq!(value.title(), Some("Task <review> & verify"));
    assert_eq!(value.progress(), Progress::new(25).unwrap());
    assert_eq!(value.start_date(), Some("2026-11-07T10:00:00Z"));
    assert_eq!(value.due_date(), Some("2026-11-08T18:00:00Z"));
    assert_eq!(raw::assigned_to(value).unwrap(), [AUTHOR]);
    assert!(matches!(
        out.apply_modern_comments(c.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(
        p.apply_modern_comments(c.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        output
    );
    assert_members_equal(
        &out.apply_modern_comments(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        &bytes,
    );
    let mut e = out.edit_modern_comments().unwrap();
    e.update(3, COMMENT, |_| {}).unwrap();
    assert_eq!(e.commit().unwrap().package().to_bytes().unwrap(), output);
    assert_eq!(p.to_bytes().unwrap(), bytes);
}
fn task_details() -> tasks::Details {
    tasks::Details {
        history: tasks::History {
            events: vec![tasks::Event {
                time: "2026-11-07T10:00:00Z".into(),
                id: "{11111111-2222-4333-8444-000000000001}".into(),
                attributed_by: tasks::User {
                    author_id: AUTHOR.into(),
                },
                anchor: Some(tasks::Anchor {
                    comment_id: COMMENT.into(),
                    extension_xml: None,
                    namespace_declarations: vec![],
                }),
                action: Some(tasks::Action::Assign(tasks::Assign {
                    author_id: AUTHOR.into(),
                })),
                extension_xml: None,
                namespace_declarations: vec![],
            }],
        },
        extension_xml: None,
        namespace_declarations: vec![],
    }
}
fn reactions(kind: &str) -> reactions::List {
    reactions::List {
        reactions: vec![reactions::Reaction {
            reaction_type: kind.into(),
            instances: vec![reactions::Instance {
                time: "2026-11-07T10:01:00Z".into(),
                author_id: AUTHOR.into(),
                extension_xml: None,
                namespace_declarations: vec![],
            }],
            namespace_declarations: vec![],
        }],
        namespace_declarations: vec![],
    }
}
fn assert_reactions(v: &reactions::List, kind: &str) {
    assert_eq!(v.reactions.len(), 1);
    assert_eq!(v.reactions[0].reaction_type, kind);
    assert_eq!(v.reactions[0].instances.len(), 1);
    assert_eq!(v.reactions[0].instances[0].author_id, AUTHOR);
    assert_eq!(v.reactions[0].instances[0].time, "2026-11-07T10:01:00Z");
}
#[test]
fn native_source_task_history_and_reaction_derivatives_add_update_clear_and_invert() {
    let p = package();
    let before = p.to_bytes().unwrap();
    let original = comment(&p);
    let mut e = p.edit_modern_comments().unwrap();
    e.update_extensions(3, COMMENT, |x| {
        x.replace_task_details(Some(TASK_URI), Some(task_details()))
            .unwrap();
        x.replace_reactions(Some(REACTION_URI), Some(reactions("👍")))
            .unwrap();
    })
    .unwrap();
    e.update_reply_extensions(3, COMMENT, REPLY, |x| {
        x.replace_reactions(Some(REACTION_URI), Some(reactions("❤️")))
            .unwrap()
    })
    .unwrap();
    let staged = e.comment(3, COMMENT).unwrap().unwrap();
    assert!(matches!(
        e.update_extensions(3, COMMENT, |x| {
            let mut task = x.task_details().unwrap().clone();
            task.history.events[0].attributed_by.author_id = "not-a-guid".into();
            let value = x
                .entries
                .iter_mut()
                .find_map(|entry| match &mut entry.payload {
                    modern::Payload::TaskDetails(v) => Some(v),
                    _ => None,
                })
                .unwrap();
            *value = task;
        }),
        Err(Error::Invalid(_))
    ));
    assert_eq!(e.comment(3, COMMENT).unwrap(), Some(staged));
    let added = e.commit().unwrap();
    let out = save(added.package());
    let bytes = out.to_bytes().unwrap();
    assert_eq!(changed(&before, &bytes), [PART]);
    assert_native_content(&out, &original);
    let c = comment(&out);
    let task = c.task_details().unwrap().unwrap();
    assert_eq!(task.history.events.len(), 1);
    let event = &task.history.events[0];
    assert_eq!(event.time, "2026-11-07T10:00:00Z");
    assert_eq!(event.id, "{11111111-2222-4333-8444-000000000001}");
    assert_eq!(event.attributed_by.author_id, AUTHOR);
    assert_eq!(event.anchor.as_ref().unwrap().comment_id, COMMENT);
    assert_eq!(
        event.action,
        Some(tasks::Action::Assign(tasks::Assign {
            author_id: AUTHOR.into()
        }))
    );
    assert_reactions(&c.reactions().unwrap().unwrap(), "👍");
    assert_reactions(&c.replies[0].reactions().unwrap().unwrap(), "❤️");
    let mut e = out.edit_modern_comments().unwrap();
    e.update_extensions(3, COMMENT, |x| {
        let mut task = x.task_details().unwrap().clone();
        task.history.events[0].action = Some(tasks::Action::Progress(Progress::new(50).unwrap()));
        x.replace_task_details(None, Some(task)).unwrap();
        x.replace_reactions(None, Some(reactions("❤️"))).unwrap();
    })
    .unwrap();
    let updated = e.commit().unwrap();
    let newer = save(updated.package());
    let c = comment(&newer);
    assert_eq!(
        c.task_details().unwrap().unwrap().history.events[0].action,
        Some(tasks::Action::Progress(Progress::new(50).unwrap()))
    );
    assert_reactions(&c.reactions().unwrap().unwrap(), "❤️");
    assert_native_content(&newer, &original);
    let mut e = newer.edit_modern_comments().unwrap();
    e.update_extensions(3, COMMENT, |x| {
        x.replace_task_details(None, None).unwrap();
        x.replace_reactions(None, None).unwrap();
    })
    .unwrap();
    e.update_reply_extensions(3, COMMENT, REPLY, |x| {
        x.replace_reactions(None, None).unwrap()
    })
    .unwrap();
    let cleared = e.commit().unwrap();
    let empty = save(cleared.package());
    assert_eq!(comment(&empty), original);
    assert_eq!(changed(&before, &empty.to_bytes().unwrap()), [PART]);
    assert_eq!(
        empty
            .apply_modern_comments(&cleared.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        newer.to_bytes().unwrap()
    );
    assert_eq!(
        newer
            .apply_modern_comments(&updated.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        bytes
    );
    assert_members_equal(
        &out.apply_modern_comments(&added.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        &before,
    );
    assert_eq!(native(FILE), before);
}
#[test]
fn native_thread_reply_crud_and_author_reference_guards_preserve_other_parts() {
    let p = package();
    let before = p.to_bytes().unwrap();
    let initial = comment(&p);
    let mut reply = initial.replies[0].clone();
    reply.id = NEW_REPLY.into();
    let mut e = p.edit_modern_comments().unwrap();
    e.add_reply(3, COMMENT, reply).unwrap();
    assert!(matches!(e.remove_author(AUTHOR), Err(Error::Invalid(_))));
    assert_eq!(e.comment(3, COMMENT).unwrap().unwrap().replies.len(), 2);
    let added = e.commit().unwrap();
    let out = save(added.package());
    assert_eq!(changed(&before, &out.to_bytes().unwrap()), [PART]);
    assert_eq!(
        comment(&out)
            .replies
            .iter()
            .map(|r| r.id.as_str())
            .collect::<Vec<_>>(),
        [REPLY, NEW_REPLY]
    );
    let mut e = out.edit_modern_comments().unwrap();
    e.update_reply(3, COMMENT, NEW_REPLY, |r| r.status = Some(Status::Resolved))
        .unwrap();
    let updated = e.commit().unwrap();
    let resolved = save(updated.package());
    assert_eq!(comment(&resolved).replies[1].status, Some(Status::Resolved));
    let mut e = resolved.edit_modern_comments().unwrap();
    assert!(e.remove_reply(3, COMMENT, NEW_REPLY).unwrap());
    assert!(!e.remove_reply(3, COMMENT, NEW_REPLY).unwrap());
    let removed = e.commit().unwrap();
    let restored = save(removed.package());
    assert_eq!(comment(&restored), initial);
    assert_eq!(
        restored
            .apply_modern_comments(&removed.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        resolved.to_bytes().unwrap()
    );
    assert_members_equal(
        &out.apply_modern_comments(&added.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        &before,
    );
    assert_eq!(p.to_bytes().unwrap(), before);
}
