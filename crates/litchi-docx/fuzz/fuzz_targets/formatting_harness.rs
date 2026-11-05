//! Bounded formatting and exact-inverse exercise for untrusted source XML.
use litchi_core::Position;
use litchi_docx::document::{BodyAnchor, NewParagraph, NewRun, raw};

pub fn exercise(data: &[u8]) -> bool {
    let Some((&selector, xml)) = data.split_first() else {
        return false;
    };
    if xml.len() > 64 * 1024 {
        return false;
    }
    let limits = raw::Limits::default();
    let Ok(mut edit) = raw::Tree::from_xml(xml.to_vec(), limits) else {
        return false;
    };
    let result = match selector % 5 {
        0 => edit
            .format_run(Position::new(0), Position::new(0), |run| {
                run.bold = Some(true)
            })
            .map(|_| ()),
        1 => edit
            .format_paragraph(Position::new(0), |paragraph| {
                paragraph.keep_next = Some(true)
            })
            .map(|_| ()),
        2 => edit
            .insert_formatted_paragraph(
                BodyAnchor::End,
                &NewParagraph::new().with_run(NewRun::new("fuzz")),
            )
            .map(|_| ()),
        3 => edit
            .append_run(Position::new(0), &NewRun::new("fuzz"))
            .map(|_| ()),
        _ => edit.remove_paragraph(Position::new(0)).map(|_| ()),
    };
    if result.is_err() {
        assert_eq!(edit.projected().xml_bytes(), xml);
        return false;
    }
    let Ok(patch) = edit.patch() else {
        return false;
    };
    let restored = edit
        .apply_patch(&patch.inverse(), limits)
        .expect("inverse of accepted edit");
    assert_eq!(restored.to_bytes(limits).unwrap(), xml);
    true
}
