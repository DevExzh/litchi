use litchi_rtf::{RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_kinsoku_sets_and_language_and_round_trips() {
    let document = RtfDocument::parse(
        r"{\rtf1\ansi{\*\fchars .,;:\u12305?}{\*\lchars \u12304?ab}\ksulang1041 Body}",
    )
    .unwrap();
    assert_eq!(document.text(), "Body");
    let kinsoku = document.kinsoku();
    assert_eq!(kinsoku.following.as_deref(), Some(".,;:】"));
    assert_eq!(kinsoku.leading.as_deref(), Some("【ab"));
    assert_eq!(kinsoku.language, Some(1041));

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.kinsoku(), kinsoku);
}

#[test]
fn parses_partial_kinsoku_metadata() {
    let document = RtfDocument::parse(r"{\rtf1{\*\lchars abc}Body}").unwrap();
    let kinsoku = document.kinsoku();
    assert_eq!(kinsoku.leading.as_deref(), Some("abc"));
    assert!(kinsoku.following.is_none());
    assert!(kinsoku.language.is_none());

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.kinsoku(), kinsoku);
}

#[test]
fn parses_upr_wrapped_cjk_kinsoku_from_word() {
    // The shape Word writes in codepage CJK document headers: the ANSI
    // branch of \upr is skipped and the Unicode branch carries the set,
    // including {\ucN\uN } encoding-switch groups.
    let document = RtfDocument::parse(concat!(
        r"{\rtf1\ansi{\upr{\*\fchars x}{\*\ud\uc0{\*\fchars ),.{\uc2\u162 }}}}",
        r"{\upr{\*\lchars x}{\*\ud\uc1{\*\lchars \u12304?(}}}Body}",
    ))
    .unwrap();
    assert_eq!(document.text(), "Body");
    let kinsoku = document.kinsoku();
    assert_eq!(kinsoku.following.as_deref(), Some("),.¢"));
    assert_eq!(kinsoku.leading.as_deref(), Some("【("));

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.kinsoku(), kinsoku);
}

#[test]
fn rejects_malformed_kinsoku_metadata() {
    let cases = [
        // Unstarred destinations.
        r"{\rtf1{\fchars x}Body}",
        r"{\rtf1{\lchars x}Body}",
        // Duplicate destinations.
        r"{\rtf1{\*\fchars x}{\*\fchars y}Body}",
        r"{\rtf1{\*\lchars x}{\*\lchars y}Body}",
        // Empty character set.
        r"{\rtf1{\*\fchars }Body}",
        // Destination nested below document root inside a body group.
        r"{\rtf1{\*\atrfstart 1}{\*\fchars x}}Body}",
        // Grouped data inside a destination.
        r"{\rtf1{\*\fchars x{y}}Body}",
        // Missing, negative, or duplicate ksulang parameter.
        r"{\rtf1\ksulang Body}",
        r"{\rtf1\ksulang-1 Body}",
        r"{\rtf1\ksulang1041\ksulang1041 Body}",
        // ksulang nested inside a body group.
        r"{\rtf1{\*\atrfstart 1}\ksulang1041 Body}",
    ];
    for rtf in cases {
        assert!(RtfDocument::parse(rtf).is_err(), "accepted malformed {rtf}");
    }
}
