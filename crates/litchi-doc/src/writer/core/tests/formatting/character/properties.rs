use super::super::super::support::*;

#[test]
fn test_character_formatting_default() {
    let fmt = CharacterFormatting::default();
    assert!(fmt.bold.is_none());
    assert!(fmt.italic.is_none());
    assert!(fmt.underline.is_none());
    assert!(fmt.font_size.is_none());
}

#[test]
fn test_paragraph_formatting_default() {
    let fmt = ParagraphFormatting::default();
    assert!(fmt.alignment.is_none());
    assert!(fmt.left_indent.is_none());
    assert!(fmt.right_indent.is_none());
    assert!(fmt.space_before.is_none());
    assert!(fmt.space_after.is_none());
}

#[test]
fn test_line_spacing_default() {
    let ls = LineSpacing::default();
    assert_eq!(ls, LineSpacing::single());
    assert_eq!(ls.dya_line, 240);
    assert!(ls.is_multiple);
}
