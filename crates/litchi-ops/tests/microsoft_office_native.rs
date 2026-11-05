//! Public operation contracts on native Microsoft Office specimens.
#![cfg(feature = "pptx")]

use litchi_ops::{BatchItem, BatchMode, Code, OpenOptions, Props, SourceMode, props};

#[test]
fn native_ink_deck_missing_notes_have_the_same_first_error_in_all_edit_forms() {
    let bytes =
        include_bytes!("../../../test-data/microsoft-office/powerpoint/powerpoint-ink-replay.pptx");
    for mode in [SourceMode::Auto, SourceMode::Eager] {
        let file =
            litchi_ops::open_bytes_with(bytes.to_vec(), &OpenOptions::new().source_mode(mode))
                .unwrap();
        let before = file.to_bytes().unwrap();
        let path = "/slide[1]/notes/paragraph[1]";
        assert_eq!(
            file.set(path, props! {"align"=>"right"})
                .unwrap_err()
                .code(),
            Code::PathNotFound
        );
        let mut first = BatchItem::new("set");
        first.path = Some(path.into());
        first.props = Some(Props::new().with("align", "right"));
        let mut second = BatchItem::new("set");
        second.path = Some("/slide[1]/shape[1]".into());
        second.props = Some(Props::new().with("text", "must not publish"));
        let error = file.batch([first, second], BatchMode::Atomic).unwrap_err();
        assert_eq!(error.code(), Code::PathNotFound);
        assert_eq!(error.item(), Some(0));
        let error = file
            .edit(|edit| {
                edit.set(path, props! {"align"=>"right"})?;
                edit.set("/slide[1]/shape[1]", props! {"text"=>"must not publish"})
            })
            .unwrap_err();
        assert_eq!(error.code(), Code::PathNotFound);
        assert_eq!(error.item(), Some(0));
        assert!(file.to_bytes().unwrap() == before);
    }
}
