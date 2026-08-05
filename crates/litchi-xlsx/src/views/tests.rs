use super::model::{PanePosition, PaneState, ViewType};

#[test]
fn view_type_tokens_round_trip() {
    for (value, token) in [
        (ViewType::Normal, "normal"),
        (ViewType::PageBreakPreview, "pageBreakPreview"),
        (ViewType::PageLayout, "pageLayout"),
    ] {
        assert_eq!(value.as_str(), token);
        assert_eq!(ViewType::parse(token), Some(value));
    }
    assert_eq!(ViewType::parse("pageLayoutView"), None);
}

#[test]
fn pane_position_tokens_round_trip() {
    for (value, token) in [
        (PanePosition::BottomRight, "bottomRight"),
        (PanePosition::TopRight, "topRight"),
        (PanePosition::BottomLeft, "bottomLeft"),
        (PanePosition::TopLeft, "topLeft"),
    ] {
        assert_eq!(value.as_str(), token);
        assert_eq!(PanePosition::parse(token), Some(value));
    }
    assert_eq!(PanePosition::parse("center"), None);
}

#[test]
fn pane_state_tokens_round_trip() {
    for (value, token) in [
        (PaneState::Split, "split"),
        (PaneState::Frozen, "frozen"),
        (PaneState::FrozenSplit, "frozenSplit"),
    ] {
        assert_eq!(value.as_str(), token);
        assert_eq!(PaneState::parse(token), Some(value));
    }
    assert_eq!(PaneState::parse("freeze"), None);
}
