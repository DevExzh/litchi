#![cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]

fn assert_send_sync_static<T: Send + Sync + 'static>() {}

#[cfg(feature = "pages")]
#[test]
fn pages_semantic_namespace_is_concrete_and_archive_free() {
    use litchi::pages::semantic::{
        Document, DocumentStats, Error, LimitKind, OpenOptions, PackageError, Position, Result,
        Root, Section, SectionSelector, SectionType, SelectorError, SelectorResult,
        SemanticLimitKind, SemanticLimits, SourceLimitKind, SourceLimits, TextPosition, TextSpan,
    };

    assert_send_sync_static::<Document>();
    assert_send_sync_static::<OpenOptions>();
    assert_send_sync_static::<SourceLimitKind>();
    assert_send_sync_static::<SourceLimits>();
    assert_send_sync_static::<LimitKind>();
    assert_send_sync_static::<DocumentStats>();
    assert_send_sync_static::<Error>();
    assert_send_sync_static::<PackageError>();
    assert_send_sync_static::<Result<()>>();
    assert_send_sync_static::<Root>();
    assert_send_sync_static::<Section>();
    assert_send_sync_static::<SectionSelector<'static>>();
    assert_send_sync_static::<SectionType>();
    assert_send_sync_static::<SemanticLimitKind>();
    assert_send_sync_static::<SemanticLimits>();
    assert_send_sync_static::<SelectorError>();
    assert_send_sync_static::<SelectorResult<()>>();
    assert_send_sync_static::<TextPosition>();
    assert_send_sync_static::<TextSpan>();
    assert_send_sync_static::<Position>();

    let options = OpenOptions::new()
        .source_limits(SourceLimits::default())
        .semantic_limits(SemanticLimits::default());
    assert_eq!(options.source_limits_value(), SourceLimits::default());
    assert_eq!(options.semantic_limits_value(), SemanticLimits::default());

    let selector = SectionSelector::name("Introduction");
    assert!(matches!(selector, SectionSelector::Name("Introduction")));
}

#[cfg(feature = "keynote")]
#[test]
fn keynote_semantic_namespace_is_concrete_and_archive_free() {
    use litchi::keynote::semantic::{
        Build, Error, LimitKind, Mode, OpenOptions, Position, Presentation, Result, Seconds,
        SemanticLimitKind, SemanticLimits, Settings, Show, Size, Slide, SlideSelector,
        SlideSelectorError, SlideSelectorResult, SourceLimitKind, SourceLimits, TextPosition,
        TextSpan, Transition, presentation,
    };

    assert_send_sync_static::<Presentation>();
    assert_send_sync_static::<OpenOptions>();
    assert_send_sync_static::<Build>();
    assert_send_sync_static::<presentation::IoKind>();
    assert_send_sync_static::<presentation::Error>();
    assert_send_sync_static::<presentation::LimitKind>();
    assert_send_sync_static::<presentation::Stats>();
    assert_send_sync_static::<LimitKind>();
    assert_send_sync_static::<SemanticLimitKind>();
    assert_send_sync_static::<SemanticLimits>();
    assert_send_sync_static::<SourceLimitKind>();
    assert_send_sync_static::<SourceLimits>();
    assert_send_sync_static::<Error>();
    assert_send_sync_static::<Mode>();
    assert_send_sync_static::<Position>();
    assert_send_sync_static::<Result<()>>();
    assert_send_sync_static::<Seconds>();
    assert_send_sync_static::<Settings>();
    assert_send_sync_static::<Show>();
    assert_send_sync_static::<Size>();
    assert_send_sync_static::<Slide>();
    assert_send_sync_static::<SlideSelector<'static>>();
    assert_send_sync_static::<SlideSelectorError>();
    assert_send_sync_static::<SlideSelectorResult<()>>();
    assert_send_sync_static::<TextPosition>();
    assert_send_sync_static::<TextSpan>();
    assert_send_sync_static::<Transition>();

    let options = OpenOptions::new()
        .source_limits(SourceLimits::default())
        .semantic_limits(SemanticLimits::default());
    assert_eq!(options.source_limits_value(), SourceLimits::default());
    assert_eq!(options.semantic_limits_value(), SemanticLimits::default());

    let selector = SlideSelector::position(Position::new(2));
    assert_eq!(selector.as_position(), Some(Position::new(2)));
    assert!(selector.as_name().is_none());
}

#[cfg(feature = "numbers")]
#[test]
fn numbers_semantic_namespace_is_concrete_and_archive_free() {
    use litchi::numbers::semantic::{
        Error, OpenOptions, ReadError, ReadLimitKind, Result, SemanticLimitKind, SemanticLimits,
        Sheet, SheetSelector, SourceLimitKind, SourceLimits, Stats, Table, TableSelector,
        TableSelectorError, Workbook,
    };

    assert_send_sync_static::<Workbook>();
    assert_send_sync_static::<OpenOptions>();
    assert_send_sync_static::<ReadError>();
    assert_send_sync_static::<SourceLimitKind>();
    assert_send_sync_static::<SourceLimits>();
    assert_send_sync_static::<Stats>();
    assert_send_sync_static::<Error>();
    assert_send_sync_static::<SemanticLimitKind>();
    assert_send_sync_static::<SemanticLimits>();
    assert_send_sync_static::<ReadLimitKind>();
    assert_send_sync_static::<Result<()>>();
    assert_send_sync_static::<Sheet>();
    assert_send_sync_static::<SheetSelector<'static>>();
    assert_send_sync_static::<Table>();
    assert_send_sync_static::<TableSelector<'static>>();
    assert_send_sync_static::<TableSelectorError>();

    let options = OpenOptions::new()
        .source_limits(SourceLimits::default())
        .semantic_limits(SemanticLimits::default());
    assert_eq!(options.source_limits_value(), SourceLimits::default());
    assert_eq!(options.semantic_limits_value(), SemanticLimits::default());

    let sheet = SheetSelector::name("Summary");
    let table = TableSelector::index(0);
    assert!(matches!(sheet, SheetSelector::Name("Summary")));
    assert_eq!(table.as_position().map(|position| position.get()), Some(0));
}
