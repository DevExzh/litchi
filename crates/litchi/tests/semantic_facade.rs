#![cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]

fn assert_send_sync_static<T: Send + Sync + 'static>() {}

#[cfg(feature = "pages")]
#[test]
fn pages_semantic_namespace_is_concrete_and_archive_free() {
    use litchi::pages::semantic::{
        Document, DocumentReadOptions, DocumentSourceLimitKind, DocumentSourceLimits,
        DocumentSourceLimitsError, DocumentStats, Error, IoKind, Position, ReadError,
        ReadLimitKind, Result, Root, Section, SectionSelector, SectionType, SelectorError,
        SelectorResult, SemanticLimitKind, SemanticLimits, SemanticLimitsError, TextPosition,
        TextSpan,
    };

    assert_send_sync_static::<Document>();
    assert_send_sync_static::<DocumentReadOptions>();
    assert_send_sync_static::<DocumentSourceLimitKind>();
    assert_send_sync_static::<DocumentSourceLimits>();
    assert_send_sync_static::<DocumentSourceLimitsError>();
    assert_send_sync_static::<DocumentStats>();
    assert_send_sync_static::<Error>();
    assert_send_sync_static::<IoKind>();
    assert_send_sync_static::<ReadError>();
    assert_send_sync_static::<ReadLimitKind>();
    assert_send_sync_static::<Result<()>>();
    assert_send_sync_static::<Root>();
    assert_send_sync_static::<Section>();
    assert_send_sync_static::<SectionSelector<'static>>();
    assert_send_sync_static::<SectionType>();
    assert_send_sync_static::<SemanticLimitKind>();
    assert_send_sync_static::<SemanticLimits>();
    assert_send_sync_static::<SemanticLimitsError>();
    assert_send_sync_static::<SelectorError>();
    assert_send_sync_static::<SelectorResult<()>>();
    assert_send_sync_static::<TextPosition>();
    assert_send_sync_static::<TextSpan>();
    assert_send_sync_static::<Position>();

    let options =
        DocumentReadOptions::new(DocumentSourceLimits::default(), SemanticLimits::default());
    assert_eq!(options.source(), DocumentSourceLimits::default());
    assert_eq!(options.semantic(), SemanticLimits::default());

    let selector = SectionSelector::name("Introduction");
    assert!(matches!(selector, SectionSelector::Name("Introduction")));
}

#[cfg(feature = "keynote")]
#[test]
fn keynote_semantic_namespace_is_concrete_and_archive_free() {
    use litchi::keynote::semantic::{
        Build, Document, DocumentIoKind, DocumentReadError, DocumentReadLimitKind,
        DocumentReadOptions, DocumentSemanticLimitKind, DocumentSemanticLimits,
        DocumentSemanticLimitsError, DocumentSourceLimitKind, DocumentSourceLimits,
        DocumentSourceLimitsError, DocumentStats, Error, Mode, Position, Result, Seconds, Settings,
        Show, Size, Slide, SlideSelector, SlideSelectorError, SlideSelectorResult, TextPosition,
        TextSpan, Transition,
    };

    assert_send_sync_static::<Document>();
    assert_send_sync_static::<DocumentReadOptions>();
    assert_send_sync_static::<Build>();
    assert_send_sync_static::<DocumentIoKind>();
    assert_send_sync_static::<DocumentReadError>();
    assert_send_sync_static::<DocumentReadLimitKind>();
    assert_send_sync_static::<DocumentSemanticLimitKind>();
    assert_send_sync_static::<DocumentSemanticLimits>();
    assert_send_sync_static::<DocumentSemanticLimitsError>();
    assert_send_sync_static::<DocumentSourceLimitKind>();
    assert_send_sync_static::<DocumentSourceLimits>();
    assert_send_sync_static::<DocumentSourceLimitsError>();
    assert_send_sync_static::<DocumentStats>();
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

    let options = DocumentReadOptions::new(
        DocumentSourceLimits::default(),
        DocumentSemanticLimits::default(),
    );
    assert_eq!(options.source(), DocumentSourceLimits::default());
    assert_eq!(options.semantic(), DocumentSemanticLimits::default());

    let selector = SlideSelector::position(Position::new(2));
    assert_eq!(selector.as_position(), Some(Position::new(2)));
    assert!(selector.as_name().is_none());
}

#[cfg(feature = "numbers")]
#[test]
fn numbers_semantic_namespace_is_concrete_and_archive_free() {
    use litchi::numbers::semantic::{
        Document, DocumentReadError, DocumentReadOptions, DocumentSourceLimitKind,
        DocumentSourceLimits, DocumentSourceLimitsError, DocumentStats, Error, LimitKind, Limits,
        LimitsError, ReadLimitKind, Result, Sheet, SheetSelector, Table, TableSelector,
        TableSelectorError,
    };

    assert_send_sync_static::<Document>();
    assert_send_sync_static::<DocumentReadOptions>();
    assert_send_sync_static::<DocumentReadError>();
    assert_send_sync_static::<DocumentSourceLimitKind>();
    assert_send_sync_static::<DocumentSourceLimits>();
    assert_send_sync_static::<DocumentSourceLimitsError>();
    assert_send_sync_static::<DocumentStats>();
    assert_send_sync_static::<Error>();
    assert_send_sync_static::<LimitKind>();
    assert_send_sync_static::<Limits>();
    assert_send_sync_static::<LimitsError>();
    assert_send_sync_static::<ReadLimitKind>();
    assert_send_sync_static::<Result<()>>();
    assert_send_sync_static::<Sheet>();
    assert_send_sync_static::<SheetSelector<'static>>();
    assert_send_sync_static::<Table>();
    assert_send_sync_static::<TableSelector<'static>>();
    assert_send_sync_static::<TableSelectorError>();

    let options = DocumentReadOptions::new(DocumentSourceLimits::default(), Limits::default());
    assert_eq!(options.source(), DocumentSourceLimits::default());
    assert_eq!(options.semantic(), Limits::default());

    let sheet = SheetSelector::name("Summary");
    let table = TableSelector::index(0);
    assert!(matches!(sheet, SheetSelector::Name("Summary")));
    assert!(matches!(table, TableSelector::Index(0)));
}
