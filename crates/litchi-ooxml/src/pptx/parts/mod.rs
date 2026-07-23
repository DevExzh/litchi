/// Parts for PowerPoint presentation documents.
///
/// This module contains wrapper types for different XML parts in a .pptx package,
/// following the structure of the python-pptx library.
pub mod chart;
pub mod chart_ex;
pub mod chart_ex_style;
pub mod comment;
pub mod presentation;
pub mod slide;
pub mod theme;

pub use chart::{
    ChartData, ChartInfo, ChartPart, ChartSeries, ChartType, generate_chart_graphic_frame,
    generate_chart_xml,
};
pub use chart_ex::{
    CHART_EX_CONTENT_TYPE, ChartExAxis, ChartExAxisScaling, ChartExAxisTitle, ChartExAxisUnit,
    ChartExAxisUnits, ChartExAxisUnitsLabel, ChartExBinning, ChartExBinningChoice, ChartExChart,
    ChartExChartSpaceFormatting, ChartExChartTitle, ChartExClosedSide, ChartExColorKind,
    ChartExColorPosition, ChartExDataLabel, ChartExDataLabelPosition, ChartExDataLabelVisibility,
    ChartExDataLabels, ChartExDataPoint, ChartExDataSet, ChartExDimension, ChartExDocument,
    ChartExDoubleOrAutomatic, ChartExDrawingPayload, ChartExElementVisibility, ChartExExternalData,
    ChartExExternalDataTarget, ChartExFormatOverride, ChartExFormula, ChartExFormulaDirection,
    ChartExGeoAddress, ChartExGeoCache, ChartExGeoCacheEntry, ChartExGeoChildEntitiesQuery,
    ChartExGeoChildEntitiesQueryResult, ChartExGeoClear, ChartExGeoData, ChartExGeoDataEntityQuery,
    ChartExGeoDataEntityQueryResult, ChartExGeoDataPointQuery, ChartExGeoDataPointToEntityQuery,
    ChartExGeoDataPointToEntityQueryResult, ChartExGeoEntity, ChartExGeoEntityType,
    ChartExGeoHierarchyEntity, ChartExGeoLocation, ChartExGeoLocationQuery,
    ChartExGeoLocationQueryResult, ChartExGeoMappingLevel, ChartExGeoParentEntitiesQueryResult,
    ChartExGeoPolygon, ChartExGeoProjection, ChartExGeography, ChartExGridlines,
    ChartExHeaderFooter, ChartExInfo, ChartExLayoutProperties, ChartExLegend, ChartExNumberFormat,
    ChartExNumericDimensionType, ChartExNumericLevel, ChartExNumericPoint, ChartExOffset,
    ChartExPageMargins, ChartExPageOrientation, ChartExPageSetup, ChartExParentLabelLayout,
    ChartExPart, ChartExPlotArea, ChartExPlotAreaRegion, ChartExPlotSurface,
    ChartExPositionAlignment, ChartExPrintSettings, ChartExQuartileMethod,
    ChartExRegionLabelLayout, ChartExSeriesDataReference, ChartExSeriesLayout, ChartExSidePosition,
    ChartExSolidColor, ChartExStringDimensionType, ChartExStringLevel, ChartExStringPoint,
    ChartExText, ChartExTickLabels, ChartExTickMarkType, ChartExTickMarks,
    ChartExValueColorPositions, ChartExValueColors,
};
pub use chart_ex_style::{
    CHART_COLOR_STYLE_CONTENT_TYPE, CHART_COLOR_STYLE_RELATIONSHIP_TYPE, CHART_STYLE_CONTENT_TYPE,
    CHART_STYLE_RELATIONSHIP_TYPE, ChartColorStyleDocument, ChartColorStyleInfo,
    ChartColorStyleMethod, ChartColorStylePart, ChartStyleColor, ChartStyleColorKind,
    ChartStyleColorTransform, ChartStyleColorTransformKind, ChartStyleColorValue,
    ChartStyleDocument, ChartStyleEntry, ChartStyleEntryKind, ChartStyleFontIndex,
    ChartStyleFontReference, ChartStyleInfo, ChartStyleMarkerLayout, ChartStyleMarkerSymbol,
    ChartStylePart, ChartStylePayload, ChartStyleReference, ChartStyleVariation,
};
pub use comment::{
    Comment, CommentAuthor, CommentAuthorsPart, CommentsPart, generate_comment_authors_xml,
    generate_comments_xml,
};
pub use presentation::PresentationPart;
pub use slide::{
    MasterVisibility, SlideHeaderFooterVisibility, SlideLayoutMetadata, SlideLayoutPart,
    SlideMasterPart, SlidePart,
};
pub use theme::{Theme, ThemeColor, ThemeFont, ThemePart};
