//! Per-property lexical validation rules.

use litchi_core::Result;

use super::super::codec::validate_spec;
use super::{Kind, Value};

impl Kind {
    pub(crate) fn parse_value(self, value: &str) -> Result<Value> {
        match self {
            Self::Dr3dAmbientColor => validate_spec(value, &[], &["color"], false, self),
            Self::Dr3dBackScale => validate_spec(value, &[], &["percent"], false, self),
            Self::Dr3dBackfaceCulling => {
                validate_spec(value, &["disabled", "enabled"], &[], false, self)
            },
            Self::Dr3dCloseBack => validate_spec(value, &[], &["boolean"], false, self),
            Self::Dr3dCloseFront => validate_spec(value, &[], &["boolean"], false, self),
            Self::Dr3dDepth => validate_spec(value, &[], &["length"], false, self),
            Self::Dr3dDiffuseColor => validate_spec(value, &[], &["color"], false, self),
            Self::Dr3dEdgeRounding => validate_spec(value, &[], &["percent"], false, self),
            Self::Dr3dEdgeRoundingMode => {
                validate_spec(value, &["attractive", "correct"], &[], false, self)
            },
            Self::Dr3dEmissiveColor => validate_spec(value, &[], &["color"], false, self),
            Self::Dr3dEndAngle => validate_spec(value, &[], &["angle"], false, self),
            Self::Dr3dHorizontalSegments => {
                validate_spec(value, &[], &["nonNegativeInteger"], false, self)
            },
            Self::Dr3dLightingMode => {
                validate_spec(value, &["double-sided", "standard"], &[], false, self)
            },
            Self::Dr3dNormalsDirection => {
                validate_spec(value, &["inverse", "normal"], &[], false, self)
            },
            Self::Dr3dNormalsKind => {
                validate_spec(value, &["flat", "object", "sphere"], &[], false, self)
            },
            Self::Dr3dShadow => validate_spec(value, &["hidden", "visible"], &[], false, self),
            Self::Dr3dShininess => validate_spec(value, &[], &["percent"], false, self),
            Self::Dr3dSpecularColor => validate_spec(value, &[], &["color"], false, self),
            Self::Dr3dTextureFilter => {
                validate_spec(value, &["disabled", "enabled"], &[], false, self)
            },
            Self::Dr3dTextureGenerationModeX => {
                validate_spec(value, &["object", "parallel", "sphere"], &[], false, self)
            },
            Self::Dr3dTextureGenerationModeY => {
                validate_spec(value, &["object", "parallel", "sphere"], &[], false, self)
            },
            Self::Dr3dTextureKind => validate_spec(
                value,
                &["color", "intensity", "luminance"],
                &[],
                false,
                self,
            ),
            Self::Dr3dTextureMode => {
                validate_spec(value, &["blend", "modulate", "replace"], &[], false, self)
            },
            Self::Dr3dVerticalSegments => {
                validate_spec(value, &[], &["nonNegativeInteger"], false, self)
            },
            Self::DrawAutoGrowHeight => validate_spec(value, &[], &["boolean"], false, self),
            Self::DrawAutoGrowWidth => validate_spec(value, &[], &["boolean"], false, self),
            Self::DrawBlue => {
                validate_spec(value, &[], &["signedZeroToHundredPercent"], false, self)
            },
            Self::DrawCaptionAngle => validate_spec(value, &[], &["angle"], false, self),
            Self::DrawCaptionAngleType => {
                validate_spec(value, &["fixed", "free"], &[], false, self)
            },
            Self::DrawCaptionEscape => {
                validate_spec(value, &[], &["length", "percent"], false, self)
            },
            Self::DrawCaptionEscapeDirection => {
                validate_spec(value, &["auto", "horizontal", "vertical"], &[], false, self)
            },
            Self::DrawCaptionFitLineLength => validate_spec(value, &[], &["boolean"], false, self),
            Self::DrawCaptionGap => validate_spec(value, &[], &["distance"], false, self),
            Self::DrawCaptionLineLength => validate_spec(value, &[], &["length"], false, self),
            Self::DrawCaptionType => validate_spec(
                value,
                &["angled-connector-line", "angled-line", "straight-line"],
                &[],
                false,
                self,
            ),
            Self::DrawColorInversion => validate_spec(value, &[], &["boolean"], false, self),
            Self::DrawColorMode => validate_spec(
                value,
                &["greyscale", "mono", "standard", "watermark"],
                &[],
                false,
                self,
            ),
            Self::DrawContrast => validate_spec(value, &[], &["percent"], false, self),
            Self::DrawDecimalPlaces => {
                validate_spec(value, &[], &["nonNegativeInteger"], false, self)
            },
            Self::DrawDrawAspect => validate_spec(
                value,
                &["content", "icon", "print-view", "thumbnail"],
                &[],
                false,
                self,
            ),
            Self::DrawEndGuide => validate_spec(value, &[], &["length"], false, self),
            Self::DrawEndLineSpacingHorizontal => {
                validate_spec(value, &[], &["distance"], false, self)
            },
            Self::DrawEndLineSpacingVertical => {
                validate_spec(value, &[], &["distance"], false, self)
            },
            Self::DrawFill => validate_spec(
                value,
                &["bitmap", "gradient", "hatch", "none", "solid"],
                &[],
                false,
                self,
            ),
            Self::DrawFillColor => validate_spec(value, &[], &["color"], false, self),
            Self::DrawFillGradientName => validate_spec(value, &[], &["styleNameRef"], false, self),
            Self::DrawFillHatchName => validate_spec(value, &[], &["styleNameRef"], false, self),
            Self::DrawFillHatchSolid => validate_spec(value, &[], &["boolean"], false, self),
            Self::DrawFillImageHeight => {
                validate_spec(value, &[], &["length", "percent"], false, self)
            },
            Self::DrawFillImageName => validate_spec(value, &[], &["styleNameRef"], false, self),
            Self::DrawFillImageRefPoint => validate_spec(
                value,
                &[
                    "bottom",
                    "bottom-left",
                    "bottom-right",
                    "center",
                    "left",
                    "right",
                    "top",
                    "top-left",
                    "top-right",
                ],
                &[],
                false,
                self,
            ),
            Self::DrawFillImageRefPointX => validate_spec(value, &[], &["percent"], false, self),
            Self::DrawFillImageRefPointY => validate_spec(value, &[], &["percent"], false, self),
            Self::DrawFillImageWidth => {
                validate_spec(value, &[], &["length", "percent"], false, self)
            },
            Self::DrawFitToContour => validate_spec(value, &[], &["boolean"], false, self),
            Self::DrawFitToSize => validate_spec(value, &[], &["boolean"], false, self),
            Self::DrawFrameDisplayBorder => validate_spec(value, &[], &["boolean"], false, self),
            Self::DrawFrameDisplayScrollbar => validate_spec(value, &[], &["boolean"], false, self),
            Self::DrawFrameMarginHorizontal => {
                validate_spec(value, &[], &["nonNegativePixelLength"], false, self)
            },
            Self::DrawFrameMarginVertical => {
                validate_spec(value, &[], &["nonNegativePixelLength"], false, self)
            },
            Self::DrawGamma => validate_spec(value, &[], &["percent"], false, self),
            Self::DrawGradientStepCount => {
                validate_spec(value, &[], &["nonNegativeInteger"], false, self)
            },
            Self::DrawGreen => {
                validate_spec(value, &[], &["signedZeroToHundredPercent"], false, self)
            },
            Self::DrawGuideDistance => validate_spec(value, &[], &["distance"], false, self),
            Self::DrawGuideOverhang => validate_spec(value, &[], &["length"], false, self),
            Self::DrawImageOpacity => {
                validate_spec(value, &[], &["zeroToHundredPercent"], false, self)
            },
            Self::DrawLineDistance => validate_spec(value, &[], &["distance"], false, self),
            Self::DrawLuminance => {
                validate_spec(value, &[], &["zeroToHundredPercent"], false, self)
            },
            Self::DrawMarkerEnd => validate_spec(value, &[], &["styleNameRef"], false, self),
            Self::DrawMarkerEndCenter => validate_spec(value, &[], &["boolean"], false, self),
            Self::DrawMarkerEndWidth => validate_spec(value, &[], &["length"], false, self),
            Self::DrawMarkerStart => validate_spec(value, &[], &["styleNameRef"], false, self),
            Self::DrawMarkerStartCenter => validate_spec(value, &[], &["boolean"], false, self),
            Self::DrawMarkerStartWidth => validate_spec(value, &[], &["length"], false, self),
            Self::DrawMeasureAlign => validate_spec(
                value,
                &["automatic", "inside", "left-outside", "right-outside"],
                &[],
                false,
                self,
            ),
            Self::DrawMeasureVerticalAlign => validate_spec(
                value,
                &["above", "automatic", "below", "center"],
                &[],
                false,
                self,
            ),
            Self::DrawOleDrawAspect => {
                validate_spec(value, &[], &["nonNegativeInteger"], false, self)
            },
            Self::DrawOpacity => validate_spec(value, &[], &["percent"], false, self),
            Self::DrawOpacityName => validate_spec(value, &[], &["styleNameRef"], false, self),
            Self::DrawParallel => validate_spec(value, &[], &["boolean"], false, self),
            Self::DrawPlacing => validate_spec(value, &["above", "below"], &[], false, self),
            Self::DrawRed => {
                validate_spec(value, &[], &["signedZeroToHundredPercent"], false, self)
            },
            Self::DrawSecondaryFillColor => validate_spec(value, &[], &["color"], false, self),
            Self::DrawShadow => validate_spec(value, &["hidden", "visible"], &[], false, self),
            Self::DrawShadowColor => validate_spec(value, &[], &["color"], false, self),
            Self::DrawShadowOffsetX => validate_spec(value, &[], &["length"], false, self),
            Self::DrawShadowOffsetY => validate_spec(value, &[], &["length"], false, self),
            Self::DrawShadowOpacity => {
                validate_spec(value, &[], &["zeroToHundredPercent"], false, self)
            },
            Self::DrawShowUnit => validate_spec(value, &[], &["boolean"], false, self),
            Self::DrawStartGuide => validate_spec(value, &[], &["length"], false, self),
            Self::DrawStartLineSpacingHorizontal => {
                validate_spec(value, &[], &["distance"], false, self)
            },
            Self::DrawStartLineSpacingVertical => {
                validate_spec(value, &[], &["distance"], false, self)
            },
            Self::DrawStroke => validate_spec(value, &["dash", "none", "solid"], &[], false, self),
            Self::DrawStrokeDash => validate_spec(value, &[], &["styleNameRef"], false, self),
            Self::DrawStrokeDashNames => validate_spec(value, &[], &["styleNameRefs"], false, self),
            Self::DrawStrokeLinejoin => validate_spec(
                value,
                &["bevel", "middle", "miter", "none", "round"],
                &[],
                false,
                self,
            ),
            Self::DrawSymbolColor => validate_spec(value, &[], &["color"], false, self),
            Self::DrawTextareaHorizontalAlign => validate_spec(
                value,
                &["center", "justify", "left", "right"],
                &[],
                false,
                self,
            ),
            Self::DrawTextareaVerticalAlign => validate_spec(
                value,
                &["bottom", "justify", "middle", "top"],
                &[],
                false,
                self,
            ),
            Self::DrawTileRepeatOffset => validate_spec(
                value,
                &["horizontal", "vertical"],
                &["zeroToHundredPercent"],
                true,
                self,
            ),
            Self::DrawUnit => validate_spec(
                value,
                &[
                    "automatic",
                    "cm",
                    "ft",
                    "inch",
                    "km",
                    "m",
                    "mi",
                    "mm",
                    "pc",
                    "pt",
                ],
                &[],
                false,
                self,
            ),
            Self::DrawVisibleAreaHeight => {
                validate_spec(value, &[], &["positiveLength"], false, self)
            },
            Self::DrawVisibleAreaLeft => {
                validate_spec(value, &[], &["nonNegativeLength"], false, self)
            },
            Self::DrawVisibleAreaTop => {
                validate_spec(value, &[], &["nonNegativeLength"], false, self)
            },
            Self::DrawVisibleAreaWidth => {
                validate_spec(value, &[], &["positiveLength"], false, self)
            },
            Self::DrawWrapInfluenceOnPosition => validate_spec(
                value,
                &["iterative", "once-concurrent", "once-successive"],
                &[],
                false,
                self,
            ),
            Self::FoBackgroundColor => {
                validate_spec(value, &["transparent"], &["color"], false, self)
            },
            Self::FoBorder => validate_spec(value, &[], &["string"], false, self),
            Self::FoBorderBottom => validate_spec(value, &[], &["string"], false, self),
            Self::FoBorderLeft => validate_spec(value, &[], &["string"], false, self),
            Self::FoBorderRight => validate_spec(value, &[], &["string"], false, self),
            Self::FoBorderTop => validate_spec(value, &[], &["string"], false, self),
            Self::FoClip => validate_spec(value, &["auto"], &["clipShape"], false, self),
            Self::FoMargin => {
                validate_spec(value, &[], &["nonNegativeLength", "percent"], false, self)
            },
            Self::FoMarginBottom => {
                validate_spec(value, &[], &["nonNegativeLength", "percent"], false, self)
            },
            Self::FoMarginLeft => validate_spec(value, &[], &["length", "percent"], false, self),
            Self::FoMarginRight => validate_spec(value, &[], &["length", "percent"], false, self),
            Self::FoMarginTop => {
                validate_spec(value, &[], &["nonNegativeLength", "percent"], false, self)
            },
            Self::FoMaxHeight => validate_spec(value, &[], &["length", "percent"], false, self),
            Self::FoMaxWidth => validate_spec(value, &[], &["length", "percent"], false, self),
            Self::FoMinHeight => validate_spec(value, &[], &["length", "percent"], false, self),
            Self::FoMinWidth => validate_spec(value, &[], &["length", "percent"], false, self),
            Self::FoPadding => validate_spec(value, &[], &["nonNegativeLength"], false, self),
            Self::FoPaddingBottom => validate_spec(value, &[], &["nonNegativeLength"], false, self),
            Self::FoPaddingLeft => validate_spec(value, &[], &["nonNegativeLength"], false, self),
            Self::FoPaddingRight => validate_spec(value, &[], &["nonNegativeLength"], false, self),
            Self::FoPaddingTop => validate_spec(value, &[], &["nonNegativeLength"], false, self),
            Self::FoWrapOption => validate_spec(value, &["no-wrap", "wrap"], &[], false, self),
            Self::StyleBackgroundTransparency => {
                validate_spec(value, &[], &["zeroToHundredPercent"], false, self)
            },
            Self::StyleBorderLineWidth => validate_spec(value, &[], &["borderWidths"], false, self),
            Self::StyleBorderLineWidthBottom => {
                validate_spec(value, &[], &["borderWidths"], false, self)
            },
            Self::StyleBorderLineWidthLeft => {
                validate_spec(value, &[], &["borderWidths"], false, self)
            },
            Self::StyleBorderLineWidthRight => {
                validate_spec(value, &[], &["borderWidths"], false, self)
            },
            Self::StyleBorderLineWidthTop => {
                validate_spec(value, &[], &["borderWidths"], false, self)
            },
            Self::StyleEditable => validate_spec(value, &[], &["boolean"], false, self),
            Self::StyleFlowWithText => validate_spec(value, &[], &["boolean"], false, self),
            Self::StyleHorizontalPos => validate_spec(
                value,
                &[
                    "center",
                    "from-inside",
                    "from-left",
                    "inside",
                    "left",
                    "outside",
                    "right",
                ],
                &[],
                false,
                self,
            ),
            Self::StyleHorizontalRel => validate_spec(
                value,
                &[
                    "char",
                    "frame",
                    "frame-content",
                    "frame-end-margin",
                    "frame-start-margin",
                    "page",
                    "page-content",
                    "page-end-margin",
                    "page-start-margin",
                    "paragraph",
                    "paragraph-content",
                    "paragraph-end-margin",
                    "paragraph-start-margin",
                ],
                &[],
                false,
                self,
            ),
            Self::StyleMirror => validate_spec(
                value,
                &["none", "vertical"],
                &["horizontal-mirror"],
                true,
                self,
            ),
            Self::StyleNumberWrappedParagraphs => {
                validate_spec(value, &["no-limit"], &["positiveInteger"], false, self)
            },
            Self::StyleOverflowBehavior => {
                validate_spec(value, &["auto-create-new-frame", "clip"], &[], false, self)
            },
            Self::StylePrintContent => validate_spec(value, &[], &["boolean"], false, self),
            Self::StyleProtect => validate_spec(
                value,
                &["content", "none", "position", "size"],
                &[],
                true,
                self,
            ),
            Self::StyleRelHeight => {
                validate_spec(value, &["scale", "scale-min"], &["percent"], false, self)
            },
            Self::StyleRelWidth => {
                validate_spec(value, &["scale", "scale-min"], &["percent"], false, self)
            },
            Self::StyleRepeat => {
                validate_spec(value, &["no-repeat", "repeat", "stretch"], &[], false, self)
            },
            Self::StyleRunThrough => {
                validate_spec(value, &["background", "foreground"], &[], false, self)
            },
            Self::StyleShadow => validate_spec(value, &[], &["shadowType"], false, self),
            Self::StyleShrinkToFit => validate_spec(value, &[], &["boolean"], false, self),
            Self::StyleVerticalPos => validate_spec(
                value,
                &["below", "bottom", "from-top", "middle", "top"],
                &[],
                false,
                self,
            ),
            Self::StyleVerticalRel => validate_spec(
                value,
                &[
                    "baseline",
                    "char",
                    "frame",
                    "frame-content",
                    "line",
                    "page",
                    "page-content",
                    "paragraph",
                    "paragraph-content",
                    "text",
                ],
                &[],
                false,
                self,
            ),
            Self::StyleWrap => validate_spec(
                value,
                &[
                    "biggest",
                    "dynamic",
                    "left",
                    "none",
                    "parallel",
                    "right",
                    "run-through",
                ],
                &[],
                false,
                self,
            ),
            Self::StyleWrapContour => validate_spec(value, &[], &["boolean"], false, self),
            Self::StyleWrapContourMode => {
                validate_spec(value, &["full", "outside"], &[], false, self)
            },
            Self::StyleWrapDynamicThreshold => {
                validate_spec(value, &[], &["nonNegativeLength"], false, self)
            },
            Self::StyleWritingMode => validate_spec(
                value,
                &["lr", "lr-tb", "page", "rl", "rl-tb", "tb", "tb-lr", "tb-rl"],
                &[],
                false,
                self,
            ),
            Self::SvgFillRule => validate_spec(value, &["evenodd", "nonzero"], &[], false, self),
            Self::SvgHeight => validate_spec(value, &[], &["length"], false, self),
            Self::SvgStrokeColor => validate_spec(value, &[], &["color"], false, self),
            Self::SvgStrokeLinecap => {
                validate_spec(value, &["butt", "round", "square"], &[], false, self)
            },
            Self::SvgStrokeOpacity => {
                validate_spec(value, &[], &["zeroToHundredPercent"], false, self)
            },
            Self::SvgStrokeWidth => validate_spec(value, &[], &["length"], false, self),
            Self::SvgWidth => validate_spec(value, &[], &["length"], false, self),
            Self::SvgX => validate_spec(value, &[], &["coordinate"], false, self),
            Self::SvgY => validate_spec(value, &[], &["coordinate"], false, self),
            Self::TextAnchorPageNumber => {
                validate_spec(value, &[], &["positiveInteger"], false, self)
            },
            Self::TextAnchorType => validate_spec(
                value,
                &["as-char", "char", "frame", "page", "paragraph"],
                &[],
                false,
                self,
            ),
            Self::TextAnimation => validate_spec(
                value,
                &["alternate", "none", "scroll", "slide"],
                &[],
                false,
                self,
            ),
            Self::TextAnimationDelay => validate_spec(value, &[], &["duration"], false, self),
            Self::TextAnimationDirection => {
                validate_spec(value, &["down", "left", "right", "up"], &[], false, self)
            },
            Self::TextAnimationRepeat => {
                validate_spec(value, &[], &["nonNegativeInteger"], false, self)
            },
            Self::TextAnimationStartInside => validate_spec(value, &[], &["boolean"], false, self),
            Self::TextAnimationSteps => validate_spec(value, &[], &["length"], false, self),
            Self::TextAnimationStopInside => validate_spec(value, &[], &["boolean"], false, self),
        }
    }
}
