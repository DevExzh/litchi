use super::ControlWord;
use crate::codec::error::{RtfError, RtfResult};

/// Match a control-word spelling and optional parameter to its typed token.
///
/// This is the RTF specification's flat control-word dispatch table. Keeping
/// the entries together makes duplicate and missing mappings auditable.
#[allow(clippy::too_many_lines)]
pub(in crate::codec::lexer::codec) fn match_control_word<'a>(
    word: &'a str,
    param: Option<i32>,
) -> RtfResult<ControlWord<'a>> {
    let param_value = param.unwrap_or(1);
    let param_bool = param.unwrap_or(1) != 0;

    #[allow(clippy::match_same_arms)]
    let control = match word {
        // Document
        "rtf" => ControlWord::Rtf(param_value),
        "ansi" => ControlWord::Ansi,
        "ansicpg" => ControlWord::AnsiCodePage(param_value),
        "mac" => ControlWord::Mac,
        "pc" => ControlWord::Pc,
        "pca" => ControlWord::Pca,

        // Headers
        "fonttbl" => ControlWord::FontTable,
        "colortbl" => ControlWord::ColorTable,
        "stylesheet" => ControlWord::StyleSheet,
        "info" => ControlWord::Info,
        "defchp" => ControlWord::DefaultCharacterProperties(param),
        "defpap" => ControlWord::DefaultParagraphProperties(param),
        "deff" => ControlWord::DefaultFont(param),
        "adeff" => ControlWord::AssociatedDefaultFont(param),
        "stshfbi" => ControlWord::StylesheetDefaultBidiFont(param),
        "stshfdbch" => ControlWord::StylesheetDefaultDoubleByteFont(param),
        "stshfhich" => ControlWord::StylesheetDefaultHighAnsiFont(param),
        "stshfloch" => ControlWord::StylesheetDefaultLowAnsiFont(param),
        "pn" => ControlWord::LegacyParagraphNumbering(param),
        "pnlvl" => ControlWord::LegacyNumberingLevel(param),
        "pnlvlblt" => ControlWord::LegacyNumberingLevelBullet(param),
        "pnlvlbody" => ControlWord::LegacyNumberingLevelBody(param),
        "pnlvlcont" => ControlWord::LegacyNumberingLevelContinue(param),
        "pnseclvl" => ControlWord::LegacySectionNumberingLevel(param_value),
        "pndec" => ControlWord::LegacyNumberingDecimal(param),
        "pnucrm" => ControlWord::LegacyNumberingUpperRoman(param),
        "pnlcrm" => ControlWord::LegacyNumberingLowerRoman(param),
        "pnucltr" => ControlWord::LegacyNumberingUpperLetter(param),
        "pnlcltr" => ControlWord::LegacyNumberingLowerLetter(param),
        "pnaiu" => {
            ControlWord::LegacyNumberingFormat(crate::LegacyParagraphNumberingFormat::Aiueo, param)
        },
        "pnaiud" => ControlWord::LegacyNumberingFormat(
            crate::LegacyParagraphNumberingFormat::AiueoDbChar,
            param,
        ),
        "pnaiueo" => ControlWord::LegacyNumberingFormat(
            crate::LegacyParagraphNumberingFormat::AiueoExtended,
            param,
        ),
        "pnaiueod" => ControlWord::LegacyNumberingFormat(
            crate::LegacyParagraphNumberingFormat::AiueoExtendedDbChar,
            param,
        ),
        "pnchosung" => ControlWord::LegacyNumberingFormat(
            crate::LegacyParagraphNumberingFormat::Chosung,
            param,
        ),
        "pncard" => ControlWord::LegacyNumberingFormat(
            crate::LegacyParagraphNumberingFormat::CardinalText,
            param,
        ),
        "pndecd" => ControlWord::LegacyNumberingFormat(
            crate::LegacyParagraphNumberingFormat::DecimalWithPeriod,
            param,
        ),
        "pnord" => ControlWord::LegacyNumberingFormat(
            crate::LegacyParagraphNumberingFormat::Ordinal,
            param,
        ),
        "pnordt" => ControlWord::LegacyNumberingFormat(
            crate::LegacyParagraphNumberingFormat::OrdinalText,
            param,
        ),
        "pncnum" => ControlWord::LegacyNumberingFormat(
            crate::LegacyParagraphNumberingFormat::ChineseCounting,
            param,
        ),
        "pndbnum" => ControlWord::LegacyNumberingFormat(
            crate::LegacyParagraphNumberingFormat::ChineseCountingDbChar,
            param,
        ),
        "pndbnumd" => ControlWord::LegacyNumberingFormat(
            crate::LegacyParagraphNumberingFormat::ChineseCountingKorean,
            param,
        ),
        "pndbnumk" => ControlWord::LegacyNumberingFormat(
            crate::LegacyParagraphNumberingFormat::ChineseCountingLegal,
            param,
        ),
        "pndbnuml" => ControlWord::LegacyNumberingFormat(
            crate::LegacyParagraphNumberingFormat::ChineseCountingThousand,
            param,
        ),
        "pndbnumt" => ControlWord::LegacyNumberingFormat(
            crate::LegacyParagraphNumberingFormat::ChineseCountingTraditional,
            param,
        ),
        "pnganada" => {
            ControlWord::LegacyNumberingFormat(crate::LegacyParagraphNumberingFormat::Ganada, param)
        },
        "pngbnum" => ControlWord::LegacyNumberingFormat(
            crate::LegacyParagraphNumberingFormat::GbCounting,
            param,
        ),
        "pngbnumd" => ControlWord::LegacyNumberingFormat(
            crate::LegacyParagraphNumberingFormat::GbCountingDbChar,
            param,
        ),
        "pngbnumk" => ControlWord::LegacyNumberingFormat(
            crate::LegacyParagraphNumberingFormat::GbCountingKorean,
            param,
        ),
        "pngbnuml" => ControlWord::LegacyNumberingFormat(
            crate::LegacyParagraphNumberingFormat::GbCountingLegal,
            param,
        ),
        "pniroha" => {
            ControlWord::LegacyNumberingFormat(crate::LegacyParagraphNumberingFormat::Iroha, param)
        },
        "pnirohad" => ControlWord::LegacyNumberingFormat(
            crate::LegacyParagraphNumberingFormat::IrohaDbChar,
            param,
        ),
        "pnzodiac" => {
            ControlWord::LegacyNumberingFormat(crate::LegacyParagraphNumberingFormat::Zodiac, param)
        },
        "pnzodiacd" => ControlWord::LegacyNumberingFormat(
            crate::LegacyParagraphNumberingFormat::ZodiacDbChar,
            param,
        ),
        "pnzodiacl" => ControlWord::LegacyNumberingFormat(
            crate::LegacyParagraphNumberingFormat::ZodiacLegal,
            param,
        ),
        "pnstart" => ControlWord::LegacyNumberingStart(param),
        "pnindent" => ControlWord::LegacyNumberingIndent(param),
        "pnsp" => ControlWord::LegacyNumberingSpace(param),
        "pnhang" => ControlWord::LegacyNumberingHanging(param),
        "pnprev" => ControlWord::LegacyNumberingPrevious(param),
        "pnacross" => ControlWord::LegacyNumberingAcross(param),
        "pnnumonce" => ControlWord::LegacyNumberingOnce(param),
        "pnrestart" => ControlWord::LegacyNumberingRestart(param),
        "pnbidia" => ControlWord::LegacyNumberingBidiA(param),
        "pnbidib" => ControlWord::LegacyNumberingBidiB(param),
        "pnql" => ControlWord::LegacyNumberingAlignLeft(param),
        "pnqc" => ControlWord::LegacyNumberingAlignCenter(param),
        "pnqr" => ControlWord::LegacyNumberingAlignRight(param),
        "pnf" => ControlWord::LegacyNumberingFont(param),
        "pnfs" => ControlWord::LegacyNumberingFontSize(param),
        "pncf" => ControlWord::LegacyNumberingColor(param),
        "pnb" => ControlWord::LegacyNumberingBold(param),
        "pni" => ControlWord::LegacyNumberingItalic(param),
        "pncaps" => ControlWord::LegacyNumberingCaps(param),
        "pnscaps" => ControlWord::LegacyNumberingSmallCaps(param),
        "pnstrike" => ControlWord::LegacyNumberingStrike(param),
        "pnul" => ControlWord::LegacyNumberingUnderlineToggle(param),
        "pnuld" => ControlWord::LegacyNumberingUnderline(
            crate::LegacyParagraphNumberingUnderline::Dotted,
            param,
        ),
        "pnuldash" => ControlWord::LegacyNumberingUnderline(
            crate::LegacyParagraphNumberingUnderline::Dashed,
            param,
        ),
        "pnuldashd" => ControlWord::LegacyNumberingUnderline(
            crate::LegacyParagraphNumberingUnderline::DashDot,
            param,
        ),
        "pnuldashdd" => ControlWord::LegacyNumberingUnderline(
            crate::LegacyParagraphNumberingUnderline::DashDotDot,
            param,
        ),
        "pnuldb" => ControlWord::LegacyNumberingUnderline(
            crate::LegacyParagraphNumberingUnderline::Double,
            param,
        ),
        "pnulhair" => ControlWord::LegacyNumberingUnderline(
            crate::LegacyParagraphNumberingUnderline::Hairline,
            param,
        ),
        "pnulnone" => ControlWord::LegacyNumberingUnderline(
            crate::LegacyParagraphNumberingUnderline::None,
            param,
        ),
        "pnulth" => ControlWord::LegacyNumberingUnderline(
            crate::LegacyParagraphNumberingUnderline::Thick,
            param,
        ),
        "pnulw" => ControlWord::LegacyNumberingUnderline(
            crate::LegacyParagraphNumberingUnderline::Words,
            param,
        ),
        "pnulwave" => ControlWord::LegacyNumberingUnderline(
            crate::LegacyParagraphNumberingUnderline::Wave,
            param,
        ),
        "pnrauth" => ControlWord::LegacyNumberingRevisionAuthor(param),
        "pnrdate" => ControlWord::LegacyNumberingRevisionDate(param),
        "pnrnfc" => ControlWord::LegacyNumberingRevisionFormat(param),
        "pnrnot" => ControlWord::LegacyNumberingRevisionNoTrack(param),
        "pnrpnbr" => ControlWord::LegacyNumberingRevisionParagraph(param),
        "pnrrgb" => ControlWord::LegacyNumberingRevisionRgb(param),
        "pnrstart" => ControlWord::LegacyNumberingRevisionStart(param),
        "pnrstop" => ControlWord::LegacyNumberingRevisionStop(param),
        "pnrxst" => ControlWord::LegacyNumberingRevisionTextStart(param),
        "pntxtb" => ControlWord::LegacyNumberingTextBefore,
        "pntxta" => ControlWord::LegacyNumberingTextAfter,
        "pntext" => ControlWord::LegacyGeneratedListText,
        "pgptbl" => ControlWord::ParagraphGroupTable,
        "pgp" => ControlWord::ParagraphGroup,
        "ipgp" => ControlWord::ParagraphGroupParent(param_value),
        "itap" => ControlWord::TableNestingLevel(param),

        // Stylesheet entries and metadata
        "s" => ControlWord::ParagraphStyle(param),
        "cs" => ControlWord::CharacterStyle(param),
        "ds" => ControlWord::SectionStyle(param),
        "ts" => ControlWord::TableStyle(param),
        "sbasedon" => ControlWord::StyleBasedOn(param_value),
        "snext" => ControlWord::StyleNext(param_value),
        "slink" => ControlWord::StyleLink(param_value),
        "additive" => ControlWord::StyleAdditive(param_bool),
        "sautoupd" => ControlWord::StyleAutoUpdate(param_bool),
        "shidden" => ControlWord::StyleHidden(param_bool),
        "slocked" => ControlWord::StyleLocked(param_bool),
        "ssemihidden" => ControlWord::StyleSemiHidden(param_bool),
        "sunhideused" => ControlWord::StyleUnhideWhenUsed(param_bool),
        "sqformat" => ControlWord::StyleQuickFormat(param_bool),
        "spriority" => ControlWord::StylePriority(param_value),
        "styrsid" => ControlWord::StyleRevisionId(param_value),
        "spersonal" => ControlWord::StylePersonal(param_bool),
        "scompose" => ControlWord::StyleCompose(param_bool),
        "sreply" => ControlWord::StyleReply(param_bool),

        // Embedded content
        "pict" => ControlWord::Picture,
        "picprop" => ControlWord::PictureProperties(param),
        "shplid" => ControlWord::PictureShapeId(param),
        "object" => {
            if param.is_some() {
                ControlWord::InvalidObjectDestinationParameter
            } else {
                ControlWord::Object
            }
        },
        "result" => {
            if param.is_some() {
                ControlWord::InvalidObjectResultDestinationParameter
            } else {
                ControlWord::Result
            }
        },
        "objclass" => ControlWord::ObjectClass,
        "docvar" => ControlWord::DocumentVariable,
        "userprops" => ControlWord::UserProperties,
        "propname" => ControlWord::PropertyName,
        "proptype" => ControlWord::PropertyType(param),
        "staticval" => ControlWord::StaticValue,
        "linkval" => ControlWord::LinkValue,
        "objname" => ControlWord::ObjectName,
        "objdata" => ControlWord::ObjectData,
        "objemb" => {
            if param.is_none() {
                ControlWord::ObjectEmbedded
            } else {
                ControlWord::InvalidObjectModifierParameter
            }
        },
        "objlink" => {
            if param.is_none() {
                ControlWord::ObjectLink
            } else {
                ControlWord::InvalidObjectModifierParameter
            }
        },
        "objautlink" => {
            if param.is_none() {
                ControlWord::ObjectAutoLink
            } else {
                ControlWord::InvalidObjectModifierParameter
            }
        },
        "objhtml" => {
            if param.is_none() {
                ControlWord::ObjectHtml
            } else {
                ControlWord::InvalidObjectModifierParameter
            }
        },
        "objsub" => ControlWord::ObjectSubscriber(param),
        "objpub" => ControlWord::ObjectPublisher(param),
        "objicemb" => ControlWord::ObjectInstallableCommand(param),
        "objocx" => ControlWord::ObjectOleControl(param),
        "linkself" => ControlWord::ObjectLinkSelf(param),
        "objw" => param.map_or(
            ControlWord::InvalidObjectModifierParameter,
            ControlWord::ObjectWidth,
        ),
        "objh" => param.map_or(
            ControlWord::InvalidObjectModifierParameter,
            ControlWord::ObjectHeight,
        ),
        "objalign" => ControlWord::ObjectAlignment(param),
        "objtransy" => ControlWord::ObjectTranslationY(param),
        "objcropt" => ControlWord::ObjectCropTop(param),
        "objcropb" => ControlWord::ObjectCropBottom(param),
        "objcropl" => ControlWord::ObjectCropLeft(param),
        "objcropr" => ControlWord::ObjectCropRight(param),
        "objscalex" => ControlWord::ObjectScaleX(param),
        "objscaley" => ControlWord::ObjectScaleY(param),
        "objlock" => {
            if param.is_none() {
                ControlWord::ObjectLocked(true)
            } else {
                ControlWord::InvalidObjectModifierParameter
            }
        },
        "objupdate" => {
            if param.is_none() {
                ControlWord::ObjectUpdate(true)
            } else {
                ControlWord::InvalidObjectModifierParameter
            }
        },
        "objsetsize" => {
            if param.is_none() {
                ControlWord::ObjectSetSize(true)
            } else {
                ControlWord::InvalidObjectModifierParameter
            }
        },
        "rsltmerge" => ControlWord::ObjectResultMerge(param),
        "rsltrtf" => ControlWord::ObjectResultRtf(param),
        "rslttxt" => ControlWord::ObjectResultText(param),
        "rsltpict" => ControlWord::ObjectResultPicture(param),
        "rsltbmp" => ControlWord::ObjectResultBitmap(param),
        "rslthtml" => ControlWord::ObjectResultHtml(param),
        "oleclsid" => ControlWord::OleClassId(param),

        // Picture properties
        "picw" => ControlWord::PictureWidth(param_value),
        "pich" => ControlWord::PictureHeight(param_value),
        "picwgoal" => ControlWord::PictureGoalWidth(param_value),
        "pichgoal" => ControlWord::PictureGoalHeight(param_value),
        "picscalex" => ControlWord::PictureScaleX(param_value),
        "picscaley" => ControlWord::PictureScaleY(param_value),
        "picscaled" => ControlWord::PictureScaled(param),
        "picbmp" => ControlWord::PictureBitmap(param),
        "picbpp" => ControlWord::PictureBitsPerPixel(param),
        "piccropl" => ControlWord::PictureCropLeft(param),
        "piccropr" => ControlWord::PictureCropRight(param),
        "piccropt" => ControlWord::PictureCropTop(param),
        "piccropb" => ControlWord::PictureCropBottom(param),
        "wbmbitspixel" => ControlWord::WindowsBitmapBitsPerPixel(param),
        "wbmplanes" => ControlWord::WindowsBitmapPlanes(param),
        "wbmwidthbytes" => ControlWord::WindowsBitmapWidthBytes(param),
        "emfblip" => ControlWord::Emfblip,
        "pngblip" => ControlWord::Pngblip,
        "jpegblip" => ControlWord::Jpegblip,
        "macpict" => ControlWord::Macpict,
        "pmmetafile" => ControlWord::Pmmetafile(param_value),
        "wmetafile" => ControlWord::Wmetafile(param_value),
        "dibitmap" => ControlWord::Dibitmap(param_value),
        "wbitmap" => ControlWord::Wbitmap(param_value),
        "bliptag" => ControlWord::BlipTag(param_value),
        "blipuid" => ControlWord::BlipUid,
        "blipupi" => ControlWord::BlipUnitsPerInch(param_value),

        // Field support
        "field" => ControlWord::Field,
        "fldinst" => ControlWord::FieldInstruction,
        "fldrslt" => ControlWord::FieldResult,
        "fldlock" => ControlWord::FieldLock(param),
        "flddirty" => ControlWord::FieldDirty(param),
        "fldedit" => ControlWord::FieldEdit(param),
        "fldpriv" => ControlWord::FieldPrivate(param),
        "formfield" => ControlWord::FormField,
        "datafield" => ControlWord::DataField,
        "fftype" => ControlWord::FormFieldType(param_value),
        "ffmaxlen" => ControlWord::FormFieldMaxLength(param_value),
        "ffprot" => ControlWord::FormFieldProtected(param_bool),
        "ffrecalc" => ControlWord::FormFieldRecalculate(param_bool),
        "ffsize" => ControlWord::FormFieldAutomaticSize(param_bool),
        "ffname" => ControlWord::FormFieldName,
        "ffformat" => ControlWord::FormFieldFormat,
        "ffdeftext" => ControlWord::FormFieldDefaultText,
        "ffdefres" => ControlWord::FormFieldDefaultResult(param_value),
        "ffres" => ControlWord::FormFieldResult(param_value),
        "ffhps" => ControlWord::FormFieldHalfPointSize(param_value),
        "ffownhelp" => ControlWord::FormFieldOwnHelp(param_bool),
        "ffownstat" => ControlWord::FormFieldOwnStatus(param_bool),
        "ffhelptext" => ControlWord::FormFieldHelpText,
        "ffstattext" => ControlWord::FormFieldStatusText,
        "ffentrymcr" => ControlWord::FormFieldEntryMacro,
        "ffexitmcr" => ControlWord::FormFieldExitMacro,
        "ffl" => ControlWord::FormFieldListEntry,
        "fftypetxt" => ControlWord::FormFieldTextType(param_value),
        "ffhaslistbox" => ControlWord::FormFieldHasListBox(param_bool),
        "generator" => ControlWord::Generator,
        "rsidtbl" => ControlWord::RevisionSaveTable,
        "rsid" => ControlWord::RevisionSaveId(param_value),
        "rsidroot" => ControlWord::RevisionSaveRoot(param_value),
        "insrsid" => ControlWord::InsertRsid(param_value),
        "delrsid" => ControlWord::DeleteRsid(param_value),
        "charrsid" => ControlWord::CharStyleRsid(param_value),
        "pararsid" => ControlWord::ParagraphRsid(param_value),
        "sectrsid" => ControlWord::SectionRsid(param_value),
        "tblrsid" => ControlWord::TableRsid(param_value),
        "xmlnstbl" => ControlWord::XmlNamespaceTable,
        "xmlns" => ControlWord::XmlNamespace(param_value),
        "xmlopen" => ControlWord::XmlOpen,
        "xmlclose" => ControlWord::XmlClose,
        "xmlattrname" => ControlWord::XmlAttributeName,
        "xmlattrvalue" => ControlWord::XmlAttributeValue,
        "mmath" | "moMath" => ControlWord::MathZoneInline,
        "mmathPara" | "moMathPara" => ControlWord::MathZoneDisplay,
        "mmathParaPr" | "moMathParaPr" => ControlWord::MathZoneParagraphProperties,
        "mmathPict" => ControlWord::MathZonePictureFallback,
        "macc" => ControlWord::MathAccent,
        "mbar" => ControlWord::MathBar,
        "mborderBox" => ControlWord::MathBorderBox,
        "mbox" => ControlWord::MathBox,
        "md" => ControlWord::MathDelimiter,
        "meqArr" => ControlWord::MathEquationArray,
        "mf" => ControlWord::MathFraction,
        "mfunc" => ControlWord::MathFunction,
        "mgroupChr" => ControlWord::MathGroupChar,
        "mlimlow" | "mlimLow" => ControlWord::MathLimitLower,
        "mlimupp" | "mlimUpp" => ControlWord::MathLimitUpper,
        "mm" => ControlWord::MathMatrix,
        "mnary" => ControlWord::MathNary,
        "mphant" => ControlWord::MathPhantom,
        "mrad" => ControlWord::MathRadical,
        "msPre" => ControlWord::MathScriptPre,
        "msSub" => ControlWord::MathScriptSub,
        "msSubSup" => ControlWord::MathScriptSubSup,
        "msSup" => ControlWord::MathScriptSup,
        "mr" => ControlWord::MathRun,
        "me" => ControlWord::MathElement,
        "mnum" => ControlWord::MathNumerator,
        "mden" => ControlWord::MathDenominator,
        "mdeg" => ControlWord::MathDegree,
        "msub" => ControlWord::MathSubscript,
        "msup" => ControlWord::MathSuperscript,
        "mlim" => ControlWord::MathLimit,
        "mfName" => ControlWord::MathFunctionName,
        "mmr" => ControlWord::MathMatrixRow,
        "maccPr" => ControlWord::MathAccentProperties,
        "mbarPr" => ControlWord::MathBarProperties,
        "mborderBoxPr" => ControlWord::MathBorderBoxProperties,
        "mboxPr" => ControlWord::MathBoxProperties,
        "mdPr" => ControlWord::MathDelimiterProperties,
        "meqArrPr" => ControlWord::MathEquationArrayProperties,
        "mfPr" => ControlWord::MathFractionProperties,
        "mfuncPr" => ControlWord::MathFunctionProperties,
        "mgroupChrPr" => ControlWord::MathGroupCharProperties,
        "mlimlowPr" | "mlimLowPr" => ControlWord::MathLimitLowerProperties,
        "mlimuppPr" | "mlimUppPr" => ControlWord::MathLimitUpperProperties,
        "mmPr" => ControlWord::MathMatrixProperties,
        "mnaryPr" => ControlWord::MathNaryProperties,
        "mphantPr" => ControlWord::MathPhantomProperties,
        "mradPr" => ControlWord::MathRadicalProperties,
        "msPrePr" => ControlWord::MathScriptPreProperties,
        "msSubPr" => ControlWord::MathScriptSubProperties,
        "msSubSupPr" => ControlWord::MathScriptSubSupProperties,
        "msSupPr" => ControlWord::MathScriptSupProperties,
        "mrPr" => ControlWord::MathRunProperties,
        "mctrlPr" => ControlWord::MathControlProperties,
        "mmcs" => ControlWord::MathMatrixColumns,
        "mmc" => ControlWord::MathMatrixColumn,
        "mmcPr" => ControlWord::MathMatrixColumnProperties,
        "margPr" => ControlWord::MathArgumentProperties,
        "margSz" => ControlWord::MathPropertyArgumentSize(param),
        "protstart" => ControlWord::ProtectionRangeStart,
        "protend" => ControlWord::ProtectionRangeEnd,
        "objalias" => ControlWord::ObjectAlias,
        "objsect" => ControlWord::ObjectSection,
        "objtime" => ControlWord::ObjectTime,
        "ebcstart" => ControlWord::EditableRegionStart(param),
        "ebcend" => ControlWord::EditableRegionEnd(param),
        "fchars" => ControlWord::KinsokuFollowing,
        "lchars" => ControlWord::KinsokuLeading,
        "ksulang" => ControlWord::KinsokuLanguage(param),
        "hl" => ControlWord::ShapeHyperlink,
        "hlloc" => ControlWord::ShapeHyperlinkLocation,
        "hlsrc" => ControlWord::ShapeHyperlinkSource,
        "hlfr" => ControlWord::ShapeHyperlinkFriendlyName,
        "tsrowd" => ControlWord::TableStyleRowDefaults(param),
        "tscfirstrow" => ControlWord::TableStyleFirstRow(param),
        "tsclastrow" => ControlWord::TableStyleLastRow(param),
        "tscfirstcol" => ControlWord::TableStyleFirstColumn(param),
        "tsclastcol" => ControlWord::TableStyleLastColumn(param),
        "tscbandhorzodd" => ControlWord::TableStyleBandHorizontalOdd(param),
        "tscbandhorzeven" => ControlWord::TableStyleBandHorizontalEven(param),
        "tscbandvertodd" => ControlWord::TableStyleBandVerticalOdd(param),
        "tscbandverteven" => ControlWord::TableStyleBandVerticalEven(param),
        "tscbandsh" => ControlWord::TableStyleBandSizeHorizontal(param),
        "tscbandsv" => ControlWord::TableStyleBandSizeVertical(param),
        "vertsect" => ControlWord::SectionVerticalRendering(param),
        "horzsect" => ControlWord::SectionHorizontalRendering(param),
        "nocolbal" => ControlWord::SectionNoColumnBalance(param),
        "sectdefaultcl" => ControlWord::SectionDefaultColumns(param),
        "noline" => ControlWord::ParagraphNoLineNumbering(param),
        "notabind" => ControlWord::ParagraphNoAutoTabIndent(param),
        "brdrbar" => ControlWord::BorderBar,
        "brdrbtw" => ControlWord::BorderBetween,
        "box" => ControlWord::BorderBox,
        "brdrhair" => ControlWord::BorderHairline,
        "softpage" => ControlWord::SoftPageBreak(param),
        "softcol" => ControlWord::SoftColumnBreak(param),
        "softline" => ControlWord::SoftLineBreak(param),
        "softlheight" => ControlWord::SoftLineHeight(param),
        "mtype" => ControlWord::MathPropertyType(param),
        "mgrow" => ControlWord::MathPropertyGrow(param),
        "mchr" => ControlWord::MathPropertyChar(param),
        "mbegChr" => ControlWord::MathPropertyBeginChar(param),
        "mendChr" => ControlWord::MathPropertyEndChar(param),
        "msepChr" => ControlWord::MathPropertySeparatorChar(param),
        "mpos" => ControlWord::MathPropertyPosition(param),
        "mvertJc" => ControlWord::MathPropertyVerticalJustify(param),
        "mbaseJc" => ControlWord::MathPropertyBaseJustify(param),
        "mjc" => ControlWord::MathPropertyJustify(param),
        "maln" => ControlWord::MathPropertyAlign(param),
        "malnScr" => ControlWord::MathPropertyAlignScript(param),
        "mdegHide" => ControlWord::MathPropertyDegreeHide(param),
        "mdiff" => ControlWord::MathPropertyDifferential(param),
        "mdiffSty" => ControlWord::MathPropertyDifferentialStyle(param),
        "mhideBot" => ControlWord::MathPropertyHideBottom(param),
        "mhideLeft" => ControlWord::MathPropertyHideLeft(param),
        "mhideRight" => ControlWord::MathPropertyHideRight(param),
        "mhideTop" => ControlWord::MathPropertyHideTop(param),
        "mlimLoc" | "mlimloc" => ControlWord::MathPropertyLimitLocation(param),
        "mplcHide" => ControlWord::MathPropertyPlaceholderHide(param),
        "msubHide" => ControlWord::MathPropertySubscriptHide(param),
        "msupHide" => ControlWord::MathPropertySuperscriptHide(param),
        "mstrikeBLTR" => ControlWord::MathPropertyStrikeBltr(param),
        "mstrikeH" => ControlWord::MathPropertyStrikeHorizontal(param),
        "mstrikeTLBR" => ControlWord::MathPropertyStrikeTlbr(param),
        "mstrikeV" => ControlWord::MathPropertyStrikeVertical(param),
        "msty" => ControlWord::MathPropertyStyle(param),
        "mscr" => ControlWord::MathPropertyScript(param),
        "mtransp" => ControlWord::MathPropertyTransparent(param),
        "mshow" => ControlWord::MathPropertyShow(param),
        "mshp" => ControlWord::MathPropertyShape(param),
        "mzeroAsc" => ControlWord::MathPropertyZeroAscent(param),
        "mzeroDesc" => ControlWord::MathPropertyZeroDescent(param),
        "mzeroWid" => ControlWord::MathPropertyZeroWidth(param),
        "mopEmu" => ControlWord::MathPropertyOperatorEmulator(param),
        "mnoBreak" => ControlWord::MathPropertyNoBreak(param),
        "mnor" => ControlWord::MathPropertyNormalText(param),
        "mlit" => ControlWord::MathPropertyLiteral(param),
        "mcGp" => ControlWord::MathPropertyMatrixColumnGap(param),
        "mcGpRule" => ControlWord::MathPropertyMatrixColumnGapRule(param),
        "mcSp" => ControlWord::MathPropertyMatrixColumnSpacing(param),
        "mcount" => ControlWord::MathPropertyMatrixCellCount(param),
        "mmcJc" => ControlWord::MathPropertyMatrixCellJustify(param),
        "mrSp" => ControlWord::MathPropertyRowSpacing(param),
        "mrSpRule" => ControlWord::MathPropertyRowSpacingRule(param),
        "mbrk" => ControlWord::MathPropertyBreak(param),
        "themedata" => ControlWord::ThemeData,
        "colorschememapping" => ControlWord::ColorSchemeMapping,
        "latentstyles" => ControlWord::LatentStyles,
        "lsdstimax" => ControlWord::LatentStyleMax(param_value),
        "lsdlockeddef" => ControlWord::LatentStyleLockedDefault(param_value),
        "lsdsemihiddendef" => ControlWord::LatentStyleSemiHiddenDefault(param_value),
        "lsdunhideuseddef" => ControlWord::LatentStyleUnhideUsedDefault(param_value),
        "lsdqformatdef" => ControlWord::LatentStyleQuickFormatDefault(param_value),
        "lsdprioritydef" => ControlWord::LatentStylePriorityDefault(param_value),
        "lsdlockedexcept" => ControlWord::LatentStyleExceptions,
        "lsdlocked" => ControlWord::LatentStyleLocked(param_value),
        "lsdsemihidden" => ControlWord::LatentStyleSemiHidden(param_value),
        "lsdunhideused" => ControlWord::LatentStyleUnhideUsed(param_value),
        "lsdqformat" => ControlWord::LatentStyleQuickFormat(param_value),
        "lsdpriority" => ControlWord::LatentStylePriority(param_value),
        "datastore" => ControlWord::DataStore,
        "mailmerge" => ControlWord::MailMerge,
        "mmconnectstr" => ControlWord::MailMergeConnectString,
        "mmconnectstrdata" => ControlWord::MailMergeConnectStringData,
        "mmdatasource" => ControlWord::MailMergeDataSource,
        "mmheadersource" => ControlWord::MailMergeHeaderSource,
        "mmlinktoquery" => ControlWord::MailMergeLinkToQuery(param_bool),
        "mmquery" => ControlWord::MailMergeQuery,
        "mmodso" => ControlWord::MailMergeDataSourceObject,
        "mmodsoactive" => ControlWord::MailMergeActiveRecord(param_value),
        "mmodsocoldelim" => ControlWord::MailMergeColumnDelimiter(param_value),
        "mmodsocolumn" => ControlWord::MailMergeColumnCount(param_value),
        "mmodsodynaddr" => ControlWord::MailMergeDynamicAddress(param_bool),
        "mmodsofhdr" => ControlWord::MailMergeFirstRowHeader(param_bool),
        "mmodsofilter" => ControlWord::MailMergeFilter,
        "mmodsofldmpdata" => ControlWord::MailMergeFieldMapData,
        "mmodsofmcolumn" => ControlWord::MailMergeFieldMapColumn(param_value),
        "mmodsohash" => ControlWord::MailMergeHash(param_value),
        "mmodsolid" => ControlWord::MailMergeId(param_value),
        "mmodsomappedname" => ControlWord::MailMergeMappedName,
        "mmodsoname" => ControlWord::MailMergeName,
        "mmodsorecipdata" => ControlWord::MailMergeRecipientData,
        "mmodsosort" => ControlWord::MailMergeSort,
        "mmodsosrc" => ControlWord::MailMergeSourceType(param_value),
        "mmodsotable" => ControlWord::MailMergeTable,
        "mmodsoudl" => ControlWord::MailMergeUdl,
        "mmodsoudldata" => ControlWord::MailMergeUdlData,
        "mmodsouniquetag" => ControlWord::MailMergeUniqueTag,
        "mmathPr" => ControlWord::MathProperties,
        "mbrkBin" => ControlWord::MathBreakBinary(param.unwrap_or(0)),
        "mbrkBinSub" => ControlWord::MathBreakBinarySubtraction(param.unwrap_or(0)),
        "mdefJc" => ControlWord::MathDefaultJustification(param.unwrap_or(1)),
        "mdispDef" => ControlWord::MathDisplayDefaults(param.unwrap_or(1)),
        "minterSp" => ControlWord::MathInterEquationSpacing(param.unwrap_or(0)),
        "mintLim" => ControlWord::MathIntegralLimitPlacement(param.unwrap_or(0)),
        "mintraSp" => ControlWord::MathIntraEquationSpacing(param.unwrap_or(0)),
        "mlMargin" => ControlWord::MathLeftMargin(param.unwrap_or(0)),
        "mmathFont" => ControlWord::MathFont(param.unwrap_or(0)),
        "mnaryLim" => ControlWord::MathNaryLimitPlacement(param.unwrap_or(1)),
        "mpostSp" => ControlWord::MathPostSpacing(param.unwrap_or(0)),
        "mpreSp" => ControlWord::MathPreSpacing(param.unwrap_or(0)),
        "mrMargin" => ControlWord::MathRightMargin(param.unwrap_or(0)),
        "msmallFrac" => ControlWord::MathSmallFractions(param.unwrap_or(0)),
        "mwrapIndent" => ControlWord::MathWrapIndent(param.unwrap_or(1440)),
        "mwrapRight" => ControlWord::MathWrapRight(param.unwrap_or(0)),
        "deftab" => ControlWord::DefaultTabWidth(param),
        "deflang" => ControlWord::DefaultLanguage(param_value),
        "deflangfe" => ControlWord::DefaultLanguageEastAsian(param_value),
        "adeflang" => ControlWord::DefaultLanguageComplexScript(param_value),
        "lang" => ControlWord::Language(param_value),
        "langfe" => ControlWord::LanguageEastAsian(param_value),
        "langnp" => ControlWord::LanguageNoProof(param_value),
        "langfenp" => ControlWord::LanguageEastAsianNoProof(param_value),
        "noproof" => ControlWord::NoProof(param_bool),
        "ltrch" => ControlWord::LeftToRightCharacter,
        "rtlch" => ControlWord::RightToLeftCharacter,
        "ltrpar" => ControlWord::LeftToRightParagraph,
        "rtlpar" => ControlWord::RightToLeftParagraph,
        "ltrdoc" => ControlWord::LeftToRightDocument,
        "rtldoc" => ControlWord::RightToLeftDocument,
        "ltrsect" => ControlWord::LeftToRightSection,
        "rtlsect" => ControlWord::RightToLeftSection,
        "ltrrow" => ControlWord::LeftToRightRow(param),
        "rtlrow" => ControlWord::RightToLeftRow(param),
        "taprtl" => ControlWord::TableRightToLeft(param_bool),
        "rtlgutter" => ControlWord::RightGutter(param_bool),
        "formprot" => ControlWord::FormProtection(param),
        "annotprot" => ControlWord::AnnotationProtection(param),
        "revprot" => ControlWord::RevisionProtection(param),
        "readprot" => ControlWord::ReadOnlyProtection(param),
        "allprot" => ControlWord::AllProtection(param),
        "enforceprot" => ControlWord::EnforceProtection(param),
        "protlevel" => ControlWord::ProtectionLevel(param),
        "password" => ControlWord::Password,
        "protusertbl" => ControlWord::ProtectionUserTable,
        "hyphauto" => ControlWord::HyphenateAutomatically(param),
        "hyphcaps" => ControlWord::HyphenateCapitalizedWords(param),
        "hyphconsec" => ControlWord::HyphenationConsecutiveLines(param),
        "hyphhotz" => ControlWord::HyphenationHotZone(param),
        "nextfile" => ControlWord::NextFile,
        "template" => ControlWord::DocumentTemplate,
        "viewkind" => ControlWord::DocumentViewKind(param),
        "viewscale" => ControlWord::DocumentViewScale(param),
        "viewzk" => ControlWord::DocumentZoomKind(param),
        "viewbksp" => ControlWord::DocumentViewBackgroundShapes(param),
        "viewnobound" => ControlWord::DocumentViewNoPageBoundaries(param),
        "donotshowmarkup" => ControlWord::HideReviewMarkup(param),
        "donotshowcomments" => ControlWord::HideReviewComments(param),
        "donotshowinsdel" => ControlWord::HideReviewInsertionsAndDeletions(param),
        "windowcaption" => ControlWord::WindowCaption,
        "xform" => ControlWord::XslTransform,
        "usexform" => ControlWord::UseXslTransform(param),
        "wgrffmtfilter" => ControlWord::StyleListFilter(param),
        "stylesortmethod" => ControlWord::StyleSortMethod(param),
        "readonlyrecommended" => ControlWord::ReadOnlyRecommended(param),
        "saveprevpict" => ControlWord::SavePreviousPicture(param),
        "writereservation" => ControlWord::WriteReservation(param),
        "writereservhash" => ControlWord::WriteReservationHash(param),
        "fromtext" => ControlWord::FromText(param),
        "fromhtml" => ControlWord::FromHtml(param),
        "doctype" => ControlWord::DocumentType(param),
        "makebackup" => ControlWord::MakeBackup(param),
        "defformat" => ControlWord::DefaultSaveFormat(param),
        "doctemp" => ControlWord::BoilerplateDocument(param),
        "muser" => ControlWord::Word97CompatibilityMode(param),
        "psover" => ControlWord::PostScriptOverText(param),
        "horzdoc" => ControlWord::HorizontalDocument(param),
        "vertdoc" => ControlWord::VerticalDocument(param),
        "jcompress" => ControlWord::CompressJustification(param),
        "jexpand" => ControlWord::ExpandJustification(param),
        "lnongrid" => ControlWord::LineBasedOnGrid(param),
        "fracwidth" => ControlWord::FractionalCharacterWidths(param),
        "ilfomacatclnup" => ControlWord::AbstractNumberingCleanup(param),
        "grfdocevents" => ControlWord::DocumentEventMask(param),
        "dgmargin" => ControlWord::DrawingGridFollowsMargins(param),
        "dgsnap" => ControlWord::SnapToDrawingGrid(param),
        "dghspace" => ControlWord::DrawingGridHorizontalSpacing(param),
        "dgvspace" => ControlWord::DrawingGridVerticalSpacing(param),
        "dghorigin" => ControlWord::DrawingGridHorizontalOrigin(param),
        "dgvorigin" => ControlWord::DrawingGridVerticalOrigin(param),
        "dghshow" => ControlWord::DrawingGridHorizontalShow(param),
        "dgvshow" => ControlWord::DrawingGridVerticalShow(param),
        "gutterprl" => ControlWord::ParallelGutter(param),
        "twoonone" => ControlWord::PrintTwoOnOne(param),
        "themelang" => ControlWord::ThemeLanguage(param),
        "themelangfe" => ControlWord::ThemeLanguageEastAsian(param),
        "themelangcs" => ControlWord::ThemeLanguageComplexScript(param),
        "relyonvml" => ControlWord::RelyOnVml(param),
        "validatexml" => ControlWord::ValidateXml(param),
        "showplaceholdtext" => ControlWord::ShowPlaceholderText(param),
        "ignoremixedcontent" => ControlWord::IgnoreMixedContent(param),
        "saveinvalidxml" => ControlWord::SaveInvalidXml(param),
        "showxmlerrors" => ControlWord::ShowXmlErrors(param),
        "donotembedsysfont" => ControlWord::DoNotEmbedSystemFonts(param),
        "donotembedlingdata" => ControlWord::DoNotEmbedLinguisticData(param),
        "trackmoves" => ControlWord::TrackMoves(param),
        "trackformatting" => ControlWord::TrackFormatting(param),
        "stylelocktheme" => ControlWord::LockDocumentTheme(param),
        "stylelockqfset" => ControlWord::LockQuickFormatSet(param),
        "usenormstyforlist" => ControlWord::UseNormalStyleForLists(param),
        "linkstyles" => ControlWord::UpdateStylesFromTemplate(param),
        "stylelock" => ControlWord::DeclareStyleRestrictions(param),
        "stylelockenforced" => ControlWord::EnforceStyleRestrictions(param),
        "stylelockbackcomp" => ControlWord::StyleRestrictionsBackwardCompatibility(param),
        "autofmtoverride" => ControlWord::AllowAutoFormatOverride(param),
        "bookfold" => ControlWord::BookFold(param),
        "bookfoldrev" => ControlWord::ReverseBookFold(param),
        "bookfoldsheets" => ControlWord::BookFoldSheets(param),
        "rempersonalinfo" => ControlWord::RemovePersonalInformation(param),
        "remdttm" => ControlWord::RemoveDateTimeInformation(param),
        "noextrasprl" => ControlWord::SuppressRaisedLoweredExtraSpacing(param),
        "sprstsp" => ControlWord::SuppressTopPageExtraSpacing(param),
        "sprsspbf" => ControlWord::SuppressSpaceBeforeAfterHardBreak(param),
        "sprslnsp" => ControlWord::SuppressWordPerfectExtraLineSpacing(param),
        "sprsbsp" => ControlWord::SuppressBottomPageExtraSpacing(param),
        "dntblnsbdb" => ControlWord::DoNotBalanceSbcsDbcs(param),
        "expshrtn" => ControlWord::ExpandSpacingAtShiftReturn(param),
        "nospaceforul" => ControlWord::DoNotAddSpaceForUnderline(param),
        "noultrlspc" => ControlWord::DoNotUnderlineTrailingSpaces(param),
        "noxlattoyen" => ControlWord::DoNotTranslateBackslashToYen(param),
        "lnbrkrule" => ControlWord::LegacyAsianLineBreakingRules(param),
        "otblrul" => ControlWord::CombineLegacyTableBorders(param),
        "alntblind" => ControlWord::DoNotAlignTableRowsIndependently(param),
        "lytcalctblwd" => ControlWord::DoNotUseRawTableWidth(param),
        "lyttblrtgr" => ControlWord::KeepTableRowsTogether(param),
        "nolnhtadjtbl" => ControlWord::DoNotAdjustTableLineHeight(param),
        "nobrkwrptbl" => ControlWord::DoNotBreakWrappedTablesAcrossPages(param),
        "nogrowautofit" => ControlWord::PreventAutofitGrowthIntoMargins(param),
        "newtblstyruls" => ControlWord::UseWord2003TableStyleRules(param),
        "splytwnine" => ControlWord::DoNotUseWord97ShapeLayout(param),
        "ftnlytwnine" => ControlWord::UseLegacyFootnoteLayout(param),
        "htmautsp" => ControlWord::UseHtmlParagraphAutoSpacing(param),
        "useltbaln" => ControlWord::PreserveLastTabAlignment(param),
        "oldas" => ControlWord::UseWord95AutoSpacing(param),
        "ApplyBrkRules" => ControlWord::ApplyThaiLineBreakingRules(param),
        "snaptogridincell" => ControlWord::SnapTextToGridInsideTable(param),
        "wrppunct" => ControlWord::AllowHangingPunctuation(param),
        "asianbrkrule" => ControlWord::UseAsianLineBreakingRules(param),
        "toplinepunct" => ControlWord::CompressPunctuationAtLineStart(param),
        "nocompatoptions" => ControlWord::NoCompatibilityOptions(param),
        "nouicompat" => ControlWord::NoUiCompatibility(param),
        "nofeaturethrottle" => ControlWord::NoFeatureThrottle(param),
        "forceupgrade" => ControlWord::ForceCompatibilityUpgrade(param),
        "noafcnsttbl" => ControlWord::PreserveAutofitTableWidthAroundShapes(param),
        "noindnmbrts" => ControlWord::UseHangingIndentAsNumberingTab(param),
        "felnbrelev" => ControlWord::UseLegacyKinsokuCharacters(param),
        "indrlsweleven" => ControlWord::UseLegacyFloatingObjectIndentation(param),
        "nocxsptable" => ControlWord::AllowContextualSpacingInTables(param),
        "notcvasp" => ControlWord::IgnoreCellVerticalAlignmentWithFloatingObjects(param),
        "notvatxbx" => ControlWord::IgnoreTextBoxVerticalAlignment(param),
        "spltpgpar" => ControlWord::SplitPageBreakParagraph(param),
        "hwelev" => ControlWord::UseFixedWidthHangul(param),
        "afelev" => ControlWord::UseLegacyAutofitWidthExpansion(param),
        "cachedcolbal" => ControlWord::UseCachedColumnBalancing(param),
        "utinl" => ControlWord::UnderlineNumberingSuffix(param),
        "notbrkcnstfrctbl" => ControlWord::DoNotSplitRowsAroundFloatingTables(param),
        "krnprsnet" => ControlWord::UseAnsiKerningPairs(param),

        // Index and table-of-contents source marks
        "xe" => ControlWord::IndexEntry,
        "xef" => ControlWord::IndexIdentifier(param),
        "bxe" => ControlWord::IndexBold(param_bool),
        "ixe" => ControlWord::IndexItalic(param_bool),
        "txe" => ControlWord::IndexReplacementText,
        "rxe" => ControlWord::IndexBookmarkRange,
        "yxe" => ControlWord::IndexYomi,
        "pxe" => ControlWord::IndexPronunciation,
        "tc" => ControlWord::TableOfContentsEntry,
        "tcn" => ControlWord::TableOfContentsEntryNoPage,
        "tcf" => ControlWord::TableOfContentsTable(param.unwrap_or(67)),
        "tcl" => ControlWord::TableOfContentsLevel(param.unwrap_or(1)),

        // Fonts
        "f" => ControlWord::FontNumber(param_value),
        "fs" => ControlWord::FontSize(param_value),
        "af" => ControlWord::AssociatedFontNumber(param),
        "afs" => ControlWord::AssociatedFontSize(param),
        "alang" => ControlWord::AssociatedLanguage(param),
        "ab" => ControlWord::AssociatedBold(param),
        "acaps" => ControlWord::AssociatedAllCaps(param),
        "acf" => ControlWord::AssociatedColor(param),
        "adn" => ControlWord::AssociatedBaselineDown(param),
        "aexpnd" => ControlWord::AssociatedExpansion(param),
        "ai" => ControlWord::AssociatedItalic(param),
        "aoutl" => ControlWord::AssociatedOutline(param),
        "ascaps" => ControlWord::AssociatedSmallCaps(param),
        "ashad" => ControlWord::AssociatedShadow(param),
        "astrike" => ControlWord::AssociatedStrike(param),
        "aul" => ControlWord::AssociatedUnderline(param),
        "auld" => ControlWord::AssociatedUnderlineDotted(param),
        "auldb" => ControlWord::AssociatedUnderlineDouble(param),
        "aulnone" => ControlWord::AssociatedUnderlineNone(param),
        "aulw" => ControlWord::AssociatedUnderlineWords(param),
        "aup" => ControlWord::AssociatedBaselineUp(param),
        "fcharset" => ControlWord::FontCharset(param_value),
        "fnil" => ControlWord::FontFamily("nil"),
        "froman" => ControlWord::FontFamily("roman"),
        "fswiss" => ControlWord::FontFamily("swiss"),
        "fmodern" => ControlWord::FontFamily("modern"),
        "fscript" => ControlWord::FontFamily("script"),
        "fdecor" => ControlWord::FontFamily("decor"),
        "flomajor" => ControlWord::FontTheme("flomajor"),
        "fhimajor" => ControlWord::FontTheme("fhimajor"),
        "fdbmajor" => ControlWord::FontTheme("fdbmajor"),
        "fbimajor" => ControlWord::FontTheme("fbimajor"),
        "flominor" => ControlWord::FontTheme("flominor"),
        "fhiminor" => ControlWord::FontTheme("fhiminor"),
        "fdbminor" => ControlWord::FontTheme("fdbminor"),
        "fbiminor" => ControlWord::FontTheme("fbiminor"),
        "ftech" => ControlWord::FontFamily("tech"),
        "fbidi" => ControlWord::FontBidi(param),
        "fprq" => ControlWord::FontPitch(param_value),
        "cpg" => ControlWord::FontCodePage(param_value),
        "falt" => ControlWord::FontAlternateName,
        "fname" => ControlWord::FontNonTaggedName,
        "panose" => ControlWord::FontPanose,
        "fontemb" => ControlWord::FontEmbedded,
        "ftnil" => ControlWord::FontEmbeddedType("nil"),
        "fttruetype" => ControlWord::FontEmbeddedType("truetype"),
        "fontfile" => ControlWord::FontFile,
        "filetbl" => ControlWord::FileTable,
        "file" => ControlWord::FileEntry,
        "fid" => ControlWord::FileId(param_value),
        "frelative" => ControlWord::FileRelative(param_value),
        "fosnum" => ControlWord::FileOperatingSystem(param_value),
        "fvalidmac" => ControlWord::FileValidMac,
        "fvaliddos" => ControlWord::FileValidDos,
        "fvalidntfs" => ControlWord::FileValidNtfs,
        "fvalidhpfs" => ControlWord::FileValidHpfs,
        "fnetwork" => ControlWord::FileNetwork,
        "fnonfilesys" => ControlWord::FileNonFileSystem,
        "upr" => ControlWord::UnicodeAlternate,
        "ud" => ControlWord::UnicodeAlternateDestination,

        // Colors
        "red" => ControlWord::Red(param_value),
        "green" => ControlWord::Green(param_value),
        "blue" => ControlWord::Blue(param_value),
        "cf" => ControlWord::ColorForeground(param_value),
        "cb" => ControlWord::ColorBackground(param),

        // Character formatting
        "b" => ControlWord::Bold(param_bool),
        "i" => ControlWord::Italic(param_bool),
        "ul" => ControlWord::Underline(param_bool),
        "ulnone" => ControlWord::UnderlineNone,
        "uldb" => ControlWord::UnderlineDouble,
        "uld" => ControlWord::UnderlineDotted,
        "uldash" => ControlWord::UnderlineDashed,
        "uldashd" => ControlWord::UnderlineDashDot,
        "uldashdd" => ControlWord::UnderlineDashDotDot,
        "ulw" => ControlWord::UnderlineWords,
        "ulth" => ControlWord::UnderlineThick,
        "ulwave" => ControlWord::UnderlineWave,
        "ulhair" => ControlWord::UnderlineHairline,
        "ulthd" => ControlWord::UnderlineThickDotted,
        "ulthdash" => ControlWord::UnderlineThickDashed,
        "ulthdashd" => ControlWord::UnderlineThickDashDot,
        "ulthdashdd" => ControlWord::UnderlineThickDashDotDot,
        "ulthldash" => ControlWord::UnderlineThickLongDash,
        "ulldash" => ControlWord::UnderlineLongDash,
        "ulhwave" => ControlWord::UnderlineHeavyWave,
        "ululdbwave" => ControlWord::UnderlineDoubleWave,
        "ulc" => ControlWord::UnderlineColor(param_value),
        "strike" => ControlWord::Strike(param_bool),
        "striked" => ControlWord::DoubleStrike(param_bool),
        "super" => ControlWord::Superscript(param_bool),
        "sub" => ControlWord::Subscript(param_bool),
        "nosupersub" => ControlWord::NoSuperSub,
        "up" => ControlWord::BaselineUp(param.unwrap_or(6)),
        "dn" => ControlWord::BaselineDown(param.unwrap_or(6)),
        "scaps" => ControlWord::SmallCaps(param_bool),
        "caps" => ControlWord::AllCaps(param_bool),
        "v" => ControlWord::Hidden(param_bool),
        "outl" => ControlWord::Outline(param_bool),
        "shad" => ControlWord::Shadow(param_bool),
        "embo" => ControlWord::Emboss(param_bool),
        "impr" => ControlWord::Imprint(param_bool),
        "expnd" => ControlWord::CharSpacing(param_value),
        "expndtw" => ControlWord::CharSpacingTwips(param_value),
        "charscalex" => ControlWord::CharScale(param_value),
        "kerning" => ControlWord::Kerning(param_value),
        "highlight" => ControlWord::Highlight(param_value),
        "chbrdr" => ControlWord::CharacterBorder(param),
        "chshdng" => ControlWord::CharacterShading(param),
        "chcfpat" => ControlWord::CharacterForegroundPattern(param),
        "chcbpat" => ControlWord::CharacterBackgroundPattern(param),
        "plain" => ControlWord::Plain,
        "loch" => ControlWord::LowAnsiCharacter(param),
        "hich" => ControlWord::HighAnsiCharacter(param),
        "dbch" => ControlWord::DoubleByteCharacter(param),
        "fcs" => ControlWord::FontComplexScript(param),
        "cgrid" => ControlWord::CharacterGrid(param),
        "animtext" => ControlWord::AnimatedText(param),
        "fittext" => ControlWord::FitText(param),
        "accnone" => ControlWord::EmphasisMark(crate::EmphasisMark::None, param),
        "accdot" => ControlWord::EmphasisMark(crate::EmphasisMark::Dot, param),
        "acccomma" => ControlWord::EmphasisMark(crate::EmphasisMark::Comma, param),
        "accunderdot" => ControlWord::EmphasisMark(crate::EmphasisMark::UnderDot, param),
        "acccircle" => ControlWord::EmphasisMark(crate::EmphasisMark::Circle, param),

        // Paragraph
        "par" => ControlWord::Par,
        "page" => ControlWord::Page(param),
        "column" => ControlWord::Column(param),
        "pard" => ControlWord::Pard,
        "ql" => ControlWord::LeftAlign,
        "qr" => ControlWord::RightAlign,
        "qc" => ControlWord::Center,
        "qj" => ControlWord::Justify,

        // Paragraph spacing/indent
        "sb" => ControlWord::SpaceBefore(param_value),
        "sa" => ControlWord::SpaceAfter(param_value),
        "sl" => ControlWord::SpaceBetween(param_value),
        "slmult" => ControlWord::LineMultiple(param_bool),
        "sbauto" => ControlWord::SpaceBeforeAuto(param),
        "saauto" => ControlWord::SpaceAfterAuto(param),
        "lisb" => ControlWord::ListSpaceBefore(param),
        "lisa" => ControlWord::ListSpaceAfter(param),
        "nosnaplinegrid" => ControlWord::NoSnapLineGrid(param),
        "contextualspace" => ControlWord::ContextualSpacing(param),
        "li" => ControlWord::LeftIndent(param_value),
        "ri" => ControlWord::RightIndent(param_value),
        "fi" => ControlWord::FirstLineIndent(param_value),
        "lin" => ControlWord::LogicalLeftIndent(param),
        "rin" => ControlWord::LogicalRightIndent(param),
        "cufi" => ControlWord::CharacterFirstLineIndent(param),
        "culi" => ControlWord::CharacterLeftIndent(param),
        "curi" => ControlWord::CharacterRightIndent(param),
        "indmirror" => ControlWord::MirrorIndents(param),

        // Paragraph additional properties
        "keep" => ControlWord::KeepTogether,
        "keepn" => ControlWord::KeepNext,
        "sbys" => ControlWord::SideBySide(param_bool),
        "pagebb" => ControlWord::PageBreakBefore,
        "widctlpar" => ControlWord::WidowControl,
        "nowidctlpar" => ControlWord::NoWidowControl(param),
        "dropcapli" => ControlWord::DropCapLines(param),
        "dropcapt" => ControlWord::DropCapType(param),
        "hyphpar" => ControlWord::ParagraphHyphenation(param),
        "aspalpha" => ControlWord::AutoSpaceAlphabetic(param),
        "aspnum" => ControlWord::AutoSpaceNumbers(param),
        "adjustright" => ControlWord::AdjustRightIndent(param),
        "wrapdefault" => ControlWord::WrapDefault(param),
        "nocwrap" => ControlWord::NoCharacterWrap(param),
        "nowwrap" => ControlWord::NoWordWrap(param),
        "nooverflow" => ControlWord::NoOverflow(param),
        "faauto" => ControlWord::FontAlignAuto(param),
        "fahang" => ControlWord::FontAlignHanging(param),
        "facenter" => ControlWord::FontAlignCenter(param),
        "faroman" => ControlWord::FontAlignRoman(param),
        "favar" => ControlWord::FontAlignVariable(param),
        "fafixed" => ControlWord::FontAlignFixed(param),

        // Tables
        "trowd" => ControlWord::TableRowDefaults,
        "irow" => ControlWord::TableRowIndex(param),
        "irowband" => ControlWord::TableRowBandIndex(param),
        "lastrow" => ControlWord::TableLastRow(param),
        "tbllkborder" => {
            ControlWord::TableAutoformatFlag(crate::TableAutoformatFlag::Border, param)
        },
        "tbllkshading" => {
            ControlWord::TableAutoformatFlag(crate::TableAutoformatFlag::Shading, param)
        },
        "tbllkfont" => ControlWord::TableAutoformatFlag(crate::TableAutoformatFlag::Font, param),
        "tbllkcolor" => ControlWord::TableAutoformatFlag(crate::TableAutoformatFlag::Color, param),
        "tbllkbestfit" => {
            ControlWord::TableAutoformatFlag(crate::TableAutoformatFlag::BestFit, param)
        },
        "tbllkhdrrows" => {
            ControlWord::TableAutoformatFlag(crate::TableAutoformatFlag::HeaderRows, param)
        },
        "tbllklastrow" => {
            ControlWord::TableAutoformatFlag(crate::TableAutoformatFlag::LastRow, param)
        },
        "tbllkhdrcols" => {
            ControlWord::TableAutoformatFlag(crate::TableAutoformatFlag::HeaderColumns, param)
        },
        "tbllklastcol" => {
            ControlWord::TableAutoformatFlag(crate::TableAutoformatFlag::LastColumn, param)
        },
        "tbllknorowband" => {
            ControlWord::TableAutoformatFlag(crate::TableAutoformatFlag::NoRowBanding, param)
        },
        "tbllknocolband" => {
            ControlWord::TableAutoformatFlag(crate::TableAutoformatFlag::NoColumnBanding, param)
        },
        "row" => ControlWord::TableRow,
        "cell" => ControlWord::TableCell,
        "cellx" => ControlWord::CellX(param_value),
        "intbl" => ControlWord::InTable,
        "trgaph" => ControlWord::TableRowGap(param),
        "trleft" => ControlWord::TableRowLeft(param),
        "trrh" => ControlWord::TableRowHeight(param),
        "trftsWidth" => ControlWord::TablePreferredWidthUnit(crate::TableDistanceScope::Row, param),
        "trwWidth" => ControlWord::TablePreferredWidthValue(crate::TableDistanceScope::Row, param),
        "trftsWidthB" => ControlWord::TableInvisibleWidthUnit(false, param),
        "trwWidthB" => ControlWord::TableInvisibleWidthValue(false, param),
        "trftsWidthA" => ControlWord::TableInvisibleWidthUnit(true, param),
        "trwWidthA" => ControlWord::TableInvisibleWidthValue(true, param),
        "clftsWidth" => {
            ControlWord::TablePreferredWidthUnit(crate::TableDistanceScope::Cell, param)
        },
        "clwWidth" => ControlWord::TablePreferredWidthValue(crate::TableDistanceScope::Cell, param),
        "trautofit" => ControlWord::TableAutoFit(param),
        "tblind" => ControlWord::TableIndentValue(param),
        "tblindtype" => ControlWord::TableIndentUnit(param),
        "trpaddl" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Left,
            },
            param,
        ),
        "trpaddr" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Right,
            },
            param,
        ),
        "trpaddt" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Top,
            },
            param,
        ),
        "trpaddb" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Bottom,
            },
            param,
        ),
        "trpaddfl" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Left,
            },
            param,
        ),
        "trpaddfr" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Right,
            },
            param,
        ),
        "trpaddft" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Top,
            },
            param,
        ),
        "trpaddfb" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Bottom,
            },
            param,
        ),
        "trspdl" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Left,
            },
            param,
        ),
        "trspdr" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Right,
            },
            param,
        ),
        "trspdt" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Top,
            },
            param,
        ),
        "trspdb" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Bottom,
            },
            param,
        ),
        "trspdfl" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Left,
            },
            param,
        ),
        "trspdfr" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Right,
            },
            param,
        ),
        "trspdft" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Top,
            },
            param,
        ),
        "trspdfb" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Bottom,
            },
            param,
        ),
        "clpadl" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Left,
            },
            param,
        ),
        "clpadr" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Right,
            },
            param,
        ),
        "clpadt" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Top,
            },
            param,
        ),
        "clpadb" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Bottom,
            },
            param,
        ),
        "clpadfl" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Left,
            },
            param,
        ),
        "clpadfr" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Right,
            },
            param,
        ),
        "clpadft" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Top,
            },
            param,
        ),
        "clpadfb" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Bottom,
            },
            param,
        ),
        "clspdl" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Left,
            },
            param,
        ),
        "clspdr" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Right,
            },
            param,
        ),
        "clspdt" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Top,
            },
            param,
        ),
        "clspdb" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Bottom,
            },
            param,
        ),
        "clspdfl" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Left,
            },
            param,
        ),
        "clspdfr" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Right,
            },
            param,
        ),
        "clspdft" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Top,
            },
            param,
        ),
        "clspdfb" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Bottom,
            },
            param,
        ),
        "tphcol" => {
            ControlWord::TableHorizontalReference(crate::TableHorizontalReference::Column, param)
        },
        "tphmrg" => {
            ControlWord::TableHorizontalReference(crate::TableHorizontalReference::Margin, param)
        },
        "tphpg" => {
            ControlWord::TableHorizontalReference(crate::TableHorizontalReference::Page, param)
        },
        "tpvmrg" => {
            ControlWord::TableVerticalReference(crate::TableVerticalReference::Margin, param)
        },
        "tpvpara" => {
            ControlWord::TableVerticalReference(crate::TableVerticalReference::Paragraph, param)
        },
        "tpvpg" => ControlWord::TableVerticalReference(crate::TableVerticalReference::Page, param),
        "tposx" => ControlWord::TableHorizontalOffset(false, param),
        "tposnegx" => ControlWord::TableHorizontalOffset(true, param),
        "tposxc" => {
            ControlWord::TableHorizontalPosition(crate::TableHorizontalPosition::Center, param)
        },
        "tposxi" => {
            ControlWord::TableHorizontalPosition(crate::TableHorizontalPosition::Inside, param)
        },
        "tposxl" => {
            ControlWord::TableHorizontalPosition(crate::TableHorizontalPosition::Left, param)
        },
        "tposxo" => {
            ControlWord::TableHorizontalPosition(crate::TableHorizontalPosition::Outside, param)
        },
        "tposxr" => {
            ControlWord::TableHorizontalPosition(crate::TableHorizontalPosition::Right, param)
        },
        "tposy" => ControlWord::TableVerticalOffset(false, param),
        "tposnegy" => ControlWord::TableVerticalOffset(true, param),
        "tposyb" => ControlWord::TableVerticalPosition(crate::TableVerticalPosition::Bottom, param),
        "tposyc" => ControlWord::TableVerticalPosition(crate::TableVerticalPosition::Center, param),
        "tposyil" => {
            ControlWord::TableVerticalPosition(crate::TableVerticalPosition::Inline, param)
        },
        "tposyin" => {
            ControlWord::TableVerticalPosition(crate::TableVerticalPosition::Inside, param)
        },
        "tposyout" => {
            ControlWord::TableVerticalPosition(crate::TableVerticalPosition::Outside, param)
        },
        "tposyt" => ControlWord::TableVerticalPosition(crate::TableVerticalPosition::Top, param),
        "tdfrmtxtLeft" => ControlWord::TableWrapDistance(crate::TableEdge::Left, param),
        "tdfrmtxtRight" => ControlWord::TableWrapDistance(crate::TableEdge::Right, param),
        "tdfrmtxtTop" => ControlWord::TableWrapDistance(crate::TableEdge::Top, param),
        "tdfrmtxtBottom" => ControlWord::TableWrapDistance(crate::TableEdge::Bottom, param),
        "tabsnoovrlp" => ControlWord::TableNoOverlap(param),
        "trhdr" => ControlWord::TableRowHeader(param),
        "trkeep" => ControlWord::TableRowKeep(param),
        "trkeepfollow" => ControlWord::TableRowKeepFollow(param),
        "trql" => ControlWord::TableRowAlignment(crate::TableRowAlignment::Left, param),
        "trqc" => ControlWord::TableRowAlignment(crate::TableRowAlignment::Center, param),
        "trqr" => ControlWord::TableRowAlignment(crate::TableRowAlignment::Right, param),
        "clvertalt" => {
            ControlWord::TableCellVerticalAlignment(crate::TableCellVerticalAlignment::Top, param)
        },
        "clvertalc" => ControlWord::TableCellVerticalAlignment(
            crate::TableCellVerticalAlignment::Center,
            param,
        ),
        "clvertalb" => ControlWord::TableCellVerticalAlignment(
            crate::TableCellVerticalAlignment::Bottom,
            param,
        ),
        "cltxlrtb" => {
            ControlWord::TableCellTextFlow(crate::TableCellTextFlow::LeftToRightTopToBottom, param)
        },
        "cltxtbrl" => {
            ControlWord::TableCellTextFlow(crate::TableCellTextFlow::RightToLeftTopToBottom, param)
        },
        "cltxbtlr" => {
            ControlWord::TableCellTextFlow(crate::TableCellTextFlow::LeftToRightBottomToTop, param)
        },
        "cltxlrtbv" => ControlWord::TableCellTextFlow(
            crate::TableCellTextFlow::LeftToRightTopToBottomVertical,
            param,
        ),
        "cltxtbrlv" => ControlWord::TableCellTextFlow(
            crate::TableCellTextFlow::TopToBottomRightToLeftVertical,
            param,
        ),
        "clFitText" => ControlWord::TableCellFitText(param),
        "clNoWrap" => ControlWord::TableCellNoWrap(param),
        "clhidemark" => ControlWord::TableCellHideMark(param),
        "clmgf" => ControlWord::TableCellMerge(
            crate::TableCellMergeAxis::Horizontal,
            crate::TableCellMergeRole::First,
            param,
        ),
        "clmrg" => ControlWord::TableCellMerge(
            crate::TableCellMergeAxis::Horizontal,
            crate::TableCellMergeRole::Continuation,
            param,
        ),
        "clvmgf" => ControlWord::TableCellMerge(
            crate::TableCellMergeAxis::Vertical,
            crate::TableCellMergeRole::First,
            param,
        ),
        "clvmrg" => ControlWord::TableCellMerge(
            crate::TableCellMergeAxis::Vertical,
            crate::TableCellMergeRole::Continuation,
            param,
        ),
        "trbrdrt" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::Row(crate::TableRowBorderSide::Top),
            param,
        ),
        "trbrdrl" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::Row(crate::TableRowBorderSide::Left),
            param,
        ),
        "trbrdrb" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::Row(crate::TableRowBorderSide::Bottom),
            param,
        ),
        "trbrdrr" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::Row(crate::TableRowBorderSide::Right),
            param,
        ),
        "trbrdrh" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::Row(crate::TableRowBorderSide::Horizontal),
            param,
        ),
        "trbrdrv" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::Row(crate::TableRowBorderSide::Vertical),
            param,
        ),
        "clbrdrt" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::Cell(crate::TableCellBorderSide::Top),
            param,
        ),
        "clbrdrl" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::Cell(crate::TableCellBorderSide::Left),
            param,
        ),
        "clbrdrb" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::Cell(crate::TableCellBorderSide::Bottom),
            param,
        ),
        "clbrdrr" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::Cell(crate::TableCellBorderSide::Right),
            param,
        ),
        "cldglu" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::Cell(
                crate::TableCellBorderSide::UpperLeftToLowerRight,
            ),
            param,
        ),
        "cldgll" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::Cell(
                crate::TableCellBorderSide::UpperRightToLowerLeft,
            ),
            param,
        ),
        "tsbrdrt" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::StyleDefault(crate::TableStyleBorderSide::Top),
            param,
        ),
        "tsbrdrl" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::StyleDefault(crate::TableStyleBorderSide::Left),
            param,
        ),
        "tsbrdrb" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::StyleDefault(crate::TableStyleBorderSide::Bottom),
            param,
        ),
        "tsbrdrr" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::StyleDefault(crate::TableStyleBorderSide::Right),
            param,
        ),
        "tsbrdrh" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::StyleDefault(
                crate::TableStyleBorderSide::HorizontalInside,
            ),
            param,
        ),
        "tsbrdrv" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::StyleDefault(
                crate::TableStyleBorderSide::VerticalInside,
            ),
            param,
        ),
        "tsbrdrdgl" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::StyleDefault(
                crate::TableStyleBorderSide::DiagonalUpperLeftToLowerRight,
            ),
            param,
        ),
        "tsbrdrdg" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::StyleDefault(
                crate::TableStyleBorderSide::DiagonalUpperRightToLowerLeft,
            ),
            param,
        ),
        "tscellpaddl" => ControlWord::TableDefaultDistanceValue(
            crate::TableDistanceKind::Padding,
            crate::TableEdge::Left,
            param,
        ),
        "tscellpaddr" => ControlWord::TableDefaultDistanceValue(
            crate::TableDistanceKind::Padding,
            crate::TableEdge::Right,
            param,
        ),
        "tscellpaddt" => ControlWord::TableDefaultDistanceValue(
            crate::TableDistanceKind::Padding,
            crate::TableEdge::Top,
            param,
        ),
        "tscellpaddb" => ControlWord::TableDefaultDistanceValue(
            crate::TableDistanceKind::Padding,
            crate::TableEdge::Bottom,
            param,
        ),
        "tscellpaddfl" => ControlWord::TableDefaultDistanceUnit(
            crate::TableDistanceKind::Padding,
            crate::TableEdge::Left,
            param,
        ),
        "tscellpaddfr" => ControlWord::TableDefaultDistanceUnit(
            crate::TableDistanceKind::Padding,
            crate::TableEdge::Right,
            param,
        ),
        "tscellpaddft" => ControlWord::TableDefaultDistanceUnit(
            crate::TableDistanceKind::Padding,
            crate::TableEdge::Top,
            param,
        ),
        "tscellpaddfb" => ControlWord::TableDefaultDistanceUnit(
            crate::TableDistanceKind::Padding,
            crate::TableEdge::Bottom,
            param,
        ),
        "tscellspcl" => ControlWord::TableDefaultDistanceValue(
            crate::TableDistanceKind::Spacing,
            crate::TableEdge::Left,
            param,
        ),
        "tscellspcr" => ControlWord::TableDefaultDistanceValue(
            crate::TableDistanceKind::Spacing,
            crate::TableEdge::Right,
            param,
        ),
        "tscellspct" => ControlWord::TableDefaultDistanceValue(
            crate::TableDistanceKind::Spacing,
            crate::TableEdge::Top,
            param,
        ),
        "tscellspcb" => ControlWord::TableDefaultDistanceValue(
            crate::TableDistanceKind::Spacing,
            crate::TableEdge::Bottom,
            param,
        ),
        "tscellspcfl" => ControlWord::TableDefaultDistanceUnit(
            crate::TableDistanceKind::Spacing,
            crate::TableEdge::Left,
            param,
        ),
        "tscellspcfr" => ControlWord::TableDefaultDistanceUnit(
            crate::TableDistanceKind::Spacing,
            crate::TableEdge::Right,
            param,
        ),
        "tscellspcft" => ControlWord::TableDefaultDistanceUnit(
            crate::TableDistanceKind::Spacing,
            crate::TableEdge::Top,
            param,
        ),
        "tscellspcfb" => ControlWord::TableDefaultDistanceUnit(
            crate::TableDistanceKind::Spacing,
            crate::TableEdge::Bottom,
            param,
        ),
        "tscellwidthfts" => ControlWord::TableDefaultCellWidthUnit(param),
        "tscellwidth" => ControlWord::TableDefaultCellWidthValue(param),
        "trshdng" => ControlWord::TableShadingAmount(crate::TableDistanceScope::Row, param),
        "trcfpat" => ControlWord::TableShadingForeground(crate::TableDistanceScope::Row, param),
        "trcbpat" => ControlWord::TableShadingBackground(crate::TableDistanceScope::Row, param),
        "trpat" => ControlWord::TableRowShadingPatternIndex(param),
        "clshdng" => ControlWord::TableShadingAmount(crate::TableDistanceScope::Cell, param),
        "clshdngraw" => ControlWord::TableShadingRawAmount(crate::TableDistanceScope::Cell, param),
        "clshdrawnil" => ControlWord::TableShadingRawNil(crate::TableDistanceScope::Cell, param),
        "clcfpat" => ControlWord::TableShadingForeground(crate::TableDistanceScope::Cell, param),
        "clcbpat" => ControlWord::TableShadingBackground(crate::TableDistanceScope::Cell, param),
        "trbghoriz" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Row,
            crate::ShadingPattern::Horizontal,
            param,
        ),
        "trbgvert" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Row,
            crate::ShadingPattern::Vertical,
            param,
        ),
        "trbgfdiag" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Row,
            crate::ShadingPattern::ForwardDiagonal,
            param,
        ),
        "trbgbdiag" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Row,
            crate::ShadingPattern::BackwardDiagonal,
            param,
        ),
        "trbgcross" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Row,
            crate::ShadingPattern::Cross,
            param,
        ),
        "trbgdcross" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Row,
            crate::ShadingPattern::DiagonalCross,
            param,
        ),
        "trbgdkhor" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Row,
            crate::ShadingPattern::DarkHorizontal,
            param,
        ),
        "trbgdkvert" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Row,
            crate::ShadingPattern::DarkVertical,
            param,
        ),
        "trbgdkfdiag" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Row,
            crate::ShadingPattern::DarkForwardDiagonal,
            param,
        ),
        "trbgdkbdiag" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Row,
            crate::ShadingPattern::DarkBackwardDiagonal,
            param,
        ),
        "trbgdkcross" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Row,
            crate::ShadingPattern::DarkCross,
            param,
        ),
        "trbgdkdcross" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Row,
            crate::ShadingPattern::DarkDiagonalCross,
            param,
        ),
        "clbghoriz" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Cell,
            crate::ShadingPattern::Horizontal,
            param,
        ),
        "clbgvert" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Cell,
            crate::ShadingPattern::Vertical,
            param,
        ),
        "clbgfdiag" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Cell,
            crate::ShadingPattern::ForwardDiagonal,
            param,
        ),
        "clbgbdiag" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Cell,
            crate::ShadingPattern::BackwardDiagonal,
            param,
        ),
        "clbgcross" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Cell,
            crate::ShadingPattern::Cross,
            param,
        ),
        "clbgdcross" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Cell,
            crate::ShadingPattern::DiagonalCross,
            param,
        ),
        "clbgdkhor" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Cell,
            crate::ShadingPattern::DarkHorizontal,
            param,
        ),
        "clbgdkvert" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Cell,
            crate::ShadingPattern::DarkVertical,
            param,
        ),
        "clbgdkfdiag" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Cell,
            crate::ShadingPattern::DarkForwardDiagonal,
            param,
        ),
        "clbgdkbdiag" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Cell,
            crate::ShadingPattern::DarkBackwardDiagonal,
            param,
        ),
        "clbgdkcross" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Cell,
            crate::ShadingPattern::DarkCross,
            param,
        ),
        "clbgdkdcross" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Cell,
            crate::ShadingPattern::DarkDiagonalCross,
            param,
        ),
        "nestcell" => ControlWord::NestedTableCell(param),
        "nestrow" => ControlWord::NestedTableRow(param),
        "nesttableprops" => ControlWord::NestedTableProperties(param),
        "nonesttables" => ControlWord::NoNestedTables(param),

        // Borders
        "brdrt" => ControlWord::BorderTop,
        "brdrb" => ControlWord::BorderBottom,
        "brdrl" => ControlWord::BorderLeft,
        "brdrr" => ControlWord::BorderRight,
        "brdrs" => ControlWord::BorderSingle,
        "brdrdot" => ControlWord::BorderDotted,
        "brdrdash" => ControlWord::BorderDashed,
        "brdrdb" => ControlWord::BorderDouble,
        "brdrtriple" => ControlWord::BorderTriple,
        "brdrwavy" => ControlWord::BorderWave,
        "brdrnone" | "brdrnil" => ControlWord::BorderNone,
        "brdrth" => ControlWord::BorderThick,
        "brdrdashsm" => ControlWord::BorderDashSmall,
        "brdrdashd" => ControlWord::BorderDotDash,
        "brdrdashdd" => ControlWord::BorderDotDotDash,
        "brdrtnthsg" => ControlWord::BorderThinThickSmall,
        "brdrthtnsg" => ControlWord::BorderThickThinSmall,
        "brdrtnthtnsg" => ControlWord::BorderThinThickThinSmall,
        "brdrtnthmg" => ControlWord::BorderThinThickMedium,
        "brdrthtnmg" => ControlWord::BorderThickThinMedium,
        "brdrtnthtnmg" => ControlWord::BorderThinThickThinMedium,
        "brdrtnthlg" => ControlWord::BorderThinThickLarge,
        "brdrthtnlg" => ControlWord::BorderThickThinLarge,
        "brdrtnthtnlg" => ControlWord::BorderThinThickThinLarge,
        "brdrwavydb" => ControlWord::BorderWavyDouble,
        "brdrdashdotstr" => ControlWord::BorderStriped,
        "brdremboss" => ControlWord::BorderEmbossed,
        "brdrengrave" => ControlWord::BorderEngraved,
        "brdroutset" => ControlWord::BorderOutset,
        "brdrinset" => ControlWord::BorderInset,
        "brdrsh" => ControlWord::BorderShadow,
        "brdrframe" => ControlWord::BorderFrame,
        "brdrw" => ControlWord::BorderWidth(param),
        "brdrcf" => ControlWord::BorderColor(param),
        "brsp" => ControlWord::BorderSpace(param),
        "pgbrdrt" => ControlWord::PageBorderTop,
        "pgbrdrl" => ControlWord::PageBorderLeft,
        "pgbrdrb" => ControlWord::PageBorderBottom,
        "pgbrdrr" => ControlWord::PageBorderRight,
        "pgbrdropt" => ControlWord::PageBorderOptions(param),
        "pgbrdrhead" => ControlWord::PageBorderSurroundHeader,
        "pgbrdrfoot" => ControlWord::PageBorderSurroundFooter,
        "pgbrdrsnap" => ControlWord::PageBorderSnap,
        "brdrart" => ControlWord::PageBorderArt(param),

        // Shading
        "shading" => ControlWord::Shading(param),
        "cfpat" => ControlWord::ForegroundPattern(param),
        "cbpat" => ControlWord::BackgroundPattern(param),

        // Tab stops
        "tql" => ControlWord::TabLeft(param),
        "tqr" => ControlWord::TabRight(param),
        "tqc" => ControlWord::TabCenter(param),
        "tqdec" => ControlWord::TabDecimal(param),
        "tb" => ControlWord::TabBar(param),
        "tx" => ControlWord::TabPosition(param),
        "tldot" => ControlWord::TabLeaderDot(param),
        "tlmdot" => ControlWord::TabLeaderMiddleDot(param),
        "tlhyph" => ControlWord::TabLeaderHyphen(param),
        "tlul" => ControlWord::TabLeaderUnderscore(param),
        "tlth" => ControlWord::TabLeaderThick(param),
        "tleq" => ControlWord::TabLeaderEqual(param),

        // Lists
        "listtable" => ControlWord::ListTable,
        "list" => ControlWord::List,
        "listtemplateid" => ControlWord::ListTemplateId(param_value),
        "listsimple" => ControlWord::ListSimple(param_bool),
        "listhybrid" => ControlWord::ListHybrid(param_bool),
        "listname" => ControlWord::ListName,
        "liststylename" => ControlWord::ListStyleName,
        "listtext" => ControlWord::GeneratedListText,
        "listpicture" | "listpict" => ControlWord::ListPicture(param),
        "shppict" => ControlWord::ShapePicture(param),
        "nonshppict" => ControlWord::NonShapePicture(param),
        "listid" => ControlWord::ListId(param_value),
        "listoverridetable" => ControlWord::ListOverrideTable,
        "listoverride" => ControlWord::ListOverride,
        "listoverridecount" => ControlWord::ListOverrideCount(param_value),
        "ls" => ControlWord::ListOverrideIndex(param_value),
        "ilvl" => ControlWord::ListLevelIndex(param_value),
        "lfolevel" => ControlWord::ListOverrideLevel,
        "listoverridestartat" => ControlWord::ListOverrideStartAt(param_bool),
        "listoverrideformat" => ControlWord::ListOverrideFormat(param_bool),
        "listlevel" => ControlWord::ListLevel,
        "levelnfc" | "levelnfcn" => ControlWord::ListLevelType(param_value),
        "leveljc" | "leveljcn" => ControlWord::ListLevelJustification(param_value),
        "levelfollow" => ControlWord::ListLevelFollow(param_value),
        "levelspace" => ControlWord::ListLevelSpace(param_value),
        "levelindent" => ControlWord::ListLevelIndent(param_value),
        "leveltext" => ControlWord::ListNumberText,
        "levelstartat" => ControlWord::ListLevelStartAt(param_value),
        "levelnumbers" => ControlWord::ListLevelNumbers,
        "levelpicture" => ControlWord::ListLevelPicture(param_value),
        "lvltentative" => ControlWord::ListLevelTentative,
        "levellegal" => ControlWord::ListLevelLegal(param_bool),
        "levelnorestart" => ControlWord::ListLevelNoRestart(param_bool),
        "levelold" => ControlWord::ListLevelOld(param_bool),
        "levelprev" => ControlWord::ListLevelPrevious(param_bool),
        "levelprevspace" => ControlWord::ListLevelPreviousSpace(param_bool),
        "leveltemplateid" => ControlWord::ListLevelTemplateId(param_value),

        // Sections
        "titlepg" => ControlWord::TitlePage,
        "endnhere" => ControlWord::SectionEndnoteHere,
        "outlinelevel" => ControlWord::OutlineLevel(param_value),
        "sectd" => ControlWord::SectionDefault,
        "sect" => ControlWord::Section,
        "sbk" => ControlWord::SectionBreak,
        "sbknone" => ControlWord::SectionContinuous,
        "sbkcol" => ControlWord::SectionColumn,
        "sbkpage" => ControlWord::SectionPage,
        "sbkeven" => ControlWord::SectionEvenPage,
        "sbkodd" => ControlWord::SectionOddPage,
        "paperw" | "pgwsxn" => ControlWord::PageWidth(param_value),
        "paperh" | "pghsxn" => ControlWord::PageHeight(param_value),
        "margl" | "marglsxn" => ControlWord::MarginLeft(param_value),
        "margr" | "margrsxn" => ControlWord::MarginRight(param_value),
        "margt" | "margtsxn" => ControlWord::MarginTop(param_value),
        "margb" | "margbsxn" => ControlWord::MarginBottom(param_value),
        "guttersxn" => ControlWord::MarginGutter(param_value),
        "binfsxn" => ControlWord::PaperSourceFirst(param),
        "binsxn" => ControlWord::PaperSourceOther(param),
        "facingp" => ControlWord::FacingPages(param_bool),
        "margmirror" => ControlWord::MirrorMargins(param),
        "gutter" => ControlWord::DocumentGutter(param),
        "headery" => ControlWord::HeaderDistance(param_value),
        "footery" => ControlWord::FooterDistance(param_value),
        "landscape" | "lndscpsxn" => ControlWord::Landscape,
        "cols" => ControlWord::Columns(param),
        "colsx" => ControlWord::ColumnSpace(param),
        "colno" => ControlWord::ColumnNumber(param),
        "colw" => ControlWord::ColumnWidth(param),
        "colsr" => ControlWord::ColumnSpaceRight(param),
        "linebetcol" => ControlWord::ColumnSeparator(param_bool),
        "pgnstarts" => ControlWord::PageNumberStart(param_value),
        "pgndec" => ControlWord::PageNumberFormat(crate::PageNumberFormat::Decimal),
        "pgnucrm" => ControlWord::PageNumberFormat(crate::PageNumberFormat::UpperRoman),
        "pgnlcrm" => ControlWord::PageNumberFormat(crate::PageNumberFormat::LowerRoman),
        "pgnucltr" => ControlWord::PageNumberFormat(crate::PageNumberFormat::UpperLetter),
        "pgnlcltr" => ControlWord::PageNumberFormat(crate::PageNumberFormat::LowerLetter),
        "pgnbidia" => ControlWord::PageNumberFormat(crate::PageNumberFormat::BidiAlphabetic),
        "pgnbidib" => ControlWord::PageNumberFormat(crate::PageNumberFormat::BidiAbjad),
        "pgnchosung" => ControlWord::PageNumberFormat(crate::PageNumberFormat::KoreanChosung),
        "pgncnum" => ControlWord::PageNumberFormat(crate::PageNumberFormat::Circle),
        "pgndbnum" => ControlWord::PageNumberFormat(crate::PageNumberFormat::KanjiDigitless),
        "pgndbnumd" => ControlWord::PageNumberFormat(crate::PageNumberFormat::KanjiWithDigit),
        "pgndbnumt" => ControlWord::PageNumberFormat(crate::PageNumberFormat::KanjiThree),
        "pgndbnumk" => ControlWord::PageNumberFormat(crate::PageNumberFormat::KanjiFour),
        "pgndecd" => ControlWord::PageNumberFormat(crate::PageNumberFormat::DoubleDecimal),
        "pgnganada" => ControlWord::PageNumberFormat(crate::PageNumberFormat::KoreanGanada),
        "pgngbnum" => ControlWord::PageNumberFormat(crate::PageNumberFormat::ChineseOne),
        "pgngbnumd" => ControlWord::PageNumberFormat(crate::PageNumberFormat::ChineseTwo),
        "pgngbnuml" => ControlWord::PageNumberFormat(crate::PageNumberFormat::ChineseThree),
        "pgngbnumk" => ControlWord::PageNumberFormat(crate::PageNumberFormat::ChineseFour),
        "pgnhindia" => ControlWord::PageNumberFormat(crate::PageNumberFormat::HindiVowels),
        "pgnhindib" => ControlWord::PageNumberFormat(crate::PageNumberFormat::HindiConsonants),
        "pgnhindic" => ControlWord::PageNumberFormat(crate::PageNumberFormat::HindiNumbers),
        "pgnhindid" => ControlWord::PageNumberFormat(crate::PageNumberFormat::HindiDescriptive),
        "pgnthaia" => ControlWord::PageNumberFormat(crate::PageNumberFormat::ThaiLetters),
        "pgnthaib" => ControlWord::PageNumberFormat(crate::PageNumberFormat::ThaiNumbers),
        "pgnthaic" => ControlWord::PageNumberFormat(crate::PageNumberFormat::ThaiDescriptive),
        "pgnvieta" => ControlWord::PageNumberFormat(crate::PageNumberFormat::VietnameseCardinal),
        "pgnzodiac" => ControlWord::PageNumberFormat(crate::PageNumberFormat::ZodiacOne),
        "pgnzodiacd" => ControlWord::PageNumberFormat(crate::PageNumberFormat::ZodiacTwo),
        "pgnzodiacl" => ControlWord::PageNumberFormat(crate::PageNumberFormat::ZodiacThree),
        "pgnrestart" => ControlWord::PageNumberRestart(crate::PageNumberRestart::Restart),
        "pgncont" => ControlWord::PageNumberRestart(crate::PageNumberRestart::Continuous),
        "pgnx" => ControlWord::PageNumberOffsetX(param_value),
        "pgny" => ControlWord::PageNumberOffsetY(param_value),
        "pgnhn" => ControlWord::PageNumberHeadingLevel(param),
        "pgnhnsh" => {
            ControlWord::PageNumberHeadingSeparator(crate::PageNumberHeadingSeparator::Hyphen)
        },
        "pgnhnsp" => {
            ControlWord::PageNumberHeadingSeparator(crate::PageNumberHeadingSeparator::Period)
        },
        "pgnhnsc" => {
            ControlWord::PageNumberHeadingSeparator(crate::PageNumberHeadingSeparator::Colon)
        },
        "pgnhnsm" => {
            ControlWord::PageNumberHeadingSeparator(crate::PageNumberHeadingSeparator::EmDash)
        },
        "pgnhnsn" => {
            ControlWord::PageNumberHeadingSeparator(crate::PageNumberHeadingSeparator::EnDash)
        },
        "sectlinegrid" => ControlWord::SectionLineGrid(param),
        "sectspecifyl" => {
            ControlWord::SectionDocumentGrid(crate::SectionDocumentGridType::LinesAndCharacters)
        },
        "sectspecifycl" => {
            ControlWord::SectionDocumentGrid(crate::SectionDocumentGridType::CharactersOnly)
        },
        "sectspecifygen" => {
            ControlWord::SectionDocumentGrid(crate::SectionDocumentGridType::Default)
        },
        "prauth" => ControlWord::ParagraphRevisionAuthor(param_value),
        "prdate" => ControlWord::ParagraphRevisionDate(param_value),
        "srauth" => ControlWord::SectionRevisionAuthor(param_value),
        "srdate" => ControlWord::SectionRevisionDate(param_value),
        "trauth" => ControlWord::TableRowRevisionAuthor(param_value),
        "trdate" => ControlWord::TableRowRevisionDate(param_value),
        "clins" => ControlWord::CellRevisionMark(crate::CellRevisionKind::Inserted),
        "clinsauth" => {
            ControlWord::CellRevisionAuthor(crate::CellRevisionKind::Inserted, param_value)
        },
        "clinsdttm" => {
            ControlWord::CellRevisionDate(crate::CellRevisionKind::Inserted, param_value)
        },
        "cldel" => ControlWord::CellRevisionMark(crate::CellRevisionKind::Deleted),
        "cldelauth" => {
            ControlWord::CellRevisionAuthor(crate::CellRevisionKind::Deleted, param_value)
        },
        "cldeldttm" => ControlWord::CellRevisionDate(crate::CellRevisionKind::Deleted, param_value),
        "clmrgd" => ControlWord::CellRevisionMark(crate::CellRevisionKind::MergeDeleted),
        "clmrgdauth" => {
            ControlWord::CellRevisionAuthor(crate::CellRevisionKind::MergeDeleted, param_value)
        },
        "clmrgddttm" => {
            ControlWord::CellRevisionDate(crate::CellRevisionKind::MergeDeleted, param_value)
        },
        "vertalt" => ControlWord::VerticalAlignTop,
        "vertalc" => ControlWord::VerticalAlignCenter,
        "vertalj" => ControlWord::VerticalAlignJustify,
        "vertalb" => ControlWord::VerticalAlignBottom,
        "linemod" => ControlWord::LineNumbering(param),
        "linex" => ControlWord::LineNumberDistance(param),
        "linestarts" => ControlWord::LineNumberStart(param),
        "linerestart" => ControlWord::LineNumberRestartSection,
        "lineppage" => ControlWord::LineNumberRestartPage,
        "linecont" => ControlWord::LineNumberContinuous,
        "header" => ControlWord::Header,
        "headerf" => ControlWord::HeaderFirst,
        "headerl" => ControlWord::HeaderLeft,
        "headerr" => ControlWord::HeaderRight,
        "footer" => ControlWord::Footer,
        "footerf" => ControlWord::FooterFirst,
        "footerl" => ControlWord::FooterLeft,
        "footerr" => ControlWord::FooterRight,

        // Footnotes and endnotes
        "footnote" => ControlWord::Footnote,
        "endnote" => ControlWord::Endnote,
        "chftn" => ControlWord::FootnoteNumber(param_value),
        "fet" => ControlWord::NoteKinds(param_value),
        "endnotes" => ControlWord::FootnotePlacement(crate::NotePlacement::EndOfSection),
        "enddoc" => ControlWord::FootnotePlacement(crate::NotePlacement::EndOfDocument),
        "ftntj" => ControlWord::FootnotePlacement(crate::NotePlacement::BeneathText),
        "ftnbj" => ControlWord::FootnotePlacement(crate::NotePlacement::BottomOfPage),
        "aendnotes" => ControlWord::EndnotePlacement(crate::NotePlacement::EndOfSection),
        "aenddoc" => ControlWord::EndnotePlacement(crate::NotePlacement::EndOfDocument),
        "aftntj" => ControlWord::EndnotePlacement(crate::NotePlacement::BeneathText),
        "aftnbj" => ControlWord::EndnotePlacement(crate::NotePlacement::BottomOfPage),
        "ftnstart" => ControlWord::FootnoteStart(param_value),
        "aftnstart" => ControlWord::EndnoteStart(param_value),
        "ftnrstcont" => ControlWord::FootnoteRestart(crate::FootnoteRestart::Continuous),
        "ftnrestart" => ControlWord::FootnoteRestart(crate::FootnoteRestart::EachSection),
        "ftnrstpg" => ControlWord::FootnoteRestart(crate::FootnoteRestart::EachPage),
        "aftnrstcont" => ControlWord::EndnoteRestart(crate::EndnoteRestart::Continuous),
        "aftnrestart" => ControlWord::EndnoteRestart(crate::EndnoteRestart::EachSection),
        "ftnnar" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::Arabic),
        "ftnnalc" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::LowercaseLetter),
        "ftnnauc" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::UppercaseLetter),
        "ftnnrlc" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::LowercaseRoman),
        "ftnnruc" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::UppercaseRoman),
        "ftnnchi" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::Chicago),
        "ftnnchosung" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::KoreanChosung),
        "ftnncnum" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::Circle),
        "ftnndbnum" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::KanjiDigitless),
        "ftnndbnumd" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::KanjiWithDigit),
        "ftnndbnumt" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::KanjiThree),
        "ftnndbnumk" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::KanjiFour),
        "ftnndbar" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::DoubleByte),
        "ftnnganada" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::KoreanGanada),
        "ftnngbnum" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::ChineseOne),
        "ftnngbnumd" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::ChineseTwo),
        "ftnngbnuml" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::ChineseThree),
        "ftnngbnumk" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::ChineseFour),
        "ftnnzodiac" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::ZodiacOne),
        "ftnnzodiacd" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::ZodiacTwo),
        "ftnnzodiacl" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::ZodiacThree),
        "aftnnar" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::Arabic),
        "aftnnalc" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::LowercaseLetter),
        "aftnnauc" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::UppercaseLetter),
        "aftnnrlc" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::LowercaseRoman),
        "aftnnruc" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::UppercaseRoman),
        "aftnnchi" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::Chicago),
        "aftnnchosung" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::KoreanChosung),
        "aftnncnum" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::Circle),
        "aftnndbnum" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::KanjiDigitless),
        "aftnndbnumd" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::KanjiWithDigit),
        "aftnndbnumt" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::KanjiThree),
        "aftnndbnumk" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::KanjiFour),
        "aftnndbar" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::DoubleByte),
        "aftnnganada" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::KoreanGanada),
        "aftnngbnum" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::ChineseOne),
        "aftnngbnumd" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::ChineseTwo),
        "aftnngbnuml" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::ChineseThree),
        "aftnngbnumk" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::ChineseFour),
        "aftnnzodiac" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::ZodiacOne),
        "aftnnzodiacd" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::ZodiacTwo),
        "aftnnzodiacl" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::ZodiacThree),
        "sftntj" => {
            ControlWord::SectionFootnotePlacement(crate::SectionFootnotePlacement::BeneathText)
        },
        "sftnbj" => {
            ControlWord::SectionFootnotePlacement(crate::SectionFootnotePlacement::BottomOfPage)
        },
        "sftnstart" => ControlWord::SectionFootnoteStart(param_value),
        "saftnstart" => ControlWord::SectionEndnoteStart(param_value),
        "sftnrstcont" => ControlWord::SectionFootnoteRestart(crate::FootnoteRestart::Continuous),
        "sftnrestart" => ControlWord::SectionFootnoteRestart(crate::FootnoteRestart::EachSection),
        "sftnrstpg" => ControlWord::SectionFootnoteRestart(crate::FootnoteRestart::EachPage),
        "saftnrstcont" => ControlWord::SectionEndnoteRestart(crate::EndnoteRestart::Continuous),
        "saftnrestart" => ControlWord::SectionEndnoteRestart(crate::EndnoteRestart::EachSection),
        "sftnnar" => ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::Arabic),
        "sftnnalc" => {
            ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::LowercaseLetter)
        },
        "sftnnauc" => {
            ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::UppercaseLetter)
        },
        "sftnnrlc" => {
            ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::LowercaseRoman)
        },
        "sftnnruc" => {
            ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::UppercaseRoman)
        },
        "sftnnchi" => ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::Chicago),
        "sftnnchosung" => {
            ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::KoreanChosung)
        },
        "sftnncnum" => ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::Circle),
        "sftnndbnum" => {
            ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::KanjiDigitless)
        },
        "sftnndbnumd" => {
            ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::KanjiWithDigit)
        },
        "sftnndbnumt" => {
            ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::KanjiThree)
        },
        "sftnndbnumk" => {
            ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::KanjiFour)
        },
        "sftnndbar" => ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::DoubleByte),
        "sftnnganada" => {
            ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::KoreanGanada)
        },
        "sftnngbnum" => {
            ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::ChineseOne)
        },
        "sftnngbnumd" => {
            ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::ChineseTwo)
        },
        "sftnngbnuml" => {
            ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::ChineseThree)
        },
        "sftnngbnumk" => {
            ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::ChineseFour)
        },
        "sftnnzodiac" => {
            ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::ZodiacOne)
        },
        "sftnnzodiacd" => {
            ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::ZodiacTwo)
        },
        "sftnnzodiacl" => {
            ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::ZodiacThree)
        },
        "saftnnar" => ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::Arabic),
        "saftnnalc" => {
            ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::LowercaseLetter)
        },
        "saftnnauc" => {
            ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::UppercaseLetter)
        },
        "saftnnrlc" => {
            ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::LowercaseRoman)
        },
        "saftnnruc" => {
            ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::UppercaseRoman)
        },
        "saftnnchi" => ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::Chicago),
        "saftnnchosung" => {
            ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::KoreanChosung)
        },
        "saftnncnum" => ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::Circle),
        "saftnndbnum" => {
            ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::KanjiDigitless)
        },
        "saftnndbnumd" => {
            ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::KanjiWithDigit)
        },
        "saftnndbnumt" => {
            ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::KanjiThree)
        },
        "saftnndbnumk" => {
            ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::KanjiFour)
        },
        "saftnndbar" => ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::DoubleByte),
        "saftnnganada" => {
            ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::KoreanGanada)
        },
        "saftnngbnum" => {
            ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::ChineseOne)
        },
        "saftnngbnumd" => {
            ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::ChineseTwo)
        },
        "saftnngbnuml" => {
            ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::ChineseThree)
        },
        "saftnngbnumk" => {
            ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::ChineseFour)
        },
        "saftnnzodiac" => {
            ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::ZodiacOne)
        },
        "saftnnzodiacd" => {
            ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::ZodiacTwo)
        },
        "saftnnzodiacl" => {
            ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::ZodiacThree)
        },
        "ftnsep" => ControlWord::FootnoteSeparator,
        "ftnsepc" => ControlWord::FootnoteContinuationSeparator,
        "ftncn" => ControlWord::FootnoteContinuationNotice,
        "aftnsep" => ControlWord::EndnoteSeparator,
        "aftnsepc" => ControlWord::EndnoteContinuationSeparator,
        "aftncn" => ControlWord::EndnoteContinuationNotice,
        "chftnsep" => ControlWord::NoteSeparatorCharacter,
        "chftnsepc" => ControlWord::NoteContinuationSeparatorCharacter,

        // Bookmarks
        "bkmkstart" => ControlWord::BookmarkStart,
        "bkmkend" => ControlWord::BookmarkEnd,
        "bkmkcolf" => ControlWord::BookmarkFirstColumn(param_value),
        "bkmkcoll" => ControlWord::BookmarkLastColumn(param_value),
        "bkmkpub" => ControlWord::BookmarkPublic,

        // Annotations
        "atn" | "annotation" => ControlWord::Annotation,
        "atndate" => ControlWord::AnnotationDate,
        "atnauthor" => ControlWord::AnnotationAuthor,
        "atnid" => ControlWord::AnnotationInitials,
        "atnref" => ControlWord::AnnotationReference,
        "atnparent" => ControlWord::AnnotationParent,
        "atrfstart" => ControlWord::AnnotationRangeStart,
        "atrfend" => ControlWord::AnnotationRangeEnd,
        "atnicn" => ControlWord::AnnotationIcon,
        "atntime" => ControlWord::AnnotationTime,
        "chatn" => ControlWord::AnnotationMark,

        // Tracked revisions
        "revtbl" => ControlWord::RevisionTable,
        "revised" => ControlWord::Revised(param_bool),
        "deleted" => ControlWord::Deleted(param_bool),
        "revauth" => ControlWord::RevisionAuthor(param_value),
        "revauthdel" => ControlWord::DeletedRevisionAuthor(param_value),
        "revdttm" => ControlWord::RevisionDate(param_value),
        "revdttmdel" => ControlWord::DeletedRevisionDate(param_value),
        "revprop" => ControlWord::RevisionProperty(param_value),

        // Shapes
        "shp" => ControlWord::Shape(param),
        "shpgrp" => ControlWord::ShapeGroup(param),
        "shpinst" => param.map_or(ControlWord::ShapeInstance, ControlWord::ShapeType),
        "shpleft" => ControlWord::ShapeLeft(param_value),
        "shptop" => ControlWord::ShapeTop(param_value),
        "shpright" => ControlWord::ShapeRight(param_value),
        "shpbottom" => ControlWord::ShapeBottom(param_value),
        "shpwidth" => ControlWord::ShapeWidth(param_value),
        "shpheight" => ControlWord::ShapeHeight(param_value),
        "shprotation" => ControlWord::ShapeRotation(param_value),
        "shpz" => ControlWord::ShapeZOrder(param_value),
        "shpwr" => ControlWord::ShapeWrap(param_value),
        "shpfblwtxt" => ControlWord::ShapeBelowText(param_bool),
        "shplockanchor" => ControlWord::ShapeLockAnchor,
        "sp" => {
            if param.is_none() {
                ControlWord::ShapeProperty
            } else {
                ControlWord::InvalidPictureShapePropertyParameter
            }
        },
        "sn" => {
            if param.is_none() {
                ControlWord::ShapePropertyName
            } else {
                ControlWord::InvalidPictureShapePropertyParameter
            }
        },
        "sv" => {
            if param.is_none() {
                ControlWord::ShapePropertyValue
            } else {
                ControlWord::InvalidPictureShapePropertyParameter
            }
        },
        "svb" => ControlWord::ShapeBinaryValue(param),
        "hsv" => ControlWord::ShapeThemeValue(param),
        "caccentone" => ControlWord::ShapeThemeColor(crate::ShapeThemeColor::Accent1, param),
        "caccenttwo" => ControlWord::ShapeThemeColor(crate::ShapeThemeColor::Accent2, param),
        "caccentthree" => ControlWord::ShapeThemeColor(crate::ShapeThemeColor::Accent3, param),
        "caccentfour" => ControlWord::ShapeThemeColor(crate::ShapeThemeColor::Accent4, param),
        "caccentfive" => ControlWord::ShapeThemeColor(crate::ShapeThemeColor::Accent5, param),
        "caccentsix" => ControlWord::ShapeThemeColor(crate::ShapeThemeColor::Accent6, param),
        "cbackgroundone" => {
            ControlWord::ShapeThemeColor(crate::ShapeThemeColor::Background1, param)
        },
        "cbackgroundtwo" => {
            ControlWord::ShapeThemeColor(crate::ShapeThemeColor::Background2, param)
        },
        "ctextone" => ControlWord::ShapeThemeColor(crate::ShapeThemeColor::Text1, param),
        "ctexttwo" => ControlWord::ShapeThemeColor(crate::ShapeThemeColor::Text2, param),
        "ctint" => ControlWord::ShapeThemeTint(param),
        "cshade" => ControlWord::ShapeThemeShade(param),
        "shprslt" => ControlWord::ShapeResult(param),
        "shptxt" => ControlWord::ShapeText(param),
        "do" => ControlWord::LegacyDrawingObject,
        "dptxbx" => ControlWord::LegacyTextBox,
        "dptxbxtext" => ControlWord::LegacyTextBoxText,
        "dobxpage" => ControlWord::LegacyAnchorXPage,
        "dobxmargin" => ControlWord::LegacyAnchorXMargin,
        "dobxcolumn" => ControlWord::LegacyAnchorXColumn,
        "dobypage" => ControlWord::LegacyAnchorYPage,
        "dobymargin" => ControlWord::LegacyAnchorYMargin,
        "dobypara" => ControlWord::LegacyAnchorYParagraph,
        "dodhgt" | "dolhgt" => ControlWord::LegacyDrawingHeight(param_value),
        "dptxbxmar" => ControlWord::LegacyTextBoxMargin(param_value),
        "dpx" => ControlWord::LegacyDrawingX(param_value),
        "dpy" => ControlWord::LegacyDrawingY(param_value),
        "dpxsize" => ControlWord::LegacyDrawingWidth(param_value),
        "dpysize" => ControlWord::LegacyDrawingHeightSize(param_value),
        "dptxlrtb" => ControlWord::LegacyTextLeftRightTopBottom,
        "dptxlrtbv" => ControlWord::LegacyTextLeftRightTopBottomVertical,
        "dptxtbrl" => ControlWord::LegacyTextTopBottomRightLeft,
        "dptxtbrlv" => ControlWord::LegacyTextTopBottomRightLeftVertical,
        "dptxbtlr" => ControlWord::LegacyTextBottomTopLeftRight,
        "dolock" => ControlWord::LegacyDrawingLock,
        "dpgroup" => ControlWord::LegacyDrawingGroup,
        "dpcount" => ControlWord::LegacyDrawingCount(param_value),
        "dppolycount" => ControlWord::LegacyDrawingCount(param_value),
        "dpendgroup" => ControlWord::LegacyDrawingEndGroup,
        "dparc" => ControlWord::LegacyDrawingArc,
        "dpcallout" => ControlWord::LegacyDrawingCallout,
        "dpellipse" => ControlWord::LegacyDrawingEllipse,
        "dpline" => ControlWord::LegacyDrawingLine,
        "dppolygon" => ControlWord::LegacyDrawingPolygon,
        "dppolyline" => ControlWord::LegacyDrawingPolyline,
        "dprect" => ControlWord::LegacyDrawingRectangle,
        "dproundr" => ControlWord::LegacyDrawingRoundRectangle,
        "dpptx" => ControlWord::LegacyDrawingPointX(param_value),
        "dppty" => ControlWord::LegacyDrawingPointY(param_value),
        "dparcflipx" => ControlWord::LegacyDrawingArcFlipX,
        "dparcflipy" => ControlWord::LegacyDrawingArcFlipY,
        "dpcotright" => ControlWord::LegacyCalloutType(crate::LegacyCalloutType::RightAngle),
        "dpcotsingle" => ControlWord::LegacyCalloutType(crate::LegacyCalloutType::Single),
        "dpcotdouble" => ControlWord::LegacyCalloutType(crate::LegacyCalloutType::Double),
        "dpcottriple" => ControlWord::LegacyCalloutType(crate::LegacyCalloutType::Triple),
        "dpcoa" => ControlWord::LegacyCalloutAngle(param_value),
        "dpcoaccent" => ControlWord::LegacyCalloutAccent,
        "dpcosmarta" => ControlWord::LegacyCalloutSmartAttach,
        "dpcobestfit" => ControlWord::LegacyCalloutBestFit,
        "dpcominusx" => ControlWord::LegacyCalloutMinusX,
        "dpcominusy" => ControlWord::LegacyCalloutMinusY,
        "dpcoborder" => ControlWord::LegacyCalloutBorder,
        "dpcodtop" => ControlWord::LegacyCalloutAttachment(crate::LegacyCalloutAttachment::Top),
        "dpcodcenter" => {
            ControlWord::LegacyCalloutAttachment(crate::LegacyCalloutAttachment::Center)
        },
        "dpcodbottom" => {
            ControlWord::LegacyCalloutAttachment(crate::LegacyCalloutAttachment::Bottom)
        },
        "dpcodabs" => {
            ControlWord::LegacyCalloutAttachment(crate::LegacyCalloutAttachment::Absolute)
        },
        "dpcodescent" => ControlWord::LegacyCalloutDescent(param_value),
        "dpcooffset" => ControlWord::LegacyCalloutOffset(param_value),
        "dpcolength" => ControlWord::LegacyCalloutLength(param_value),
        "dplinesolid" => ControlWord::LegacyDrawingLineStyle(crate::LegacyDrawingLineStyle::Solid),
        "dplinehollow" => {
            ControlWord::LegacyDrawingLineStyle(crate::LegacyDrawingLineStyle::Hollow)
        },
        "dplinedash" => ControlWord::LegacyDrawingLineStyle(crate::LegacyDrawingLineStyle::Dashed),
        "dplinedot" => ControlWord::LegacyDrawingLineStyle(crate::LegacyDrawingLineStyle::Dotted),
        "dplinedado" => ControlWord::LegacyDrawingLineStyle(crate::LegacyDrawingLineStyle::DashDot),
        "dplinedadodo" => {
            ControlWord::LegacyDrawingLineStyle(crate::LegacyDrawingLineStyle::DashDotDot)
        },
        "dplinegray" => ControlWord::LegacyDrawingLineGray(param_value),
        "dplinecor" => ControlWord::LegacyDrawingLineRed(param_value),
        "dplinecog" => ControlWord::LegacyDrawingLineGreen(param_value),
        "dplinecob" => ControlWord::LegacyDrawingLineBlue(param_value),
        "dplinepal" => ControlWord::LegacyDrawingLinePalette,
        "dplinew" => ControlWord::LegacyDrawingLineWidth(param_value),
        "dpfillfggray" => ControlWord::LegacyDrawingFillForegroundGray(param_value),
        "dpfillfgcr" => ControlWord::LegacyDrawingFillForegroundRed(param_value),
        "dpfillfgcg" => ControlWord::LegacyDrawingFillForegroundGreen(param_value),
        "dpfillfgcb" => ControlWord::LegacyDrawingFillForegroundBlue(param_value),
        "dpfillfgpal" => ControlWord::LegacyDrawingFillForegroundPalette,
        "dpfillbggray" => ControlWord::LegacyDrawingFillBackgroundGray(param_value),
        "dpfillbgcr" => ControlWord::LegacyDrawingFillBackgroundRed(param_value),
        "dpfillbgcg" => ControlWord::LegacyDrawingFillBackgroundGreen(param_value),
        "dpfillbgcb" => ControlWord::LegacyDrawingFillBackgroundBlue(param_value),
        "dpfillbgpal" => ControlWord::LegacyDrawingFillBackgroundPalette,
        "dpfillpat" => ControlWord::LegacyDrawingFillPattern(param_value),
        "dpastartsol" => {
            ControlWord::LegacyDrawingStartArrowFill(crate::LegacyDrawingArrowFill::Solid)
        },
        "dpastarthol" => {
            ControlWord::LegacyDrawingStartArrowFill(crate::LegacyDrawingArrowFill::Hollow)
        },
        "dpastartl" => ControlWord::LegacyDrawingStartArrowLength(param_value),
        "dpastartw" => ControlWord::LegacyDrawingStartArrowWidth(param_value),
        "dpaendsol" => ControlWord::LegacyDrawingEndArrowFill(crate::LegacyDrawingArrowFill::Solid),
        "dpaendhol" => {
            ControlWord::LegacyDrawingEndArrowFill(crate::LegacyDrawingArrowFill::Hollow)
        },
        "dpaendl" => ControlWord::LegacyDrawingEndArrowLength(param_value),
        "dpaendw" => ControlWord::LegacyDrawingEndArrowWidth(param_value),
        "dpshadow" => ControlWord::LegacyDrawingShadow,
        "dpshadx" => ControlWord::LegacyDrawingShadowX(param_value),
        "dpshady" => ControlWord::LegacyDrawingShadowY(param_value),
        "background" => ControlWord::BackgroundDestination(param),

        // Document info
        "title" => ControlWord::Title,
        "subject" => ControlWord::Subject,
        "author" => ControlWord::Author,
        "manager" => ControlWord::Manager,
        "company" => ControlWord::Company,
        "operator" => ControlWord::Operator,
        "category" => ControlWord::Category,
        "keywords" => ControlWord::Keywords,
        "comment" => ControlWord::Comment,
        "doccomm" => ControlWord::DocComment,
        "hlinkbase" => ControlWord::HyperlinkBase,
        "creatim" => ControlWord::CreationTime,
        "revtim" => ControlWord::RevisionTime,
        "printim" => ControlWord::PrintTime,
        "buptim" => ControlWord::BackupTime,
        "version" => ControlWord::InfoVersion(param_value),
        "vern" => ControlWord::InfoRevision(param_value),
        "edmins" => ControlWord::EditingTime(param_value),
        "nofpages" => ControlWord::NumberOfPages(param_value),
        "nofwords" => ControlWord::NumberOfWords(param_value),
        "nofchars" => ControlWord::NumberOfCharacters(param_value),
        "nofcharsws" => ControlWord::NumberOfCharactersWithSpaces(param_value),
        "id" => ControlWord::DocumentId(param_value),
        "yr" => ControlWord::Year(param_value),
        "mo" => ControlWord::Month(param_value),
        "dy" => ControlWord::Day(param_value),
        "hr" => ControlWord::Hour(param_value),
        "min" => ControlWord::Minute(param_value),
        "sec" => ControlWord::Second(param_value),

        // Unicode
        "u" => ControlWord::Unicode(param.ok_or_else(|| {
            RtfError::InvalidUnicode("RTF \\u requires a numeric parameter".to_string())
        })?),
        "uc" => ControlWord::UnicodeSkip(param_value),

        // Special
        "tab" => ControlWord::Tab,
        "line" => ControlWord::Line,
        "emdash" => ControlWord::EmDash,
        "endash" => ControlWord::EnDash,
        "emspace" => ControlWord::EmSpace,
        "enspace" => ControlWord::EnSpace,
        "qmspace" => ControlWord::QuarterEmSpace,
        "bullet" => ControlWord::Bullet,
        "ltrmark" => ControlWord::LeftToRightMark,
        "rtlmark" => ControlWord::RightToLeftMark,
        "zwj" => ControlWord::ZeroWidthJoiner,
        "zwnj" => ControlWord::ZeroWidthNonJoiner,
        "zwbo" => ControlWord::ZeroWidthBreakOpportunity,
        "zwnbo" => ControlWord::ZeroWidthNoBreakOpportunity,
        "chdate" => ControlWord::CurrentDate,
        "chdpl" => ControlWord::CurrentDateLong,
        "chdpa" => ControlWord::CurrentDateAbbreviated,
        "chtime" => ControlWord::CurrentTime,

        // Binary data
        "bin" => ControlWord::Binary(param_value),

        // Unknown
        _ => ControlWord::Unknown(word, param),
    };

    Ok(control)
}
