use super::ControlWord;
use super::repair::{binary, flag, toggle, toggled_selector, value};
use crate::codec::error::{RtfError, RtfResult};
use crate::leniency::ParameterDefect;

/// `\fs` default when the parameter is missing: 12 points, as in
/// LibreOffice's keyword table.
const DEFAULT_FONT_SIZE_HALF_POINTS: i32 = 24;

fn reject_numeric_parameter(word: &str, param: Option<i32>) -> RtfResult<()> {
    if param.is_some() {
        return Err(RtfError::MalformedDocument(format!(
            "RTF \\{word} does not accept a numeric parameter"
        )));
    }
    Ok(())
}

/// Match a control-word spelling and optional parameter to its typed token.
///
/// This is the RTF specification's flat control-word dispatch table. Keeping
/// the entries together makes duplicate and missing mappings auditable.
///
/// Controls whose parameter shape the parser validates strictly pass their
/// parameter through the [`super::repair`] helpers, which apply the
/// Word-compatible interpretation of a malformed shape and report it through
/// `repair`.
#[allow(
    clippy::too_many_lines,
    reason = "the RTF specification's flat control-word dispatch table is inherently long; keeping entries together aids auditing"
)]
pub(in crate::codec::lexer::codec) fn match_control_word<'a>(
    word: &'a str,
    param: Option<i32>,
    repair: &mut Option<ParameterDefect>,
) -> RtfResult<ControlWord<'a>> {
    let param_value = param.unwrap_or(1);
    let param_bool = param.unwrap_or(1) != 0;

    #[allow(
        clippy::match_same_arms,
        reason = "distinct RTF control words may map to equal tokens; keeping one arm per spelling keeps the table auditable"
    )]
    let control = match word {
        // Document
        "rtf" => ControlWord::Rtf(param.ok_or_else(|| {
            RtfError::MalformedDocument(
                "RTF \\rtf control requires a version parameter".to_string(),
            )
        })?),
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
        "deff" => ControlWord::DefaultFont(value(param, 0, repair)),
        "adeff" => ControlWord::AssociatedDefaultFont(value(param, 0, repair)),
        "stshfbi" => ControlWord::StylesheetDefaultBidiFont(value(param, 0, repair)),
        "stshfdbch" => ControlWord::StylesheetDefaultDoubleByteFont(value(param, 0, repair)),
        "stshfhich" => ControlWord::StylesheetDefaultHighAnsiFont(value(param, 0, repair)),
        "stshfloch" => ControlWord::StylesheetDefaultLowAnsiFont(value(param, 0, repair)),
        "pn" => ControlWord::LegacyParagraphNumbering(param),
        "pnlvl" => ControlWord::LegacyNumberingLevel(param),
        "pnlvlblt" => ControlWord::LegacyNumberingLevelBullet(param),
        "pnlvlbody" => ControlWord::LegacyNumberingLevelBody(param),
        "pnlvlcont" => ControlWord::LegacyNumberingLevelContinue(param),
        "pnseclvl" => ControlWord::LegacySectionNumberingLevel(param_value),
        "pndec" => ControlWord::LegacyNumberingDecimal(flag(param, repair)),
        "pnucrm" => ControlWord::LegacyNumberingUpperRoman(flag(param, repair)),
        "pnlcrm" => ControlWord::LegacyNumberingLowerRoman(flag(param, repair)),
        "pnucltr" => ControlWord::LegacyNumberingUpperLetter(flag(param, repair)),
        "pnlcltr" => ControlWord::LegacyNumberingLowerLetter(flag(param, repair)),
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
        "pnstart" => ControlWord::LegacyNumberingStart(value(param, 0, repair)),
        "pnindent" => ControlWord::LegacyNumberingIndent(value(param, 0, repair)),
        "pnsp" => ControlWord::LegacyNumberingSpace(value(param, 0, repair)),
        "pnhang" => ControlWord::LegacyNumberingHanging(flag(param, repair)),
        "pnprev" => ControlWord::LegacyNumberingPrevious(flag(param, repair)),
        "pnacross" => ControlWord::LegacyNumberingAcross(param),
        "pnnumonce" => ControlWord::LegacyNumberingOnce(param),
        "pnrestart" => ControlWord::LegacyNumberingRestart(param),
        "pnbidia" => ControlWord::LegacyNumberingBidiA(param),
        "pnbidib" => ControlWord::LegacyNumberingBidiB(param),
        "pnql" => ControlWord::LegacyNumberingAlignLeft(flag(param, repair)),
        "pnqc" => ControlWord::LegacyNumberingAlignCenter(flag(param, repair)),
        "pnqr" => ControlWord::LegacyNumberingAlignRight(flag(param, repair)),
        "pnf" => ControlWord::LegacyNumberingFont(value(param, 0, repair)),
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
        "itap" => ControlWord::TableNestingLevel(value(param, 1, repair)),

        // Stylesheet entries and metadata
        "s" => ControlWord::ParagraphStyle(value(param, 0, repair)),
        "cs" => ControlWord::CharacterStyle(value(param, 0, repair)),
        "ds" => ControlWord::SectionStyle(value(param, 0, repair)),
        "ts" => ControlWord::TableStyle(value(param, 0, repair)),
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
        "picbpp" => ControlWord::PictureBitsPerPixel(value(param, 1, repair)),
        "piccropl" => ControlWord::PictureCropLeft(value(param, 0, repair)),
        "piccropr" => ControlWord::PictureCropRight(value(param, 0, repair)),
        "piccropt" => ControlWord::PictureCropTop(value(param, 0, repair)),
        "piccropb" => ControlWord::PictureCropBottom(value(param, 0, repair)),
        "wbmbitspixel" => ControlWord::WindowsBitmapBitsPerPixel(value(param, 1, repair)),
        "wbmplanes" => ControlWord::WindowsBitmapPlanes(value(param, 1, repair)),
        "wbmwidthbytes" => ControlWord::WindowsBitmapWidthBytes(value(param, 0, repair)),
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
        "xmlnstbl" => {
            reject_numeric_parameter("xmlnstbl", flag(param, repair))?;
            ControlWord::XmlNamespaceTable
        },
        "xmlns" => ControlWord::XmlNamespace(param.ok_or_else(|| {
            RtfError::MalformedDocument(
                "RTF \\xmlns control requires a namespace parameter".to_string(),
            )
        })?),
        "xmlopen" => {
            reject_numeric_parameter("xmlopen", flag(param, repair))?;
            ControlWord::XmlOpen
        },
        "xmlclose" => {
            reject_numeric_parameter("xmlclose", flag(param, repair))?;
            ControlWord::XmlClose
        },
        "xmlattr" => match param {
            Some(value) => ControlWord::XmlAttribute(value),
            None => ControlWord::XmlAttributeGroup,
        },
        "xmlattrns" => ControlWord::XmlAttributeNamespace(param.ok_or_else(|| {
            RtfError::MalformedDocument(
                "RTF \\xmlattrns control requires a namespace parameter".to_string(),
            )
        })?),
        "xmlattrname" => {
            reject_numeric_parameter("xmlattrname", flag(param, repair))?;
            ControlWord::XmlAttributeName
        },
        "xmlattrvalue" => {
            reject_numeric_parameter("xmlattrvalue", flag(param, repair))?;
            ControlWord::XmlAttributeValue
        },
        "factoidname" => {
            reject_numeric_parameter("factoidname", flag(param, repair))?;
            ControlWord::FactoidName
        },
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
        "ebcstart" => ControlWord::EditableRegionStart(flag(param, repair)),
        "ebcend" => ControlWord::EditableRegionEnd(flag(param, repair)),
        "fchars" => ControlWord::KinsokuFollowing,
        "lchars" => ControlWord::KinsokuLeading,
        "ksulang" => ControlWord::KinsokuLanguage(value(param, 0, repair)),
        "hl" => ControlWord::ShapeHyperlink,
        "hlloc" => ControlWord::ShapeHyperlinkLocation,
        "hlsrc" => ControlWord::ShapeHyperlinkSource,
        "hlfr" => ControlWord::ShapeHyperlinkFriendlyName,
        "tsrowd" => ControlWord::TableStyleRowDefaults(flag(param, repair)),
        "tscfirstrow" => ControlWord::TableStyleFirstRow(flag(param, repair)),
        "tsclastrow" => ControlWord::TableStyleLastRow(flag(param, repair)),
        "tscfirstcol" => ControlWord::TableStyleFirstColumn(flag(param, repair)),
        "tsclastcol" => ControlWord::TableStyleLastColumn(flag(param, repair)),
        "tscbandhorzodd" => ControlWord::TableStyleBandHorizontalOdd(flag(param, repair)),
        "tscbandhorzeven" => ControlWord::TableStyleBandHorizontalEven(flag(param, repair)),
        "tscbandvertodd" => ControlWord::TableStyleBandVerticalOdd(flag(param, repair)),
        "tscbandverteven" => ControlWord::TableStyleBandVerticalEven(flag(param, repair)),
        "tscbandsh" => ControlWord::TableStyleBandSizeHorizontal(value(param, 0, repair)),
        "tscbandsv" => ControlWord::TableStyleBandSizeVertical(value(param, 0, repair)),
        "vertsect" => ControlWord::SectionVerticalRendering(flag(param, repair)),
        "horzsect" => ControlWord::SectionHorizontalRendering(flag(param, repair)),
        "nocolbal" => ControlWord::SectionNoColumnBalance(flag(param, repair)),
        "sectdefaultcl" => ControlWord::SectionDefaultColumns(flag(param, repair)),
        "noline" => ControlWord::ParagraphNoLineNumbering(flag(param, repair)),
        "notabind" => ControlWord::ParagraphNoAutoTabIndent(flag(param, repair)),
        "brdrbar" => ControlWord::BorderBar,
        "brdrbtw" => ControlWord::BorderBetween,
        "box" => ControlWord::BorderBox,
        "brdrhair" => ControlWord::BorderHairline,
        "softpage" => ControlWord::SoftPageBreak(flag(param, repair)),
        "softcol" => ControlWord::SoftColumnBreak(flag(param, repair)),
        "softline" => ControlWord::SoftLineBreak(flag(param, repair)),
        "softlheight" => ControlWord::SoftLineHeight(value(param, 0, repair)),
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
        "deftab" => ControlWord::DefaultTabWidth(value(param, 720, repair)),
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
        "ltrrow" => ControlWord::LeftToRightRow(flag(param, repair)),
        "rtlrow" => ControlWord::RightToLeftRow(flag(param, repair)),
        "taprtl" => ControlWord::TableRightToLeft(param_bool),
        "rtlgutter" => ControlWord::RightGutter(param_bool),
        "formprot" => ControlWord::FormProtection(toggle(param, repair)),
        "annotprot" => ControlWord::AnnotationProtection(toggle(param, repair)),
        "revprot" => ControlWord::RevisionProtection(toggle(param, repair)),
        "readprot" => ControlWord::ReadOnlyProtection(toggle(param, repair)),
        "allprot" => ControlWord::AllProtection(toggle(param, repair)),
        "enforceprot" => ControlWord::EnforceProtection(binary(param, 0, repair)),
        "protlevel" => ControlWord::ProtectionLevel(value(param, 0, repair)),
        "password" => ControlWord::Password,
        "passwordhash" => {
            reject_numeric_parameter("passwordhash", flag(param, repair))?;
            ControlWord::PasswordHash
        },
        "protusertbl" => ControlWord::ProtectionUserTable,
        "hyphauto" => ControlWord::HyphenateAutomatically(toggle(param, repair)),
        "hyphcaps" => ControlWord::HyphenateCapitalizedWords(toggle(param, repair)),
        "hyphconsec" => ControlWord::HyphenationConsecutiveLines(value(param, 0, repair)),
        "hyphhotz" => ControlWord::HyphenationHotZone(value(param, 0, repair)),
        "nextfile" => ControlWord::NextFile,
        "template" => ControlWord::DocumentTemplate,
        "viewkind" => ControlWord::DocumentViewKind(value(param, 0, repair)),
        "viewscale" => ControlWord::DocumentViewScale(value(param, 100, repair)),
        "viewzk" => ControlWord::DocumentZoomKind(value(param, 0, repair)),
        "viewbksp" => ControlWord::DocumentViewBackgroundShapes(binary(param, 0, repair)),
        "viewnobound" => ControlWord::DocumentViewNoPageBoundaries(flag(param, repair)),
        "donotshowmarkup" => ControlWord::HideReviewMarkup(flag(param, repair)),
        "donotshowcomments" => ControlWord::HideReviewComments(flag(param, repair)),
        "donotshowinsdel" => ControlWord::HideReviewInsertionsAndDeletions(flag(param, repair)),
        "windowcaption" => ControlWord::WindowCaption,
        "xform" => ControlWord::XslTransform,
        "usexform" => ControlWord::UseXslTransform(flag(param, repair)),
        "wgrffmtfilter" => ControlWord::StyleListFilter(param),
        "stylesortmethod" => ControlWord::StyleSortMethod(param),
        "readonlyrecommended" => ControlWord::ReadOnlyRecommended(flag(param, repair)),
        "saveprevpict" => ControlWord::SavePreviousPicture(flag(param, repair)),
        "writereservation" => ControlWord::WriteReservation(param),
        "writereservhash" => ControlWord::WriteReservationHash(param),
        "fromtext" => ControlWord::FromText(flag(param, repair)),
        "fromhtml" => ControlWord::FromHtml(param),
        "htmlrtf" => ControlWord::HtmlRtf(param_bool),
        "htmltag" => ControlWord::HtmlTag(param.ok_or_else(|| {
            RtfError::MalformedDocument("RTF htmltag requires a numeric parameter".to_string())
        })?),
        "doctype" => ControlWord::DocumentType(param),
        "makebackup" => ControlWord::MakeBackup(flag(param, repair)),
        "defformat" => ControlWord::DefaultSaveFormat(flag(param, repair)),
        "doctemp" => ControlWord::BoilerplateDocument(flag(param, repair)),
        "muser" => ControlWord::Word97CompatibilityMode(flag(param, repair)),
        "psover" => ControlWord::PostScriptOverText(flag(param, repair)),
        "horzdoc" => ControlWord::HorizontalDocument(flag(param, repair)),
        "vertdoc" => ControlWord::VerticalDocument(flag(param, repair)),
        "jcompress" => ControlWord::CompressJustification(flag(param, repair)),
        "jexpand" => ControlWord::ExpandJustification(flag(param, repair)),
        "lnongrid" => ControlWord::LineBasedOnGrid(flag(param, repair)),
        "fracwidth" => ControlWord::FractionalCharacterWidths(flag(param, repair)),
        "ilfomacatclnup" => ControlWord::AbstractNumberingCleanup(binary(param, 0, repair)),
        "grfdocevents" => ControlWord::DocumentEventMask(param),
        "dgmargin" => ControlWord::DrawingGridFollowsMargins(flag(param, repair)),
        "dgsnap" => ControlWord::SnapToDrawingGrid(flag(param, repair)),
        "dghspace" => ControlWord::DrawingGridHorizontalSpacing(value(param, 120, repair)),
        "dgvspace" => ControlWord::DrawingGridVerticalSpacing(value(param, 120, repair)),
        "dghorigin" => ControlWord::DrawingGridHorizontalOrigin(value(param, 1701, repair)),
        "dgvorigin" => ControlWord::DrawingGridVerticalOrigin(value(param, 1984, repair)),
        "dghshow" => ControlWord::DrawingGridHorizontalShow(value(param, 3, repair)),
        "dgvshow" => ControlWord::DrawingGridVerticalShow(value(param, 0, repair)),
        "gutterprl" => ControlWord::ParallelGutter(flag(param, repair)),
        "twoonone" => ControlWord::PrintTwoOnOne(flag(param, repair)),
        "themelang" => ControlWord::ThemeLanguage(value(param, 0, repair)),
        "themelangfe" => ControlWord::ThemeLanguageEastAsian(value(param, 0, repair)),
        "themelangcs" => ControlWord::ThemeLanguageComplexScript(value(param, 0, repair)),
        "relyonvml" => ControlWord::RelyOnVml(binary(param, 0, repair)),
        "validatexml" => ControlWord::ValidateXml(binary(param, 0, repair)),
        "showplaceholdtext" => ControlWord::ShowPlaceholderText(binary(param, 0, repair)),
        "ignoremixedcontent" => ControlWord::IgnoreMixedContent(binary(param, 0, repair)),
        "saveinvalidxml" => ControlWord::SaveInvalidXml(binary(param, 1, repair)),
        "showxmlerrors" => ControlWord::ShowXmlErrors(binary(param, 0, repair)),
        "donotembedsysfont" => ControlWord::DoNotEmbedSystemFonts(binary(param, 0, repair)),
        "donotembedlingdata" => ControlWord::DoNotEmbedLinguisticData(binary(param, 0, repair)),
        "trackmoves" => ControlWord::TrackMoves(binary(param, 0, repair)),
        "trackformatting" => ControlWord::TrackFormatting(binary(param, 0, repair)),
        "stylelocktheme" => ControlWord::LockDocumentTheme(flag(param, repair)),
        "stylelockqfset" => ControlWord::LockQuickFormatSet(flag(param, repair)),
        "usenormstyforlist" => ControlWord::UseNormalStyleForLists(flag(param, repair)),
        "linkstyles" => ControlWord::UpdateStylesFromTemplate(flag(param, repair)),
        "stylelock" => ControlWord::DeclareStyleRestrictions(flag(param, repair)),
        "stylelockenforced" => ControlWord::EnforceStyleRestrictions(flag(param, repair)),
        "stylelockbackcomp" => {
            ControlWord::StyleRestrictionsBackwardCompatibility(flag(param, repair))
        },
        "autofmtoverride" => ControlWord::AllowAutoFormatOverride(flag(param, repair)),
        "bookfold" => ControlWord::BookFold(flag(param, repair)),
        "bookfoldrev" => ControlWord::ReverseBookFold(flag(param, repair)),
        "bookfoldsheets" => ControlWord::BookFoldSheets(value(param, 0, repair)),
        "rempersonalinfo" => ControlWord::RemovePersonalInformation(flag(param, repair)),
        "remdttm" => ControlWord::RemoveDateTimeInformation(flag(param, repair)),
        "noextrasprl" => ControlWord::SuppressRaisedLoweredExtraSpacing(flag(param, repair)),
        "sprstsp" => ControlWord::SuppressTopPageExtraSpacing(flag(param, repair)),
        "sprsspbf" => ControlWord::SuppressSpaceBeforeAfterHardBreak(flag(param, repair)),
        "sprslnsp" => ControlWord::SuppressWordPerfectExtraLineSpacing(flag(param, repair)),
        "sprsbsp" => ControlWord::SuppressBottomPageExtraSpacing(flag(param, repair)),
        "dntblnsbdb" => ControlWord::DoNotBalanceSbcsDbcs(flag(param, repair)),
        "expshrtn" => ControlWord::ExpandSpacingAtShiftReturn(flag(param, repair)),
        "nospaceforul" => ControlWord::DoNotAddSpaceForUnderline(flag(param, repair)),
        "noultrlspc" => ControlWord::DoNotUnderlineTrailingSpaces(flag(param, repair)),
        "noxlattoyen" => ControlWord::DoNotTranslateBackslashToYen(flag(param, repair)),
        "lnbrkrule" => ControlWord::LegacyAsianLineBreakingRules(flag(param, repair)),
        "otblrul" => ControlWord::CombineLegacyTableBorders(flag(param, repair)),
        "alntblind" => ControlWord::DoNotAlignTableRowsIndependently(flag(param, repair)),
        "lytcalctblwd" => ControlWord::DoNotUseRawTableWidth(flag(param, repair)),
        "lyttblrtgr" => ControlWord::KeepTableRowsTogether(flag(param, repair)),
        "nolnhtadjtbl" => ControlWord::DoNotAdjustTableLineHeight(flag(param, repair)),
        "nobrkwrptbl" => ControlWord::DoNotBreakWrappedTablesAcrossPages(flag(param, repair)),
        "nogrowautofit" => ControlWord::PreventAutofitGrowthIntoMargins(flag(param, repair)),
        "newtblstyruls" => ControlWord::UseWord2003TableStyleRules(flag(param, repair)),
        "splytwnine" => ControlWord::DoNotUseWord97ShapeLayout(flag(param, repair)),
        "ftnlytwnine" => ControlWord::UseLegacyFootnoteLayout(flag(param, repair)),
        "htmautsp" => ControlWord::UseHtmlParagraphAutoSpacing(flag(param, repair)),
        "useltbaln" => ControlWord::PreserveLastTabAlignment(flag(param, repair)),
        "oldas" => ControlWord::UseWord95AutoSpacing(flag(param, repair)),
        "ApplyBrkRules" => ControlWord::ApplyThaiLineBreakingRules(flag(param, repair)),
        "snaptogridincell" => ControlWord::SnapTextToGridInsideTable(flag(param, repair)),
        "wrppunct" => ControlWord::AllowHangingPunctuation(flag(param, repair)),
        "asianbrkrule" => ControlWord::UseAsianLineBreakingRules(flag(param, repair)),
        "toplinepunct" => ControlWord::CompressPunctuationAtLineStart(flag(param, repair)),
        "nocompatoptions" => ControlWord::NoCompatibilityOptions(flag(param, repair)),
        "nouicompat" => ControlWord::NoUiCompatibility(flag(param, repair)),
        "nofeaturethrottle" => ControlWord::NoFeatureThrottle(binary(param, 1, repair)),
        "forceupgrade" => ControlWord::ForceCompatibilityUpgrade(flag(param, repair)),
        "noafcnsttbl" => ControlWord::PreserveAutofitTableWidthAroundShapes(flag(param, repair)),
        "noindnmbrts" => ControlWord::UseHangingIndentAsNumberingTab(flag(param, repair)),
        "felnbrelev" => ControlWord::UseLegacyKinsokuCharacters(flag(param, repair)),
        "indrlsweleven" => ControlWord::UseLegacyFloatingObjectIndentation(flag(param, repair)),
        "nocxsptable" => ControlWord::AllowContextualSpacingInTables(flag(param, repair)),
        "notcvasp" => {
            ControlWord::IgnoreCellVerticalAlignmentWithFloatingObjects(flag(param, repair))
        },
        "notvatxbx" => ControlWord::IgnoreTextBoxVerticalAlignment(flag(param, repair)),
        "spltpgpar" => ControlWord::SplitPageBreakParagraph(flag(param, repair)),
        "hwelev" => ControlWord::UseFixedWidthHangul(flag(param, repair)),
        "afelev" => ControlWord::UseLegacyAutofitWidthExpansion(flag(param, repair)),
        "cachedcolbal" => ControlWord::UseCachedColumnBalancing(flag(param, repair)),
        "utinl" => ControlWord::UnderlineNumberingSuffix(flag(param, repair)),
        "notbrkcnstfrctbl" => ControlWord::DoNotSplitRowsAroundFloatingTables(flag(param, repair)),
        "krnprsnet" => ControlWord::UseAnsiKerningPairs(flag(param, repair)),

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
        "fs" => ControlWord::FontSize(param.unwrap_or_else(|| {
            *repair = Some(ParameterDefect::MissingParameter);
            DEFAULT_FONT_SIZE_HALF_POINTS
        })),
        "af" => ControlWord::AssociatedFontNumber(value(param, 0, repair)),
        "afs" => ControlWord::AssociatedFontSize(value(param, 24, repair)),
        "alang" => ControlWord::AssociatedLanguage(value(param, 0, repair)),
        "ab" => ControlWord::AssociatedBold(toggle(param, repair)),
        "acaps" => ControlWord::AssociatedAllCaps(toggle(param, repair)),
        "acf" => ControlWord::AssociatedColor(value(param, 0, repair)),
        "adn" => ControlWord::AssociatedBaselineDown(value(param, 6, repair)),
        "aexpnd" => ControlWord::AssociatedExpansion(value(param, 0, repair)),
        "ai" => ControlWord::AssociatedItalic(toggle(param, repair)),
        "aoutl" => ControlWord::AssociatedOutline(toggle(param, repair)),
        "ascaps" => ControlWord::AssociatedSmallCaps(toggle(param, repair)),
        "ashad" => ControlWord::AssociatedShadow(toggle(param, repair)),
        "astrike" => ControlWord::AssociatedStrike(toggle(param, repair)),
        "aul" => ControlWord::AssociatedUnderline(toggle(param, repair)),
        "auld" => {
            if toggled_selector(param, repair) {
                ControlWord::AssociatedUnderlineDotted(None)
            } else {
                ControlWord::AssociatedUnderlineNone(None)
            }
        },
        "auldb" => {
            if toggled_selector(param, repair) {
                ControlWord::AssociatedUnderlineDouble(None)
            } else {
                ControlWord::AssociatedUnderlineNone(None)
            }
        },
        "aulnone" => ControlWord::AssociatedUnderlineNone(flag(param, repair)),
        "aulw" => {
            if toggled_selector(param, repair) {
                ControlWord::AssociatedUnderlineWords(None)
            } else {
                ControlWord::AssociatedUnderlineNone(None)
            }
        },
        "aup" => ControlWord::AssociatedBaselineUp(value(param, 6, repair)),
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
        "cb" => ControlWord::ColorBackground(value(param, 0, repair)),

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
        "nosupersub" => ControlWord::NoSuperSub(flag(param, repair)),
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
        "chbrdr" => ControlWord::CharacterBorder(flag(param, repair)),
        "chshdng" => ControlWord::CharacterShading(value(param, 0, repair)),
        "chcfpat" => ControlWord::CharacterForegroundPattern(value(param, 0, repair)),
        "chcbpat" => ControlWord::CharacterBackgroundPattern(value(param, 0, repair)),
        "chbghoriz" => ControlWord::CharacterShadingPattern(
            crate::ShadingPattern::Horizontal,
            flag(param, repair),
        ),
        "chbgvert" => ControlWord::CharacterShadingPattern(
            crate::ShadingPattern::Vertical,
            flag(param, repair),
        ),
        "chbgfdiag" => ControlWord::CharacterShadingPattern(
            crate::ShadingPattern::ForwardDiagonal,
            flag(param, repair),
        ),
        "chbgbdiag" => ControlWord::CharacterShadingPattern(
            crate::ShadingPattern::BackwardDiagonal,
            flag(param, repair),
        ),
        "chbgcross" => {
            ControlWord::CharacterShadingPattern(crate::ShadingPattern::Cross, flag(param, repair))
        },
        "chbgdcross" => ControlWord::CharacterShadingPattern(
            crate::ShadingPattern::DiagonalCross,
            flag(param, repair),
        ),
        "chbgdkhoriz" => ControlWord::CharacterShadingPattern(
            crate::ShadingPattern::DarkHorizontal,
            flag(param, repair),
        ),
        "chbgdkvert" => ControlWord::CharacterShadingPattern(
            crate::ShadingPattern::DarkVertical,
            flag(param, repair),
        ),
        "chbgdkfdiag" => ControlWord::CharacterShadingPattern(
            crate::ShadingPattern::DarkForwardDiagonal,
            flag(param, repair),
        ),
        "chbgdkbdiag" => ControlWord::CharacterShadingPattern(
            crate::ShadingPattern::DarkBackwardDiagonal,
            flag(param, repair),
        ),
        "chbgdkcross" => ControlWord::CharacterShadingPattern(
            crate::ShadingPattern::DarkCross,
            flag(param, repair),
        ),
        "chbgdkdcross" => ControlWord::CharacterShadingPattern(
            crate::ShadingPattern::DarkDiagonalCross,
            flag(param, repair),
        ),
        "plain" => ControlWord::Plain,
        "loch" => ControlWord::LowAnsiCharacter(flag(param, repair)),
        "hich" => ControlWord::HighAnsiCharacter(flag(param, repair)),
        "dbch" => ControlWord::DoubleByteCharacter(flag(param, repair)),
        "fcs" => ControlWord::FontComplexScript(binary(param, 0, repair)),
        "cgrid" => ControlWord::CharacterGrid(param),
        "animtext" => ControlWord::AnimatedText(value(param, 0, repair)),
        "fittext" => ControlWord::FitText(value(param, 0, repair)),
        "accnone" => ControlWord::EmphasisMark(crate::EmphasisMark::None, flag(param, repair)),
        "accdot" => ControlWord::EmphasisMark(
            if toggled_selector(param, repair) {
                crate::EmphasisMark::Dot
            } else {
                crate::EmphasisMark::None
            },
            None,
        ),
        "acccomma" => ControlWord::EmphasisMark(
            if toggled_selector(param, repair) {
                crate::EmphasisMark::Comma
            } else {
                crate::EmphasisMark::None
            },
            None,
        ),
        "accunderdot" => ControlWord::EmphasisMark(
            if toggled_selector(param, repair) {
                crate::EmphasisMark::UnderDot
            } else {
                crate::EmphasisMark::None
            },
            None,
        ),
        "acccircle" => ControlWord::EmphasisMark(
            if toggled_selector(param, repair) {
                crate::EmphasisMark::Circle
            } else {
                crate::EmphasisMark::None
            },
            None,
        ),

        // Paragraph
        "par" => ControlWord::Par,
        "page" => ControlWord::Page(flag(param, repair)),
        "column" => ControlWord::Column(flag(param, repair)),
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
        "sbauto" => ControlWord::SpaceBeforeAuto(binary(param, 1, repair)),
        "saauto" => ControlWord::SpaceAfterAuto(binary(param, 1, repair)),
        "lisb" => ControlWord::ListSpaceBefore(value(param, 0, repair)),
        "lisa" => ControlWord::ListSpaceAfter(value(param, 0, repair)),
        "nosnaplinegrid" => ControlWord::NoSnapLineGrid(flag(param, repair)),
        "contextualspace" => ControlWord::ContextualSpacing(flag(param, repair)),
        "li" => ControlWord::LeftIndent(param_value),
        "ri" => ControlWord::RightIndent(param_value),
        "fi" => ControlWord::FirstLineIndent(param_value),
        "lin" => ControlWord::LogicalLeftIndent(value(param, 0, repair)),
        "rin" => ControlWord::LogicalRightIndent(value(param, 0, repair)),
        "cufi" => ControlWord::CharacterFirstLineIndent(value(param, 0, repair)),
        "culi" => ControlWord::CharacterLeftIndent(value(param, 0, repair)),
        "curi" => ControlWord::CharacterRightIndent(value(param, 0, repair)),
        "indmirror" => ControlWord::MirrorIndents(flag(param, repair)),

        // Paragraph additional properties
        "keep" => ControlWord::KeepTogether,
        "keepn" => ControlWord::KeepNext,
        "sbys" => ControlWord::SideBySide(param_bool),
        "pagebb" => ControlWord::PageBreakBefore,
        "widctlpar" => ControlWord::WidowControl,
        "nowidctlpar" => ControlWord::NoWidowControl(param),
        "dropcapli" => ControlWord::DropCapLines(value(param, 0, repair)),
        "dropcapt" => ControlWord::DropCapType(value(param, 0, repair)),
        "hyphpar" => ControlWord::ParagraphHyphenation(toggle(param, repair)),
        "aspalpha" => ControlWord::AutoSpaceAlphabetic(toggle(param, repair)),
        "aspnum" => ControlWord::AutoSpaceNumbers(toggle(param, repair)),
        "adjustright" => ControlWord::AdjustRightIndent(toggle(param, repair)),
        "wrapdefault" => ControlWord::WrapDefault(flag(param, repair)),
        "nocwrap" => ControlWord::NoCharacterWrap(flag(param, repair)),
        "nowwrap" => ControlWord::NoWordWrap(flag(param, repair)),
        "nooverflow" => ControlWord::NoOverflow(flag(param, repair)),
        "faauto" => ControlWord::FontAlignAuto(flag(param, repair)),
        "fahang" => ControlWord::FontAlignHanging(flag(param, repair)),
        "facenter" => ControlWord::FontAlignCenter(flag(param, repair)),
        "faroman" => ControlWord::FontAlignRoman(flag(param, repair)),
        "favar" => ControlWord::FontAlignVariable(flag(param, repair)),
        "fafixed" => ControlWord::FontAlignFixed(flag(param, repair)),

        // Positioned paragraph objects and frames
        "absw" => ControlWord::FrameAbsoluteWidth(value(param, 0, repair)),
        "absh" => ControlWord::FrameAbsoluteHeight(value(param, 0, repair)),
        "phmrg" => ControlWord::FrameHorizontalMargin(flag(param, repair)),
        "phpg" => ControlWord::FrameHorizontalPage(flag(param, repair)),
        "phcol" => ControlWord::FrameHorizontalColumn(flag(param, repair)),
        "posx" => ControlWord::FrameHorizontalOffset(value(param, 0, repair)),
        "posnegx" => ControlWord::FrameHorizontalNegativeOffset(value(param, 0, repair)),
        "posxc" => ControlWord::FrameHorizontalCenter(flag(param, repair)),
        "posxi" => ControlWord::FrameHorizontalInside(flag(param, repair)),
        "posxo" => ControlWord::FrameHorizontalOutside(flag(param, repair)),
        "posxl" => ControlWord::FrameHorizontalLeft(flag(param, repair)),
        "posxr" => ControlWord::FrameHorizontalRight(flag(param, repair)),
        "pvmrg" => ControlWord::FrameVerticalMargin(flag(param, repair)),
        "pvpg" => ControlWord::FrameVerticalPage(flag(param, repair)),
        "pvpara" => ControlWord::FrameVerticalParagraph(flag(param, repair)),
        "posy" => ControlWord::FrameVerticalOffset(value(param, 0, repair)),
        "posnegy" => ControlWord::FrameVerticalNegativeOffset(value(param, 0, repair)),
        "posyil" => ControlWord::FrameVerticalInline(flag(param, repair)),
        "posyt" => ControlWord::FrameVerticalTop(flag(param, repair)),
        "posyc" => ControlWord::FrameVerticalCenter(flag(param, repair)),
        "posyb" => ControlWord::FrameVerticalBottom(flag(param, repair)),
        "posyin" => ControlWord::FrameVerticalInside(flag(param, repair)),
        "posyout" => ControlWord::FrameVerticalOutside(flag(param, repair)),
        "abslock" => ControlWord::FrameAbsoluteLock(binary(param, 1, repair)),
        "nowrap" => ControlWord::FrameNoWrap(flag(param, repair)),
        "dxfrtext" => ControlWord::FrameHorizontalTextDistance(value(param, 0, repair)),
        "dfrmtxtx" => ControlWord::FrameHorizontalTextOffset(value(param, 0, repair)),
        "dfrmtxty" => ControlWord::FrameVerticalTextOffset(value(param, 0, repair)),
        "wraparound" => ControlWord::FrameWrapAround(flag(param, repair)),
        "wraptight" => ControlWord::FrameWrapTight(flag(param, repair)),
        "wrapthrough" => ControlWord::FrameWrapThrough(flag(param, repair)),
        "overlay" => ControlWord::FrameOverlay(flag(param, repair)),
        "absnoovrlp" => ControlWord::FrameNoOverlap(binary(param, 1, repair)),
        "frmtxlrtb" => ControlWord::FrameTextFlowLrtb(flag(param, repair)),
        "frmtxtbrl" => ControlWord::FrameTextFlowTbrl(flag(param, repair)),
        "frmtxbtlr" => ControlWord::FrameTextFlowBtlr(flag(param, repair)),
        "frmtxlrtbv" => ControlWord::FrameTextFlowLrtbv(flag(param, repair)),
        "frmtxtbrlv" => ControlWord::FrameTextFlowTbrlv(flag(param, repair)),

        // Tables
        "trowd" => ControlWord::TableRowDefaults,
        "irow" => ControlWord::TableRowIndex(value(param, 0, repair)),
        "irowband" => ControlWord::TableRowBandIndex(value(param, 0, repair)),
        "lastrow" => ControlWord::TableLastRow(flag(param, repair)),
        "tbllkborder" => ControlWord::TableAutoformatFlag(
            crate::TableAutoformatFlag::Border,
            flag(param, repair),
        ),
        "tbllkshading" => ControlWord::TableAutoformatFlag(
            crate::TableAutoformatFlag::Shading,
            flag(param, repair),
        ),
        "tbllkfont" => {
            ControlWord::TableAutoformatFlag(crate::TableAutoformatFlag::Font, flag(param, repair))
        },
        "tbllkcolor" => {
            ControlWord::TableAutoformatFlag(crate::TableAutoformatFlag::Color, flag(param, repair))
        },
        "tbllkbestfit" => ControlWord::TableAutoformatFlag(
            crate::TableAutoformatFlag::BestFit,
            flag(param, repair),
        ),
        "tbllkhdrrows" => ControlWord::TableAutoformatFlag(
            crate::TableAutoformatFlag::HeaderRows,
            flag(param, repair),
        ),
        "tbllklastrow" => ControlWord::TableAutoformatFlag(
            crate::TableAutoformatFlag::LastRow,
            flag(param, repair),
        ),
        "tbllkhdrcols" => ControlWord::TableAutoformatFlag(
            crate::TableAutoformatFlag::HeaderColumns,
            flag(param, repair),
        ),
        "tbllklastcol" => ControlWord::TableAutoformatFlag(
            crate::TableAutoformatFlag::LastColumn,
            flag(param, repair),
        ),
        "tbllknorowband" => ControlWord::TableAutoformatFlag(
            crate::TableAutoformatFlag::NoRowBanding,
            flag(param, repair),
        ),
        "tbllknocolband" => ControlWord::TableAutoformatFlag(
            crate::TableAutoformatFlag::NoColumnBanding,
            flag(param, repair),
        ),
        "row" => ControlWord::TableRow,
        "cell" => ControlWord::TableCell,
        "cellx" => ControlWord::CellX(param_value),
        "intbl" => ControlWord::InTable,
        "trgaph" => ControlWord::TableRowGap(value(param, 0, repair)),
        "trleft" => ControlWord::TableRowLeft(value(param, 0, repair)),
        "trrh" => ControlWord::TableRowHeight(value(param, 0, repair)),
        "trftsWidth" => ControlWord::TablePreferredWidthUnit(
            crate::TableDistanceScope::Row,
            value(param, 0, repair),
        ),
        "trwWidth" => ControlWord::TablePreferredWidthValue(
            crate::TableDistanceScope::Row,
            value(param, 0, repair),
        ),
        "trftsWidthB" => ControlWord::TableInvisibleWidthUnit(false, value(param, 0, repair)),
        "trwWidthB" => ControlWord::TableInvisibleWidthValue(false, value(param, 0, repair)),
        "trftsWidthA" => ControlWord::TableInvisibleWidthUnit(true, value(param, 0, repair)),
        "trwWidthA" => ControlWord::TableInvisibleWidthValue(true, value(param, 0, repair)),
        "clftsWidth" => ControlWord::TablePreferredWidthUnit(
            crate::TableDistanceScope::Cell,
            value(param, 0, repair),
        ),
        "clwWidth" => ControlWord::TablePreferredWidthValue(
            crate::TableDistanceScope::Cell,
            value(param, 0, repair),
        ),
        "trautofit" => ControlWord::TableAutoFit(binary(param, 1, repair)),
        "tblind" => ControlWord::TableIndentValue(param),
        "tblindtype" => ControlWord::TableIndentUnit(value(param, 0, repair)),
        "trpaddl" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Left,
            },
            value(param, 0, repair),
        ),
        "trpaddr" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Right,
            },
            value(param, 0, repair),
        ),
        "trpaddt" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Top,
            },
            value(param, 0, repair),
        ),
        "trpaddb" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Bottom,
            },
            value(param, 0, repair),
        ),
        "trpaddfl" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Left,
            },
            value(param, 0, repair),
        ),
        "trpaddfr" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Right,
            },
            value(param, 0, repair),
        ),
        "trpaddft" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Top,
            },
            value(param, 0, repair),
        ),
        "trpaddfb" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Bottom,
            },
            value(param, 0, repair),
        ),
        "trspdl" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Left,
            },
            value(param, 0, repair),
        ),
        "trspdr" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Right,
            },
            value(param, 0, repair),
        ),
        "trspdt" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Top,
            },
            value(param, 0, repair),
        ),
        "trspdb" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Bottom,
            },
            value(param, 0, repair),
        ),
        "trspdfl" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Left,
            },
            value(param, 0, repair),
        ),
        "trspdfr" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Right,
            },
            value(param, 0, repair),
        ),
        "trspdft" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Top,
            },
            value(param, 0, repair),
        ),
        "trspdfb" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Row,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Bottom,
            },
            value(param, 0, repair),
        ),
        "clpadl" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Left,
            },
            value(param, 0, repair),
        ),
        "clpadr" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Right,
            },
            value(param, 0, repair),
        ),
        "clpadt" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Top,
            },
            value(param, 0, repair),
        ),
        "clpadb" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Bottom,
            },
            value(param, 0, repair),
        ),
        "clpadfl" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Left,
            },
            value(param, 0, repair),
        ),
        "clpadfr" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Right,
            },
            value(param, 0, repair),
        ),
        "clpadft" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Top,
            },
            value(param, 0, repair),
        ),
        "clpadfb" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Padding,
                edge: crate::TableEdge::Bottom,
            },
            value(param, 0, repair),
        ),
        "clspl" | "clspdl" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Left,
            },
            value(param, 0, repair),
        ),
        "clspr" | "clspdr" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Right,
            },
            value(param, 0, repair),
        ),
        "clspt" | "clspdt" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Top,
            },
            value(param, 0, repair),
        ),
        "clspb" | "clspdb" => ControlWord::TableDistanceValue(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Bottom,
            },
            value(param, 0, repair),
        ),
        "clspfl" | "clspdfl" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Left,
            },
            value(param, 0, repair),
        ),
        "clspfr" | "clspdfr" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Right,
            },
            value(param, 0, repair),
        ),
        "clspft" | "clspdft" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Top,
            },
            value(param, 0, repair),
        ),
        "clspfb" | "clspdfb" => ControlWord::TableDistanceUnit(
            crate::TableDistanceTarget {
                scope: crate::TableDistanceScope::Cell,
                kind: crate::TableDistanceKind::Spacing,
                edge: crate::TableEdge::Bottom,
            },
            value(param, 0, repair),
        ),
        "tphcol" => ControlWord::TableHorizontalReference(
            crate::TableHorizontalReference::Column,
            flag(param, repair),
        ),
        "tphmrg" => ControlWord::TableHorizontalReference(
            crate::TableHorizontalReference::Margin,
            flag(param, repair),
        ),
        "tphpg" => ControlWord::TableHorizontalReference(
            crate::TableHorizontalReference::Page,
            flag(param, repair),
        ),
        "tpvmrg" => ControlWord::TableVerticalReference(
            crate::TableVerticalReference::Margin,
            flag(param, repair),
        ),
        "tpvpara" => ControlWord::TableVerticalReference(
            crate::TableVerticalReference::Paragraph,
            flag(param, repair),
        ),
        "tpvpg" => ControlWord::TableVerticalReference(
            crate::TableVerticalReference::Page,
            flag(param, repair),
        ),
        "tposx" => ControlWord::TableHorizontalOffset(false, value(param, 0, repair)),
        "tposnegx" => ControlWord::TableHorizontalOffset(true, value(param, 0, repair)),
        "tposxc" => ControlWord::TableHorizontalPosition(
            crate::TableHorizontalPosition::Center,
            flag(param, repair),
        ),
        "tposxi" => ControlWord::TableHorizontalPosition(
            crate::TableHorizontalPosition::Inside,
            flag(param, repair),
        ),
        "tposxl" => ControlWord::TableHorizontalPosition(
            crate::TableHorizontalPosition::Left,
            flag(param, repair),
        ),
        "tposxo" => ControlWord::TableHorizontalPosition(
            crate::TableHorizontalPosition::Outside,
            flag(param, repair),
        ),
        "tposxr" => ControlWord::TableHorizontalPosition(
            crate::TableHorizontalPosition::Right,
            flag(param, repair),
        ),
        "tposy" => ControlWord::TableVerticalOffset(false, value(param, 0, repair)),
        "tposnegy" => ControlWord::TableVerticalOffset(true, value(param, 0, repair)),
        "tposyb" => ControlWord::TableVerticalPosition(
            crate::TableVerticalPosition::Bottom,
            flag(param, repair),
        ),
        "tposyc" => ControlWord::TableVerticalPosition(
            crate::TableVerticalPosition::Center,
            flag(param, repair),
        ),
        "tposyil" => ControlWord::TableVerticalPosition(
            crate::TableVerticalPosition::Inline,
            flag(param, repair),
        ),
        "tposyin" => ControlWord::TableVerticalPosition(
            crate::TableVerticalPosition::Inside,
            flag(param, repair),
        ),
        "tposyout" => ControlWord::TableVerticalPosition(
            crate::TableVerticalPosition::Outside,
            flag(param, repair),
        ),
        "tposyt" => ControlWord::TableVerticalPosition(
            crate::TableVerticalPosition::Top,
            flag(param, repair),
        ),
        "tdfrmtxtLeft" => {
            ControlWord::TableWrapDistance(crate::TableEdge::Left, value(param, 0, repair))
        },
        "tdfrmtxtRight" => {
            ControlWord::TableWrapDistance(crate::TableEdge::Right, value(param, 0, repair))
        },
        "tdfrmtxtTop" => {
            ControlWord::TableWrapDistance(crate::TableEdge::Top, value(param, 0, repair))
        },
        "tdfrmtxtBottom" => {
            ControlWord::TableWrapDistance(crate::TableEdge::Bottom, value(param, 0, repair))
        },
        "tabsnoovrlp" => ControlWord::TableNoOverlap(toggle(param, repair)),
        "trhdr" => ControlWord::TableRowHeader(flag(param, repair)),
        "trkeep" => ControlWord::TableRowKeep(flag(param, repair)),
        "trkeepfollow" => ControlWord::TableRowKeepFollow(flag(param, repair)),
        "trql" => {
            ControlWord::TableRowAlignment(crate::TableRowAlignment::Left, flag(param, repair))
        },
        "trqc" => {
            ControlWord::TableRowAlignment(crate::TableRowAlignment::Center, flag(param, repair))
        },
        "trqr" => {
            ControlWord::TableRowAlignment(crate::TableRowAlignment::Right, flag(param, repair))
        },
        "clvertalt" => ControlWord::TableCellVerticalAlignment(
            crate::TableCellVerticalAlignment::Top,
            flag(param, repair),
        ),
        "clvertalc" => ControlWord::TableCellVerticalAlignment(
            crate::TableCellVerticalAlignment::Center,
            flag(param, repair),
        ),
        "clvertalb" => ControlWord::TableCellVerticalAlignment(
            crate::TableCellVerticalAlignment::Bottom,
            flag(param, repair),
        ),
        "cltxlrtb" => ControlWord::TableCellTextFlow(
            crate::TableCellTextFlow::LeftToRightTopToBottom,
            flag(param, repair),
        ),
        "cltxtbrl" => ControlWord::TableCellTextFlow(
            crate::TableCellTextFlow::RightToLeftTopToBottom,
            flag(param, repair),
        ),
        "cltxbtlr" => ControlWord::TableCellTextFlow(
            crate::TableCellTextFlow::LeftToRightBottomToTop,
            flag(param, repair),
        ),
        "cltxlrtbv" => ControlWord::TableCellTextFlow(
            crate::TableCellTextFlow::LeftToRightTopToBottomVertical,
            flag(param, repair),
        ),
        "cltxtbrlv" => ControlWord::TableCellTextFlow(
            crate::TableCellTextFlow::TopToBottomRightToLeftVertical,
            flag(param, repair),
        ),
        "clFitText" => ControlWord::TableCellFitText(flag(param, repair)),
        "clNoWrap" => ControlWord::TableCellNoWrap(flag(param, repair)),
        "clhidemark" => ControlWord::TableCellHideMark(flag(param, repair)),
        "clmgf" => ControlWord::TableCellMerge(
            crate::TableCellMergeAxis::Horizontal,
            crate::TableCellMergeRole::First,
            flag(param, repair),
        ),
        "clmrg" => ControlWord::TableCellMerge(
            crate::TableCellMergeAxis::Horizontal,
            crate::TableCellMergeRole::Continuation,
            flag(param, repair),
        ),
        "clvmgf" => ControlWord::TableCellMerge(
            crate::TableCellMergeAxis::Vertical,
            crate::TableCellMergeRole::First,
            flag(param, repair),
        ),
        "clvmrg" => ControlWord::TableCellMerge(
            crate::TableCellMergeAxis::Vertical,
            crate::TableCellMergeRole::Continuation,
            flag(param, repair),
        ),
        "trbrdrt" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::Row(crate::TableRowBorderSide::Top),
            flag(param, repair),
        ),
        "trbrdrl" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::Row(crate::TableRowBorderSide::Left),
            flag(param, repair),
        ),
        "trbrdrb" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::Row(crate::TableRowBorderSide::Bottom),
            flag(param, repair),
        ),
        "trbrdrr" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::Row(crate::TableRowBorderSide::Right),
            flag(param, repair),
        ),
        "trbrdrh" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::Row(crate::TableRowBorderSide::Horizontal),
            flag(param, repair),
        ),
        "trbrdrv" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::Row(crate::TableRowBorderSide::Vertical),
            flag(param, repair),
        ),
        "clbrdrt" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::Cell(crate::TableCellBorderSide::Top),
            flag(param, repair),
        ),
        "clbrdrl" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::Cell(crate::TableCellBorderSide::Left),
            flag(param, repair),
        ),
        "clbrdrb" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::Cell(crate::TableCellBorderSide::Bottom),
            flag(param, repair),
        ),
        "clbrdrr" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::Cell(crate::TableCellBorderSide::Right),
            flag(param, repair),
        ),
        "cldglu" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::Cell(
                crate::TableCellBorderSide::UpperLeftToLowerRight,
            ),
            flag(param, repair),
        ),
        "cldgll" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::Cell(
                crate::TableCellBorderSide::UpperRightToLowerLeft,
            ),
            flag(param, repair),
        ),
        "tsbrdrt" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::StyleDefault(crate::TableStyleBorderSide::Top),
            flag(param, repair),
        ),
        "tsbrdrl" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::StyleDefault(crate::TableStyleBorderSide::Left),
            flag(param, repair),
        ),
        "tsbrdrb" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::StyleDefault(crate::TableStyleBorderSide::Bottom),
            flag(param, repair),
        ),
        "tsbrdrr" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::StyleDefault(crate::TableStyleBorderSide::Right),
            flag(param, repair),
        ),
        "tsbrdrh" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::StyleDefault(
                crate::TableStyleBorderSide::HorizontalInside,
            ),
            flag(param, repair),
        ),
        "tsbrdrv" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::StyleDefault(
                crate::TableStyleBorderSide::VerticalInside,
            ),
            flag(param, repair),
        ),
        "tsbrdrdgl" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::StyleDefault(
                crate::TableStyleBorderSide::DiagonalUpperLeftToLowerRight,
            ),
            flag(param, repair),
        ),
        "tsbrdrdg" => ControlWord::TableBorder(
            crate::table::TableBorderTarget::StyleDefault(
                crate::TableStyleBorderSide::DiagonalUpperRightToLowerLeft,
            ),
            flag(param, repair),
        ),
        "tscellpaddl" => ControlWord::TableDefaultDistanceValue(
            crate::TableDistanceKind::Padding,
            crate::TableEdge::Left,
            value(param, 0, repair),
        ),
        "tscellpaddr" => ControlWord::TableDefaultDistanceValue(
            crate::TableDistanceKind::Padding,
            crate::TableEdge::Right,
            value(param, 0, repair),
        ),
        "tscellpaddt" => ControlWord::TableDefaultDistanceValue(
            crate::TableDistanceKind::Padding,
            crate::TableEdge::Top,
            value(param, 0, repair),
        ),
        "tscellpaddb" => ControlWord::TableDefaultDistanceValue(
            crate::TableDistanceKind::Padding,
            crate::TableEdge::Bottom,
            value(param, 0, repair),
        ),
        "tscellpaddfl" => ControlWord::TableDefaultDistanceUnit(
            crate::TableDistanceKind::Padding,
            crate::TableEdge::Left,
            value(param, 0, repair),
        ),
        "tscellpaddfr" => ControlWord::TableDefaultDistanceUnit(
            crate::TableDistanceKind::Padding,
            crate::TableEdge::Right,
            value(param, 0, repair),
        ),
        "tscellpaddft" => ControlWord::TableDefaultDistanceUnit(
            crate::TableDistanceKind::Padding,
            crate::TableEdge::Top,
            value(param, 0, repair),
        ),
        "tscellpaddfb" => ControlWord::TableDefaultDistanceUnit(
            crate::TableDistanceKind::Padding,
            crate::TableEdge::Bottom,
            value(param, 0, repair),
        ),
        "tscellspcl" => ControlWord::TableDefaultDistanceValue(
            crate::TableDistanceKind::Spacing,
            crate::TableEdge::Left,
            value(param, 0, repair),
        ),
        "tscellspcr" => ControlWord::TableDefaultDistanceValue(
            crate::TableDistanceKind::Spacing,
            crate::TableEdge::Right,
            value(param, 0, repair),
        ),
        "tscellspct" => ControlWord::TableDefaultDistanceValue(
            crate::TableDistanceKind::Spacing,
            crate::TableEdge::Top,
            value(param, 0, repair),
        ),
        "tscellspcb" => ControlWord::TableDefaultDistanceValue(
            crate::TableDistanceKind::Spacing,
            crate::TableEdge::Bottom,
            value(param, 0, repair),
        ),
        "tscellspcfl" => ControlWord::TableDefaultDistanceUnit(
            crate::TableDistanceKind::Spacing,
            crate::TableEdge::Left,
            value(param, 0, repair),
        ),
        "tscellspcfr" => ControlWord::TableDefaultDistanceUnit(
            crate::TableDistanceKind::Spacing,
            crate::TableEdge::Right,
            value(param, 0, repair),
        ),
        "tscellspcft" => ControlWord::TableDefaultDistanceUnit(
            crate::TableDistanceKind::Spacing,
            crate::TableEdge::Top,
            value(param, 0, repair),
        ),
        "tscellspcfb" => ControlWord::TableDefaultDistanceUnit(
            crate::TableDistanceKind::Spacing,
            crate::TableEdge::Bottom,
            value(param, 0, repair),
        ),
        "tscellwidthfts" => ControlWord::TableDefaultCellWidthUnit(value(param, 0, repair)),
        "tscellwidth" => ControlWord::TableDefaultCellWidthValue(value(param, 0, repair)),
        "trshdng" => {
            ControlWord::TableShadingAmount(crate::TableDistanceScope::Row, value(param, 0, repair))
        },
        "trcfpat" => ControlWord::TableShadingForeground(
            crate::TableDistanceScope::Row,
            value(param, 0, repair),
        ),
        "trcbpat" => ControlWord::TableShadingBackground(
            crate::TableDistanceScope::Row,
            value(param, 0, repair),
        ),
        "trpat" => ControlWord::TableRowShadingPatternIndex(value(param, 0, repair)),
        "clshdng" => ControlWord::TableShadingAmount(
            crate::TableDistanceScope::Cell,
            value(param, 0, repair),
        ),
        "clshdngraw" => ControlWord::TableShadingRawAmount(
            crate::TableDistanceScope::Cell,
            value(param, 0, repair),
        ),
        "clshdrawnil" => {
            ControlWord::TableShadingRawNil(crate::TableDistanceScope::Cell, flag(param, repair))
        },
        "clcfpat" => ControlWord::TableShadingForeground(
            crate::TableDistanceScope::Cell,
            value(param, 0, repair),
        ),
        "clcbpat" => ControlWord::TableShadingBackground(
            crate::TableDistanceScope::Cell,
            value(param, 0, repair),
        ),
        "trbghoriz" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Row,
            crate::ShadingPattern::Horizontal,
            flag(param, repair),
        ),
        "trbgvert" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Row,
            crate::ShadingPattern::Vertical,
            flag(param, repair),
        ),
        "trbgfdiag" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Row,
            crate::ShadingPattern::ForwardDiagonal,
            flag(param, repair),
        ),
        "trbgbdiag" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Row,
            crate::ShadingPattern::BackwardDiagonal,
            flag(param, repair),
        ),
        "trbgcross" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Row,
            crate::ShadingPattern::Cross,
            flag(param, repair),
        ),
        "trbgdcross" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Row,
            crate::ShadingPattern::DiagonalCross,
            flag(param, repair),
        ),
        "trbgdkhor" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Row,
            crate::ShadingPattern::DarkHorizontal,
            flag(param, repair),
        ),
        "trbgdkvert" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Row,
            crate::ShadingPattern::DarkVertical,
            flag(param, repair),
        ),
        "trbgdkfdiag" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Row,
            crate::ShadingPattern::DarkForwardDiagonal,
            flag(param, repair),
        ),
        "trbgdkbdiag" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Row,
            crate::ShadingPattern::DarkBackwardDiagonal,
            flag(param, repair),
        ),
        "trbgdkcross" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Row,
            crate::ShadingPattern::DarkCross,
            flag(param, repair),
        ),
        "trbgdkdcross" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Row,
            crate::ShadingPattern::DarkDiagonalCross,
            flag(param, repair),
        ),
        "clbghoriz" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Cell,
            crate::ShadingPattern::Horizontal,
            flag(param, repair),
        ),
        "clbgvert" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Cell,
            crate::ShadingPattern::Vertical,
            flag(param, repair),
        ),
        "clbgfdiag" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Cell,
            crate::ShadingPattern::ForwardDiagonal,
            flag(param, repair),
        ),
        "clbgbdiag" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Cell,
            crate::ShadingPattern::BackwardDiagonal,
            flag(param, repair),
        ),
        "clbgcross" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Cell,
            crate::ShadingPattern::Cross,
            flag(param, repair),
        ),
        "clbgdcross" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Cell,
            crate::ShadingPattern::DiagonalCross,
            flag(param, repair),
        ),
        "clbgdkhor" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Cell,
            crate::ShadingPattern::DarkHorizontal,
            flag(param, repair),
        ),
        "clbgdkvert" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Cell,
            crate::ShadingPattern::DarkVertical,
            flag(param, repair),
        ),
        "clbgdkfdiag" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Cell,
            crate::ShadingPattern::DarkForwardDiagonal,
            flag(param, repair),
        ),
        "clbgdkbdiag" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Cell,
            crate::ShadingPattern::DarkBackwardDiagonal,
            flag(param, repair),
        ),
        "clbgdkcross" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Cell,
            crate::ShadingPattern::DarkCross,
            flag(param, repair),
        ),
        "clbgdkdcross" => ControlWord::TableShadingPattern(
            crate::TableDistanceScope::Cell,
            crate::ShadingPattern::DarkDiagonalCross,
            flag(param, repair),
        ),
        "nestcell" => ControlWord::NestedTableCell(flag(param, repair)),
        "nestrow" => ControlWord::NestedTableRow(flag(param, repair)),
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
        "brdrw" => ControlWord::BorderWidth(value(param, 0, repair)),
        "brdrcf" => ControlWord::BorderColor(value(param, 0, repair)),
        "brsp" => ControlWord::BorderSpace(value(param, 0, repair)),
        "pgbrdrt" => ControlWord::PageBorderTop,
        "pgbrdrl" => ControlWord::PageBorderLeft,
        "pgbrdrb" => ControlWord::PageBorderBottom,
        "pgbrdrr" => ControlWord::PageBorderRight,
        "pgbrdropt" => ControlWord::PageBorderOptions(value(param, 0, repair)),
        "pgbrdrhead" => ControlWord::PageBorderSurroundHeader,
        "pgbrdrfoot" => ControlWord::PageBorderSurroundFooter,
        "pgbrdrsnap" => ControlWord::PageBorderSnap,
        "brdrart" => ControlWord::PageBorderArt(param),

        // Shading
        "shading" => ControlWord::Shading(value(param, 0, repair)),
        "cfpat" => ControlWord::ForegroundPattern(value(param, 0, repair)),
        "cbpat" => ControlWord::BackgroundPattern(value(param, 0, repair)),
        "bghoriz" => ControlWord::ParagraphShadingPattern(
            crate::ShadingPattern::Horizontal,
            flag(param, repair),
        ),
        "bgvert" => ControlWord::ParagraphShadingPattern(
            crate::ShadingPattern::Vertical,
            flag(param, repair),
        ),
        "bgfdiag" => ControlWord::ParagraphShadingPattern(
            crate::ShadingPattern::ForwardDiagonal,
            flag(param, repair),
        ),
        "bgbdiag" => ControlWord::ParagraphShadingPattern(
            crate::ShadingPattern::BackwardDiagonal,
            flag(param, repair),
        ),
        "bgcross" => {
            ControlWord::ParagraphShadingPattern(crate::ShadingPattern::Cross, flag(param, repair))
        },
        "bgdcross" => ControlWord::ParagraphShadingPattern(
            crate::ShadingPattern::DiagonalCross,
            flag(param, repair),
        ),
        "bgdkhoriz" => ControlWord::ParagraphShadingPattern(
            crate::ShadingPattern::DarkHorizontal,
            flag(param, repair),
        ),
        "bgdkvert" => ControlWord::ParagraphShadingPattern(
            crate::ShadingPattern::DarkVertical,
            flag(param, repair),
        ),
        "bgdkfdiag" => ControlWord::ParagraphShadingPattern(
            crate::ShadingPattern::DarkForwardDiagonal,
            flag(param, repair),
        ),
        "bgdkbdiag" => ControlWord::ParagraphShadingPattern(
            crate::ShadingPattern::DarkBackwardDiagonal,
            flag(param, repair),
        ),
        "bgdkcross" => ControlWord::ParagraphShadingPattern(
            crate::ShadingPattern::DarkCross,
            flag(param, repair),
        ),
        "bgdkdcross" => ControlWord::ParagraphShadingPattern(
            crate::ShadingPattern::DarkDiagonalCross,
            flag(param, repair),
        ),

        // Tab stops
        "tql" => ControlWord::TabLeft(flag(param, repair)),
        "tqr" => ControlWord::TabRight(flag(param, repair)),
        "tqc" => ControlWord::TabCenter(flag(param, repair)),
        "tqdec" => ControlWord::TabDecimal(flag(param, repair)),
        "tb" => ControlWord::TabBar(value(param, 0, repair)),
        "tx" => ControlWord::TabPosition(value(param, 0, repair)),
        "tldot" => ControlWord::TabLeaderDot(flag(param, repair)),
        "tlmdot" => ControlWord::TabLeaderMiddleDot(flag(param, repair)),
        "tlhyph" => ControlWord::TabLeaderHyphen(flag(param, repair)),
        "tlul" => ControlWord::TabLeaderUnderscore(flag(param, repair)),
        "tlth" => ControlWord::TabLeaderThick(flag(param, repair)),
        "tleq" => ControlWord::TabLeaderEqual(flag(param, repair)),

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
        "binfsxn" => ControlWord::PaperSourceFirst(value(param, 0, repair)),
        "binsxn" => ControlWord::PaperSourceOther(value(param, 0, repair)),
        "facingp" => ControlWord::FacingPages(param_bool),
        "margmirror" => ControlWord::MirrorMargins(flag(param, repair)),
        "gutter" => ControlWord::DocumentGutter(value(param, 0, repair)),
        "headery" => ControlWord::HeaderDistance(param_value),
        "footery" => ControlWord::FooterDistance(param_value),
        "landscape" | "lndscpsxn" => ControlWord::Landscape,
        "cols" => ControlWord::Columns(param),
        "colsx" => ControlWord::ColumnSpace(param),
        "colno" => ControlWord::ColumnNumber(value(param, 0, repair)),
        "colw" => ControlWord::ColumnWidth(value(param, 0, repair)),
        "colsr" => ControlWord::ColumnSpaceRight(value(param, 0, repair)),
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

        // Tracked move bookmarks
        "mvfmf" => {
            reject_numeric_parameter("mvfmf", flag(param, repair))?;
            ControlWord::MoveFromStart
        },
        "mvfml" => {
            reject_numeric_parameter("mvfml", flag(param, repair))?;
            ControlWord::MoveFromEnd
        },
        "mvtof" => {
            reject_numeric_parameter("mvtof", flag(param, repair))?;
            ControlWord::MoveToStart
        },
        "mvtol" => {
            reject_numeric_parameter("mvtol", flag(param, repair))?;
            ControlWord::MoveToEnd
        },

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
        "lquote" => ControlWord::LeftSingleQuote,
        "rquote" => ControlWord::RightSingleQuote,
        "ldblquote" => ControlWord::LeftDoubleQuote,
        "rdblquote" => ControlWord::RightDoubleQuote,
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
