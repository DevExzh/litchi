//! The DOCX formatting engine of `document::Edit` (unified-ops L-DOCX-W5)
//! through the public API: the property/value pairs of OfficeCLI's
//! `examples/word/run-formatting.sh` and `paragraph-formatting.sh` on a blank
//! document, publication through both package owners, tracked-change
//! preservation on a real producer's output, and the signature policy.

use std::io::Cursor;

use litchi_core::Position;
use litchi_docx::color::Theme;
use litchi_docx::document::{BodyAnchor, BodyBlock, Edit, NewParagraph, NewRun, raw};
use litchi_docx::paragraph::{LineSpacingRule, ParagraphSpacing};
use litchi_docx::properties::{
    Border, BorderStyle, Color, EighthPoints, Emphasis, Fonts, HalfPoints, HexColor, Highlight,
    Indentation, Justification, Language, Numbering, ParagraphBorders, Points, Shading,
    ShadingPattern, TabKind, TabStop, TextAlignment, TextEffect, TextboxTightWrap, ThemeFont,
    Twips, VerticalAlign,
};
use litchi_docx::{Package, RunUnderline, RunUnderlineColor, UnderlineStyle, source_backed};
use litchi_opc::OpcError;

type RunSet = fn(&mut litchi_docx::properties::Run);
type ParagraphSet = fn(&mut litchi_docx::properties::Paragraph);

/// One OfficeCLI `add` line: the property/value pairs as the script writes
/// them, for the record, and their typed equivalents.
struct Line {
    keys: &'static str,
    paragraph: ParagraphSet,
    runs: Vec<(&'static str, RunSet)>,
}

fn none_paragraph(_: &mut litchi_docx::properties::Paragraph) {}
fn none_run(_: &mut litchi_docx::properties::Run) {}

fn line(keys: &'static str, paragraph: ParagraphSet, run: RunSet) -> Line {
    Line {
        keys,
        paragraph,
        runs: vec![(keys, run)],
    }
}

fn rgb(hex: u32) -> HexColor {
    let [_, red, green, blue] = hex.to_be_bytes();
    HexColor::Rgb([red, green, blue])
}

fn color(hex: u32) -> Option<Color> {
    Some(Color {
        value: rgb(hex),
        theme: None,
        theme_tint: None,
        theme_shade: None,
    })
}

fn underline(style: UnderlineStyle, color: Option<[u8; 3]>) -> Option<RunUnderline> {
    Some(RunUnderline {
        style,
        color: color.map(RunUnderlineColor::Rgb),
        theme_color: None,
        theme_tint: None,
        theme_shade: None,
    })
}

fn border(style: &str, size: Option<i64>, hex: Option<u32>, space: Option<i64>) -> Border {
    let mut border = Border::new(BorderStyle::from_xml(style).unwrap());
    border.size = size.map(EighthPoints);
    border.color = hex.map(rgb);
    border.space = space.map(Points);
    border
}

fn fill(hex: u32) -> Option<Shading> {
    let mut shading = Shading::new(ShadingPattern::Clear);
    shading.fill = Some(rgb(hex));
    Some(shading)
}

fn fonts(ascii: Option<&str>, east_asia: Option<&str>, cs: Option<&str>) -> Fonts {
    let mut fonts = Fonts::default();
    fonts.ascii = ascii.map(str::to_owned);
    fonts.h_ansi = ascii.map(str::to_owned);
    fonts.east_asia = east_asia.map(str::to_owned);
    fonts.complex_script = cs.map(str::to_owned);
    fonts
}

/// `examples/word/run-formatting.sh`. Not ported, because the typed model
/// does not own them: the Word 2010 effects (`textFill`, `textOutline`,
/// `w14glow`, `w14reflection`, `w14shadow`), `eastAsianLayout.*` and run
/// `fitText`.
fn run_formatting() -> Vec<Line> {
    vec![
        line(
            "align=center bold=true size=20",
            |paragraph| paragraph.justification = Some(Justification::Center),
            |run| {
                run.bold = Some(true);
                run.size = Some(HalfPoints(40));
            },
        ),
        line(
            "heading: bold=true size=14 color=1F4E79 spaceBefore=8pt",
            |paragraph| {
                paragraph.spacing = Some(ParagraphSpacing {
                    before: Some(160),
                    ..ParagraphSpacing::default()
                });
            },
            |run| {
                run.bold = Some(true);
                run.size = Some(HalfPoints(28));
                run.color = color(0x1F_4E79);
            },
        ),
        line("bold=true", none_paragraph, |run| run.bold = Some(true)),
        line("italic=true", none_paragraph, |run| run.italic = Some(true)),
        line("bold=true italic=true", none_paragraph, |run| {
            run.bold = Some(true);
            run.italic = Some(true);
        }),
        line("underline=single", none_paragraph, |run| {
            run.underline = underline(UnderlineStyle::Single, None);
        }),
        line("underline=double", none_paragraph, |run| {
            run.underline = underline(UnderlineStyle::Double, None);
        }),
        line("underline=thick", none_paragraph, |run| {
            run.underline = underline(UnderlineStyle::Thick, None);
        }),
        line("underline=dotted", none_paragraph, |run| {
            run.underline = underline(UnderlineStyle::Dotted, None);
        }),
        line(
            "underline=wave underline.color=FF0000",
            none_paragraph,
            |run| run.underline = underline(UnderlineStyle::Wave, Some([0xFF, 0, 0])),
        ),
        line("strike=true", none_paragraph, |run| run.strike = Some(true)),
        line("dstrike=true", none_paragraph, |run| {
            run.double_strike = Some(true);
        }),
        line("caps=true", none_paragraph, |run| run.caps = Some(true)),
        line("smallcaps=true", none_paragraph, |run| {
            run.small_caps = Some(true);
        }),
        Line {
            keys: "E = mc + run 2 superscript=true",
            paragraph: none_paragraph,
            runs: vec![
                ("E = mc", none_run),
                ("2", |run| {
                    run.vertical_align = Some(VerticalAlign::Superscript);
                }),
            ],
        },
        Line {
            keys: "H + run 2 subscript=true + run O",
            paragraph: none_paragraph,
            runs: vec![
                ("H", none_run),
                ("2", |run| {
                    run.vertical_align = Some(VerticalAlign::Subscript)
                }),
                ("O", none_run),
            ],
        },
        line("color=C00000 size=16", none_paragraph, |run| {
            run.color = color(0xC0_0000);
            run.size = Some(HalfPoints(32));
        }),
        line("highlight=yellow", none_paragraph, |run| {
            run.highlight = Some(Highlight::Yellow);
        }),
        line(
            "font.latin=Georgia font.eastAsia=SimSun size=14",
            none_paragraph,
            |run| {
                run.fonts = Some(fonts(Some("Georgia"), Some("SimSun"), None));
                run.size = Some(HalfPoints(28));
            },
        ),
        line("emboss=true", none_paragraph, |run| run.emboss = Some(true)),
        line("imprint=true", none_paragraph, |run| {
            run.imprint = Some(true)
        }),
        line("outline=true", none_paragraph, |run| {
            run.outline = Some(true)
        }),
        line("shadow=true", none_paragraph, |run| run.shadow = Some(true)),
        line("charSpacing=2pt", none_paragraph, |run| {
            run.character_spacing = Some(Twips(40));
        }),
        line("position=3pt", none_paragraph, |run| {
            run.position = Some(HalfPoints(6));
        }),
        line("lang=en-US", none_paragraph, |run| {
            let mut language = Language::default();
            language.latin = Some("en-US".into());
            run.language = Some(language);
        }),
        line(
            "bold.cs=true italic.cs=true size.cs=14pt",
            none_paragraph,
            |run| {
                run.bold_complex_script = Some(true);
                run.italic_complex_script = Some(true);
                run.size_complex_script = Some(HalfPoints(28));
            },
        ),
        line(
            "rtl=true direction=rtl",
            |paragraph| paragraph.bidi = Some(true),
            |run| run.right_to_left = Some(true),
        ),
        line(
            "font.asciiTheme=minorHAnsi font.hAnsiTheme=minorHAnsi font.csTheme=minorBidi font.eaTheme=minorEastAsia",
            none_paragraph,
            |run| {
                let mut fonts = Fonts::default();
                fonts.ascii_theme = Some(ThemeFont::MinorHAnsi);
                fonts.h_ansi_theme = Some(ThemeFont::MinorHAnsi);
                fonts.complex_script_theme = Some(ThemeFont::MinorBidi);
                fonts.east_asia_theme = Some(ThemeFont::MinorEastAsia);
                run.fonts = Some(fonts);
            },
        ),
        line("font=Calibri", none_paragraph, |run| {
            run.fonts = Some(fonts(Some("Calibri"), Some("Calibri"), None));
        }),
        line("font.cs=Arial font.ea=SimSun", none_paragraph, |run| {
            run.fonts = Some(fonts(None, Some("SimSun"), Some("Arial")));
        }),
        line(
            "lang.latin=en-US lang.ea=zh-CN lang.cs=ar-SA",
            none_paragraph,
            |run| {
                let mut language = Language::default();
                language.latin = Some("en-US".into());
                language.east_asia = Some("zh-CN".into());
                language.bidi = Some("ar-SA".into());
                run.language = Some(language);
            },
        ),
        line(
            "shading=FFFF00",
            |paragraph| paragraph.shading = fill(0xFF_FF00),
            none_run,
        ),
        line("vanish=true", none_paragraph, |run| run.vanish = Some(true)),
        line("noproof=true", none_paragraph, |run| {
            run.no_proof = Some(true)
        }),
        line("vertAlign=superscript", none_paragraph, |run| {
            run.vertical_align = Some(VerticalAlign::Superscript);
        }),
        line("kern=28", none_paragraph, |run| {
            run.kerning = Some(HalfPoints(28));
        }),
        Line {
            keys: "Boxed run + run bdr=single + run bdr=single;4;FF0000;0",
            paragraph: none_paragraph,
            runs: vec![
                ("Boxed run: ", none_run),
                ("single border", |run| {
                    run.border = Some(border("single", None, None, None));
                }),
                ("  red 0.5pt", |run| {
                    run.border = Some(border("single", Some(4), Some(0xFF_0000), Some(0)));
                }),
            ],
        },
        Line {
            keys: "Run character style + run rStyle=Emphasis",
            paragraph: none_paragraph,
            runs: vec![
                ("Run character style: ", none_run),
                ("Emphasis", |run| {
                    run.style = Some("Emphasis".into());
                }),
            ],
        },
        line("em=dot", none_paragraph, |run| {
            run.emphasis = Some(Emphasis::Dot)
        }),
        line("em=underDot", none_paragraph, |run| {
            run.emphasis = Some(Emphasis::UnderDot);
        }),
        line("em=circle", none_paragraph, |run| {
            run.emphasis = Some(Emphasis::Circle);
        }),
        line("effect=blinkBackground", none_paragraph, |run| {
            run.effect = Some(TextEffect::BlinkBackground);
        }),
        line("webHidden=true", none_paragraph, |run| {
            run.web_hidden = Some(true);
        }),
        Line {
            keys: "Layout grid + run snapToGrid=false + run specVanish=true",
            paragraph: none_paragraph,
            runs: vec![
                ("Layout grid + special vanish: ", none_run),
                ("snapToGrid=false", |run| run.snap_to_grid = Some(false)),
                ("  specVanish", |run| run.special_vanish = Some(true)),
            ],
        },
    ]
}

/// `examples/word/paragraph-formatting.sh`. Not ported: `framePr.*` and
/// `divId` (not in the typed model) and `listStyle`/`start`, which create
/// numbering definitions (L-DOCX-W10).
fn paragraph_formatting() -> Vec<Line> {
    vec![
        line(
            "align=left",
            |paragraph| paragraph.justification = Some(Justification::Left),
            none_run,
        ),
        line(
            "align=center",
            |paragraph| paragraph.justification = Some(Justification::Center),
            none_run,
        ),
        line(
            "align=right",
            |paragraph| paragraph.justification = Some(Justification::Right),
            none_run,
        ),
        line(
            "align=both",
            |paragraph| paragraph.justification = Some(Justification::Both),
            none_run,
        ),
        line(
            "indent=1cm",
            |paragraph| {
                let mut indentation = Indentation::default();
                indentation.left = Some(Twips(567));
                paragraph.indentation = Some(indentation);
            },
            none_run,
        ),
        line(
            "rightIndent=2cm",
            |paragraph| {
                let mut indentation = Indentation::default();
                indentation.right = Some(Twips(1134));
                paragraph.indentation = Some(indentation);
            },
            none_run,
        ),
        line(
            "firstLineIndent=1cm",
            |paragraph| {
                let mut indentation = Indentation::default();
                indentation.first_line = Some(Twips(567));
                paragraph.indentation = Some(indentation);
            },
            none_run,
        ),
        line(
            "indent=1cm hangingIndent=1cm",
            |paragraph| {
                let mut indentation = Indentation::default();
                indentation.left = Some(Twips(567));
                indentation.hanging = Some(Twips(567));
                paragraph.indentation = Some(indentation);
            },
            none_run,
        ),
        line(
            "spaceBefore=18pt spaceAfter=6pt",
            |paragraph| {
                paragraph.spacing = Some(ParagraphSpacing {
                    before: Some(360),
                    after: Some(120),
                    ..ParagraphSpacing::default()
                });
            },
            none_run,
        ),
        line(
            "lineSpacing=1.5x",
            |paragraph| {
                paragraph.spacing = Some(ParagraphSpacing {
                    line: Some(360),
                    line_rule: Some(LineSpacingRule::Auto),
                    ..ParagraphSpacing::default()
                });
            },
            none_run,
        ),
        line(
            "keepNext=true",
            |paragraph| paragraph.keep_next = Some(true),
            none_run,
        ),
        line(
            "keepLines=true",
            |paragraph| paragraph.keep_lines = Some(true),
            none_run,
        ),
        line(
            "widowControl=true",
            |paragraph| paragraph.widow_control = Some(true),
            none_run,
        ),
        line("bold=true color=C00000 size=13", none_paragraph, |run| {
            run.bold = Some(true);
            run.color = color(0xC0_0000);
            run.size = Some(HalfPoints(26));
        }),
        line("italic=true highlight=yellow", none_paragraph, |run| {
            run.italic = Some(true);
            run.highlight = Some(Highlight::Yellow);
        }),
        line(
            "shading.fill=D9D9D9",
            |paragraph| paragraph.shading = fill(0xD9_D9D9),
            none_run,
        ),
        line(
            "shading.fill=DDEBF7",
            |paragraph| paragraph.shading = fill(0xDD_EBF7),
            none_run,
        ),
        line(
            "markRPr.bold=true markRPr.color=C00000",
            |paragraph| {
                let mark = paragraph.mark.get_or_insert_with(Default::default);
                mark.bold = Some(true);
                mark.color = color(0xC0_0000);
            },
            none_run,
        ),
        line(
            "outlineLvl=1",
            |paragraph| paragraph.outline_level = Some(1),
            none_run,
        ),
        line("strike=true", none_paragraph, |run| run.strike = Some(true)),
        line(
            "underline=wave underline.color=#FF0000",
            none_paragraph,
            |run| run.underline = underline(UnderlineStyle::Wave, Some([0xFF, 0, 0])),
        ),
        line(
            "bold.cs=true italic.cs=true size.cs=14pt direction=rtl",
            |paragraph| paragraph.bidi = Some(true),
            |run| {
                run.bold_complex_script = Some(true);
                run.italic_complex_script = Some(true);
                run.size_complex_script = Some(HalfPoints(28));
            },
        ),
        line(
            "contextualSpacing=true",
            |paragraph| paragraph.contextual_spacing = Some(true),
            none_run,
        ),
        line(
            "lineSpacing=14pt lineRule=atLeast",
            |paragraph| {
                paragraph.spacing = Some(ParagraphSpacing {
                    line: Some(280),
                    line_rule: Some(LineSpacingRule::AtLeast),
                    ..ParagraphSpacing::default()
                });
            },
            none_run,
        ),
        line(
            "pageBreakBefore=true",
            |paragraph| paragraph.page_break_before = Some(true),
            none_run,
        ),
        line(
            "wordWrap=false",
            |paragraph| paragraph.word_wrap = Some(false),
            none_run,
        ),
        line(
            "firstLineChars=200 hangingChars=100",
            |paragraph| {
                let mut indentation = Indentation::default();
                indentation.first_line_chars = Some(200);
                indentation.hanging_chars = Some(100);
                paragraph.indentation = Some(indentation);
            },
            none_run,
        ),
        line("font=Times New Roman", none_paragraph, |run| {
            run.fonts = Some(fonts(
                Some("Times New Roman"),
                Some("Times New Roman"),
                None,
            ));
        }),
        line(
            "font.latin=Calibri font.ea=SimSun font.cs=Arial",
            none_paragraph,
            |run| run.fonts = Some(fonts(Some("Calibri"), Some("SimSun"), Some("Arial"))),
        ),
        line(
            "font.asciiTheme=minorHAnsi font.hAnsiTheme=minorHAnsi font.eaTheme=minorEastAsia font.csTheme=minorBidi",
            none_paragraph,
            |run| {
                let mut fonts = Fonts::default();
                fonts.ascii_theme = Some(ThemeFont::MinorHAnsi);
                fonts.h_ansi_theme = Some(ThemeFont::MinorHAnsi);
                fonts.east_asia_theme = Some(ThemeFont::MinorEastAsia);
                fonts.complex_script_theme = Some(ThemeFont::MinorBidi);
                run.fonts = Some(fonts);
            },
        ),
        line(
            "style=Heading1",
            |paragraph| paragraph.style = Some("Heading1".into()),
            none_run,
        ),
        line("rStyle=Emphasis", none_paragraph, |run| {
            run.style = Some("Emphasis".into());
        }),
        line(
            "shd=FFFF00",
            |paragraph| paragraph.shading = fill(0xFF_FF00),
            none_run,
        ),
        line(
            "shading.val=pct15 shading.fill=DDEBF7 shading.color=C00000",
            |paragraph| {
                let mut shading = Shading::new(ShadingPattern::Percent15);
                shading.fill = Some(rgb(0xDD_EBF7));
                shading.color = Some(rgb(0xC0_0000));
                paragraph.shading = Some(shading);
            },
            none_run,
        ),
        line(
            "tabs=720,1440",
            |paragraph| {
                paragraph.tabs = Some(vec![
                    TabStop::new(TabKind::Left, Twips(720)),
                    TabStop::new(TabKind::Left, Twips(1440)),
                ]);
            },
            none_run,
        ),
        line(
            "markRPr.italic markRPr.strike markRPr.underline=single markRPr.size=14pt markRPr.highlight=yellow markRPr.font.latin=Georgia markRPr.font.ea=SimSun markRPr.font.cs=Arial",
            |paragraph| {
                let mark = paragraph.mark.get_or_insert_with(Default::default);
                mark.italic = Some(true);
                mark.strike = Some(true);
                mark.underline = underline(UnderlineStyle::Single, None);
                mark.size = Some(HalfPoints(28));
                mark.highlight = Some(Highlight::Yellow);
                mark.fonts = Some(fonts(Some("Georgia"), Some("SimSun"), Some("Arial")));
            },
            none_run,
        ),
        line(
            "border=single",
            |paragraph| {
                let side = Some(border("single", None, None, None));
                let mut borders = ParagraphBorders::default();
                borders.top = side;
                borders.left = side;
                borders.bottom = side;
                borders.right = side;
                paragraph.borders = Some(borders);
            },
            none_run,
        ),
        line(
            "border=single;8;FF0000",
            |paragraph| {
                let side = Some(border("single", Some(8), Some(0xFF_0000), None));
                let mut borders = ParagraphBorders::default();
                borders.top = side;
                borders.left = side;
                borders.bottom = side;
                borders.right = side;
                paragraph.borders = Some(borders);
            },
            none_run,
        ),
        line(
            "border.top=single;8;0070C0 border.bottom=single;8;0070C0",
            |paragraph| {
                let side = Some(border("single", Some(8), Some(0x00_70C0), None));
                let mut borders = ParagraphBorders::default();
                borders.top = side;
                borders.bottom = side;
                paragraph.borders = Some(borders);
            },
            none_run,
        ),
        line(
            "textAlignment=center",
            |paragraph| paragraph.text_alignment = Some(TextAlignment::Center),
            none_run,
        ),
        line(
            "textAlignment=top",
            |paragraph| paragraph.text_alignment = Some(TextAlignment::Top),
            none_run,
        ),
        line(
            "kinsoku=false",
            |paragraph| paragraph.kinsoku = Some(false),
            none_run,
        ),
        line(
            "autoSpaceDE=false autoSpaceDN=false",
            |paragraph| {
                paragraph.auto_space_east_asian_latin = Some(false);
                paragraph.auto_space_east_asian_numbers = Some(false);
            },
            none_run,
        ),
        line(
            "overflowPunct=true topLinePunct=true",
            |paragraph| {
                paragraph.overflow_punctuation = Some(true);
                paragraph.top_line_punctuation = Some(true);
            },
            none_run,
        ),
        line(
            "suppressLineNumbers=true suppressAutoHyphens=true",
            |paragraph| {
                paragraph.suppress_line_numbers = Some(true);
                paragraph.suppress_auto_hyphens = Some(true);
            },
            none_run,
        ),
        line(
            "mirrorIndents=true adjustRightInd=false snapToGrid=false",
            |paragraph| {
                paragraph.mirror_indents = Some(true);
                paragraph.adjust_right_indent = Some(false);
                paragraph.snap_to_grid = Some(false);
            },
            none_run,
        ),
        line(
            "textboxTightWrap=allLines",
            |paragraph| paragraph.textbox_tight_wrap = Some(TextboxTightWrap::AllLines),
            none_run,
        ),
        line(
            "numId=1 numLevel=0",
            |paragraph| {
                let mut numbering = Numbering::default();
                numbering.id = Some(1);
                numbering.level = Some(0);
                paragraph.numbering = Some(numbering);
            },
            none_run,
        ),
    ]
}

fn new_paragraph(line: &Line) -> NewParagraph {
    let mut paragraph = NewParagraph::new();
    (line.paragraph)(paragraph.properties_mut());
    for (text, set) in &line.runs {
        let mut run = NewRun::new(*text);
        set(run.properties_mut());
        paragraph.push_run(run);
    }
    paragraph
}

fn saved(package: &mut Package) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    package.write_to(&mut output).unwrap();
    output.into_inner()
}

fn blank() -> Vec<u8> {
    saved(&mut Package::new().unwrap())
}

/// The paragraphs and direct runs of a saved document, read back through the
/// public readers.
fn read_back(
    bytes: &[u8],
) -> Vec<(
    litchi_docx::properties::Paragraph,
    Vec<(String, litchi_docx::properties::Run)>,
)> {
    let package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let snapshot = package.document_snapshot().unwrap();
    snapshot
        .paragraphs()
        .iter()
        .map(|paragraph| {
            (
                paragraph.direct_properties().unwrap(),
                paragraph
                    .runs()
                    .unwrap()
                    .iter()
                    .map(|run| (run.text().unwrap(), run.direct_properties().unwrap()))
                    .collect(),
            )
        })
        .collect()
}

fn expected(
    line: &Line,
) -> (
    litchi_docx::properties::Paragraph,
    Vec<(String, litchi_docx::properties::Run)>,
) {
    let paragraph = new_paragraph(line);
    (
        paragraph.properties().clone(),
        paragraph
            .runs()
            .iter()
            .map(|run| (run.text().to_owned(), run.properties().clone()))
            .collect(),
    )
}

/// The OfficeCLI property sets, inserted with their properties (`add`) on a
/// blank document and saved, read back exactly; the same sets applied to
/// plain paragraphs afterwards (`set`) produce byte-identical paragraphs.
#[test]
fn officecli_property_sets_read_back_from_a_blank_document() {
    for (script, lines) in [
        ("run-formatting.sh", run_formatting()),
        ("paragraph-formatting.sh", paragraph_formatting()),
    ] {
        // `add`: each paragraph inserted with its properties and runs.
        let mut added = Package::from_reader(Cursor::new(blank())).unwrap();
        let base = added.document_snapshot().unwrap();
        let existing = base.paragraph_count();
        let mut edit = added.edit().unwrap();
        for line in &lines {
            edit.insert_formatted_paragraph(BodyAnchor::End, &new_paragraph(line))
                .unwrap_or_else(|error| panic!("{script} {}: {error}", line.keys));
        }
        let commit = edit.commit().unwrap();
        let added_document = commit.patch().target().xml_bytes().to_vec();
        added = added.apply(commit.patch()).unwrap().into_package();
        let added_bytes = saved(&mut added);
        let readback = read_back(&added_bytes);
        assert_eq!(readback.len(), existing + lines.len());
        for (line, actual) in lines.iter().zip(&readback[existing..]) {
            assert_eq!(*actual, expected(line), "{script} {}", line.keys);
        }

        // `set`: plain paragraphs and runs formatted in place, runs after
        // the first appended to the existing paragraph.
        let mut set = Package::from_reader(Cursor::new(blank())).unwrap();
        let mut edit = set.edit().unwrap();
        for line in &lines {
            let mut plain = NewParagraph::new();
            if let Some((text, _)) = line.runs.first() {
                plain.push_run(NewRun::new(*text));
            }
            let position = edit
                .insert_formatted_paragraph(BodyAnchor::End, &plain)
                .unwrap();
            for (text, _) in line.runs.iter().skip(1) {
                edit.append_run(position, &NewRun::new(*text)).unwrap();
            }
            edit.format_paragraph(position, line.paragraph).unwrap();
            for (index, (_, run)) in line.runs.iter().enumerate() {
                edit.format_run(position, Position::new(index), *run)
                    .unwrap_or_else(|error| panic!("{script} {}: {error}", line.keys));
            }
        }
        let commit = edit.commit().unwrap();
        // Every property element lands in schema order with canonical
        // attributes either way, so the two documents are the same bytes.
        assert_eq!(
            commit.patch().target().xml_bytes(),
            added_document.as_slice(),
            "{script}: `set` and `add` disagree"
        );
        set = set.apply(commit.patch()).unwrap().into_package();
        assert_eq!(read_back(&saved(&mut set)), readback, "{script}");
    }
}

fn fixture(relative: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../..")
            .join(relative),
    )
    .unwrap_or_else(|error| panic!("{relative}: {error}"))
}

/// Count every occurrence of each exact element spelling.
fn occurrences(haystack: &[u8], needle: &[u8]) -> usize {
    haystack
        .windows(needle.len())
        .filter(|window| *window == needle)
        .count()
}

/// The exact markup of every element named `local` in `xml`.
fn elements(xml: &[u8], local: &str) -> Vec<Vec<u8>> {
    let open = format!("<w:{local}");
    let close = format!("</w:{local}>");
    let mut found = Vec::new();
    let mut cursor = 0usize;
    while let Some(offset) = xml[cursor..]
        .windows(open.len())
        .position(|window| window == open.as_bytes())
    {
        let start = cursor + offset;
        let after = xml[start + open.len()];
        if !matches!(after, b' ' | b'>' | b'/') {
            cursor = start + open.len();
            continue;
        }
        let tag_end = start + xml[start..].iter().position(|byte| *byte == b'>').unwrap() + 1;
        let end = if xml[tag_end - 2] == b'/' {
            tag_end
        } else {
            tag_end
                + xml[tag_end..]
                    .windows(close.len())
                    .position(|window| window == close.as_bytes())
                    .unwrap()
                + close.len()
        };
        found.push(xml[start..end].to_vec());
        cursor = end;
    }
    found
}

/// OfficeCLI's own `revisions.docx`, a real producer's output with tracked
/// insertions, deletions and property changes: formatting every paragraph
/// and every direct run keeps every change-tracking element byte for byte.
#[test]
fn tracked_changes_of_a_real_producer_survive_formatting() {
    let source = fixture("test-data/officecli/examples/word/revisions.docx");
    let package = Package::from_reader(Cursor::new(source.as_slice())).unwrap();
    let base = package.document_snapshot().unwrap();
    let tracked = ["rPrChange", "pPrChange", "ins", "del"];
    let before = tracked.map(|local| elements(base.xml_bytes(), local));
    assert!(
        before.iter().map(Vec::len).sum::<usize>() >= 20,
        "the fixture should carry tracked changes"
    );
    let mut edit = package.edit().unwrap();
    let mut runs = 0usize;
    for position in 0..base.paragraph_count() {
        let paragraph = Position::new(position);
        edit.format_paragraph(paragraph, |properties| {
            properties.keep_lines = Some(true);
            properties.mark.get_or_insert_with(Default::default).caps = Some(true);
        })
        .unwrap();
        let mut run = 0usize;
        while edit
            .format_run(paragraph, Position::new(run), |properties| {
                properties.bold = Some(!properties.bold.unwrap_or(false));
                properties.size = Some(HalfPoints(23));
                properties.color = color(0x12_3456);
            })
            .is_ok()
        {
            run += 1;
        }
        runs += run;
    }
    assert!(runs >= 10, "{runs} direct runs formatted");
    let commit = edit.commit().unwrap();
    let after = commit.patch().target().xml_bytes();
    for (local, before) in tracked.iter().zip(&before) {
        for element in before {
            assert!(
                occurrences(after, element) >= occurrences(base.xml_bytes(), element),
                "w:{local} changed: {}",
                String::from_utf8_lossy(element)
            );
        }
        assert_eq!(elements(after, local).len(), before.len(), "w:{local}");
    }
    let restored = commit
        .patch()
        .inverse()
        .apply(commit.patch().target())
        .unwrap();
    assert_eq!(restored.xml_bytes(), base.xml_bytes());
}

/// Formatting and structure edits publish through the source-backed owner
/// with every other package member raw-copied, and through the owning
/// package; both outputs read back identically.
#[test]
fn both_owners_publish_formatting_and_structure_identically() {
    for relative in [
        "test-data/ooxml/docx/drawing.docx",
        "test-data/poi/test-data/document/saut_page.docx",
    ] {
        let source = fixture(relative);
        fn stage<P: Clone>(edit: &mut Edit<P>) {
            edit.format_run(Position::new(0), Position::new(0), |run| {
                run.bold = Some(true);
                run.size = Some(HalfPoints(30));
            })
            .unwrap_or_else(|error| panic!("formatting edit failed: {error}"));
            edit.format_paragraph(Position::new(0), |paragraph| {
                paragraph.justification = Some(Justification::Center);
            })
            .unwrap();
            let mut new = NewParagraph::new().with_run(NewRun::new("inserted"));
            new.properties_mut().style = Some("Title".into());
            let inserted = edit
                .insert_formatted_paragraph(
                    BodyAnchor::After(BodyBlock::Paragraph(Position::new(0))),
                    &new,
                )
                .unwrap();
            edit.append_run(inserted, &NewRun::new(" tail")).unwrap();
            edit.move_paragraph(
                inserted,
                BodyAnchor::Before(BodyBlock::Paragraph(Position::new(0))),
            )
            .unwrap();
        }

        let owned = Package::from_reader(Cursor::new(source.as_slice())).unwrap();
        let mut owned_edit = owned.edit().unwrap();
        stage(&mut owned_edit);
        let commit = owned_edit.commit().unwrap();
        let mut owned = owned;
        owned = owned.apply(commit.patch()).unwrap().into_package();
        let owned_bytes = saved(&mut owned);

        let backed = source_backed::Package::from_reader(Cursor::new(source.as_slice())).unwrap();
        let mut backed_edit = backed.edit().unwrap();
        stage(&mut backed_edit);
        let commit = backed_edit.commit().unwrap();
        let mut backed_bytes = Vec::new();
        backed
            .apply(commit.patch())
            .and_then(|committed| {
                let committed = committed.into_package();
                committed.write_to(&mut backed_bytes)?;
                committed.document_snapshot()
            })
            .unwrap();

        assert_eq!(
            read_back(&owned_bytes),
            read_back(&backed_bytes),
            "{relative}"
        );
        let reopened = Package::from_reader(Cursor::new(backed_bytes)).unwrap();
        let snapshot = reopened.document_snapshot().unwrap();
        let first = snapshot.paragraph(Position::new(0)).unwrap();
        assert_eq!(first.text().unwrap(), "inserted tail");
        assert_eq!(
            first.direct_properties().unwrap().style.as_deref(),
            Some("Title")
        );
    }
}

fn signed() -> Vec<u8> {
    fixture("test-data/poi/test-data/xmldsign/hello-world-signed.docx")
}

/// A changed publication of a signed package is refused by both owners until
/// the signature is removed explicitly; an exact no-op keeps the signature.
#[test]
fn signed_packages_refuse_formatting_until_explicitly_unsigned() {
    let source = signed();
    let mut package = Package::from_reader(Cursor::new(source.as_slice())).unwrap();
    assert!(package.is_signed());
    let original = saved(&mut package);
    let mut noop = package.edit().unwrap();
    noop.format_run(Position::new(0), Position::new(0), |_| {})
        .unwrap();
    package = package
        .apply(noop.commit().unwrap().patch())
        .unwrap()
        .into_package();
    assert!(package.is_signed());
    assert_eq!(saved(&mut package), original);

    fn stage<P: Clone>(edit: &mut Edit<P>) {
        edit.format_run(Position::new(0), Position::new(0), |run| {
            run.italic = Some(true);
        })
        .unwrap();
        edit.insert_formatted_paragraph(BodyAnchor::End, &NewParagraph::new())
            .unwrap();
    }
    let mut package_edit = package.edit().unwrap();
    stage(&mut package_edit);
    assert!(matches!(
        package_edit.commit(),
        Err(litchi_docx::document::Error::Document(
            litchi_docx::Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy)
        ))
    ));
    assert!(package.is_signed());
    assert_eq!(saved(&mut package), original);

    let unsigned_for_patch = package.unsign();
    let mut unsigned_edit = unsigned_for_patch.edit().unwrap();
    stage(&mut unsigned_edit);
    let commit = unsigned_edit.commit().unwrap();
    assert!(package.apply(commit.patch()).is_err());

    let backed = source_backed::Package::from_reader(Cursor::new(source.as_slice())).unwrap();
    let mut backed_edit = backed.edit().unwrap();
    stage(&mut backed_edit);
    assert!(matches!(
        backed_edit.commit(),
        Err(litchi_docx::document::Error::Document(
            litchi_docx::Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy)
        ))
    ));
    assert!(matches!(
        backed.apply(commit.patch()),
        Err(litchi_docx::document::Error::PatchConflict)
    ));

    package = package.unsign();
    package = package.apply(commit.patch()).unwrap().into_package();
    assert!(!package.is_signed());
    let reopened = Package::from_reader(Cursor::new(saved(&mut package))).unwrap();
    let snapshot = reopened.document_snapshot().unwrap();
    let paragraph = snapshot.paragraph(Position::new(0)).unwrap();
    assert_eq!(
        paragraph.runs().unwrap()[0]
            .direct_properties()
            .unwrap()
            .italic,
        Some(true)
    );
}

/// Formatting a theme-coloured run keeps the theme reference and transforms.
#[test]
fn theme_colours_round_trip_through_the_engine() {
    let mut edit = raw::Tree::from_xml(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:color w:val="4472C4" w:themeColor="accent1"/></w:rPr><w:t>x</w:t></w:r></w:p></w:body></w:document>"#.as_bytes().to_vec(),
        raw::Limits::default(),
    )
    .unwrap();
    edit.format_run(Position::new(0), Position::new(0), |run| {
        let color = run.color.as_mut().unwrap();
        color.theme = Some(Theme::Accent2);
        color.theme_shade = Some(0xBF);
    })
    .unwrap();
    assert_eq!(
        std::str::from_utf8(&edit.to_bytes(raw::Limits::default()).unwrap()).unwrap(),
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:color w:val="4472C4" w:themeColor="accent2" w:themeShade="BF"/></w:rPr><w:t>x</w:t></w:r></w:p></w:body></w:document>"#
    );
}

#[test]
fn canonical_owner_commits_retain_source_and_support_inverse() {
    fn stage<P: Clone>(edit: &mut Edit<P>) {
        let first = edit
            .insert_formatted_paragraph(
                BodyAnchor::End,
                &NewParagraph::new().with_run(NewRun::new("PRIVATE")),
            )
            .unwrap();
        edit.format_run(first, Position::new(0), |run| run.bold = Some(true))
            .unwrap();
        edit.format_paragraph(first, |paragraph| {
            paragraph.justification = Some(Justification::Center)
        })
        .unwrap();
        edit.append_run(first, &NewRun::new("tail")).unwrap();
        edit.insert_run(first, Position::new(0), &NewRun::new("head"))
            .unwrap();
        edit.remove_run(first, Position::new(0)).unwrap();
        let second = edit
            .insert_formatted_paragraph(BodyAnchor::End, &NewParagraph::new())
            .unwrap();
        edit.move_paragraph(second, BodyAnchor::Before(BodyBlock::Paragraph(first)))
            .unwrap();
        edit.remove_paragraph(Position::new(first.get())).unwrap();
    }
    let package = Package::new().unwrap();
    let before = package.document_snapshot().unwrap().xml_bytes().to_vec();
    let mut edit = package.edit().unwrap();
    stage(&mut edit);
    let commit = edit.commit().unwrap();
    assert_eq!(package.document_snapshot().unwrap().xml_bytes(), before);
    assert!(commit.diagnostics().changed());
    let inverse = commit.patch().inverse();
    let changed = commit.into_package();
    let restored = changed.apply(&inverse).unwrap().into_package();
    assert_eq!(restored.document_snapshot().unwrap().xml_bytes(), before);

    let source = source_backed::Package::from_bytes(&blank()).unwrap();
    let before = source.document_snapshot().unwrap().xml_bytes().to_vec();
    let mut edit = source.edit().unwrap();
    stage(&mut edit);
    let commit = edit.commit().unwrap();
    assert_eq!(source.document_snapshot().unwrap().xml_bytes(), before);
    let inverse = commit.patch().inverse();
    let changed = commit.into_package();
    let restored = changed.apply(&inverse).unwrap().into_package();
    assert_eq!(restored.document_snapshot().unwrap().xml_bytes(), before);
}
